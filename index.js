const STRICT_USER_ID = String(process.env.STRICT_USER_ID || "false").toLowerCase() === "true";
require("dotenv").config();
const express = require("express");
const bodyParser = require("body-parser");
const cors = require("cors");
const path = require("path");
const crypto = require("crypto");
const { DateTime, Interval } = require("luxon");

const { router: authRouter } = require("./auth");
const tokenStore = require("./tokenStore");
const graph = require("./graph");

const app = express();


const API_KEY  = process.env.API_KEY || "CROWDdoingtools";
const PORT     = Number(process.env.PORT || 3001);
const HOST     = process.env.HOST || "0.0.0.0";
const BASE_URL = process.env.PUBLIC_BASE_URL || `http://localhost:${PORT}`;

/* Middleware*/
app.use(cors({ origin: "*", allowedHeaders: ["Content-Type", "Authorization"] }));
app.use(bodyParser.json());
app.use(authRouter);


app.use((req, _res, next) => {
  console.log(`${new Date().toISOString()} - ${req.method} ${req.url}`);
  next();
});

/*Token helper*/
function getValidToken(user_id) {
  const rec = tokenStore.get(user_id);
  if (!rec) return null;
  const skewMs = 60 * 1000; // treat as expired if < 60s remain
  if (rec.expiry - skewMs > Date.now()) return rec.access_token;
  return null; // expired or near-expiry
}

// Cleanup expired tokens on boot
try {
  const removed = tokenStore.deleteExpired();
  if (removed) console.log(`[tokenStore] removed ${removed} expired token(s)`);
} catch (e) {
  console.warn("token cleanup failed:", e.message);
}

/*  Relative date helpers*/
function toIso(dt) { return dt.toISO({ suppressMilliseconds: true }); }
function computeRange({ tz, intent, n, on, since, start, end }) {
  const now = DateTime.now().setZone(tz || "America/Chicago");
  let s, e;

  switch ((intent || "").trim()) {
    case "today":       s = now.startOf("day"); e = now.endOf("day"); break;
    case "yesterday":   { const y = now.minus({ days: 1 }); s = y.startOf("day"); e = y.endOf("day"); break; }
    case "this_week":   s = now.startOf("week"); e = now.endOf("week"); break;
    case "last_week":   { const w = now.minus({ weeks: 1 }); s = w.startOf("week"); e = w.endOf("week"); break; }
    case "this_month":  s = now.startOf("month"); e = now.endOf("month"); break;
    case "last_month":  { const m = now.minus({ months: 1 }); s = m.startOf("month"); e = m.endOf("month"); break; }
    case "last_n_days": { const d = Number(n) || 7; s = now.minus({ days: d }).startOf("day"); e = now.endOf("day"); break; }
    case "on_date":     { const d = DateTime.fromISO(on, { zone: tz }); if (!d.isValid) throw new Error("Invalid 'on' date"); s = d.startOf("day"); e = d.endOf("day"); break; }
    case "since_date":  { const d = DateTime.fromISO(since, { zone: tz }); if (!d.isValid) throw new Error("Invalid 'since' date"); s = d.startOf("day"); e = now.endOf("day"); break; }
    case "between":     {
      const ds = DateTime.fromISO(start, { zone: tz });
      const de = DateTime.fromISO(end,   { zone: tz });
      if (!ds.isValid || !de.isValid) throw new Error("Invalid 'between' dates");
      s = ds.startOf("day"); e = de.endOf("day");
      if (Interval.fromDateTimes(s, e).length("minutes") < 0) [s, e] = [e.startOf("day"), s.endOf("day")];
      break;
    }
    default:            s = now.minus({ days: 7 }).startOf("day"); e = now.endOf("day");
  }
  return { startIso: toIso(s), endIso: toIso(e), tz: now.zoneName };
}

/* Tool endpoint */
app.post("/execute_tool", async (req, res) => {
  // 1) API key check
  // console.log("[EXECUTE_TOOL] raw body:", JSON.stringify(req.body));
  const authHeader = req.headers.authorization || "";
  const providedKey = authHeader.startsWith("Bearer ") ? authHeader.slice(7) : null;
  if (!providedKey || providedKey !== API_KEY) {
    return res.status(401).json({
      error: "Invalid API key",
      requires_login: true,
      login_url: `${BASE_URL}/login?user_id=temp`
    });
  }

  // 2) LLM MUST send a user_id. If missing/new, we mint one and ask user to log in.
  const STRICT_USER_ID = true; 
  const UUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i;
  const BAD_TOKENS = new Set(["", "new", "temp", "current", "me", "self"]);

  let user_id = (req.body?.user_id || "").toString().trim();
  // console.log("[AUTH] candidate user_id from client:", user_id || "(missing)");

  if (STRICT_USER_ID) {
    if (!user_id || BAD_TOKENS.has(user_id.toLowerCase()) || !UUID_RE.test(user_id)) {
      const minted = crypto.randomUUID();
      return res.status(400).json({
        error: "user_id_required",
        message:
          "Always send the SAME `user_id` you previously received as `user_id_used`. " +
          "Use the `user_id` below to login, then reuse it on every request.",
        requires_login: true,
        user_id: minted,
        login_url: `${BASE_URL}/login?user_id=${encodeURIComponent(minted)}`
      });
    }
  }

  // 3) Validate token for that user_id
  const token = getValidToken(user_id);
  // console.log("[AUTH] using user_id:", user_id, "| hasToken:", !!token);
  if (!token) {
    // Token missing/expired — ask to login again, but KEEP the same user_id
    return res.json({
      requires_login: true,
      user_id,
      login_url: `${BASE_URL}/login?user_id=${encodeURIComponent(user_id)}`
    });
  }

  // Helper to always return the id used (so the client can persist it).

  const ok = (payload) => {
    res.setHeader("X-User-Id-Used", user_id);
    return res.json({ user_id_used: user_id, ...payload });
  };

  // 4) Execute action
  try {
    const { action, inputs } = req.body || {};

    if (action === "read") {
      const top = Number(inputs?.top || 10);
      const data = await graph.readLatest(token, top);
      return ok(data);
    }

    if (action === "read_sent") {
      const top = Number(inputs?.top || 10);
      const data = await graph.readSentLatest(token, top);
      return ok(data);
    }

    if (action === "read_all") {
      const max = Number(inputs?.max || 1000);
      const data = await graph.readAllMailbox({ access_token: token, max });
      return ok(data);
    }

    if (action === "list_folders") {
      const folders = await graph.listAllFolders(token);
      return ok({ folders });
    }

    if (action === "read_folder_all") {
      const folder = String(inputs?.folder || "Inbox");
      const max = Number(inputs?.max || 1000);
      const data = await graph.readFolderAll({ access_token: token, folderName: folder, max });
      return ok(data);
    }

    if (action === "read_folder_id_all") {
      const folderId = String(inputs?.folderId || "");
      const max = Number(inputs?.max || 1000);
      if (!folderId) return res.status(400).json({ error: "folderId is required" });
      const data = await graph.readFolderByIdAll({ access_token: token, folderId, max });
      return ok(data);
    }

    if (action === "list_search_folders") {
      const folders = await graph.listSearchFolders(token);
      return ok({ folders });
    }

    if (action === "read_search_folder_id_all") {
      const folderId = String(inputs?.folderId || "");
      const max = Number(inputs?.max || 1000);
      if (!folderId) return res.status(400).json({ error: "folderId is required" });
      const data = await graph.readSearchFolderByIdAll({ access_token: token, folderId, max });
      return ok(data);
    }

    if (action === "read_relative") {
      const { startIso, endIso, tz } = computeRange({
        tz: (inputs?.tz || "America/Chicago").trim(),
        intent: inputs?.intent,
        n: inputs?.n,
        on: inputs?.on,
        since: inputs?.since,
        start: inputs?.start,
        end: inputs?.end
      });
      const top = Number(inputs?.top || 5000);
      const data = await graph.filterAllMailByDate({ access_token: token, startIso, endIso, top });
      return ok({ range: { startIso, endIso, tz }, ...data });
    }

    if (action === "search") {
      const q = String(inputs?.q || "").trim();
      const top = Number(inputs?.top || 50);
      const data = await graph.searchAllMail({ access_token: token, query: q, top });
      return ok(data);
    }

    if (action === "search_all") {
      const q = String(inputs?.q || "").trim();
      const max = Number(inputs?.max || 1000);
      const data = await graph.searchAllMailAllPages({ access_token: token, query: q, max });
      return ok(data);
    }

    if (action === "search_by_date") {
      const startIso = String(inputs?.startIso || "").trim();
      const endIso   = String(inputs?.endIso || "").trim();
      const top      = Number(inputs?.top || 200);
      if (!startIso || !endIso) return res.status(400).json({ error: "startIso and endIso are required (ISO 8601)" });
      const data = await graph.filterAllMailByDate({ access_token: token, startIso, endIso, top });
      return ok(data);
    }

    if (action === "search_sender_email") {
      const email = String(inputs?.email || "").trim();
      const limit = Number(inputs?.limit || 2000);
      if (!email) return res.status(400).json({ error: "email is required" });
      const data = await graph.searchBySenderEmail({ access_token: token, email, limit });
      return ok(data);
    }

    if (action === "search_sender_name_bootstrap") {
      const name = String(inputs?.name || "").trim();
      const maxAqs = Number(inputs?.maxAqs || 300);
      const perSenderLimit = Number(inputs?.perSenderLimit || 2000);
      if (!name) return res.status(400).json({ error: "name is required" });
      const data = await graph.searchSenderByNameBootstrap({
        access_token: token, name, maxAqs, perSenderLimit
      });
      return ok(data);
    }

    return ok({ error: "Invalid action" });
  } catch (err) {
    const payload = err?.response?.data || err?.message || "Unknown error";
    console.error("execute_tool error:", payload);
    return res.status(400).json({ error: payload });
  }
});


app.get("/openapi.json", (_req, res) => {
  res.sendFile(path.join(__dirname, "openapi.json"));
});


app.listen(PORT, HOST, () => {
  console.log(`Server running at ${BASE_URL}`);
});
