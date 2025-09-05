/**
  * All Graph calls live here. We centralize retry/backoff and pagination so the rest
 * of the app can call these small helpers. Everything returns plain JSON the LLM can
 * easily consume.
 */

const axios = require("axios");

/* -------------------- Helpers -------------------- */
/**
 * GET with backoff on 429/5xx. We honor Retry-After seconds if provided; otherwise
 * exponential backoff: 1, 2, 4, 8, 16 (capped at 16).
 */
async function httpGetWithBackoff(url, headers, maxRetries = 4) {
  let attempt = 0;
  for (;;) {
    try {
      return await axios.get(url, { headers, timeout: 30000 });
    } catch (e) {
      const s = e.response?.status;
      if ((s === 429 || s >= 500) && attempt < maxRetries) {
        const retryAfter = Number(e.response?.headers?.["retry-after"]) || Math.min(2 ** attempt, 16);
        await new Promise(r => setTimeout(r, retryAfter * 1000));
        attempt++;
        continue;
      }
      throw e;
    }
  }
}

/** Generic collection paging. */
async function* listCollection({ url, headers }) {
  for (;;) {
    const resp = await httpGetWithBackoff(url, headers);
    const items = resp.data?.value || [];
    for (const it of items) yield it;
    const next = resp.data?.["@odata.nextLink"];
    if (!next) break;
    url = next; /
  }
}

/** Messages paging (same as listCollection, just exists for clarity) */
async function* listMessages({ url, headers }) {
  for (;;) {
    const resp = await httpGetWithBackoff(url, headers);
    const items = resp.data?.value || [];
    for (const m of items) yield m;
    const next = resp.data?.["@odata.nextLink"];
    if (!next) break;
    url = next;
  }
}

/** Normalize a Graph message into a compact shape */
function mapMsg(m) {
  return {
    id: m.id,
    received: m.receivedDateTime,
    sent: m.sentDateTime || m.createdDateTime,
    from: m.from?.emailAddress?.address,
    subject: m.subject,
    preview: m.bodyPreview,
    folderId: m.parentFolderId
  };
}

/* -------------------- Readers (mailbox) -------------------- */

/** Latest messages across ALL folders (first N). */
async function readLatest(access_token, top = 10) {
  const url = new URL("https://graph.microsoft.com/v1.0/me/messages");
  url.searchParams.set("$orderby", "receivedDateTime desc");
  url.searchParams.set("$top", String(top));
  url.searchParams.set("$select", "id,receivedDateTime,sentDateTime,createdDateTime,subject,bodyPreview,from,parentFolderId");

  const headers = { Authorization: `Bearer ${access_token}`, Prefer: 'outlook.body-content-type="text"' };

  const out = [];
  for await (const m of listMessages({ url: url.toString(), headers })) {
    out.push(mapMsg(m));
    if (out.length >= top) break;
  }
  return { results: out, count: out.length };
}

/** Latest from Sent Items (first N). */
async function readSentLatest(access_token, top = 10) {
  const url = new URL("https://graph.microsoft.com/v1.0/me/mailFolders/SentItems/messages");
  url.searchParams.set("$orderby", "sentDateTime desc");
  url.searchParams.set("$top", String(top));
  url.searchParams.set("$select", "id,receivedDateTime,sentDateTime,createdDateTime,subject,bodyPreview,from,parentFolderId");

  const headers = { Authorization: `Bearer ${access_token}`, Prefer: 'outlook.body-content-type="text"' };

  const out = [];
  for await (const m of listMessages({ url: url.toString(), headers })) {
    out.push(mapMsg(m));
    if (out.length >= top) break;
  }
  return { results: out, count: out.length };
}

/** Whole mailbox (deep paginate) up to `max`. */
async function readAllMailbox({ access_token, max = 1000 }) {
  let url = new URL("https://graph.microsoft.com/v1.0/me/messages");
  url.searchParams.set("$orderby", "receivedDateTime desc");
  url.searchParams.set("$top", "100");
  url.searchParams.set("$select", "id,receivedDateTime,sentDateTime,createdDateTime,subject,bodyPreview,from,parentFolderId");

  const headers = { Authorization: `Bearer ${access_token}`, Prefer: 'outlook.body-content-type="text"' };

  const out = [];
  for await (const m of listMessages({ url: url.toString(), headers })) {
    out.push(mapMsg(m));
    if (out.length >= max) break;
  }
  return { results: out, count: out.length };
}

/** Specific folder by display name (OK, but localized). Prefer readFolderByIdAll. */
async function readFolderAll({ access_token, folderName = "Inbox", max = 1000 }) {
  const enc = encodeURIComponent(folderName);
  let url = new URL(`https://graph.microsoft.com/v1.0/me/mailFolders/${enc}/messages`);
  url.searchParams.set("$orderby", "receivedDateTime desc");
  url.searchParams.set("$top", "100");
  url.searchParams.set("$select", "id,receivedDateTime,sentDateTime,createdDateTime,subject,bodyPreview,from,parentFolderId");

  const headers = { Authorization: `Bearer ${access_token}`, Prefer: 'outlook.body-content-type="text"' };

  const out = [];
  for await (const m of listMessages({ url: url.toString(), headers })) {
    out.push(mapMsg(m));
    if (out.length >= max) break;
  }
  return { results: out, count: out.length, folder: folderName };
}



/** AQS/keyword search (first N). */
async function searchAllMail({ access_token, query, top = 50 }) {
  let url = new URL("https://graph.microsoft.com/v1.0/me/messages");
  url.searchParams.set("$search", `"${query}"`);
  url.searchParams.set("$top", String(Math.min(top, 100)));
  url.searchParams.set("$select", "id,receivedDateTime,subject,bodyPreview,from,parentFolderId");

  const headers = {
    Authorization: `Bearer ${access_token}`,
    Prefer: 'outlook.body-content-type="text"',
    ConsistencyLevel: "eventual",
  };

  const out = [];
  for (;;) {
    const resp = await httpGetWithBackoff(url.toString(), headers);
    for (const m of resp.data?.value || []) {
      out.push(mapMsg(m));
      if (out.length >= top) {
        out.sort((a, b) => (a.received < b.received ? 1 : -1));
        return { results: out, count: out.length };
      }
    }
    const next = resp.data?.["@odata.nextLink"];
    if (!next) break;
    url = new URL(next);
  }
  out.sort((a, b) => (a.received < b.received ? 1 : -1));
  return { results: out, count: out.length };
}


/** Deep AQS (all pages up to max). */
async function searchAllMailAllPages({ access_token, query, max = 1000 }) {
  let url = new URL("https://graph.microsoft.com/v1.0/me/messages");
  url.searchParams.set("$search", `"${query}"`);
  url.searchParams.set("$top", "100");
  url.searchParams.set("$select", "id,receivedDateTime,subject,bodyPreview,from,parentFolderId");

  const headers = {
    Authorization: `Bearer ${access_token}`,
    Prefer: 'outlook.body-content-type="text"',
    ConsistencyLevel: "eventual"
  };

  const out = [];
  for (;;) {
    const resp = await httpGetWithBackoff(url.toString(), headers);
    for (const m of resp.data?.value || []) {
      out.push(mapMsg(m));
      if (out.length >= max) return { results: out, count: out.length };
    }
    const next = resp.data?.["@odata.nextLink"];
    if (!next) break;
    url = new URL(next);
  }
  return { results: out, count: out.length };
}

/** Absolute date window (ALL folders). */
async function filterAllMailByDate({ access_token, startIso, endIso, top = 200 }) {
  const url = new URL("https://graph.microsoft.com/v1.0/me/messages");
  url.searchParams.set("$filter", `receivedDateTime ge ${startIso} and receivedDateTime le ${endIso}`);
  url.searchParams.set("$orderby", "receivedDateTime desc");
  url.searchParams.set("$top", String(top));
  url.searchParams.set("$select", "id,receivedDateTime,subject,bodyPreview,from,parentFolderId");

  const headers = { Authorization: `Bearer ${access_token}`, Prefer: 'outlook.body-content-type="text"' };

  const out = [];
  for await (const m of listMessages({ url: url.toString(), headers })) {
    out.push(mapMsg(m));
    if (out.length >= top) break;
  }
  return { results: out, count: out.length };
}



/** Exact sender email (exhaustive crawl via $filter). */

async function searchBySenderEmail({ access_token, email, limit = 2000, startIso, endIso }) {
  if (!email) throw new Error("email is required");
  let filter = `from/emailAddress/address eq '${email.replace(/'/g, "''")}'`;
  if (startIso && endIso) {
    filter += ` and receivedDateTime ge ${startIso} and receivedDateTime le ${endIso}`;
  }

  let url = new URL("https://graph.microsoft.com/v1.0/me/messages");
  url.searchParams.set("$filter", filter);
  url.searchParams.set("$select", "id,receivedDateTime,subject,bodyPreview,from,parentFolderId");
  url.searchParams.set("$top", "100");

  const headers = { Authorization: `Bearer ${access_token}`, Prefer: 'outlook.body-content-type="text"' };

  const out = [];
  for (;;) {
    const resp = await httpGetWithBackoff(url.toString(), headers);
    for (const m of resp.data?.value || []) {
      out.push(mapMsg(m));
      if (out.length >= limit) {
        out.sort((a, b) => (a.received < b.received ? 1 : -1));
        return { results: out, count: out.length, email };
      }
    }
    const next = resp.data?.["@odata.nextLink"];
    if (!next) break;
    url = new URL(next);
  }
  out.sort((a, b) => (a.received < b.received ? 1 : -1));
  return { results: out, count: out.length, email };
}

/**
 * Name bootstrap (no /me/people, no User.Read):
 * 1) AQS search for "from:<name>" across all folders (sample up to maxAqs)
 * 2) Collect unique From: addresses
 * 3) For each discovered address, run exact-email crawl to fetch complete history
 */
async function searchSenderByNameBootstrap({ access_token, name, maxAqs = 300, perSenderLimit = 2000 }) {
  if (!name) throw new Error("name is required");

  const aqs = await searchAllMailAllPages({
    access_token,
    query: `from:${name}`,
    max: maxAqs
  });

  const senders = new Set();
  for (const m of aqs.results) {
    const addr = (m.from || "").trim();
    if (addr) senders.add(addr.toLowerCase());
  }
  if (senders.size === 0) return { results: [], count: 0, discoveredSenders: [] };

  const all = [];
  for (const email of senders) {
    const res = await searchBySenderEmail({ access_token, email, limit: perSenderLimit });
    all.push(...res.results);
  }

  // Dedupe by ID, sort by received desc
  const seen = new Set();
  const uniq = all.filter(m => (seen.has(m.id) ? false : (seen.add(m.id), true)));
  uniq.sort((a, b) => (a.received < b.received ? 1 : -1));

  return { results: uniq, count: uniq.length, discoveredSenders: Array.from(senders) };
}

/* -------------------- Folders & Search Folders -------------------- */

/** List ALL folders (names & ids). */
async function listAllFolders(access_token) {
  let url = new URL("https://graph.microsoft.com/v1.0/me/mailFolders");
  url.searchParams.set("$top", "100");
  url.searchParams.set("$select", "id,displayName,childFolderCount,totalItemCount,unreadItemCount");

  const headers = { Authorization: `Bearer ${access_token}` };

  const out = [];
  for await (const f of listCollection({ url: url.toString(), headers })) {
    out.push(f);
  }
  out.sort((a, b) => (a.displayName || "").localeCompare(b.displayName || ""));
  return out;
}

/** Read messages from a folder by its opaque ID (recommended). */
async function readFolderByIdAll({ access_token, folderId, max = 1000 }) {
  if (!folderId) throw new Error("folderId is required");

  let url = new URL(`https://graph.microsoft.com/v1.0/me/mailFolders/${encodeURIComponent(folderId)}/messages`);
  url.searchParams.set("$orderby", "receivedDateTime desc");
  url.searchParams.set("$top", "100");
  url.searchParams.set("$select", "id,receivedDateTime,sentDateTime,createdDateTime,subject,bodyPreview,from,parentFolderId");

  const headers = { Authorization: `Bearer ${access_token}`, Prefer: 'outlook.body-content-type="text"' };

  const out = [];
  for await (const m of listMessages({ url: url.toString(), headers })) {
    out.push(mapMsg(m));
    if (out.length >= max) break;
  }
  return { results: out, count: out.length, folderId };
}

/**
 * List Search Folders (virtual folders). We first try the well-known container,
 * then fall back to a heuristic on displayName.
 */
async function listSearchFolders(access_token) {
  const headers = { Authorization: `Bearer ${access_token}` };
  const out = [];

  // Try: /me/mailFolders('searchfolders')/childFolders
  try {
    let url = new URL("https://graph.microsoft.com/v1.0/me/mailFolders('searchfolders')/childFolders");
    url.searchParams.set("$top", "100");
    url.searchParams.set("$select", "id,displayName,childFolderCount,totalItemCount,unreadItemCount");
    for await (const f of listCollection({ url: url.toString(), headers })) out.push(f);
  } catch (_) {
    
  }

  if (out.length === 0) {
    // Fallback: scan all folders and pick those that look like search folders by name.
    const all = await listAllFolders(access_token);
    const candidates = all.filter(f => {
      const name = (f.displayName || "").toLowerCase();
      return name.includes("search") && name.includes("folder");
    });
    return candidates;
  }

  out.sort((a, b) => (a.displayName || "").localeCompare(b.displayName || ""));
  return out;
}

/** Read messages from a Search Folder by ID (same as normal folder by ID). */
async function readSearchFolderByIdAll({ access_token, folderId, max = 1000 }) {
  return readFolderByIdAll({ access_token, folderId, max });
}


module.exports = {
  // Readers
  readLatest,
  readSentLatest,
  readAllMailbox,
  readFolderAll,

  // Search
  searchAllMail,
  searchAllMailAllPages,
  filterAllMailByDate,

  // Sender
  searchBySenderEmail,
  searchSenderByNameBootstrap,

  // Folders
  listAllFolders,
  readFolderByIdAll,
  listSearchFolders,
  readSearchFolderByIdAll,

  // Utilities (exported for reuse if needed)
  httpGetWithBackoff,
  listMessages,
  mapMsg
};

