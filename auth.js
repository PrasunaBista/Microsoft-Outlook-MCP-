/**
 * auth.js
 *
 * Public-client OAuth 2.0 code flow .
 * We require the user to provide ?user_id=... to /login so we can deterministically
 * store tokens under that user_id. The LLM/tool must reuse this id on all calls.
 */

require("dotenv").config();
const express = require("express");
const axios = require("axios");
const qs = require("qs");
const tokenStore = require("./tokenStore");

const router = express.Router();

// Env
const CLIENT_ID    = process.env.CLIENT_ID;
const TENANT_ID    = process.env.TENANT_ID;
const REDIRECT_URI = process.env.REDIRECT_URI || "http://localhost:3001/auth/callback";
const SCOPES       = (process.env.SCOPES || "Mail.ReadWrite").trim();

const AUTH_URL  = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/authorize`;
const TOKEN_URL = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;

/**
 * GET /login?user_id=<id>
 * - Opens Microsoft sign-in page for the given user_id (state).
 */
router.get("/login", (req, res) => {
  const user_id = (req.query.user_id || "").toString().trim();
  if (!user_id) {
    return res.status(400).json({ error: "Missing user_id. Call /execute_tool first to get login_url + user_id." });
  }

  const params = new URLSearchParams({
    client_id: CLIENT_ID,
    response_type: "code",
    redirect_uri: REDIRECT_URI,
    response_mode: "query",
    scope: SCOPES,
    state: user_id,
    prompt: "select_account"
  });

  res.redirect(`${AUTH_URL}?${params.toString()}`);
});

/**
 * GET /auth/callback?code=...&state=<user_id>[&format=json]
 * - Microsoft redirects here after successful login.
 * - We exchange the auth code for tokens and persist under `state` (the user_id).
 * - Returns a small success HTML so users can close the tab, or JSON if format=json.
 */
router.get("/auth/callback", async (req, res) => {
  const { code, state: user_id, error, error_description, format } = req.query || {};
  if (error) return res.status(400).send(`OAuth error: ${error} - ${error_description}`);
  if (!code || !user_id) return res.status(400).send("Missing authorization code or state (user_id).");

  try {
    const data = qs.stringify({
      client_id: CLIENT_ID,
      scope: SCOPES,
      code,
      redirect_uri: REDIRECT_URI,
      grant_type: "authorization_code"
    });

    const tokenResp = await axios.post(TOKEN_URL, data, {
      headers: { "Content-Type": "application/x-www-form-urlencoded" }
    });

    const { access_token, refresh_token, expires_in } = tokenResp.data;
    tokenStore.set(user_id, {
      access_token,
      refresh_token: refresh_token || null,
      expiry: Date.now() + (expires_in * 1000),
      scopes: SCOPES
    });

    if ((format || "").toLowerCase() === "json") {
      return res.json({ status: "logged_in", user_id, expires_in });
    }

   
    const minutes = Math.max(1, Math.floor(expires_in / 60));
    return res
      .status(200)
      .type("html")
      .send(`<!doctype html>
<html lang="en"><head>
<meta charset="utf-8" />
<title>Signed in</title>
<meta name="viewport" content="width=device-width, initial-scale=1" />
<style>
  body { font-family: system-ui, -apple-system, Segoe UI, Roboto, sans-serif; background:#0b1220; color:#e6edf3; display:flex; align-items:center; justify-content:center; height:100vh; margin:0;}
  .card { background:#111827; border:1px solid #1f2937; border-radius:14px; padding:24px; max-width:540px; box-shadow: 0 10px 30px rgba(0,0,0,.35);}
  .title { font-size:20px; font-weight:700; margin-bottom:8px; }
  .sub { color:#9ca3af; margin-bottom:16px; }
  code { background:#0b1220; border:1px solid #1f2937; padding:2px 6px; border-radius:8px; font-size:12px; }
  .row { margin:10px 0; }
  .btn { background:#2563eb; color:white; border:0; padding:10px 14px; border-radius:10px; cursor:pointer; }
  .btn:active { transform: translateY(1px); }
</style>
</head><body>
  <div class="card">
    <div class="title">You're signed in 🎉</div>
    <div class="sub">You can close this tab now.</div>
    <div class="row">Session active for about <strong>${minutes} minute${minutes === 1 ? "" : "s"}</strong>. After that you may need to sign in again.</div>
  </div>
<script>
  document.getElementById('copy').onclick = async () => {
    try { await navigator.clipboard.writeText(document.getElementById('uid').textContent); alert('User ID copied!'); }
    catch { alert('Copy failed.'); }
  };
  document.getElementById('close').onclick = () => window.close();
  try { window.opener && window.opener.postMessage({ type:'oauth_success', user_id: '${user_id}', expires_in: ${expires_in} }, '*'); } catch {}
</script>
</body></html>`);
  } catch (err) {
    console.error("OAuth callback error:", err.response?.data || err.message);
    return res.status(400).json({ error: err.response?.data || err.message });
  }
});

module.exports = { router, TOKEN_URL, CLIENT_ID, SCOPES, REDIRECT_URI };
