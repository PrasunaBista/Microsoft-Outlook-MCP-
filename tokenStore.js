/**
 * tokenStore.js
 *
 * Tiny SQLite-backed token store. Keeps a strict 1:1 mapping of user_id → tokens.
 * I do NOT auto-refresh tokens here (no offline_access by default). When expired,
 * the API responds with { requires_login: true } and the same user_id so the client
 * can login again. This is simple, safe, and predictable.
 */

const fs = require("fs");
const path = require("path");
const Database = require("better-sqlite3");

const DATA_DIR = path.join(__dirname, "data");
const DB_PATH = path.join(DATA_DIR, "tokens.db");

if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });

const db = new Database(DB_PATH);
db.pragma("journal_mode = WAL");


db.exec(`
CREATE TABLE IF NOT EXISTS tokens (
  user_id TEXT PRIMARY KEY,
  access_token TEXT NOT NULL,
  refresh_token TEXT,
  expiry INTEGER NOT NULL,   -- epoch ms
  scopes TEXT,
  created_at INTEGER NOT NULL,
  updated_at INTEGER NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_tokens_expiry ON tokens(expiry);
`);

// Prepared statements
const upsertStmt = db.prepare(`
INSERT INTO tokens (user_id, access_token, refresh_token, expiry, scopes, created_at, updated_at)
VALUES (@user_id, @access_token, @refresh_token, @expiry, @scopes, @ts, @ts)
ON CONFLICT(user_id) DO UPDATE SET
  access_token = excluded.access_token,
  refresh_token = excluded.refresh_token,
  expiry = excluded.expiry,
  scopes = excluded.scopes,
  updated_at = excluded.updated_at
`);

const getStmt = db.prepare(`SELECT * FROM tokens WHERE user_id = ?`);
const delStmt = db.prepare(`DELETE FROM tokens WHERE user_id = ?`);
const delExpiredStmt = db.prepare(`DELETE FROM tokens WHERE expiry <= ?`);
const countStmt = db.prepare(`SELECT COUNT(*) as c FROM tokens`);

module.exports = {
  set(user_id, { access_token, refresh_token = null, expiry, scopes }) {
    const ts = Date.now();
    upsertStmt.run({
      user_id, access_token, refresh_token, expiry, scopes, ts
    });
  },
  get(user_id) {
    return getStmt.get(user_id) || null; // { user_id, access_token, ... }
  },
  delete(user_id) {
    delStmt.run(user_id);
  },
  deleteExpired() {
    return delExpiredStmt.run(Date.now()).changes;
  },
  count() {
    return countStmt.get().c || 0;
  },
  _db: db // exported for troubleshooting if needed
};


