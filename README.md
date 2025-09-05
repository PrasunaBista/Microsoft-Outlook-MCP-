Microsoft 365 Mail MCP Tool (Node.js)

A minimal MCP-style HTTP service that lets an LLM (or any client) read/search a Microsoft 365 mailbox using delegated OAuth.
It enforces a stable user_id per conversation/session so the LLM won’t keep asking to log in.

Auth flow: Microsoft OAuth (authorization code)

Token store: SQLite (file), no silent refresh by default

Endpoints: /execute_tool, /login, /auth/callback, /openapi.json

Client hinting: OpenAPI includes strict “client rules” so tools reuse the last user_id_used.
