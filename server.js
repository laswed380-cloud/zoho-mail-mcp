import express from "express";
import { randomUUID, createHash } from "node:crypto";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";
import { ImapFlow } from "imapflow";
import nodemailer from "nodemailer";
import { z } from "zod";

const EMAIL = process.env.ZOHO_EMAIL;
const PASSWORD = process.env.ZOHO_APP_PASSWORD;
const REGION = (process.env.ZOHO_REGION || "in").toLowerCase();
const API_KEY = process.env.MCP_API_KEY;
const OAUTH_CLIENT_ID = process.env.OAUTH_CLIENT_ID || "claude-ai";
const OAUTH_CLIENT_SECRET = process.env.OAUTH_CLIENT_SECRET || API_KEY;
const PORT = parseInt(process.env.PORT || "3000", 10);

// Deterministic token — derived from secrets, survives server restarts
const STATIC_TOKEN = createHash("sha256")
  .update(`${OAUTH_CLIENT_ID}:${OAUTH_CLIENT_SECRET}:${API_KEY}`)
  .digest("hex");

// Auth codes are short-lived and in-memory (fine — codes are used once immediately)
const authCodes = new Map();

if (!EMAIL || !PASSWORD) {
  console.error("Missing ZOHO_EMAIL or ZOHO_APP_PASSWORD");
  process.exit(1);
}

const IMAP_HOST = `imap.zoho.${REGION}`;
const SMTP_HOST = `smtp.zoho.${REGION}`;

// ─── IMAP helpers ───────────────────────────────────────────────────────────

function imapClient() {
  return new ImapFlow({
    host: IMAP_HOST, port: 993, secure: true,
    auth: { user: EMAIL, pass: PASSWORD },
    logger: false,
  });
}

const smtpTransport = nodemailer.createTransport({
  host: SMTP_HOST, port: 465, secure: true,
  auth: { user: EMAIL, pass: PASSWORD },
});

async function withMailbox(folder, fn) {
  const client = imapClient();
  await client.connect();
  try {
    const lock = await client.getMailboxLock(folder);
    try { return await fn(client); }
    finally { lock.release(); }
  } finally {
    await client.logout().catch(() => client.close());
  }
}

function formatAddrs(addrs) {
  if (!addrs) return "";
  return addrs.map(a => a.name ? `${a.name} <${a.address}>` : a.address).join(", ");
}

function msgSummary(msg) {
  return {
    uid: msg.uid,
    flags: Array.from(msg.flags || []),
    from: formatAddrs(msg.envelope?.from),
    to: formatAddrs(msg.envelope?.to),
    cc: formatAddrs(msg.envelope?.cc),
    subject: msg.envelope?.subject || "(no subject)",
    date: msg.envelope?.date,
    messageId: msg.envelope?.messageId,
  };
}

// ─── MCP server factory ─────────────────────────────────────────────────────

function createMcpServer() {
  const server = new McpServer({ name: "zoho-mail", version: "1.0.0" });

  server.registerTool("list_folders", {
    title: "List mail folders",
    description: "List all folders in the Zoho mailbox.",
    inputSchema: {},
  }, async () => {
    const client = imapClient();
    await client.connect();
    try {
      const list = await client.list();
      return { content: [{ type: "text", text: JSON.stringify(list.map(f => ({ path: f.path, name: f.name, specialUse: f.specialUse })), null, 2) }] };
    } finally { await client.logout().catch(() => client.close()); }
  });

  server.registerTool("list_messages", {
    title: "List recent messages",
    description: "List the most recent messages from a folder (newest first). Returns headers only.",
    inputSchema: {
      folder: z.string().default("INBOX"),
      limit: z.number().int().min(1).max(100).default(20),
      unseenOnly: z.boolean().default(false),
    },
  }, async ({ folder, limit, unseenOnly }) => {
    return await withMailbox(folder, async (client) => {
      const uids = await client.search(unseenOnly ? { seen: false } : { all: true }, { uid: true });
      const recent = uids.slice(-limit).reverse();
      const out = [];
      for await (const msg of client.fetch(recent, { envelope: true, flags: true, uid: true }, { uid: true })) {
        out.push(msgSummary(msg));
      }
      return { content: [{ type: "text", text: JSON.stringify({ folder, count: out.length, total: uids.length, messages: out }, null, 2) }] };
    });
  });

  server.registerTool("read_message", {
    title: "Read message body",
    description: "Fetch the full body and headers of a message by UID.",
    inputSchema: {
      folder: z.string().default("INBOX"),
      uid: z.number().int().positive(),
      markRead: z.boolean().default(false),
    },
  }, async ({ folder, uid, markRead }) => {
    return await withMailbox(folder, async (client) => {
      const downloaded = await client.download(uid, undefined, { uid: true });
      const chunks = [];
      for await (const chunk of downloaded.content) chunks.push(chunk);
      const raw = Buffer.concat(chunks).toString("utf8");
      if (markRead) await client.messageFlagsAdd(uid, ["\\Seen"], { uid: true });
      return { content: [{ type: "text", text: raw.slice(0, 20000) }] };
    });
  });

  server.registerTool("search_messages", {
    title: "Search messages",
    description: "Search by from/to/subject/body/date/unseen.",
    inputSchema: {
      folder: z.string().default("INBOX"),
      from: z.string().optional(),
      to: z.string().optional(),
      subject: z.string().optional(),
      body: z.string().optional(),
      since: z.string().optional().describe("YYYY-MM-DD"),
      before: z.string().optional().describe("YYYY-MM-DD"),
      unseen: z.boolean().optional(),
      limit: z.number().int().min(1).max(100).default(25),
    },
  }, async ({ folder, from, to, subject, body, since, before, unseen, limit }) => {
    return await withMailbox(folder, async (client) => {
      const q = {};
      if (from) q.from = from;
      if (to) q.to = to;
      if (subject) q.subject = subject;
      if (body) q.body = body;
      if (since) q.since = new Date(since);
      if (before) q.before = new Date(before);
      if (unseen) q.seen = false;
      const uids = await client.search(q, { uid: true });
      const recent = uids.slice(-limit).reverse();
      const out = [];
      for await (const msg of client.fetch(recent, { envelope: true, flags: true, uid: true }, { uid: true })) {
        out.push(msgSummary(msg));
      }
      return { content: [{ type: "text", text: JSON.stringify({ matched: uids.length, messages: out }, null, 2) }] };
    });
  });

  server.registerTool("send_email", {
    title: "Send email",
    description: "Send an email via Zoho SMTP.",
    inputSchema: {
      to: z.union([z.string(), z.array(z.string())]),
      subject: z.string(),
      text: z.string().optional(),
      html: z.string().optional(),
      cc: z.union([z.string(), z.array(z.string())]).optional(),
      bcc: z.union([z.string(), z.array(z.string())]).optional(),
      replyTo: z.string().optional(),
      inReplyTo: z.string().optional(),
    },
  }, async (args) => {
    if (!args.text && !args.html) return { content: [{ type: "text", text: "Provide text or html body." }], isError: true };
    const info = await smtpTransport.sendMail({ from: EMAIL, ...args });
    return { content: [{ type: "text", text: JSON.stringify({ ok: true, messageId: info.messageId, accepted: info.accepted }, null, 2) }] };
  });

  server.registerTool("mark_message", {
    title: "Mark message",
    description: "Add/remove flags: \\Seen, \\Flagged, \\Answered.",
    inputSchema: {
      folder: z.string().default("INBOX"),
      uid: z.number().int().positive(),
      add: z.array(z.string()).optional(),
      remove: z.array(z.string()).optional(),
    },
  }, async ({ folder, uid, add, remove }) => {
    return await withMailbox(folder, async (client) => {
      if (add?.length) await client.messageFlagsAdd(uid, add, { uid: true });
      if (remove?.length) await client.messageFlagsRemove(uid, remove, { uid: true });
      return { content: [{ type: "text", text: JSON.stringify({ ok: true, uid, added: add, removed: remove }) }] };
    });
  });

  server.registerTool("move_message", {
    title: "Move message",
    description: "Move a message between folders.",
    inputSchema: {
      fromFolder: z.string().default("INBOX"),
      uid: z.number().int().positive(),
      toFolder: z.string(),
    },
  }, async ({ fromFolder, uid, toFolder }) => {
    return await withMailbox(fromFolder, async (client) => {
      const res = await client.messageMove(uid, toFolder, { uid: true });
      return { content: [{ type: "text", text: JSON.stringify({ ok: true, ...res }) }] };
    });
  });

  server.registerTool("delete_message", {
    title: "Delete message",
    description: "Delete a message (trash by default, permanent=true to expunge).",
    inputSchema: {
      folder: z.string().default("INBOX"),
      uid: z.number().int().positive(),
      permanent: z.boolean().default(false),
    },
  }, async ({ folder, uid, permanent }) => {
    return await withMailbox(folder, async (client) => {
      if (permanent) {
        await client.messageFlagsAdd(uid, ["\\Deleted"], { uid: true });
        await client.messageDelete(uid, { uid: true });
      } else {
        await client.messageMove(uid, "Trash", { uid: true }).catch(async () => {
          await client.messageFlagsAdd(uid, ["\\Deleted"], { uid: true });
        });
      }
      return { content: [{ type: "text", text: JSON.stringify({ ok: true, uid, permanent }) }] };
    });
  });

  return server;
}

// ─── Express HTTP server ────────────────────────────────────────────────────

const app = express();
app.set("trust proxy", 1); // trust Fly.io / Render reverse proxy for HTTPS
app.use(express.json());

// Auth middleware — accepts API key OR valid OAuth Bearer token
app.use((req, res, next) => {
  const skip = ["/health", "/", "/.well-known/oauth-authorization-server",
                "/oauth/authorize", "/oauth/token"];
  if (skip.some(p => req.path === p || req.path.startsWith(p))) return next();

  const header = req.headers.authorization || "";
  const apiKey = req.headers["x-api-key"];

  // Direct API key (for Claude Code / testing)
  if (API_KEY && apiKey === API_KEY) return next();

  // OAuth Bearer token — validate against deterministic static token
  const bearer = header.replace(/^Bearer\s+/i, "");
  if (bearer) {
    if (bearer === STATIC_TOKEN) return next();
    return res.status(401).json({ error: "invalid_token" });
  }

  res.status(401).json({ error: "unauthorized" });
});

app.get("/", (_req, res) => res.json({ ok: true, service: "zoho-mail-mcp" }));
app.get("/health", (_req, res) => res.json({ ok: true, email: EMAIL, region: REGION }));

// ─── OAuth 2.0 endpoints (required for Claude.ai MCP connector) ─────────────

// Discovery metadata
app.get("/.well-known/oauth-authorization-server", (req, res) => {
  const base = `${req.protocol}://${req.get("host")}`;
  res.json({
    issuer: base,
    authorization_endpoint: `${base}/oauth/authorize`,
    token_endpoint: `${base}/oauth/token`,
    response_types_supported: ["code"],
    grant_types_supported: ["authorization_code", "client_credentials"],
    code_challenge_methods_supported: ["S256", "plain"],
  });
});

// Authorization endpoint — auto-approves for this personal single-user server
app.get("/oauth/authorize", (req, res) => {
  const { client_id, redirect_uri, state, response_type } = req.query;
  if (client_id !== OAUTH_CLIENT_ID) return res.status(401).send("Unknown client");
  const code = randomUUID();
  authCodes.set(code, { clientId: client_id, expiresAt: Date.now() + 60_000 });
  const url = new URL(redirect_uri);
  url.searchParams.set("code", code);
  if (state) url.searchParams.set("state", state);
  res.redirect(url.toString());
});

// Token endpoint
app.post("/oauth/token", express.urlencoded({ extended: false }), (req, res) => {
  const { grant_type, code, client_id, client_secret, redirect_uri } = req.body;

  // Validate client
  if (client_id !== OAUTH_CLIENT_ID || client_secret !== OAUTH_CLIENT_SECRET) {
    return res.status(401).json({ error: "invalid_client" });
  }

  if (grant_type === "authorization_code") {
    const entry = authCodes.get(code);
    if (!entry || entry.expiresAt < Date.now()) {
      return res.status(400).json({ error: "invalid_grant" });
    }
    authCodes.delete(code);
    // Return deterministic token — survives server restarts
    return res.json({ access_token: STATIC_TOKEN, token_type: "bearer", expires_in: 315360000 });
  }

  if (grant_type === "client_credentials") {
    return res.json({ access_token: STATIC_TOKEN, token_type: "bearer", expires_in: 315360000 });
  }

  res.status(400).json({ error: "unsupported_grant_type" });
});

// Session store: sessionId -> { server, transport }
const sessions = new Map();

function getOrCreateSession(sessionId) {
  if (sessions.has(sessionId)) return sessions.get(sessionId);
  const mcpServer = createMcpServer();
  const transport = new StreamableHTTPServerTransport({
    sessionIdGenerator: () => sessionId,
    onsessioninitialized: (id) => {
      if (!sessions.has(id)) sessions.set(id, { mcpServer, transport });
    },
  });
  mcpServer.connect(transport);
  sessions.set(sessionId, { mcpServer, transport });
  return { mcpServer, transport };
}

// MCP endpoint — handles initialize (POST), SSE stream (GET), close (DELETE)
app.all("/mcp", async (req, res) => {
  try {
    const sessionId = req.headers["mcp-session-id"] || randomUUID();
    const { transport } = getOrCreateSession(sessionId);
    await transport.handleRequest(req, res, req.body);
  } catch (err) {
    console.error("MCP error:", err);
    if (!res.headersSent) res.status(500).json({ error: err.message });
  }
});

// Cleanup stale sessions every 30 min
setInterval(() => {
  for (const [id, { transport }] of sessions) {
    try { if (!transport.isConnected?.()) { transport.close().catch(() => {}); sessions.delete(id); } }
    catch { sessions.delete(id); }
  }
}, 30 * 60 * 1000);

app.listen(PORT, () => {
  console.log(`Zoho Mail MCP server running on port ${PORT}`);
  console.log(`Email: ${EMAIL} | Region: ${REGION} | Auth: ${API_KEY ? "enabled" : "DISABLED"}`);
});
