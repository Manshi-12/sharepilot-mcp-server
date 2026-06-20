import express, { Request, Response, NextFunction } from "express";
import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
} from "@modelcontextprotocol/sdk/types.js";
import dotenv from "dotenv";

import { TOOL_SCHEMAS, executeTool } from "./tools/registry.js";
import { runChatAgent, ChatMessage, AgentEvent } from "./chat/agent.js";

dotenv.config();

// ── Fix #16 — Validate all required env vars on startup ───────────────────────
const REQUIRED_ENV_VARS = ["TENANT_ID", "CLIENT_ID", "CLIENT_SECRET", "SITE_ID", "SITE_URL"];
const missingVars = REQUIRED_ENV_VARS.filter((v) => !process.env[v]);
if (missingVars.length > 0) {
  console.error(
    `[SharePilot] FATAL: Missing required environment variables: ${missingVars.join(", ")}. ` +
    `Set them in .env (local) or Azure App Service → Configuration → Application Settings.`
  );
  process.exit(1);
}

const PORT = process.env.PORT ? parseInt(process.env.PORT, 10) : 3000;

// ── Fix #1 — Shared secret auth on /mcp endpoint ─────────────────────────────
// Set MCP_SECRET in your Azure App Service env vars and in your Foundry MCP config.
// If MCP_SECRET is not set at all, the check is skipped (dev/local mode).
const MCP_SECRET = process.env.MCP_SECRET || "";

function mcpAuthMiddleware(req: Request, res: Response, next: NextFunction): void {
  if (!MCP_SECRET) {
    // No secret configured — skip check (local dev mode)
    next();
    return;
  }
  const incoming = req.headers["x-mcp-secret"] || req.headers["authorization"];
  const token = typeof incoming === "string"
    ? incoming.replace(/^Bearer\s+/i, "")
    : "";

  if (token !== MCP_SECRET) {
    res.status(401).json({ error: "Unauthorized" });
    return;
  }
  next();
}

function createMcpServer(): Server {
  const server = new Server(
    { name: "sharepilot-mcp-server", version: "0.1.0" },
    { capabilities: { tools: {} } }
  );

  server.setRequestHandler(ListToolsRequestSchema, async () => ({
    tools: TOOL_SCHEMAS,
  }));

  server.setRequestHandler(CallToolRequestSchema, async (request) => {
    const { name, arguments: args } = request.params;

    try {
      const result = await executeTool(name, args as any);
      return { content: [{ type: "text", text: JSON.stringify(result, null, 2) }] };
    } catch (err: any) {
      const graphError = err?.response?.data
        ? JSON.stringify(err.response.data)
        : err.message || String(err);
      const status = err?.response?.status ? ` (HTTP ${err.response.status})` : "";
      return {
        content: [{ type: "text", text: `Error executing tool "${name}"${status}: ${graphError}` }],
        isError: true,
      };
    }
  });

  return server;
}

async function main() {
  const app = express();
  app.use(express.json());

  // Fix #12 — correct tool count in health check
  app.get("/", (_req, res) => {
    res.json({ status: "ok", service: "sharepilot-mcp-server", tools: 10 });
  });

  // Fix #1 — auth middleware applied only to /mcp
  app.post("/mcp", mcpAuthMiddleware, async (req, res) => {
    try {
      const server = createMcpServer();
      const transport = new StreamableHTTPServerTransport({
        sessionIdGenerator: undefined,
      });
      res.on("close", () => { transport.close(); server.close(); });
      await server.connect(transport);
      await transport.handleRequest(req, res, req.body);
    } catch (err: any) {
      console.error("Error handling MCP request:", err);
      if (!res.headersSent) {
        res.status(500).json({
          jsonrpc: "2.0",
          error: { code: -32603, message: "Internal server error" },
          id: null,
        });
      }
    }
  });

  app.get("/mcp", (_req, res) => {
    res.status(405).json({ error: "Method not allowed" });
  });

  // ── /chat — streaming REST endpoint for the Next.js frontend ────────────────
  // The frontend doesn't speak the MCP transport protocol, so this wraps the
  // same tools in a plain SSE stream: POST { messages } -> a stream of
  // { type: "tool_call" | "tool_result" | "token" | "usage" | "done", ... }
  // events, ending the HTTP response when the agent loop finishes.
  // Protected by the same shared secret as /mcp.
  app.post("/chat", mcpAuthMiddleware, async (req: Request, res: Response) => {
    const messages: ChatMessage[] = req.body?.messages;
    if (!Array.isArray(messages) || messages.length === 0) {
      res.status(400).json({ error: "messages array is required." });
      return;
    }

    res.writeHead(200, {
      "Content-Type": "text/event-stream",
      "Cache-Control": "no-cache, no-transform",
      Connection: "keep-alive",
      "X-Accel-Buffering": "no", // disable proxy buffering (App Service/nginx)
    });

    const send = (event: AgentEvent) => {
      res.write(`data: ${JSON.stringify(event)}\n\n`);
    };

    try {
      await runChatAgent(messages, send);
    } catch (err: any) {
      const message = err?.message || "Internal server error";
      send({ type: "done", content: `⚠️ ${message}` } as AgentEvent);
    } finally {
      res.end();
    }
  });

  app.listen(PORT, () => {
    console.log(`SharePilot MCP server listening on port ${PORT}`);
    console.log(`MCP endpoint: POST http://localhost:${PORT}/mcp`);
    if (MCP_SECRET) {
      console.log(`MCP endpoint is protected — x-mcp-secret header required.`);
    } else {
      console.log(`WARNING: MCP_SECRET not set — endpoint is unprotected. Set it in production.`);
    }
  });
}

main().catch((err) => {
  console.error("Fatal error:", err);
  process.exit(1);
});
