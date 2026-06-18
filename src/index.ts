import express, { Request, Response, NextFunction } from "express";
import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
} from "@modelcontextprotocol/sdk/types.js";
import dotenv from "dotenv";

import { searchFile, searchFileToolSchema } from "./tools/searchFile.js";
import { readFile, readFileToolSchema } from "./tools/readFile.js";
import { createListItem, createListItemToolSchema } from "./tools/createListItem.js";
import { uploadFile, uploadFileToolSchema } from "./tools/uploadFile.js";
import { getListItems, getListItemsToolSchema } from "./tools/getListItems.js";
import { uploadListItemImage, uploadListItemImageToolSchema } from "./tools/uploadListItemImage.js";
import { createList, createListToolSchema } from "./tools/createList.js";
import { updateListItem, updateListItemToolSchema } from "./tools/updateListItem.js";
import { deleteListItem, deleteListItemToolSchema } from "./tools/deleteListItem.js";
import { deleteFile, deleteFileToolSchema } from "./tools/deleteFile.js";

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
    tools: [
      searchFileToolSchema,
      readFileToolSchema,
      createListItemToolSchema,
      uploadFileToolSchema,
      getListItemsToolSchema,
      uploadListItemImageToolSchema,
      createListToolSchema,
      updateListItemToolSchema,
      deleteListItemToolSchema,
      deleteFileToolSchema,
    ],
  }));

  server.setRequestHandler(CallToolRequestSchema, async (request) => {
    const { name, arguments: args } = request.params;

    try {
      switch (name) {
        case "search_file": {
          const result = await searchFile(
            (args as any).filename,
            (args as any).libraryName
          );
          return { content: [{ type: "text", text: JSON.stringify(result, null, 2) }] };
        }
        case "read_file": {
          const result = await readFile(
            (args as any).fileId,
            (args as any).driveId
          );
          return { content: [{ type: "text", text: JSON.stringify(result, null, 2) }] };
        }
        case "create_list_item": {
          const result = await createListItem(
            (args as any).listName,
            (args as any).fields
          );
          return { content: [{ type: "text", text: JSON.stringify(result, null, 2) }] };
        }
        case "upload_file": {
          const result = await uploadFile(
            (args as any).filename,
            (args as any).content,
            (args as any).libraryName,
            (args as any).isBase64,
            (args as any).mimeType
          );
          return { content: [{ type: "text", text: JSON.stringify(result, null, 2) }] };
        }
        case "get_list_items": {
          const result = await getListItems(
            (args as any).listName,
            (args as any).search,
            (args as any).top
          );
          return { content: [{ type: "text", text: JSON.stringify(result, null, 2) }] };
        }
        case "upload_list_item_image": {
          const result = await uploadListItemImage(
            (args as any).listName,
            (args as any).itemId,
            (args as any).imageFieldName,
            (args as any).fileName,
            (args as any).base64Content,
            (args as any).mimeType
          );
          return { content: [{ type: "text", text: JSON.stringify(result, null, 2) }] };
        }
        case "create_list": {
          const result = await createList(
            (args as any).displayName,
            (args as any).template,
            (args as any).description,
            (args as any).columns
          );
          return { content: [{ type: "text", text: JSON.stringify(result, null, 2) }] };
        }
        case "update_list_item": {
          const result = await updateListItem(
            (args as any).listName,
            (args as any).itemId,
            (args as any).fields
          );
          return { content: [{ type: "text", text: JSON.stringify(result, null, 2) }] };
        }
        case "delete_list_item": {
          const result = await deleteListItem(
            (args as any).listName,
            (args as any).itemId
          );
          return { content: [{ type: "text", text: JSON.stringify(result, null, 2) }] };
        }
        case "delete_file": {
          const result = await deleteFile(
            (args as any).libraryName,
            (args as any).fileId
          );
          return { content: [{ type: "text", text: JSON.stringify(result, null, 2) }] };
        }
        default:
          throw new Error(`Unknown tool: ${name}`);
      }
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
