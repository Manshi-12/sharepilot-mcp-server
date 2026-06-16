import express from "express";
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

dotenv.config();

const PORT = process.env.PORT ? parseInt(process.env.PORT, 10) : 3000;

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

  app.get("/", (_req, res) => {
    res.json({ status: "ok", service: "sharepilot-mcp-server", tools: 5 });
    res.json({ status: "ok", service: "sharepilot-mcp-server", tools: 6 });
  });

  app.post("/mcp", async (req, res) => {
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
  });
}

main().catch((err) => {
  console.error("Fatal error:", err);
  process.exit(1);
});