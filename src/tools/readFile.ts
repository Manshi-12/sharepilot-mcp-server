import { getGraphClient } from "../auth/graphClient.js";

export const readFileToolSchema = {
  name: "read_file",
  description:
    "Downloads a file from the 'Company Knowledge Base' SharePoint Document Library by its " +
    "file ID and returns its content/metadata for further extraction by the AI agent. " +
    "Use search_file first to obtain the fileId and driveId.",
  inputSchema: {
    type: "object",
    properties: {
      fileId: {
        type: "string",
        description: "The SharePoint drive item ID of the file to read (from search_file result).",
      },
      driveId: {
        type: "string",
        description: "The drive ID containing the file (from search_file result).",
      },
    },
    required: ["fileId", "driveId"],
  },
};

export async function readFile(fileId: string, driveId: string) {
  const client = await getGraphClient();

  const itemRes = await client.get(`/drives/${driveId}/items/${fileId}`);
  const item = itemRes.data;

  return {
    id: item.id,
    name: item.name,
    downloadUrl: item["@microsoft.graph.downloadUrl"],
    mimeType: item.file?.mimeType,
    size: item.size,
    webUrl: item.webUrl,
    lastModifiedDateTime: item.lastModifiedDateTime,
  };
}