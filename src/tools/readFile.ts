import { getGraphClient } from "../auth/graphClient.js";

const SITE_ID = process.env.SITE_ID || "";

export const readFileToolSchema = {
  name: "read_file",
  description:
    "Downloads a file from the 'Company Knowledge Base' SharePoint Document Library by its " +
    "file ID and returns its raw content/metadata for further extraction by the AI agent. " +
    "Use search_file first to obtain the file ID.",
  inputSchema: {
    type: "object",
    properties: {
      fileId: {
        type: "string",
        description: "The SharePoint drive item ID of the file to read (from search_file).",
      },
    },
    required: ["fileId"],
  },
};

export async function readFile(fileId: string) {
  const client = await getGraphClient();

  const driveRes = await client.get(`/sites/${SITE_ID}/drive`);
  const driveId = driveRes.data.id;

  const itemRes = await client.get(`/drives/${driveId}/items/${fileId}`);
  const item = itemRes.data;

  return {
    id: item.id,
    name: item.name,
    downloadUrl: item["@microsoft.graph.downloadUrl"],
    mimeType: item.file?.mimeType,
    size: item.size,
  };
}