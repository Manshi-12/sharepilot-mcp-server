import { getGraphClient } from "../auth/graphClient.js";

export const readFileToolSchema = {
  name: "read_file",
  description:
    "Reads a file from a SharePoint Document Library and returns its metadata and download URL. " +
    "Use search_file first to get the fileId and driveId. Works with any library on the site.",
  inputSchema: {
    type: "object",
    properties: {
      fileId: {
        type: "string",
        description: "The SharePoint drive item ID of the file (from search_file result).",
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