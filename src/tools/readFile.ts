import { getGraphClient } from "../auth/graphClient.js";
import axios from "axios";

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

  let textContent: string | null = null;

  const downloadUrl = item["@microsoft.graph.downloadUrl"];
  const mimeType = item.file?.mimeType || "";

  if (downloadUrl && (
    mimeType.includes("text") ||
    mimeType.includes("json") ||
    mimeType.includes("csv") ||
    item.name?.endsWith(".txt") ||
    item.name?.endsWith(".csv") ||
    item.name?.endsWith(".json") ||
    item.name?.endsWith(".md")
  )) {
    try {
      const contentRes = await axios.get(downloadUrl, { responseType: "text", timeout: 15000 });
      textContent = typeof contentRes.data === "string"
        ? contentRes.data.slice(0, 8000) // cap at 8000 chars so model context isn't flooded
        : JSON.stringify(contentRes.data).slice(0, 8000);
    } catch {
      textContent = null; // non-fatal — metadata still returned
    }
  }

  return {
    id: item.id,
    name: item.name,
    downloadUrl: item["@microsoft.graph.downloadUrl"],
    mimeType: item.file?.mimeType,
    size: item.size,
    webUrl: item.webUrl,
    lastModifiedDateTime: item.lastModifiedDateTime,
    textContent,
  };
}