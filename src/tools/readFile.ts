import { getGraphClient } from "../auth/graphClient.js";
import axios from "axios";
import mammoth from "mammoth";
import * as XLSX from "xlsx";

export const readFileToolSchema = {
  name: "read_file",
  description:
    "Reads a file from a SharePoint Document Library and returns its metadata, download URL, and extracted text content. " +
    "Supports .txt, .csv, .json, .md, .docx, and .xlsx files. " +
    "Use search_file first to get the fileId and driveId.",
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

async function extractTextContent(
  downloadUrl: string,
  fileName: string,
  mimeType: string
): Promise<string | null> {
  const name = fileName.toLowerCase();

  try {
    // ── Plain text types ───────────────────────────────────────────────────
    if (
      mimeType.includes("text") ||
      mimeType.includes("json") ||
      mimeType.includes("csv") ||
      name.endsWith(".txt") ||
      name.endsWith(".csv") ||
      name.endsWith(".json") ||
      name.endsWith(".md")
    ) {
      const res = await axios.get(downloadUrl, { responseType: "text", timeout: 15000 });
      const raw = typeof res.data === "string" ? res.data : JSON.stringify(res.data);
      return raw.slice(0, 8000);
    }

    // ── DOCX ──────────────────────────────────────────────────────────────
    if (name.endsWith(".docx") || mimeType.includes("wordprocessingml")) {
      const res = await axios.get(downloadUrl, { responseType: "arraybuffer", timeout: 15000 });
      const result = await mammoth.extractRawText({ buffer: Buffer.from(res.data) });
      return result.value.slice(0, 8000);
    }

    // ── XLSX ──────────────────────────────────────────────────────────────
    if (name.endsWith(".xlsx") || name.endsWith(".xls") || mimeType.includes("spreadsheetml")) {
      const res = await axios.get(downloadUrl, { responseType: "arraybuffer", timeout: 15000 });
      const workbook = XLSX.read(res.data, { type: "buffer" });
      const lines: string[] = [];
      for (const sheetName of workbook.SheetNames) {
        const sheet = workbook.Sheets[sheetName];
        const csv = XLSX.utils.sheet_to_csv(sheet);
        lines.push(`### Sheet: ${sheetName}\n${csv}`);
      }
      return lines.join("\n\n").slice(0, 8000);
    }

  } catch {
    // Non-fatal — return null and let caller handle gracefully
  }

  return null;
}

export async function readFile(fileId: string, driveId: string) {
  const client = await getGraphClient();

  const itemRes = await client.get(`/drives/${driveId}/items/${fileId}`);
  const item = itemRes.data;

  const downloadUrl = item["@microsoft.graph.downloadUrl"];
  const mimeType = item.file?.mimeType || "";

  const textContent = await extractTextContent(downloadUrl, item.name, mimeType);

  return {
    id: item.id,
    name: item.name,
    downloadUrl,
    mimeType,
    size: item.size,
    webUrl: item.webUrl,
    lastModifiedDateTime: item.lastModifiedDateTime,
    textContent,
  };
}