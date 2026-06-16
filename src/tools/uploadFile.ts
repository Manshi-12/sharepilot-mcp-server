import { getGraphClient } from "../auth/graphClient.js";
import { resolveDrive } from "../utils/resolve.js";

const EXTENSION_MIME: Record<string, string> = {
  ".txt": "text/plain",
  ".md": "text/markdown",
  ".csv": "text/csv",
  ".json": "application/json",
  ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  ".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
  ".pdf": "application/pdf",
  ".png": "image/png",
  ".jpg": "image/jpeg",
  ".jpeg": "image/jpeg",
  ".gif": "image/gif",
};

function inferMimeType(filename: string): string {
  const ext = filename.slice(filename.lastIndexOf(".")).toLowerCase();
  return EXTENSION_MIME[ext] || "application/octet-stream";
}

export const uploadFileToolSchema = {
  name: "upload_file",
  description:
    "Uploads a file to any Document Library on the SharePoint site by display name. " +
    "For plain text content (e.g. notes, reports the agent generated), pass it directly " +
    "as text and leave isBase64 false. For a real binary file the user attached " +
    "(.docx, .pdf, .png, .xlsx, etc.), pass its base64-encoded content and set isBase64 " +
    "to true so it isn't corrupted on upload.",
  inputSchema: {
    type: "object",
    properties: {
      filename: {
        type: "string",
        description: "Name to give the file, e.g. 'report.docx' or 'notes.txt'.",
      },
      content: {
        type: "string",
        description: "The file content — plain text, or base64 if isBase64 is true.",
      },
      libraryName: {
        type: "string",
        description: "Target Document Library to upload into, e.g. 'Company Knowledge Base'. Required — ask the user which library if they didn't say.",
      },
      isBase64: {
        type: "boolean",
        description: "Set true if 'content' is base64-encoded binary data rather than plain text. Defaults to false.",
      },
      mimeType: {
        type: "string",
        description: "Optional explicit MIME type. If omitted, it's inferred from the filename's extension.",
      },
    },
    required: ["filename", "content", "libraryName"],
  },
};

export async function uploadFile(
  filename: string,
  content: string,
  libraryName: string,
  isBase64: boolean = false,
  mimeType?: string
) {
  const client = await getGraphClient();
  const drive = await resolveDrive(client, libraryName);

  const buffer = isBase64 ? Buffer.from(content, "base64") : Buffer.from(content, "utf-8");
  const contentType = mimeType || inferMimeType(filename);

  const res = await client.put(
    `/drives/${drive.id}/root:/${encodeURIComponent(filename)}:/content`,
    buffer,
    { headers: { "Content-Type": contentType } }
  );

  return {
    id: res.data.id,
    name: res.data.name,
    webUrl: res.data.webUrl,
    size: res.data.size,
    libraryName: drive.name,
    status: "uploaded",
  };
}
