import axios from "axios";
import { readFile } from "./readFile.js";

const AZURE_OPENAI_ENDPOINT = (process.env.AZURE_OPENAI_ENDPOINT || "").replace(/\/+$/, "");
const AZURE_OPENAI_API_KEY = process.env.AZURE_OPENAI_API_KEY || "";
const AZURE_OPENAI_DEPLOYMENT = process.env.AZURE_OPENAI_DEPLOYMENT || "";
const AZURE_OPENAI_API_VERSION = process.env.AZURE_OPENAI_API_VERSION || "2024-08-01-preview";

export const summarizeFileToolSchema = {
  name: "summarize_file",
  description:
    "Reads a file from a SharePoint Document Library and returns an AI-generated summary of its contents. " +
    "Use search_file first to get the fileId and driveId. Works best with text, CSV, JSON, DOCX, and XLSX files.",
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
      instruction: {
        type: "string",
        description:
          "Optional. A specific question or instruction about the file, e.g. 'List all employee names' or 'What is the total budget?'. If omitted, a general summary is produced.",
      },
    },
    required: ["fileId", "driveId"],
  },
};

export async function summarizeFile(
  fileId: string,
  driveId: string,
  instruction?: string
) {
  // Step 1 — reuse readFile to get content
  const fileData = await readFile(fileId, driveId) as any;

  if (!fileData.textContent) {
    return {
      fileName: fileData.name,
      summary: null,
      message:
        "This file type cannot be read as text (e.g. images, PDFs, or unsupported binary formats). " +
        "Summary is not available.",
    };
  }

  // Step 2 — call Azure OpenAI to summarize
  if (!AZURE_OPENAI_ENDPOINT || !AZURE_OPENAI_API_KEY || !AZURE_OPENAI_DEPLOYMENT) {
    throw new Error("Azure OpenAI is not configured. Cannot generate summary.");
  }

  const prompt = instruction
    ? `The following is the content of a file named "${fileData.name}".\n\n${fileData.textContent}\n\nInstruction: ${instruction}`
    : `Summarize the following file named "${fileData.name}" in clear, concise bullet points.\n\n${fileData.textContent}`;

  const url = `${AZURE_OPENAI_ENDPOINT}/openai/deployments/${AZURE_OPENAI_DEPLOYMENT}/chat/completions?api-version=${AZURE_OPENAI_API_VERSION}`;

  const { data } = await axios.post(
    url,
    {
      messages: [
        { role: "system", content: "You are a helpful assistant that summarizes SharePoint documents accurately." },
        { role: "user", content: prompt },
      ],
      temperature: 0.2,
      max_tokens: 800,
    },
    {
      headers: {
        "Content-Type": "application/json",
        "api-key": AZURE_OPENAI_API_KEY,
      },
      timeout: 30000,
    }
  );

  const summary = data.choices?.[0]?.message?.content || "Could not generate summary.";

  return {
    fileName: fileData.name,
    mimeType: fileData.mimeType,
    size: fileData.size,
    summary,
    instruction: instruction || null,
  };
}