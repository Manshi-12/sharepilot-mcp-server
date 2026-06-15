import { getGraphClient } from "../auth/graphClient.js";

const SITE_ID = process.env.SITE_ID || "";

const KNOWN_DRIVES: Record<string, string> = {
  "company knowledge base": "b!_QyOrq2LXkiz1turmiRqpReoxLAcnd5AqBzBOABhT83c0dgJ0DkPQqqH8PQDZMHP",
  "documents": "b!_QyOrq2LXkiz1turmiRqpReoxLAcnd5AqBzBOABhT807pG5Xm1gIR4AxoEnqyICJ",
};

export const uploadFileToolSchema = {
  name: "upload_file",
  description:
    "Uploads a text file or creates a new document in a SharePoint Document Library. " +
    "Provide the filename, content as text, and target library name. " +
    "Available libraries: 'Company Knowledge Base', 'Documents'.",
  inputSchema: {
    type: "object",
    properties: {
      filename: {
        type: "string",
        description: "Name of the file to create, e.g. 'report.txt' or 'notes.md'.",
      },
      content: {
        type: "string",
        description: "Text content to write into the file.",
      },
      libraryName: {
        type: "string",
        description: "Target Document Library: 'Company Knowledge Base' or 'Documents'. Defaults to 'Company Knowledge Base'.",
      },
    },
    required: ["filename", "content"],
  },
};

export async function uploadFile(
  filename: string,
  content: string,
  libraryName: string = "Company Knowledge Base"
) {
  const client = await getGraphClient();

  const key = libraryName.toLowerCase();
  let driveId = KNOWN_DRIVES[key];

  if (!driveId) {
    const drivesRes = await client.get(`/sites/${SITE_ID}/drives`);
    const drives = drivesRes.data.value || [];
    const found = drives.find((d: any) => d.name.toLowerCase() === key);
    if (!found) throw new Error(`Library "${libraryName}" not found.`);
    driveId = found.id;
  }

  const res = await client.put(
    `/drives/${driveId}/root:/${encodeURIComponent(filename)}:/content`,
    content,
    {
      headers: {
        "Content-Type": "text/plain",
      },
    }
  );

  return {
    id: res.data.id,
    name: res.data.name,
    webUrl: res.data.webUrl,
    size: res.data.size,
    libraryName: libraryName,
    status: "uploaded",
  };
}