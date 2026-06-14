import { getGraphClient } from "../auth/graphClient.js";

const SITE_ID = process.env.SITE_ID || "";

export const searchFileToolSchema = {
  name: "search_file",
  description:
    "Searches the 'Company Knowledge Base' SharePoint Document Library for a file by name " +
    "(or partial name) and returns basic metadata (id, name, webUrl, last modified) for matching files.",
  inputSchema: {
    type: "object",
    properties: {
      filename: {
        type: "string",
        description: "Full or partial name of the file to search for, e.g. 'project_brief.docx'",
      },
    },
    required: ["filename"],
  },
};

export async function searchFile(filename: string) {
  const client = await getGraphClient();

  const drivesRes = await client.get(`/sites/${SITE_ID}/drives`);
  const drives = drivesRes.data.value || [];

  const targetDrive = drives.find(
    (d: any) =>
      d.name === "Company Knowledge Base" ||
      d.name === "CompanyKnowledgeBase"
  ) || drives[0];

  const driveId = targetDrive.id;

  const searchRes = await client.get(
    `/drives/${driveId}/root/search(q='${encodeURIComponent(filename)}')`
  );

  const items = (searchRes.data.value || []).map((item: any) => ({
    id: item.id,
    name: item.name,
    webUrl: item.webUrl,
    lastModifiedDateTime: item.lastModifiedDateTime,
    size: item.size,
    driveId: driveId,
  }));

  return {
    matchCount: items.length,
    files: items,
  };
}