import { getGraphClient } from "../auth/graphClient.js";

const SITE_ID = process.env.SITE_ID || "";

export const searchFileToolSchema = {
  name: "search_file",
  description:
    "Searches a specific Document Library on the SharePoint site for a file by name. " +
    "Can search any library on the site. Returns file metadata including id and driveId needed for read_file.",
  inputSchema: {
    type: "object",
    properties: {
      filename: {
        type: "string",
        description: "Full or partial name of the file to search for, e.g. 'HR_Policy.docx'",
      },
      libraryName: {
        type: "string",
        description: "Name of the Document Library to search in, e.g. 'Company Knowledge Base', 'Documents'. If not specified, searches all libraries.",
      },
    },
    required: ["filename"],
  },
};

export async function searchFile(filename: string, libraryName?: string) {
  const client = await getGraphClient();

  const drivesRes = await client.get(`/sites/${SITE_ID}/drives`);
  const drives = drivesRes.data.value || [];

  const targetDrives = libraryName
    ? drives.filter((d: any) =>
        d.name.toLowerCase() === libraryName.toLowerCase() ||
        d.name.toLowerCase().replace(/\s/g, "") === libraryName.toLowerCase().replace(/\s/g, "")
      )
    : drives;

  const allFiles: any[] = [];

  for (const drive of targetDrives) {
    try {
      const searchRes = await client.get(
        `/drives/${drive.id}/root/search(q='${encodeURIComponent(filename)}')`
      );
      const items = (searchRes.data.value || []).map((item: any) => ({
        id: item.id,
        name: item.name,
        webUrl: item.webUrl,
        lastModifiedDateTime: item.lastModifiedDateTime,
        size: item.size,
        driveId: drive.id,
        libraryName: drive.name,
      }));
      allFiles.push(...items);
    } catch (_) {}
  }

  return { matchCount: allFiles.length, files: allFiles };
}