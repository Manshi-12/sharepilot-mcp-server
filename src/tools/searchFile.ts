import { getGraphClient } from "../auth/graphClient.js";

const SITE_ID = process.env.SITE_ID || "";

const KNOWN_DRIVES: Record<string, string> = {
  "company knowledge base": "b!_QyOrq2LXkiz1turmiRqpReoxLAcnd5AqBzBOABhT83c0dgJ0DkPQqqH8PQDZMHP",
  "documents": "b!_QyOrq2LXkiz1turmiRqpReoxLAcnd5AqBzBOABhT807pG5Xm1gIR4AxoEnqyICJ",
};

export const searchFileToolSchema = {
  name: "search_file",
  description:
    "Searches a Document Library on the SharePoint site for a file by name. " +
    "Available libraries: 'Company Knowledge Base', 'Documents'. " +
    "Returns file metadata including id and driveId needed for read_file.",
  inputSchema: {
    type: "object",
    properties: {
      filename: {
        type: "string",
        description: "Full or partial name of the file to search for, e.g. 'HR_Policy.docx'",
      },
      libraryName: {
        type: "string",
        description: "Name of the Document Library: 'Company Knowledge Base' or 'Documents'. If not specified, searches all libraries.",
      },
    },
    required: ["filename"],
  },
};

export async function searchFile(filename: string, libraryName?: string) {
  const client = await getGraphClient();

  let driveEntries: { name: string; id: string }[] = [];

  if (libraryName) {
    const key = libraryName.toLowerCase();
    const driveId = KNOWN_DRIVES[key];
    if (driveId) {
      driveEntries = [{ name: libraryName, id: driveId }];
    } else {
      const drivesRes = await client.get(`/sites/${SITE_ID}/drives`);
      const drives = drivesRes.data.value || [];
      const found = drives.find((d: any) => d.name.toLowerCase() === key);
      if (found) driveEntries = [{ name: found.name, id: found.id }];
    }
  } else {
    driveEntries = Object.entries(KNOWN_DRIVES).map(([name, id]) => ({ name, id }));
  }

  const allFiles: any[] = [];
  const searchLower = filename.toLowerCase();

  for (const drive of driveEntries) {
    try {
      // List all files in root, filter by name client-side
      const listRes = await client.get(`/drives/${drive.id}/root/children`);
      const items = (listRes.data.value || [])
        .filter((item: any) => item.name?.toLowerCase().includes(searchLower))
        .map((item: any) => ({
          id: item.id,
          name: item.name,
          webUrl: item.webUrl,
          lastModifiedDateTime: item.lastModifiedDateTime,
          size: item.size,
          driveId: drive.id,
          libraryName: drive.name,
        }));
      allFiles.push(...items);
    } catch (e: any) {
      // Fallback to search API
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
  }

  return { matchCount: allFiles.length, files: allFiles };
}