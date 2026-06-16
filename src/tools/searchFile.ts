import { getGraphClient } from "../auth/graphClient.js";
import { getAllDrives, resolveDrive } from "../utils/resolve.js";

export const searchFileToolSchema = {
  name: "search_file",
  description:
    "Searches Document Libraries on the SharePoint site for a file by name. " +
    "Works with any library that exists on the site — pass libraryName to search one " +
    "specific library, or omit it to search every library on the site. " +
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
        description: "Name of the Document Library to search. If omitted, searches all libraries on the site.",
      },
    },
    required: ["filename"],
  },
};

export async function searchFile(filename: string, libraryName?: string) {
  const client = await getGraphClient();

  const driveEntries = libraryName
    ? [await resolveDrive(client, libraryName)]
    : await getAllDrives(client);

  const allFiles: any[] = [];
  const searchLower = filename.toLowerCase();

  for (const drive of driveEntries) {
    try {
      // Listing root children and filtering client-side is more reliable for
      // app-only (client credentials) auth than the /root/search() endpoint.
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
      } catch (_) {
        // Skip a library we genuinely can't read rather than failing the whole search.
      }
    }
  }

  return { matchCount: allFiles.length, files: allFiles };
}
