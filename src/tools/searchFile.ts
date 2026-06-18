import { getGraphClient } from "../auth/graphClient.js";
import { getAllDrives, resolveDrive } from "../utils/resolve.js";

export const searchFileToolSchema = {
  name: "search_file",
  description:
    "Searches Document Libraries on the SharePoint site for a file by name. " +
    "Works with any library that exists on the site — pass libraryName to search one " +
    "specific library, or omit it to search every library on the site. " +
    "Searches recursively through all folders and subfolders. " +
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

/**
 * Recursively walks a folder and collects all items whose name matches searchLower.
 * Fix #6 — old code only checked root/children, missing anything inside subfolders.
 */
async function searchFolderRecursive(
  client: any,
  driveId: string,
  driveName: string,
  folderPath: string,
  searchLower: string,
  depth: number = 0
): Promise<any[]> {
  // Guard against infinite recursion on circular folder structures (shouldn't happen
  // in SharePoint but defensive programming is good)
  if (depth > 10) return [];

  let items: any[] = [];

  try {
    const url = folderPath === "/"
      ? `/drives/${driveId}/root/children`
      : `/drives/${driveId}/root:${folderPath}:/children`;

    const res = await client.get(url);
    const children = res.data.value || [];

    for (const item of children) {
      if (item.folder) {
        // It's a folder — recurse into it
        const subPath = folderPath === "/"
          ? `/${item.name}`
          : `${folderPath}/${item.name}`;
        const nested = await searchFolderRecursive(
          client, driveId, driveName, subPath, searchLower, depth + 1
        );
        items = items.concat(nested);
      } else if (item.name?.toLowerCase().includes(searchLower)) {
        items.push({
          id: item.id,
          name: item.name,
          webUrl: item.webUrl,
          lastModifiedDateTime: item.lastModifiedDateTime,
          size: item.size,
          driveId,
          libraryName: driveName,
          // Include folder path so user knows where the file lives
          folderPath: folderPath === "/" ? "/" : folderPath,
        });
      }
    }
  } catch {
    // Skip folders we can't read rather than failing the whole search
  }

  return items;
}

export async function searchFile(filename: string, libraryName?: string) {
  const client = await getGraphClient();

  const driveEntries = libraryName
    ? [await resolveDrive(client, libraryName)]
    : await getAllDrives(client);

  const allFiles: any[] = [];
  const searchLower = filename.toLowerCase();

  for (const drive of driveEntries) {
    try {
      const results = await searchFolderRecursive(
        client, drive.id, drive.name, "/", searchLower
      );
      allFiles.push(...results);
    } catch {
      // If recursive walk fails entirely, fall back to Graph's search endpoint
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
      } catch {
        // Skip library entirely if both approaches fail
      }
    }
  }

  return { matchCount: allFiles.length, files: allFiles };
}
