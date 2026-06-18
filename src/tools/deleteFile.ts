import { getGraphClient } from "../auth/graphClient.js";
import { resolveDrive, getAllLists, clearResolverCache } from "../utils/resolve.js";

const SITE_ID = process.env.SITE_ID || "";

export const deleteFileToolSchema = {
  name: "delete_file",
  description:
    "Deletes a specific file from a SharePoint Document Library, OR deletes an entire Document Library. " +
    "To delete a single file: provide libraryName + fileId (get fileId from search_file). " +
    "To delete the entire library: provide libraryName only (leave fileId empty). " +
    "WARNING: Deleting a library is permanent and removes ALL files inside it.",
  inputSchema: {
    type: "object",
    properties: {
      libraryName: {
        type: "string",
        description: "Display name of the Document Library, e.g. 'Company Knowledge Base'.",
      },
      fileId: {
        type: "string",
        description:
          "The drive item ID of the file to delete (from search_file result). " +
          "Leave this out to delete the ENTIRE library.",
      },
    },
    required: ["libraryName"],
  },
};

export async function deleteFile(libraryName: string, fileId?: string) {
  const client = await getGraphClient();
  const drive = await resolveDrive(client, libraryName);

  if (fileId) {
    // Delete a single file via drive item
    await client.delete(`/drives/${drive.id}/items/${fileId}`);
    return {
      success: true,
      deleted: "file",
      fileId,
      libraryName: drive.name,
      message: `File was permanently deleted from "${drive.name}".`,
    };
  } else {
    // Fix #10 — Graph does NOT support DELETE /drives/{id}.
    // Document libraries are backed by SharePoint lists — delete the backing list instead.
    const lists = await getAllLists(client);
    const backingList = lists.find(
      (l) => l.displayName === drive.name || l.name === drive.name
    );

    if (!backingList) {
      throw new Error(
        `Could not find the backing list for library "${drive.name}". ` +
        `The library may be a system library that cannot be deleted this way.`
      );
    }

    await client.delete(`/sites/${SITE_ID}/lists/${backingList.id}`);
    clearResolverCache();

    return {
      success: true,
      deleted: "library",
      libraryName: drive.name,
      message: `The library "${drive.name}" and all its files were permanently deleted.`,
    };
  }
}
