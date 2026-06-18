import { getGraphClient } from "../auth/graphClient.js";
import { resolveDrive, clearResolverCache } from "../utils/resolve.js";

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
    // Delete a single file
    await client.delete(`/drives/${drive.id}/items/${fileId}`);
    return {
      success: true,
      deleted: "file",
      fileId,
      libraryName: drive.name,
      message: `File was permanently deleted from "${drive.name}".`,
    };
  } else {
    // Delete the entire library — libraries are backed by a list in Graph
    // We delete the drive's root list which removes the library
    await client.delete(`/drives/${drive.id}`);
    clearResolverCache();
    return {
      success: true,
      deleted: "library",
      libraryName: drive.name,
      message: `The library "${drive.name}" and all its files were permanently deleted.`,
    };
  }
}