import { getGraphClient } from "../auth/graphClient.js";
import { resolveDrive } from "../utils/resolve.js";

export const renameFileToolSchema = {
  name: "rename_file",
  description:
    "Renames a file in a SharePoint document library. " +
    "Use when the user asks to rename a file or change a file's name. " +
    "The file stays in the same library — only its name changes.",
  inputSchema: {
    type: "object",
    properties: {
      fileId: { type: "string", description: "The file ID returned by search_file." },
      libraryName: { type: "string", description: "Display name of the document library the file is in." },
      newFileName: { type: "string", description: "The new file name including extension, e.g. 'Q2_Report_Final.docx'." },
    },
    required: ["fileId", "libraryName", "newFileName"],
  },
};

export async function renameFile(fileId: string, libraryName: string, newFileName: string) {
  const client = await getGraphClient();
  const drive = await resolveDrive(client, libraryName);

  const res = await client.patch(
    `/drives/${drive.id}/items/${fileId}`,
    { name: newFileName }
  );

  return {
    status: "renamed",
    newFileName: res.data.name,
    libraryName: drive.name,
    webUrl: res.data.webUrl,
  };
}