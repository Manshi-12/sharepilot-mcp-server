import { getGraphClient } from "../auth/graphClient.js";
import { resolveDrive } from "../utils/resolve.js";

export const moveFileToolSchema = {
  name: "move_file",
  description:
    "Moves a file from one SharePoint document library to another on the same site. " +
    "The file is removed from the source library and placed in the destination library. " +
    "Use when the user asks to move a file or document to a different library.",
  inputSchema: {
    type: "object",
    properties: {
      fileId: { type: "string", description: "The file ID returned by search_file." },
      sourceLibraryName: { type: "string", description: "Display name of the source document library." },
      destinationLibraryName: { type: "string", description: "Display name of the destination document library." },
      fileName: { type: "string", description: "The file name including extension, e.g. 'report.docx'." },
    },
    required: ["fileId", "sourceLibraryName", "destinationLibraryName", "fileName"],
  },
};

export async function moveFile(
  fileId: string,
  sourceLibraryName: string,
  destinationLibraryName: string,
  fileName: string
) {
  const client = await getGraphClient();
  const sourceDrive = await resolveDrive(client, sourceLibraryName);
  const destDrive = await resolveDrive(client, destinationLibraryName);

  const res = await client.patch(
    `/drives/${sourceDrive.id}/items/${fileId}`,
    {
      parentReference: { driveId: destDrive.id, id: "root" },
      name: fileName,
    }
  );

  return {
    status: "moved",
    fileName: res.data.name,
    from: sourceLibraryName,
    to: destinationLibraryName,
    webUrl: res.data.webUrl,
  };
}