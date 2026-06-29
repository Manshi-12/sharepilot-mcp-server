import { getGraphClient } from "../auth/graphClient.js";
import { resolveDrive } from "../utils/resolve.js";

export const copyFileToolSchema = {
  name: "copy_file",
  description:
    "Copies a file from one SharePoint document library to another on the same site. " +
    "The original file stays in the source library. " +
    "Use when the user asks to copy a file or document to another library.",
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

export async function copyFile(
  fileId: string,
  sourceLibraryName: string,
  destinationLibraryName: string,
  fileName: string
) {
  const client = await getGraphClient();
  const sourceDrive = await resolveDrive(client, sourceLibraryName);
  const destDrive = await resolveDrive(client, destinationLibraryName);

  // Graph copy is async — it returns 202 Accepted with a monitor URL
  const res = await client.post(
    `/drives/${sourceDrive.id}/items/${fileId}/copy`,
    {
      parentReference: { driveId: destDrive.id, id: "root" },
      name: fileName,
    },
    { headers: { Prefer: "respond-async" } }
  );

  // 202 means copy started in background — confirm to user
  const monitorUrl = res.headers?.location || null;

  return {
    status: "copy_started",
    fileName,
    from: sourceLibraryName,
    to: destinationLibraryName,
    message:
      "The file copy has been started. Large files may take a moment to appear in the destination library.",
    monitorUrl,
  };
}