import { getGraphClient } from "../auth/graphClient.js";
import { resolveDrive } from "../utils/resolve.js";

export const getFileVersionsToolSchema = {
  name: "get_file_versions",
  description:
    "Returns the version history of a file in a SharePoint document library. " +
    "Shows each version's number, who last modified it, when, and the file size. " +
    "Use when the user asks about version history, previous versions, or who changed a file.",
  inputSchema: {
    type: "object",
    properties: {
      libraryName: { type: "string", description: "Display name of the document library." },
      fileId: { type: "string", description: "The file ID returned by search_file." },
    },
    required: ["libraryName", "fileId"],
  },
};

export async function getFileVersions(libraryName: string, fileId: string) {
  const client = await getGraphClient();
  const drive = await resolveDrive(client, libraryName);

  const res = await client.get(`/drives/${drive.id}/items/${fileId}/versions`);
  const versions = (res.data.value || []).map((v: any) => ({
    versionId: v.id,
    lastModifiedBy: v.lastModifiedBy?.user?.displayName || "Unknown",
    lastModifiedAt: v.lastModifiedDateTime,
    size: v.size ? `${(v.size / 1024).toFixed(1)} KB` : null,
  }));

  return { libraryName, fileId, totalVersions: versions.length, versions };
}