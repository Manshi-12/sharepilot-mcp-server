import { getGraphClient } from "../auth/graphClient.js";
import { resolveList, resolveDrive } from "../utils/resolve.js";

const SITE_ID = process.env.SITE_ID || "";
const SITE_URL = process.env.SITE_URL || ""; // e.g. https://yourorg.sharepoint.com/sites/YourSite

export const uploadListItemImageToolSchema = {
  name: "upload_list_item_image",
  description:
    "Uploads an image (as base64) and attaches it to an Image column on a SharePoint list item. " +
    "First uploads the image file to the Site Assets library, then sets the Image field on the specified item.",
  inputSchema: {
    type: "object",
    properties: {
      listName: {
        type: "string",
        description: "Display name of the SharePoint list, e.g. 'Project Tasks'.",
      },
      itemId: {
        type: "string",
        description: "The numeric ID of the list item to update.",
      },
      imageFieldName: {
        type: "string",
        description: "Display name of the Image column, e.g. 'Thumbnail' or 'ProjectImage'.",
      },
      fileName: {
        type: "string",
        description: "Filename with extension, e.g. 'photo.jpg'.",
      },
      base64Content: {
        type: "string",
        description: "Base64-encoded image content (without the data:image/... prefix).",
      },
      mimeType: {
        type: "string",
        description: "MIME type of the image, e.g. 'image/jpeg', 'image/png'.",
      },
    },
    required: ["listName", "itemId", "imageFieldName", "fileName", "base64Content", "mimeType"],
  },
};

export async function uploadListItemImage(
  listName: string,
  itemId: string,
  imageFieldName: string,
  fileName: string,
  base64Content: string,
  mimeType: string
) {
  const client = await getGraphClient();

  // Step 1: Resolve the list to get its ID
  const list = await resolveList(client, listName);

  // Step 2: Find the image column's internal name
  const colsRes = await client.get(`/sites/${SITE_ID}/lists/${list.id}/columns`);
  const columns = colsRes.data.value || [];
  const imageCol = columns.find(
    (c: any) =>
      (c.displayName || "").toLowerCase() === imageFieldName.toLowerCase() ||
      (c.name || "").toLowerCase() === imageFieldName.toLowerCase()
  );
  if (!imageCol) {
    throw new Error(
      `Image column "${imageFieldName}" not found on list "${listName}". ` +
      `Available columns: ${columns.map((c: any) => c.displayName).join(", ")}`
    );
  }
  const internalName = imageCol.name;

  // Step 3: Upload the image to Site Assets library
  const drive = await resolveDrive(client, "Site Assets");
  const imageBuffer = Buffer.from(base64Content, "base64");

  // Use the Graph upload session for reliability
  const uploadRes = await client.put(
    `/drives/${drive.id}/root:/${fileName}:/content`,
    imageBuffer,
    {
      headers: {
        "Content-Type": mimeType,
        "Content-Length": imageBuffer.length.toString(),
      },
    }
  );

  const uploadedItem = uploadRes.data;
  const uploadedWebUrl: string = uploadedItem.webUrl || "";
  if (!uploadedWebUrl) {
    throw new Error("Upload succeeded but Graph did not return a webUrl. Cannot set image field.");
  }

  const parsedUrl = new URL(uploadedWebUrl);
  const serverUrl = parsedUrl.origin;
  const serverRelativeUrl = parsedUrl.pathname;

  const imageFieldValue = {
    type: "thumbnail",
    fileName: fileName,
    serverUrl: serverUrl,
    serverRelativeUrl: serverRelativeUrl,
  };

  await client.patch(
    `/sites/${SITE_ID}/lists/${list.id}/items/${itemId}/fields`,
    { [internalName]: JSON.stringify(imageFieldValue) }
  );

  await client.patch(
    `/sites/${SITE_ID}/lists/${list.id}/items/${itemId}/fields`,
    { [internalName]: JSON.stringify(imageFieldValue) }
  );

  return {
    success: true,
    itemId,
    listName: list.displayName,
    imageField: imageFieldName,
    uploadedTo: uploadedItem.webUrl,
    imageFieldValue,
  };
}