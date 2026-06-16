import { getGraphClient } from "../auth/graphClient.js";
import { resolveList, resolveDrive } from "../utils/resolve.js";

const SITE_ID = process.env.SITE_ID || "";

export const uploadListItemImageToolSchema = {
  name: "upload_list_item_image",
  description:
    "Uploads an image (as base64) and attaches it to an Image column on a SharePoint list item. " +
    "First uploads the image file to the Site Assets library, then sets the Image field on the specified item. " +
    "The fileName must be a simple name with no spaces, e.g. 'photo.jpg' or 'task_image.png'.",
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
        description: "Display name of the Image column, e.g. 'TaskImage'.",
      },
      fileName: {
        type: "string",
        description: "Filename with extension and NO spaces, e.g. 'photo.jpg' or 'ayush_shah.png'.",
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

/** Sanitizes a filename — replaces spaces/special chars with underscores, keeps extension. */
function sanitizeFileName(name: string): string {
  return name.trim().replace(/\s+/g, "_").replace(/[^a-zA-Z0-9._-]/g, "_");
}

export async function uploadListItemImage(
  listName: string,
  itemId: string,
  imageFieldName: string,
  fileName: string,
  base64Content: string,
  mimeType: string
) {
  const client = await getGraphClient();

  // Always sanitize filename — no spaces allowed in SharePoint image field paths
  const safeFileName = sanitizeFileName(fileName);

  // Step 1: Resolve the list
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

  // Step 3: Upload image to Site Assets library
  let drive;
  try {
    drive = await resolveDrive(client, "Site Assets");
  } catch {
    drive = await resolveDrive(client, "Documents");
  }

  const imageBuffer = Buffer.from(base64Content, "base64");

  const uploadRes = await client.put(
    `/drives/${drive.id}/root:/${safeFileName}:/content`,
    imageBuffer,
    {
      headers: {
        "Content-Type": mimeType,
        "Content-Length": imageBuffer.length.toString(),
      },
    }
  );

  const uploadedItem = uploadRes.data;

  // Step 4: Build serverUrl and serverRelativeUrl from Graph's response
  // Graph returns webUrl like: https://tenant.sharepoint.com/sites/Site/Site Assets/file.png
  // We need: serverUrl = https://tenant.sharepoint.com
  //          serverRelativeUrl = /sites/Site/Site Assets/file.png  (URL-decoded)
  const uploadedWebUrl: string = uploadedItem.webUrl || "";
  if (!uploadedWebUrl) {
    throw new Error(
      "Upload succeeded but Graph did not return a webUrl. Cannot set image field."
    );
  }

  const parsedUrl = new URL(uploadedWebUrl);
  const serverUrl = parsedUrl.origin; // https://dwivedimanshi12outlook.sharepoint.com
  // pathname is already URL-decoded by the URL constructor — this is correct for SharePoint
  const serverRelativeUrl = decodeURIComponent(parsedUrl.pathname);

  // Step 5: Build the thumbnail JSON SharePoint expects for image columns
  const imageFieldValue = {
    type: "thumbnail",
    fileName: safeFileName,
    serverUrl: serverUrl,
    serverRelativeUrl: serverRelativeUrl,
  };

  // Step 6: Patch the list item's image field — only once
  await client.patch(
    `/sites/${SITE_ID}/lists/${list.id}/items/${itemId}/fields`,
    { [internalName]: JSON.stringify(imageFieldValue) }
  );

  return {
    success: true,
    itemId,
    listName: list.displayName,
    imageField: imageFieldName,
    uploadedTo: uploadedWebUrl,
    imageUrl: serverUrl + serverRelativeUrl,
    imageFieldValue,
  };
}