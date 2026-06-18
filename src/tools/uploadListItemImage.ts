import { getGraphClient, getSharePointAccessToken } from "../auth/graphClient.js";
import { resolveList, resolveDrive, parseImageFieldValue } from "../utils/resolve.js";

const SITE_ID = process.env.SITE_ID || "";
const SITE_URL = process.env.SITE_URL || "";

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

function stripDataUriPrefix(content: string): string {
  const commaIndex = content.indexOf(",");
  if (content.startsWith("data:") && commaIndex !== -1) {
    return content.slice(commaIndex + 1);
  }
  return content;
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
      `Image column "${imageFieldName}" not found on list "${listName}".`
    );
  }
  const internalName = imageCol.name;

  // Step 3: Upload image to Site Assets library
  let drive: any;
  try {
    drive = await resolveDrive(client, "Site Assets");
  } catch {
    drive = await resolveDrive(client, "Documents");
  }

  const imageBuffer = Buffer.from(stripDataUriPrefix(base64Content), "base64");

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
  const uploadedWebUrl: string = uploadedItem.webUrl || "";
  if (!uploadedWebUrl) {
    throw new Error(
      "Upload succeeded but Graph did not return a webUrl. Cannot set image field."
    );
  }

  const parsedUrl = new URL(uploadedWebUrl);
  const serverUrl = parsedUrl.origin; // https://dwivedimanshi12outlook.sharepoint.com
  const serverRelativeUrl = decodeURIComponent(parsedUrl.pathname);

  // Step 5: Build the thumbnail JSON SharePoint expects for image columns
  const imageFieldValue = {
    type: "thumbnail",
    fileName: safeFileName,
    serverUrl: serverUrl,
    serverRelativeUrl: serverRelativeUrl,
  };

  // Step 6: Use SharePoint REST API validateUpdateListItem to set the image field.
  // WHY NOT Graph PATCH: Graph PATCH accepts the JSON string without error but
  // SharePoint internally cannot resolve the image from it — the column stays
  // broken/empty in the UI. validateUpdateListItem is how SharePoint's own UI
  // sets thumbnail columns and is the only reliable way to do it via API.
  // NOTE: this REST call requires a certificate-based app-only token, not a
  // client-secret-based one — see getSharePointAccessToken() in graphClient.ts.
  const token = await getSharePointAccessToken();

  const restUrl = `${SITE_URL}/_api/lists/getbytitle('${encodeURIComponent(list.displayName)}')/items(${itemId})/validateUpdateListItem`;

  const restResponse = await fetch(restUrl, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json;odata=verbose",
      Accept: "application/json;odata=verbose",
    },
    body: JSON.stringify({
      formValues: [
        {
          FieldName: internalName,
          FieldValue: JSON.stringify(imageFieldValue),
        },
      ],
      bNewDocumentUpdate: false,
      checkInComment: "",
    }),
  });

  if (!restResponse.ok) {
    const errText = await restResponse.text();
    throw new Error(
      `Failed to set image field on item ${itemId}: ${restResponse.status} — ${errText}`
    );
  }

  // Step 7: Re-read the field to get the URL that SharePoint actually stored.
  // SharePoint rewrites the thumbnail value server-side after you set it,
  // so we always re-read rather than trusting the URL we uploaded to.
  const verifyRes = await client.get(
    `/sites/${SITE_ID}/lists/${list.id}/items/${itemId}`,
    { params: { $expand: `fields($select=${internalName})` } }
  );
  const finalRawValue = verifyRes.data.fields?.[internalName];
  const finalImage = parseImageFieldValue(finalRawValue);

  return {
    success: true,
    itemId,
    listName: list.displayName,
    imageField: imageFieldName,
    imageUrl: finalImage?.url || (serverUrl + serverRelativeUrl),
    note:
      "Image is now saved and will render correctly in the SharePoint list. " +
      "You may need to be signed into SharePoint to view the direct link.",
  };
}