import { getGraphClient } from "../auth/graphClient.js";
import { resolveList, getListColumns, parseImageFieldValue, getUserMap } from "../utils/resolve.js";

const SITE_ID = process.env.SITE_ID || "";
const SITE_URL = process.env.SITE_URL || "";

export const getListItemByIdToolSchema = {
  name: "get_list_item_by_id",
  description:
    "Fetches a single item from a SharePoint list by its numeric item ID. " +
    "Returns all fields with friendly display names. Use this when the user asks to " +
    "see, view, or get details of a specific item by ID.",
  inputSchema: {
    type: "object",
    properties: {
      listName: { type: "string", description: "Display name of the SharePoint list." },
      itemId: { type: "number", description: "Numeric ID of the list item." },
    },
    required: ["listName", "itemId"],
  },
};

export async function getListItemById(listName: string, itemId: number) {
  const client = await getGraphClient();
  const list = await resolveList(client, listName);
  const columns = await getListColumns(client, list.id);

  const nameMap = new Map<string, string>();
  const colMap = new Map<string, typeof columns[0]>();
  for (const col of columns) {
    nameMap.set(col.internalName, col.displayName);
    colMap.set(col.internalName, col);
  }

  const res = await client.get(
    `/sites/${SITE_ID}/lists/${list.id}/items/${itemId}`,
    { params: { $expand: "fields" } }
  );

  const item = res.data;
  const fields = item.fields || {};
  const cleaned: Record<string, any> = {};

  let userMap = new Map<number, { name: string; email: string }>();
  try { userMap = await getUserMap(client); } catch { }

  for (const [internalName, value] of Object.entries(fields)) {
    if (
      internalName.startsWith("@") ||
      ["ContentType","Attachments","Edit","LinkTitle","LinkTitleNoMenu",
       "ItemChildCount","FolderChildCount","_ComplianceFlags","_ComplianceTag",
       "_ComplianceTagWrittenTime","_ComplianceTagUserId","AppAuthor","AppEditor"].includes(internalName)
    ) continue;

    const displayName = nameMap.get(internalName) || internalName;
    const col = colMap.get(internalName);

    if (col?.type === "personOrGroup") {
      if (typeof value === "number") {
        cleaned[displayName] = userMap.get(value)?.name ?? `User ${value}`;
      } else if (Array.isArray(value)) {
        cleaned[displayName] = value.map((v: any) =>
          typeof v === "number" ? (userMap.get(v)?.name ?? `User ${v}`) : (v?.LookupValue || String(v))
        );
      } else {
        cleaned[displayName] = (value as any)?.LookupValue || String(value);
      }
      continue;
    }

    if (internalName.endsWith("LookupId") || internalName.endsWith("LookupIds")) continue;

    if (typeof value === "string" && value.includes("serverRelativeUrl")) {
      const parsedImage = parseImageFieldValue(value);
      if (parsedImage) { cleaned[displayName] = { imageUrl: parsedImage.url, fileName: parsedImage.fileName }; continue; }
    }

    if (col?.type === "hyperlinkOrPicture" && value && typeof value === "object" && (value as any).Url) {
      cleaned[displayName] = { url: (value as any).Url, label: (value as any).Description };
      continue;
    }

    cleaned[displayName] = value;
  }

  const viewUrl = SITE_URL ? `${SITE_URL}/Lists/${list.name}/DispForm.aspx?ID=${itemId}` : item.webUrl;

  return { listName: list.displayName, itemId, fields: cleaned, viewUrl };
}