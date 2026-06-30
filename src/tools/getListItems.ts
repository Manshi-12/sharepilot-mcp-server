import { getGraphClient } from "../auth/graphClient.js";
import { resolveList, getListColumns, parseImageFieldValue, getUserMap } from "../utils/resolve.js";

const SITE_ID = process.env.SITE_ID || "";
const SITE_URL = process.env.SITE_URL || "";

export const getListItemsToolSchema = {
  name: "get_list_items",
  description:
    "Fetches items (rows) from any SharePoint List on the site by its display name — " +
    "works for any list, not just a specific one. Returns each row's fields using their " +
    "friendly display names (Title, Status, Priority, Due Date, etc.), with choice, number, " +
    "date, yes/no, image/hyperlink, lookup and person fields all included. " +
    "Use this whenever the user asks to see, fetch, list, or summarize data from a SharePoint list.",
  inputSchema: {
    type: "object",
    properties: {
      listName: {
        type: "string",
        description: "Display name of the SharePoint list, e.g. 'Project Tasks'.",
      },
      search: {
        type: "string",
        description:
          "Optional. If provided, only rows where at least one field contains this text " +
          "(case-insensitive) are returned.",
      },
      top: {
        type: "number",
        description: "Optional. Maximum number of rows to return. Defaults to 10. Do not request more than 10 initially unless the user explicitly asks.",
      },
    },
    required: ["listName"],
  },
};

function buildItemDisplayUrl(listInternalName: string, itemId: string): string {
  if (!SITE_URL) return "";
  return `${SITE_URL}/Lists/${listInternalName}/DispForm.aspx?ID=${itemId}`;
}

export async function getListItems(listName: string, search?: string, top: number = 10) {
  const client = await getGraphClient();

  const list = await resolveList(client, listName);
  const columns = await getListColumns(client, list.id);

  const colMap = new Map<string, typeof columns[0]>();
  for (const col of columns) colMap.set(col.internalName, col);

  const nameMap = new Map<string, string>();
  for (const col of columns) nameMap.set(col.internalName, col.displayName);

  const res = await client.get(
    `/sites/${SITE_ID}/lists/${list.id}/items`,
    { params: { $expand: "fields", $top: 200 } }
  );

  const rawItems = res.data.value || [];

  // Fix #9 — use shared cached getUserMap instead of fetching users inline every time
  let userMap = new Map<number, string>();
  try {
    userMap = await getUserMap(client);
  } catch {
    // If user map fails, person fields will show IDs — not fatal
  }

  const searchLower = search?.toLowerCase();

  const items = rawItems
    .map((item: any) => {
      const fields = item.fields || {};
      const cleaned: Record<string, any> = {};

      for (const [internalName, value] of Object.entries(fields)) {
        if (
          internalName.startsWith("@") ||
          [
            "ContentType", "Attachments", "Edit", "LinkTitle", "LinkTitleNoMenu",
            "ItemChildCount", "FolderChildCount", "_ComplianceFlags", "_ComplianceTag",
            "_ComplianceTagWrittenTime", "_ComplianceTagUserId", "AppAuthor", "AppEditor",
          ].includes(internalName)
        ) {
          continue;
        }

        const displayName = nameMap.get(internalName) || internalName;
        const col = colMap.get(internalName);

        if (col?.type === "personOrGroup") {
          if (typeof value === "number") {
            cleaned[displayName] = userMap.get(value) ?? `User ${value}`;
          } else if (Array.isArray(value)) {
            cleaned[displayName] = value.map((v: any) =>
              typeof v === "number"
                ? (userMap.get(v) ?? `User ${v}`)
                : (v?.LookupValue || v?.Title || String(v))
            );
          } else if (value && typeof value === "object") {
            cleaned[displayName] = (value as any).LookupValue || (value as any).Title || String(value);
          } else {
            cleaned[displayName] = value;
          }
          continue;
        }

        if (internalName.endsWith("LookupId") || internalName.endsWith("LookupIds")) {
          continue;
        }

        if (typeof value === "string" && value.includes("serverRelativeUrl")) {
          const parsedImage = parseImageFieldValue(value);
          if (parsedImage) {
            cleaned[displayName] = { imageUrl: parsedImage.url, fileName: parsedImage.fileName };
            continue;
          }
        }

        if (col?.type === "hyperlinkOrPicture") {
          if (value && typeof value === "object" && (value as any).Url) {
            cleaned[displayName] = { url: (value as any).Url, label: (value as any).Description };
            continue;
          }
        }

        if (value && typeof value === "object" && ("Url" in value || "Description" in value)) {
          cleaned[displayName] = {
            url: (value as any).Url,
            label: (value as any).Description,
          };
          continue;
        }

        cleaned[displayName] = value;
      }

      const displayUrl = buildItemDisplayUrl(list.name, item.id);

      return {
        id: item.id,
        viewUrl: displayUrl || item.webUrl,
        fields: cleaned,
      };
    });

  const filteredItems = items.filter((item: any) => {
    if (!searchLower) return true;
    return Object.values(item.fields).some((v) =>
      String(v ?? "").toLowerCase().includes(searchLower)
    );
  });

  const totalMatches = filteredItems.length;
  const slicedItems = filteredItems.slice(0, top);

  return {
    listName: list.displayName,
    totalMatchCount: totalMatches,
    returnedCount: slicedItems.length,
    items: slicedItems,
  };
}
