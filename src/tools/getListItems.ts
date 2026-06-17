import { getGraphClient } from "../auth/graphClient.js";
import { resolveList, getListColumns, parseImageFieldValue } from "../utils/resolve.js";

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
        description: "Optional. Maximum number of rows to return. Defaults to 100.",
      },
    },
    required: ["listName"],
  },
};

/**
 * Builds the human-readable SharePoint list item display URL.
 * Graph's webUrl returns a .000 file download — this is the correct form URL.
 */
function buildItemDisplayUrl(listInternalName: string, itemId: string): string {
  if (!SITE_URL) return "";
  // listInternalName from Graph is URL-encoded sometimes, keep as-is
  return `${SITE_URL}/Lists/${listInternalName}/DispForm.aspx?ID=${itemId}`;
}

export async function getListItems(listName: string, search?: string, top: number = 100) {
  const client = await getGraphClient();

  const list = await resolveList(client, listName);
  const columns = await getListColumns(client, list.id);

  // Map internalName → column metadata (for type checking and display names)
  const colMap = new Map<string, typeof columns[0]>();
  for (const col of columns) colMap.set(col.internalName, col);

  // Also build internalName → displayName for fast lookup
  const nameMap = new Map<string, string>();
  for (const col of columns) nameMap.set(col.internalName, col.displayName);

  // Fetch all items — $expand: fields gives us all field values
  const res = await client.get(
    `/sites/${SITE_ID}/lists/${list.id}/items`,
    { params: { $expand: "fields", $top: Math.min(top, 200) } }
  );

  const rawItems = res.data.value || [];

  // Pre-fetch User Information List once to resolve person IDs to names
  // We do this by fetching items from the User Information List
  let userMap = new Map<number, string>(); // numeric ID → display name
  try {
    const allLists = await client.get(`/sites/${SITE_ID}/lists?$select=id,displayName,list`);
    const userInfoList = (allLists.data.value || []).find(
      (l: any) => (l.displayName || "").toLowerCase() === "user information list"
    );
    if (userInfoList) {
      const usersRes = await client.get(
        `/sites/${SITE_ID}/lists/${userInfoList.id}/items`,
        { params: { "$expand": "fields($select=Title,EMail)", "$top": 2000 } }
      );
      for (const u of usersRes.data.value || []) {
        const numId = Number(u.id);
        const name = u.fields?.Title || u.fields?.EMail || `User ${numId}`;
        userMap.set(numId, name);
      }
    }
  } catch {
    // If user info list fetch fails, we'll just show IDs — not a fatal error
  }

  const searchLower = search?.toLowerCase();

  const items = rawItems
    .map((item: any) => {
      const fields = item.fields || {};
      const cleaned: Record<string, any> = {};

      for (const [internalName, value] of Object.entries(fields)) {
        // Skip Graph/SharePoint internal bookkeeping fields
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

        // --- Person/Group fields ---
        // Graph returns these as numeric IDs in fields (e.g. AssignedToLookupId: 11)
        // The field itself comes as the raw value; the LookupId variant is separate.
        // We detect them by column type and resolve to display names.
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

        // Skip raw LookupId fields — we already handle person fields above by type
        if (internalName.endsWith("LookupId") || internalName.endsWith("LookupIds")) {
          continue;
        }

        // --- Image fields ---
        // SharePoint image columns store a JSON string with serverUrl + serverRelativeUrl.
        // Uses the SAME parser as upload_list_item_image so both tools always agree.
        if (typeof value === "string" && value.includes("serverRelativeUrl")) {
          const parsedImage = parseImageFieldValue(value);
          if (parsedImage) {
            cleaned[displayName] = { imageUrl: parsedImage.url, fileName: parsedImage.fileName };
            continue;
          }
        }

        // hyperlinkOrPicture columns (non-image, plain URL type)
        if (col?.type === "hyperlinkOrPicture") {
          if (value && typeof value === "object" && (value as any).Url) {
            cleaned[displayName] = { url: (value as any).Url, label: (value as any).Description };
            continue;
          }
        }

        // --- Hyperlink fields (Url/Description objects) ---
        if (value && typeof value === "object" && ("Url" in value || "Description" in value)) {
          cleaned[displayName] = {
            url: (value as any).Url,
            label: (value as any).Description,
          };
          continue;
        }

        // Everything else: pass through as-is
        cleaned[displayName] = value;
      }

      // Build proper display URL (not the .000 Graph webUrl)
      const displayUrl = buildItemDisplayUrl(list.name, item.id);

      return {
        id: item.id,
        viewUrl: displayUrl || item.webUrl,
        fields: cleaned,
      };
    })
    .filter((item: any) => {
      if (!searchLower) return true;
      return Object.values(item.fields).some((v) =>
        String(v ?? "").toLowerCase().includes(searchLower)
      );
    });

  return {
    listName: list.displayName,
    matchCount: items.length,
    items,
  };
}