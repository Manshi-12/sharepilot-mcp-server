import { getGraphClient } from "../auth/graphClient.js";
import { resolveList, getListColumns } from "../utils/resolve.js";

const SITE_ID = process.env.SITE_ID || "";

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

function isImageOrHyperlinkValue(v: any): boolean {
  return v && typeof v === "object" && ("Url" in v || "Description" in v);
}

export async function getListItems(listName: string, search?: string, top: number = 100) {
  const client = await getGraphClient();

  const list = await resolveList(client, listName);
  const columns = await getListColumns(client, list.id);

  // Build internalName -> displayName so the response speaks in friendly field names,
  // not SharePoint's internal field_2 / field_10 style names.
  const nameMap = new Map<string, string>();
  for (const col of columns) nameMap.set(col.internalName, col.displayName);

  const res = await client.get(
    `/sites/${SITE_ID}/lists/${list.id}/items`,
    { params: { $expand: "fields", $top: Math.min(top, 200) } }
  );

  const rawItems = res.data.value || [];
  const searchLower = search?.toLowerCase();

  const items = rawItems
    .map((item: any) => {
      const fields = item.fields || {};
      const cleaned: Record<string, any> = {};

      for (const [internalName, value] of Object.entries(fields)) {
        // Skip Graph/SharePoint bookkeeping fields nobody asked for.
        if (
          internalName.startsWith("@") ||
          ["ContentType", "Attachments", "Edit", "LinkTitle", "LinkTitleNoMenu", "ItemChildCount", "FolderChildCount", "_ComplianceFlags", "_ComplianceTag", "_ComplianceTagWrittenTime", "_ComplianceTagUserId", "AppAuthor", "AppEditor"].includes(internalName)
        ) {
          continue;
        }

        const displayName = nameMap.get(internalName) || internalName;

        // Image / hyperlink columns come back as { Description, Url } objects — keep them
        // structured rather than dropping or mangling them (fixes requirement: image fields
        // should be retrievable too).
        if (isImageOrHyperlinkValue(value)) {
          cleaned[displayName] = { url: (value as any).Url, label: (value as any).Description };
        } else {
          cleaned[displayName] = value;
        }
      }

      return {
        id: item.id,
        webUrl: item.webUrl,
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
