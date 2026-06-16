import { getGraphClient } from "../auth/graphClient.js";
import { resolveList } from "../utils/resolve.js";

const SITE_ID = process.env.SITE_ID || "";

export const getListItemsToolSchema = {
  name: "get_list_items",
  description:
    "Fetches items (rows) from any SharePoint list on the site. " +
    "Returns human-readable field values — person fields show display names, " +
    "image fields show direct URLs, lookup fields show their text values.",
  inputSchema: {
    type: "object",
    properties: {
      listName: {
        type: "string",
        description: "Display name of the SharePoint list, e.g. 'Project Tasks'.",
      },
      search: {
        type: "string",
        description: "Optional: filter items by a search term (matched against Title).",
      },
      top: {
        type: "number",
        description: "Optional: max number of items to return (default 20).",
      },
    },
    required: ["listName"],
  },
};

/** Parses a SharePoint image field JSON string into a plain URL string. */
function parseImageField(raw: any): string | null {
  if (!raw) return null;
  // If it's already a plain URL string
  if (typeof raw === "string") {
    try {
      const parsed = JSON.parse(raw);
      // SharePoint image JSON: { serverUrl, serverRelativeUrl, ... }
      if (parsed.serverUrl && parsed.serverRelativeUrl) {
        return parsed.serverUrl + parsed.serverRelativeUrl;
      }
      if (parsed.url) return parsed.url;
    } catch {
      // Not JSON — might be a raw URL already
      if (raw.startsWith("http")) return raw;
    }
  }
  if (typeof raw === "object") {
    if (raw.serverUrl && raw.serverRelativeUrl) return raw.serverUrl + raw.serverRelativeUrl;
    if (raw.url) return raw.url;
  }
  return null;
}

/** Cleans up a single item's fields into human-readable form. */
function cleanFields(fields: Record<string, any>): Record<string, any> {
  const cleaned: Record<string, any> = {};

  for (const [key, value] of Object.entries(fields)) {
    // Skip Graph internal bookkeeping fields
    if (key.startsWith("@") || key === "Edit" || key === "LinkTitleNoMenu" ||
        key === "LinkTitle" || key === "_UIVersionString" || key === "ItemChildCount" ||
        key === "FolderChildCount" || key === "appAuthorId" || key === "appEditorId") {
      continue;
    }

    // Person/lookup fields: Graph returns both "AssignedTo" (display name expand)
    // and "AssignedToLookupId" (numeric ID). Keep only the display name version.
    if (key.endsWith("LookupId") || key.endsWith("LookupIds")) {
      continue; // skip raw ID fields — we show the expanded name instead
    }

    // Image fields: Graph returns a JSON string — parse it to a clean URL
    if (typeof value === "string" && value.includes("serverRelativeUrl")) {
      const url = parseImageField(value);
      cleaned[key] = url ?? value;
      continue;
    }

    // Person fields expanded by Graph come as objects with { LookupValue, Email, ... }
    if (typeof value === "object" && value !== null && !Array.isArray(value)) {
      if (value.LookupValue) {
        cleaned[key] = value.LookupValue;
        continue;
      }
      if (value.Title) {
        cleaned[key] = value.Title;
        continue;
      }
    }

    // Array of person/lookup objects
    if (Array.isArray(value)) {
      cleaned[key] = value.map((v: any) =>
        typeof v === "object" ? (v.LookupValue || v.Title || v.Email || JSON.stringify(v)) : v
      );
      continue;
    }

    cleaned[key] = value;
  }

  return cleaned;
}

export async function getListItems(
  listName: string,
  search?: string,
  top: number = 20
) {
  const client = await getGraphClient();
  const list = await resolveList(client, listName);

  // Get columns to identify person and image fields for proper $expand
  const colsRes = await client.get(`/sites/${SITE_ID}/lists/${list.id}/columns`);
  const columns = colsRes.data.value || [];

  // Build $expand for person/lookup fields so Graph returns display names
  const personCols = columns
    .filter((c: any) => c.personOrGroup || c.lookup)
    .map((c: any) => c.name);

  let expandParam = "fields";
  if (personCols.length > 0) {
    // Expand person fields so we get LookupValue (display name) not just numeric IDs
    expandParam = `fields($expand=${personCols.join(",")})`;
  }

  const params: Record<string, any> = {
    $expand: expandParam,
    $top: top,
  };

  if (search) {
    params.$filter = `startswith(fields/Title,'${search}')`;
  }

  const res = await client.get(`/sites/${SITE_ID}/lists/${list.id}/items`, { params });
  const items = res.data.value || [];

  const cleaned = items.map((item: any) => ({
    id: item.id,
    webUrl: item.webUrl,
    fields: cleanFields(item.fields || {}),
  }));

  return {
    listName: list.displayName,
    totalReturned: cleaned.length,
    items: cleaned,
  };
}