import { getGraphClient } from "../auth/graphClient.js";

const SITE_ID = process.env.SITE_ID || "";

export const debugListItemToolSchema = {
  name: "debug_list_item",
  description:
    "Fetches all items OR a specific item from ANY SharePoint list. " +
    "Use listName to specify which list. Optionally provide itemId to fetch one item. " +
    "If no itemId given, returns all items in the list. " +
    "Also returns raw column internal names so we can debug field mapping.",
  inputSchema: {
    type: "object",
    properties: {
      listName: {
        type: "string",
        description: "Display name of the SharePoint list, e.g. 'Project Tasks' or 'testList'.",
      },
      itemId: {
        type: "string",
        description: "Optional. The numeric ID of a specific list item to read.",
      },
    },
    required: ["listName"],
  },
};

export async function debugListItem(listName: string, itemId?: string) {
  const client = await getGraphClient();

  // Step 1: Find the list by display name
  const listsRes = await client.get(`/sites/${SITE_ID}/lists?$select=id,name,displayName`);
  const lists = listsRes.data.value || [];
  const found = lists.find(
    (l: any) =>
      l.displayName?.toLowerCase() === listName.toLowerCase() ||
      l.name?.toLowerCase() === listName.toLowerCase()
  );

  if (!found) {
    throw new Error(
      `List "${listName}" not found. Available: ${lists.map((l: any) => l.displayName).join(", ")}`
    );
  }

  const listId = found.id;

  // Step 2: Get FULL column definitions including readOnly, hidden, required
  const colRes = await client.get(
    `/sites/${SITE_ID}/lists/${listId}/columns`
  );
  const columnDefinitions = (colRes.data.value || []).map((c: any) => ({
    internalName: c.name,
    displayName: c.displayName,
    type: c.type,
    readOnly: c.readOnly || false,
    hidden: c.hidden || false,
    required: c.required || false,
    displayAs: c.choice?.displayAs || undefined,
    choices: c.choice?.choices || undefined,
  }));

  if (itemId) {
    const itemRes = await client.get(
      `/sites/${SITE_ID}/lists/${listId}/items/${itemId}/fields`
    );
    return {
      listName: found.displayName,
      listId,
      item: itemRes.data,
      columnDefinitions,
    };
  } else {
    const itemsRes = await client.get(
      `/sites/${SITE_ID}/lists/${listId}/items?expand=fields&$top=50`
    );
    const items = (itemsRes.data.value || []).map((i: any) => i.fields);
    return {
      listName: found.displayName,
      listId,
      totalItems: items.length,
      items,
      columnDefinitions,
    };
  }
}