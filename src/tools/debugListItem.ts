import { getGraphClient } from "../auth/graphClient.js";

const SITE_ID = process.env.SITE_ID || "";

export const debugListItemToolSchema = {
  name: "debug_list_item",
  description:
    "Fetches all items OR a specific item from ANY SharePoint list. " +
    "Use listName to specify which list. Optionally provide itemId to fetch one item. " +
    "If no itemId given, returns all items in the list.",
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

  // Resolve list by display name
  const listsRes = await client.get(
    `/sites/${SITE_ID}/lists?$select=id,name,displayName`
  );
  const lists = listsRes.data.value || [];
  const found = lists.find(
    (l: any) =>
      l.displayName?.toLowerCase() === listName.toLowerCase() ||
      l.name?.toLowerCase() === listName.toLowerCase()
  );

  if (!found) {
    throw new Error(
      `List "${listName}" not found. Available lists: ${lists.map((l: any) => l.displayName).join(", ")}`
    );
  }

  const listId = found.id;

  // Get column definitions
  const colRes = await client.get(
    `/sites/${SITE_ID}/lists/${listId}/columns?$select=name,displayName,type,choice`
  );
  const columnDefinitions = (colRes.data.value || []).map((c: any) => ({
    internalName: c.name,
    displayName: c.displayName,
    type: c.type,
    displayAs: c.choice?.displayAs || undefined,
    choices: c.choice?.choices || undefined,
  }));

  if (itemId) {
    // Fetch specific item
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
    // Fetch ALL items
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