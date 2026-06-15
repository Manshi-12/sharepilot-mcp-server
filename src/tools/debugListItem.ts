import { getGraphClient } from "../auth/graphClient.js";

const SITE_ID = process.env.SITE_ID || "";
const LIST_ID = "066a3b58-72a3-4fba-a3fc-3acae90be4bf";

export const debugListItemToolSchema = {
  name: "debug_list_item",
  description: "Reads raw field data from a SharePoint list item for debugging. Returns all internal field names and values.",
  inputSchema: {
    type: "object",
    properties: {
      itemId: {
        type: "string",
        description: "The numeric ID of the list item to read.",
      },
    },
    required: ["itemId"],
  },
};

export async function debugListItem(itemId: string) {
  const client = await getGraphClient();

  // Get raw fields
  const res = await client.get(
    `/sites/${SITE_ID}/lists/${LIST_ID}/items/${itemId}/fields`
  );

  // Get list column definitions
  const colRes = await client.get(
    `/sites/${SITE_ID}/lists/${LIST_ID}/columns?$select=name,displayName,type,choice`
  );

  const columns = (colRes.data.value || []).map((c: any) => ({
    internalName: c.name,
    displayName: c.displayName,
    type: c.type,
    choices: c.choice?.choices || undefined,
  }));

  return {
    rawFields: res.data,
    columnDefinitions: columns,
  };
}