import { getGraphClient } from "../auth/graphClient.js";
import { resolveList, getListColumns as fetchColumns } from "../utils/resolve.js";

export const getListColumnsToolSchema = {
  name: "get_list_columns",
  description:
    "Returns all columns (fields/schema) of a SharePoint list or document library. " +
    "Shows each column's display name, internal name, type, whether it's required, " +
    "and available choices for choice columns. Use when the user asks what columns or " +
    "fields a list has, or wants to understand the structure of a list.",
  inputSchema: {
    type: "object",
    properties: {
      listName: { type: "string", description: "Display name of the SharePoint list or library." },
    },
    required: ["listName"],
  },
};

export async function getListColumns(listName: string) {
  const client = await getGraphClient();
  const list = await resolveList(client, listName);
  const columns = await fetchColumns(client, list.id);

  return {
    listName: list.displayName,
    totalColumns: columns.length,
    columns: columns.map((c) => ({
      displayName: c.displayName,
      internalName: c.internalName,
      type: c.type,
      required: c.required,
      ...(c.choices ? { choices: c.choices } : {}),
    })),
  };
}