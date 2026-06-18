import { getGraphClient } from "../auth/graphClient.js";
import { resolveList, clearResolverCache } from "../utils/resolve.js";

const SITE_ID = process.env.SITE_ID || "";

export const deleteListItemToolSchema = {
  name: "delete_list_item",
  description:
    "Deletes a specific item/row from a SharePoint List, OR deletes an entire List or Document Library. " +
    "To delete a single item: provide listName + itemId. " +
    "To delete the entire list/library: provide listName only (leave itemId empty). " +
    "WARNING: Deleting a list/library is permanent and removes ALL its items/files.",
  inputSchema: {
    type: "object",
    properties: {
      listName: {
        type: "string",
        description: "Display name of the SharePoint list or library to target.",
      },
      itemId: {
        type: "string",
        description:
          "The numeric ID of the specific item to delete. " +
          "Leave this out to delete the ENTIRE list/library instead.",
      },
    },
    required: ["listName"],
  },
};

export async function deleteListItem(listName: string, itemId?: string) {
  const client = await getGraphClient();
  const list = await resolveList(client, listName);

  if (itemId) {
    // Delete a single item
    await client.delete(
      `/sites/${SITE_ID}/lists/${list.id}/items/${itemId}`
    );
    return {
      success: true,
      deleted: "item",
      itemId,
      listName: list.displayName,
      message: `Item #${itemId} was permanently deleted from "${list.displayName}".`,
    };
  } else {
    // Delete the entire list/library
    await client.delete(`/sites/${SITE_ID}/lists/${list.id}`);
    // Clear cache so the deleted list doesn't appear in future lookups
    clearResolverCache();
    return {
      success: true,
      deleted: "list",
      listName: list.displayName,
      message: `The list/library "${list.displayName}" and all its contents were permanently deleted.`,
    };
  }
}