import { getGraphClient } from "../auth/graphClient.js";
import {
  resolveList,
  getListColumns,
  findColumn,
  resolvePersonId,
  ColumnInfo,
} from "../utils/resolve.js";
import { coerceValue } from "../utils/coerce.js";

const SITE_ID = process.env.SITE_ID || "";

export const updateListItemToolSchema = {
  name: "update_list_item",
  description:
    "Updates fields on an existing item/row in any SharePoint List. " +
    "You must provide the list name and the item's numeric ID (get this from get_list_items). " +
    "Only the fields you pass will be changed — all other fields stay untouched.",
  inputSchema: {
    type: "object",
    properties: {
      listName: {
        type: "string",
        description: "Display name of the SharePoint list, e.g. 'Project Tasks'.",
      },
      itemId: {
        type: "string",
        description: "The numeric ID of the list item to update (from get_list_items).",
      },
      fields: {
        type: "object",
        description:
          "Field display-name / new value pairs to update. Only provided fields will change.",
        additionalProperties: true,
      },
    },
    required: ["listName", "itemId", "fields"],
  },
};

export async function updateListItem(
  listName: string,
  itemId: string,
  fields: Record<string, any>
) {
  const client = await getGraphClient();

  const list = await resolveList(client, listName);
  const columns = await getListColumns(client, list.id);

  const verifiedFields: Record<string, any> = {};
  const fieldErrors: Record<string, string> = {};

  for (const [key, rawValue] of Object.entries(fields)) {
    const col = findColumn(columns, key);
    if (!col) {
      fieldErrors[key] = `No column named "${key}" exists on the "${list.displayName}" list.`;
      continue;
    }

    if (col.type === "personOrGroup") {
      const rawNames = Array.isArray(rawValue) ? rawValue : [rawValue];
      const resolvedIds: number[] = [];
      const unresolvedNames: string[] = [];

      for (const n of rawNames) {
        const id = await resolvePersonId(client, String(n));
        if (id !== null) resolvedIds.push(id);
        else unresolvedNames.push(String(n));
      }

      if (unresolvedNames.length > 0) {
        fieldErrors[key] =
          `Could not find a site user matching: ${unresolvedNames.join(", ")}. ` +
          `Use their exact display name or email.`;
        continue;
      }

      try {
        await client.patch(
          `/sites/${SITE_ID}/lists/${list.id}/items/${itemId}/fields`,
          { [`${col.internalName}LookupId`]: col.multi ? resolvedIds : resolvedIds[0] }
        );
        verifiedFields[key] = rawValue;
      } catch (e: any) {
        fieldErrors[key] =
          e?.response?.data?.error?.message || e.message || "Failed to set this field.";
      }
      continue;
    }

    // Validate and coerce before hitting SharePoint
    let value: any;
    try {
      value = coerceValue(col, rawValue);
    } catch (coerceError: any) {
      fieldErrors[key] = coerceError.message;
      continue;
    }

    try {
      // Graph silently rejects hyperlink/picture field writes unless this
      // header is present — it's a documented Graph quirk, not optional.
      const patchConfig =
        col.type === "hyperlinkOrPicture"
          ? { headers: { Prefer: "apiversion=2.1" } }
          : undefined;

      await client.patch(
        `/sites/${SITE_ID}/lists/${list.id}/items/${itemId}/fields`,
        { [col.internalName]: value },
        patchConfig
      );
      verifiedFields[key] = value;
    } catch (firstError: any) {
      if (col.type === "choice" || col.type === "multiChoice") {
        try {
          const arrayValue = Array.isArray(value) ? value : [value];
          await client.patch(
            `/sites/${SITE_ID}/lists/${list.id}/items/${itemId}/fields`,
            { [col.internalName]: arrayValue }
          );
          verifiedFields[key] = arrayValue;
          continue;
        } catch (secondError: any) {
          fieldErrors[key] =
            secondError?.response?.data?.error?.message ||
            secondError.message ||
            "Failed to set this field.";
          continue;
        }
      }
      fieldErrors[key] =
        firstError?.response?.data?.error?.message ||
        firstError.message ||
        "Failed to set this field.";
    }
  }

  const missingFields = Object.keys(fieldErrors);
  const status = missingFields.length === 0 ? "fully_updated" : "partially_updated";

  return {
    id: itemId,
    listName: list.displayName,
    status,
    verifiedFields,
    missingFields,
    fieldErrors,
  };
}