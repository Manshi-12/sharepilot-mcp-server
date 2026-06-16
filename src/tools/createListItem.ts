import { getGraphClient } from "../auth/graphClient.js";
import { resolveList, getListColumns, findColumn, ColumnInfo } from "../utils/resolve.js";

const SITE_ID = process.env.SITE_ID || "";

export const createListItemToolSchema = {
  name: "create_list_item",
  description:
    "Creates a new item/row in any SharePoint List on the site by display name. " +
    "Field names should match the list's actual column display names (e.g. 'Title', " +
    "'Priority', 'Due Date'). The server automatically discovers the list's real schema " +
    "(choice, multi-choice, date, number, yes/no, lookup, person fields) — works for any " +
    "list, not just one specific one.",
  inputSchema: {
    type: "object",
    properties: {
      listName: {
        type: "string",
        description: "Display name of the SharePoint list, e.g. 'Project Tasks'.",
      },
      fields: {
        type: "object",
        description: "Field display-name / value pairs to set on the new item.",
        additionalProperties: true,
      },
    },
    required: ["listName", "fields"],
  },
};

/** Coerces a raw value into the shape SharePoint expects for a given column type. */
function coerceValue(col: ColumnInfo, value: any): any {
  switch (col.type) {
    case "boolean":
      if (typeof value === "boolean") return value;
      return String(value).trim().toLowerCase() === "true";
    case "number":
    case "currency":
      return typeof value === "number" ? value : Number(value);
    case "dateTime": {
      // Accept "YYYY-MM-DD" or a full ISO string; normalize to ISO with time.
      const d = new Date(value);
      return isNaN(d.getTime()) ? value : d.toISOString();
    }
    case "multiChoice":
      return Array.isArray(value) ? value : [value];
    default:
      return value;
  }
}

const CHOICE_TYPES = new Set(["choice", "multiChoice"]);

export async function createListItem(listName: string, fields: Record<string, any>) {
  const client = await getGraphClient();

  const list = await resolveList(client, listName);
  const columns = await getListColumns(client, list.id);

  const safeFields: Record<string, any> = {};      // first-pass POST (non-choice columns)
  const choiceFields: Record<string, { internalName: string; value: any; type: ColumnInfo["type"] }> = {};
  const unmatchedFields: Record<string, any> = {};  // keys we couldn't map to any real column

  for (const [key, rawValue] of Object.entries(fields)) {
    if (key.toLowerCase() === "title") {
      safeFields["Title"] = rawValue;
      continue;
    }

    const col = findColumn(columns, key);
    if (!col) {
      unmatchedFields[key] = rawValue;
      continue;
    }

    const value = coerceValue(col, rawValue);

    if (CHOICE_TYPES.has(col.type)) {
      // Choice-type columns frequently 500 when sent in the initial POST (this is the
      // checkbox-display-type quirk you ran into) — set them in a second PATCH instead.
      choiceFields[key] = { internalName: col.internalName, value, type: col.type };
    } else {
      safeFields[col.internalName] = value;
    }
  }

  const res = await client.post(
    `/sites/${SITE_ID}/lists/${list.id}/items`,
    { fields: safeFields }
  );

  const itemId = res.data.id;

  const verifiedFields: Record<string, any> = { ...safeFields };
  const fieldErrors: Record<string, string> = {};

  for (const [displayKey, { internalName, value, type }] of Object.entries(choiceFields)) {
    try {
      await client.patch(
        `/sites/${SITE_ID}/lists/${list.id}/items/${itemId}/fields`,
        { [internalName]: value }
      );
      verifiedFields[displayKey] = value;
    } catch {
      // Fallback: some choice columns insist on array format even when single-value.
      try {
        const arrayValue = Array.isArray(value) ? value : [value];
        await client.patch(
          `/sites/${SITE_ID}/lists/${list.id}/items/${itemId}/fields`,
          { [internalName]: arrayValue }
        );
        verifiedFields[displayKey] = arrayValue;
      } catch (e: any) {
        fieldErrors[displayKey] = e?.response?.data?.error?.message || e.message || "Failed to set this field.";
      }
    }
  }

  for (const key of Object.keys(unmatchedFields)) {
    fieldErrors[key] = `No column named "${key}" exists on the "${list.displayName}" list.`;
  }

  const missingFields = Object.keys(fieldErrors);
  const status = missingFields.length === 0 ? "fully_created" : "partially_created";

  return {
    id: itemId,
    webUrl: res.data.webUrl,
    listName: list.displayName,
    status,
    verifiedFields,
    missingFields,
    fieldErrors,
  };
}
