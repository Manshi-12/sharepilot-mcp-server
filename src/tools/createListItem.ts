import { getGraphClient } from "../auth/graphClient.js";
import { resolveList, getListColumns, findColumn, resolvePersonId, ColumnInfo } from "../utils/resolve.js";

const SITE_ID = process.env.SITE_ID || "";

export const createListItemToolSchema = {
  name: "create_list_item",
  description:
    "Creates a new item/row in any SharePoint List on the site by display name. " +
    "Field names should match the list's actual column display names (e.g. 'Title', " +
    "'Priority', 'Due Date', 'Assigned To'). The server automatically discovers the " +
    "list's real schema — choice, multi-choice, date, number, yes/no, lookup, and " +
    "person fields are all handled correctly. For person fields, pass the person's " +
    "display name or email as it appears on the site.",
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
      const d = new Date(value);
      return isNaN(d.getTime()) ? value : d.toISOString();
    }

    case "choice": {
      // Normalize against known choices (case-insensitive) so "high" matches "High"
      if (col.choices && col.choices.length > 0) {
        const strVal = String(value).trim();
        const match = col.choices.find(
          (c) => c.toLowerCase() === strVal.toLowerCase()
        );
        return match ?? strVal; // use matched casing, or original if not found
      }
      return String(value).trim();
    }

    case "multiChoice": {
      const arr = Array.isArray(value) ? value : [value];
      if (col.choices && col.choices.length > 0) {
        return arr.map((v) => {
          const strVal = String(v).trim();
          const match = col.choices!.find(
            (c) => c.toLowerCase() === strVal.toLowerCase()
          );
          return match ?? strVal;
        });
      }
      return arr;
    }

    default:
      return value;
  }
}

async function patchField(
  client: any,
  listId: string,
  itemId: string,
  body: Record<string, any>
): Promise<void> {
  await client.patch(`/sites/${SITE_ID}/lists/${listId}/items/${itemId}/fields`, body);
}

export async function createListItem(listName: string, fields: Record<string, any>) {
  const client = await getGraphClient();

  const list = await resolveList(client, listName);
  const columns = await getListColumns(client, list.id);

  // Pull Title out separately — it's the one field safe to include in the very first
  // POST that creates the item shell. Every other field is set afterward, one at a
  // time, so a problem with any single field can never block the rest.
  const titleKey = Object.keys(fields).find((k) => k.toLowerCase() === "title");
  const titleValue = titleKey ? fields[titleKey] : "";

  let itemId: string;
  let webUrl: string;
  try {
    const res = await client.post(`/sites/${SITE_ID}/lists/${list.id}/items`, {
      fields: { Title: titleValue },
    });
    itemId = res.data.id;
    webUrl = res.data.webUrl;
  } catch (e: any) {
    throw new Error(
      `Could not create the item at all (failed even with just Title set): ` +
      (e?.response?.data?.error?.message || e.message)
    );
  }

  const verifiedFields: Record<string, any> = { Title: titleValue };
  const fieldErrors: Record<string, string> = {};

  for (const [key, rawValue] of Object.entries(fields)) {
    if (titleKey && key === titleKey) continue;

    const col = findColumn(columns, key);
    if (!col) {
      fieldErrors[key] = `No column named "${key}" exists on the "${list.displayName}" list.`;
      continue;
    }

    // Person/Group fields need the site-user's numeric ID, not their name.
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
          `Use their exact display name or email as it appears on the SharePoint site.`;
        continue;
      }

      try {
        await patchField(client, list.id, itemId, {
          [`${col.internalName}LookupId`]: col.multi ? resolvedIds : resolvedIds[0],
        });
        verifiedFields[key] = rawValue;
      } catch (e: any) {
        fieldErrors[key] = e?.response?.data?.error?.message || e.message || "Failed to set this field.";
      }
      continue;
    }

    let value: any;
    try {
      value = coerceValue(col, rawValue);
    } catch (coerceError: any) {
      fieldErrors[key] = coerceError.message;
      continue;
    }

    try {
      await patchField(client, list.id, itemId, { [col.internalName]: value });
      verifiedFields[key] = value;
    } catch (firstError: any) {
      if (col.type === "choice" || col.type === "multiChoice") {
        try {
          const arrayValue = Array.isArray(value) ? value : [value];
          await patchField(client, list.id, itemId, { [col.internalName]: arrayValue });
          verifiedFields[key] = arrayValue;
          continue;
        } catch (secondError: any) {
          fieldErrors[key] =
            secondError?.response?.data?.error?.message || secondError.message || "Failed to set this field.";
          continue;
        }
      }
      fieldErrors[key] = firstError?.response?.data?.error?.message || firstError.message || "Failed to set this field.";
    }
  }

  const missingFields = Object.keys(fieldErrors);
  const status = missingFields.length === 0 ? "fully_created" : "partially_created";

  return {
    id: itemId,
    webUrl,
    listName: list.displayName,
    status,
    verifiedFields,
    missingFields,
    fieldErrors,
  };
}
