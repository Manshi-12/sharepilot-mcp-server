import { getGraphClient } from "../auth/graphClient.js";
import { resolveList, getListColumns, findColumn, resolvePersonId, ColumnInfo } from "../utils/resolve.js";
import { coerceValue } from "../utils/coerce.js";

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

async function patchField(
  client: any,
  listId: string,
  itemId: string,
  body: Record<string, any>,
  extraHeaders?: Record<string, string>
): Promise<void> {
  await client.patch(
    `/sites/${SITE_ID}/lists/${listId}/items/${itemId}/fields`,
    body,
    extraHeaders ? { headers: extraHeaders } : undefined
  );
}

export async function createListItem(listName: string, fields: Record<string, any>) {
  const client = await getGraphClient();

  const list = await resolveList(client, listName);
  const columns = await getListColumns(client, list.id);

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
      `Could not create the item (failed even with just Title set): ` +
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

    // Fix #coerce — validate choices before hitting SharePoint
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
      const extraHeaders =
        col.type === "hyperlinkOrPicture" ? { Prefer: "apiversion=2.1" } : undefined;

      await patchField(client, list.id, itemId, { [col.internalName]: value }, extraHeaders);
      verifiedFields[key] = value;
    } catch (firstError: any) {
      // Checkbox-style choice columns may need array — retry once
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