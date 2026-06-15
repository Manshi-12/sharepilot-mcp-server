import { getGraphClient } from "../auth/graphClient.js";

const SITE_ID = process.env.SITE_ID || "";

const FIELD_NAME_MAP: Record<string, string> = {
  "status": "field_2",
  "priority": "field_3",
  "duedate": "field_4",
  "due date": "field_4",
  "description": "field_5",
  "percentcomplete": "field_6",
  "percent complete": "field_6",
  "%complete": "field_6",
  "isapproved": "field_8",
  "is approved": "field_8",
  "taskcategory": "field_9",
  "task category": "field_9",
  "departmentname": "field_10",
  "department name": "field_10",
  "department": "field_10",
  "title": "Title",
};

const KNOWN_LISTS: Record<string, string> = {
  "project tasks": "066a3b58-72a3-4fba-a3fc-3acae90be4bf",
};

export const createListItemToolSchema = {
  name: "create_list_item",
  description:
    "Creates a new item/row in a SharePoint List. " +
    "For 'Project Tasks' list use these fields: " +
    "Title (text, required), Description (text), DueDate (YYYY-MM-DD), " +
    "PercentComplete (0-100), IsApproved (true/false), " +
    "Status ('Not started'|'In-Progress'|'Completed'|'Blocked'), " +
    "Priority ('High'|'Medium'|'Low'), " +
    "TaskCategory ('Development'|'Testing'|'Design'|'Documentation'|'Meeting'), " +
    "DepartmentName ('IT'|'HR'|'Finance'|'Marketing'|'Operations').",
  inputSchema: {
    type: "object",
    properties: {
      listName: { type: "string", description: "Name of the SharePoint list." },
      fields: { type: "object", description: "Field name-value pairs.", additionalProperties: true },
    },
    required: ["listName", "fields"],
  },
};

export async function createListItem(listName: string, fields: Record<string, any>) {
  const client = await getGraphClient();

  // --- Resolve list ID ---
  const listKey = listName.toLowerCase();
  let listId = KNOWN_LISTS[listKey];

  if (!listId) {
    const listsRes = await client.get(`/sites/${SITE_ID}/lists?$select=id,name,displayName`);
    const lists = listsRes.data.value || [];
    const found = lists.find(
      (l: any) =>
        l.displayName?.toLowerCase() === listKey ||
        l.name?.toLowerCase() === listKey
    );
    if (!found) throw new Error(`List "${listName}" not found.`);
    listId = found.id;
  }

  // --- Fetch real column definitions to know which fields are multi-choice ---
  const colRes = await client.get(
    `/sites/${SITE_ID}/lists/${listId}/columns?$select=name,displayName,type,choice`
  );
  const columns: any[] = colRes.data.value || [];

  // Build set of internal names that are multi-select choice fields
  const multiChoiceFields = new Set<string>();
  const singleChoiceFields = new Set<string>();
  for (const col of columns) {
    if (col.choice) {
      if (col.choice.displayAs === "checkBoxes") {
        multiChoiceFields.add(col.name);
      } else {
        singleChoiceFields.add(col.name);
      }
    }
  }

  // --- Map friendly names → internal names, wrap multi-choice as arrays ---
  const mappedFields: Record<string, any> = {};
  for (const [key, value] of Object.entries(fields)) {
    const internalName = FIELD_NAME_MAP[key.toLowerCase()] || key;

    if (multiChoiceFields.has(internalName)) {
      // Multi-select: must be array
      mappedFields[internalName] = Array.isArray(value) ? value : [String(value)];
    } else if (singleChoiceFields.has(internalName)) {
      // Single-select: must be plain string
      mappedFields[internalName] = Array.isArray(value) ? value[0] : String(value);
    } else {
      mappedFields[internalName] = value;
    }
  }

  // --- POST all fields in one request ---
  let itemId: string;
  let postError: string | null = null;

  try {
    const res = await client.post(
      `/sites/${SITE_ID}/lists/${listId}/items`,
      { fields: mappedFields }
    );
    itemId = res.data.id;
  } catch (postErr: any) {
    postError = postErr?.response?.data?.error?.message || postErr?.message || String(postErr);

    // Fallback: POST only Title, then PATCH rest one by one
    const safePost: Record<string, any> = { Title: mappedFields["Title"] || "Untitled" };
    const fallbackRes = await client.post(
      `/sites/${SITE_ID}/lists/${listId}/items`,
      { fields: safePost }
    );
    itemId = fallbackRes.data.id;

    // PATCH each field individually and log results
    for (const [fieldName, value] of Object.entries(mappedFields)) {
      if (fieldName === "Title") continue;
      try {
        await client.patch(
          `/sites/${SITE_ID}/lists/${listId}/items/${itemId}/fields`,
          { [fieldName]: value }
        );
      } catch (patchErr: any) {
        // log but continue
        console.error(`PATCH failed for ${fieldName}:`, patchErr?.response?.data);
      }
    }
  }

  // --- Read back to verify ---
  const verifyRes = await client.get(
    `/sites/${SITE_ID}/lists/${listId}/items/${itemId}/fields`
  );
  const saved = verifyRes.data;

  // Build verified snapshot of only the fields we tried to set
  const verifiedFields: Record<string, any> = {};
  for (const [key] of Object.entries(fields)) {
    const internalName = FIELD_NAME_MAP[key.toLowerCase()] || key;
    const rawVal = saved[internalName];
    verifiedFields[key] = Array.isArray(rawVal) && rawVal.length === 1 ? rawVal[0] : rawVal ?? null;
  }

  const missingFields = Object.entries(verifiedFields)
    .filter(([_, v]) => v === null || v === undefined || v === "")
    .map(([k]) => k);

  return {
    id: itemId,
    webUrl: `https://dwivedimanshi12outlook.sharepoint.com/Lists/${encodeURIComponent(listName)}/DispForm.aspx?ID=${itemId}`,
    listName,
    status: missingFields.length === 0 ? "fully_created" : "partially_created",
    verifiedFields,
    missingFields: missingFields.length > 0 ? missingFields : undefined,
    postError: postError ?? undefined,
    note: missingFields.length === 0
      ? "✅ All fields saved successfully."
      : `⚠️ Fields not saved: ${missingFields.join(", ")}. postError: ${postError}`,
  };
}