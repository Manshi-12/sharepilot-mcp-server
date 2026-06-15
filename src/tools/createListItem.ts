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

// These are multi-select choice fields — must be sent as arrays
const MULTI_CHOICE_FIELDS = new Set(["field_2", "field_3", "field_9", "field_10"]);

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
      listName: {
        type: "string",
        description: "Name of the SharePoint list, e.g. 'Project Tasks'.",
      },
      fields: {
        type: "object",
        description: "Field name-value pairs using the field names described above.",
        additionalProperties: true,
      },
    },
    required: ["listName", "fields"],
  },
};

export async function createListItem(listName: string, fields: Record<string, any>) {
  const client = await getGraphClient();

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

  // Step 1: Map friendly names → internal SharePoint field names
  const allMapped: Record<string, any> = {};
  for (const [key, value] of Object.entries(fields)) {
    const internalName = FIELD_NAME_MAP[key.toLowerCase()] || key;

    // Multi-select choice fields MUST be arrays in the POST body
    if (MULTI_CHOICE_FIELDS.has(internalName)) {
      allMapped[internalName] = Array.isArray(value) ? value : [value];
    } else {
      allMapped[internalName] = value;
    }
  }

  // Step 2: POST everything in one shot (including choice fields as arrays)
  const res = await client.post(
    `/sites/${SITE_ID}/lists/${listId}/items`,
    { fields: allMapped }
  );

  const itemId = res.data.id;

  // Step 3: Read back what was actually saved to verify
  const verifyRes = await client.get(
    `/sites/${SITE_ID}/lists/${listId}/items/${itemId}/fields`
  );
  const saved = verifyRes.data;

  // Normalize: unwrap single-element arrays for display
  const unwrap = (v: any) => Array.isArray(v) ? (v.length === 1 ? v[0] : v) : v;

  const verifiedFields = {
    Title: saved.Title || null,
    Status: unwrap(saved.field_2) || null,
    Priority: unwrap(saved.field_3) || null,
    DueDate: saved.field_4 || null,
    Description: saved.field_5 || null,
    PercentComplete: saved.field_6 ?? null,
    TaskCategory: unwrap(saved.field_9) || null,
    DepartmentName: unwrap(saved.field_10) || null,
  };

  const missingFields = Object.entries(verifiedFields)
    .filter(([_, v]) => v === null || v === undefined)
    .map(([k]) => k);

  return {
    id: itemId,
    webUrl: `https://dwivedimanshi12outlook.sharepoint.com/Lists/Project%20Tasks/DispForm.aspx?ID=${itemId}`,
    listName,
    status: missingFields.length === 0 ? "fully_created" : "partially_created",
    verifiedFields,
    missingFields: missingFields.length > 0 ? missingFields : undefined,
    note: missingFields.length === 0
      ? "✅ All fields saved successfully."
      : `⚠️ These fields were not provided or saved: ${missingFields.join(", ")}`,
  };
}