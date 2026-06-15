import { getGraphClient } from "../auth/graphClient.js";

const SITE_ID = process.env.SITE_ID || "";

// This map is ONLY for Project Tasks list where display names differ from internal names
// For all other lists, we use the field name as-is and let SharePoint match it
const PROJECT_TASKS_FIELD_MAP: Record<string, string> = {
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
    "Creates a new item/row in ANY SharePoint List. " +
    "For 'Project Tasks' list use these fields: " +
    "Title (text, required), Description (text), DueDate (YYYY-MM-DD), " +
    "PercentComplete (0-100), IsApproved (true/false), " +
    "Status ('Not started'|'In-Progress'|'Completed'|'Blocked'), " +
    "Priority ('High'|'Medium'|'Low'), " +
    "TaskCategory ('Development'|'Testing'|'Design'|'Documentation'|'Meeting'), " +
    "DepartmentName ('IT'|'HR'|'Finance'|'Marketing'|'Operations'). " +
    "For other lists, use the exact column internal names as field keys.",
  inputSchema: {
    type: "object",
    properties: {
      listName: { type: "string", description: "Display name of the SharePoint list." },
      fields: { type: "object", description: "Field name-value pairs.", additionalProperties: true },
    },
    required: ["listName", "fields"],
  },
};

export async function createListItem(listName: string, fields: Record<string, any>) {
  const client = await getGraphClient();

  // ── 1. Resolve list ID ──────────────────────────────────────────────────────
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

  // ── 2. Fetch real column definitions ────────────────────────────────────────
  const colRes = await client.get(`/sites/${SITE_ID}/lists/${listId}/columns`);
  const columns: any[] = colRes.data.value || [];

  // Build lookup: internalName → { displayAs, choices }
  const colSchema: Record<string, { displayAs: string; choices: string[] }> = {};
  for (const col of columns) {
    if (col.choice) {
      colSchema[col.name] = {
        displayAs: col.choice.displayAs || "dropDownMenu",
        choices: col.choice.choices || [],
      };
    }
  }

  // Also build lookup: displayName.toLowerCase() → internalName
  // So if user passes "Task Category" we find "field_9" or "TaskCategory"
  const displayToInternal: Record<string, string> = {};
  for (const col of columns) {
    displayToInternal[col.displayName.toLowerCase()] = col.name;
    displayToInternal[col.name.toLowerCase()] = col.name; // also map internalName → itself
  }

  // ── 3. Map user-provided field names → correct internal names ───────────────
  const isProjectTasks = listKey === "project tasks";

  const mappedFields: Record<string, any> = {};
  for (const [userKey, value] of Object.entries(fields)) {
    let internalName: string;

    if (isProjectTasks) {
      // Use the hardcoded map for Project Tasks (field_2, field_3, etc.)
      internalName = PROJECT_TASKS_FIELD_MAP[userKey.toLowerCase()] || userKey;
    } else {
      // For all other lists: match by displayName or internalName from real schema
      internalName = displayToInternal[userKey.toLowerCase()] || userKey;
    }

    // Check how this field expects its value
    const schema = colSchema[internalName];
    if (schema) {
      if (schema.displayAs === "checkBoxes") {
        // checkBoxes = multi-select → MUST be array
        mappedFields[internalName] = Array.isArray(value) ? value : [String(value)];
      } else {
        // dropDownMenu = single-select → MUST be plain string
        mappedFields[internalName] = Array.isArray(value) ? value[0] : String(value);
      }
    } else {
      // Not a choice field, send as-is
      mappedFields[internalName] = value;
    }
  }

  // ── 4. POST everything in one request ───────────────────────────────────────
  const res = await client.post(
    `/sites/${SITE_ID}/lists/${listId}/items`,
    { fields: mappedFields }
  );
  const itemId = res.data.id;

  // ── 5. Read back to verify ──────────────────────────────────────────────────
  const verifyRes = await client.get(
    `/sites/${SITE_ID}/lists/${listId}/items/${itemId}/fields`
  );
  const saved = verifyRes.data;

  // Build verified snapshot for only the fields the user tried to set
  const verifiedFields: Record<string, any> = {};
  for (const [userKey] of Object.entries(fields)) {
    let internalName: string;
    if (isProjectTasks) {
      internalName = PROJECT_TASKS_FIELD_MAP[userKey.toLowerCase()] || userKey;
    } else {
      internalName = displayToInternal[userKey.toLowerCase()] || userKey;
    }
    const raw = saved[internalName];
    // Unwrap single-element arrays for clean display
    verifiedFields[userKey] = Array.isArray(raw) && raw.length === 1 ? raw[0] : (raw ?? null);
  }

  const missingFields = Object.entries(verifiedFields)
    .filter(([_, v]) => v === null || v === undefined || v === "")
    .map(([k]) => k);

  return {
    id: itemId,
    webUrl: `https://dwivedimanshi12outlook.sharepoint.com/Lists/${encodeURIComponent(listName.replace(/ /g, "%20"))}/DispForm.aspx?ID=${itemId}`,
    listName,
    status: missingFields.length === 0 ? "fully_created" : "partially_created",
    verifiedFields,
    missingFields: missingFields.length > 0 ? missingFields : undefined,
    note: missingFields.length === 0
      ? "✅ All fields saved successfully."
      : `⚠️ Fields not saved: ${missingFields.join(", ")}`,
  };
}