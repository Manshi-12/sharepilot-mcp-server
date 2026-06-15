import { getGraphClient } from "../auth/graphClient.js";

const SITE_ID = process.env.SITE_ID || "";

// Hardcoded map ONLY for Project Tasks (internal names differ from display names)
const PROJECT_TASKS_FIELD_MAP: Record<string, string> = {
  "title": "Title",
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
};

// System fields that SharePoint manages internally — never write to these
const SYSTEM_FIELDS = new Set([
  "ID", "Id", "id",
  "ContentType", "ContentTypeId",
  "Modified", "Created",
  "Author", "AuthorId",
  "Editor", "EditorId",
  "AppAuthor", "AppAuthorId",
  "AppEditor", "AppEditorId",
  "_UIVersionString", "_UIVersion",
  "Attachments", "Edit",
  "LinkTitle", "LinkTitleNoMenu", "LinkFilename",
  "DocIcon", "ItemChildCount", "FolderChildCount",
  "_ColorTag", "ComplianceAssetId",
  "_ComplianceFlags", "_ComplianceTag",
  "_ComplianceTagWrittenTime", "_ComplianceTagUserId",
  "_IsRecord",
]);

const KNOWN_LISTS: Record<string, string> = {
  "project tasks": "066a3b58-72a3-4fba-a3fc-3acae90be4bf",
};

export const createListItemToolSchema = {
  name: "create_list_item",
  description:
    "Creates a new item/row in ANY SharePoint List. " +
    "For 'Project Tasks' list, supported fields: " +
    "Title (required), Description, DueDate (YYYY-MM-DD), PercentComplete (0-100), IsApproved (true/false), " +
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

  // ── 1. Resolve list ID ───────────────────────────────────────────────────────
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

  // ── 2. Fetch column schema — only writable, non-hidden columns ───────────────
  const colRes = await client.get(`/sites/${SITE_ID}/lists/${listId}/columns`);
  const columns: any[] = colRes.data.value || [];

  // Only keep user-editable columns
  const writableColumns = columns.filter(
    (c: any) => !c.readOnly && !c.hidden && !SYSTEM_FIELDS.has(c.name)
  );

  // Map: internalName → choice schema
  const choiceSchema: Record<string, { displayAs: string; choices: string[] }> = {};
  for (const col of writableColumns) {
    if (col.choice) {
      choiceSchema[col.name] = {
        displayAs: col.choice.displayAs || "dropDownMenu",
        choices: col.choice.choices || [],
      };
    }
  }

  // Map: displayName.toLowerCase() → internalName (only for writable columns)
  const displayToInternal: Record<string, string> = {};
  for (const col of writableColumns) {
    displayToInternal[col.displayName.toLowerCase()] = col.name;
    displayToInternal[col.name.toLowerCase()] = col.name;
  }

  // ── 3. Map user fields → internal names, skip system fields ─────────────────
  const isProjectTasks = listKey === "project tasks";
  const mappedFields: Record<string, any> = {};

  for (const [userKey, value] of Object.entries(fields)) {
    // Skip if user accidentally passed a system field
    if (SYSTEM_FIELDS.has(userKey)) continue;

    let internalName: string;
    if (isProjectTasks) {
      internalName = PROJECT_TASKS_FIELD_MAP[userKey.toLowerCase()] ?? userKey;
    } else {
      internalName = displayToInternal[userKey.toLowerCase()] ?? userKey;
    }

    // Skip system fields even after mapping
    if (SYSTEM_FIELDS.has(internalName)) continue;

    // Apply correct format based on choice schema
    const schema = choiceSchema[internalName];
    if (schema) {
      if (schema.displayAs === "checkBoxes") {
        // Multi-select: MUST be array of strings
        mappedFields[internalName] = Array.isArray(value)
          ? value.map(String)
          : [String(value)];
      } else {
        // Single-select (dropDownMenu): MUST be plain string, NOT array
        mappedFields[internalName] = Array.isArray(value)
          ? String(value[0])
          : String(value);
      }
    } else {
      // Regular field: send as-is
      mappedFields[internalName] = value;
    }
  }

  // ── 4. POST all fields in one request ───────────────────────────────────────
  const res = await client.post(
    `/sites/${SITE_ID}/lists/${listId}/items`,
    { fields: mappedFields }
  );
  const itemId = res.data.id;

  // ── 5. Verify: read back what was actually saved ─────────────────────────────
  const verifyRes = await client.get(
    `/sites/${SITE_ID}/lists/${listId}/items/${itemId}/fields`
  );
  const saved = verifyRes.data;

  // Build verified map using user-friendly keys
  const verifiedFields: Record<string, any> = {};
  for (const [userKey] of Object.entries(fields)) {
    if (SYSTEM_FIELDS.has(userKey)) continue;

    let internalName: string;
    if (isProjectTasks) {
      internalName = PROJECT_TASKS_FIELD_MAP[userKey.toLowerCase()] ?? userKey;
    } else {
      internalName = displayToInternal[userKey.toLowerCase()] ?? userKey;
    }

    const raw = saved[internalName];
    verifiedFields[userKey] =
      Array.isArray(raw) && raw.length === 1 ? raw[0] : (raw ?? null);
  }

  const missingFields = Object.entries(verifiedFields)
    .filter(([_, v]) => v === null || v === undefined || v === "")
    .map(([k]) => k);

  const sharePointUrl = `https://dwivedimanshi12outlook.sharepoint.com/Lists/${listName.replace(/ /g, "%20")}/DispForm.aspx?ID=${itemId}`;

  return {
    id: itemId,
    webUrl: sharePointUrl,
    listName,
    status: missingFields.length === 0 ? "fully_created" : "partially_created",
    verifiedFields,
    missingFields: missingFields.length > 0 ? missingFields : undefined,
    note: missingFields.length === 0
      ? "✅ All fields saved successfully."
      : `⚠️ Fields not saved: ${missingFields.join(", ")}`,
  };
}