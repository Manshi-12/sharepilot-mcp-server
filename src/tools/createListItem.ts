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

// These choice fields MUST be PATCHed after the initial POST
const CHOICE_FIELDS = new Set(["field_2", "field_3", "field_9", "field_10"]);

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
    if (!found) {
      throw new Error(
        `List "${listName}" not found. Available: ${lists.map((l: any) => l.displayName).join(", ")}`
      );
    }
    listId = found.id;
  }

  // Map all friendly field names to internal SharePoint names
  const allMapped: Record<string, any> = {};
  for (const [key, value] of Object.entries(fields)) {
    const internalName = FIELD_NAME_MAP[key.toLowerCase()] || key;
    allMapped[internalName] = value;
  }

  // Separate safe fields (POST) from choice fields (PATCH)
  const safeFields: Record<string, any> = {};
  const choiceFields: Record<string, any> = {};

  for (const [key, value] of Object.entries(allMapped)) {
    if (CHOICE_FIELDS.has(key)) {
      choiceFields[key] = value;
    } else {
      safeFields[key] = value;
    }
  }

  // Step 1: POST to create the item with safe fields only
  const res = await client.post(
    `/sites/${SITE_ID}/lists/${listId}/items`,
    { fields: safeFields }
  );

  const itemId = res.data.id;
  const errors: string[] = [];

  // Step 2: PATCH all choice fields together in ONE request (not one-by-one)
  if (Object.keys(choiceFields).length > 0) {
    try {
      await client.patch(
        `/sites/${SITE_ID}/lists/${listId}/items/${itemId}/fields`,
        choiceFields
      );
    } catch (err: any) {
      // If batch PATCH fails, try each field individually
      for (const [fieldName, value] of Object.entries(choiceFields)) {
        let fieldSet = false;
        // Try as plain string first
        try {
          await client.patch(
            `/sites/${SITE_ID}/lists/${listId}/items/${itemId}/fields`,
            { [fieldName]: value }
          );
          fieldSet = true;
        } catch {
          // Try as array
          try {
            await client.patch(
              `/sites/${SITE_ID}/lists/${listId}/items/${itemId}/fields`,
              { [fieldName]: [value] }
            );
            fieldSet = true;
          } catch (e2: any) {
            const errMsg = e2?.response?.data?.error?.message || e2?.message || String(e2);
            errors.push(`${fieldName}="${value}" failed: ${errMsg}`);
          }
        }
      }
    }
  }

  // Verify what was actually saved by reading back the item
  let verifiedFields: Record<string, any> = {};
  try {
    const verifyRes = await client.get(
      `/sites/${SITE_ID}/lists/${listId}/items/${itemId}/fields`
    );
    const data = verifyRes.data;
    verifiedFields = {
      Title: data.Title,
      Status: data.field_2,
      Priority: data.field_3,
      DueDate: data.field_4,
      Description: data.field_5,
      PercentComplete: data.field_6,
      TaskCategory: data.field_9,
      DepartmentName: data.field_10,
    };
  } catch {
    // verification failed, not critical
  }

  // Build honest status
  const missing = Object.entries(verifiedFields)
    .filter(([_, v]) => v === null || v === undefined || v === "")
    .map(([k]) => k);

  return {
    id: itemId,
    webUrl: `https://dwivedimanshi12outlook.sharepoint.com/Lists/Project%20Tasks/DispForm.aspx?ID=${itemId}`,
    listName: listName,
    status: errors.length === 0 && missing.length === 0 ? "fully_created" : "partially_created",
    verifiedFields,
    fieldErrors: errors.length > 0 ? errors : undefined,
    missingFields: missing.length > 0 ? missing : undefined,
    note: missing.length > 0
      ? `⚠️ These fields were NOT saved: ${missing.join(", ")}. The item was created but is incomplete.`
      : "✅ All fields saved successfully.",
  };
}