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

// Choice fields that need special handling
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

  // Separate safe fields (go in POST) from choice fields (PATCH after)
  const safeFields: Record<string, any> = {};
  const choiceFields: Record<string, any> = {};

  for (const [key, value] of Object.entries(allMapped)) {
    if (CHOICE_FIELDS.has(key)) {
      choiceFields[key] = value;
    } else {
      safeFields[key] = value;
    }
  }

  // Step 1: POST to create the item with safe fields
  const res = await client.post(
    `/sites/${SITE_ID}/lists/${listId}/items`,
    { fields: safeFields }
  );
  const itemId = res.data.id;

  // Step 2: Get real column definitions so we use exact internal names
  const colRes = await client.get(
    `/sites/${SITE_ID}/lists/${listId}/columns?$select=name,displayName,type,choice`
  );
  const columns: any[] = colRes.data.value || [];

  // Build a map: field_X -> { internalName, choices[] }
  const choiceColMap: Record<string, { internalName: string; choices: string[] }> = {};
  for (const col of columns) {
    if (col.choice?.choices?.length > 0) {
      choiceColMap[col.name] = {
        internalName: col.name,
        choices: col.choice.choices,
      };
    }
  }

  // Step 3: PATCH each choice field individually with validated value
  const patchResults: Record<string, any> = {};

  for (const [fieldName, value] of Object.entries(choiceFields)) {
    const colInfo = choiceColMap[fieldName];

    // Validate value against allowed choices
    if (colInfo) {
      const match = colInfo.choices.find(
        (c: string) => c.toLowerCase() === String(value).toLowerCase()
      );

      if (!match) {
        patchResults[fieldName] = {
          attempted: value,
          result: "error",
          reason: `Value "${value}" not in allowed choices: [${colInfo.choices.join(", ")}]`,
        };
        continue;
      }

      // Use exactly-matched choice value (correct casing)
      const exactValue = match;

      // Try plain string PATCH
      try {
        await client.patch(
          `/sites/${SITE_ID}/lists/${listId}/items/${itemId}/fields`,
          { [fieldName]: exactValue }
        );
        patchResults[fieldName] = { attempted: exactValue, result: "success" };
      } catch (e1: any) {
        const err1 = e1?.response?.data?.error?.message || e1?.message || String(e1);
        // Try array format
        try {
          await client.patch(
            `/sites/${SITE_ID}/lists/${listId}/items/${itemId}/fields`,
            { [fieldName]: [exactValue] }
          );
          patchResults[fieldName] = { attempted: [exactValue], result: "success" };
        } catch (e2: any) {
          const err2 = e2?.response?.data?.error?.message || e2?.message || String(e2);
          patchResults[fieldName] = {
            attempted: exactValue,
            result: "error",
            reason: `Plain string failed: "${err1}" | Array format failed: "${err2}"`,
          };
        }
      }
    } else {
      // Column not found in schema, try anyway
      try {
        await client.patch(
          `/sites/${SITE_ID}/lists/${listId}/items/${itemId}/fields`,
          { [fieldName]: value }
        );
        patchResults[fieldName] = { attempted: value, result: "success" };
      } catch (e: any) {
        patchResults[fieldName] = {
          attempted: value,
          result: "error",
          reason: e?.response?.data?.error?.message || e?.message,
        };
      }
    }
  }

  // Step 4: Read back what was actually saved
  const verifyRes = await client.get(
    `/sites/${SITE_ID}/lists/${listId}/items/${itemId}/fields`
  );
  const saved = verifyRes.data;

  const verifiedFields = {
    Title: saved.Title || null,
    Status: saved.field_2 || null,
    Priority: saved.field_3 || null,
    DueDate: saved.field_4 || null,
    Description: saved.field_5 || null,
    PercentComplete: saved.field_6 || null,
    TaskCategory: saved.field_9 || null,
    DepartmentName: saved.field_10 || null,
  };

  const failedFields = Object.entries(patchResults)
    .filter(([_, v]) => v.result === "error")
    .map(([k, v]) => `${k}: ${v.reason}`);

  return {
    id: itemId,
    webUrl: `https://dwivedimanshi12outlook.sharepoint.com/Lists/Project%20Tasks/DispForm.aspx?ID=${itemId}`,
    listName,
    status: failedFields.length === 0 ? "fully_created" : "partially_created",
    verifiedFields,
    choiceFieldResults: patchResults,
    errors: failedFields.length > 0 ? failedFields : undefined,
  };
}