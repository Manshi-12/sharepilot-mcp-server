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

const PROBLEMATIC_FIELDS = new Set(["field_2", "field_3", "field_9", "field_10"]);

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

  const allMapped: Record<string, any> = {};
  for (const [key, value] of Object.entries(fields)) {
    const internalName = FIELD_NAME_MAP[key.toLowerCase()] || key;
    allMapped[internalName] = value;
  }

  const safeFields: Record<string, any> = {};
  const skippedFields: Record<string, any> = {};

  for (const [key, value] of Object.entries(allMapped)) {
    if (PROBLEMATIC_FIELDS.has(key)) {
      skippedFields[key] = value;
    } else {
      safeFields[key] = value;
    }
  }

  const res = await client.post(
    `/sites/${SITE_ID}/lists/${listId}/items`,
    { fields: safeFields }
  );

  const itemId = res.data.id;

  const successfulChoices: Record<string, any> = {};
  const failedChoices: Record<string, any> = {};

  for (const [fieldName, value] of Object.entries(skippedFields)) {
    try {
      await client.patch(
        `/sites/${SITE_ID}/lists/${listId}/items/${itemId}/fields`,
        { [fieldName]: value }
      );
      successfulChoices[fieldName] = value;
    } catch {
      try {
        await client.patch(
          `/sites/${SITE_ID}/lists/${listId}/items/${itemId}/fields`,
          { [fieldName]: [value] }
        );
        successfulChoices[fieldName] = [value];
      } catch {
        failedChoices[fieldName] = value;
      }
    }
  }

  return {
    id: itemId,
    webUrl: res.data.webUrl,
    listName: listName,
    fieldsCreated: safeFields,
    choiceFieldsSet: successfulChoices,
    choiceFieldsFailed: failedChoices,
    status: "created",
  };
}