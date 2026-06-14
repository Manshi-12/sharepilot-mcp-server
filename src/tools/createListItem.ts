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
    "Creates a new item/row in a SharePoint List on the site. " +
    "For 'Project Tasks' list, use these field names: " +
    "Title, Status, Priority, DueDate, Description, PercentComplete, IsApproved, TaskCategory, DepartmentName. " +
    "Valid values — Priority: 'High'|'Medium'|'Low'. " +
    "Status: 'Not started'|'In-Progress'|'Completed'|'Blocked'. " +
    "DepartmentName: 'IT'|'HR'|'Finance'|'Marketing'|'Operations'. " +
    "TaskCategory: 'Development'|'Testing'|'Design'|'Documentation'|'Meeting'.",
  inputSchema: {
    type: "object",
    properties: {
      listName: {
        type: "string",
        description: "Name of the SharePoint list, e.g. 'Project Tasks'.",
      },
      fields: {
        type: "object",
        description: "Field name-value pairs. Use: Title, Status, Priority, DueDate, Description, PercentComplete, IsApproved, TaskCategory, DepartmentName.",
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

  // Map display field names → internal SharePoint field names
  const mappedFields: Record<string, any> = {};
  for (const [key, value] of Object.entries(fields)) {
    const internalName = FIELD_NAME_MAP[key.toLowerCase()] || key;
    mappedFields[internalName] = value;
  }

  const res = await client.post(
    `/sites/${SITE_ID}/lists/${listId}/items`,
    { fields: mappedFields }
  );

  return {
    id: res.data.id,
    webUrl: res.data.webUrl,
    listName: listName,
    fieldsCreated: mappedFields,
    status: "created",
  };
}