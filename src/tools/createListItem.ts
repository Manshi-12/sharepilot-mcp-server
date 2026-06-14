import { getGraphClient } from "../auth/graphClient.js";

const SITE_ID = process.env.SITE_ID || "";

export const createListItemToolSchema = {
  name: "create_list_item",
  description:
    "Creates a new item/row in any SharePoint List on the site. " +
    "Specify the list name and provide field values as key-value pairs. " +
    "Works with any list on the site — e.g. 'Project Tasks', 'Events', etc.",
  inputSchema: {
    type: "object",
    properties: {
      listName: {
        type: "string",
        description: "Name of the SharePoint list to create the item in, e.g. 'Project Tasks'.",
      },
      fields: {
        type: "object",
        description:
          "Key-value pairs of field/column names and their values. " +
          "For 'Project Tasks' list the available fields are: " +
          "Title (text), Description (text), Priority ('High'|'Medium'|'Low'), " +
          "Status ('Not started'|'In-Progress'|'Completed'|'Blocked'), " +
          "DueDate (YYYY-MM-DD), ProjectCode (text), " +
          "DepartmentName ('IT'|'HR'|'Finance'|'Marketing'|'Operations'), " +
          "TaskCategory ('Development'|'Testing'|'Design'|'Documentation'|'Meeting'), " +
          "PercentComplete (number 0-100), Budget (number), IsApproved (true/false).",
        additionalProperties: true,
      },
    },
    required: ["listName", "fields"],
  },
};

export async function createListItem(listName: string, fields: Record<string, any>) {
  const client = await getGraphClient();

  const listsRes = await client.get(`/sites/${SITE_ID}/lists?$select=id,name,displayName`);
  const lists = listsRes.data.value || [];

  const targetList = lists.find(
    (l: any) =>
      l.displayName?.toLowerCase() === listName.toLowerCase() ||
      l.name?.toLowerCase() === listName.toLowerCase()
  );

  if (!targetList) {
    throw new Error(
      `List "${listName}" not found. Available lists: ${lists.map((l: any) => l.displayName).join(", ")}`
    );
  }

  const res = await client.post(`/sites/${SITE_ID}/lists/${targetList.id}/items`, { fields });

  return {
    id: res.data.id,
    webUrl: res.data.webUrl,
    listName: listName,
    status: "created",
  };
}