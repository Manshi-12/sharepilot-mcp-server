import { getGraphClient } from "../auth/graphClient.js";

const SITE_ID = process.env.SITE_ID || "";

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

  // Project Tasks: hardcoded known-good field map
  // field_2=Status, field_3=Priority, field_9=TaskCategory, field_10=DepartmentName
  // ALL are checkBoxes type → must be sent as string arrays
  const isProjectTasks = listKey === "project tasks";

  let postBody: Record<string, any> = {};

  if (isProjectTasks) {
    // Map ONLY the fields user provided, nothing else
    for (const [key, value] of Object.entries(fields)) {
      const k = key.toLowerCase().replace(/\s/g, "");
      switch (k) {
        case "title":
          postBody["Title"] = String(value); break;
        case "status":
          postBody["field_2"] = [String(value)]; break;
        case "priority":
          postBody["field_3"] = [String(value)]; break;
        case "duedate":
          postBody["field_4"] = String(value); break;
        case "description":
          postBody["field_5"] = String(value); break;
        case "percentcomplete":
        case "%complete":
          postBody["field_6"] = Number(value); break;
        case "isapproved":
          postBody["field_8"] = Boolean(value); break;
        case "taskcategory":
          postBody["field_9"] = [String(value)]; break;
        case "departmentname":
        case "department":
          postBody["field_10"] = [String(value)]; break;
        default:
          postBody[key] = value;
      }
    }
  } else {
    // For other lists: send Title + known fields as plain strings
    for (const [key, value] of Object.entries(fields)) {
      postBody[key] = value;
    }
  }

  // POST
  const res = await client.post(
    `/sites/${SITE_ID}/lists/${listId}/items`,
    { fields: postBody }
  );
  const itemId = res.data.id;

  // Read back to verify
  const verifyRes = await client.get(
    `/sites/${SITE_ID}/lists/${listId}/items/${itemId}/fields`
  );
  const saved = verifyRes.data;

  const unwrap = (v: any) => Array.isArray(v) ? (v.length === 1 ? v[0] : v) : v;

  const verifiedFields: Record<string, any> = {};
  if (isProjectTasks) {
    verifiedFields.Title = saved.Title ?? null;
    verifiedFields.Status = unwrap(saved.field_2) ?? null;
    verifiedFields.Priority = unwrap(saved.field_3) ?? null;
    verifiedFields.DueDate = saved.field_4 ?? null;
    verifiedFields.Description = saved.field_5 ?? null;
    verifiedFields.PercentComplete = saved.field_6 ?? null;
    verifiedFields.TaskCategory = unwrap(saved.field_9) ?? null;
    verifiedFields.DepartmentName = unwrap(saved.field_10) ?? null;
  } else {
    for (const key of Object.keys(fields)) {
      verifiedFields[key] = saved[key] ?? null;
    }
  }

  const missingFields = Object.entries(verifiedFields)
    .filter(([_, v]) => v === null || v === undefined || v === "")
    .map(([k]) => k);

  return {
    id: itemId,
    webUrl: `https://dwivedimanshi12outlook.sharepoint.com/Lists/${listName.replace(/ /g, "%20")}/DispForm.aspx?ID=${itemId}`,
    listName,
    status: missingFields.length === 0 ? "fully_created" : "partially_created",
    verifiedFields,
    missingFields: missingFields.length > 0 ? missingFields : undefined,
    note: missingFields.length === 0
      ? "✅ All fields saved successfully."
      : `⚠️ Fields not saved: ${missingFields.join(", ")}`,
  };
}