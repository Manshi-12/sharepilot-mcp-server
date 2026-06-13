import { getGraphClient } from "../auth/graphClient.js";

const SITE_ID = process.env.SITE_ID || "";
const LIST_ID = process.env.LIST_ID || ""; // "Project Tasks" list ID

export const createListItemToolSchema = {
  name: "create_list_item",
  description:
    "Creates a new item in the 'Project Tasks' SharePoint list using data extracted " +
    "from a document. Fields map to the list's actual columns: Title, Description, " +
    "Priority, Status, Due Date, Project Code, Department Name, Task Category, " +
    "% Complete, Budget, and Is Approved. Only Title is required; other fields are optional.",
  inputSchema: {
    type: "object",
    properties: {
      title: {
        type: "string",
        description: "Task title (maps to the 'Title' column).",
      },
      description: {
        type: "string",
        description: "Task description (maps to 'Description', multi-line text).",
      },
      priority: {
        type: "string",
        description: "Priority value. Must be exactly one of: 'High', 'Medium', 'Low'.",
      },
      status: {
        type: "string",
        description: "Status value. Must be exactly one of: 'Not started', 'In-Progress', 'Completed', 'Blocked'.",
      },
      dueDate: {
        type: "string",
        description: "Due date in YYYY-MM-DD format (maps to 'Due Date').",
      },
      projectCode: {
        type: "string",
        description: "Project code (maps to 'Project Code', single line of text).",
      },
      departmentName: {
        type: "string",
        description: "Department name. Must be exactly one of: 'IT', 'HR', 'Finance', 'Marketing', 'Operations'.",
      },
      taskCategory: {
        type: "string",
        description: "Task category. Must be exactly one of: 'Development', 'Testing', 'Design', 'Documentation', 'Meeting'.",
      },
      percentComplete: {
        type: "number",
        description: "Completion percentage as a number, e.g. 0, 50, 100 (maps to '% Complete').",
      },
      budget: {
        type: "number",
        description: "Budget amount (maps to 'Budget', currency).",
      },
      isApproved: {
        type: "boolean",
        description: "Whether the task is approved (maps to 'Is Approved', Yes/No).",
      },
    },
    required: ["title"],
  },
};

interface CreateListItemInput {
  title: string;
  description?: string;
  priority?: string;
  status?: string;
  dueDate?: string;
  projectCode?: string;
  departmentName?: string;
  taskCategory?: string;
  percentComplete?: number;
  budget?: number;
  isApproved?: boolean;
}

/**
 * Creates a new item in the "Project Tasks" SharePoint list.
 * Returns a cleaned response with just the new item's ID and webUrl (Task 4: optimization).
 *
 * NOTE: Internal column names (e.g. DueDate vs Due_x0020_Date, PercentComplete vs
 * "_x0025_Complete") may differ from these guesses. Verify exact internal names via:
 *   GET /sites/{SITE_ID}/lists/{LIST_ID}/columns
 * and adjust the field keys below if the POST returns a 400 "field not found" error.
 */
export async function createListItem(input: CreateListItemInput) {
  const client = await getGraphClient();

  const fields: Record<string, any> = {
    Title: input.title,
  };

  if (input.description !== undefined) fields["Description"] = input.description;
  if (input.priority !== undefined) fields["Priority"] = input.priority;
  if (input.status !== undefined) fields["Status"] = input.status;
  if (input.dueDate !== undefined) fields["DueDate"] = input.dueDate;
  if (input.projectCode !== undefined) fields["ProjectCode"] = input.projectCode;
  if (input.departmentName !== undefined) fields["DepartmentName"] = input.departmentName;
  if (input.taskCategory !== undefined) fields["TaskCategory"] = input.taskCategory;
  if (input.percentComplete !== undefined) fields["PercentComplete"] = input.percentComplete;
  if (input.budget !== undefined) fields["Budget"] = input.budget;
  if (input.isApproved !== undefined) fields["IsApproved"] = input.isApproved;

  const res = await client.post(`/sites/${SITE_ID}/lists/${LIST_ID}/items`, { fields });

  return {
    id: res.data.id,
    webUrl: res.data.webUrl,
    status: "created",
  };
}