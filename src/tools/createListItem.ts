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

  // Map friendly names to internal SharePoint field names
  const allMapped: Record<string, any> = {};
  for (const [key, value] of Object.entries(fields)) {
    const internalName = FIELD_NAME_MAP[key.toLowerCase()] || key;
    allMapped[internalName] = value;
  }

  // --- STRATEGY: Send ALL fields together in the initial POST ---
  // SharePoint accepts choice fields in the POST body directly.
  // The two-step PATCH approach fails for checkbox-style choice fields.
  try {
    const res = await client.post(
      `/sites/${SITE_ID}/lists/${listId}/items`,
      { fields: allMapped }
    );

    const itemId = res.data.id;

    // Read back to verify what was actually saved
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
      PercentComplete: saved.field_6 ?? null,
      TaskCategory: saved.field_9 || null,
      DepartmentName: saved.field_10 || null,
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
      note: missingFields.length > 0
        ? `⚠️ These fields were not saved: ${missingFields.join(", ")}`
        : "✅ All fields saved successfully.",
    };

  } catch (postErr: any) {
    // If POST with all fields fails, fall back to safe fields only + PATCH approach
    // but this time capture the EXACT error for each choice field
    const postErrMsg = postErr?.response?.data?.error?.message || postErr?.message || String(postErr);
    const postErrCode = postErr?.response?.data?.error?.code || "unknown";

    // Try POST with only safe (non-choice) fields
    const CHOICE_FIELDS = new Set(["field_2", "field_3", "field_9", "field_10"]);
    const safeFields: Record<string, any> = {};
    const choiceFields: Record<string, any> = {};

    for (const [key, value] of Object.entries(allMapped)) {
      if (CHOICE_FIELDS.has(key)) choiceFields[key] = value;
      else safeFields[key] = value;
    }

    const res2 = await client.post(
      `/sites/${SITE_ID}/lists/${listId}/items`,
      { fields: safeFields }
    );
    const itemId = res2.data.id;

    // Try PATCH for each choice field and record exact error
    const patchLog: Record<string, any> = {};
    for (const [fieldName, value] of Object.entries(choiceFields)) {
      // Attempt 1: plain string
      try {
        await client.patch(
          `/sites/${SITE_ID}/lists/${listId}/items/${itemId}/fields`,
          { [fieldName]: value }
        );
        patchLog[fieldName] = { result: "success", value };
        continue;
      } catch (e1: any) {
        const msg1 = e1?.response?.data?.error?.message || e1?.message;
        const code1 = e1?.response?.data?.error?.code || "";

        // Attempt 2: array
        try {
          await client.patch(
            `/sites/${SITE_ID}/lists/${listId}/items/${itemId}/fields`,
            { [fieldName]: [value] }
          );
          patchLog[fieldName] = { result: "success_as_array", value: [value] };
          continue;
        } catch (e2: any) {
          const msg2 = e2?.response?.data?.error?.message || e2?.message;

          // Attempt 3: object format { Value: "..." }
          try {
            await client.patch(
              `/sites/${SITE_ID}/lists/${listId}/items/${itemId}/fields`,
              { [fieldName]: { Value: value } }
            );
            patchLog[fieldName] = { result: "success_as_object", value: { Value: value } };
            continue;
          } catch (e3: any) {
            const msg3 = e3?.response?.data?.error?.message || e3?.message;
            patchLog[fieldName] = {
              result: "ALL_ATTEMPTS_FAILED",
              attempt1_string: msg1,
              attempt2_array: msg2,
              attempt3_object: msg3,
            };
          }
        }
      }
    }

    // Read back what was actually saved
    const verifyRes = await client.get(
      `/sites/${SITE_ID}/lists/${listId}/items/${itemId}/fields`
    );
    const saved = verifyRes.data;

    return {
      id: itemId,
      webUrl: `https://dwivedimanshi12outlook.sharepoint.com/Lists/Project%20Tasks/DispForm.aspx?ID=${itemId}`,
      listName,
      status: "partially_created",
      initialPostError: { code: postErrCode, message: postErrMsg },
      verifiedFields: {
        Title: saved.Title || null,
        Status: saved.field_2 || null,
        Priority: saved.field_3 || null,
        DueDate: saved.field_4 || null,
        Description: saved.field_5 || null,
        PercentComplete: saved.field_6 ?? null,
        TaskCategory: saved.field_9 || null,
        DepartmentName: saved.field_10 || null,
      },
      choicePatchLog: patchLog,
      note: "⚠️ Initial POST with all fields failed. Fell back to safe-fields POST + per-field PATCH. See choicePatchLog for exact errors.",
    };
  }
}