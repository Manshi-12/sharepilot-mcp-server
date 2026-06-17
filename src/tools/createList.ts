import { getGraphClient } from "../auth/graphClient.js";

const SITE_ID = process.env.SITE_ID!;

// ─── Tool Schema ────────────────────────────────────────────────────────────

export const createListToolSchema = {
    name: "create_list",
    description:
        "Creates a new SharePoint List or Document Library on the site with custom columns. " +
        "Set template to 'genericList' for a regular list, or 'documentLibrary' for a library.",
    inputSchema: {
        type: "object",
        properties: {
            displayName: {
                type: "string",
                description: "The display name of the new list or library (e.g. 'Project Tracker')",
            },
            description: {
                type: "string",
                description: "Optional description for the list/library",
            },
            template: {
                type: "string",
                enum: ["genericList", "documentLibrary"],
                description:
                    "'genericList' creates a standard list. 'documentLibrary' creates a document library.",
            },
            columns: {
                type: "array",
                description: "Custom columns to add to the list after creation.",
                items: {
                    type: "object",
                    properties: {
                        name: {
                            type: "string",
                            description: "Internal name for the column (no spaces, e.g. 'ProjectStatus')",
                        },
                        type: {
                            type: "string",
                            enum: ["text", "number", "boolean", "dateTime", "choice", "personOrGroup"],
                            description: "Column data type",
                        },
                        description: {
                            type: "string",
                            description: "Optional description for the column",
                        },
                        choices: {
                            type: "array",
                            items: { type: "string" },
                            description: "Required if type is 'choice'. List of choice values.",
                        },
                    },
                    required: ["name", "type"],
                },
            },
        },
        required: ["displayName", "template"],
    },
};

// ─── Column builder ──────────────────────────────────────────────────────────

function buildColumnDefinition(col: {
    name: string;
    type: string;
    description?: string;
    choices?: string[];
}) {
    const base: any = {
        name: col.name,
        description: col.description ?? "",
    };

    switch (col.type) {
        case "text":
            base.text = {};
            break;
        case "number":
            base.number = {};
            break;
        case "boolean":
            base.boolean = {};
            break;
        case "dateTime":
            base.dateTime = { format: "dateTime" };
            break;
        case "choice":
            base.choice = {
                allowTextEntry: false,
                choices: col.choices ?? [],
                displayAs: "dropDownMenu",
            };
            break;
        case "personOrGroup":
            base.personOrGroup = { allowMultipleSelection: false, displayAs: "nameWithPresence" };
            break;
        default:
            base.text = {};
    }

    return base;
}

// ─── Main function ───────────────────────────────────────────────────────────

export async function createList(
    displayName: string,
    template: "genericList" | "documentLibrary",
    description?: string,
    columns?: { name: string; type: string; description?: string; choices?: string[] }[]
) {
    const client = await getGraphClient(); // <-- await here, returns AxiosInstance

    // 1. Create the list/library
    const listBody: any = {
        displayName,
        description: description ?? "",
        list: {
            template,
        },
    };

    const createResponse = await client.post(`/sites/${SITE_ID}/lists`, listBody);
    const created = createResponse.data;

    const listId: string = created.id;
    const results: any[] = [];

    // 2. Add custom columns if provided
    if (columns && columns.length > 0) {
        for (const col of columns) {
            try {
                const colDef = buildColumnDefinition(col);
                const colResponse = await client.post(
                    `/sites/${SITE_ID}/lists/${listId}/columns`,
                    colDef
                );
                results.push({ column: col.name, status: "created", id: colResponse.data.id });
            } catch (err: any) {
                results.push({
                    column: col.name,
                    status: "failed",
                    error: err?.response?.data ?? err.message,
                });
            }
        }
    }

    return {
        success: true,
        listId,
        displayName: created.displayName,
        webUrl: created.webUrl,
        template,
        columnsAdded: results,
    };
}