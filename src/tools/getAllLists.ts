import { getGraphClient } from "../auth/graphClient.js";

export const getAllListsToolSchema = {
  name: "get_all_lists",
  description: "Returns all SharePoint lists and document libraries available on the site.",
  inputSchema: { type: "object", properties: {}, required: [] },
};

export async function getAllLists() {
  const client = await getGraphClient();
  const res = await client.get(
    `/sites/${process.env.SITE_ID}/lists?$select=id,name,displayName,list&$top=50`
  );
  return res.data.value.map((l: any) => ({
    name: l.displayName,
    type: l.list?.template === "documentLibrary" ? "Document Library" : "List",
  }));
}