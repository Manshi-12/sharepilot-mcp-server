import { getGraphClient } from "../auth/graphClient.js";

const SITE_ID = process.env.SITE_ID || "";

export const getSiteUsersToolSchema = {
  name: "get_site_users",
  description:
    "Returns all users who have access to this SharePoint site. " +
    "Use when the user asks who has access or wants to list site members.",
  inputSchema: { type: "object", properties: {}, required: [] },
};

export async function getSiteUsers() {
  const client = await getGraphClient();

  // Fetch ALL lists including hidden ones to find User Information List
  const res = await client.get(`/sites/${SITE_ID}/lists`, {
    params: { "$select": "id,name,displayName", "$top": 200 },
  });

  const allLists = res.data.value || [];

  // Try multiple possible names/slugs across tenant languages
  const userList = allLists.find((l: any) =>
    ["users", "user information list", "userinformationlist"].includes(
      (l.name || "").toLowerCase()
    ) ||
    ["users", "user information list", "userinformationlist"].includes(
      (l.displayName || "").toLowerCase()
    )
  );

  if (!userList) {
    // Last resort: dump list names for debugging
    const names = allLists.map((l: any) => `${l.name} / ${l.displayName}`).join(", ");
    throw new Error(`Could not find User Information List. Lists found: ${names}`);
  }

  const itemsRes = await client.get(`/sites/${SITE_ID}/lists/${userList.id}/items`, {
    params: {
      "$expand": "fields($select=Title,EMail,IsSiteAdmin)",
      "$top": 500,
    },
  });

  const users = (itemsRes.data.value || [])
    .filter((u: any) => u.fields?.Title && u.fields?.EMail)
    .map((u: any) => ({
      name: u.fields.Title,
      email: u.fields.EMail,
      isAdmin: u.fields.IsSiteAdmin || false,
    }));

  return { totalUsers: users.length, users };
}