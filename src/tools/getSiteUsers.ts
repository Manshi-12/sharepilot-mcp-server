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

  // Use SharePoint siteUsers endpoint — no Group.Read.All needed
  const res = await client.get(`/sites/${SITE_ID}/lists`, {
    params: { "$select": "id,name", "$top": 100 },
  });

  // Find the hidden "users" list (User Information List internal name is always "users")
  const userList = (res.data.value || []).find(
    (l: any) => l.name === "users"
  );

  if (!userList) {
    throw new Error("Could not locate the User Information List on this site.");
  }

  const itemsRes = await client.get(`/sites/${SITE_ID}/lists/${userList.id}/items`, {
    params: {
      "$expand": "fields($select=Title,EMail,IsSiteAdmin,UserExpiration)",
      "$top": 500,
      "$filter": "fields/ContentType eq 'Person'",
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