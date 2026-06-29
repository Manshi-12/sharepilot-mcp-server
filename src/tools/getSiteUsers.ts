import { getGraphClient } from "../auth/graphClient.js";

const SITE_ID = process.env.SITE_ID || "";

export const getSiteUsersToolSchema = {
  name: "get_site_users",
  description:
    "Returns all users and members who have access to this SharePoint site. " +
    "Shows display name and email. Use when the user asks who has access, " +
    "list site members, or wants to find someone's name or email on the site.",
  inputSchema: {
    type: "object",
    properties: {},
    required: [],
  },
};

export async function getSiteUsers() {
  const client = await getGraphClient();

  // Find user list by internal name "users" — works across all tenant languages
  const listsRes = await client.get(`/sites/${SITE_ID}/lists`, {
    params: { "$select": "id,displayName,name", "$top": 100 },
  });

  const userList = (listsRes.data.value || []).find(
    (l: any) =>
      l.name?.toLowerCase() === "users" ||
      l.displayName?.toLowerCase() === "user information list" ||
      l.displayName?.toLowerCase().includes("user information")
  );

  if (userList) {
    const itemsRes = await client.get(`/sites/${SITE_ID}/lists/${userList.id}/items`, {
      params: { "$expand": "fields($select=Title,EMail,IsSiteAdmin)", "$top": 500 },
    });

    const users = (itemsRes.data.value || [])
      .filter((u: any) => u.fields?.Title)
      .map((u: any) => ({
        name: u.fields.Title,
        email: u.fields.EMail || null,
        isAdmin: u.fields.IsSiteAdmin || false,
      }));

    return { totalUsers: users.length, users };
  }

  // Fallback: site permissions API
  const permRes = await client.get(`/sites/${SITE_ID}/permissions`);
  const users: any[] = [];
  for (const perm of permRes.data.value || []) {
    const identities = perm.grantedToIdentitiesV2 || perm.grantedToIdentities || [];
    for (const identity of identities) {
      if (identity.user) {
        users.push({
          name: identity.user.displayName || "Unknown",
          email: identity.user.email || null,
        });
      }
    }
  }

  return { totalUsers: users.length, users };
}