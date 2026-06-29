import { getGraphClient } from "../auth/graphClient.js";

const SITE_ID = process.env.SITE_ID || "";

export const getSiteUsersToolSchema = {
  name: "get_site_users",
  description:
    "Returns all users and members who have access to this SharePoint site. " +
    "Shows display name and email. Use when the user asks who has access, " +
    "list site members, or wants to find someone's name or email on the site.",
  inputSchema: { type: "object", properties: {}, required: [] },
};

export async function getSiteUsers() {
  const client = await getGraphClient();

  // Use site groups — works without User Information List
  const groupsRes = await client.get(`/sites/${SITE_ID}/groups`);
  const groups = groupsRes.data.value || [];

  const userSet = new Map<string, { name: string; email: string | null }>();

  for (const group of groups) {
    try {
      const membersRes = await client.get(`/groups/${group.id}/members`, {
        params: { "$select": "displayName,mail,userPrincipalName", "$top": 100 },
      });
      for (const m of membersRes.data.value || []) {
        if (m.displayName) {
          userSet.set(m.id || m.displayName, {
            name: m.displayName,
            email: m.mail || m.userPrincipalName || null,
          });
        }
      }
    } catch {
      // skip groups we can't read
    }
  }

  const users = Array.from(userSet.values());
  return { totalUsers: users.length, users };
}