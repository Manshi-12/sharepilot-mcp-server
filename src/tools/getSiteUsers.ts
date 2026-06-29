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

  // Correct Graph endpoint for site members — no hidden list needed
  const res = await client.get(`/sites/${SITE_ID}/permissions`);
  const permissions = res.data.value || [];

  const userMap = new Map<string, { name: string; email: string | null; role: string }>();

  for (const perm of permissions) {
    const roles: string[] = perm.roles || [];
    const role = roles.includes("owner") ? "Owner" : roles.includes("write") ? "Member" : "Visitor";

    const identities = [
      ...(perm.grantedToV2?.user ? [perm.grantedToV2.user] : []),
      ...((perm.grantedToIdentitiesV2 || []).map((i: any) => i.user).filter(Boolean)),
    ];

    for (const user of identities) {
      if (user?.displayName) {
        userMap.set(user.displayName, {
          name: user.displayName,
          email: user.email || null,
          role,
        });
      }
    }
  }

  const users = Array.from(userMap.values());
  return { totalUsers: users.length, users };
}