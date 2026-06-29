import { getGraphClient } from "../auth/graphClient.js";
import { getUserMap } from "../utils/resolve.js";

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
  const userMap = await getUserMap(client);

  const users = Array.from(userMap.entries()).map(([id, name]) => ({ id, name }));

  return { totalUsers: users.length, users };
}