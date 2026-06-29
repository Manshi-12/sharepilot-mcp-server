import { getGraphClient } from "../auth/graphClient.js";

const SITE_ID = process.env.SITE_ID || "";

export const searchSiteContentToolSchema = {
  name: "search_site_content",
  description:
    "Performs a full-text search across ALL content on this SharePoint site — " +
    "files, list items, and pages — using a keyword or phrase. " +
    "Use when the user wants to find anything on the site without knowing where it lives, " +
    "or asks a broad question like 'find anything about X' or 'search the whole site for Y'.",
  inputSchema: {
    type: "object",
    properties: {
      query: {
        type: "string",
        description: "The keyword or phrase to search for across the entire site.",
      },
      top: {
        type: "number",
        description: "Maximum number of results to return. Defaults to 10.",
      },
    },
    required: ["query"],
  },
};

export async function searchSiteContent(query: string, top: number = 10) {
  const client = await getGraphClient();

  const res = await client.post(`/search/query`, {
    requests: [
      {
        entityTypes: ["driveItem", "listItem", "site"],
        query: { queryString: query },
        from: 0,
        size: top,
        fields: ["title", "name", "webUrl", "lastModifiedDateTime", "createdBy", "contentType"],
      },
    ],
  });

  const hits = res.data?.value?.[0]?.hitsContainers?.[0]?.hits || [];

  const results = hits.map((hit: any) => {
    const r = hit.resource || {};
    return {
      title: r.name || r.displayName || r.subject || "Untitled",
      type: hit.resource["@odata.type"]?.replace("#microsoft.graph.", "") || "unknown",
      webUrl: r.webUrl || null,
      lastModified: r.lastModifiedDateTime || null,
      summary: hit.summary || null,
    };
  });

  return { query, totalResults: results.length, results };
}