import { getGraphClient } from "../auth/graphClient.js";

const SITE_ID = process.env.SITE_ID || "";
const SITE_URL = process.env.SITE_URL || "";

export const searchSiteContentToolSchema = {
  name: "search_site_content",
  description:
    "Performs a full-text search across ALL content on this SharePoint site — " +
    "files, list items, and pages — using a keyword or phrase. " +
    "Use when the user wants to find anything on the site without knowing where it lives.",
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

  // Scope search to this specific site using contentSources
  const siteScope = SITE_URL
    ? `${SITE_URL.replace(/\/$/, "")}`
    : null;

  const requestBody: any = {
    requests: [
      {
        entityTypes: ["driveItem", "listItem"],
        query: {
          queryString: siteScope
            ? `${query} site:${siteScope}`
            : query,
        },
        from: 0,
        size: top,
        fields: ["title", "name", "webUrl", "lastModifiedDateTime", "createdBy"],
      },
    ],
  };

  const res = await client.post(`/search/query`, requestBody);

  const hits = res.data?.value?.[0]?.hitsContainers?.[0]?.hits || [];
  const total = res.data?.value?.[0]?.hitsContainers?.[0]?.total || 0;

  const results = hits.map((hit: any) => {
    const r = hit.resource || {};
    return {
      title: r.name || r.displayName || r.subject || "Untitled",
      type: (hit.resource["@odata.type"] || "").replace("#microsoft.graph.", ""),
      webUrl: r.webUrl || null,
      lastModified: r.lastModifiedDateTime || null,
      summary: hit.summary || null,
    };
  });

  return { query, totalResults: total, returned: results.length, results };
}