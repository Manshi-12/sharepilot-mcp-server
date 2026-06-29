import { getGraphClient } from "../auth/graphClient.js";

export const searchSiteContentToolSchema = {
  name: "search_site_content",
  description:
    "Searches across ALL content on this SharePoint site — files and list items — using a keyword or phrase. " +
    "Use when the user wants to find anything on the site without knowing exactly where it lives.",
  inputSchema: {
    type: "object",
    properties: {
      query: { type: "string", description: "Keyword or phrase to search for." },
      top: { type: "number", description: "Max results to return. Defaults to 10." },
    },
    required: ["query"],
  },
};

export async function searchSiteContent(query: string, top: number = 10) {
  const client = await getGraphClient();

  // No contentSources for driveItem/listItem — region required for app-only
  const res = await client.post(`/search/query`, {
    requests: [
      {
        entityTypes: ["driveItem", "listItem"],
        query: { queryString: query },
        from: 0,
        size: top,
        region: "NAM",
      },
    ],
  });

  const hits = res.data?.value?.[0]?.hitsContainers?.[0]?.hits || [];
  const total = res.data?.value?.[0]?.hitsContainers?.[0]?.total || 0;

  const results = hits.map((hit: any) => {
    const r = hit.resource || {};
    return {
      title: r.name || r.displayName || "Untitled",
      type: (r["@odata.type"] || "").replace("#microsoft.graph.", ""),
      webUrl: r.webUrl || null,
      lastModified: r.lastModifiedDateTime || null,
      summary: hit.summary || null,
    };
  });

  return { query, totalResults: total, returned: results.length, results };
}