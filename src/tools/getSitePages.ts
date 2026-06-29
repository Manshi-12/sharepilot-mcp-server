import { getGraphClient } from "../auth/graphClient.js";

const SITE_ID = process.env.SITE_ID || "";

export const getSitePagesToolSchema = {
  name: "get_site_pages",
  description:
    "Returns all pages and news posts on this SharePoint site. " +
    "Use when the user asks about site pages, news posts, or what's published.",
  inputSchema: {
    type: "object",
    properties: {
      type: {
        type: "string",
        enum: ["all", "news", "pages"],
        description: "Filter: 'news', 'pages', or 'all'. Defaults to 'all'.",
      },
    },
    required: [],
  },
};

export async function getSitePages(type: "all" | "news" | "pages" = "all") {
  const client = await getGraphClient();

  // Use no $select at all — let Graph return default fields to avoid property errors
  const res = await client.get(`/sites/${SITE_ID}/pages`, {
    params: { $top: 50, $orderby: "lastModifiedDateTime desc" },
  });

  let pages = (res.data.value || []).map((p: any) => ({
    id: p.id,
    title: p.title || p.name || "Untitled",
    // promotionKind may not exist on older API versions — use odata type instead
    type: (p["@odata.type"] || "").includes("NewsLinkPage") || p.promotionKind === "newsPost"
      ? "News"
      : "Page",
    lastModified: p.lastModifiedDateTime || p.createdDateTime || null,
    webUrl: p.webUrl,
  }));

  if (type === "news") pages = pages.filter((p: any) => p.type === "News");
  if (type === "pages") pages = pages.filter((p: any) => p.type === "Page");

  return { total: pages.length, filter: type, pages };
}