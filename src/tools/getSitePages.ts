import { getGraphClient } from "../auth/graphClient.js";

const SITE_ID = process.env.SITE_ID || "";

export const getSitePagesToolSchema = {
  name: "get_site_pages",
  description:
    "Returns all pages and news posts published on this SharePoint site. " +
    "Shows title, type (Page or News), author, and a direct link. " +
    "Use when the user asks about site pages, news posts, or what's been published.",
  inputSchema: {
    type: "object",
    properties: {
      type: {
        type: "string",
        enum: ["all", "news", "pages"],
        description: "Filter: 'news' for news posts only, 'pages' for site pages only, 'all' for both. Defaults to 'all'.",
      },
    },
    required: [],
  },
};

export async function getSitePages(type: "all" | "news" | "pages" = "all") {
  const client = await getGraphClient();

  // Minimal $select — only fields that exist on baseSitePage
  const res = await client.get(`/sites/${SITE_ID}/pages`, {
    params: {
      $select: "id,title,webUrl,createdDateTime,lastModifiedDateTime,promotionKind",
      $top: 50,
      $orderby: "lastModifiedDateTime desc",
    },
  });

  let pages = (res.data.value || []).map((p: any) => ({
    id: p.id,
    title: p.title || "Untitled",
    type: p.promotionKind === "newsPost" ? "News" : "Page",
    lastModified: p.lastModifiedDateTime || p.createdDateTime || null,
    webUrl: p.webUrl,
  }));

  if (type === "news") pages = pages.filter((p: any) => p.type === "News");
  if (type === "pages") pages = pages.filter((p: any) => p.type === "Page");

  return { total: pages.length, filter: type, pages };
}