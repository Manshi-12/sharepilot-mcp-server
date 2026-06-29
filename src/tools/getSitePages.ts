import { getGraphClient } from "../auth/graphClient.js";

const SITE_ID = process.env.SITE_ID || "";

export const getSitePagesToolSchema = {
  name: "get_site_pages",
  description:
    "Returns all pages and news posts published on this SharePoint site. " +
    "Shows title, type (Page or News), author, published date, and a direct link. " +
    "Use when the user asks about site pages, news posts, announcements, or what's been published.",
  inputSchema: {
    type: "object",
    properties: {
      type: {
        type: "string",
        enum: ["all", "news", "pages"],
        description: "Filter by type: 'news' for news posts only, 'pages' for site pages only, 'all' for both. Defaults to 'all'.",
      },
    },
    required: [],
  },
};

export async function getSitePages(type: "all" | "news" | "pages" = "all") {
  const client = await getGraphClient();

  const res = await client.get(
    `/sites/${SITE_ID}/pages`,
    {
      params: {
        $select: "id,title,webUrl,publishedDateTime,createdDateTime,promotionKind,createdBy",
        $top: 50,
        $orderby: "lastModifiedDateTime desc",
      },
    }
  );

  let pages = (res.data.value || []).map((p: any) => ({
    id: p.id,
    title: p.title || "Untitled",
    type: p.promotionKind === "newsPost" ? "News" : "Page",
    author: p.createdBy?.user?.displayName || "Unknown",
    publishedAt: p.publishedDateTime || p.createdDateTime || null,
    webUrl: p.webUrl,
  }));

  if (type === "news") pages = pages.filter((p: any) => p.type === "News");
  if (type === "pages") pages = pages.filter((p: any) => p.type === "Page");

  return { total: pages.length, filter: type, pages };
}