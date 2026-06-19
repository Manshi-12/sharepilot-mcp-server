import { searchFile, searchFileToolSchema } from "./searchFile.js";
import { readFile, readFileToolSchema } from "./readFile.js";
import { createListItem, createListItemToolSchema } from "./createListItem.js";
import { uploadFile, uploadFileToolSchema } from "./uploadFile.js";
import { getListItems, getListItemsToolSchema } from "./getListItems.js";
import { uploadListItemImage, uploadListItemImageToolSchema } from "./uploadListItemImage.js";
import { createList, createListToolSchema } from "./createList.js";
import { updateListItem, updateListItemToolSchema } from "./updateListItem.js";
import { deleteListItem, deleteListItemToolSchema } from "./deleteListItem.js";
import { deleteFile, deleteFileToolSchema } from "./deleteFile.js";

// Single source of truth for every tool's schema — used by both the raw
// MCP transport (/mcp) and the OpenAI-style function-calling agent (/chat).
export const TOOL_SCHEMAS = [
  searchFileToolSchema,
  readFileToolSchema,
  createListItemToolSchema,
  uploadFileToolSchema,
  getListItemsToolSchema,
  uploadListItemImageToolSchema,
  createListToolSchema,
  updateListItemToolSchema,
  deleteListItemToolSchema,
  deleteFileToolSchema,
];

/**
 * Executes a tool by name with the given arguments and returns a
 * JSON-serializable result. Throws on unknown tool name or tool failure —
 * callers are responsible for catching and formatting errors for their
 * respective protocol (MCP content block vs. OpenAI tool message).
 */
export async function executeTool(name: string, args: any): Promise<any> {
  switch (name) {
    case "search_file":
      return searchFile(args.filename, args.libraryName);
    case "read_file":
      return readFile(args.fileId, args.driveId);
    case "create_list_item":
      return createListItem(args.listName, args.fields);
    case "upload_file":
      return uploadFile(args.filename, args.content, args.libraryName, args.isBase64, args.mimeType);
    case "get_list_items":
      return getListItems(args.listName, args.search, args.top);
    case "upload_list_item_image":
      return uploadListItemImage(
        args.listName,
        args.itemId,
        args.imageFieldName,
        args.fileName,
        args.base64Content,
        args.mimeType
      );
    case "create_list":
      return createList(args.displayName, args.template, args.description, args.columns);
    case "update_list_item":
      return updateListItem(args.listName, args.itemId, args.fields);
    case "delete_list_item":
      return deleteListItem(args.listName, args.itemId);
    case "delete_file":
      return deleteFile(args.libraryName, args.fileId);
    default:
      throw new Error(`Unknown tool: ${name}`);
  }
}
