import { AxiosInstance } from "axios";

const SITE_ID = process.env.SITE_ID || "";
const CACHE_TTL_MS = 5 * 60 * 1000; // 5 minutes — long enough to avoid hammering Graph,
                                     // short enough that a newly-created library/list shows up fast.

interface CacheEntry<T> {
  data: T;
  expiresAt: number;
}

export interface DriveInfo {
  id: string;
  name: string;
}

export interface ListInfo {
  id: string;
  name: string;        // internal name
  displayName: string; // friendly name users type
}

export interface ColumnInfo {
  internalName: string;
  displayName: string;
  type: "choice" | "multiChoice" | "boolean" | "number" | "currency" | "dateTime"
      | "lookup" | "personOrGroup" | "hyperlinkOrPicture" | "text";
  required: boolean;
  choices?: string[];
}

let driveCache: CacheEntry<DriveInfo[]> | null = null;
let listCache: CacheEntry<ListInfo[]> | null = null;
const columnCache = new Map<string, CacheEntry<ColumnInfo[]>>();

/** Strips spacing/case/punctuation so "Task Category", "TaskCategory" and "task_category" all match. */
export function normalizeKey(s: string): string {
  return s.trim().toLowerCase().replace(/[\s_-]+/g, "");
}

export function clearResolverCache(): void {
  driveCache = null;
  listCache = null;
  columnCache.clear();
}

// ---------- Drives (Document Libraries) ----------

export async function getAllDrives(client: AxiosInstance): Promise<DriveInfo[]> {
  if (driveCache && driveCache.expiresAt > Date.now()) return driveCache.data;

  const res = await client.get(`/sites/${SITE_ID}/drives?$select=id,name`);
  const drives: DriveInfo[] = (res.data.value || []).map((d: any) => ({ id: d.id, name: d.name }));

  driveCache = { data: drives, expiresAt: Date.now() + CACHE_TTL_MS };
  return drives;
}

/**
 * Finds a Document Library by display name. No hardcoded names anywhere —
 * works for any library that currently exists on the site, including ones
 * created after this server was deployed.
 */
export async function resolveDrive(client: AxiosInstance, libraryName: string): Promise<DriveInfo> {
  const drives = await getAllDrives(client);
  const key = normalizeKey(libraryName);

  const exact = drives.find((d) => normalizeKey(d.name) === key);
  if (exact) return exact;

  const partial = drives.find((d) => normalizeKey(d.name).includes(key));
  if (partial) return partial;

  throw new Error(
    `Document library "${libraryName}" was not found on this site. ` +
    `Available libraries: ${drives.map((d) => d.name).join(", ") || "(none found)"}`
  );
}

// ---------- Lists ----------

export async function getAllLists(client: AxiosInstance): Promise<ListInfo[]> {
  if (listCache && listCache.expiresAt > Date.now()) return listCache.data;

  const res = await client.get(`/sites/${SITE_ID}/lists?$select=id,name,displayName,list`);
  const lists: ListInfo[] = (res.data.value || [])
    // Hide system/hidden lists so they never collide with user-named lists.
    .filter((l: any) => !l.list?.hidden)
    .map((l: any) => ({ id: l.id, name: l.name, displayName: l.displayName || l.name }));

  listCache = { data: lists, expiresAt: Date.now() + CACHE_TTL_MS };
  return lists;
}

/**
 * Finds a List by display name. Works for any list on the site — Project Tasks
 * today, anything else tomorrow — with zero code changes required.
 */
export async function resolveList(client: AxiosInstance, listName: string): Promise<ListInfo> {
  const lists = await getAllLists(client);
  const key = normalizeKey(listName);

  const exact = lists.find((l) => normalizeKey(l.displayName) === key || normalizeKey(l.name) === key);
  if (exact) return exact;

  const partial = lists.find(
    (l) => normalizeKey(l.displayName).includes(key) || normalizeKey(l.name).includes(key)
  );
  if (partial) return partial;

  throw new Error(
    `List "${listName}" was not found on this site. ` +
    `Available lists: ${lists.map((l) => l.displayName).join(", ") || "(none found)"}`
  );
}

// ---------- List Columns (schema) ----------

/**
 * Reads a list's actual column definitions from Graph and classifies each one
 * by real type (choice / multiChoice / boolean / number / dateTime / lookup /
 * personOrGroup / hyperlinkOrPicture / text). This replaces every hardcoded
 * field_N map — the server now learns the schema of ANY list at runtime.
 */
export async function getListColumns(client: AxiosInstance, listId: string): Promise<ColumnInfo[]> {
  const cached = columnCache.get(listId);
  if (cached && cached.expiresAt > Date.now()) return cached.data;

  const res = await client.get(`/sites/${SITE_ID}/lists/${listId}/columns`);
  const raw = res.data.value || [];

  const columns: ColumnInfo[] = raw
    // Skip Graph's own bookkeeping columns (ContentType, Attachments, etc.)
    .filter((c: any) => !c.readOnly && !c.hidden)
    .map((c: any) => {
      let type: ColumnInfo["type"] = "text";
      let choices: string[] | undefined;

      if (c.choice) {
        type = c.choice.allowMultipleValues ? "multiChoice" : "choice";
        choices = c.choice.choices;
      } else if (c.boolean) type = "boolean";
      else if (c.number) type = "number";
      else if (c.currency) type = "currency";
      else if (c.dateTime) type = "dateTime";
      else if (c.lookup) type = "lookup";
      else if (c.personOrGroup) type = "personOrGroup";
      else if (c.hyperlinkOrPicture) type = "hyperlinkOrPicture";

      return {
        internalName: c.name,
        displayName: c.displayName || c.name,
        type,
        required: !!c.required,
        choices,
      };
    });

  columnCache.set(listId, { data: columns, expiresAt: Date.now() + CACHE_TTL_MS });
  return columns;
}

/** Matches a user/agent-supplied field key (e.g. "Task Category") to its column definition. */
export function findColumn(columns: ColumnInfo[], key: string): ColumnInfo | undefined {
  const k = normalizeKey(key);
  return columns.find((c) => normalizeKey(c.displayName) === k || normalizeKey(c.internalName) === k);
}
