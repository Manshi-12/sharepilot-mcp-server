import { AxiosInstance } from "axios";

const SITE_ID = process.env.SITE_ID || "";
const CACHE_TTL_MS = 5 * 60 * 1000;

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
  name: string;
  displayName: string;
}

export interface ColumnInfo {
  internalName: string;
  displayName: string;
  type: "choice" | "multiChoice" | "boolean" | "number" | "currency" | "dateTime"
  | "lookup" | "personOrGroup" | "hyperlinkOrPicture" | "text";
  required: boolean;
  choices?: string[];
  multi?: boolean;
}

let driveCache: CacheEntry<DriveInfo[]> | null = null;
let listCache: CacheEntry<ListInfo[]> | null = null;
const columnCache = new Map<string, CacheEntry<ColumnInfo[]>>();

// Fix #9 — cache the user map so we don't fetch 2000 users on every person field
let userMapCache: CacheEntry<Map<number, string>> | null = null;

export function normalizeKey(s: string): string {
  return s.trim().toLowerCase().replace(/[\s_-]+/g, "");
}

export function clearResolverCache(): void {
  driveCache = null;
  listCache = null;
  columnCache.clear();
  userMapCache = null;
}

// ---------- Drives ----------

export async function getAllDrives(client: AxiosInstance): Promise<DriveInfo[]> {
  if (driveCache && driveCache.expiresAt > Date.now()) return driveCache.data;

  const res = await client.get(`/sites/${SITE_ID}/drives?$select=id,name`);
  const drives: DriveInfo[] = (res.data.value || []).map((d: any) => ({ id: d.id, name: d.name }));

  driveCache = { data: drives, expiresAt: Date.now() + CACHE_TTL_MS };
  return drives;
}

export async function resolveDrive(client: AxiosInstance, libraryName: string): Promise<DriveInfo> {
  const drives = await getAllDrives(client);
  const key = normalizeKey(libraryName);

  const exact = drives.find((d) => normalizeKey(d.name) === key);
  if (exact) return exact;

  const partial = drives.find((d) => normalizeKey(d.name).includes(key));
  if (partial) return partial;

  throw new Error(
    `Document library "${libraryName}" was not found on this site. Please check the name and try again.`
  );
}

// ---------- Lists ----------

export async function getAllLists(client: AxiosInstance): Promise<ListInfo[]> {
  if (listCache && listCache.expiresAt > Date.now()) return listCache.data;

  const res = await client.get(`/sites/${SITE_ID}/lists?$select=id,name,displayName,list`);
  const lists: ListInfo[] = (res.data.value || [])
    .filter((l: any) => !l.list?.hidden)
    .map((l: any) => ({ id: l.id, name: l.name, displayName: l.displayName || l.name }));

  listCache = { data: lists, expiresAt: Date.now() + CACHE_TTL_MS };
  return lists;
}

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
    `List "${listName}" was not found on this site. Please check the name and try again.`
  );
}

// ---------- List Columns ----------

export async function getListColumns(client: AxiosInstance, listId: string): Promise<ColumnInfo[]> {
  const cached = columnCache.get(listId);
  if (cached && cached.expiresAt > Date.now()) return cached.data;

  const res = await client.get(`/sites/${SITE_ID}/lists/${listId}/columns`);
  const raw = res.data.value || [];

  const columns: ColumnInfo[] = raw
    .filter((c: any) => !c.readOnly && !c.hidden)
    .map((c: any) => {
      let type: ColumnInfo["type"] = "text";
      let choices: string[] | undefined;
      let multi: boolean | undefined;

      if (c.choice) {
        type = c.choice.allowMultipleValues ? "multiChoice" : "choice";
        choices = c.choice.choices;
      } else if (c.boolean) type = "boolean";
      else if (c.number) type = "number";
      else if (c.currency) type = "currency";
      else if (c.dateTime) type = "dateTime";
      else if (c.lookup) type = "lookup";
      else if (c.personOrGroup) {
        type = "personOrGroup";
        multi = !!c.personOrGroup.allowMultipleSelections;
      } else if (c.hyperlinkOrPicture) type = "hyperlinkOrPicture";

      return {
        internalName: c.name,
        displayName: c.displayName || c.name,
        type,
        required: !!c.required,
        choices,
        multi,
      };
    });

  columnCache.set(listId, { data: columns, expiresAt: Date.now() + CACHE_TTL_MS });
  return columns;
}

export function findColumn(columns: ColumnInfo[], key: string): ColumnInfo | undefined {
  const k = normalizeKey(key);
  return columns.find((c) => normalizeKey(c.displayName) === k || normalizeKey(c.internalName) === k);
}

// ---------- Person field resolution ----------

async function getUserInfoListId(client: AxiosInstance): Promise<string> {
  const res = await client.get(`/sites/${SITE_ID}/lists?$select=id,displayName`);
  const match = (res.data.value || []).find(
    (l: any) => (l.displayName || "").toLowerCase() === "user information list"
  );
  if (!match) throw new Error("Could not locate this site's User Information List.");
  return match.id;
}

// Fix #9 — build user map once and cache it for 5 minutes
export async function getUserMap(client: AxiosInstance): Promise<Map<number, string>> {
  if (userMapCache && userMapCache.expiresAt > Date.now()) return userMapCache.data;

  const listId = await getUserInfoListId(client);
  const res = await client.get(`/sites/${SITE_ID}/lists/${listId}/items`, {
    params: { "$expand": "fields($select=Title,EMail)", "$top": 2000 },
  });

  const map = new Map<number, string>();
  for (const u of res.data.value || []) {
    const numId = Number(u.id);
    const name = u.fields?.Title || u.fields?.EMail || `User ${numId}`;
    map.set(numId, name);
  }

  userMapCache = { data: map, expiresAt: Date.now() + CACHE_TTL_MS };
  return map;
}

export async function resolvePersonId(client: AxiosInstance, nameOrEmail: string): Promise<number | null> {
  try {
    const map = await getUserMap(client);
    const key = nameOrEmail.trim().toLowerCase();

    for (const [id, name] of map.entries()) {
      if (name.toLowerCase() === key) return id;
    }
    // partial match fallback
    for (const [id, name] of map.entries()) {
      if (name.toLowerCase().includes(key)) return id;
    }
    return null;
  } catch {
    return null;
  }
}

// ---------- Image / Thumbnail field parsing ----------

export interface ParsedImage {
  url: string;
  fileName: string;
}

export function parseImageFieldValue(raw: any): ParsedImage | null {
  if (!raw) return null;

  let value = raw;
  if (typeof raw === "string") {
    try {
      value = JSON.parse(raw);
    } catch {
      if (raw.startsWith("http")) return { url: raw, fileName: "" };
      return null;
    }
  }

  if (value && typeof value === "object") {
    if (value.serverUrl && value.serverRelativeUrl) {
      return { url: value.serverUrl + value.serverRelativeUrl, fileName: value.fileName || "" };
    }
    if (value.Url) {
      return { url: value.Url, fileName: value.Description || "" };
    }
  }

  return null;
}
