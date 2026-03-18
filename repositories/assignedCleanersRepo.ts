import * as sharepoint from "../lib/sharepoint";
import type { SiteCleanerAssignment } from "../types";

const SITE_CLEANERS_LIST_NAME = "CleanTrack Site Cleaners";

/** Display name -> internal name. Cached in-memory. */
let cachedAssignmentsFieldMap: Record<string, string> | null = null;

function getField<T>(fields: Record<string, unknown>, ...keys: string[]): T | undefined {
  const f = fields as Record<string, T>;
  for (const k of keys) {
    if (f[k] !== undefined && f[k] !== null) return f[k];
  }
  return undefined;
}

function getString(fields: Record<string, unknown>, ...keys: string[]): string {
  const v = getField<unknown>(fields, ...keys);
  if (v == null) return "";
  return String(v).trim();
}

function getBoolean(fields: Record<string, unknown>, key: string | undefined, defaultVal: boolean): boolean {
  if (!key) return defaultVal;
  const v = (fields as Record<string, unknown>)[key];
  if (v === undefined || v === null) return defaultVal;
  if (v === false || v === "No" || String(v).toLowerCase() === "no") return false;
  if (v === true || v === "Yes" || String(v).toLowerCase() === "yes") return true;
  return defaultVal;
}

function requireMapKey(map: Record<string, string>, displayName: string): string {
  const internal = map[displayName];
  if (!internal) {
    throw new Error(`SharePoint column not found: ${displayName}. Check list column display name.`);
  }
  return internal;
}

function resolveLookupDisplayName(raw: unknown): string {
  if (raw == null) return "";
  if (typeof raw === "string") return raw.trim();
  if (typeof raw === "number") return String(raw);
  if (typeof raw !== "object") return "";
  const obj = raw as Record<string, unknown>;
  if (typeof obj.LookupValue === "string" && obj.LookupValue.trim()) {
    return obj.LookupValue.trim();
  }
  if (typeof obj.Title === "string" && obj.Title.trim()) {
    return obj.Title.trim();
  }
  if (typeof obj["Site Name"] === "string" && String(obj["Site Name"]).trim()) {
    return String(obj["Site Name"]).trim();
  }
  if (typeof obj.Site_x0020_Name === "string" && String(obj.Site_x0020_Name).trim()) {
    return String(obj.Site_x0020_Name).trim();
  }
  if (typeof obj.siteName === "string" && obj.siteName.trim()) {
    return obj.siteName.trim();
  }
  if (typeof obj.name === "string" && obj.name.trim()) {
    return obj.name.trim();
  }
  return "";
}

function getLookupIdAndName(
  fields: Record<string, unknown>,
  baseInternalName: string
): { id: string; name: string; raw: unknown } {
  const keys = Object.keys(fields);
  const idKey =
    keys.find((k) => k === `${baseInternalName}LookupId` || k === `${baseInternalName}Id`) ??
    keys.find((k) => k.toLowerCase() === `${baseInternalName.toLowerCase()}lookupid`);
  const nameKey =
    keys.find((k) => k === baseInternalName) ??
    keys.find((k) => k.toLowerCase() === baseInternalName.toLowerCase());

  let rawId: unknown = idKey ? fields[idKey] : undefined;
  let rawLookup: unknown = nameKey ? fields[nameKey] : undefined;

  if (rawId == null && rawLookup && typeof rawLookup === "object" && "LookupId" in (rawLookup as object)) {
    rawId = (rawLookup as { LookupId?: number | string }).LookupId;
  }

  const id = rawId != null ? sharepoint.normalizeListItemId(String(rawId).trim()) : "";
  const name = resolveLookupDisplayName(rawLookup);

  return { id, name, raw: rawLookup };
}

let itemToAssignmentDebugCounter = 0;

async function getSiteAndListId(accessToken: string): Promise<{ siteId: string; listId: string }> {
  const siteId = await sharepoint.getSiteId(accessToken);
  const listId = await sharepoint.getListIdByName(accessToken, siteId, SITE_CLEANERS_LIST_NAME);
  if (!listId) throw new Error(`List "${SITE_CLEANERS_LIST_NAME}" not found.`);
  return { siteId, listId };
}

/** Get displayName -> internal name map for CleanTrack Site Cleaners. Cached in-memory. */
export async function getSiteCleanerAssignmentsFieldMap(accessToken: string): Promise<Record<string, string>> {
  if (cachedAssignmentsFieldMap) return cachedAssignmentsFieldMap;
  const { siteId, listId } = await getSiteAndListId(accessToken);
  const columns = await sharepoint.getListColumns(accessToken, siteId, listId);
  const map: Record<string, string> = {};
  for (const col of columns) {
    if (col.displayName) map[col.displayName] = col.name;
  }
  cachedAssignmentsFieldMap = map;
  return map;
}

function itemToAssignment(
  item: sharepoint.GraphListItem,
  map: Record<string, string>
): SiteCleanerAssignment {
  const fields = item.fields ?? {};
  const assignmentNameKey = map["Assignment Name"] ?? "Assignment Name";
  const rawAssignmentName =
    getString(fields, assignmentNameKey, "Title", "LinkTitle", "Assignment_x0020_Name") ||
    getString(fields, "Title") ||
    "";

  const siteInternal = map["Site"] ?? "Site";
  const cleanerInternal = map["Cleaner"] ?? "Cleaner";

  const { id: siteId, name: siteName, raw: rawSiteLookup } = getLookupIdAndName(
    fields as Record<string, unknown>,
    siteInternal
  );
  const { id: cleanerId, name: cleanerName } = getLookupIdAndName(
    fields as Record<string, unknown>,
    cleanerInternal
  );

  const activeKey = map["Active"] ?? "Active";

  // Targeted debug: first few records for site lookup resolution
  if (itemToAssignmentDebugCounter < 3) {
    console.log("[SiteCleaners] itemToAssignment debug:", {
      itemId: item.id,
      siteInternal,
      cleanerInternal,
      rawFieldsSite: (fields as Record<string, unknown>)[siteInternal],
      resolvedSiteId: siteId,
      resolvedSiteName: siteName,
      rawSiteLookup,
    });
    itemToAssignmentDebugCounter++;
  }

  return {
    id: sharepoint.normalizeListItemId(item.id),
    assignmentName: rawAssignmentName,
    siteId,
    siteName,
    cleanerId,
    cleanerName,
    active: getBoolean(fields, activeKey, true),
  };
}

/** Fetch Site Cleaner assignments from CleanTrack Site Cleaners. */
export async function getSiteCleanerAssignments(
  accessToken: string,
  options?: { activeOnly?: boolean; siteId?: string }
): Promise<SiteCleanerAssignment[]> {
  const { siteId, listId } = await getSiteAndListId(accessToken);
  const map = await getSiteCleanerAssignmentsFieldMap(accessToken);
  const items = await sharepoint.getListItems(accessToken, siteId, listId);

  // Debug: raw Site Cleaner assignment records
  if (items.length > 0) {
    console.log("[SiteCleaners] raw items sample:", {
      count: items.length,
      first: items[0],
    });
  } else {
    console.log("[SiteCleaners] no items returned from list");
  }

  let assignments = items.map((item) => itemToAssignment(item, map));

  // Debug: mapped assignment objects
  if (assignments.length > 0) {
    console.log("[SiteCleaners] mapped assignments sample:", {
      count: assignments.length,
      first: assignments[0],
    });
  }

  if (options?.activeOnly) {
    assignments = assignments.filter((a) => a.active);
  }
  if (options?.siteId) {
    const target = sharepoint.normalizeListItemId(options.siteId);
    assignments = assignments.filter(
      (a) => target && sharepoint.normalizeListItemId(String(a.siteId)) === target
    );
  }
  return assignments;
}

export interface SiteCleanerAssignmentPayload {
  siteId: string;
  cleanerId: string;
  assignmentName?: string;
  active?: boolean;
}

/** Build fields object using internal column names for CleanTrack Site Cleaners. */
function payloadToFields(
  payload: SiteCleanerAssignmentPayload,
  map: Record<string, string>
): Record<string, unknown> {
  const fields: Record<string, unknown> = {};

  if (payload.assignmentName !== undefined) {
    let k = requireMapKey(map, "Assignment Name");
    if (k === "LinkTitle") k = "Title";
    fields[k] = payload.assignmentName;
  }

  const siteInternal = map["Site"] ?? "Site";
  const cleanerInternal = map["Cleaner"] ?? "Cleaner";

  const siteKey = siteInternal === "Site" ? "SiteLookupId" : `${siteInternal}LookupId`;
  const cleanerKey = cleanerInternal === "Cleaner" ? "CleanerLookupId" : `${cleanerInternal}LookupId`;

  if (payload.siteId !== undefined) {
    const val = /^\d+$/.test(payload.siteId) ? parseInt(payload.siteId, 10) : payload.siteId;
    fields[siteKey] = val;
  }
  if (payload.cleanerId !== undefined) {
    const val = /^\d+$/.test(payload.cleanerId) ? parseInt(payload.cleanerId, 10) : payload.cleanerId;
    fields[cleanerKey] = val;
  }

  if (payload.active !== undefined) {
    const k = map["Active"];
    if (k) fields[k] = payload.active;
  }

  return fields;
}

/** Create a new Site Cleaner assignment (list item). */
export async function createSiteCleanerAssignment(
  accessToken: string,
  payload: SiteCleanerAssignmentPayload
): Promise<SiteCleanerAssignment> {
  const { siteId, listId } = await getSiteAndListId(accessToken);
  const map = await getSiteCleanerAssignmentsFieldMap(accessToken);

  const fields = payloadToFields(
    {
      ...payload,
      active: payload.active !== false,
    },
    map
  );

  const created = await sharepoint.createListItem(accessToken, siteId, listId, fields);
  return itemToAssignment({ id: created.id, fields: created.fields ?? {} }, map);
}

/** Update an existing Site Cleaner assignment (e.g. Active flag). */
export async function updateSiteCleanerAssignment(
  accessToken: string,
  id: string,
  payload: Partial<Pick<SiteCleanerAssignmentPayload, "active">>
): Promise<void> {
  const { siteId, listId } = await getSiteAndListId(accessToken);
  const map = await getSiteCleanerAssignmentsFieldMap(accessToken);
  const fields = payloadToFields(
    {
      siteId: "" as string,
      cleanerId: "" as string,
      assignmentName: undefined,
      ...(payload.active !== undefined ? { active: payload.active } : {}),
    },
    map
  );
  const filteredFields: Record<string, unknown> = {};
  for (const [k, v] of Object.entries(fields)) {
    if (v !== "" && v !== undefined) filteredFields[k] = v;
  }
  if (Object.keys(filteredFields).length === 0) return;
  await sharepoint.updateListItem(accessToken, siteId, listId, id, filteredFields);
}

/** Convenience: map of siteId -> active cleanerIds. */
export async function getActiveCleanerIdsBySite(
  accessToken: string
): Promise<Record<string, string[]>> {
  const assignments = await getSiteCleanerAssignments(accessToken, { activeOnly: true });
  const result: Record<string, string[]> = {};
  for (const a of assignments) {
    if (!a.siteId || !a.cleanerId || !a.active) continue;
    if (!result[a.siteId]) result[a.siteId] = [];
    if (!result[a.siteId].includes(a.cleanerId)) {
      result[a.siteId].push(a.cleanerId);
    }
  }
  return result;
}
