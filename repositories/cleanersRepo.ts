import * as sharepoint from "../lib/sharepoint";

const CLEANERS_LIST_NAME = "CleanTrack Cleaners";

/** Display name -> internal name. Cached in-memory. */
let cachedCleanersFieldMap: Record<string, string> | null = null;

export interface CleanerItem {
  id: string;
  cleanerName: string;
  payRatePerHour: number;
  accountName: string;
  bsb: string;
  accountNumber: string;
  active: boolean;
  type: "cleaner" | "contractor";
}

/** Graph/SharePoint may return a string or a wrapped choice value. */
function coerceChoiceString(v: unknown): string {
  if (v == null) return "";
  if (typeof v === "object" && v !== null && "value" in v) {
    const inner = (v as { value?: unknown }).value;
    return inner == null ? "" : String(inner).trim();
  }
  return String(v).trim();
}

function normalizeWorkerType(v: unknown): "cleaner" | "contractor" {
  const s = coerceChoiceString(v).toLowerCase();
  return s === "contractor" ? "contractor" : "cleaner";
}

/** SharePoint Choice options use title case; app model uses lowercase. */
function workerTypeToSharePointChoice(t: "cleaner" | "contractor"): string {
  return t === "contractor" ? "Contractor" : "Cleaner";
}

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

function getNumber(fields: Record<string, unknown>, ...keys: string[]): number {
  const v = getField<unknown>(fields, ...keys);
  if (v == null || v === "") return 0;
  if (typeof v === "number" && !Number.isNaN(v)) return v;
  const n = parseFloat(String(v).replace(/[^0-9.-]/g, ""));
  return Number.isNaN(n) ? 0 : n;
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

function itemToCleaner(item: sharepoint.GraphListItem, map: Record<string, string>): CleanerItem {
  const fields = item.fields ?? {};
  const nameKey = map["Cleaner Name"] ?? "Cleaner Name";
  const cleanerName =
    getString(fields, nameKey, "LinkTitle", "Title", "Cleaner_x0020_Name") ||
    getString(fields, "Title") ||
    "";
  const payRateKey = map["Pay Rate"] ?? "Pay Rate";
  const accountKey = map["Account Name"] ?? "Account Name";
  const bsbKey = map["BSB"] ?? "BSB";
  const accountNumKey = map["Account Number"] ?? "Account Number";
  const activeKey = map["Active"] ?? "Active";
  const workerTypeInternal = map["Worker Type"];
  const rawWorkerType = workerTypeInternal
    ? getField<unknown>(fields, workerTypeInternal)
    : getField<unknown>(fields, "Worker_x0020_Type");
  return {
    id: item.id,
    cleanerName,
    payRatePerHour: getNumber(fields, payRateKey, "Pay_x0020_Rate", "Pay Rate"),
    accountName: getString(fields, accountKey, "Account_x0020_Name", "Account Name"),
    bsb: getString(fields, bsbKey, "BSB"),
    accountNumber: getString(fields, accountNumKey, "Account_x0020_Number", "Account Number"),
    active: getBoolean(fields, activeKey, true),
    type: normalizeWorkerType(rawWorkerType),
  };
}

async function getSiteAndListId(accessToken: string): Promise<{ siteId: string; listId: string }> {
  const siteId = await sharepoint.getSiteId(accessToken);
  const listId = await sharepoint.getListIdByName(accessToken, siteId, CLEANERS_LIST_NAME);
  if (!listId) throw new Error(`List "${CLEANERS_LIST_NAME}" not found.`);
  return { siteId, listId };
}

/** Get displayName -> internal name map for CleanTrack Cleaners. Cached in-memory. */
export async function getCleanersFieldMap(accessToken: string): Promise<Record<string, string>> {
  if (cachedCleanersFieldMap) return cachedCleanersFieldMap;
  const { siteId, listId } = await getSiteAndListId(accessToken);
  const columns = await sharepoint.getListColumns(accessToken, siteId, listId);
  const map: Record<string, string> = {};
  for (const col of columns) {
    if (col.displayName) map[col.displayName] = col.name;
  }
  cachedCleanersFieldMap = map;
  return map;
}

/** Fetch all cleaners from CleanTrack Cleaners list. */
export async function getCleaners(accessToken: string): Promise<CleanerItem[]> {
  const { siteId, listId } = await getSiteAndListId(accessToken);
  const map = await getCleanersFieldMap(accessToken);
  const items = await sharepoint.getListItems(accessToken, siteId, listId);
  return items.map((item) => itemToCleaner(item, map));
}

export interface CleanerPayload {
  cleanerName: string;
  payRatePerHour?: number;
  accountName?: string;
  bsb?: string;
  accountNumber?: string;
  active?: boolean;
  type?: "cleaner" | "contractor";
}

/** Build fields using internal names. Never send LinkTitle; use Title for cleaner name. */
function payloadToFields(
  payload: CleanerPayload,
  map: Record<string, string>
): Record<string, unknown> {
  const fields: Record<string, unknown> = {};
  if (payload.cleanerName !== undefined) {
    let k = requireMapKey(map, "Cleaner Name");
    if (k === "LinkTitle") k = "Title";
    fields[k] = payload.cleanerName;
  }
  if (payload.payRatePerHour !== undefined) {
    const k = map["Pay Rate"];
    if (k) fields[k] = payload.payRatePerHour;
  }
  if (payload.accountName !== undefined) {
    const k = map["Account Name"];
    if (k) fields[k] = payload.accountName;
  }
  if (payload.bsb !== undefined) {
    const k = map["BSB"];
    if (k) fields[k] = payload.bsb;
  }
  if (payload.accountNumber !== undefined) {
    const k = map["Account Number"];
    if (k) fields[k] = payload.accountNumber;
  }
  if (payload.active !== undefined) {
    const k = map["Active"];
    if (k) fields[k] = payload.active;
  }
  if (payload.type !== undefined) {
    const k = map["Worker Type"];
    if (k) fields[k] = workerTypeToSharePointChoice(payload.type);
  }
  return fields;
}

/** Create a new cleaner (list item). */
export async function createCleaner(accessToken: string, payload: CleanerPayload): Promise<CleanerItem> {
  const { siteId, listId } = await getSiteAndListId(accessToken);
  const map = await getCleanersFieldMap(accessToken);
  const fields = payloadToFields(
    {
      cleanerName: payload.cleanerName,
      payRatePerHour: payload.payRatePerHour ?? 0,
      accountName: payload.accountName ?? "",
      bsb: payload.bsb ?? "",
      accountNumber: payload.accountNumber ?? "",
      active: payload.active !== false,
      type: payload.type ?? "cleaner",
    },
    map
  );
  const created = await sharepoint.createListItem(accessToken, siteId, listId, fields);
  return itemToCleaner({ id: created.id, fields: created.fields ?? {} }, map);
}

/** Update an existing cleaner (list item). */
export async function updateCleaner(
  accessToken: string,
  itemId: string,
  payload: CleanerPayload
): Promise<void> {
  const { siteId, listId } = await getSiteAndListId(accessToken);
  const map = await getCleanersFieldMap(accessToken);
  const fields = payloadToFields(payload, map);
  if (Object.keys(fields).length === 0) return;
  await sharepoint.updateListItem(accessToken, siteId, listId, itemId, fields);
}

/** Delete a cleaner (list item). Item is moved to SharePoint recycle bin. */
export async function deleteCleaner(accessToken: string, itemId: string): Promise<void> {
  const { siteId, listId } = await getSiteAndListId(accessToken);
  await sharepoint.deleteListItem(accessToken, siteId, listId, itemId);
}
