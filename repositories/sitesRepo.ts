import * as sharepoint from "../lib/sharepoint";

const SITES_LIST_NAME = "CleanTrack Sites";

/** Optional budget for toAppSite (SiteBudgetHours from budgetsRepo). */
export type SiteBudgetForApp = {
  fortnightCap: number;
  sunday: number;
  monday: number;
  tuesday: number;
  wednesday: number;
  thursday: number;
  friday: number;
  saturday: number;
  visitFrequency?: string;
};

/** Display name -> internal name. Cached in-memory. */
let cachedSitesFieldMap: Record<string, string> | null = null;

export interface Site {
  id: string;
  siteName: string;
  address: string;
  state: string;
  active: boolean;
  monthlyRevenue: number | null;
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

function getBoolean(fields: Record<string, unknown>, key: string | undefined, defaultVal: boolean): boolean {
  if (!key) return defaultVal;
  const v = (fields as Record<string, unknown>)[key];
  if (v === undefined || v === null) return defaultVal;
  if (v === false || v === "No" || String(v).toLowerCase() === "no") return false;
  if (v === true || v === "Yes" || String(v).toLowerCase() === "yes") return true;
  return defaultVal;
}

function getNumber(fields: Record<string, unknown>, ...keys: string[]): number | null {
  const v = getField<unknown>(fields, ...keys);
  if (v == null || v === "") return null;
  if (typeof v === "number" && !Number.isNaN(v)) return v;
  const n = parseFloat(String(v).replace(/[^0-9.-]/g, ""));
  return Number.isNaN(n) ? null : n;
}

function requireMapKey(map: Record<string, string>, displayName: string): string {
  const internal = map[displayName];
  if (!internal) {
    throw new Error(`SharePoint column not found: ${displayName}. Check list column display name.`);
  }
  return internal;
}

function itemToSite(item: sharepoint.GraphListItem, map: Record<string, string>): Site {
  const fields = item.fields ?? {};
  const siteNameKey = map["Site Name"] ?? "Site Name";
  const siteName =
    getString(fields, siteNameKey, "Site_x0020_Name", "Title") ||
    getString(fields, "Title") ||
    "";
  const addressKey = map["Address"] ?? "Address";
  const stateKey = map["State"] ?? "State";
  const activeKey = map["Active"] ?? "Active";
  const monthlyKey = map["Monthly Revenue"] ?? "Monthly Revenue";
  return {
    id: item.id,
    siteName,
    address: getString(fields, addressKey, "Address"),
    state: getString(fields, stateKey, "State"),
    active: getBoolean(fields, activeKey, true),
    monthlyRevenue: getNumber(
      fields,
      monthlyKey,
      "Monthly_x0020_Revenue",
      "Monthly Revenue"
    ),
  };
}

/** Resolve site id and CleanTrack Sites list id. */
async function getSiteAndListId(accessToken: string): Promise<{ siteId: string; listId: string }> {
  const siteId = await sharepoint.getSiteId(accessToken);
  const listId = await sharepoint.getListIdByName(accessToken, siteId, SITES_LIST_NAME);
  if (!listId) throw new Error(`List "${SITES_LIST_NAME}" not found.`);
  return { siteId, listId };
}

/** Get displayName -> internal name map for CleanTrack Sites. Cached in-memory. */
export async function getSitesFieldMap(accessToken: string): Promise<Record<string, string>> {
  if (cachedSitesFieldMap) return cachedSitesFieldMap;
  const { siteId, listId } = await getSiteAndListId(accessToken);
  const columns = await sharepoint.getListColumns(accessToken, siteId, listId);
  const map: Record<string, string> = {};
  for (const col of columns) {
    if (col.displayName) map[col.displayName] = col.name;
  }
  cachedSitesFieldMap = map;
  return map;
}

/** Fetch all sites from CleanTrack Sites list. Optionally filter to assigned site IDs (for Manager role). */
export async function getSites(
  accessToken: string,
  options?: { assignedSiteIds?: string[] }
): Promise<Site[]> {
  const { siteId, listId } = await getSiteAndListId(accessToken);
  const map = await getSitesFieldMap(accessToken);
  const items = await sharepoint.getListItems(accessToken, siteId, listId);
  let list = items.map((item) => itemToSite(item, map));
  if (options?.assignedSiteIds && options.assignedSiteIds.length > 0) {
    const set = new Set(options.assignedSiteIds);
    list = list.filter((s) => set.has(s.id));
  }
  return list;
}

/** Payload for create/update. */
export interface SitePayload {
  siteName: string;
  address?: string;
  state?: string;
  active?: boolean;
  monthlyRevenue?: number | null;
}

/** Build fields object using internal column names. Throws if a required display name is missing from map.
 * Never sends LinkTitle (read-only); uses Title for the site name when creating/updating. */
function payloadToFields(
  payload: SitePayload,
  map: Record<string, string>
): Record<string, unknown> {
  const fields: Record<string, unknown> = {};
  if (payload.siteName !== undefined) {
    let k = requireMapKey(map, "Site Name");
    if (k === "LinkTitle") k = "Title";
    fields[k] = payload.siteName;
  }
  if (payload.address !== undefined) {
    const k = map["Address"];
    if (k) fields[k] = payload.address;
  }
  if (payload.state !== undefined) {
    const k = map["State"];
    if (k) fields[k] = payload.state;
  }
  if (payload.active !== undefined) {
    const k = map["Active"];
    if (k) fields[k] = payload.active;
  }
  if (payload.monthlyRevenue !== undefined) {
    const k = map["Monthly Revenue"];
    if (k) fields[k] = payload.monthlyRevenue ?? null;
  }
  return fields;
}

/** Create a new site (list item). Uses siteId from getSiteId(token) and listId from getListIdByName(..., "CleanTrack Sites"). */
export async function createSite(accessToken: string, payload: SitePayload): Promise<Site> {
  const { siteId, listId } = await getSiteAndListId(accessToken);
  const map = await getSitesFieldMap(accessToken);
  const fields = payloadToFields(
    {
      siteName: payload.siteName,
      address: payload.address ?? "",
      state: payload.state ?? "",
      active: payload.active !== false,
      monthlyRevenue: payload.monthlyRevenue ?? null,
    },
    map
  );
  const created = await sharepoint.createListItem(accessToken, siteId, listId, fields);
  return itemToSite({ id: created.id, fields: created.fields ?? {} }, map);
}

/** Update an existing site by list item id. */
export async function updateSite(
  accessToken: string,
  id: string,
  payload: Partial<SitePayload>
): Promise<void> {
  const { siteId, listId } = await getSiteAndListId(accessToken);
  const map = await getSitesFieldMap(accessToken);
  const fields = payloadToFields(payload as SitePayload, map);
  if (Object.keys(fields).length === 0) return;
  await sharepoint.updateListItem(accessToken, siteId, listId, id, fields);
}

/** Set Active flag only. */
export async function setSiteActive(accessToken: string, id: string, active: boolean): Promise<void> {
  return updateSite(accessToken, id, { active });
}

/** Delete a site (list item). Item is moved to SharePoint recycle bin. */
export async function deleteSite(accessToken: string, siteListItemId: string): Promise<void> {
  const { siteId, listId } = await getSiteAndListId(accessToken);
  await sharepoint.deleteListItem(accessToken, siteId, listId, siteListItemId);
}

/** Map repo Site to app Site shape (id, name, address, is_active, monthly_revenue, budgeted_hours, daily_budgets, etc.). */
export function toAppSite(
  s: Site,
  budget?: SiteBudgetForApp
): {
  id: string;
  name: string;
  address: string;
  is_active: boolean;
  monthly_revenue: number;
  budgeted_hours_per_fortnight: number;
  daily_budgets: number[];
  assigned_cleaner_ids: string[];
  financial_budget: number;
  cleaner_rates: Record<string, number>;
  visit_frequency?: string;
} {
  const fortnightHours = budget ? budget.fortnightCap : 0;
  const dailyBudgets = budget
    ? [budget.sunday, budget.monday, budget.tuesday, budget.wednesday, budget.thursday, budget.friday, budget.saturday]
    : [0, 0, 0, 0, 0, 0, 0];
  return {
    id: s.id,
    name: s.siteName,
    address: s.address,
    is_active: s.active,
    monthly_revenue: s.monthlyRevenue ?? 0,
    budgeted_hours_per_fortnight: fortnightHours,
    daily_budgets: dailyBudgets,
    assigned_cleaner_ids: [],
    financial_budget: 0,
    cleaner_rates: {},
    ...(budget?.visitFrequency && { visit_frequency: budget.visitFrequency }),
  };
}
