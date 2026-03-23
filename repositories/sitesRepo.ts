import * as sharepoint from "../lib/sharepoint";

const SITES_LIST_NAME = "CleanTrack Sites";

/** Optional budget for toAppSite (SiteBudgetHours from budgetsRepo). */
export type SiteBudgetForApp = {
  fortnightCap: number;
  sunday?: number;
  monday?: number;
  tuesday?: number;
  wednesday?: number;
  thursday?: number;
  friday?: number;
  saturday?: number;
  /** Optional Week 2 hours when Visit Frequency is Fortnightly. */
  week2Sunday?: number;
  week2Monday?: number;
  week2Tuesday?: number;
  week2Wednesday?: number;
  week2Thursday?: number;
  week2Friday?: number;
  week2Saturday?: number;
  visitFrequency?: string;
  /** Hourly rates for budgeted labour ($/hr). */
  weekdayLabourRate?: number;
  saturdayLabourRate?: number;
  sundayLabourRate?: number;
  phLabourRate?: number;
  /** @deprecated Legacy SharePoint column; merged into weekday / weekend in UI. */
  budgetLabourRate?: number;
  weekendLabourRate?: number;
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
  noServicePeriods?: SiteNoServicePeriod[];
}

export interface SiteNoServicePeriod {
  label?: string;
  start_date: string;
  end_date: string;
  reason?: string;
  /** "manual" | "school_holidays_auto" — optional for backward compatibility */
  source?: string;
  state?: string;
  year?: number;
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

function normalizeNoServicePeriod(p: unknown): SiteNoServicePeriod | null {
  if (!p || typeof p !== "object") return null;
  const o = p as Record<string, unknown>;
  const start = String(o.start_date ?? "").trim();
  const end = String(o.end_date ?? "").trim();
  if (!start || !end) return null;
  const label = String(o.label ?? "").trim();
  const reason = String(o.reason ?? "").trim();
  const source = String(o.source ?? "").trim();
  const state = String(o.state ?? "").trim();
  let year: number | undefined;
  const yr = o.year;
  if (typeof yr === "number" && !Number.isNaN(yr)) year = yr;
  else if (typeof yr === "string" && yr.trim()) {
    const n = parseInt(yr, 10);
    if (!Number.isNaN(n)) year = n;
  }
  return {
    start_date: start,
    end_date: end,
    ...(label ? { label } : {}),
    ...(reason ? { reason } : {}),
    ...(source ? { source } : {}),
    ...(state ? { state } : {}),
    ...(year !== undefined ? { year } : {}),
  };
}

/** Display names we accept (SharePoint column display name → internal name via map). Includes common typos. */
const NO_SERVICE_PERIODS_DISPLAY_NAMES = [
  "No Service Periods",
  "NoServicePeriods",
  "No Service Perioids",
  "No Service Periouds",
] as const;

function resolveNoServicePeriodsFieldKey(map: Record<string, string>): string | undefined {
  for (const name of NO_SERVICE_PERIODS_DISPLAY_NAMES) {
    const internal = map[name];
    if (internal) return internal;
  }
  return undefined;
}

function parseNoServicePeriods(
  fields: Record<string, unknown>,
  map: Record<string, string>
): SiteNoServicePeriod[] {
  const key = resolveNoServicePeriodsFieldKey(map);
  if (!key) return [];
  const raw = getField<unknown>(fields, key);
  if (raw == null || raw === "") return [];
  let arr: unknown[] = [];
  if (Array.isArray(raw)) arr = raw;
  else if (typeof raw === "string") {
    try {
      const parsed = JSON.parse(raw);
      if (Array.isArray(parsed)) arr = parsed;
    } catch {
      arr = [];
    }
  }
  return arr.map(normalizeNoServicePeriod).filter((x): x is SiteNoServicePeriod => !!x);
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
    id: sharepoint.normalizeListItemId(item.id),
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
    noServicePeriods: parseNoServicePeriods(fields, map),
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
  noServicePeriods?: SiteNoServicePeriod[];
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
    if (k) {
      const val = payload.monthlyRevenue;
      fields[k] = val == null ? null : Math.round(Number(val) * 100) / 100;
    }
  }
  if (payload.noServicePeriods !== undefined) {
    const k = resolveNoServicePeriodsFieldKey(map);
    if (k) {
      const cleaned = payload.noServicePeriods
        .map((p) => normalizeNoServicePeriod(p))
        .filter((x): x is SiteNoServicePeriod => !!x);
      fields[k] = cleaned.length > 0 ? JSON.stringify(cleaned) : null;
    }
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
  daily_budgets_week2?: number[];
  assigned_cleaner_ids: string[];
  financial_budget: number;
  cleaner_rates: Record<string, number>;
  visit_frequency?: string;
  budget_labour_rate?: number;
  budget_weekday_labour_rate?: number;
  budget_saturday_labour_rate?: number;
  budget_sunday_labour_rate?: number;
  budget_ph_labour_rate?: number;
  no_service_periods?: SiteNoServicePeriod[];
} {
  const fortnightHours = budget ? budget.fortnightCap : 0;
  const dailyBudgets = budget
    ? [
        budget.sunday ?? 0,
        budget.monday ?? 0,
        budget.tuesday ?? 0,
        budget.wednesday ?? 0,
        budget.thursday ?? 0,
        budget.friday ?? 0,
        budget.saturday ?? 0,
      ]
    : [0, 0, 0, 0, 0, 0, 0];
  const dailyBudgetsWeek2 =
    budget &&
    [
      budget.week2Sunday,
      budget.week2Monday,
      budget.week2Tuesday,
      budget.week2Wednesday,
      budget.week2Thursday,
      budget.week2Friday,
      budget.week2Saturday,
    ].some((v) => v != null)
      ? [
          budget.week2Sunday ?? 0,
          budget.week2Monday ?? 0,
          budget.week2Tuesday ?? 0,
          budget.week2Wednesday ?? 0,
          budget.week2Thursday ?? 0,
          budget.week2Friday ?? 0,
          budget.week2Saturday ?? 0,
        ]
      : undefined;
  return {
    id: s.id,
    name: s.siteName,
    address: s.address,
    is_active: s.active,
    monthly_revenue: s.monthlyRevenue ?? 0,
    budgeted_hours_per_fortnight: fortnightHours,
    daily_budgets: dailyBudgets,
    ...(dailyBudgetsWeek2 ? { daily_budgets_week2: dailyBudgetsWeek2 } : {}),
    assigned_cleaner_ids: [],
    financial_budget: 0,
    cleaner_rates: {},
    ...(budget?.visitFrequency && { visit_frequency: budget.visitFrequency }),
    /* Weekday Labour Rate (schema) → budget_labour_rate + budget_weekday_labour_rate; legacy Budget Labour Rate fallback */
    ...((budget?.weekdayLabourRate ?? budget?.budgetLabourRate) != null && { budget_labour_rate: budget.weekdayLabourRate ?? budget.budgetLabourRate ?? 0 }),
    ...(budget?.weekdayLabourRate != null && budget.weekdayLabourRate >= 0 && { budget_weekday_labour_rate: budget.weekdayLabourRate }),
    ...(budget?.saturdayLabourRate != null && budget.saturdayLabourRate >= 0 && { budget_saturday_labour_rate: budget.saturdayLabourRate }),
    ...(budget?.sundayLabourRate != null && budget.sundayLabourRate >= 0 && { budget_sunday_labour_rate: budget.sundayLabourRate }),
    ...(budget?.phLabourRate != null && budget.phLabourRate >= 0 && { budget_ph_labour_rate: budget.phLabourRate }),
    ...(s.noServicePeriods && s.noServicePeriods.length > 0 && { no_service_periods: s.noServicePeriods }),
  };
}
