import * as sharepoint from "../lib/sharepoint";

/** end is exclusive */
export type DateRange = { start: Date; end: Date };

export interface DashboardMetrics {
  portfolioRevenue: number;
  laborExpenses: number;
  netGrossProfit: number;
  profitMargin: number;
}

const SITES_LIST_NAME = "CleanTrack Sites";
const CLEANERS_LIST_NAME = "CleanTrack Cleaners";
const TIMESHEETS_LIST_NAME = "CleanTrack Timesheet Entries";

function toDate(v: unknown): Date | null {
  if (v == null) return null;
  if (typeof v === "string") {
    const d = new Date(v);
    return isNaN(d.getTime()) ? null : d;
  }
  return null;
}

function toNum(v: unknown): number {
  if (typeof v === "number" && !Number.isNaN(v)) return v;
  if (typeof v === "string") {
    const n = parseFloat(String(v).replace(/[^0-9.-]/g, ""));
    return Number.isNaN(n) ? 0 : n;
  }
  return 0;
}

function isActive(f: Record<string, unknown>): boolean {
  const key = Object.keys(f).find((k) => k === "Active" || k.toLowerCase() === "active");
  if (!key) return true;
  const v = f[key];
  if (v === false || v === "No" || String(v).toLowerCase() === "no") return false;
  return true;
}

/** start <= d < end */
function dateInRange(d: Date, range: DateRange): boolean {
  const t = d.getTime();
  return t >= range.start.getTime() && t < range.end.getTime();
}

function daysInMonth(date: Date): number {
  const y = date.getFullYear();
  const m = date.getMonth();
  return new Date(y, m + 1, 0).getDate();
}

function monthStart(d: Date): Date {
  return new Date(d.getFullYear(), d.getMonth(), 1);
}

function monthEnd(d: Date): Date {
  return new Date(d.getFullYear(), d.getMonth() + 1, 0, 23, 59, 59, 999);
}

/** Overlap in days between range [start, end) and month. */
function overlapDaysWithMonth(
  rangeStart: Date,
  rangeEndExclusive: Date,
  monthDate: Date
): number {
  const mStart = monthStart(monthDate);
  const mEnd = monthEnd(monthDate);
  const start = Math.max(rangeStart.getTime(), mStart.getTime());
  const end = Math.min(rangeEndExclusive.getTime(), mEnd.getTime() + 1);
  if (end <= start) return 0;
  return Math.round((end - start) / (24 * 60 * 60 * 1000));
}

/** Pro-rated revenue from CleanTrack Sites: Monthly Revenue, Active; range end exclusive. */
function computePortfolioRevenue(
  siteItems: sharepoint.GraphListItem[],
  range: DateRange
): number {
  if (siteItems.length === 0) return 0;
  const monthlyRevKey =
    Object.keys(siteItems[0].fields ?? {}).find(
      (k) => k === "Monthly Revenue" || k.toLowerCase() === "monthly revenue"
    ) ?? "Monthly Revenue";

  let total = 0;
  for (const item of siteItems) {
    const f = item.fields ?? {};
    if (!isActive(f)) continue;
    const monthlyRev = toNum(f[monthlyRevKey]);
    if (monthlyRev <= 0) continue;

    const months = new Set<string>();
    let d = new Date(range.start.getFullYear(), range.start.getMonth(), 1);
    while (d < range.end) {
      months.add(`${d.getFullYear()}-${d.getMonth()}`);
      d = new Date(d.getFullYear(), d.getMonth() + 1, 1);
    }

    for (const key of months) {
      const [y, m] = key.split("-").map(Number);
      const monthDate = new Date(y, m, 1);
      const daysInMo = daysInMonth(monthDate);
      const overlap = overlapDaysWithMonth(range.start, range.end, monthDate);
      total += monthlyRev * (overlap / daysInMo);
    }
  }
  return total;
}

/** Resolve lookup id or display name from a SharePoint lookup field. */
function getLookupIdOrName(f: Record<string, unknown>, columnBase: string): { id: string; name: string } {
  const idKey =
    Object.keys(f).find((k) => k === `${columnBase}LookupId` || k === `${columnBase}Id`) ??
    Object.keys(f).find((k) => k.toLowerCase() === `${columnBase.toLowerCase()}lookupid`);
  const nameKey =
    Object.keys(f).find((k) => k === columnBase) ??
    Object.keys(f).find((k) => k.toLowerCase() === columnBase.toLowerCase());
  const id = idKey ? String(f[idKey] ?? "").trim() : "";
  const name = nameKey ? String(f[nameKey] ?? "").trim() : "";
  return { id, name };
}

/** Build map: cleaner id -> pay rate, and cleaner name (normalized) -> pay rate. */
function buildCleanerRateMap(
  cleanerItems: sharepoint.GraphListItem[]
): Record<string, number> {
  const map: Record<string, number> = {};
  if (cleanerItems.length === 0) return map;
  const nameKey =
    Object.keys(cleanerItems[0].fields ?? {}).find(
      (k) => k === "Cleaner Name" || k.toLowerCase() === "cleaner name"
    ) ?? "Cleaner Name";
  const payRateKey =
    Object.keys(cleanerItems[0].fields ?? {}).find(
      (k) => k === "Pay Rate" || k.toLowerCase() === "pay rate"
    ) ?? "Pay Rate";

  for (const item of cleanerItems) {
    const f = item.fields ?? {};
    if (!isActive(f)) continue;
    const name = String(f[nameKey] ?? "").trim();
    const rate = toNum(f[payRateKey]);
    if (item.id) map[item.id] = rate;
    if (name) map[name.toLowerCase()] = rate;
  }
  return map;
}

/** Labor = sum(Hours * Cleaner.PayRate) for entries in range. */
function computeLaborExpenses(
  timesheetItems: sharepoint.GraphListItem[],
  cleanerRateMap: Record<string, number>,
  range: DateRange
): { labor: number; error?: string } {
  if (timesheetItems.length === 0) return { labor: 0 };
  const workDateKey =
    Object.keys(timesheetItems[0].fields ?? {}).find(
      (k) => k === "Work Date" || k.toLowerCase() === "work date"
    ) ?? "Work Date";
  const hoursKey =
    Object.keys(timesheetItems[0].fields ?? {}).find(
      (k) => k === "Hours" || k.toLowerCase() === "hours"
    ) ?? "Hours";

  let labor = 0;
  const missingCleaners = new Set<string>();

  for (const item of timesheetItems) {
    const f = item.fields ?? {};
    const workDate = toDate(f[workDateKey]);
    if (!workDate || !dateInRange(workDate, range)) continue;
    const hours = toNum(f[hoursKey]);
    const { id: cleanerId, name: cleanerName } = getLookupIdOrName(f, "Cleaner");
    const rate =
      (cleanerId && cleanerRateMap[cleanerId] !== undefined ? cleanerRateMap[cleanerId] : null) ??
      (cleanerName ? cleanerRateMap[cleanerName.toLowerCase()] : undefined);
    if (rate === undefined || (typeof rate === "number" && Number.isNaN(rate))) {
      const label = cleanerName || cleanerId || "unknown";
      missingCleaners.add(label);
      continue;
    }
    labor += hours * rate;
  }

  if (missingCleaners.size > 0) {
    return {
      labor,
      error: `Missing Pay Rate for cleaner(s): ${[...missingCleaners].join(", ")}. Ensure CleanTrack Cleaners has Pay Rate and Active=Yes.`,
    };
  }
  return { labor };
}

/** Flat entry shape for Dashboard/SiteDetail (matches TimeEntry with siteId, cleanerId). */
export interface TimesheetEntryFlat {
  id: string;
  date: string;
  hours: number;
  siteId: string;
  cleanerId: string;
  pay_rate_snapshot?: number;
}

/** Load timesheet entries from SharePoint for the given range; returns flat list for dashboard/views. */
export async function getTimesheetEntriesForRange(
  accessToken: string,
  range: DateRange
): Promise<TimesheetEntryFlat[]> {
  const siteId = await sharepoint.getSiteId(accessToken);
  const timesheetsListId = await sharepoint.getListIdByName(accessToken, siteId, TIMESHEETS_LIST_NAME);
  const cleanersListId = await sharepoint.getListIdByName(accessToken, siteId, CLEANERS_LIST_NAME);
  if (!timesheetsListId) return [];

  const [timesheetItems, cleanerItems] = await Promise.all([
    sharepoint.getListItems(accessToken, siteId, timesheetsListId),
    cleanersListId ? sharepoint.getListItems(accessToken, siteId, cleanersListId) : [],
  ]);

  const cleanerRateMap: Record<string, number> = {};
  if (cleanerItems.length > 0) {
    const nameKey = Object.keys(cleanerItems[0].fields ?? {}).find(
      (k) => k === "Cleaner Name" || k.toLowerCase() === "cleaner name"
    ) ?? "Cleaner Name";
    const payRateKey = Object.keys(cleanerItems[0].fields ?? {}).find(
      (k) => k === "Pay Rate" || k.toLowerCase() === "pay rate"
    ) ?? "Pay Rate";
    for (const item of cleanerItems) {
      const f = item.fields ?? {};
      if (!isActive(f)) continue;
      const rate = toNum(f[payRateKey]);
      if (item.id) cleanerRateMap[item.id] = rate;
    }
  }

  const workDateKey = "Work Date";
  const hoursKey = "Hours";
  const result: TimesheetEntryFlat[] = [];

  for (const item of timesheetItems) {
    const f = (item.fields ?? {}) as Record<string, unknown>;
    const workDate = toDate(f[workDateKey] ?? f["WorkDate"]);
    if (!workDate || !dateInRange(workDate, range)) continue;
    const hours = toNum(f[hoursKey] ?? f["Hours"]);
    const { id: siteIdVal } = getLookupIdOrName(f, "Site");
    const { id: cleanerIdVal } = getLookupIdOrName(f, "Cleaner");
    if (!siteIdVal || !cleanerIdVal) continue;
    const rate = cleanerRateMap[cleanerIdVal] ?? 0;
    const dateStr = workDate.toISOString().slice(0, 10);
    result.push({
      id: item.id!,
      date: dateStr,
      hours,
      siteId: siteIdVal,
      cleanerId: cleanerIdVal,
      pay_rate_snapshot: rate || undefined,
    });
  }
  return result;
}

/** Payload for one timesheet row (site + cleaner + date + hours). */
export interface TimesheetEntryPayload {
  siteId: string;
  cleanerId: string;
  date: string;
  hours: number;
}

/**
 * Save timesheet entries to CleanTrack Timesheet Entries. For each entry, updates existing item if found (same site, cleaner, date) or creates new.
 */
export async function saveTimesheetEntriesToSharePoint(
  accessToken: string,
  range: DateRange,
  entries: TimesheetEntryPayload[]
): Promise<{ saved: number; error?: string }> {
  if (entries.length === 0) return { saved: 0 };
  const siteId = await sharepoint.getSiteId(accessToken);
  const listId = await sharepoint.getListIdByName(accessToken, siteId, TIMESHEETS_LIST_NAME);
  if (!listId) return { saved: 0, error: "CleanTrack Timesheet Entries list not found." };

  const columns = await sharepoint.getListColumns(accessToken, siteId, listId);
  const map: Record<string, string> = {};
  for (const c of columns) {
    if (c.displayName) map[c.displayName] = c.name;
  }
  const workDateKey = map["Work Date"] ?? "WorkDate" ?? "Work_x0020_Date";
  const hoursKey = map["Hours"] ?? "Hours";
  const siteKey = map["Site"] === "Site" ? "SiteLookupId" : (map["Site"] ?? "SiteLookupId");
  const cleanerKey = map["Cleaner"] === "Cleaner" ? "CleanerLookupId" : (map["Cleaner"] ?? "CleanerLookupId");

  const existing = await getTimesheetEntriesForRange(accessToken, range);
  const keyToId = new Map<string, string>();
  for (const e of existing) {
    keyToId.set(`${e.siteId}|${e.cleanerId}|${e.date}`, e.id);
  }

  let saved = 0;
  for (const entry of entries) {
    const key = `${entry.siteId}|${entry.cleanerId}|${entry.date}`;
    const existingId = keyToId.get(key);
    const siteLookupVal = /^\d+$/.test(entry.siteId) ? parseInt(entry.siteId, 10) : entry.siteId;
    const cleanerLookupVal = /^\d+$/.test(entry.cleanerId) ? parseInt(entry.cleanerId, 10) : entry.cleanerId;
    const workDate = entry.date; // yyyy-MM-dd

    try {
      if (existingId) {
        await sharepoint.updateListItem(accessToken, siteId, listId, existingId, {
          [hoursKey]: entry.hours,
        });
      } else {
        await sharepoint.createListItem(accessToken, siteId, listId, {
          [workDateKey]: workDate,
          [hoursKey]: entry.hours,
          [siteKey]: siteLookupVal,
          [cleanerKey]: cleanerLookupVal,
        });
      }
      saved++;
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      return { saved, error: msg };
    }
  }
  return { saved };
}

export async function getDashboardMetrics(
  accessToken: string,
  range: DateRange,
  options?: { assignedSiteIds?: string[] }
): Promise<{ metrics: DashboardMetrics; error?: string }> {
  const zeros = {
    portfolioRevenue: 0,
    laborExpenses: 0,
    netGrossProfit: 0,
    profitMargin: 0,
  };

  try {
    const siteId = await sharepoint.getSiteId(accessToken);

    const sitesListId = await sharepoint.getListIdByName(accessToken, siteId, SITES_LIST_NAME);
    const cleanersListId = await sharepoint.getListIdByName(accessToken, siteId, CLEANERS_LIST_NAME);
    const timesheetsListId = await sharepoint.getListIdByName(accessToken, siteId, TIMESHEETS_LIST_NAME);

    if (!sitesListId || !cleanersListId || !timesheetsListId) {
      const missing = [];
      if (!sitesListId) missing.push("CleanTrack Sites");
      if (!cleanersListId) missing.push("CleanTrack Cleaners");
      if (!timesheetsListId) missing.push("CleanTrack Timesheet Entries");
      return { metrics: zeros, error: `Lists not found: ${missing.join(", ")}.` };
    }

    const [siteItems, cleanerItems, timesheetItems] = await Promise.all([
      sharepoint.getListItems(accessToken, siteId, sitesListId),
      sharepoint.getListItems(accessToken, siteId, cleanersListId),
      sharepoint.getListItems(accessToken, siteId, timesheetsListId),
    ]);

    let siteItemsForRevenue = siteItems;
    let timesheetItemsForLabor = timesheetItems;
    if (options?.assignedSiteIds && options.assignedSiteIds.length > 0) {
      const siteIdSet = new Set(options.assignedSiteIds);
      siteItemsForRevenue = siteItems.filter((item) => item.id && siteIdSet.has(item.id));
      const firstEntry = timesheetItems[0];
      const siteLookupKey = firstEntry
        ? Object.keys(firstEntry.fields ?? {}).find(
            (k) => k === "SiteLookupId" || k === "SiteId" || k.toLowerCase() === "sitelookupid"
          )
        : undefined;
      if (siteLookupKey) {
        timesheetItemsForLabor = timesheetItems.filter((item) => {
          const f = item.fields ?? {};
          const val = (f as Record<string, unknown>)[siteLookupKey];
          const id = val != null ? String(val) : "";
          return siteIdSet.has(id);
        });
      }
    }

    const portfolioRevenue = computePortfolioRevenue(siteItemsForRevenue, range);

    const cleanerRateMap = buildCleanerRateMap(cleanerItems);
    const laborResult = computeLaborExpenses(timesheetItemsForLabor, cleanerRateMap, range);
    const laborExpenses = laborResult.labor;

    const netGrossProfit = portfolioRevenue - laborExpenses;
    const profitMargin = portfolioRevenue > 0 ? netGrossProfit / portfolioRevenue : 0;

    return {
      metrics: {
        portfolioRevenue,
        laborExpenses,
        netGrossProfit,
        profitMargin,
      },
      error: laborResult.error,
    };
  } catch (err) {
    return {
      metrics: zeros,
      error: err instanceof Error ? err.message : "Failed to load KPIs.",
    };
  }
}
