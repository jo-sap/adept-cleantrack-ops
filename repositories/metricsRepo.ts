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
  /** Optional display names from the Site and Cleaner lookups, used as a fallback join key if ids drift. */
  siteName?: string;
  cleanerName?: string;
  pay_rate_snapshot?: number;
  adhocJobId?: string;
  adhocJobName?: string;
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

  const [timesheetItems, cleanerItems, timesheetColumns] = await Promise.all([
    sharepoint.getListItems(accessToken, siteId, timesheetsListId),
    cleanersListId ? sharepoint.getListItems(accessToken, siteId, cleanersListId) : [],
    sharepoint.getListColumns(accessToken, siteId, timesheetsListId),
  ]);

  const tsMap: Record<string, string> = {};
  for (const c of timesheetColumns) {
    if (c.displayName) tsMap[c.displayName] = c.name;
  }
  const workDateInternal =
    tsMap["Work Date"] ?? "WorkDate" ?? "Work_x0020_Date";
  const hoursInternal = tsMap["Hours"] ?? "Hours";
  const adHocJobCol = tsMap["Ad Hoc Job"] ?? "Ad_x0020_Hoc_x0020_Job";
  const adHocJobIdKey = `${adHocJobCol}LookupId`;
  const adHocJobNameKey = adHocJobCol;

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

  const workDateKey = workDateInternal;
  const hoursKey = hoursInternal;
  const result: TimesheetEntryFlat[] = [];

  for (const item of timesheetItems) {
    const f = (item.fields ?? {}) as Record<string, unknown>;
    const rawWorkDate = f[workDateKey] ?? f["WorkDate"];
    const workDate = toDate(rawWorkDate);
    if (!workDate || !dateInRange(workDate, range)) continue;
    const hours = toNum(f[hoursKey] ?? f["Hours"]);
    const { id: siteIdRaw, name: siteNameRaw } = getLookupIdOrName(f, "Site");
    const { id: cleanerIdRaw, name: cleanerNameRaw } = getLookupIdOrName(f, "Cleaner");
    if (!siteIdRaw || !cleanerIdRaw) continue;
    // Normalize lookup ids so they match app-level site.id / cleaner.id (Graph may return path-style ids).
    const siteIdVal = sharepoint.normalizeListItemId(siteIdRaw);
    const cleanerIdVal = sharepoint.normalizeListItemId(cleanerIdRaw);
    const rate = cleanerRateMap[cleanerIdVal] ?? 0;
    // Use the raw SharePoint date string (first 10 chars) to avoid timezone shifts,
    // falling back to the ISO date only if needed.
    const dateStr =
      typeof rawWorkDate === "string" && rawWorkDate.length >= 10
        ? rawWorkDate.slice(0, 10)
        : workDate.toISOString().slice(0, 10);
    const adhocJobIdVal = f[adHocJobIdKey] != null ? String(f[adHocJobIdKey]).trim() : "";
    let adhocJobNameVal = "";
    const ahVal = f[adHocJobNameKey];
    if (ahVal != null) {
      if (typeof ahVal === "object" && ahVal !== null && "LookupValue" in (ahVal as object))
        adhocJobNameVal = String((ahVal as { LookupValue?: string }).LookupValue ?? "").trim();
      else adhocJobNameVal = String(ahVal).trim();
    }
    result.push({
      id: item.id!,
      date: dateStr,
      hours,
      siteId: siteIdVal,
      cleanerId: cleanerIdVal,
      siteName: siteNameRaw || undefined,
      cleanerName: cleanerNameRaw || undefined,
      pay_rate_snapshot: rate || undefined,
      ...(adhocJobIdVal
        ? { adhocJobId: adhocJobIdVal, adhocJobName: adhocJobNameVal || undefined }
        : {}),
    });
  }
  if (typeof process !== "undefined" && process.env?.NODE_ENV === "development" && result.length > 0) {
    const withAdHoc = result.filter((e) => e.adhocJobId);
    console.log("[CleanTrack Timesheet Entries] range count:", result.length, "with Ad Hoc Job:", withAdHoc.length, "sample:", result[0]);
  }
  return result;
}

/** Payload for one timesheet row (site + cleaner + date + hours; optional ad hoc job). */
export interface TimesheetEntryPayload {
  siteId: string;
  cleanerId: string;
  date: string;
  hours: number;
  /** Optional CleanTrack Ad Hoc Jobs list item id (lookup). */
  adhocJobId?: string | null;
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
  const adHocJobDisplay = "Ad Hoc Job";
  const adHocJobInternal = map[adHocJobDisplay] ?? "Ad_x0020_Hoc_x0020_Job";
  const adHocJobKey = adHocJobInternal === "Ad Hoc Job" ? "Ad_x0020_Hoc_x0020_JobLookupId" : `${adHocJobInternal}LookupId`;

  const existing = await getTimesheetEntriesForRange(accessToken, range);
  const keyToId = new Map<string, string>();
  for (const e of existing) {
    keyToId.set(`${e.siteId}|${e.cleanerId}|${e.date}`, e.id);
  }

  let saved = 0;
  const errors: string[] = [];
  for (const entry of entries) {
    const key = `${entry.siteId}|${entry.cleanerId}|${entry.date}`;
    const existingId = keyToId.get(key);
    const siteLookupVal = /^\d+$/.test(entry.siteId) ? parseInt(entry.siteId, 10) : entry.siteId;
    const cleanerLookupVal = /^\d+$/.test(entry.cleanerId) ? parseInt(entry.cleanerId, 10) : entry.cleanerId;
    const workDate = entry.date; // yyyy-MM-dd
    const adHocLookupVal =
      entry.adhocJobId != null && entry.adhocJobId !== ""
        ? (/^\d+$/.test(entry.adhocJobId) ? parseInt(entry.adhocJobId, 10) : entry.adhocJobId)
        : null;

    try {
      if (existingId) {
        const updateFields: Record<string, unknown> = { [hoursKey]: entry.hours };
        if (adHocJobKey) updateFields[adHocJobKey] = adHocLookupVal;
        if (typeof console !== "undefined") {
          console.log("[Timesheets Save] PATCH entry", {
            id: existingId,
            workDate,
            hours: entry.hours,
            siteId: entry.siteId,
            cleanerId: entry.cleanerId,
          });
        }
        await sharepoint.updateListItem(accessToken, siteId, listId, existingId, updateFields);
      } else {
        const createFields: Record<string, unknown> = {
          [workDateKey]: workDate,
          [hoursKey]: entry.hours,
          [siteKey]: siteLookupVal,
          [cleanerKey]: cleanerLookupVal,
        };
        if (adHocJobKey && adHocLookupVal != null) createFields[adHocJobKey] = adHocLookupVal;
        if (typeof console !== "undefined") {
          console.log("[Timesheets Save] POST entry", {
            workDate,
            hours: entry.hours,
            siteId: entry.siteId,
            cleanerId: entry.cleanerId,
          });
        }
        await sharepoint.createListItem(accessToken, siteId, listId, createFields);
      }
      saved++;
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      const ctx = `date=${workDate}, siteId=${entry.siteId}, cleanerId=${entry.cleanerId}`;
      const full = `[Timesheets Save] Failed for ${ctx}: ${msg}`;
      if (typeof console !== "undefined") {
        console.error(full);
      }
      errors.push(full);
      // continue with other entries so partial saves still go through
    }
  }
  if (errors.length > 0) {
    return { saved, error: errors.join(" | ") };
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
