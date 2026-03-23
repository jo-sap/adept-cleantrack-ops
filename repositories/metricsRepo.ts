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

function pad2(n: number): string {
  return String(n).padStart(2, "0");
}

/** Format a Date as YYYY-MM-DD in local time (avoids timezone shifts from toISOString()). */
function toLocalYmd(d: Date): string {
  return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
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
  const startYmd = toLocalYmd(range.start);
  const endYmdExclusive = toLocalYmd(range.end);

  for (const item of timesheetItems) {
    const f = (item.fields ?? {}) as Record<string, unknown>;
    const rawWorkDate = f[workDateKey] ?? f["WorkDate"];
    const workDate = toDate(rawWorkDate);
    if (!workDate) continue;
    const hours = toNum(f[hoursKey] ?? f["Hours"]);
    const { id: siteIdRaw, name: siteNameRaw } = getLookupIdOrName(f, "Site");
    const { id: cleanerIdRaw, name: cleanerNameRaw } = getLookupIdOrName(f, "Cleaner");
    if (!siteIdRaw || !cleanerIdRaw) continue;
    // Normalize lookup ids so they match app-level site.id / cleaner.id (Graph may return path-style ids).
    const siteIdVal = sharepoint.normalizeListItemId(siteIdRaw);
    const cleanerIdVal = sharepoint.normalizeListItemId(cleanerIdRaw);
    const rate = cleanerRateMap[cleanerIdVal] ?? 0;
    // Normalize Work Date to YYYY-MM-DD.
    // IMPORTANT: avoid Date.toISOString() here because it can shift the day (timezone).
    // Graph may return ISO strings ("2026-03-09T00:00:00Z") or locale-ish strings ("3/9/2026").
    const dateStr =
      typeof rawWorkDate === "string" && /^\d{4}-\d{2}-\d{2}/.test(rawWorkDate)
        ? rawWorkDate.slice(0, 10)
        : toLocalYmd(workDate);
    // Range check using YYYY-MM-DD string comparison (timezone-safe).
    if (dateStr < startYmd || dateStr >= endYmdExclusive) continue;
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
  const siteInternal = map["Site"] ?? "Site";
  const cleanerInternal = map["Cleaner"] ?? "Cleaner";
  const siteLookupIdKey = siteInternal === "Site" ? "SiteLookupId" : `${siteInternal}LookupId`;
  const cleanerLookupIdKey = cleanerInternal === "Cleaner" ? "CleanerLookupId" : `${cleanerInternal}LookupId`;
  const adHocJobDisplay = "Ad Hoc Job";
  const adHocJobInternal = map[adHocJobDisplay] ?? "Ad_x0020_Hoc_x0020_Job";
  const adHocJobKey = adHocJobInternal === "Ad Hoc Job" ? "Ad_x0020_Hoc_x0020_JobLookupId" : `${adHocJobInternal}LookupId`;

  // IMPORTANT: We use "replace" semantics to avoid duplicate creation when Graph read-after-write is laggy
  // or when date parsing causes mismatched keys. For each (site, cleaner, adhoc) in this save:
  // 1) delete all existing entries in this range for that composite
  // 2) create fresh entries for dates with hours > 0
  const startYmd = toLocalYmd(range.start);
  const endYmdExclusive = toLocalYmd(range.end);

  const byComposite = new Map<string, TimesheetEntryPayload[]>();
  for (const e of entries) {
    const sId = sharepoint.normalizeListItemId(e.siteId);
    const cId = sharepoint.normalizeListItemId(e.cleanerId);
    const aId = e.adhocJobId != null ? sharepoint.normalizeListItemId(e.adhocJobId) : "";
    const composite = `${sId}|${cId}|${aId}`;
    if (!byComposite.has(composite)) byComposite.set(composite, []);
    byComposite.get(composite)!.push({ ...e, siteId: sId, cleanerId: cId, adhocJobId: aId || null });
  }

  // Fetch current list items once, then filter client-side (avoids Graph 400 on unindexed lookup filters).
  const fieldsSelect = Array.from(
    new Set(
      [
        workDateKey,
        siteLookupIdKey,
        cleanerLookupIdKey,
        hoursKey,
        ...(adHocJobKey ? [adHocJobKey] : []),
      ].filter(Boolean)
    )
  );
  const existingItems = await sharepoint.getListItems(accessToken, siteId, listId, fieldsSelect);

  let saved = 0;
  const errors: string[] = [];
  for (const [composite, group] of byComposite.entries()) {
    const [sId, cId, aId] = composite.split("|");
    if (!sId || !cId) continue;
    const siteLookupVal = /^\d+$/.test(sId) ? parseInt(sId, 10) : sId;
    const cleanerLookupVal = /^\d+$/.test(cId) ? parseInt(cId, 10) : cId;
    const adHocLookupVal =
      aId && aId.trim()
        ? (/^\d+$/.test(aId) ? parseInt(aId, 10) : aId)
        : null;

    // 1) Delete existing items for this composite in this range (client-side filtered),
    //    sending deletes in Graph $batch chunks (20 per request).
    const toDeleteIds: string[] = [];
    for (const item of existingItems) {
      const f = (item.fields ?? {}) as Record<string, unknown>;
      const siteVal = Number(f[siteLookupIdKey]);
      const cleanerVal = Number(f[cleanerLookupIdKey]);
      if (!Number.isFinite(siteVal) || !Number.isFinite(cleanerVal)) continue;
      if (siteVal !== Number(siteLookupVal) || cleanerVal !== Number(cleanerLookupVal)) continue;

      if (adHocJobKey) {
        const ahRaw = f[adHocJobKey];
        const ahVal = ahRaw == null || ahRaw === "" ? null : Number(ahRaw);
        const want = adHocLookupVal == null ? null : Number(adHocLookupVal);
        if (want == null) {
          if (ahVal != null && Number.isFinite(ahVal)) continue;
        } else {
          if (!(Number.isFinite(ahVal) && ahVal === want)) continue;
        }
      }

      const rawWorkDate = f[workDateKey];
      const d = toDate(rawWorkDate);
      if (!d) continue;
      const ymd =
        typeof rawWorkDate === "string" && /^\d{4}-\d{2}-\d{2}/.test(rawWorkDate)
          ? rawWorkDate.slice(0, 10)
          : toLocalYmd(d);
      if (ymd < startYmd || ymd >= endYmdExclusive) continue;

      if (item.id) toDeleteIds.push(item.id);
    }

    try {
      let batchIdx = 0;
      for (let i = 0; i < toDeleteIds.length; i += 20) {
        const chunk = toDeleteIds.slice(i, i + 20);
        const reqs = chunk.map((id, j) => ({
          id: `del-${batchIdx}-${j}`,
          method: "DELETE" as const,
          url: `/sites/${siteId}/lists/${listId}/items/${id}`,
        }));
        const res = await sharepoint.graphBatch(accessToken, reqs);
        for (const r of res) {
          if (r.status >= 400) {
            errors.push(`[Timesheets Save] Delete failed composite=${composite} status=${r.status}`);
          }
        }
        batchIdx++;
      }
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      errors.push(`[Timesheets Save] Failed to delete existing rows for composite=${composite}: ${msg}`);
    }

    // 2) Create fresh items for hours > 0 using Graph $batch (20 per request).
    const toCreate = group.filter((e) => e.hours > 0);
    try {
      let batchIdx = 0;
      for (let i = 0; i < toCreate.length; i += 20) {
        const chunk = toCreate.slice(i, i + 20);
        const reqs = chunk.map((entry, j) => {
          const workDate = entry.date; // yyyy-MM-dd
          const createFields: Record<string, unknown> = {
            [workDateKey]: workDate,
            [hoursKey]: entry.hours,
            [siteLookupIdKey]: siteLookupVal,
            [cleanerLookupIdKey]: cleanerLookupVal,
          };
          if (adHocJobKey && adHocLookupVal != null) createFields[adHocJobKey] = adHocLookupVal;
          return {
            id: `post-${batchIdx}-${j}`,
            method: "POST" as const,
            url: `/sites/${siteId}/lists/${listId}/items`,
            headers: { "Content-Type": "application/json" },
            body: { fields: createFields },
          };
        });
        const res = await sharepoint.graphBatch(accessToken, reqs);
        for (const r of res) {
          if (r.status >= 400) {
            errors.push(`[Timesheets Save] Create failed composite=${composite} status=${r.status}`);
          } else {
            saved++;
          }
        }
        batchIdx++;
      }
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      errors.push(`[Timesheets Save] Failed to create rows for composite=${composite}: ${msg}`);
    }
  }

  if (errors.length > 0) return { saved, error: errors.join(" | ") };
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
