/**
 * CleanTrack Ad Hoc Jobs – fetch, create, update.
 * List: "CleanTrack Ad Hoc Jobs"
 * Job Name = SharePoint Title. Site = lookup to CleanTrack Sites. Assigned Manager = lookup to CleanTrack Users.
 * All relationship keys use lookup IDs only.
 */

import * as sharepoint from "../lib/sharepoint";
import { getSharePointAccessToken } from "../lib/graph";
import { getCleanTrackUserIdToNameMap } from "./usersRepo";
import type { AdHocJob } from "../types";

const AD_HOC_JOBS_LIST_NAME = "CleanTrack Ad Hoc Jobs";

/** Temporary debug logs. Set to false to disable. */
const DEBUG_ADHOC = true;

/**
 * Display titles the app uses for CleanTrack Ad Hoc Jobs. Wired in getFieldMap (case/spacing-insensitive)
 * so Graph internal names resolve even when SharePoint titles differ slightly.
 */
const AD_HOC_DISPLAY_NAMES_FOR_WIRING: string[] = [
  "Job Name",
  "Job Type",
  "Company Name",
  "Client Name",
  "Site",
  "Assigned Manager",
  "Manual Site Name",
  "Manual Site Address",
  "Manual Site State",
  "Requested By Name",
  "Requested By Email",
  "Request Channel",
  "Requested Date",
  "Scheduled Date",
  "Completed Date",
  "Status",
  "Budgeted Hours",
  "Actual Hours",
  "Service Provider",
  "Charge Rate Per Hour",
  "Cost Rate Per Hour",
  "Charge",
  "Cost",
  "Gross Profit",
  "Markup %",
  "GP %",
  "Description",
  "Approval Method",
  "Approval Proof Required",
  "Approval Proof Uploaded",
  "Approval Reference",
  "Notes for Information",
  "Active",
  "Timesheet Applicable",
  "Recurrence Frequency",
  "Recurrence End Date",
  "Hours Per Service Day",
  "Weekday Hours",
  "Recurrence Weekdays",
  "Monthly Mode",
  "Monthly Day Of Month",
  "Monthly Week Of Month",
  "Monthly Weekday",
  "Monthly Hours",
  "Weekday Charge Rate Override",
  "Saturday Charge Rate Override",
  "Sunday Charge Rate Override",
  "Public Holiday Charge Rate Override",
  "Weekday Cost Rate Override",
  "Saturday Cost Rate Override",
  "Sunday Cost Rate Override",
  "Public Holiday Cost Rate Override",
];

export interface AdHocJobFilters {
  /** Filter by calendar month if any of requested / scheduled / completed / recurrence-end date falls in YYYY-MM. */
  month?: string;
  status?: string;
  assignedManagerId?: string;
  siteId?: string;
}

function toDateStr(v: unknown): string | null {
  if (v == null) return null;
  if (typeof v === "string") {
    const s = v.trim();
    if (!s) return null;
    const d = new Date(s);
    return isNaN(d.getTime()) ? null : d.toISOString().slice(0, 10);
  }
  return null;
}

function toNum(v: unknown): number | null {
  if (v == null || v === "") return null;
  if (typeof v === "number" && !Number.isNaN(v)) return v;
  const n = parseFloat(String(v).replace(/[^0-9.-]/g, ""));
  return Number.isNaN(n) ? null : n;
}

function toNumArray(v: unknown): number[] | null {
  if (v == null || v === "") return null;
  if (Array.isArray(v)) {
    const nums = v.map((x) => (typeof x === "number" ? x : parseInt(String(x), 10))).filter((n) => Number.isFinite(n));
    return nums.length ? nums : null;
  }
  if (typeof v === "string") {
    const s = v.trim();
    if (!s) return null;
    try {
      const parsed = JSON.parse(s);
      if (Array.isArray(parsed)) {
        const nums = parsed.map((x) => (typeof x === "number" ? x : parseInt(String(x), 10))).filter((n) => Number.isFinite(n));
        return nums.length ? nums : null;
      }
    } catch {
      // fall through to comma-split parsing
    }
    const nums = s
      .split(",")
      .map((p) => parseInt(p.trim(), 10))
      .filter((n) => Number.isFinite(n));
    return nums.length ? nums : null;
  }
  return null;
}

function toJsonObject(v: unknown): Record<string, number> | null {
  if (v == null || v === "") return null;
  if (typeof v === "object" && v !== null && !Array.isArray(v)) {
    const out: Record<string, number> = {};
    for (const [k, val] of Object.entries(v as Record<string, unknown>)) {
      const n = typeof val === "number" ? val : parseFloat(String(val));
      if (Number.isFinite(n)) out[k] = n;
    }
    return Object.keys(out).length ? out : null;
  }
  if (typeof v === "string") {
    const s = v.trim();
    if (!s) return null;
    try {
      const parsed = JSON.parse(s);
      return toJsonObject(parsed);
    } catch {
      return null;
    }
  }
  return null;
}

function toBool(v: unknown): boolean {
  if (v === true || v === "Yes") return true;
  if (typeof v === "string") return v.trim().toLowerCase() === "yes";
  return false;
}

/** True if yyyy-MM-dd (or ISO prefix) falls in the given calendar month. */
function isoDateInMonth(iso: string | null | undefined, year: number, month: number): boolean {
  if (!iso) return false;
  const part = String(iso).trim().slice(0, 10);
  const [y, m] = part.split("-").map(Number);
  if (!Number.isFinite(y) || !Number.isFinite(m)) return false;
  return y === year && m === month;
}

/**
 * Month filter: include jobs that touch this month on any primary date (not request-only).
 * Request-only matching hid once-off work scheduled in March when Requested Date was another month.
 */
function adHocJobTouchesMonth(job: AdHocJob, year: number, month: number): boolean {
  return (
    isoDateInMonth(job.requestedDate, year, month) ||
    isoDateInMonth(job.scheduledDate, year, month) ||
    isoDateInMonth(job.completedDate, year, month) ||
    isoDateInMonth(job.recurrenceEndDate, year, month)
  );
}

function getStr(f: Record<string, unknown>, ...keys: string[]): string {
  for (const k of keys) {
    const v = f[k];
    if (v != null && v !== "") return String(v).trim();
  }
  return "";
}

function getLookupId(f: Record<string, unknown>, baseKey: string): string {
  const rawKeys = Object.keys(f);
  const compact = baseKey.replace(/\s+/g, "");
  const candidates = [
    `${baseKey}LookupId`,
    `${baseKey}Id`,
    `${compact}LookupId`,
    `${compact}Id`,
  ];
  const idKey = rawKeys.find(
    (k) => candidates.includes(k) || candidates.some((c) => k.toLowerCase() === c.toLowerCase())
  );
  if (!idKey) return "";
  const v = f[idKey];
  if (v == null) return "";
  if (typeof v === "number") return String(v);
  return String(v).trim();
}

function getLookupValue(f: Record<string, unknown>, baseKey: string): string {
  const rawKeys = Object.keys(f);
  const compact = baseKey.replace(/\s+/g, "");
  const candidates = [baseKey, compact];
  const nameKey = rawKeys.find(
    (k) => candidates.includes(k) || candidates.some((c) => k.toLowerCase() === c.toLowerCase())
  );
  if (!nameKey) return "";
  const v = f[nameKey];
  if (v == null) return "";
  if (typeof v === "object" && v !== null && "LookupValue" in (v as object))
    return String((v as { LookupValue?: string }).LookupValue ?? "").trim();
  return String(v).trim();
}

/** Normalize list column display labels for case/whitespace-insensitive matching. */
function normalizeColumnLabel(displayName: string): string {
  return displayName.trim().toLowerCase().replace(/\s+/g, " ");
}

/**
 * Build display name -> internal name map from list columns.
 * Resolves "Manual Site Name" / "Manual Site Address" even when SharePoint uses different
 * display casing, alternate titles, or compact internal names (e.g. ManualSiteName vs Manual_x0020_...).
 */
async function getFieldMap(
  accessToken: string,
  siteId: string,
  listId: string
): Promise<Record<string, string>> {
  const columns = await sharepoint.getListColumns(accessToken, siteId, listId);
  const map: Record<string, string> = {};
  const byNormDisplay = new Map<string, string>();

  for (const c of columns) {
    if (!c.displayName || !c.name) continue;
    map[c.displayName] = c.name;
    const norm = normalizeColumnLabel(c.displayName);
    if (!byNormDisplay.has(norm)) byNormDisplay.set(norm, c.name);
  }

  const wireCanonical = (canonical: string, ...aliasDisplays: string[]) => {
    if (map[canonical]) return;
    for (const a of aliasDisplays) {
      const internal = byNormDisplay.get(normalizeColumnLabel(a));
      if (internal) {
        map[canonical] = internal;
        return;
      }
    }
  };

  wireCanonical(
    "Manual Site Name",
    "Manual Site Name",
    "Manual site name",
    "Unlisted Site Name",
    "Manual Site",
    "Site Name (Manual)"
  );
  wireCanonical(
    "Manual Site Address",
    "Manual Site Address",
    "Manual site address",
    "Unlisted Site Address",
    "Site Address (Manual)"
  );
  wireCanonical("Manual Site State", "Manual Site State", "Manual site state", "Site State", "State");

  if (!map["Manual Site Name"]) {
    for (const c of columns) {
      if (!c.name) continue;
      const lower = c.name.toLowerCase().replace(/_x0020_/g, "");
      if (
        lower === "manualsitename" ||
        lower === "manual_site_name" ||
        (lower.includes("manual") && lower.includes("site") && lower.includes("name") && !lower.includes("address"))
      ) {
        map["Manual Site Name"] = c.name;
        break;
      }
    }
  }
  if (!map["Manual Site Address"]) {
    for (const c of columns) {
      if (!c.name) continue;
      const lower = c.name.toLowerCase().replace(/_x0020_/g, "");
      if (
        lower === "manualsiteaddress" ||
        lower === "manual_site_address" ||
        (lower.includes("manual") && lower.includes("site") && lower.includes("address"))
      ) {
        map["Manual Site Address"] = c.name;
        break;
      }
    }
  }
  if (!map["Manual Site State"]) {
    for (const c of columns) {
      if (!c.name) continue;
      const lower = c.name.toLowerCase().replace(/_x0020_/g, "");
      if (
        lower === "manualsitestate" ||
        lower === "manual_site_state" ||
        lower === "state" ||
        (lower.includes("manual") && lower.includes("site") && lower.includes("state"))
      ) {
        map["Manual Site State"] = c.name;
        break;
      }
    }
  }

  // List uses "Approval Reference Notes"; app field is approvalReference → "Approval Reference".
  wireCanonical("Approval Reference", "Approval Reference", "Approval Reference Notes");

  for (const label of AD_HOC_DISPLAY_NAMES_FOR_WIRING) {
    wireCanonical(label, label);
  }

  return map;
}

/** Map SharePoint list item to AdHocJob. Resolve manager name from CleanTrack Users when possible. */
function itemToAdHocJob(
  item: sharepoint.GraphListItem,
  map: Record<string, string>,
  managerIdToName: Record<string, string>
): AdHocJob {
  const f = (item.fields ?? {}) as Record<string, unknown>;
  const titleKey = map["Job Name"] ?? "Title";
  const jobName = getStr(f, titleKey, "Title");

  const siteId = getLookupId(f, "Site");
  const siteName = getLookupValue(f, "Site");
  const assignedManagerId = getLookupId(f, "Assigned Manager");
  let assignedManagerName = getLookupValue(f, "Assigned Manager");
  if (!assignedManagerName && assignedManagerId) {
    const norm = sharepoint.normalizeListItemId(assignedManagerId);
    assignedManagerName = managerIdToName[norm] ?? managerIdToName[assignedManagerId] ?? "";
  }

  const jobTypeKey = map["Job Type"] ?? "Job_x0020_Type";
  const companyKey = map["Company Name"] ?? "Company_x0020_Name";
  const clientKey = map["Client Name"] ?? "Client_x0020_Name";
  const manualSiteKey = map["Manual Site Name"];
  const manualSiteAddressKey = map["Manual Site Address"];
  const manualSiteStateKey = map["Manual Site State"] ?? map["State"];
  const reqNameKey = map["Requested By Name"] ?? "Requested_x0020_By_x0020_Name";
  const reqEmailKey = map["Requested By Email"] ?? "Requested_x0020_By_x0020_Email";
  const channelKey = map["Request Channel"] ?? "Request_x0020_Channel";
  const reqDateKey = map["Requested Date"] ?? "Requested_x0020_Date";
  const schedKey = map["Scheduled Date"] ?? "Scheduled_x0020_Date";
  const completedKey = map["Completed Date"] ?? "Completed_x0020_Date";
  const statusKey = map["Status"] ?? "Status";
  const budgetHrsKey = map["Budgeted Hours"] ?? "Budgeted_x0020_Hours";
  const actualHrsKey = map["Actual Hours"] ?? "Actual_x0020_Hours";
  const serviceProviderKey = map["Service Provider"] ?? "Service_x0020_Provider";
  const chargeRateKey = map["Charge Rate Per Hour"] ?? "Charge_x0020_Rate_x0020_Per_x0020_Hour";
  const costRateKey = map["Cost Rate Per Hour"] ?? "Cost_x0020_Rate_x0020_Per_x0020_Hour";
  const chargeKey = map["Charge"] ?? "Charge";
  const costKey = map["Cost"] ?? "Cost";
  const gpKey = map["Gross Profit"] ?? "Gross_x0020_Profit";
  const markupKey = map["Markup %"] ?? "Markup_x0020__x0025_";
  const gpPctKey = map["GP %"] ?? "GP_x0020__x0025_";
  const descKey = map["Description"] ?? "Description";
  const approvalMethodKey = map["Approval Method"];
  const proofReqKey = map["Approval Proof Required"];
  const proofUpKey = map["Approval Proof Uploaded"] ?? "Approval_x0020_Proof_x0020_Uploaded";
  const approvalRefKey = map["Approval Reference"] ?? "Approval_x0020_Reference";
  const notesKey = map["Notes for Information"] ?? "Notes_x0020_for_x0020_Information";
  const activeKey = map["Active"] ?? "Active";
  const timesheetApplicableKey = map["Timesheet Applicable"];
  const recurrenceFreqKey = map["Recurrence Frequency"] ?? "Recurrence_x0020_Frequency";
  const recurrenceEndKey = map["Recurrence End Date"] ?? "Recurrence_x0020_End_x0020_Date";
  const hoursPerDayKey = map["Hours Per Service Day"] ?? "Hours_x0020_Per_x0020_Service_x0020_Day";
  const weekdaysKey = map["Recurrence Weekdays"] ?? "Recurrence_x0020_Weekdays";
  const monthlyModeKey = map["Monthly Mode"] ?? "Monthly_x0020_Mode";
  const monthlyDomKey = map["Monthly Day Of Month"] ?? "Monthly_x0020_Day_x0020_Of_x0020_Month";
  const monthlyWomKey = map["Monthly Week Of Month"] ?? "Monthly_x0020_Week_x0020_Of_x0020_Month";
  const monthlyWdKey = map["Monthly Weekday"] ?? "Monthly_x0020_Weekday";
  const monthlyHoursKey = map["Monthly Hours"] ?? "Monthly_x0020_Hours";
  const weekdayHoursKey = map["Weekday Hours"] ?? "Weekday_x0020_Hours";

  const wdChargeKey = map["Weekday Charge Rate Override"] ?? "Weekday_x0020_Charge_x0020_Rate_x0020_Override";
  const satChargeKey = map["Saturday Charge Rate Override"] ?? "Saturday_x0020_Charge_x0020_Rate_x0020_Override";
  const sunChargeKey = map["Sunday Charge Rate Override"] ?? "Sunday_x0020_Charge_x0020_Rate_x0020_Override";
  const phChargeKey = map["Public Holiday Charge Rate Override"] ?? "Public_x0020_Holiday_x0020_Charge_x0020_Rate_x0020_Override";
  const wdCostKey = map["Weekday Cost Rate Override"] ?? "Weekday_x0020_Cost_x0020_Rate_x0020_Override";
  const satCostKey = map["Saturday Cost Rate Override"] ?? "Saturday_x0020_Cost_x0020_Rate_x0020_Override";
  const sunCostKey = map["Sunday Cost Rate Override"] ?? "Sunday_x0020_Cost_x0020_Rate_x0020_Override";
  const phCostKey = map["Public Holiday Cost Rate Override"] ?? "Public_x0020_Holiday_x0020_Cost_x0020_Rate_x0020_Override";

  return {
    id: sharepoint.normalizeListItemId(item.id),
    jobName,
    jobType: getStr(f, jobTypeKey, "Job Type"),
    companyName: getStr(f, companyKey),
    clientName: getStr(f, clientKey),
    siteId: siteId || null,
    siteName: siteName || "",
    manualSiteName:
      getStr(
        f,
        ...(manualSiteKey ? [manualSiteKey] : []),
        "Manual_x0020_Site_x0020_Name",
        "ManualSiteName"
      ) || undefined,
    manualSiteAddress:
      getStr(
        f,
        ...(manualSiteAddressKey ? [manualSiteAddressKey] : []),
        "Manual_x0020_Site_x0020_Address",
        "ManualSiteAddress"
      ) || undefined,
    manualSiteState:
      getStr(
        f,
        ...(manualSiteStateKey ? [manualSiteStateKey] : []),
        "Manual_x0020_Site_x0020_State",
        "ManualSiteState",
        "State"
      ) || undefined,
    description: getStr(f, descKey),
    requestedByName: getStr(f, reqNameKey),
    requestedByEmail: getStr(f, reqEmailKey),
    requestChannel: getStr(f, channelKey),
    requestedDate: toDateStr(f[reqDateKey] ?? f["RequestedDate"]),
    assignedManagerId: assignedManagerId || null,
    assignedManagerName,
    scheduledDate: toDateStr(f[schedKey] ?? f["ScheduledDate"]),
    completedDate: toDateStr(f[completedKey] ?? f["CompletedDate"]),
    status: getStr(f, statusKey, "Status"),
    budgetedHours: toNum(f[budgetHrsKey] ?? f["BudgetedHours"]),
    actualHours: toNum(f[actualHrsKey] ?? f["ActualHours"]),
    serviceProvider: getStr(f, serviceProviderKey),
    chargeRatePerHour: toNum(f[chargeRateKey]),
    costRatePerHour: toNum(f[costRateKey]),
    charge: toNum(f[chargeKey] ?? f["Charge"]),
    cost: toNum(f[costKey] ?? f["Cost"]),
    grossProfit: toNum(f[gpKey] ?? f["GrossProfit"]),
    markupPercent: toNum(f[markupKey]),
    gpPercent: toNum(f[gpPctKey]),
    approvalMethod: getStr(f, ...(approvalMethodKey ? [approvalMethodKey] : []), "ApprovalMethod") || undefined,
    approvalProofRequired:
      toBool(f[proofReqKey ?? "ApprovalProofRequired"]) ||
      /^required$/i.test(getStr(f, ...(approvalMethodKey ? [approvalMethodKey] : []), "ApprovalMethod")),
    approvalProofUploaded: toBool(f[proofUpKey] ?? f["ApprovalProofUploaded"]),
    approvalReference: getStr(f, approvalRefKey),
    notesForInformation: getStr(f, notesKey),
    active: toBool(f[activeKey] ?? f["Active"]) !== false,
    // Backward-compat: if column is missing, treat as applicable.
    timesheetApplicable:
      timesheetApplicableKey
        ? toBool(f[timesheetApplicableKey] ?? f["TimesheetApplicable"])
        : true,
    recurrenceFrequency: (getStr(f, recurrenceFreqKey) as any) || null,
    recurrenceStartDate: toDateStr(f[schedKey] ?? f["ScheduledDate"]),
    recurrenceEndDate: toDateStr(f[recurrenceEndKey]),
    hoursPerServiceDay: toNum(f[hoursPerDayKey]),
    recurrenceWeekdays: toNumArray(f[weekdaysKey]),
    weekdayHours: toJsonObject(f[weekdayHoursKey]),
    monthlyMode: (getStr(f, monthlyModeKey) as any) || null,
    monthlyDayOfMonth: toNum(f[monthlyDomKey]),
    monthlyWeekOfMonth: (getStr(f, monthlyWomKey) as any) || null,
    monthlyWeekday: toNum(f[monthlyWdKey]),
    monthlyHours: toNum(f[monthlyHoursKey]),

    weekdayChargeRateOverride: toNum(f[wdChargeKey]),
    saturdayChargeRateOverride: toNum(f[satChargeKey]),
    sundayChargeRateOverride: toNum(f[sunChargeKey]),
    publicHolidayChargeRateOverride: toNum(f[phChargeKey]),
    weekdayCostRateOverride: toNum(f[wdCostKey]),
    saturdayCostRateOverride: toNum(f[satCostKey]),
    sundayCostRateOverride: toNum(f[sunCostKey]),
    publicHolidayCostRateOverride: toNum(f[phCostKey]),
  };
}

/**
 * Fetch Ad Hoc Jobs with optional filters. Manager role: pass assignedManagerId to restrict to their jobs.
 */
export async function getAdHocJobs(
  accessToken: string,
  filters?: AdHocJobFilters
): Promise<AdHocJob[]> {
  const siteId = await sharepoint.getSiteId(accessToken);
  const listId = await sharepoint.getListIdByName(accessToken, siteId, AD_HOC_JOBS_LIST_NAME);
  if (!listId) return [];

  const [items, map, managerIdToName] = await Promise.all([
    sharepoint.getListItems(accessToken, siteId, listId),
    getFieldMap(accessToken, siteId, listId),
    getCleanTrackUserIdToNameMap(accessToken),
  ]);

  if (DEBUG_ADHOC && items.length > 0) {
    console.log("[CleanTrack Ad Hoc Jobs] raw record count:", items.length);
    console.log("[CleanTrack Ad Hoc Jobs] sample raw:", JSON.stringify(items[0], null, 2));
    console.log("[CleanTrack Ad Hoc Jobs] field map keys:", Object.keys(map));
  }

  let list: AdHocJob[] = items.map((item) =>
    itemToAdHocJob(item, map, managerIdToName)
  );

  if (DEBUG_ADHOC && list.length > 0) {
    console.log("[CleanTrack Ad Hoc Jobs] mapped sample:", JSON.stringify(list[0], null, 2));
  }

  if (filters?.month) {
    const [y, m] = filters.month.split("-").map(Number);
    list = list.filter((j) => adHocJobTouchesMonth(j, y, m));
    if (DEBUG_ADHOC) console.log("[CleanTrack Ad Hoc Jobs] after month filter:", filters.month, "count:", list.length);
  }
  if (filters?.status) {
    list = list.filter((j) => j.status.toLowerCase() === filters!.status!.toLowerCase());
  }
  if (filters?.assignedManagerId) {
    const norm = sharepoint.normalizeListItemId(filters.assignedManagerId);
    list = list.filter((j) => j.assignedManagerId && sharepoint.normalizeListItemId(j.assignedManagerId) === norm);
  }
  if (filters?.siteId) {
    const norm = sharepoint.normalizeListItemId(filters.siteId);
    list = list.filter((j) => j.siteId && sharepoint.normalizeListItemId(j.siteId) === norm);
  }

  return list;
}

/** Payload for create/update. Use lookup IDs for Site and Assigned Manager. */
export interface AdHocJobPayload {
  jobName: string;
  jobType?: string;
  companyName?: string;
  clientName?: string;
  siteId?: string | null;
  manualSiteName?: string;
  manualSiteAddress?: string;
  manualSiteState?: string;
  assignedManagerId?: string | null;
  requestedByName?: string;
  requestedByEmail?: string;
  requestChannel?: string;
  requestedDate?: string | null;
  scheduledDate?: string | null;
  completedDate?: string | null;
  status?: string;
  budgetedHours?: number | null;
  /** Recurring schedule fields (optional; only used when schedule type = Recurring). */
  recurrenceFrequency?: 'Weekly' | 'Fortnightly' | 'Monthly' | null;
  recurrenceEndDate?: string | null;
  /** Legacy: single hours-per-day (kept for backward compatibility). */
  hoursPerServiceDay?: number | null;
  recurrenceWeekdays?: number[] | null;
  /** Per-weekday hours for weekly/fortnightly. Keys are day indexes 0..6. */
  weekdayHours?: Record<string, number> | null;
  monthlyMode?: 'day_of_month' | 'nth_weekday' | null;
  monthlyDayOfMonth?: number | null;
  monthlyWeekOfMonth?: 'First' | 'Second' | 'Third' | 'Fourth' | 'Last' | null;
  monthlyWeekday?: number | null;
  monthlyHours?: number | null;

  /** Optional day-type rate overrides (fixed rates, not multipliers). */
  weekdayChargeRateOverride?: number | null;
  saturdayChargeRateOverride?: number | null;
  sundayChargeRateOverride?: number | null;
  publicHolidayChargeRateOverride?: number | null;
  weekdayCostRateOverride?: number | null;
  saturdayCostRateOverride?: number | null;
  sundayCostRateOverride?: number | null;
  publicHolidayCostRateOverride?: number | null;
  description?: string;
  actualHours?: number | null;
  serviceProvider?: string;
  chargeRatePerHour?: number | null;
  costRatePerHour?: number | null;
  charge?: number | null;
  cost?: number | null;
  grossProfit?: number | null;
  markupPercent?: number | null;
  gpPercent?: number | null;
  approvalProofRequired?: boolean;
  approvalProofUploaded?: boolean;
  approvalMethod?: string;
  approvalReference?: string;
  notesForInformation?: string;
  active?: boolean;
  timesheetApplicable?: boolean;
}

export interface AdHocAttachment {
  fileName: string;
  url: string;
}

function payloadToFields(
  payload: AdHocJobPayload,
  map: Record<string, string>
): Record<string, unknown> {
  const titleKey = map["Job Name"] ?? "Title";
  const fields: Record<string, unknown> = {};
  if (payload.jobName !== undefined) fields[titleKey === "LinkTitle" ? "Title" : titleKey] = payload.jobName;
  if (payload.jobType !== undefined) fields[map["Job Type"] ?? "Job_x0020_Type"] = payload.jobType ?? "";
  if (payload.companyName !== undefined) {
    const k = map["Company Name"] ?? "Company_x0020_Name";
    fields[k] = payload.companyName ?? "";
  }
  if (payload.clientName !== undefined) {
    const k = map["Client Name"] ?? "Client_x0020_Name";
    fields[k] = payload.clientName ?? "";
  }

  const siteInternal = map["Site"] ?? "Site";
  const siteLookupKey = siteInternal === "Site" ? "SiteLookupId" : `${siteInternal}LookupId`;
  if (payload.siteId !== undefined && payload.siteId != null && payload.siteId !== "") {
    const num = /^\d+$/.test(payload.siteId) ? parseInt(payload.siteId, 10) : payload.siteId;
    fields[siteLookupKey] = num;
  }
  // Manual/unlisted site — only send if the list has a resolved column (wrong internal names 400).
  if (payload.manualSiteName !== undefined && map["Manual Site Name"]) {
    fields[map["Manual Site Name"]] = payload.manualSiteName ?? "";
  }
  if (payload.manualSiteAddress !== undefined && map["Manual Site Address"]) {
    fields[map["Manual Site Address"]] = payload.manualSiteAddress ?? "";
  }
  const manualSiteStateWriteKey = map["Manual Site State"] ?? map["State"];
  if (payload.manualSiteState !== undefined && manualSiteStateWriteKey) {
    fields[manualSiteStateWriteKey] = payload.manualSiteState ?? "";
  }

  const mgrInternal = map["Assigned Manager"] ?? "Assigned_x0020_Manager";
  const mgrLookupKey = mgrInternal === "Assigned Manager" ? "Assigned_x0020_ManagerLookupId" : `${mgrInternal}LookupId`;
  if (payload.assignedManagerId !== undefined && payload.assignedManagerId != null && payload.assignedManagerId !== "") {
    const num = /^\d+$/.test(payload.assignedManagerId) ? parseInt(payload.assignedManagerId, 10) : payload.assignedManagerId;
    fields[mgrLookupKey] = num;
  }

  const set = (displayName: string, internalFallback: string, value: unknown) => {
    const k = map[displayName] ?? internalFallback;
    if (k && value !== undefined) fields[k] = value;
  };
  /** Only write if this list has the column (Graph 400 if internal name is wrong). */
  const setIfMapped = (displayName: string, value: unknown) => {
    const k = map[displayName];
    if (!k || value === undefined) return;
    fields[k] = value;
  };
  /**
   * Partial PATCHes (e.g. only `{ approvalProofUploaded: true }` after file upload) must not
   * send other keys; `undefined ?? null` would otherwise clear dates, money fields, and text in SharePoint.
   */
  if (payload.requestedByName !== undefined) {
    set("Requested By Name", "Requested_x0020_By_x0020_Name", payload.requestedByName ?? "");
  }
  if (payload.requestedByEmail !== undefined) {
    set("Requested By Email", "Requested_x0020_By_x0020_Email", payload.requestedByEmail ?? "");
  }
  if (payload.requestChannel !== undefined) {
    set("Request Channel", "Request_x0020_Channel", payload.requestChannel ?? "");
  }
  if (payload.requestedDate !== undefined) {
    set("Requested Date", "Requested_x0020_Date", payload.requestedDate ?? null);
  }
  if (payload.scheduledDate !== undefined) {
    set("Scheduled Date", "Scheduled_x0020_Date", payload.scheduledDate ?? null);
  }
  if (payload.completedDate !== undefined) {
    set("Completed Date", "Completed_x0020_Date", payload.completedDate ?? null);
  }
  if (payload.status !== undefined) {
    set("Status", "Status", payload.status ?? "Requested");
  }
  if (payload.budgetedHours !== undefined) {
    set("Budgeted Hours", "Budgeted_x0020_Hours", payload.budgetedHours ?? null);
  }
  // Schedule / recurrence / overrides — only if columns exist (this tenant may omit or use different internals).
  if (payload.recurrenceFrequency !== undefined) {
    setIfMapped("Recurrence Frequency", payload.recurrenceFrequency ?? null);
  }
  if (payload.recurrenceEndDate !== undefined) {
    setIfMapped("Recurrence End Date", payload.recurrenceEndDate ?? null);
  }
  if (payload.hoursPerServiceDay !== undefined) {
    setIfMapped("Hours Per Service Day", payload.hoursPerServiceDay ?? null);
  }
  if (payload.weekdayHours !== undefined) {
    const k = map["Weekday Hours"];
    if (k) fields[k] = payload.weekdayHours ? JSON.stringify(payload.weekdayHours) : "";
  }
  if (payload.recurrenceWeekdays !== undefined) {
    const k = map["Recurrence Weekdays"];
    if (k) fields[k] = payload.recurrenceWeekdays ? payload.recurrenceWeekdays.join(",") : "";
  }
  if (payload.monthlyMode !== undefined) {
    setIfMapped("Monthly Mode", payload.monthlyMode ?? null);
  }
  if (payload.monthlyDayOfMonth !== undefined) {
    setIfMapped("Monthly Day Of Month", payload.monthlyDayOfMonth ?? null);
  }
  if (payload.monthlyWeekOfMonth !== undefined) {
    setIfMapped("Monthly Week Of Month", payload.monthlyWeekOfMonth ?? null);
  }
  if (payload.monthlyWeekday !== undefined) {
    setIfMapped("Monthly Weekday", payload.monthlyWeekday ?? null);
  }
  if (payload.monthlyHours !== undefined) {
    setIfMapped("Monthly Hours", payload.monthlyHours ?? null);
  }

  if (payload.weekdayChargeRateOverride !== undefined) {
    setIfMapped("Weekday Charge Rate Override", payload.weekdayChargeRateOverride ?? null);
  }
  if (payload.saturdayChargeRateOverride !== undefined) {
    setIfMapped("Saturday Charge Rate Override", payload.saturdayChargeRateOverride ?? null);
  }
  if (payload.sundayChargeRateOverride !== undefined) {
    setIfMapped("Sunday Charge Rate Override", payload.sundayChargeRateOverride ?? null);
  }
  if (payload.publicHolidayChargeRateOverride !== undefined) {
    setIfMapped("Public Holiday Charge Rate Override", payload.publicHolidayChargeRateOverride ?? null);
  }
  if (payload.weekdayCostRateOverride !== undefined) {
    setIfMapped("Weekday Cost Rate Override", payload.weekdayCostRateOverride ?? null);
  }
  if (payload.saturdayCostRateOverride !== undefined) {
    setIfMapped("Saturday Cost Rate Override", payload.saturdayCostRateOverride ?? null);
  }
  if (payload.sundayCostRateOverride !== undefined) {
    setIfMapped("Sunday Cost Rate Override", payload.sundayCostRateOverride ?? null);
  }
  if (payload.publicHolidayCostRateOverride !== undefined) {
    setIfMapped("Public Holiday Cost Rate Override", payload.publicHolidayCostRateOverride ?? null);
  }
  if (payload.description !== undefined) {
    set("Description", "Description", payload.description ?? "");
  }
  if (payload.actualHours !== undefined) {
    set("Actual Hours", "Actual_x0020_Hours", payload.actualHours ?? null);
  }
  if (payload.serviceProvider !== undefined) {
    set("Service Provider", "Service_x0020_Provider", payload.serviceProvider ?? "");
  }
  if (payload.chargeRatePerHour !== undefined) {
    set("Charge Rate Per Hour", "Charge_x0020_Rate_x0020_Per_x0020_Hour", payload.chargeRatePerHour ?? null);
  }
  if (payload.costRatePerHour !== undefined) {
    set("Cost Rate Per Hour", "Cost_x0020_Rate_x0020_Per_x0020_Hour", payload.costRatePerHour ?? null);
  }
  if (payload.charge !== undefined) {
    set("Charge", "Charge", payload.charge ?? null);
  }
  if (payload.cost !== undefined) {
    set("Cost", "Cost", payload.cost ?? null);
  }
  if (payload.grossProfit !== undefined) {
    set("Gross Profit", "Gross_x0020_Profit", payload.grossProfit ?? null);
  }
  if (payload.markupPercent !== undefined) {
    set("Markup %", "Markup_x0020__x0025_", payload.markupPercent ?? null);
  }
  if (payload.gpPercent !== undefined) {
    set("GP %", "GP_x0020__x0025_", payload.gpPercent ?? null);
  }
  if (payload.approvalProofRequired !== undefined) {
    setIfMapped("Approval Proof Required", payload.approvalProofRequired);
    if (!map["Approval Proof Required"] && map["Approval Method"] && payload.approvalMethod === undefined) {
      fields[map["Approval Method"]] = payload.approvalProofRequired ? "Required" : "Not Required";
    }
  }
  if (payload.approvalMethod !== undefined) {
    setIfMapped("Approval Method", payload.approvalMethod ?? "");
  }
  if (payload.approvalProofUploaded !== undefined) {
    set("Approval Proof Uploaded", "Approval_x0020_Proof_x0020_Uploaded", payload.approvalProofUploaded ?? false);
  }
  // Only send Approval Reference if we have a trusted internal name from SharePoint;
  // avoid hard-coded fallback that may not exist.
  if (map["Approval Reference"] && payload.approvalReference !== undefined) {
    fields[map["Approval Reference"]] = payload.approvalReference ?? "";
  }
  if (payload.notesForInformation !== undefined) {
    set("Notes for Information", "Notes_x0020_for_x0020_Information", payload.notesForInformation ?? "");
  }
  if (payload.active !== undefined) {
    set("Active", "Active", payload.active !== false);
  }
  if (payload.timesheetApplicable !== undefined) {
    setIfMapped("Timesheet Applicable", payload.timesheetApplicable);
  }
  return fields;
}

/** Resolve internal column names for Ad Hoc Jobs list (for create/update). */
async function getAdHocFieldMap(accessToken: string): Promise<Record<string, string>> {
  const siteId = await sharepoint.getSiteId(accessToken);
  const listId = await sharepoint.getListIdByName(accessToken, siteId, AD_HOC_JOBS_LIST_NAME);
  if (!listId) throw new Error(`List "${AD_HOC_JOBS_LIST_NAME}" not found.`);
  return getFieldMap(accessToken, siteId, listId);
}

export async function createAdHocJob(
  accessToken: string,
  payload: AdHocJobPayload
): Promise<AdHocJob> {
  const siteId = await sharepoint.getSiteId(accessToken);
  const listId = await sharepoint.getListIdByName(accessToken, siteId, AD_HOC_JOBS_LIST_NAME);
  if (!listId) throw new Error(`List "${AD_HOC_JOBS_LIST_NAME}" not found.`);

  const map = await getAdHocFieldMap(accessToken);
  const full: AdHocJobPayload = {
    status: "Requested",
    active: true,
    timesheetApplicable: true,
    ...payload,
  };
  const fields = payloadToFields(full, map);

  if (DEBUG_ADHOC) {
    if (full.manualSiteName?.trim() && !map["Manual Site Name"]) {
      console.warn("[CleanTrack Ad Hoc Jobs] No list column resolved for Manual Site Name; value not sent.");
    }
    if (full.manualSiteAddress?.trim() && !map["Manual Site Address"]) {
      console.warn("[CleanTrack Ad Hoc Jobs] No list column resolved for Manual Site Address; value not sent.");
    }
    if (full.manualSiteState?.trim() && !(map["Manual Site State"] || map["State"])) {
      console.warn("[CleanTrack Ad Hoc Jobs] No list column resolved for Manual Site State; value not sent.");
    }
    console.log("[CleanTrack Ad Hoc Jobs] create payload fields:", JSON.stringify(fields, null, 2));
  }

  const created = await sharepoint.createListItem(accessToken, siteId, listId, fields);
  const managerIdToName = await getCleanTrackUserIdToNameMap(accessToken);
  return itemToAdHocJob(created, map, managerIdToName);
}

export async function updateAdHocJob(
  accessToken: string,
  itemId: string,
  payload: Partial<AdHocJobPayload>
): Promise<void> {
  const siteId = await sharepoint.getSiteId(accessToken);
  const listId = await sharepoint.getListIdByName(accessToken, siteId, AD_HOC_JOBS_LIST_NAME);
  if (!listId) throw new Error(`List "${AD_HOC_JOBS_LIST_NAME}" not found.`);

  const map = await getAdHocFieldMap(accessToken);
  const fields = payloadToFields(payload as AdHocJobPayload, map);
  if (Object.keys(fields).length === 0) return;

  if (DEBUG_ADHOC) {
    console.log("[CleanTrack Ad Hoc Jobs] update itemId:", itemId, "fields:", JSON.stringify(fields, null, 2));
  }

  await sharepoint.updateListItem(accessToken, siteId, listId, itemId, fields);
}

/**
 * Upload proof files as attachments to an Ad Hoc Job list item.
 * Uses SharePoint REST (Graph does not support list item attachments).
 * Call after create or when editing; then set Approval Proof Uploaded on the item if needed.
 */
export async function uploadAdHocJobAttachments(
  accessToken: string,
  itemId: string,
  files: File[]
): Promise<void> {
  if (files.length === 0) return;
  const webUrl = await sharepoint.getSiteWebUrl(accessToken);
  const spoToken = await getSharePointAccessToken();
  if (!spoToken) {
    throw new Error(
      "SharePoint attachment upload failed: missing SharePoint API token. Ask an admin to grant Sites.ReadWrite.All for the SharePoint resource."
    );
  }
  const formDigest = await sharepoint.getSharePointFormDigest(webUrl, spoToken);
  const listTitle = AD_HOC_JOBS_LIST_NAME;
  const normId = sharepoint.normalizeListItemId(itemId);
  for (const file of files) {
    const buf = await file.arrayBuffer();
    const original = file.name.trim() || "proof";
    const lastDot = original.lastIndexOf(".");
    const base = lastDot > 0 ? original.slice(0, lastDot) : original;
    const ext = lastDot > 0 ? original.slice(lastDot) : "";
    const uniqueSuffix = Date.now().toString();
    const safeName = `${base}-${uniqueSuffix}${ext}`;
    await sharepoint.addListItemAttachment(
      webUrl,
      listTitle,
      normId,
      safeName,
      buf,
      spoToken,
      formDigest
    );
  }
}

/** Fetch existing attachments for an Ad Hoc Job (for viewing in UI). */
export async function getAdHocJobAttachments(
  accessToken: string,
  itemId: string
): Promise<AdHocAttachment[]> {
  const webUrl = await sharepoint.getSiteWebUrl(accessToken);
  const spoToken = await getSharePointAccessToken();
  if (!spoToken) return [];
  const listTitle = AD_HOC_JOBS_LIST_NAME;
  const normId = sharepoint.normalizeListItemId(itemId);
  const escapeForOData = (s: string) => s.replace(/'/g, "''");
  const safeTitle = escapeForOData(listTitle.trim());
  const url = `${webUrl}/_api/web/lists/GetByTitle('${safeTitle}')/items(${normId})/AttachmentFiles`;
  const res = await fetch(url, {
    method: "GET",
    headers: {
      Authorization: `Bearer ${spoToken}`,
      Accept: "application/json;odata=nometadata",
    },
  });
  const text = await res.text();
  if (!res.ok) {
    console.error("[SP REST] list attachments error", res.status, text);
    return [];
  }
  let json: any;
  try {
    json = text ? JSON.parse(text) : {};
  } catch {
    return [];
  }
  const items: Array<{ FileName?: string; ServerRelativeUrl?: string }> =
    json.value ??
    (json.d && Array.isArray(json.d.results) ? json.d.results : []);
  const siteOrigin = (() => {
    try {
      return new URL(webUrl).origin;
    } catch {
      return webUrl;
    }
  })();
  return items
    .filter((x) => x.FileName && x.ServerRelativeUrl)
    .map((x) => ({
      fileName: x.FileName as string,
      // ServerRelativeUrl already starts with /sites/... ; joining with webUrl duplicates path and 404s.
      url: `${siteOrigin}${x.ServerRelativeUrl}`,
    }));
}

/** Delete a single attachment from an Ad Hoc Job list item by file name. */
export async function deleteAdHocJobAttachment(
  accessToken: string,
  itemId: string,
  fileName: string
): Promise<void> {
  const webUrl = await sharepoint.getSiteWebUrl(accessToken);
  const spoToken = await getSharePointAccessToken();
  if (!spoToken) return;
  const listTitle = AD_HOC_JOBS_LIST_NAME;
  const normId = sharepoint.normalizeListItemId(itemId);
  const escapeForOData = (s: string) => s.replace(/'/g, "''");
  const safeTitle = escapeForOData(listTitle.trim());
  const safeName = escapeForOData(fileName.trim());
  const url = `${webUrl}/_api/web/lists/GetByTitle('${safeTitle}')/items(${normId})/AttachmentFiles/getByFileName('${safeName}')`;
  const res = await fetch(url, {
    method: "DELETE",
    headers: {
      Authorization: `Bearer ${spoToken}`,
    },
  });
  if (!res.ok) {
    const text = await res.text();
    console.error("[SP REST] delete attachment error", res.status, text);
  }
}

/** Delete an Ad Hoc Job list item. Moves item to SharePoint recycle bin. */
export async function deleteAdHocJob(
  accessToken: string,
  itemId: string
): Promise<void> {
  const siteId = await sharepoint.getSiteId(accessToken);
  const listId = await sharepoint.getListIdByName(accessToken, siteId, AD_HOC_JOBS_LIST_NAME);
  if (!listId) throw new Error(`List "${AD_HOC_JOBS_LIST_NAME}" not found.`);
  await sharepoint.deleteListItem(accessToken, siteId, listId, itemId);
}
