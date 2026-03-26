import * as sharepoint from "../lib/sharepoint";

const BUDGETS_LIST_NAME = "CleanTrack Site Budgets";

/**
 * Expected SharePoint column display names (for a tidy list see docs/SHAREPOINT-SITE-BUDGETS-LIST.md):
 * Identity: Budget Name, Site, Active
 * Visit: Visit Frequency (Choice: Weekly|Fortnightly|Monthly), Hours per Visit (Number)
 * Week 1: Monday Hours .. Sunday Hours
 * Week 2: Week 2 Monday Hours .. Week 2 Sunday Hours
 */

export interface BudgetPayload {
  budgetName: string;
  siteListItemId: string;
  sundayHours?: number;
  mondayHours?: number;
  tuesdayHours?: number;
  wednesdayHours?: number;
  thursdayHours?: number;
  fridayHours?: number;
  saturdayHours?: number;
  active?: boolean;
  /** Weekly (default) | Fortnightly | Monthly */
  visitFrequency?: "Weekly" | "Fortnightly" | "Monthly";
  /** Used when Visit Frequency is Fortnightly or Monthly */
  hoursPerVisit?: number;
  /**
   * Optional contract-monthly recurrence (plans hours on specific monthly dates).
   * When unset/null, Monthly behaves as the existing period-cap implementation.
   */
  monthlyMode?: "day_of_month" | "nth_weekday" | null;
  monthlyWeekOfMonth?: "First" | "Second" | "Third" | "Fourth" | "Last" | null;
  monthlyWeekday?: number | null; // 0=Sun..6=Sat
  monthlyDayOfMonth?: number | null; // 1..31
  /** Monthly exception (delta) add-on for contract sites. */
  monthlyExceptionHoursDelta?: number | null;
  monthlyExceptionMode?: "day_of_month" | "nth_weekday" | null;
  monthlyExceptionWeekOfMonth?: "First" | "Second" | "Third" | "Fourth" | "Last" | null;
  monthlyExceptionWeekday?: number | null; // 0=Sun..6=Sat
  monthlyExceptionDayOfMonth?: number | null; // 1..31
  /** Week 2 hours when Visit Frequency is Fortnightly */
  week2SundayHours?: number;
  week2MondayHours?: number;
  week2TuesdayHours?: number;
  week2WednesdayHours?: number;
  week2ThursdayHours?: number;
  week2FridayHours?: number;
  week2SaturdayHours?: number;
  /** Optional fortnight cost budget ($). */
  fortnightCostBudget?: number;
  /** Hourly rates for budgeted labour ($/hr). Mon–Fri use Weekday; Sat/Sun and PH use their own rates. */
  weekdayLabourRate?: number;
  saturdayLabourRate?: number;
  sundayLabourRate?: number;
  phLabourRate?: number;
}

async function getSiteAndListId(accessToken: string): Promise<{ siteId: string; listId: string }> {
  const siteId = await sharepoint.getSiteId(accessToken);
  const listId = await sharepoint.getListIdByName(accessToken, siteId, BUDGETS_LIST_NAME);
  if (!listId) throw new Error(`List "${BUDGETS_LIST_NAME}" not found.`);
  return { siteId, listId };
}

/** Create a Site Budget row (planned hours per weekday). */
export async function createSiteBudget(
  accessToken: string,
  payload: BudgetPayload
): Promise<void> {
  const { siteId, listId } = await getSiteAndListId(accessToken);
  const columns = await sharepoint.getListColumns(accessToken, siteId, listId);
  const map: Record<string, string> = {};
  for (const c of columns) {
    const display = c.displayName?.trim();
    if (display) {
      map[display] = c.name;
      map[display.toLowerCase()] = c.name;
    }
    if (c.name) map[c.name] = c.name;
  }
  const budgetNameKey = map["Budget Name"] ?? "Title";
  const siteKey = map["Site"] ?? "SiteLookupId";
  const monKey = map["Monday Hours"] ?? "MondayHours";
  const tueKey = map["Tuesday Hours"] ?? "TuesdayHours";
  const wedKey = map["Wednesday Hours"] ?? "WednesdayHours";
  const thuKey = map["Thursday Hours"] ?? "ThursdayHours";
  const friKey = map["Friday Hours"] ?? "FridayHours";
  const satKey = map["Saturday Hours"] ?? "SaturdayHours";
  const sunKey = map["Sunday Hours"] ?? "SundayHours";
  const activeKey = map["Active"] ?? "Active";
  const visitFreqKey = map["Visit Frequency"] ?? map["VisitFrequency"] ?? "Visit_x0020_Frequency";
  const hoursPerVisitKey = map["Hours per Visit"] ?? map["HoursPerVisit"] ?? "Hours_x0020_per_x0020_Visit";
  const monthlyModeKey = map["Monthly Mode"] ?? map["Monthly_x0020_Mode"] ?? undefined;
  const monthlyWomKey =
    map["Monthly Week Of Month"] ?? map["Monthly_x0020_Week_x0020_Of_x0020_Month"] ?? undefined;
  const monthlyWdKey = map["Monthly Weekday"] ?? map["Monthly_x0020_Weekday"] ?? undefined;
  const monthlyDomKey =
    map["Monthly Day Of Month"] ?? map["Monthly_x0020_Day_x0020_Of_x0020_Month"] ?? undefined;
  const ci = (s: string) => s.trim().toLowerCase();
  const pickFirst = (labels: string[]) => {
    for (const l of labels) {
      const v = map[l] ?? map[ci(l)];
      if (v) return v;
    }
    return undefined;
  };

  const monthlyExceptionHoursDeltaKey = pickFirst([
    "Monthly Exception Hours Delta",
    "Monthly_x0020_Exception_x0020_Hours_x0020_Delta",
  ]);
  const monthlyExceptionModeKey = pickFirst(["Monthly Exception Mode", "Monthly_x0020_Exception_x0020_Mode"]);
  const monthlyExceptionWomKey = pickFirst([
    "Monthly Exception Week Of Month",
    "Monthly_x0020_Exception_x0020_Week_x0020_Of_x0020_Month",
  ]);
  const monthlyExceptionWdKey = pickFirst(["Monthly Exception Weekday", "Monthly_x0020_Exception_x0020_Weekday"]);
  const monthlyExceptionDomKey = pickFirst([
    "Monthly Exception Day Of Month",
    "Monthly_x0020_Exception_x0020_Day_x0020_Of_x0020_Month",
  ]);
  const w2MonKey = map["Week 2 Monday Hours"] ?? "Week2MondayHours";
  const w2TueKey = map["Week 2 Tuesday Hours"] ?? "Week2TuesdayHours";
  const w2WedKey = map["Week 2 Wednesday Hours"] ?? "Week2WednesdayHours";
  const w2ThuKey = map["Week 2 Thursday Hours"] ?? "Week2ThursdayHours";
  const w2FriKey = map["Week 2 Friday Hours"] ?? "Week2FridayHours";
  const w2SatKey = map["Week 2 Saturday Hours"] ?? "Week2SaturdayHours";
  const w2SunKey = map["Week 2 Sunday Hours"] ?? "Week2SundayHours";
  /** Only write if column exists (Graph rejects unknown field names). */
  const fortnightCostBudgetKey = map["Fortnight Cost Budget"] ?? map["FortnightCostBudget"] ?? undefined;
  const weekdayLabourRateKey = map["Weekday Labour Rate"] ?? map["WeekdayLabourRate"] ?? undefined;
  const saturdayLabourRateKey = map["Saturday Labour Rate"] ?? map["SaturdayLabourRate"] ?? undefined;
  const sundayLabourRateKey = map["Sunday Labour Rate"] ?? map["SundayLabourRate"] ?? undefined;
  const phLabourRateKey = map["PH Labour Rate"] ?? map["PHLabourRate"] ?? map["Public Holiday Labour Rate"] ?? undefined;

  const nameFieldKey = budgetNameKey === "LinkTitle" ? "Title" : budgetNameKey;
  const numSiteId = parseInt(payload.siteListItemId, 10);
  const siteLookupVal = Number.isNaN(numSiteId) ? payload.siteListItemId : numSiteId;

  // Graph requires lookup columns to be set with the "LookupId" suffix (e.g. SiteLookupId), not "Site"
  const siteLookupKey = siteKey === "Site" ? "SiteLookupId" : siteKey;
  const fields: Record<string, unknown> = {
    [nameFieldKey]: payload.budgetName,
    [siteLookupKey]: siteLookupVal,
    [monKey]: payload.mondayHours ?? 0,
    [tueKey]: payload.tuesdayHours ?? 0,
    [wedKey]: payload.wednesdayHours ?? 0,
    [thuKey]: payload.thursdayHours ?? 0,
    [friKey]: payload.fridayHours ?? 0,
    [satKey]: payload.saturdayHours ?? 0,
    [sunKey]: payload.sundayHours ?? 0,
    [activeKey]: payload.active !== false,
  };
  if (payload.visitFrequency !== undefined && payload.visitFrequency !== "") {
    fields[visitFreqKey] = payload.visitFrequency;
  }
  if (payload.hoursPerVisit !== undefined && payload.hoursPerVisit !== null && payload.hoursPerVisit !== "") {
    fields[hoursPerVisitKey] = Number(payload.hoursPerVisit);
  }
  if (monthlyModeKey && payload.monthlyMode !== undefined && payload.monthlyMode !== null) {
    fields[monthlyModeKey] = payload.monthlyMode;
  }
  if (monthlyWomKey && payload.monthlyWeekOfMonth !== undefined && payload.monthlyWeekOfMonth !== null) {
    fields[monthlyWomKey] = payload.monthlyWeekOfMonth;
  }
  if (monthlyWdKey && payload.monthlyWeekday !== undefined && payload.monthlyWeekday !== null) {
    fields[monthlyWdKey] = payload.monthlyWeekday;
  }
  if (monthlyDomKey && payload.monthlyDayOfMonth !== undefined && payload.monthlyDayOfMonth !== null) {
    fields[monthlyDomKey] = payload.monthlyDayOfMonth;
  }
  if (
    monthlyExceptionHoursDeltaKey &&
    payload.monthlyExceptionHoursDelta !== undefined &&
    payload.monthlyExceptionHoursDelta !== null
  ) {
    fields[monthlyExceptionHoursDeltaKey] = payload.monthlyExceptionHoursDelta;
  }
  if (monthlyExceptionModeKey && payload.monthlyExceptionMode !== undefined && payload.monthlyExceptionMode !== null) {
    fields[monthlyExceptionModeKey] = payload.monthlyExceptionMode;
  }
  if (
    monthlyExceptionWomKey &&
    payload.monthlyExceptionWeekOfMonth !== undefined &&
    payload.monthlyExceptionWeekOfMonth !== null
  ) {
    fields[monthlyExceptionWomKey] = payload.monthlyExceptionWeekOfMonth;
  }
  if (
    monthlyExceptionWdKey &&
    payload.monthlyExceptionWeekday !== undefined &&
    payload.monthlyExceptionWeekday !== null
  ) {
    fields[monthlyExceptionWdKey] = payload.monthlyExceptionWeekday;
  }
  if (
    monthlyExceptionDomKey &&
    payload.monthlyExceptionDayOfMonth !== undefined &&
    payload.monthlyExceptionDayOfMonth !== null
  ) {
    fields[monthlyExceptionDomKey] = payload.monthlyExceptionDayOfMonth;
  }
  if (payload.week2SundayHours !== undefined) fields[w2SunKey] = payload.week2SundayHours;
  if (payload.week2MondayHours !== undefined) fields[w2MonKey] = payload.week2MondayHours;
  if (payload.week2TuesdayHours !== undefined) fields[w2TueKey] = payload.week2TuesdayHours;
  if (payload.week2WednesdayHours !== undefined) fields[w2WedKey] = payload.week2WednesdayHours;
  if (payload.week2ThursdayHours !== undefined) fields[w2ThuKey] = payload.week2ThursdayHours;
  if (payload.week2FridayHours !== undefined) fields[w2FriKey] = payload.week2FridayHours;
  if (payload.week2SaturdayHours !== undefined) fields[w2SatKey] = payload.week2SaturdayHours;
  if (fortnightCostBudgetKey && payload.fortnightCostBudget !== undefined && payload.fortnightCostBudget !== null && payload.fortnightCostBudget !== "") {
    const val = Number(payload.fortnightCostBudget);
    if (!Number.isNaN(val) && val >= 0) fields[fortnightCostBudgetKey] = Math.round(val * 100) / 100;
  }
  if (weekdayLabourRateKey && payload.weekdayLabourRate !== undefined && payload.weekdayLabourRate !== null && payload.weekdayLabourRate !== "") {
    const val = Number(payload.weekdayLabourRate);
    if (!Number.isNaN(val) && val >= 0) fields[weekdayLabourRateKey] = Math.round(val * 100) / 100;
  }
  if (saturdayLabourRateKey && payload.saturdayLabourRate !== undefined && payload.saturdayLabourRate !== null && payload.saturdayLabourRate !== "") {
    const val = Number(payload.saturdayLabourRate);
    if (!Number.isNaN(val) && val >= 0) fields[saturdayLabourRateKey] = Math.round(val * 100) / 100;
  }
  if (sundayLabourRateKey && payload.sundayLabourRate !== undefined && payload.sundayLabourRate !== null && payload.sundayLabourRate !== "") {
    const val = Number(payload.sundayLabourRate);
    if (!Number.isNaN(val) && val >= 0) fields[sundayLabourRateKey] = Math.round(val * 100) / 100;
  }
  if (phLabourRateKey && payload.phLabourRate !== undefined && payload.phLabourRate !== null && payload.phLabourRate !== "") {
    const val = Number(payload.phLabourRate);
    if (!Number.isNaN(val) && val >= 0) fields[phLabourRateKey] = Math.round(val * 100) / 100;
  }

  await sharepoint.createListItem(accessToken, siteId, listId, fields);
}

export interface SiteBudgetHours {
  siteListItemId: string;
  budgetListItemId?: string;
  /** Omitted when the column is empty / absent in SharePoint (UI shows blank). */
  sunday?: number;
  monday?: number;
  tuesday?: number;
  wednesday?: number;
  thursday?: number;
  friday?: number;
  saturday?: number;
  weekTotal: number;
  fortnightCap: number;
  /** Weekly | Fortnightly | Monthly */
  visitFrequency?: string;
  hoursPerVisit?: number;
  monthlyMode?: "day_of_month" | "nth_weekday" | null;
  monthlyWeekOfMonth?: "First" | "Second" | "Third" | "Fourth" | "Last" | null;
  monthlyWeekday?: number | null;
  monthlyDayOfMonth?: number | null;
  /** Monthly exception (delta) add-on for contract sites. */
  monthlyExceptionHoursDelta?: number | null;
  monthlyExceptionMode?: "day_of_month" | "nth_weekday" | null;
  monthlyExceptionWeekOfMonth?: "First" | "Second" | "Third" | "Fourth" | "Last" | null;
  monthlyExceptionWeekday?: number | null;
  monthlyExceptionDayOfMonth?: number | null;
  /** Week 2 day hours when Fortnightly */
  week2Sunday?: number;
  week2Monday?: number;
  week2Tuesday?: number;
  week2Wednesday?: number;
  week2Thursday?: number;
  week2Friday?: number;
  week2Saturday?: number;
  /** Optional fortnight cost budget ($). */
  fortnightCostBudget?: number;
  /** Hourly rates for budgeted labour ($/hr). Mon–Fri / Sat / Sun / Public Holiday. */
  weekdayLabourRate?: number;
  saturdayLabourRate?: number;
  sundayLabourRate?: number;
  phLabourRate?: number;
  /** @deprecated Use weekdayLabourRate. Kept for reading old lists. */
  budgetLabourRate?: number;
  /** @deprecated Use saturdayLabourRate/sundayLabourRate. Kept for reading old lists. */
  weekendLabourRate?: number;
}

function toNum(v: unknown): number {
  if (typeof v === "number" && !Number.isNaN(v)) return v;
  if (typeof v === "string") {
    const n = parseFloat(String(v).replace(/[^0-9.-]/g, ""));
    return Number.isNaN(n) ? 0 : n;
  }
  return 0;
}

/** Extract number from SharePoint Currency or Number column (Graph may return number or { Value?, Amount? }). */
function toCurrencyNum(v: unknown): number {
  if (v == null) return 0;
  if (typeof v === "number" && !Number.isNaN(v)) return v;
  if (typeof v === "string") return toNum(v);
  if (typeof v === "object" && v !== null) {
    const o = v as Record<string, unknown>;
    const val = o["Value"] ?? o["value"] ?? o["Amount"] ?? o["amount"];
    if (val != null) return toNum(val);
  }
  return 0;
}

/** Fields that may exist on a “secondary” budget row when the primary row wins on weekTotal. */
const BUDGET_MERGE_OPTIONAL_KEYS: (keyof SiteBudgetHours)[] = [
  "fortnightCostBudget",
  "weekdayLabourRate",
  "saturdayLabourRate",
  "sundayLabourRate",
  "phLabourRate",
  "budgetLabourRate",
  "weekendLabourRate",
  "sunday",
  "monday",
  "tuesday",
  "wednesday",
  "thursday",
  "friday",
  "saturday",
  "week2Sunday",
  "week2Monday",
  "week2Tuesday",
  "week2Wednesday",
  "week2Thursday",
  "week2Friday",
  "week2Saturday",
  "hoursPerVisit",
  "monthlyMode",
  "monthlyWeekOfMonth",
  "monthlyWeekday",
  "monthlyDayOfMonth",
  "monthlyExceptionHoursDelta",
  "monthlyExceptionMode",
  "monthlyExceptionWeekOfMonth",
  "monthlyExceptionWeekday",
  "monthlyExceptionDayOfMonth",
];

/**
 * When multiple budget list items point at the same site, we pick the row with the highest weekTotal for
 * core hour totals — but cost, rates, or per-day values may live on another row. Merge so we don't drop them.
 */
function mergeBudgetRowsForSameSite(
  existing: SiteBudgetHours | undefined,
  incoming: SiteBudgetHours
): SiteBudgetHours {
  if (!existing) return incoming;
  const primary =
    incoming.weekTotal > existing.weekTotal
      ? incoming
      : existing.weekTotal > incoming.weekTotal
        ? existing
        : existing;
  const secondary =
    incoming.weekTotal > existing.weekTotal
      ? existing
      : existing.weekTotal > incoming.weekTotal
        ? incoming
        : incoming;
  let merged: SiteBudgetHours = { ...primary };
  if (!merged.visitFrequency && secondary.visitFrequency) {
    merged = { ...merged, visitFrequency: secondary.visitFrequency };
  }
  for (const key of BUDGET_MERGE_OPTIONAL_KEYS) {
    const pv = merged[key];
    const sv = secondary[key];
    const hasPrimary = pv !== undefined && pv !== null;
    const hasSecondary =
      sv !== undefined && sv !== null && (typeof sv !== "number" || !Number.isNaN(sv));
    if (!hasPrimary && hasSecondary) merged = { ...merged, [key]: sv } as SiteBudgetHours;
  }
  return merged;
}

/** Discover Fortnight Cost Budget when internal column name differs per tenant. */
function findFortnightCostBudgetFromFields(f: Record<string, unknown>): number | undefined {
  for (const key of Object.keys(f)) {
    const lower = key.toLowerCase();
    if (!lower.includes("fortnight")) continue;
    if (!lower.includes("cost")) continue;
    if (!lower.includes("budget")) continue;
    const v = f[key];
    if (v === undefined || v === null || v === "") continue;
    const n = toCurrencyNum(v);
    if (!Number.isNaN(n) && n >= 0) return n;
  }
  return undefined;
}

/** Display name + internal name lookup (same as create/update budget). */
function buildBudgetColumnMap(columns: Array<{ name: string; displayName: string }>): Record<string, string> {
  const map: Record<string, string> = {};
  for (const c of columns) {
    const display = c.displayName?.trim();
    if (display) {
      map[display] = c.name;
      map[display.toLowerCase()] = c.name;
    }
    if (c.name) map[c.name] = c.name;
  }
  return map;
}

/**
 * Read a list item field using SharePoint column display names (resolves real internal Graph field keys).
 * Graph item.fields keys are internal names, which often differ from our hardcoded guesses.
 */
function readBudgetFieldRaw(
  f: Record<string, unknown>,
  colMap: Record<string, string>,
  displayNames: string[]
): unknown {
  for (const dn of displayNames) {
    const key = colMap[dn] ?? colMap[dn.toLowerCase()];
    if (!key) continue;
    if (!Object.prototype.hasOwnProperty.call(f, key)) continue;
    return f[key];
  }
  return undefined;
}

/** Map SharePoint Choice / text values to app monthly mode ids. */
function normalizeMonthlyRecurrenceMode(raw: unknown): "day_of_month" | "nth_weekday" | null {
  if (raw == null) return null;
  const s = String(raw).trim();
  if (!s) return null;
  const compact = s.toLowerCase().replace(/\s+/g, "_").replace(/-/g, "_");
  if (compact === "day_of_month" || compact === "dayofmonth") return "day_of_month";
  if (compact === "nth_weekday" || compact === "nthweekday") return "nth_weekday";
  const lower = s.toLowerCase();
  if (lower.includes("nth") && lower.includes("week")) return "nth_weekday";
  if (lower.includes("day") && lower.includes("month")) return "day_of_month";
  return null;
}

/**
 * Returns display names of expected monthly-exception columns that are missing from the list schema.
 * Use after save to explain why values did not persist (writes are skipped when columns do not exist).
 */
export async function getMissingMonthlyExceptionBudgetColumns(
  accessToken: string,
  exceptionMode?: "day_of_month" | "nth_weekday" | ""
): Promise<string[]> {
  const { siteId, listId } = await getSiteAndListId(accessToken);
  const columns = await sharepoint.getListColumns(accessToken, siteId, listId);
  const colMap = buildBudgetColumnMap(columns);
  const base = ["Monthly Exception Hours Delta", "Monthly Exception Mode"];
  let requiredDisplayNames: string[] = base;
  if (exceptionMode === "day_of_month") {
    requiredDisplayNames = [...base, "Monthly Exception Day Of Month"];
  } else if (exceptionMode === "nth_weekday") {
    requiredDisplayNames = [...base, "Monthly Exception Week Of Month", "Monthly Exception Weekday"];
  } else {
    // Unknown/empty: be conservative and require everything.
    requiredDisplayNames = [
      ...base,
      "Monthly Exception Day Of Month",
      "Monthly Exception Week Of Month",
      "Monthly Exception Weekday",
    ];
  }
  return requiredDisplayNames.filter((dn) => !colMap[dn]);
}

/** Load all site budgets and return hours per site (prefer budget with highest week total when multiple exist).
 * Also includes budgets with empty Site, keyed by "name:{Budget Name}" so they can be matched and updated. */
export async function getSiteBudgets(
  accessToken: string
): Promise<Record<string, SiteBudgetHours>> {
  const { siteId, listId } = await getSiteAndListId(accessToken);
  const columns = await sharepoint.getListColumns(accessToken, siteId, listId);
  const colMap = buildBudgetColumnMap(columns);
  const items = await sharepoint.getListItems(accessToken, siteId, listId);
  const result: Record<string, SiteBudgetHours> = {};
  const activeKey = "Active";

  /** Normalize choice value from SharePoint (e.g. "monthly" -> "Monthly"). */
  function normalizeVisitFreq(raw: string | undefined): string | undefined {
    if (raw == null || String(raw).trim() === "") return undefined;
    const s = String(raw).trim().toLowerCase();
    if (s === "weekly") return "Weekly";
    if (s === "fortnightly") return "Fortnightly";
    if (s === "monthly") return "Monthly";
    return String(raw).trim();
  }

  function normalizeSiteId(sid: unknown): string {
    if (sid == null) return "";
    if (typeof sid === "object" && "LookupId" in (sid as object)) return String((sid as { LookupId?: number }).LookupId ?? "");
    if (typeof sid === "number") return String(sid);
    return String(sid).trim();
  }

  function getBudgetName(f: Record<string, unknown>): string {
    const name = f["Budget Name"] ?? f["Title"] ?? f["BudgetName"] ?? f["LinkTitle"];
    return String(name ?? "").trim();
  }

  const hourKeyVariants: Record<string, string[]> = {
    sun: ["Sunday Hours", "SundayHours", "Sunday_x0020_Hours"],
    mon: ["Monday Hours", "MondayHours", "Monday_x0020_Hours"],
    tue: ["Tuesday Hours", "TuesdayHours", "Tuesday_x0020_Hours"],
    wed: ["Wednesday Hours", "WednesdayHours", "Wednesday_x0020_Hours"],
    thu: ["Thursday Hours", "ThursdayHours", "Thursday_x0020_Hours"],
    fri: ["Friday Hours", "FridayHours", "Friday_x0020_Hours"],
    sat: ["Saturday Hours", "SaturdayHours", "Saturday_x0020_Hours"],
  };
  const week2KeyVariants: Record<string, string[]> = {
    sun: ["Week 2 Sunday Hours", "Week2SundayHours", "Week_x0020_2_x0020_Sunday_x0020_Hours"],
    mon: ["Week 2 Monday Hours", "Week2MondayHours", "Week_x0020_2_x0020_Monday_x0020_Hours"],
    tue: ["Week 2 Tuesday Hours", "Week2TuesdayHours", "Week_x0020_2_x0020_Tuesday_x0020_Hours"],
    wed: ["Week 2 Wednesday Hours", "Week2WednesdayHours", "Week_x0020_2_x0020_Wednesday_x0020_Hours"],
    thu: ["Week 2 Thursday Hours", "Week2ThursdayHours", "Week_x0020_2_x0020_Thursday_x0020_Hours"],
    fri: ["Week 2 Friday Hours", "Week2FridayHours", "Week_x0020_2_x0020_Friday_x0020_Hours"],
    sat: ["Week 2 Saturday Hours", "Week2SaturdayHours", "Week_x0020_2_x0020_Saturday_x0020_Hours"],
  };
  const visitFreqKeys = ["VisitFrequency", "Visit_x0020_Frequency", "Visit Frequency"];
  const hoursPerVisitKeys = ["HoursPerVisit", "Hours_x0020_per_x0020_Visit", "Hours per Visit"];
  const monthlyModeKeys = ["MonthlyMode", "Monthly_x0020_Mode", "Monthly Mode"];
  const monthlyWomKeys = [
    "MonthlyWeekOfMonth",
    "Monthly_x0020_Week_x0020_Of_x0020_Month",
    "Monthly Week Of Month",
  ];
  const monthlyWdKeys = ["MonthlyWeekday", "Monthly_x0020_Weekday", "Monthly Weekday"];
  const monthlyDomKeys = [
    "MonthlyDayOfMonth",
    "Monthly_x0020_Day_x0020_Of_x0020_Month",
    "Monthly Day Of Month",
  ];
  const monthlyExceptionHoursDeltaKeys = [
    "MonthlyExceptionHoursDelta",
    "Monthly_x0020_Exception_x0020_Hours_x0020_Delta",
    "Monthly Exception Hours Delta",
  ];
  const monthlyExceptionModeKeys = [
    "MonthlyExceptionMode",
    "Monthly_x0020_Exception_x0020_Mode",
    "Monthly Exception Mode",
  ];
  const monthlyExceptionWomKeys = [
    "MonthlyExceptionWeekOfMonth",
    "Monthly_x0020_Exception_x0020_Week_x0020_Of_x0020_Month",
    "Monthly Exception Week Of Month",
  ];
  const monthlyExceptionWdKeys = [
    "MonthlyExceptionWeekday",
    "Monthly_x0020_Exception_x0020_Weekday",
    "Monthly Exception Weekday",
  ];
  const monthlyExceptionDomKeys = [
    "MonthlyExceptionDayOfMonth",
    "Monthly_x0020_Exception_x0020_Day_x0020_Of_x0020_Month",
    "Monthly Exception Day Of Month",
  ];
  /** Schema: Weekday Labour Rate (renamed from Budget Labour Rate), Saturday/Sunday/PH Labour Rate, Fortnight Cost Budget — all Currency in SharePoint. */
  const fortnightCostBudgetKeys = [
    "Fortnight Cost Budget",
    "FortnightCostBudget",
    "Fortnight_x0020_Cost_x0020_Budget",
    "Fortnightly Cost Budget",
    "FortnightlyCostBudget",
    "Fortnightly_x0020_Cost_x0020_Budget",
  ];
  // Graph list item fields use INTERNAL column names. Include common encoded-name variants.
  const weekdayLabourRateKeys = [
    "Weekday Labour Rate",
    "WeekdayLabourRate",
    "Weekday_x0020_Labour_x0020_Rate",
    "Budget Labour Rate",
    "BudgetLabourRate",
    "Budget_x0020_Labour_x0020_Rate",
  ];
  const saturdayLabourRateKeys = [
    "Saturday Labour Rate",
    "SaturdayLabourRate",
    "Saturday_x0020_Labour_x0020_Rate",
    "Weekend Labour Rate",
    "WeekendLabourRate",
    "Weekend_x0020_Labour_x0020_Rate",
  ];
  const sundayLabourRateKeys = [
    "Sunday Labour Rate",
    "SundayLabourRate",
    "Sunday_x0020_Labour_x0020_Rate",
    "Weekend Labour Rate",
    "WeekendLabourRate",
    "Weekend_x0020_Labour_x0020_Rate",
  ];
  const phLabourRateKeys = [
    "PH Labour Rate",
    "PHLabourRate",
    "PH_x0020_Labour_x0020_Rate",
    "Public Holiday Labour Rate",
    "PublicHolidayLabourRate",
    "Public_x0020_Holiday_x0020_Labour_x0020_Rate",
  ];
  const budgetLabourRateKeys = ["Budget Labour Rate", "BudgetLabourRate", "Budget_x0020_Labour_x0020_Rate"];
  const weekendLabourRateKeys = ["Weekend Labour Rate", "WeekendLabourRate", "Weekend_x0020_Labour_x0020_Rate"];

  for (const item of items) {
    const f = (item.fields ?? {}) as Record<string, unknown>;
    const activeVal = f[activeKey];
    if (activeVal === false || activeVal === "No" || String(activeVal).toLowerCase() === "no") continue;
    const sid = f["SiteLookupId"] ?? f["SiteId"] ?? f["Site"];
    const siteIdStr = normalizeSiteId(sid);

    /** Only return a number when SharePoint sent a value for that column (null/absent → undefined, 0 → 0). */
    const getHoursOptional = (keys: string[]): number | undefined => {
      for (const k of keys) {
        if (!Object.prototype.hasOwnProperty.call(f, k)) continue;
        const v = f[k];
        if (v === null || v === undefined || v === "") continue;
        return toNum(v);
      }
      return undefined;
    };
    const sun = getHoursOptional(hourKeyVariants.sun);
    const mon = getHoursOptional(hourKeyVariants.mon);
    const tue = getHoursOptional(hourKeyVariants.tue);
    const wed = getHoursOptional(hourKeyVariants.wed);
    const thu = getHoursOptional(hourKeyVariants.thu);
    const fri = getHoursOptional(hourKeyVariants.fri);
    const sat = getHoursOptional(hourKeyVariants.sat);
    const weekTotal =
      (sun ?? 0) + (mon ?? 0) + (tue ?? 0) + (wed ?? 0) + (thu ?? 0) + (fri ?? 0) + (sat ?? 0);

    const sun2 = getHoursOptional(week2KeyVariants.sun);
    const mon2 = getHoursOptional(week2KeyVariants.mon);
    const tue2 = getHoursOptional(week2KeyVariants.tue);
    const wed2 = getHoursOptional(week2KeyVariants.wed);
    const thu2 = getHoursOptional(week2KeyVariants.thu);
    const fri2 = getHoursOptional(week2KeyVariants.fri);
    const sat2 = getHoursOptional(week2KeyVariants.sat);
    const week2Total =
      (sun2 ?? 0) + (mon2 ?? 0) + (tue2 ?? 0) + (wed2 ?? 0) + (thu2 ?? 0) + (fri2 ?? 0) + (sat2 ?? 0);

    const getFirst = (keys: string[]): string | undefined => {
      for (const k of keys) {
        const v = f[k];
        if (v !== undefined && v !== null && String(v).trim() !== "") return String(v).trim();
      }
      return undefined;
    };
    const getFirstNum = (keys: string[]): number | undefined => {
      for (const k of keys) {
        const v = f[k];
        if (v !== undefined && v !== null && v !== "") return toNum(v);
      }
      return undefined;
    };
    /** Currency columns (Fortnight Cost Budget, labour rates) — handle Graph number or { Value } shape. */
    const getFirstCurrencyNum = (keys: string[]): number | undefined => {
      for (const k of keys) {
        const v = f[k];
        if (v !== undefined && v !== null && v !== "") {
          const n = toCurrencyNum(v);
          if (!Number.isNaN(n)) return n;
        }
      }
      return undefined;
    };
    const visitFrequencyRaw = getFirst(visitFreqKeys);
    const visitFrequency = normalizeVisitFreq(visitFrequencyRaw);
    const hoursPerVisit = getFirstNum(hoursPerVisitKeys);

    const monthlyModeMapped = readBudgetFieldRaw(f, colMap, ["Monthly Mode"]);
    const monthlyMode = normalizeMonthlyRecurrenceMode(
      monthlyModeMapped !== undefined ? monthlyModeMapped : getFirst(monthlyModeKeys)
    );

    const monthlyWomMapped = readBudgetFieldRaw(f, colMap, ["Monthly Week Of Month"]);
    const monthlyWomRaw =
      monthlyWomMapped !== undefined ? monthlyWomMapped : getFirst(monthlyWomKeys);
    const monthlyWomStr = monthlyWomRaw != null && String(monthlyWomRaw).trim() !== "" ? String(monthlyWomRaw).trim() : "";
    const monthlyWeekOfMonth = (
      ["First", "Second", "Third", "Fourth", "Last"] as const
    ).find((w) => w.toLowerCase() === monthlyWomStr.toLowerCase());

    const monthlyWeekdayMapped = readBudgetFieldRaw(f, colMap, ["Monthly Weekday"]);
    const monthlyWeekday =
      monthlyWeekdayMapped !== undefined && monthlyWeekdayMapped !== null && monthlyWeekdayMapped !== ""
        ? toNum(monthlyWeekdayMapped)
        : getFirstNum(monthlyWdKeys);

    const monthlyDomMapped = readBudgetFieldRaw(f, colMap, ["Monthly Day Of Month"]);
    const monthlyDomRaw =
      monthlyDomMapped !== undefined ? monthlyDomMapped : undefined;
    const monthlyDayOfMonth =
      monthlyDomRaw !== undefined && monthlyDomRaw !== null && monthlyDomRaw !== ""
        ? toNum(monthlyDomRaw)
        : getFirstNum(monthlyDomKeys);

    const monthlyExcDeltaRaw = readBudgetFieldRaw(f, colMap, ["Monthly Exception Hours Delta"]);
    const monthlyExceptionHoursDelta =
      monthlyExcDeltaRaw !== undefined && monthlyExcDeltaRaw !== null && monthlyExcDeltaRaw !== ""
        ? toNum(monthlyExcDeltaRaw)
        : getFirstNum(monthlyExceptionHoursDeltaKeys);

    const monthlyExcModeMapped = readBudgetFieldRaw(f, colMap, ["Monthly Exception Mode"]);
    const monthlyExceptionMode = normalizeMonthlyRecurrenceMode(
      monthlyExcModeMapped !== undefined ? monthlyExcModeMapped : getFirst(monthlyExceptionModeKeys)
    );

    const monthlyExcWomMapped = readBudgetFieldRaw(f, colMap, ["Monthly Exception Week Of Month"]);
    const monthlyExcWomRaw =
      monthlyExcWomMapped !== undefined ? monthlyExcWomMapped : getFirst(monthlyExceptionWomKeys);
    const monthlyExcWomStr =
      monthlyExcWomRaw != null && String(monthlyExcWomRaw).trim() !== "" ? String(monthlyExcWomRaw).trim() : "";
    const monthlyExceptionWeekOfMonth = (
      ["First", "Second", "Third", "Fourth", "Last"] as const
    ).find((w) => w.toLowerCase() === monthlyExcWomStr.toLowerCase());

    const monthlyExcWdRaw = readBudgetFieldRaw(f, colMap, ["Monthly Exception Weekday"]);
    const monthlyExceptionWeekday =
      monthlyExcWdRaw !== undefined && monthlyExcWdRaw !== null && monthlyExcWdRaw !== ""
        ? toNum(monthlyExcWdRaw)
        : getFirstNum(monthlyExceptionWdKeys);

    const monthlyExcDomRaw = readBudgetFieldRaw(f, colMap, ["Monthly Exception Day Of Month"]);
    const monthlyExceptionDayOfMonth =
      monthlyExcDomRaw !== undefined && monthlyExcDomRaw !== null && monthlyExcDomRaw !== ""
        ? toNum(monthlyExcDomRaw)
        : getFirstNum(monthlyExceptionDomKeys);
    const fortnightCostBudget =
      getFirstCurrencyNum(fortnightCostBudgetKeys) ??
      getFirstNum(fortnightCostBudgetKeys) ??
      findFortnightCostBudgetFromFields(f);
    const weekdayLabourRate = getFirstCurrencyNum(weekdayLabourRateKeys) ?? getFirstCurrencyNum(budgetLabourRateKeys);
    const saturdayLabourRate = getFirstCurrencyNum(saturdayLabourRateKeys) ?? getFirstCurrencyNum(weekendLabourRateKeys);
    const sundayLabourRate = getFirstCurrencyNum(sundayLabourRateKeys) ?? getFirstCurrencyNum(weekendLabourRateKeys);
    const phLabourRate = getFirstCurrencyNum(phLabourRateKeys);
    const budgetLabourRate = getFirstCurrencyNum(budgetLabourRateKeys);
    const weekendLabourRate = getFirstCurrencyNum(weekendLabourRateKeys);

    let fortnightCap: number;
    if (visitFrequency === "Fortnightly" && (week2Total > 0 || weekTotal > 0)) {
      fortnightCap = weekTotal + week2Total;
    } else if (visitFrequency === "Fortnightly" && hoursPerVisit != null && hoursPerVisit > 0) {
      fortnightCap = hoursPerVisit;
    } else if (visitFrequency === "Monthly" && hoursPerVisit != null && hoursPerVisit > 0) {
      fortnightCap = hoursPerVisit / 2;
    } else {
      fortnightCap = weekTotal * 2;
    }

    const entry: SiteBudgetHours = {
      siteListItemId: siteIdStr,
      budgetListItemId: item.id,
      ...(sun !== undefined && { sunday: sun }),
      ...(mon !== undefined && { monday: mon }),
      ...(tue !== undefined && { tuesday: tue }),
      ...(wed !== undefined && { wednesday: wed }),
      ...(thu !== undefined && { thursday: thu }),
      ...(fri !== undefined && { friday: fri }),
      ...(sat !== undefined && { saturday: sat }),
      weekTotal,
      fortnightCap,
      ...(visitFrequency && { visitFrequency }),
      ...(hoursPerVisit != null && { hoursPerVisit }),
      ...(monthlyMode && { monthlyMode }),
      ...(monthlyWeekOfMonth && { monthlyWeekOfMonth }),
      ...(monthlyWeekday != null && { monthlyWeekday }),
      ...(monthlyDayOfMonth != null && { monthlyDayOfMonth }),
      ...(monthlyExceptionHoursDelta != null && { monthlyExceptionHoursDelta }),
      ...(monthlyExceptionMode && { monthlyExceptionMode }),
      ...(monthlyExceptionWeekOfMonth && { monthlyExceptionWeekOfMonth }),
      ...(monthlyExceptionWeekday != null && { monthlyExceptionWeekday }),
      ...(monthlyExceptionDayOfMonth != null && { monthlyExceptionDayOfMonth }),
      ...(fortnightCostBudget != null && fortnightCostBudget >= 0 && { fortnightCostBudget }),
      ...(weekdayLabourRate != null && weekdayLabourRate >= 0 && { weekdayLabourRate }),
      ...(saturdayLabourRate != null && saturdayLabourRate >= 0 && { saturdayLabourRate }),
      ...(sundayLabourRate != null && sundayLabourRate >= 0 && { sundayLabourRate }),
      ...(phLabourRate != null && phLabourRate >= 0 && { phLabourRate }),
      ...(budgetLabourRate != null && budgetLabourRate >= 0 && { budgetLabourRate }),
      ...(weekendLabourRate != null && weekendLabourRate >= 0 && { weekendLabourRate }),
      ...(sun2 !== undefined && { week2Sunday: sun2 }),
      ...(mon2 !== undefined && { week2Monday: mon2 }),
      ...(tue2 !== undefined && { week2Tuesday: tue2 }),
      ...(wed2 !== undefined && { week2Wednesday: wed2 }),
      ...(thu2 !== undefined && { week2Thursday: thu2 }),
      ...(fri2 !== undefined && { week2Friday: fri2 }),
      ...(sat2 !== undefined && { week2Saturday: sat2 }),
    };

    if (siteIdStr) {
      result[siteIdStr] = mergeBudgetRowsForSameSite(result[siteIdStr], entry);
    }
    const budgetName = getBudgetName(f);
    if (budgetName) {
      const nameKey = "name:" + budgetName;
      const namedEntry = { ...entry, siteListItemId: siteIdStr || "" };
      result[nameKey] = mergeBudgetRowsForSameSite(result[nameKey], namedEntry);
    }
  }
  return result;
}

export interface UpdateBudgetPayload {
  sundayHours?: number;
  mondayHours?: number;
  tuesdayHours?: number;
  wednesdayHours?: number;
  thursdayHours?: number;
  fridayHours?: number;
  saturdayHours?: number;
  active?: boolean;
  /** When set, updates the budget's Site lookup so it is found by site id on next load. */
  siteListItemId?: string;
  /** Weekly | Fortnightly | Monthly */
  visitFrequency?: "Weekly" | "Fortnightly" | "Monthly";
  hoursPerVisit?: number;
  monthlyMode?: "day_of_month" | "nth_weekday" | null;
  monthlyWeekOfMonth?: "First" | "Second" | "Third" | "Fourth" | "Last" | null;
  monthlyWeekday?: number | null;
  monthlyDayOfMonth?: number | null;
  /** Monthly exception (delta) add-on for contract sites. */
  monthlyExceptionHoursDelta?: number | null;
  monthlyExceptionMode?: "day_of_month" | "nth_weekday" | null;
  monthlyExceptionWeekOfMonth?: "First" | "Second" | "Third" | "Fourth" | "Last" | null;
  monthlyExceptionWeekday?: number | null;
  monthlyExceptionDayOfMonth?: number | null;
  week2SundayHours?: number;
  week2MondayHours?: number;
  week2TuesdayHours?: number;
  week2WednesdayHours?: number;
  week2ThursdayHours?: number;
  week2FridayHours?: number;
  week2SaturdayHours?: number;
  fortnightCostBudget?: number;
  weekdayLabourRate?: number;
  saturdayLabourRate?: number;
  sundayLabourRate?: number;
  phLabourRate?: number;
}

/** Update an existing site budget by its list item id. */
export async function updateSiteBudget(
  accessToken: string,
  budgetListItemId: string,
  payload: UpdateBudgetPayload
): Promise<void> {
  const { siteId, listId } = await getSiteAndListId(accessToken);
  const columns = await sharepoint.getListColumns(accessToken, siteId, listId);
  const map: Record<string, string> = {};
  for (const c of columns) {
    const display = c.displayName?.trim();
    if (display) {
      map[display] = c.name;
      map[display.toLowerCase()] = c.name;
    }
    if (c.name) map[c.name] = c.name;
  }
  const monKey = map["Monday Hours"] ?? "MondayHours";
  const tueKey = map["Tuesday Hours"] ?? "TuesdayHours";
  const wedKey = map["Wednesday Hours"] ?? "WednesdayHours";
  const thuKey = map["Thursday Hours"] ?? "ThursdayHours";
  const friKey = map["Friday Hours"] ?? "FridayHours";
  const satKey = map["Saturday Hours"] ?? "SaturdayHours";
  const sunKey = map["Sunday Hours"] ?? "SundayHours";
  const activeKey = map["Active"] ?? "Active";
  const visitFreqKey = map["Visit Frequency"] ?? map["VisitFrequency"] ?? "Visit_x0020_Frequency";
  const hoursPerVisitKey = map["Hours per Visit"] ?? map["HoursPerVisit"] ?? "Hours_x0020_per_x0020_Visit";

  const monthlyModeKey = map["Monthly Mode"] ?? map["Monthly_x0020_Mode"] ?? undefined;
  const monthlyWomKey =
    map["Monthly Week Of Month"] ?? map["Monthly_x0020_Week_x0020_Of_x0020_Month"] ?? undefined;
  const monthlyWdKey = map["Monthly Weekday"] ?? map["Monthly_x0020_Weekday"] ?? undefined;
  const monthlyDomKey =
    map["Monthly Day Of Month"] ?? map["Monthly_x0020_Day_x0020_Of_x0020_Month"] ?? undefined;

  const ci = (s: string) => s.trim().toLowerCase();
  const pickFirst = (labels: string[]) => {
    for (const l of labels) {
      const v = map[l] ?? map[ci(l)];
      if (v) return v;
    }
    return undefined;
  };

  const monthlyExceptionHoursDeltaKey = pickFirst([
    "Monthly Exception Hours Delta",
    "Monthly_x0020_Exception_x0020_Hours_x0020_Delta",
  ]);
  const monthlyExceptionModeKey = pickFirst(["Monthly Exception Mode", "Monthly_x0020_Exception_x0020_Mode"]);
  const monthlyExceptionWomKey = pickFirst([
    "Monthly Exception Week Of Month",
    "Monthly_x0020_Exception_x0020_Week_x0020_Of_x0020_Month",
  ]);
  const monthlyExceptionWdKey = pickFirst(["Monthly Exception Weekday", "Monthly_x0020_Exception_x0020_Weekday"]);
  const monthlyExceptionDomKey = pickFirst([
    "Monthly Exception Day Of Month",
    "Monthly_x0020_Exception_x0020_Day_x0020_Of_x0020_Month",
  ]);

  const w2MonKey = map["Week 2 Monday Hours"] ?? "Week2MondayHours";
  const w2TueKey = map["Week 2 Tuesday Hours"] ?? "Week2TuesdayHours";
  const w2WedKey = map["Week 2 Wednesday Hours"] ?? "Week2WednesdayHours";
  const w2ThuKey = map["Week 2 Thursday Hours"] ?? "Week2ThursdayHours";
  const w2FriKey = map["Week 2 Friday Hours"] ?? "Week2FridayHours";
  const w2SatKey = map["Week 2 Saturday Hours"] ?? "Week2SaturdayHours";
  const w2SunKey = map["Week 2 Sunday Hours"] ?? "Week2SundayHours";
  /** Only write if column exists (Graph rejects unknown field names). */
  const fortnightCostBudgetKey = map["Fortnight Cost Budget"] ?? map["FortnightCostBudget"] ?? undefined;
  const weekdayLabourRateKey = map["Weekday Labour Rate"] ?? map["WeekdayLabourRate"] ?? undefined;
  const saturdayLabourRateKey = map["Saturday Labour Rate"] ?? map["SaturdayLabourRate"] ?? undefined;
  const sundayLabourRateKey = map["Sunday Labour Rate"] ?? map["SundayLabourRate"] ?? undefined;
  const phLabourRateKey = map["PH Labour Rate"] ?? map["PHLabourRate"] ?? map["Public Holiday Labour Rate"] ?? undefined;

  const fields: Record<string, unknown> = {
    [monKey]: payload.mondayHours ?? 0,
    [tueKey]: payload.tuesdayHours ?? 0,
    [wedKey]: payload.wednesdayHours ?? 0,
    [thuKey]: payload.thursdayHours ?? 0,
    [friKey]: payload.fridayHours ?? 0,
    [satKey]: payload.saturdayHours ?? 0,
    [sunKey]: payload.sundayHours ?? 0,
    [activeKey]: payload.active !== false,
  };
  if (payload.siteListItemId != null && payload.siteListItemId !== "") {
    const num = parseInt(payload.siteListItemId, 10);
    fields["SiteLookupId"] = Number.isNaN(num) ? payload.siteListItemId : num;
  }
  if (payload.visitFrequency !== undefined && payload.visitFrequency !== "") {
    fields[visitFreqKey] = payload.visitFrequency;
  }
  if (payload.hoursPerVisit !== undefined && payload.hoursPerVisit !== null) {
    const val = payload.hoursPerVisit;
    const num = typeof val === "number" ? val : val === "" ? 0 : Number(val);
    if (!Number.isNaN(num)) fields[hoursPerVisitKey] = num;
  }
  if (monthlyModeKey && payload.monthlyMode !== undefined && payload.monthlyMode !== null) {
    fields[monthlyModeKey] = payload.monthlyMode;
  }
  if (monthlyWomKey && payload.monthlyWeekOfMonth !== undefined && payload.monthlyWeekOfMonth !== null) {
    fields[monthlyWomKey] = payload.monthlyWeekOfMonth;
  }
  if (monthlyWdKey && payload.monthlyWeekday !== undefined && payload.monthlyWeekday !== null) {
    fields[monthlyWdKey] = payload.monthlyWeekday;
  }
  if (monthlyDomKey && payload.monthlyDayOfMonth !== undefined && payload.monthlyDayOfMonth !== null) {
    fields[monthlyDomKey] = payload.monthlyDayOfMonth;
  }
  if (
    monthlyExceptionHoursDeltaKey &&
    payload.monthlyExceptionHoursDelta !== undefined &&
    payload.monthlyExceptionHoursDelta !== null
  ) {
    fields[monthlyExceptionHoursDeltaKey] = payload.monthlyExceptionHoursDelta;
  }
  if (
    monthlyExceptionModeKey &&
    payload.monthlyExceptionMode !== undefined &&
    payload.monthlyExceptionMode !== null
  ) {
    fields[monthlyExceptionModeKey] = payload.monthlyExceptionMode;
  }
  if (
    monthlyExceptionWomKey &&
    payload.monthlyExceptionWeekOfMonth !== undefined &&
    payload.monthlyExceptionWeekOfMonth !== null
  ) {
    fields[monthlyExceptionWomKey] = payload.monthlyExceptionWeekOfMonth;
  }
  if (
    monthlyExceptionWdKey &&
    payload.monthlyExceptionWeekday !== undefined &&
    payload.monthlyExceptionWeekday !== null
  ) {
    fields[monthlyExceptionWdKey] = payload.monthlyExceptionWeekday;
  }
  if (
    monthlyExceptionDomKey &&
    payload.monthlyExceptionDayOfMonth !== undefined &&
    payload.monthlyExceptionDayOfMonth !== null
  ) {
    fields[monthlyExceptionDomKey] = payload.monthlyExceptionDayOfMonth;
  }
  if (payload.week2SundayHours !== undefined) fields[w2SunKey] = payload.week2SundayHours;
  if (payload.week2MondayHours !== undefined) fields[w2MonKey] = payload.week2MondayHours;
  if (payload.week2TuesdayHours !== undefined) fields[w2TueKey] = payload.week2TuesdayHours;
  if (payload.week2WednesdayHours !== undefined) fields[w2WedKey] = payload.week2WednesdayHours;
  if (payload.week2ThursdayHours !== undefined) fields[w2ThuKey] = payload.week2ThursdayHours;
  if (payload.week2FridayHours !== undefined) fields[w2FriKey] = payload.week2FridayHours;
  if (payload.week2SaturdayHours !== undefined) fields[w2SatKey] = payload.week2SaturdayHours;
  if (fortnightCostBudgetKey && payload.fortnightCostBudget !== undefined && payload.fortnightCostBudget !== null && payload.fortnightCostBudget !== "") {
    const val = typeof payload.fortnightCostBudget === "number" ? payload.fortnightCostBudget : Number(payload.fortnightCostBudget);
    if (!Number.isNaN(val) && val >= 0) fields[fortnightCostBudgetKey] = Math.round(val * 100) / 100;
  }
  if (weekdayLabourRateKey && payload.weekdayLabourRate !== undefined && payload.weekdayLabourRate !== null && payload.weekdayLabourRate !== "") {
    const val = typeof payload.weekdayLabourRate === "number" ? payload.weekdayLabourRate : Number(payload.weekdayLabourRate);
    if (!Number.isNaN(val) && val >= 0) fields[weekdayLabourRateKey] = Math.round(val * 100) / 100;
  }
  if (saturdayLabourRateKey && payload.saturdayLabourRate !== undefined && payload.saturdayLabourRate !== null && payload.saturdayLabourRate !== "") {
    const val = typeof payload.saturdayLabourRate === "number" ? payload.saturdayLabourRate : Number(payload.saturdayLabourRate);
    if (!Number.isNaN(val) && val >= 0) fields[saturdayLabourRateKey] = Math.round(val * 100) / 100;
  }
  if (sundayLabourRateKey && payload.sundayLabourRate !== undefined && payload.sundayLabourRate !== null && payload.sundayLabourRate !== "") {
    const val = typeof payload.sundayLabourRate === "number" ? payload.sundayLabourRate : Number(payload.sundayLabourRate);
    if (!Number.isNaN(val) && val >= 0) fields[sundayLabourRateKey] = Math.round(val * 100) / 100;
  }
  if (phLabourRateKey && payload.phLabourRate !== undefined && payload.phLabourRate !== null && payload.phLabourRate !== "") {
    const val = typeof payload.phLabourRate === "number" ? payload.phLabourRate : Number(payload.phLabourRate);
    if (!Number.isNaN(val) && val >= 0) fields[phLabourRateKey] = Math.round(val * 100) / 100;
  }

  await sharepoint.updateListItem(accessToken, siteId, listId, budgetListItemId, fields);
}
