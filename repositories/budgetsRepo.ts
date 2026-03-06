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
  /** Week 2 hours when Visit Frequency is Fortnightly */
  week2SundayHours?: number;
  week2MondayHours?: number;
  week2TuesdayHours?: number;
  week2WednesdayHours?: number;
  week2ThursdayHours?: number;
  week2FridayHours?: number;
  week2SaturdayHours?: number;
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
    if (c.displayName) map[c.displayName] = c.name;
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
  const w2MonKey = map["Week 2 Monday Hours"] ?? "Week2MondayHours";
  const w2TueKey = map["Week 2 Tuesday Hours"] ?? "Week2TuesdayHours";
  const w2WedKey = map["Week 2 Wednesday Hours"] ?? "Week2WednesdayHours";
  const w2ThuKey = map["Week 2 Thursday Hours"] ?? "Week2ThursdayHours";
  const w2FriKey = map["Week 2 Friday Hours"] ?? "Week2FridayHours";
  const w2SatKey = map["Week 2 Saturday Hours"] ?? "Week2SaturdayHours";
  const w2SunKey = map["Week 2 Sunday Hours"] ?? "Week2SundayHours";

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
  if (payload.week2SundayHours !== undefined) fields[w2SunKey] = payload.week2SundayHours;
  if (payload.week2MondayHours !== undefined) fields[w2MonKey] = payload.week2MondayHours;
  if (payload.week2TuesdayHours !== undefined) fields[w2TueKey] = payload.week2TuesdayHours;
  if (payload.week2WednesdayHours !== undefined) fields[w2WedKey] = payload.week2WednesdayHours;
  if (payload.week2ThursdayHours !== undefined) fields[w2ThuKey] = payload.week2ThursdayHours;
  if (payload.week2FridayHours !== undefined) fields[w2FriKey] = payload.week2FridayHours;
  if (payload.week2SaturdayHours !== undefined) fields[w2SatKey] = payload.week2SaturdayHours;

  await sharepoint.createListItem(accessToken, siteId, listId, fields);
}

export interface SiteBudgetHours {
  siteListItemId: string;
  budgetListItemId?: string;
  sunday: number;
  monday: number;
  tuesday: number;
  wednesday: number;
  thursday: number;
  friday: number;
  saturday: number;
  weekTotal: number;
  fortnightCap: number;
  /** Weekly | Fortnightly | Monthly */
  visitFrequency?: string;
  hoursPerVisit?: number;
  /** Week 2 day hours when Fortnightly */
  week2Sunday?: number;
  week2Monday?: number;
  week2Tuesday?: number;
  week2Wednesday?: number;
  week2Thursday?: number;
  week2Friday?: number;
  week2Saturday?: number;
}

function toNum(v: unknown): number {
  if (typeof v === "number" && !Number.isNaN(v)) return v;
  if (typeof v === "string") {
    const n = parseFloat(String(v).replace(/[^0-9.-]/g, ""));
    return Number.isNaN(n) ? 0 : n;
  }
  return 0;
}

/** Load all site budgets and return hours per site (prefer budget with highest week total when multiple exist).
 * Also includes budgets with empty Site, keyed by "name:{Budget Name}" so they can be matched and updated. */
export async function getSiteBudgets(
  accessToken: string
): Promise<Record<string, SiteBudgetHours>> {
  const { siteId, listId } = await getSiteAndListId(accessToken);
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

  for (const item of items) {
    const f = (item.fields ?? {}) as Record<string, unknown>;
    const activeVal = f[activeKey];
    if (activeVal === false || activeVal === "No" || String(activeVal).toLowerCase() === "no") continue;
    const sid = f["SiteLookupId"] ?? f["SiteId"] ?? f["Site"];
    const siteIdStr = normalizeSiteId(sid);

    const getHours = (keys: string[]) => {
      for (const k of keys) {
        const v = f[k];
        if (v !== undefined && v !== null) return toNum(v);
      }
      return 0;
    };
    const sun = getHours(hourKeyVariants.sun);
    const mon = getHours(hourKeyVariants.mon);
    const tue = getHours(hourKeyVariants.tue);
    const wed = getHours(hourKeyVariants.wed);
    const thu = getHours(hourKeyVariants.thu);
    const fri = getHours(hourKeyVariants.fri);
    const sat = getHours(hourKeyVariants.sat);
    const weekTotal = sun + mon + tue + wed + thu + fri + sat;

    const sun2 = getHours(week2KeyVariants.sun);
    const mon2 = getHours(week2KeyVariants.mon);
    const tue2 = getHours(week2KeyVariants.tue);
    const wed2 = getHours(week2KeyVariants.wed);
    const thu2 = getHours(week2KeyVariants.thu);
    const fri2 = getHours(week2KeyVariants.fri);
    const sat2 = getHours(week2KeyVariants.sat);
    const week2Total = sun2 + mon2 + tue2 + wed2 + thu2 + fri2 + sat2;

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
    const visitFrequencyRaw = getFirst(visitFreqKeys);
    const visitFrequency = normalizeVisitFreq(visitFrequencyRaw);
    const hoursPerVisit = getFirstNum(hoursPerVisitKeys);

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
      sunday: sun,
      monday: mon,
      tuesday: tue,
      wednesday: wed,
      thursday: thu,
      friday: fri,
      saturday: sat,
      weekTotal,
      fortnightCap,
      ...(visitFrequency && { visitFrequency }),
      ...(hoursPerVisit != null && { hoursPerVisit }),
      ...(week2Total > 0 || sun2 || mon2 || tue2 || wed2 || thu2 || fri2 || sat2
        ? {
            week2Sunday: sun2,
            week2Monday: mon2,
            week2Tuesday: tue2,
            week2Wednesday: wed2,
            week2Thursday: thu2,
            week2Friday: fri2,
            week2Saturday: sat2,
          }
        : {}),
    };

    if (siteIdStr) {
      const existing = result[siteIdStr];
      if (!existing || existing.weekTotal < weekTotal) result[siteIdStr] = entry;
    }
    const budgetName = getBudgetName(f);
    if (budgetName) {
      const nameKey = "name:" + budgetName;
      const existing = result[nameKey];
      if (!existing || existing.weekTotal < weekTotal) result[nameKey] = { ...entry, siteListItemId: siteIdStr || "" };
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
  week2SundayHours?: number;
  week2MondayHours?: number;
  week2TuesdayHours?: number;
  week2WednesdayHours?: number;
  week2ThursdayHours?: number;
  week2FridayHours?: number;
  week2SaturdayHours?: number;
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
    if (c.displayName) map[c.displayName] = c.name;
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
  const w2MonKey = map["Week 2 Monday Hours"] ?? "Week2MondayHours";
  const w2TueKey = map["Week 2 Tuesday Hours"] ?? "Week2TuesdayHours";
  const w2WedKey = map["Week 2 Wednesday Hours"] ?? "Week2WednesdayHours";
  const w2ThuKey = map["Week 2 Thursday Hours"] ?? "Week2ThursdayHours";
  const w2FriKey = map["Week 2 Friday Hours"] ?? "Week2FridayHours";
  const w2SatKey = map["Week 2 Saturday Hours"] ?? "Week2SaturdayHours";
  const w2SunKey = map["Week 2 Sunday Hours"] ?? "Week2SundayHours";

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
  if (payload.week2SundayHours !== undefined) fields[w2SunKey] = payload.week2SundayHours;
  if (payload.week2MondayHours !== undefined) fields[w2MonKey] = payload.week2MondayHours;
  if (payload.week2TuesdayHours !== undefined) fields[w2TueKey] = payload.week2TuesdayHours;
  if (payload.week2WednesdayHours !== undefined) fields[w2WedKey] = payload.week2WednesdayHours;
  if (payload.week2ThursdayHours !== undefined) fields[w2ThuKey] = payload.week2ThursdayHours;
  if (payload.week2FridayHours !== undefined) fields[w2FriKey] = payload.week2FridayHours;
  if (payload.week2SaturdayHours !== undefined) fields[w2SatKey] = payload.week2SaturdayHours;

  await sharepoint.updateListItem(accessToken, siteId, listId, budgetListItemId, fields);
}
