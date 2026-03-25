import React, { useState, useEffect, useCallback, useRef, useMemo } from "react";
import { Plus, Edit3, X, Loader2, Trash2, Layers, Search, ChevronUp, ChevronDown, UserMinus, UserPlus } from "lucide-react";
import { useRole } from "../contexts/RoleContext";
import { useAppAuth } from "../contexts/AppAuthContext";
import { getGraphAccessToken } from "../lib/graph";
import { msalInstance } from "../lib/msal";
import { GRAPH_SCOPES } from "../src/auth/graphScopes";
import { InteractionRequiredAuthError } from "@azure/msal-browser";
import {
  getSites,
  createSite,
  updateSite,
  setSiteActive,
  deleteSite,
  type Site,
  type SitePayload,
  type SiteNoServicePeriod,
} from "../repositories/sitesRepo";
import { getSchoolHolidayPeriods, getSupportedSchoolHolidayYears } from "../lib/schoolHolidays";
import { getAssignedSiteIdsForManager, createSiteManagerAssignment, getAssignedManagersForSite, deleteSiteManagerAssignment, fetchSiteManagerAssignments, joinAssignmentsToSites } from "../repositories/siteManagersRepo";
import { getCleanTrackManagers } from "../repositories/usersRepo";
import { createSiteBudget, getSiteBudgets, updateSiteBudget, type SiteBudgetHours } from "../repositories/budgetsRepo";
import { getSiteCleanerAssignments } from "../repositories/assignedCleanersRepo";
import { normalizeListItemId } from "../lib/sharepoint";
import { AU_STATES } from "../lib/auStates";
const DAY_LABELS = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"] as const;
type DayKey = (typeof DAY_LABELS)[number];
/** Display order: Monday first, Sunday last. */
const DAY_KEYS: DayKey[] = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"];

type VisitFreq = "Weekly" | "Fortnightly" | "Monthly";

/** Empty = not set in UI / SharePoint; saved `0` stays numeric `0`. */
type DayHourField = number | "";
const EMPTY_DAY_HOURS: Record<DayKey, DayHourField> = { Sun: "", Mon: "", Tue: "", Wed: "", Thu: "", Fri: "", Sat: "" };

function dayHourToNum(v: DayHourField): number {
  return v === "" ? 0 : v;
}

function sumDayHours(rec: Record<DayKey, DayHourField>): number {
  return DAY_KEYS.reduce((s, d) => s + dayHourToNum(rec[d]), 0);
}

function budgetHourToField(n: number | undefined): DayHourField {
  return n !== undefined ? n : "";
}

function budgetHasAnyWeek2(b: SiteBudgetHours): boolean {
  return (
    b.week2Sunday !== undefined ||
    b.week2Monday !== undefined ||
    b.week2Tuesday !== undefined ||
    b.week2Wednesday !== undefined ||
    b.week2Thursday !== undefined ||
    b.week2Friday !== undefined ||
    b.week2Saturday !== undefined
  );
}

/**
 * Display rates for list/cards; matches edit-modal fallbacks (legacy Budget / Weekend labour columns).
 */
function effectiveBudgetLabourRates(budget: SiteBudgetHours | undefined): {
  weekday: number | undefined;
  saturday: number | undefined;
  sunday: number | undefined;
  ph: number | undefined;
} {
  if (!budget) {
    return { weekday: undefined, saturday: undefined, sunday: undefined, ph: undefined };
  }
  const weekdayRaw = budget.weekdayLabourRate ?? budget.budgetLabourRate;
  const weekday =
    weekdayRaw != null && typeof weekdayRaw === "number" && weekdayRaw >= 0 ? weekdayRaw : undefined;
  const weekend =
    budget.weekendLabourRate != null &&
    typeof budget.weekendLabourRate === "number" &&
    budget.weekendLabourRate >= 0
      ? budget.weekendLabourRate
      : undefined;
  const saturday =
    budget.saturdayLabourRate != null && budget.saturdayLabourRate >= 0
      ? budget.saturdayLabourRate
      : weekend;
  const sunday =
    budget.sundayLabourRate != null && budget.sundayLabourRate >= 0 ? budget.sundayLabourRate : weekend;
  const ph = budget.phLabourRate != null && budget.phLabourRate >= 0 ? budget.phLabourRate : undefined;
  return { weekday, saturday, sunday, ph };
}

type BudgetLabourRatesDisplay = ReturnType<typeof effectiveBudgetLabourRates>;

/** Single-line labour rates for table cells (MF / Sa / Su / PH). */
function LabourRatesInlineCell({ rates }: { rates: BudgetLabourRatesDisplay }) {
  const fmt = (r: number | undefined) => (r != null && r >= 0 ? `$${Number(r).toFixed(2)}` : "—");
  const piece = (abbr: string, expand: string, rate: number | undefined) => (
    <span className="whitespace-nowrap">
      <abbr title={expand} className="text-gray-500 font-medium no-underline cursor-help">
        {abbr}
      </abbr>{" "}
      <span className="font-semibold text-gray-900">{fmt(rate)}</span>
    </span>
  );
  return (
    <div
      className="text-[11px] md:text-[10px] text-gray-800 tabular-nums leading-snug whitespace-normal md:whitespace-nowrap"
      title={`Mon–Fri ${fmt(rates.weekday)} · Sat ${fmt(rates.saturday)} · Sun ${fmt(rates.sunday)} · PH ${fmt(rates.ph)} (per hour)`}
    >
      <span className="inline-flex flex-wrap md:flex-nowrap items-baseline gap-x-2 gap-y-0.5">
        {piece("MF", "Monday–Friday", rates.weekday)}
        <span className="text-gray-300 select-none" aria-hidden>
          ·
        </span>
        {piece("Sa", "Saturday", rates.saturday)}
        <span className="text-gray-300 select-none" aria-hidden>
          ·
        </span>
        {piece("Su", "Sunday", rates.sunday)}
        <span className="text-gray-300 select-none" aria-hidden>
          ·
        </span>
        {piece("PH", "Public holiday", rates.ph)}
      </span>
    </div>
  );
}

function isFortnightlyBudgetVisit(b: SiteBudgetHours): boolean {
  return String(b.visitFrequency ?? "")
    .toLowerCase()
    .includes("fortnight");
}

function SiteBudgetDayHoursGrid({ budget }: { budget: SiteBudgetHours }) {
  type DayH = { day: string; h: number | undefined };
  const week1: DayH[] = [
    { day: "Mon", h: budget.monday },
    { day: "Tue", h: budget.tuesday },
    { day: "Wed", h: budget.wednesday },
    { day: "Thu", h: budget.thursday },
    { day: "Fri", h: budget.friday },
    { day: "Sat", h: budget.saturday },
    { day: "Sun", h: budget.sunday },
  ];
  const week2: DayH[] = [
    { day: "Mon", h: budget.week2Monday },
    { day: "Tue", h: budget.week2Tuesday },
    { day: "Wed", h: budget.week2Wednesday },
    { day: "Thu", h: budget.week2Thursday },
    { day: "Fri", h: budget.week2Friday },
    { day: "Sat", h: budget.week2Saturday },
    { day: "Sun", h: budget.week2Sunday },
  ];
  const showTwoWeekRows = isFortnightlyBudgetVisit(budget) && budgetHasAnyWeek2(budget);

  const cell = ({ day, h }: DayH) => {
    const filled = h != null && h > 0;
    const short = day.slice(0, 2);
    return (
      <div
        key={day}
        className={`min-h-[2.65rem] px-1 py-1 flex flex-col items-center justify-center text-center ${
          filled ? "text-blue-700" : "text-gray-400"
        } bg-white`}
        title={h !== undefined ? `${day}: ${h}h` : `${day}: not set`}
      >
        <span className="text-[7px] sm:text-[8px] font-bold uppercase tracking-tight text-gray-500 leading-none">{short}</span>
        <span className="text-[10px] sm:text-[11px] font-bold tabular-nums leading-none mt-0.5">
          {h !== undefined ? String(h) : "—"}
        </span>
      </div>
    );
  };

  const row = (days: DayH[], label: string | null) => (
    <div key={label ?? "single"}>
      {label ? (
        <div className="text-[9px] font-semibold text-gray-400 uppercase tracking-wider mb-0.5 text-center">{label}</div>
      ) : null}
      <div className="grid grid-cols-7 gap-px bg-[#edeef0] rounded-md border border-[#edeef0] overflow-hidden shadow-sm">
        {days.map((d) => cell(d))}
      </div>
    </div>
  );

  const ariaSummary = (days: DayH[]) => days.map(({ day, h }) => `${day} ${h ?? "—"}h`).join(", ");
  const ariaLabel = showTwoWeekRows
    ? `Week 1: ${ariaSummary(week1)}. Week 2: ${ariaSummary(week2)}`
    : `Planned hours per day: ${ariaSummary(week1)}`;

  return (
    <div className="w-full max-w-[min(100%,22rem)] md:max-w-[26rem] mx-auto space-y-1.5" aria-label={ariaLabel}>
      {showTwoWeekRows ? (
        <>
          {row(week1, "Week 1")}
          {row(week2, "Week 2")}
        </>
      ) : (
        row(week1, null)
      )}
    </div>
  );
}

interface BulkSiteRow {
  id: string;
  siteName: string;
  address: string;
  monthlyRevenue: number | "";
  fortnightCostBudget: number | "";
  /** Weekday (Mon–Fri) labour rate — saved as Weekday Labour Rate in SharePoint. */
  weekdayLabourRate: number | "";
  state: string;
  active: boolean;
  visitFrequency: VisitFreq;
  dailyHours: Record<DayKey, DayHourField>;
  dailyHoursWeek2: Record<DayKey, DayHourField>;
  hoursPerVisit: number | "";
}

type NoServiceMode = "manual" | "school_auto";

interface NoServicePeriodRow {
  id: string;
  label: string;
  startDate: string;
  endDate: string;
  reason: string;
  source?: "manual" | "school_holidays_auto";
  state?: string;
  year?: number;
}

function inferNoServiceModeFromSite(site: Site): NoServiceMode {
  const list = site.noServicePeriods;
  if (!list?.length) return "manual";
  return list.some((p) => p.source === "school_holidays_auto") ? "school_auto" : "manual";
}

function inferNoServiceYearFromSite(site: Site): number {
  const y = site.noServicePeriods?.find(
    (p) => p.source === "school_holidays_auto" && typeof p.year === "number"
  )?.year;
  return y ?? new Date().getFullYear();
}

const WRITE_PERMISSION_HINT =
  "Admin consent may be required for Sites.ReadWrite.All in the Azure AD app registration. If consent is already granted, sign out and sign in again to get a token with write scope.";

interface SiteManagerProps {
  /** Called after sites data changes. Pass true to refresh without showing the global loading spinner (keeps search and list visible). */
  onUpdateSite?: (silent?: boolean) => void;
  onViewSite?: (siteId: string) => void;
  /** When this changes, sites and assigned cleaners are refetched (e.g. after returning from Site Detail). */
  refreshTrigger?: number;
}

const SiteManager: React.FC<SiteManagerProps> = ({ onUpdateSite, onViewSite, refreshTrigger }) => {
  const { isAdmin: isAdminFromRole } = useRole();
  const { user } = useAppAuth();
  const isAdmin = isAdminFromRole || user?.role === "Admin";

  const [sites, setSites] = useState<Site[]>([]);
  const [budgetsBySiteId, setBudgetsBySiteId] = useState<Record<string, SiteBudgetHours>>({});
  const [assignedManagersBySiteId, setAssignedManagersBySiteId] = useState<
    Record<string, { assignedManagers: { managerName: string; assignmentItemId: string }[] }>
  >({});
  const [assignedCleanersBySiteId, setAssignedCleanersBySiteId] = useState<
    Record<string, string[]>
  >({});
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [toast, setToast] = useState<string | null>(null);
  const [modalMode, setModalMode] = useState<"add" | "edit" | null>(null);
  const [editingSite, setEditingSite] = useState<Site | null>(null);
  const [form, setForm] = useState<SitePayload>({
    siteName: "",
    address: "",
    state: "",
    active: true,
    monthlyRevenue: null,
  });
  const [submitLoading, setSubmitLoading] = useState(false);
  const [submitError, setSubmitError] = useState<string | null>(null);
  const [managers, setManagers] = useState<{ id: string; fullName: string; email: string }[]>([]);
  const [dailyHours, setDailyHours] = useState<Record<DayKey, DayHourField>>({ ...EMPTY_DAY_HOURS });
  const [dailyHoursWeek2, setDailyHoursWeek2] = useState<Record<DayKey, DayHourField>>({ ...EMPTY_DAY_HOURS });
  const [fortnightCostBudget, setFortnightCostBudget] = useState<number | "">("");
  const [selectedManagerIds, setSelectedManagerIds] = useState<string[]>([]);
  const [visitFrequency, setVisitFrequency] = useState<VisitFreq>("Weekly");
  const [hoursPerVisit, setHoursPerVisit] = useState<number | "">("");
  const [weekdayLabourRate, setWeekdayLabourRate] = useState<number | "">("");
  const [saturdayLabourRate, setSaturdayLabourRate] = useState<number | "">("");
  const [sundayLabourRate, setSundayLabourRate] = useState<number | "">("");
  const [phLabourRate, setPhLabourRate] = useState<number | "">("");
  const [noServicePeriods, setNoServicePeriods] = useState<NoServicePeriodRow[]>([]);
  const [noServiceMode, setNoServiceMode] = useState<NoServiceMode>("manual");
  const [noServiceYear, setNoServiceYear] = useState<number>(() => new Date().getFullYear());
  const [noServiceHolidayMessage, setNoServiceHolidayMessage] = useState<string | null>(null);

  const [bulkModalOpen, setBulkModalOpen] = useState(false);
  const [bulkRows, setBulkRows] = useState<BulkSiteRow[]>([]);
  const [bulkManagerIds, setBulkManagerIds] = useState<string[]>([]);
  const [bulkSubmitLoading, setBulkSubmitLoading] = useState(false);
  const [bulkProgress, setBulkProgress] = useState<{ current: number; total: number } | null>(null);
  const [bulkError, setBulkError] = useState<string | null>(null);
  const [siteSearchQuery, setSiteSearchQuery] = useState("");
  type SiteSortKey = "name" | "address" | "state" | "monthlyRevenue" | "fortnightlyCap";
  const [siteSortBy, setSiteSortBy] = useState<SiteSortKey>("name");
  const [siteSortDir, setSiteSortDir] = useState<"asc" | "desc">("asc");
  const [selectedSiteIds, setSelectedSiteIds] = useState<string[]>([]);
  const [bulkDeleteLoading, setBulkDeleteLoading] = useState(false);
  const editingSiteIdRef = useRef<string | null>(null);

  const createNoServiceRow = useCallback(
    (): NoServicePeriodRow => ({
      id: `nsp-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
      label: "",
      startDate: "",
      endDate: "",
      reason: "",
    }),
    []
  );

  const showToast = useCallback((msg: string) => {
    setToast(msg);
    setTimeout(() => setToast(null), 4000);
  }, []);

  const loadSites = useCallback(async (silent = false) => {
    const token = await getGraphAccessToken();
    if (!token) {
      setError("Sign in with Microsoft to view sites.");
      setSites([]);
      setLoading(false);
      return;
    }
    if (!silent) {
      setLoading(true);
      setError(null);
    }
    try {
      let data = await getSites(token);
      if (!isAdmin && user?.role === "Manager" && user?.email) {
        const isAllSites = user.permissionScope?.trim().toLowerCase() === "allsites";
        if (!isAllSites) {
          const assignedIds = await getAssignedSiteIdsForManager(token, user.email);
          data = assignedIds.length > 0 ? data.filter((s) => assignedIds.includes(s.id)) : [];
        }
      }
      setSites(data);
      const [budgets, managersList, assignments, cleanerAssignments] = await Promise.all([
        getSiteBudgets(token).catch(() => ({})),
        getCleanTrackManagers(token).catch(() => []),
        fetchSiteManagerAssignments(token).catch(() => []),
        getSiteCleanerAssignments(token, { activeOnly: true }).catch((err) => {
          console.warn("[SiteManager] getSiteCleanerAssignments failed:", err);
          return [];
        }),
      ]);
      setBudgetsBySiteId(budgets);
      setManagers(managersList);
      const joined = joinAssignmentsToSites(data, assignments);
      setAssignedManagersBySiteId(
        Object.fromEntries(
          Object.entries(joined).map(([id, v]) => [
            id,
            { assignedManagers: v.assignedManagers.map((a) => ({ managerName: a.managerName, assignmentItemId: a.id })) },
          ])
        )
      );
      const cleanersBySite: Record<string, string[]> = {};
      for (const a of cleanerAssignments) {
        if (!a.siteId || !a.cleanerName || !a.active) continue;
        const key = normalizeListItemId(String(a.siteId));
        if (!key) continue;
        if (!cleanersBySite[key]) cleanersBySite[key] = [];
        if (!cleanersBySite[key].includes(a.cleanerName)) {
          cleanersBySite[key].push(a.cleanerName);
        }
      }
      setAssignedCleanersBySiteId(cleanersBySite);
    } catch (e) {
      const msg = e instanceof Error ? e.message : "Failed to load sites.";
      setError(msg);
      setSites([]);
    } finally {
      if (!silent) setLoading(false);
    }
  }, [isAdmin, user?.role, user?.email]);

  useEffect(() => {
    loadSites();
  }, [loadSites]);

  /** Refetch when parent signals (e.g. user returned to Sites from Site Detail after assigning cleaners). */
  useEffect(() => {
    if (refreshTrigger != null && refreshTrigger > 0) {
      loadSites(true);
    }
  }, [refreshTrigger, loadSites]);

  const filteredSites = React.useMemo(() => {
    const q = siteSearchQuery.trim().toLowerCase();
    if (!q) return sites;
    return sites.filter(
      (s) =>
        (s.siteName?.toLowerCase().includes(q) ?? false) ||
        (s.address?.toLowerCase().includes(q) ?? false)
    );
  }, [sites, siteSearchQuery]);

  const sortedSites = React.useMemo(() => {
    const list = [...filteredSites];
    list.sort((a, b) => {
      const budgetA = budgetsBySiteId[String(a.id)] ?? budgetsBySiteId["name:" + ((a.siteName || "").trim() + " Budget")];
      const budgetB = budgetsBySiteId[String(b.id)] ?? budgetsBySiteId["name:" + ((b.siteName || "").trim() + " Budget")];
      const capA = budgetA?.fortnightCap ?? 0;
      const capB = budgetB?.fortnightCap ?? 0;
      let cmp = 0;
      if (siteSortBy === "name") {
        const na = (a.siteName || "Unnamed site").toLowerCase();
        const nb = (b.siteName || "Unnamed site").toLowerCase();
        cmp = na.localeCompare(nb);
      } else if (siteSortBy === "address") {
        const aa = (a.address || "").toLowerCase();
        const ab = (b.address || "").toLowerCase();
        cmp = aa.localeCompare(ab);
      } else if (siteSortBy === "state") {
        const sa = (a.state || "").toLowerCase();
        const sb = (b.state || "").toLowerCase();
        cmp = sa.localeCompare(sb);
      } else if (siteSortBy === "monthlyRevenue") {
        const ra = Number(a.monthlyRevenue ?? 0);
        const rb = Number(b.monthlyRevenue ?? 0);
        cmp = ra - rb;
      } else {
        cmp = capA - capB;
      }
      return siteSortDir === "asc" ? cmp : -cmp;
    });
    return list;
  }, [filteredSites, siteSortBy, siteSortDir, budgetsBySiteId]);

  const handleSiteSort = (key: SiteSortKey) => {
    if (siteSortBy === key) {
      setSiteSortDir((d) => (d === "asc" ? "desc" : "asc"));
    } else {
      setSiteSortBy(key);
      setSiteSortDir(key === "fortnightlyCap" || key === "monthlyRevenue" ? "desc" : "asc");
    }
  };

  const validateManualNoServiceRows = (
    rows: NoServicePeriodRow[]
  ): { periods: SiteNoServicePeriod[]; error?: string } => {
    const periods: SiteNoServicePeriod[] = [];
    for (const row of rows) {
      const label = row.label.trim();
      const start = row.startDate.trim();
      const end = row.endDate.trim();
      const reason = row.reason.trim();
      const isEmpty = !label && !start && !end && !reason;
      if (isEmpty) continue;
      if (!start || !end) {
        return { periods: [], error: "No Service Period rows require both Start Date and End Date." };
      }
      if (end < start) {
        return { periods: [], error: "No Service Period End Date must be on or after Start Date." };
      }
      periods.push({
        ...(label ? { label } : {}),
        start_date: start,
        end_date: end,
        ...(reason ? { reason } : {}),
        source: "manual",
      });
    }
    return { periods };
  };

  const validateAutoNoServiceRows = (
    rows: NoServicePeriodRow[],
    siteState: string,
    year: number
  ): { periods: SiteNoServicePeriod[]; error?: string } => {
    const st = siteState.trim();
    if (!st) {
      return {
        periods: [],
        error: "Select State on this site before saving School Holidays (Auto).",
      };
    }
    if (rows.length === 0) {
      return {
        periods: [],
        error: "Populate school holidays before saving, or switch to Manual mode.",
      };
    }
    const periods: SiteNoServicePeriod[] = [];
    const upper = st.toUpperCase();
    for (const row of rows) {
      const start = row.startDate.trim();
      const end = row.endDate.trim();
      if (!start || !end) {
        return {
          periods: [],
          error: "Each period needs Start and End dates. Click Populate School Holidays again.",
        };
      }
      if (end < start) {
        return { periods: [], error: "No Service Period End Date must be on or after Start Date." };
      }
      periods.push({
        ...(row.label.trim() ? { label: row.label.trim() } : {}),
        start_date: start,
        end_date: end,
        reason: row.reason.trim() || "School holidays",
        source: "school_holidays_auto",
        state: upper,
        year,
      });
    }
    return { periods };
  };

  const handlePopulateSchoolHolidays = () => {
    setNoServiceHolidayMessage(null);
    const st = form.state.trim();
    if (!st) {
      setNoServiceHolidayMessage("Select State (above) before populating school holidays.");
      return;
    }
    const res = getSchoolHolidayPeriods(st, noServiceYear);
    if (!res.ok) {
      setNoServiceHolidayMessage(res.error);
      return;
    }
    if (res.periods.length === 0) {
      setNoServiceHolidayMessage("No school holiday periods were returned for that state and year.");
      return;
    }
    setNoServicePeriods(
      res.periods.map((p, i) => ({
        id: `nsp-auto-${Date.now()}-${i}`,
        label: p.label ?? "",
        startDate: p.start_date,
        endDate: p.end_date,
        reason: p.reason ?? "School holidays",
        source: "school_holidays_auto" as const,
        state: p.state,
        year: p.year,
      }))
    );
    setNoServiceHolidayMessage(
      `Loaded ${res.periods.length} period(s) for ${st.toUpperCase()} ${noServiceYear}. Click Save to store.`
    );
  };

  const schoolHolidayYearOptions = useMemo(() => {
    const cy = new Date().getFullYear();
    const embedded = getSupportedSchoolHolidayYears();
    const set = new Set<number>([...embedded, cy, cy + 1, cy + 2, cy - 1, noServiceYear]);
    return [...set].sort((a, b) => a - b);
  }, [noServiceYear]);

  const openAdd = () => {
    setModalMode("add");
    setEditingSite(null);
    editingSiteIdRef.current = null;
    setForm({
      siteName: "",
      address: "",
      state: "",
      active: true,
      monthlyRevenue: null,
    });
    setDailyHours({ ...EMPTY_DAY_HOURS });
    setDailyHoursWeek2({ ...EMPTY_DAY_HOURS });
    setFortnightCostBudget("");
    setVisitFrequency("Weekly");
    setHoursPerVisit("");
    setWeekdayLabourRate("");
    setSaturdayLabourRate("");
    setSundayLabourRate("");
    setPhLabourRate("");
    setNoServicePeriods([]);
    setNoServiceMode("manual");
    setNoServiceYear(new Date().getFullYear());
    setNoServiceHolidayMessage(null);
    setSelectedManagerIds([]);
    setSubmitError(null);
    getGraphAccessToken().then((token) => {
      if (token) getCleanTrackManagers(token).then(setManagers).catch(() => setManagers([]));
      else setManagers([]);
    });
  };

  const openEdit = (site: Site) => {
    setModalMode("edit");
    setEditingSite(site);
    editingSiteIdRef.current = site.id;
    setForm({
      siteName: site.siteName,
      address: site.address,
      state: site.state,
      active: site.active,
      monthlyRevenue: site.monthlyRevenue,
    });
    const budget = budgetsBySiteId[String(site.id)] ?? budgetsBySiteId["name:" + (site.siteName.trim() + " Budget")];
    setDailyHours(
      budget
        ? {
            Sun: budgetHourToField(budget.sunday),
            Mon: budgetHourToField(budget.monday),
            Tue: budgetHourToField(budget.tuesday),
            Wed: budgetHourToField(budget.wednesday),
            Thu: budgetHourToField(budget.thursday),
            Fri: budgetHourToField(budget.friday),
            Sat: budgetHourToField(budget.saturday),
          }
        : { ...EMPTY_DAY_HOURS }
    );
    setDailyHoursWeek2(
      budget && budgetHasAnyWeek2(budget)
        ? {
            Sun: budgetHourToField(budget.week2Sunday),
            Mon: budgetHourToField(budget.week2Monday),
            Tue: budgetHourToField(budget.week2Tuesday),
            Wed: budgetHourToField(budget.week2Wednesday),
            Thu: budgetHourToField(budget.week2Thursday),
            Fri: budgetHourToField(budget.week2Friday),
            Sat: budgetHourToField(budget.week2Saturday),
          }
        : { ...EMPTY_DAY_HOURS }
    );
    setFortnightCostBudget(budget?.fortnightCostBudget != null ? budget.fortnightCostBudget : "");
    setSelectedManagerIds([]);
    const freq = budget?.visitFrequency;
    const normalizedFreq: VisitFreq =
      freq === "Fortnightly" || freq === "Monthly" ? freq : "Weekly";
    setVisitFrequency(normalizedFreq);
    setHoursPerVisit(
      budget?.hoursPerVisit != null && budget.hoursPerVisit > 0 ? budget.hoursPerVisit : ""
    );
    setWeekdayLabourRate(
      budget?.weekdayLabourRate != null && budget.weekdayLabourRate >= 0
        ? budget.weekdayLabourRate
        : budget?.budgetLabourRate != null && budget.budgetLabourRate >= 0
          ? budget.budgetLabourRate
          : ""
    );
    setSaturdayLabourRate(
      budget?.saturdayLabourRate != null && budget.saturdayLabourRate >= 0 ? budget.saturdayLabourRate : (budget?.weekendLabourRate != null && budget.weekendLabourRate >= 0 ? budget.weekendLabourRate : "")
    );
    setSundayLabourRate(
      budget?.sundayLabourRate != null && budget.sundayLabourRate >= 0 ? budget.sundayLabourRate : (budget?.weekendLabourRate != null && budget.weekendLabourRate >= 0 ? budget.weekendLabourRate : "")
    );
    setPhLabourRate(
      budget?.phLabourRate != null && budget.phLabourRate >= 0 ? budget.phLabourRate : ""
    );
    setNoServiceMode(inferNoServiceModeFromSite(site));
    setNoServiceYear(inferNoServiceYearFromSite(site));
    setNoServiceHolidayMessage(null);
    setNoServicePeriods(
      (site.noServicePeriods ?? []).map((p, idx) => ({
        id: `nsp-${site.id}-${idx}`,
        label: String(p.label ?? ""),
        startDate: String(p.start_date ?? ""),
        endDate: String(p.end_date ?? ""),
        reason: String(p.reason ?? ""),
        source:
          p.source === "school_holidays_auto" || p.source === "manual" ? p.source : undefined,
        state: p.state,
        year: typeof p.year === "number" ? p.year : undefined,
      }))
    );
    setSubmitError(null);
    getGraphAccessToken().then(async (token) => {
      if (!token) return;
      const siteIdBeingLoaded = site.id;
      const [managersList, assignments] = await Promise.all([
        getCleanTrackManagers(token).catch(() => []),
        getAssignedManagersForSite(token, site.id).catch(() => []),
      ]);
      setManagers(managersList);
      if (editingSiteIdRef.current === siteIdBeingLoaded) {
        setSelectedManagerIds(assignments.map((a) => a.managerId));
      }
    });
  };

  const closeModal = () => {
    setModalMode(null);
    setEditingSite(null);
    editingSiteIdRef.current = null;
    setSubmitError(null);
    setNoServiceHolidayMessage(null);
  };

  /** Create a new empty bulk row. */
  const createEmptyBulkRow = useCallback((): BulkSiteRow => ({
    id: `bulk-${Date.now()}-${Math.random().toString(36).slice(2)}`,
    siteName: "",
    address: "",
    monthlyRevenue: "",
    fortnightCostBudget: "",
    weekdayLabourRate: "",
    state: "",
    active: true,
    visitFrequency: "Weekly",
    dailyHours: { ...EMPTY_DAY_HOURS },
    dailyHoursWeek2: { ...EMPTY_DAY_HOURS },
    hoursPerVisit: "",
  }), []);

  const openBulkModal = () => {
    setBulkModalOpen(true);
    setBulkRows([createEmptyBulkRow()]);
    setBulkManagerIds([]);
    setBulkProgress(null);
    setBulkError(null);
    getGraphAccessToken().then((token) => {
      if (token) getCleanTrackManagers(token).then(setManagers).catch(() => setManagers([]));
      else setManagers([]);
    });
  };

  const updateBulkRow = useCallback((id: string, updates: Partial<BulkSiteRow>) => {
    setBulkRows((prev) => prev.map((r) => (r.id === id ? { ...r, ...updates } : r)));
  }, []);

  const addBulkRow = useCallback(() => {
    setBulkRows((prev) => [...prev, createEmptyBulkRow()]);
  }, [createEmptyBulkRow]);

  const removeBulkRow = useCallback((id: string) => {
    setBulkRows((prev) => (prev.length <= 1 ? prev : prev.filter((r) => r.id !== id)));
  }, []);

  const handleBulkSubmit = async () => {
    const rows = bulkRows.filter((r) => r.siteName.trim());
    if (rows.length === 0) {
      setBulkError("Enter at least one site with a Site Name.");
      return;
    }
    const token = await getGraphAccessToken();
    if (!token) {
      setBulkError("Sign in with Microsoft to add sites.");
      return;
    }
    setBulkSubmitLoading(true);
    setBulkError(null);
    setBulkProgress({ current: 0, total: rows.length });
    let successCount = 0;
    for (let i = 0; i < rows.length; i++) {
      setBulkProgress({ current: i + 1, total: rows.length });
      const row = rows[i];
      const siteName = row.siteName.trim();
      const isMonthly = row.visitFrequency === "Monthly";
      const isFortnightly = row.visitFrequency === "Fortnightly";
      try {
        const created = await createSite(token, {
          siteName,
          address: row.address.trim() || undefined,
          state: row.state.trim() || undefined,
          active: row.active !== false,
          monthlyRevenue: row.monthlyRevenue === "" ? null : Number(row.monthlyRevenue),
        });
        const siteId = created.id;
        try {
          await createSiteBudget(token, {
            budgetName: `${siteName} Budget`,
            siteListItemId: siteId,
            sundayHours: isMonthly ? 0 : dayHourToNum(row.dailyHours.Sun),
            mondayHours: isMonthly ? 0 : dayHourToNum(row.dailyHours.Mon),
            tuesdayHours: isMonthly ? 0 : dayHourToNum(row.dailyHours.Tue),
            wednesdayHours: isMonthly ? 0 : dayHourToNum(row.dailyHours.Wed),
            thursdayHours: isMonthly ? 0 : dayHourToNum(row.dailyHours.Thu),
            fridayHours: isMonthly ? 0 : dayHourToNum(row.dailyHours.Fri),
            saturdayHours: isMonthly ? 0 : dayHourToNum(row.dailyHours.Sat),
            active: true,
            visitFrequency: row.visitFrequency,
            ...(isMonthly && row.hoursPerVisit !== "" && { hoursPerVisit: Number(row.hoursPerVisit) }),
            ...(isFortnightly && {
              week2SundayHours: dayHourToNum(row.dailyHoursWeek2.Sun),
              week2MondayHours: dayHourToNum(row.dailyHoursWeek2.Mon),
              week2TuesdayHours: dayHourToNum(row.dailyHoursWeek2.Tue),
              week2WednesdayHours: dayHourToNum(row.dailyHoursWeek2.Wed),
              week2ThursdayHours: dayHourToNum(row.dailyHoursWeek2.Thu),
              week2FridayHours: dayHourToNum(row.dailyHoursWeek2.Fri),
              week2SaturdayHours: dayHourToNum(row.dailyHoursWeek2.Sat),
            }),
            ...(row.weekdayLabourRate !== "" && !Number.isNaN(Number(row.weekdayLabourRate)) && { weekdayLabourRate: Number(row.weekdayLabourRate) }),
          });
        } catch (budgetErr) {
          console.warn("Bulk: site budget create failed for", siteName, budgetErr);
        }
        for (const managerId of bulkManagerIds) {
          try {
            const managerFullName = managers.find((m) => m.id === managerId)?.fullName ?? "";
            await createSiteManagerAssignment(token, siteId, managerId, {
              siteName,
              managerFullName,
            });
          } catch {
            // ignore per-manager errors
          }
        }
        successCount++;
      } catch (e) {
        const msg = e instanceof Error ? e.message : "Unknown error";
        setBulkError(`Site "${siteName}": ${msg}`);
        setBulkSubmitLoading(false);
        setBulkProgress(null);
        return;
      }
    }
    setBulkProgress(null);
    setBulkSubmitLoading(false);
    setBulkModalOpen(false);
    showToast(successCount === rows.length ? `Added ${successCount} sites.` : `Added ${successCount} of ${rows.length} sites.`);
    await loadSites(true);
    onUpdateSite?.(true);
  };

  const handleSubmit = async () => {
    if (!form.siteName.trim()) {
      setSubmitError("Site Name is required.");
      return;
    }
    const noServiceValidation =
      noServiceMode === "manual"
        ? validateManualNoServiceRows(noServicePeriods)
        : validateAutoNoServiceRows(noServicePeriods, form.state, noServiceYear);
    if (noServiceValidation.error) {
      setSubmitError(noServiceValidation.error);
      return;
    }
    const token = await getGraphAccessToken();
    if (!token) {
      setSubmitError("Not signed in. Sign in with Microsoft to save.");
      return;
    }
    setSubmitLoading(true);
    setSubmitError(null);
    try {
      if (modalMode === "add") {
        const created = await createSite(token, {
          siteName: form.siteName.trim(),
          address: form.address.trim() || undefined,
          state: form.state || undefined,
          active: form.active !== false,
          monthlyRevenue: form.monthlyRevenue ?? null,
          noServicePeriods: noServiceValidation.periods,
        });
        const siteId = created.id;
        const isMonthly = visitFrequency === "Monthly";
        const isFortnightly = visitFrequency === "Fortnightly";
        try {
          await createSiteBudget(token, {
            budgetName: `${form.siteName.trim()} Budget`,
            siteListItemId: siteId,
            sundayHours: isMonthly ? 0 : dayHourToNum(dailyHours.Sun),
            mondayHours: isMonthly ? 0 : dayHourToNum(dailyHours.Mon),
            tuesdayHours: isMonthly ? 0 : dayHourToNum(dailyHours.Tue),
            wednesdayHours: isMonthly ? 0 : dayHourToNum(dailyHours.Wed),
            thursdayHours: isMonthly ? 0 : dayHourToNum(dailyHours.Thu),
            fridayHours: isMonthly ? 0 : dayHourToNum(dailyHours.Fri),
            saturdayHours: isMonthly ? 0 : dayHourToNum(dailyHours.Sat),
            active: true,
            visitFrequency,
            ...(isMonthly && hoursPerVisit !== "" && { hoursPerVisit: Number(hoursPerVisit) }),
            ...(isFortnightly && {
              week2SundayHours: dayHourToNum(dailyHoursWeek2.Sun),
              week2MondayHours: dayHourToNum(dailyHoursWeek2.Mon),
              week2TuesdayHours: dayHourToNum(dailyHoursWeek2.Tue),
              week2WednesdayHours: dayHourToNum(dailyHoursWeek2.Wed),
              week2ThursdayHours: dayHourToNum(dailyHoursWeek2.Thu),
              week2FridayHours: dayHourToNum(dailyHoursWeek2.Fri),
              week2SaturdayHours: dayHourToNum(dailyHoursWeek2.Sat),
            }),
            ...(weekdayLabourRate !== "" && !Number.isNaN(Number(weekdayLabourRate)) && Number(weekdayLabourRate) >= 0 && { weekdayLabourRate: Number(weekdayLabourRate) }),
            ...(saturdayLabourRate !== "" && saturdayLabourRate != null && !Number.isNaN(Number(saturdayLabourRate)) && Number(saturdayLabourRate) >= 0 && { saturdayLabourRate: Number(saturdayLabourRate) }),
            ...(sundayLabourRate !== "" && sundayLabourRate != null && !Number.isNaN(Number(sundayLabourRate)) && Number(sundayLabourRate) >= 0 && { sundayLabourRate: Number(sundayLabourRate) }),
            ...(phLabourRate !== "" && phLabourRate != null && !Number.isNaN(Number(phLabourRate)) && Number(phLabourRate) >= 0 && { phLabourRate: Number(phLabourRate) }),
          });
        } catch (budgetErr) {
          console.warn("Site budget create failed (site was created):", budgetErr);
        }
        for (const managerId of selectedManagerIds) {
          try {
            const managerFullName = managers.find((m) => m.id === managerId)?.fullName ?? "";
            await createSiteManagerAssignment(token, siteId, managerId, {
              siteName: form.siteName.trim(),
              managerFullName,
            });
          } catch (assignErr) {
            console.warn("Manager assignment failed for", managerId, assignErr);
          }
        }
        showToast("Site added.");
      } else if (modalMode === "edit" && editingSite) {
        await updateSite(token, editingSite.id, {
          siteName: form.siteName.trim(),
          address: form.address.trim(),
          state: form.state || undefined,
          active: form.active,
          monthlyRevenue: form.monthlyRevenue ?? null,
          noServicePeriods: noServiceValidation.periods,
        });
        const budget = budgetsBySiteId[String(editingSite.id)] ?? budgetsBySiteId["name:" + (editingSite.siteName.trim() + " Budget")];
        const isMonthly = visitFrequency === "Monthly";
        const isFortnightly = visitFrequency === "Fortnightly";
        const budgetPayload = {
          sundayHours: isMonthly ? 0 : dayHourToNum(dailyHours.Sun),
          mondayHours: isMonthly ? 0 : dayHourToNum(dailyHours.Mon),
          tuesdayHours: isMonthly ? 0 : dayHourToNum(dailyHours.Tue),
          wednesdayHours: isMonthly ? 0 : dayHourToNum(dailyHours.Wed),
          thursdayHours: isMonthly ? 0 : dayHourToNum(dailyHours.Thu),
          fridayHours: isMonthly ? 0 : dayHourToNum(dailyHours.Fri),
          saturdayHours: isMonthly ? 0 : dayHourToNum(dailyHours.Sat),
          active: true,
          siteListItemId: editingSite.id,
          visitFrequency,
          hoursPerVisit: isMonthly && hoursPerVisit !== "" ? Number(hoursPerVisit) : (isFortnightly || visitFrequency === "Weekly" ? 0 : undefined),
          week2SundayHours: isFortnightly ? dayHourToNum(dailyHoursWeek2.Sun) : 0,
          week2MondayHours: isFortnightly ? dayHourToNum(dailyHoursWeek2.Mon) : 0,
          week2TuesdayHours: isFortnightly ? dayHourToNum(dailyHoursWeek2.Tue) : 0,
          week2WednesdayHours: isFortnightly ? dayHourToNum(dailyHoursWeek2.Wed) : 0,
          week2ThursdayHours: isFortnightly ? dayHourToNum(dailyHoursWeek2.Thu) : 0,
          week2FridayHours: isFortnightly ? dayHourToNum(dailyHoursWeek2.Fri) : 0,
          week2SaturdayHours: isFortnightly ? dayHourToNum(dailyHoursWeek2.Sat) : 0,
          ...(weekdayLabourRate !== "" && weekdayLabourRate != null && !Number.isNaN(Number(weekdayLabourRate)) && Number(weekdayLabourRate) >= 0 && { weekdayLabourRate: Number(weekdayLabourRate) }),
          ...(fortnightCostBudget !== "" && fortnightCostBudget != null && !Number.isNaN(Number(fortnightCostBudget)) && { fortnightCostBudget: Number(fortnightCostBudget) }),
          ...(saturdayLabourRate !== "" && saturdayLabourRate != null && !Number.isNaN(Number(saturdayLabourRate)) && Number(saturdayLabourRate) >= 0 && { saturdayLabourRate: Number(saturdayLabourRate) }),
          ...(sundayLabourRate !== "" && sundayLabourRate != null && !Number.isNaN(Number(sundayLabourRate)) && Number(sundayLabourRate) >= 0 && { sundayLabourRate: Number(sundayLabourRate) }),
          ...(phLabourRate !== "" && phLabourRate != null && !Number.isNaN(Number(phLabourRate)) && Number(phLabourRate) >= 0 && { phLabourRate: Number(phLabourRate) }),
        };
        if (budget?.budgetListItemId) {
          await updateSiteBudget(token, budget.budgetListItemId, budgetPayload);
        } else {
          await createSiteBudget(token, {
            budgetName: `${form.siteName.trim()} Budget`,
            siteListItemId: editingSite.id,
            sundayHours: budgetPayload.sundayHours,
            mondayHours: budgetPayload.mondayHours,
            tuesdayHours: budgetPayload.tuesdayHours,
            wednesdayHours: budgetPayload.wednesdayHours,
            thursdayHours: budgetPayload.thursdayHours,
            fridayHours: budgetPayload.fridayHours,
            saturdayHours: budgetPayload.saturdayHours,
            active: budgetPayload.active,
            visitFrequency: budgetPayload.visitFrequency,
            ...(budgetPayload.hoursPerVisit !== undefined && { hoursPerVisit: budgetPayload.hoursPerVisit }),
            week2SundayHours: budgetPayload.week2SundayHours,
            week2MondayHours: budgetPayload.week2MondayHours,
            week2TuesdayHours: budgetPayload.week2TuesdayHours,
            week2WednesdayHours: budgetPayload.week2WednesdayHours,
            week2ThursdayHours: budgetPayload.week2ThursdayHours,
            week2FridayHours: budgetPayload.week2FridayHours,
            week2SaturdayHours: budgetPayload.week2SaturdayHours,
            ...(budgetPayload.fortnightCostBudget !== undefined && { fortnightCostBudget: budgetPayload.fortnightCostBudget }),
            ...(budgetPayload.weekdayLabourRate !== undefined && { weekdayLabourRate: budgetPayload.weekdayLabourRate }),
            ...(budgetPayload.saturdayLabourRate !== undefined && { saturdayLabourRate: budgetPayload.saturdayLabourRate }),
            ...(budgetPayload.sundayLabourRate !== undefined && { sundayLabourRate: budgetPayload.sundayLabourRate }),
            ...(budgetPayload.phLabourRate !== undefined && { phLabourRate: budgetPayload.phLabourRate }),
          });
        }
        const currentAssignments = await getAssignedManagersForSite(token, editingSite.id);
        const currentIds = new Set(currentAssignments.map((a) => a.managerId));
        const selectedSet = new Set(selectedManagerIds);
        for (const a of currentAssignments) {
          if (!selectedSet.has(a.managerId)) {
            try {
              await deleteSiteManagerAssignment(token, a.itemId);
            } catch (err) {
              console.warn("Remove manager assignment failed:", err);
            }
          }
        }
        const managerErrors: string[] = [];
        for (const managerId of selectedManagerIds) {
          if (!currentIds.has(managerId)) {
            try {
              const managerFullName = managers.find((m) => m.id === managerId)?.fullName ?? "";
              await createSiteManagerAssignment(token, editingSite.id, managerId, {
                siteName: form.siteName.trim(),
                managerFullName,
              });
            } catch (err) {
              const msg = err instanceof Error ? err.message : String(err);
              console.warn("Add manager assignment failed:", err);
              managerErrors.push(`${managerId}: ${msg}`);
            }
          }
        }
        if (managerErrors.length > 0) {
          showToast(`Site updated but manager assignment failed: ${managerErrors.join("; ")}`);
        } else {
          showToast("Site updated.");
        }
      }
      closeModal();
      await loadSites(true);
      onUpdateSite?.(true);
    } catch (e) {
      const msg = e instanceof Error ? e.message : "Save failed.";
      const isPermissionError =
        msg.includes("write permission") || msg.includes("403") || msg.includes("401");
      setSubmitError(
        isPermissionError
          ? `Graph write permission missing. ${WRITE_PERMISSION_HINT}`
          : msg
      );
    } finally {
      setSubmitLoading(false);
    }
  };

  const retryWithNewToken = async () => {
    const account = msalInstance.getAllAccounts()[0];
    if (!account) {
      setSubmitError("Please sign in first using the main sign-in.");
      return;
    }
    setSubmitError(null);
    try {
      await msalInstance.acquireTokenPopup({ scopes: GRAPH_SCOPES });
      showToast("Token refreshed. Retrying save…");
      await handleSubmit();
    } catch (err) {
      if (err instanceof InteractionRequiredAuthError || (err && typeof (err as Error).message === "string")) {
        setSubmitError("Sign-in was cancelled or failed. " + WRITE_PERMISSION_HINT);
      } else {
        setSubmitError(err instanceof Error ? err.message : "Failed to get new token.");
      }
    }
  };

  const handleSetActive = async (site: Site, active: boolean) => {
    const token = await getGraphAccessToken();
    if (!token) return;
    setSubmitError(null);
    try {
      await setSiteActive(token, site.id, active);
      showToast(active ? "Site activated." : "Site deactivated.");
      await loadSites(true);
      onUpdateSite?.(true);
    } catch (e) {
      const msg = e instanceof Error ? e.message : "Update failed.";
      const isPermissionError =
        msg.includes("write permission") || msg.includes("403") || msg.includes("401");
      showToast(isPermissionError ? `Permission missing. ${WRITE_PERMISSION_HINT}` : msg);
    }
  };

  const handleDeleteSite = async (site: Site) => {
    if (!window.confirm(`Delete "${site.siteName || "Unnamed site"}"? The site will be removed from the list (moved to recycle bin). Related budgets and timesheet entries may still reference it.`)) return;
    const token = await getGraphAccessToken();
    if (!token) {
      showToast("Sign in with Microsoft to delete sites.");
      return;
    }
    try {
      await deleteSite(token, site.id);
      showToast("Site deleted.");
      await loadSites(true);
      onUpdateSite?.(true);
    } catch (e) {
      const msg = e instanceof Error ? e.message : "Delete failed.";
      showToast(msg.includes("403") || msg.includes("401") ? `Permission missing. ${WRITE_PERMISSION_HINT}` : msg);
    }
  };

  const selectedSet = React.useMemo(() => new Set(selectedSiteIds), [selectedSiteIds]);
  const handleRemoveManagerFromSite = useCallback(
    async (siteName: string, assignmentItemId: string) => {
      const token = await getGraphAccessToken();
      if (!token) {
        showToast("Sign in with Microsoft to remove managers.");
        return;
      }
      try {
        await deleteSiteManagerAssignment(token, assignmentItemId);
        showToast("Manager removed from site.");
        await loadSites(true);
        onUpdateSite?.(true);
      } catch (e) {
        const msg = e instanceof Error ? e.message : "Failed to remove manager.";
        showToast(msg);
      }
    },
    [loadSites, onUpdateSite, showToast]
  );

  const selectedManagerIdsSet = useMemo(
    () => new Set(selectedManagerIds),
    [selectedManagerIds]
  );
  const selectAllFiltered = () => {
    if (selectedSet.size === sortedSites.length) {
      setSelectedSiteIds([]);
    } else {
      setSelectedSiteIds(sortedSites.map((s) => s.id));
    }
  };

  const handleBulkDelete = async () => {
    if (selectedSiteIds.length === 0) return;
    if (
      !window.confirm(
        `Delete ${selectedSiteIds.length} selected site(s)? Sites will be removed from the list (moved to recycle bin). Related budgets and timesheet entries may still reference them.`
      )
    ) return;
    const token = await getGraphAccessToken();
    if (!token) {
      showToast("Sign in with Microsoft to delete sites.");
      return;
    }
    setBulkDeleteLoading(true);
    setError(null);
    let deleted = 0;
    let failed = 0;
    try {
      for (const id of selectedSiteIds) {
        try {
          await deleteSite(token, id);
          deleted++;
        } catch {
          failed++;
        }
      }
      setSelectedSiteIds([]);
      await loadSites(true);
      onUpdateSite?.(true);
      if (failed > 0) {
        showToast(`Deleted ${deleted} site(s). ${failed} failed.`);
      } else {
        showToast(deleted === 1 ? "Site deleted." : `Deleted ${deleted} sites.`);
      }
    } catch (e) {
      const msg = e instanceof Error ? e.message : "Bulk delete failed.";
      showToast(msg.includes("403") || msg.includes("401") ? `Permission missing. ${WRITE_PERMISSION_HINT}` : msg);
    } finally {
      setBulkDeleteLoading(false);
    }
  };

  return (
    <div className="space-y-6 sm:space-y-8">
      <div className="flex flex-col sm:flex-row sm:justify-between sm:items-end gap-4">
        <div className="min-w-0">
          <h2 className="text-[20px] sm:text-[22px] font-semibold text-gray-900">
            Sites &amp; Budgets
          </h2>
          <p className="text-gray-500 text-[13px] mt-1">
            Sites &amp; Budgets — Configure service windows and financial caps per site.
          </p>
        </div>
        {isAdmin && (
          <div className="flex flex-col sm:flex-row sm:items-center gap-1.5 w-full sm:w-auto flex-shrink-0">
            <button
              onClick={openBulkModal}
              className="w-full sm:w-auto justify-center so-btn-secondary px-3 py-2 text-sm font-medium flex items-center gap-1.5 touch-manipulation"
            >
              <Layers size={16} /> Bulk Add
            </button>
            <button
              onClick={openAdd}
              className="w-full sm:w-auto justify-center so-btn-primary px-3 py-2 text-sm font-medium flex items-center gap-1.5 touch-manipulation shadow-sm"
            >
              <Plus size={16} /> New Site
            </button>
          </div>
        )}
      </div>

      <div className="relative">
        <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" size={18} />
        <input
          type="search"
          placeholder="Search sites by name or address…"
          value={siteSearchQuery}
          onChange={(e) => setSiteSearchQuery(e.target.value)}
          className="w-full pl-10 pr-4 py-2.5 so-input text-sm text-gray-900 placeholder-gray-400 bg-white"
          aria-label="Search sites"
        />
      </div>

      {isAdmin && selectedSiteIds.length > 0 && (
        <div className="hidden md:flex sticky top-12 z-20 flex-wrap items-center gap-2 py-2 px-3 bg-amber-50 border border-amber-200 rounded-lg shadow-sm">
          <span className="text-sm font-medium text-amber-800">
            {selectedSiteIds.length} selected
          </span>
          <button
            type="button"
            onClick={handleBulkDelete}
            disabled={bulkDeleteLoading}
            className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-bold bg-red-600 text-white hover:bg-red-700 disabled:opacity-50"
          >
            {bulkDeleteLoading ? <Loader2 className="animate-spin" size={14} /> : <Trash2 size={14} />}
            Delete selected
          </button>
          <button
            type="button"
            onClick={() => setSelectedSiteIds([])}
            className="text-xs font-medium text-amber-800 hover:text-amber-900"
          >
            Clear selection
          </button>
        </div>
      )}

      {toast && (
        <div className="bg-green-50 border border-green-200 text-green-800 px-4 py-2 rounded-lg text-sm">
          {toast}
        </div>
      )}

      {loading ? (
        <div className="flex items-center gap-2 text-gray-500">
          <Loader2 className="animate-spin" size={20} /> Loading sites…
        </div>
      ) : error ? (
        <div className="bg-amber-50 border border-amber-200 text-amber-800 px-4 py-3 rounded-lg">
          {error}
        </div>
      ) : filteredSites.length === 0 ? (
        <div className="text-gray-500 py-8">
          {siteSearchQuery.trim() ? "No sites match your search." : "No sites found."}
        </div>
      ) : (
        <>
        <div className="md:hidden space-y-2">
          {sortedSites.map((site) => {
            const budget = budgetsBySiteId[String(site.id)];
            const fortnightCap = budget?.fortnightCap ?? 0;
            const labourRates = effectiveBudgetLabourRates(budget);
            return (
              <div
                key={site.id}
                className="border border-[#edeef0] rounded-lg bg-white px-3 py-3"
              >
                <div className="flex items-start justify-between gap-2">
                  <div className="min-w-0">
                    {onViewSite ? (
                      <button
                        type="button"
                        onClick={() => onViewSite(site.id)}
                        className="text-left text-sm font-bold text-gray-900 break-words hover:underline"
                      >
                        {site.siteName || "Unnamed site"}
                      </button>
                    ) : (
                      <span className="text-sm font-bold text-gray-900 break-words">
                        {site.siteName || "Unnamed site"}
                      </span>
                    )}
                    <p className="text-[11px] text-gray-500 mt-1">
                      {site.state || "—"} · {fortnightCap}h cap
                    </p>
                    {isAdmin && site.monthlyRevenue != null && (
                      <p className="text-[11px] text-gray-500 mt-0.5">
                        Revenue:{" "}
                        {new Intl.NumberFormat("en-AU", {
                          style: "currency",
                          currency: "AUD",
                          minimumFractionDigits: 2,
                          maximumFractionDigits: 2,
                        }).format(Number(site.monthlyRevenue))}
                      </p>
                    )}
                    {budget && (
                      <div className="mt-1.5">
                        <LabourRatesInlineCell rates={labourRates} />
                      </div>
                    )}
                  </div>
                  {isAdmin && (
                    <div className="flex items-center gap-1 shrink-0">
                      <button
                        onClick={() => openEdit(site)}
                        className="touch-target p-2 rounded text-blue-600 hover:text-blue-800 hover:bg-blue-50 inline-flex items-center justify-center"
                        aria-label={`Edit ${site.siteName}`}
                        title="Edit"
                      >
                        <Edit3 size={18} />
                      </button>
                      <button
                        onClick={() => handleSetActive(site, !site.active)}
                        className="touch-target p-2 rounded text-gray-600 hover:text-gray-900 hover:bg-gray-100 inline-flex items-center justify-center"
                        aria-label={site.active ? `Deactivate ${site.siteName}` : `Activate ${site.siteName}`}
                        title={site.active ? "Deactivate" : "Activate"}
                      >
                        {site.active ? <UserMinus size={18} /> : <UserPlus size={18} />}
                      </button>
                      <button
                        onClick={() => handleDeleteSite(site)}
                        className="touch-target p-2 rounded text-red-600 hover:text-red-800 hover:bg-red-50 inline-flex items-center justify-center"
                        aria-label={`Delete ${site.siteName}`}
                        title="Delete site"
                      >
                        <Trash2 size={18} />
                      </button>
                    </div>
                  )}
                </div>
              </div>
            );
          })}
        </div>
        <div className="hidden md:block so-table bg-white table-scroll-mobile">
          <table className="w-full border-collapse min-w-[860px] table-auto md:table-fixed md:min-w-[1180px] text-center">
            <colgroup className="hidden md:contents">
              {isAdmin && <col style={{ width: '4%' }} />}
              <col style={{ width: isAdmin ? '15%' : '18%' }} />
              <col style={{ width: '5%' }} />
              {isAdmin && <col style={{ width: '8%' }} />}
              <col style={{ width: '8%' }} />
              <col style={{ width: isAdmin ? '24%' : '28%' }} />
              <col style={{ width: isAdmin ? '22%' : '26%' }} />
              <col style={{ width: isAdmin ? '8%' : '15%' }} />
              {isAdmin && <col style={{ width: '6%' }} />}
            </colgroup>
            <thead>
              <tr className="border-b border-[#edeef0]">
                {isAdmin && (
                  <th className="px-2 py-2 md:px-1.5 md:py-1.5 w-10 align-middle">
                    <div className="flex justify-center">
                      <input
                        type="checkbox"
                        checked={sortedSites.length > 0 && selectedSet.size === sortedSites.length}
                        onChange={selectAllFiltered}
                        className="rounded border-gray-300"
                        aria-label="Select all"
                      />
                    </div>
                  </th>
                )}
                <th className="px-2 py-2 md:px-1.5 md:py-1.5 align-middle">
                  <button
                    type="button"
                    onClick={() => handleSiteSort("name")}
                    className="text-[10px] font-semibold text-gray-700 uppercase tracking-widest inline-flex items-center justify-center gap-0.5 mx-auto hover:text-gray-900 focus:outline-none focus:ring-2 focus:ring-gray-900/20 rounded"
                  >
                    Site name
                    {siteSortBy === "name" && (siteSortDir === "asc" ? <ChevronUp size={12} /> : <ChevronDown size={12} />)}
                  </button>
                </th>
                <th className="hidden md:table-cell px-1.5 py-1.5 align-middle">
                  <button
                    type="button"
                    onClick={() => handleSiteSort("state")}
                    className="text-[10px] font-semibold text-gray-700 uppercase tracking-widest inline-flex items-center justify-center gap-0.5 mx-auto hover:text-gray-900 focus:outline-none focus:ring-2 focus:ring-gray-900/20 rounded"
                  >
                    State
                    {siteSortBy === "state" && (siteSortDir === "asc" ? <ChevronUp size={12} /> : <ChevronDown size={12} />)}
                  </button>
                </th>
                {isAdmin && (
                  <th className="hidden md:table-cell px-1.5 py-1.5 text-center align-middle">
                    <button
                      type="button"
                      onClick={() => handleSiteSort("monthlyRevenue")}
                      className="text-[10px] font-semibold text-gray-700 uppercase tracking-widest inline-flex items-center justify-center gap-0.5 hover:text-gray-900 focus:outline-none focus:ring-2 focus:ring-gray-900/20 rounded"
                    >
                      Monthly revenue
                      {siteSortBy === "monthlyRevenue" && (siteSortDir === "asc" ? <ChevronUp size={12} /> : <ChevronDown size={12} />)}
                    </button>
                  </th>
                )}
                <th className="hidden md:table-cell px-1.5 py-1.5 text-[10px] font-semibold text-gray-700 uppercase tracking-widest text-center align-middle">
                  Assigned managers
                </th>
                <th className="px-2 py-2 md:px-1.5 md:py-1.5 text-[10px] font-semibold text-gray-700 uppercase tracking-widest text-center align-middle">
                  Daily hours
                </th>
                <th className="hidden md:table-cell px-1.5 py-1.5 text-[10px] font-semibold text-gray-700 uppercase tracking-widest align-middle">
                  Labour rates
                </th>
                <th className="px-2 py-2 md:px-1.5 md:py-1.5 text-center align-middle">
                  <button
                    type="button"
                    onClick={() => handleSiteSort("fortnightlyCap")}
                    className="text-[10px] font-semibold text-gray-700 uppercase tracking-widest inline-flex items-center justify-center gap-0.5 hover:text-gray-900 focus:outline-none focus:ring-2 focus:ring-gray-900/20 rounded"
                  >
                    Fortnight cap
                    {siteSortBy === "fortnightlyCap" && (siteSortDir === "asc" ? <ChevronUp size={12} /> : <ChevronDown size={12} />)}
                  </button>
                </th>
                {isAdmin && (
                  <th className="px-2 py-2 md:py-1.5 md:pl-2 md:pr-4 text-[10px] font-semibold text-gray-700 uppercase tracking-widest align-middle">
                    Actions
                  </th>
                )}
              </tr>
            </thead>
            <tbody className="divide-y divide-[#edeef0]">
          {sortedSites.map((site) => {
            const budget = budgetsBySiteId[String(site.id)];
            const fortnightCap = budget?.fortnightCap ?? 0;
            const labourRates = effectiveBudgetLabourRates(budget);

            return (
              <tr key={site.id} className="transition-colors">
                {isAdmin && (
                  <td className="px-2 py-2 md:px-1.5 md:py-1.5 align-middle">
                    <div className="flex justify-center">
                      <input
                        type="checkbox"
                        checked={selectedSet.has(site.id)}
                        onChange={() => {
                          setSelectedSiteIds((prev) =>
                            prev.includes(site.id) ? prev.filter((id) => id !== site.id) : [...prev, site.id]
                          );
                        }}
                        className="rounded border-gray-300"
                        aria-label={`Select ${site.siteName || "Unnamed site"}`}
                      />
                    </div>
                  </td>
                )}
                <td className="px-2 py-2 md:px-1.5 md:py-1.5 align-middle min-w-[120px] text-center">
                  {onViewSite ? (
                    <button
                      type="button"
                      onClick={() => onViewSite(site.id)}
                      className="block w-full text-sm md:text-xs font-bold text-gray-900 break-words hover:underline text-center"
                    >
                      {site.siteName || "Unnamed site"}
                    </button>
                  ) : (
                    <span className="block w-full text-sm md:text-xs font-bold text-gray-900 break-words text-center">
                      {site.siteName || "Unnamed site"}
                    </span>
                  )}
                </td>
                <td className="hidden md:table-cell px-1.5 py-1.5 align-middle text-center whitespace-nowrap">
                  <span className="text-[11px] font-medium text-gray-700 inline-block">{site.state || "—"}</span>
                </td>
                {isAdmin && (
                  <td className="hidden md:table-cell px-1.5 py-1.5 align-middle text-center whitespace-nowrap">
                    {site.monthlyRevenue != null ? (
                      <span className="text-[11px] font-medium text-gray-700 inline-block">
                        {new Intl.NumberFormat("en-AU", {
                          style: "currency",
                          currency: "AUD",
                          minimumFractionDigits: 2,
                          maximumFractionDigits: 2,
                        }).format(Number(site.monthlyRevenue))}
                      </span>
                    ) : (
                      <span className="text-[10px] text-gray-400">—</span>
                    )}
                  </td>
                )}
                  <td className="hidden md:table-cell px-1.5 py-1.5 align-middle text-center">
                  {(() => {
                    const { assignedManagers } = assignedManagersBySiteId[site.id] ?? { assignedManagers: [] };
                    if (assignedManagers.length === 0) {
                      return <span className="text-[10px] text-gray-400">No assigned managers</span>;
                    }
                    return (
                      <div className="flex flex-wrap gap-1 items-center justify-center">
                        {assignedManagers.map((a, i) => {
                          const name = a.managerName || "";
                          const parts = name.trim().split(/\s+/);
                          const initials =
                            parts.length === 1
                              ? (parts[0][0] ?? "").toUpperCase()
                              : `${(parts[0][0] ?? "").toUpperCase()}${(parts[parts.length - 1][0] ?? "").toUpperCase()}`;
                          return (
                            <button
                              key={a.assignmentItemId ?? i}
                              type="button"
                              onClick={() => handleRemoveManagerFromSite(site.siteName, a.assignmentItemId)}
                              className="w-7 h-7 rounded-md bg-gray-100 text-[10px] font-bold text-gray-700 flex items-center justify-center border border-gray-200 hover:bg-red-50 hover:text-red-600 hover:border-red-200"
                              aria-label={`Remove ${a.managerName} from site`}
                              title={`${a.managerName} – click to remove`}
                            >
                              {initials || "?"}
                            </button>
                          );
                        })}
                      </div>
                    );
                  })()}
                </td>
                <td className="px-2 py-2 md:pl-1.5 md:pr-5 md:py-1.5 align-middle text-center md:min-w-[16rem]">
                  {budget ? (
                    <SiteBudgetDayHoursGrid budget={budget} />
                  ) : (
                    <span className="text-[10px] text-gray-400">—</span>
                  )}
                </td>
                <td className="hidden md:table-cell py-1.5 align-middle md:min-w-[19rem] px-3 md:pl-6 md:pr-4">
                  <div className="flex justify-center w-full">
                    <LabourRatesInlineCell rates={labourRates} />
                  </div>
                </td>
                <td className="px-2 py-2 md:px-1.5 md:py-1.5 align-middle text-center whitespace-nowrap">
                  <span className="text-sm md:text-xs font-bold text-gray-900 inline-block">{fortnightCap}h</span>
                </td>
                {isAdmin && (
                  <td className="px-2 py-2 md:py-1.5 md:pl-2 md:pr-4 align-middle whitespace-nowrap">
                    <div className="flex items-center justify-center gap-0 flex-nowrap">
                      <button
                        onClick={() => openEdit(site)}
                        className="p-1 rounded text-blue-600 hover:text-blue-800 hover:bg-blue-50 inline-flex items-center justify-center min-w-[28px] min-h-[28px]"
                        aria-label={`Edit ${site.siteName}`}
                        title="Edit"
                      >
                        <Edit3 size={15} className="shrink-0" strokeWidth={2} />
                      </button>
                      <button
                        onClick={() => handleSetActive(site, !site.active)}
                        className="p-1 rounded text-gray-600 hover:text-gray-900 hover:bg-gray-100 inline-flex items-center justify-center min-w-[28px] min-h-[28px]"
                        aria-label={site.active ? `Deactivate ${site.siteName}` : `Activate ${site.siteName}`}
                        title={site.active ? "Deactivate" : "Activate"}
                      >
                        {site.active ? <UserMinus size={15} className="shrink-0" strokeWidth={2} /> : <UserPlus size={15} className="shrink-0" strokeWidth={2} />}
                      </button>
                      <button
                        onClick={() => handleDeleteSite(site)}
                        className="p-1 rounded text-red-600 hover:text-red-800 hover:bg-red-50 inline-flex items-center justify-center min-w-[28px] min-h-[28px]"
                        aria-label={`Delete ${site.siteName}`}
                        title="Delete site"
                      >
                        <Trash2 size={15} className="shrink-0" strokeWidth={2} />
                      </button>
                    </div>
                  </td>
                )}
              </tr>
            );
          })}
            </tbody>
          </table>
        </div>
        <p className="text-gray-400 text-xs flex items-center gap-1 mt-3">
          <span className="inline-block w-4 h-4 rounded bg-gray-200 flex items-center justify-center text-[10px]">i</span>
          Daily budgets are persistent across fortnight cycles. Click column headers to sort.
        </p>
        </>
      )}

      {modalMode && (
        <div
          className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 p-2 sm:p-4"
          onClick={closeModal}
        >
          <div
            className="bg-white rounded-xl shadow-xl w-full max-w-[96vw] sm:max-w-2xl mx-auto overflow-y-auto max-h-[90vh]"
            onClick={(e) => e.stopPropagation()}
          >
            <div className="flex justify-between items-center p-4 sm:p-6 border-b border-[#edeef0]">
              <h3 className="text-lg font-bold text-gray-900">
                {modalMode === "add" ? "Create New Site" : "Edit Site"}
              </h3>
              <button
                onClick={closeModal}
                className="text-gray-400 hover:text-gray-600 p-1"
                aria-label="Close"
              >
                <X size={20} />
              </button>
            </div>

            <div className="p-4 sm:p-6 space-y-4 sm:space-y-6">
              <div className="rounded-lg border border-[#edeef0] p-3 sm:p-4 bg-white">
                <h4 className="text-xs font-bold text-gray-500 uppercase tracking-widest mb-3">Basic Site Information</h4>
                <div className="space-y-4">
                  <div>
                    <label className="block text-xs font-bold text-gray-500 uppercase tracking-widest mb-1">
                      Site Name <span className="text-red-500">*</span>
                    </label>
                    <input
                      type="text"
                      value={form.siteName}
                      onChange={(e) => setForm((f) => ({ ...f, siteName: e.target.value }))}
                      className="w-full border border-[#edeef0] rounded-lg px-3 py-2 text-sm"
                      placeholder="e.g. Westfield Mall"
                    />
                  </div>
                  <div>
                    <label className="block text-xs font-bold text-gray-500 uppercase tracking-widest mb-1">
                      Site Address
                    </label>
                    <input
                      type="text"
                      value={form.address}
                      onChange={(e) => setForm((f) => ({ ...f, address: e.target.value }))}
                      className="w-full border border-[#edeef0] rounded-lg px-3 py-2 text-sm"
                      placeholder="Street, suburb"
                    />
                  </div>
                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                    <div>
                      <label className="block text-xs font-bold text-gray-500 uppercase tracking-widest mb-1">
                        $ Monthly Revenue
                      </label>
                      <input
                        type="number"
                        min={0}
                        step={0.01}
                        value={form.monthlyRevenue ?? ""}
                        onChange={(e) => {
                          const v = e.target.value;
                          setForm((f) => ({
                            ...f,
                            monthlyRevenue: v === "" ? null : parseFloat(v) || 0,
                          }));
                        }}
                        className="w-full border border-[#edeef0] rounded-lg px-3 py-2 text-sm"
                        placeholder="0"
                      />
                    </div>
                    <div>
                      <label className="block text-xs font-bold text-gray-500 uppercase tracking-widest mb-1">
                        $ Fortnight Cost Budget
                      </label>
                      <input
                        type="number"
                        min={0}
                        step={0.01}
                        value={fortnightCostBudget}
                        onChange={(e) => {
                          const v = e.target.value;
                          setFortnightCostBudget(v === "" ? "" : parseFloat(v) || 0);
                        }}
                        className="w-full border border-[#edeef0] rounded-lg px-3 py-2 text-sm"
                        placeholder="0"
                      />
                    </div>
                  </div>
                  <div>
                    <label className="block text-xs font-bold text-gray-500 uppercase tracking-widest mb-1">
                      State
                    </label>
                    <select
                      value={form.state}
                      onChange={(e) => setForm((f) => ({ ...f, state: e.target.value }))}
                      className="w-full border border-[#edeef0] rounded-lg px-3 py-2 text-sm"
                    >
                      <option value="">—</option>
                      {AU_STATES.map((s) => (
                        <option key={s} value={s}>{s}</option>
                      ))}
                    </select>
                  </div>
                  <div className="flex items-center gap-2">
                    <input
                      type="checkbox"
                      id="active"
                      checked={form.active !== false}
                      onChange={(e) => setForm((f) => ({ ...f, active: e.target.checked }))}
                      className="rounded border-gray-300"
                    />
                    <label htmlFor="active" className="text-sm text-gray-700">Active</label>
                  </div>
                </div>
              </div>

              <div className="rounded-lg border border-[#edeef0] p-3 sm:p-4 bg-white">
                <h4 className="text-xs font-bold text-gray-500 uppercase tracking-widest mb-3">Visit pattern</h4>
                <div className="space-y-3">
                  <div>
                    <label className="block text-xs font-bold text-gray-500 uppercase tracking-widest mb-1">
                      Visit frequency
                    </label>
                    <select
                      value={visitFrequency}
                      onChange={(e) => {
                        const next = e.target.value as VisitFreq;
                        setVisitFrequency(next);
                        if (next === "Fortnightly") {
                          if (sumDayHours(dailyHoursWeek2) === 0) setDailyHoursWeek2({ ...dailyHours });
                        }
                      }}
                      className="w-full border border-[#edeef0] rounded-lg px-3 py-2 text-sm"
                    >
                      <option value="Weekly">Weekly</option>
                      <option value="Fortnightly">Fortnightly</option>
                      <option value="Monthly">Monthly</option>
                    </select>
                  </div>
                  {visitFrequency === "Monthly" && (
                    <div>
                      <label className="block text-xs font-bold text-gray-500 uppercase tracking-widest mb-1">
                        Hours per visit
                      </label>
                      <input
                        type="number"
                        min={0}
                        step={0.5}
                        value={hoursPerVisit}
                        onChange={(e) => {
                          const v = e.target.value;
                          setHoursPerVisit(v === "" ? "" : parseFloat(v) || 0);
                        }}
                        className="w-full border border-[#edeef0] rounded-lg px-3 py-2 text-sm"
                        placeholder="e.g. 5"
                      />
                    </div>
                  )}
                  <h4 className="text-xs font-bold text-gray-500 uppercase tracking-widest mb-2">Labour rates ($/hr)</h4>
                  <p className="text-[10px] text-gray-400 mb-2">Used for budgeted labour cost, profit margin and dashboard. Weekday = Mon–Fri; PH = public holiday.</p>
                  <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
                    <div>
                      <label className="block text-xs font-bold text-gray-500 uppercase tracking-widest mb-1">Weekday ($/hr)</label>
                      <input
                        type="number"
                        min={0}
                        step={0.5}
                        value={weekdayLabourRate}
                        onChange={(e) => setWeekdayLabourRate(e.target.value === "" ? "" : parseFloat(e.target.value) || 0)}
                        className="w-full border border-[#edeef0] rounded-lg px-3 py-2 text-sm"
                        placeholder="Mon–Fri"
                      />
                    </div>
                    <div>
                      <label className="block text-xs font-bold text-gray-500 uppercase tracking-widest mb-1">Saturday ($/hr)</label>
                      <input
                        type="number"
                        min={0}
                        step={0.5}
                        value={saturdayLabourRate}
                        onChange={(e) => setSaturdayLabourRate(e.target.value === "" ? "" : parseFloat(e.target.value) || 0)}
                        className="w-full border border-[#edeef0] rounded-lg px-3 py-2 text-sm"
                        placeholder="Sat"
                      />
                    </div>
                    <div>
                      <label className="block text-xs font-bold text-gray-500 uppercase tracking-widest mb-1">Sunday ($/hr)</label>
                      <input
                        type="number"
                        min={0}
                        step={0.5}
                        value={sundayLabourRate}
                        onChange={(e) => setSundayLabourRate(e.target.value === "" ? "" : parseFloat(e.target.value) || 0)}
                        className="w-full border border-[#edeef0] rounded-lg px-3 py-2 text-sm"
                        placeholder="Sun"
                      />
                    </div>
                    <div>
                      <label className="block text-xs font-bold text-gray-500 uppercase tracking-widest mb-1">PH ($/hr)</label>
                      <input
                        type="number"
                        min={0}
                        step={0.5}
                        value={phLabourRate}
                        onChange={(e) => setPhLabourRate(e.target.value === "" ? "" : parseFloat(e.target.value) || 0)}
                        className="w-full border border-[#edeef0] rounded-lg px-3 py-2 text-sm"
                        placeholder="Public holiday"
                      />
                    </div>
                  </div>
                </div>
              </div>

              {(visitFrequency === "Weekly" || visitFrequency === "Fortnightly") && (
              <div className="rounded-lg border border-[#edeef0] p-3 sm:p-4 bg-white">
                <h4 className="text-xs font-bold text-gray-500 uppercase tracking-widest mb-3">Daily Service Hours</h4>
                {visitFrequency === "Weekly" && (
                <div className="grid grid-cols-7 gap-2">
                  {DAY_KEYS.map((day) => (
                    <div key={day}>
                      <label className="block text-[10px] font-bold text-gray-400 uppercase mb-1">{day}</label>
                      <input
                        type="number"
                        min={0}
                        step={0.5}
                        value={dailyHours[day]}
                        onChange={(e) => {
                          const v = e.target.value;
                          setDailyHours((h) => ({ ...h, [day]: v === "" ? "" : parseFloat(v) || 0 }));
                        }}
                        className="w-full border border-[#edeef0] rounded-lg px-2 py-1.5 text-sm"
                      />
                    </div>
                  ))}
                </div>
                )}
                {visitFrequency === "Fortnightly" && (
                <div className="space-y-4">
                  <div>
                    <p className="text-[10px] font-bold text-gray-500 uppercase mb-2">Week 1</p>
                    <div className="grid grid-cols-7 gap-2">
                      {DAY_KEYS.map((day) => (
                        <div key={"w1-" + day}>
                          <label className="block text-[10px] font-bold text-gray-400 uppercase mb-1">{day}</label>
                          <input
                            type="number"
                            min={0}
                            step={0.5}
                            value={dailyHours[day]}
                            onChange={(e) => {
                              const v = e.target.value;
                              setDailyHours((h) => ({ ...h, [day]: v === "" ? "" : parseFloat(v) || 0 }));
                            }}
                            className="w-full border border-[#edeef0] rounded-lg px-2 py-1.5 text-sm"
                          />
                        </div>
                      ))}
                    </div>
                  </div>
                  <div>
                    <p className="text-[10px] font-bold text-gray-500 uppercase mb-2">Week 2</p>
                    <div className="grid grid-cols-7 gap-2">
                      {DAY_KEYS.map((day) => (
                        <div key={"w2-" + day}>
                          <label className="block text-[10px] font-bold text-gray-400 uppercase mb-1">{day}</label>
                          <input
                            type="number"
                            min={0}
                            step={0.5}
                            value={dailyHoursWeek2[day]}
                            onChange={(e) => {
                              const v = e.target.value;
                              setDailyHoursWeek2((h) => ({ ...h, [day]: v === "" ? "" : parseFloat(v) || 0 }));
                            }}
                            className="w-full border border-[#edeef0] rounded-lg px-2 py-1.5 text-sm"
                          />
                        </div>
                      ))}
                    </div>
                  </div>
                </div>
                )}
              </div>
              )}

              <div className="rounded-lg border border-[#edeef0] p-3 sm:p-4 bg-white">
                <h4 className="text-xs font-bold text-gray-500 uppercase tracking-widest mb-3">No Service Periods</h4>
                <p className="text-[10px] text-gray-400 mb-2">
                  Used for school holidays, shutdowns, client closures, and seasonal pauses. Matching dates will show as 0h / On Target in Timesheets.
                </p>
                <div className="mb-3">
                  <label className="block text-[10px] font-bold text-gray-400 uppercase mb-1">No service mode</label>
                  <div className="inline-flex rounded-lg border border-[#edeef0] p-0.5 bg-gray-50/80">
                    <button
                      type="button"
                      onClick={() => {
                        setNoServiceMode("manual");
                        setNoServiceHolidayMessage(null);
                      }}
                      className={`px-3 py-1.5 text-xs font-bold rounded-md transition-colors ${
                        noServiceMode === "manual"
                          ? "bg-gray-900 text-white"
                          : "text-gray-600 hover:text-gray-900"
                      }`}
                    >
                      Manual
                    </button>
                    <button
                      type="button"
                      onClick={() => {
                        setNoServiceMode("school_auto");
                        setNoServiceHolidayMessage(null);
                      }}
                      className={`px-3 py-1.5 text-xs font-bold rounded-md transition-colors ${
                        noServiceMode === "school_auto"
                          ? "bg-gray-900 text-white"
                          : "text-gray-600 hover:text-gray-900"
                      }`}
                    >
                      School Holidays (Auto)
                    </button>
                  </div>
                </div>

                {noServiceMode === "school_auto" && (
                  <div className="mb-3 space-y-2 rounded-lg border border-[#edeef0] bg-gray-50/50 px-3 py-2">
                    <p className="text-[10px] text-gray-500">
                      Uses the site <strong>State</strong> and selected <strong>Year</strong> to populate school holiday no-service periods. Click Save to write them to the site.
                    </p>
                    <div className="flex flex-wrap items-end gap-3">
                      <div>
                        <span className="block text-[10px] font-bold text-gray-400 uppercase mb-1">State (from site)</span>
                        <span className="inline-block text-sm font-semibold text-gray-800 min-w-[3rem]">
                          {form.state.trim() ? form.state.trim().toUpperCase() : "—"}
                        </span>
                      </div>
                      <div>
                        <label className="block text-[10px] font-bold text-gray-400 uppercase mb-1">Year</label>
                        <select
                          value={noServiceYear}
                          onChange={(e) => {
                            setNoServiceYear(Number(e.target.value));
                            setNoServiceHolidayMessage(null);
                          }}
                          className="border border-[#edeef0] rounded-lg px-2 py-1.5 text-sm bg-white min-w-[5.5rem]"
                        >
                          {schoolHolidayYearOptions.map((y) => (
                            <option key={y} value={y}>
                              {y}
                            </option>
                          ))}
                        </select>
                      </div>
                      <button
                        type="button"
                        onClick={handlePopulateSchoolHolidays}
                        disabled={!form.state.trim()}
                        className="px-3 py-1.5 text-xs font-bold text-white bg-gray-900 rounded-lg hover:bg-gray-800 disabled:opacity-40 disabled:cursor-not-allowed"
                      >
                        Populate School Holidays
                      </button>
                    </div>
                    {!form.state.trim() && (
                      <p className="text-[10px] text-amber-700">Select <strong>State</strong> in the site details above before populating.</p>
                    )}
                    {noServiceHolidayMessage && (
                      <p
                        className={`text-[10px] ${
                          noServiceHolidayMessage.startsWith("Loaded")
                            ? "text-green-700"
                            : "text-amber-800"
                        }`}
                      >
                        {noServiceHolidayMessage}
                      </p>
                    )}
                  </div>
                )}

                <div className="space-y-2">
                  {noServicePeriods.map((row) => (
                    <div key={row.id} className="grid grid-cols-1 sm:grid-cols-12 gap-2 items-end">
                      <div className="sm:col-span-3">
                        <label className="block text-[10px] font-bold text-gray-400 uppercase mb-1">Label</label>
                        <input
                          type="text"
                          value={row.label}
                          readOnly={noServiceMode === "school_auto"}
                          onChange={(e) => {
                            const v = e.target.value;
                            setNoServicePeriods((prev) => prev.map((p) => (p.id === row.id ? { ...p, label: v } : p)));
                          }}
                          className="w-full border border-[#edeef0] rounded-lg px-2 py-1.5 text-sm read-only:bg-gray-50 read-only:text-gray-700"
                          placeholder="Term Holidays"
                        />
                      </div>
                      <div className="sm:col-span-3">
                        <label className="block text-[10px] font-bold text-gray-400 uppercase mb-1">Start Date</label>
                        <input
                          type="date"
                          value={row.startDate}
                          readOnly={noServiceMode === "school_auto"}
                          onChange={(e) => {
                            const v = e.target.value;
                            setNoServicePeriods((prev) => prev.map((p) => (p.id === row.id ? { ...p, startDate: v } : p)));
                          }}
                          className="w-full border border-[#edeef0] rounded-lg px-2 py-1.5 text-sm read-only:bg-gray-50"
                        />
                      </div>
                      <div className="sm:col-span-3">
                        <label className="block text-[10px] font-bold text-gray-400 uppercase mb-1">End Date</label>
                        <input
                          type="date"
                          value={row.endDate}
                          readOnly={noServiceMode === "school_auto"}
                          onChange={(e) => {
                            const v = e.target.value;
                            setNoServicePeriods((prev) => prev.map((p) => (p.id === row.id ? { ...p, endDate: v } : p)));
                          }}
                          className="w-full border border-[#edeef0] rounded-lg px-2 py-1.5 text-sm read-only:bg-gray-50"
                        />
                      </div>
                      <div className="sm:col-span-2">
                        <label className="block text-[10px] font-bold text-gray-400 uppercase mb-1">Reason</label>
                        <input
                          type="text"
                          value={row.reason}
                          readOnly={noServiceMode === "school_auto"}
                          onChange={(e) => {
                            const v = e.target.value;
                            setNoServicePeriods((prev) => prev.map((p) => (p.id === row.id ? { ...p, reason: v } : p)));
                          }}
                          className="w-full border border-[#edeef0] rounded-lg px-2 py-1.5 text-sm read-only:bg-gray-50"
                          placeholder="Optional"
                        />
                      </div>
                      <div className="sm:col-span-1">
                        <button
                          type="button"
                          disabled={noServiceMode === "school_auto"}
                          onClick={() => setNoServicePeriods((prev) => prev.filter((p) => p.id !== row.id))}
                          className="w-full px-2 py-1.5 text-xs font-bold text-gray-600 hover:bg-gray-100 rounded-lg border border-[#edeef0] disabled:opacity-40 disabled:cursor-not-allowed"
                          title="Remove period"
                        >
                          Remove
                        </button>
                      </div>
                    </div>
                  ))}
                  {noServiceMode === "manual" && (
                    <button
                      type="button"
                      onClick={() => setNoServicePeriods((prev) => [...prev, createNoServiceRow()])}
                      className="px-3 py-1.5 text-xs font-bold text-gray-700 hover:bg-gray-100 rounded-lg border border-[#edeef0]"
                    >
                      + Add Period
                    </button>
                  )}
                </div>
              </div>

              <div className="rounded-lg border border-[#edeef0] p-3 sm:p-4 bg-white">
                <h4 className="text-xs font-bold text-gray-500 uppercase tracking-widest mb-3">Personnel &amp; Rates</h4>
                <p className="text-xs text-gray-500 mb-2">Assign managers to this site.</p>
                <div className="space-y-2">
                  {managers.length === 0 && (
                    <p className="text-xs text-gray-400">No managers in CleanTrack Users (Role = Manager).</p>
                  )}
                  {managers.map((m) => {
                    const isChecked = selectedManagerIdsSet.has(m.id);
                    return (
                    <label key={m.id} className="flex items-center gap-2 cursor-pointer">
                      <input
                        type="checkbox"
                        checked={isChecked}
                        onChange={(e) => {
                          if (e.target.checked) {
                            setSelectedManagerIds((prev) => (prev.includes(m.id) ? prev : [...prev, m.id]));
                          } else {
                            setSelectedManagerIds((prev) => prev.filter((id) => id !== m.id));
                          }
                        }}
                        className="rounded border-gray-300"
                      />
                      <span className="text-sm text-gray-800">{m.fullName}</span>
                    </label>
                  ); })}
                </div>
              </div>
            </div>

            {submitError && (
              <div className="px-4 sm:px-6 pb-4 space-y-2">
                <div className="bg-amber-50 border border-amber-200 text-amber-800 px-3 py-2 rounded-lg text-sm">
                  {submitError}
                </div>
                {(submitError.includes("Graph write permission") || submitError.includes("403") || submitError.includes("401")) && (
                  <button
                    type="button"
                    onClick={retryWithNewToken}
                    className="text-sm font-medium text-amber-800 hover:text-amber-900 underline"
                  >
                    Sign in again for write access
                  </button>
                )}
              </div>
            )}

            <div className="p-4 sm:p-6 pt-0 flex flex-col-reverse sm:flex-row sm:justify-end gap-2">
              <button
                onClick={closeModal}
                className="w-full sm:w-auto px-4 py-2 text-sm font-medium text-gray-600 hover:bg-gray-100 rounded-lg"
              >
                Cancel
              </button>
              <button
                onClick={handleSubmit}
                disabled={submitLoading}
                className="w-full sm:w-auto so-btn-primary px-6 py-2.5 rounded-lg text-sm font-bold disabled:opacity-50 flex items-center justify-center gap-2"
              >
                {submitLoading && <Loader2 className="animate-spin" size={14} />}
                {modalMode === "add" ? "Launch New Site" : "Save"}
              </button>
            </div>
          </div>
        </div>
      )}

      {bulkModalOpen && (
        <div
          className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 p-2 sm:p-4"
          onClick={() => !bulkSubmitLoading && setBulkModalOpen(false)}
        >
          <div
            className="bg-white rounded-xl shadow-xl w-full max-w-[96vw] sm:max-w-[1600px] mx-auto overflow-y-auto max-h-[90vh]"
            onClick={(e) => e.stopPropagation()}
          >
            <div className="flex justify-between items-center p-4 sm:p-6 border-b border-[#edeef0]">
              <h3 className="text-lg font-bold text-gray-900">Bulk Add Sites</h3>
              <button
                onClick={() => !bulkSubmitLoading && setBulkModalOpen(false)}
                className="text-gray-400 hover:text-gray-600 p-1"
                aria-label="Close"
                disabled={bulkSubmitLoading}
              >
                <X size={20} />
              </button>
            </div>

            <div className="p-4 sm:p-6 space-y-6">
              <div>
                <h4 className="text-xs font-bold text-gray-500 uppercase tracking-widest mb-2">Sites</h4>
                <div className="md:hidden space-y-3">
                  {bulkRows.map((row, rowIdx) => (
                    <div key={row.id} className="border border-[#edeef0] rounded-lg p-3 space-y-3 bg-white">
                      <div className="flex items-center justify-between">
                        <p className="text-xs font-semibold text-gray-600 uppercase tracking-wider">
                          Site {rowIdx + 1}
                        </p>
                        <button
                          type="button"
                          onClick={() => removeBulkRow(row.id)}
                          disabled={bulkRows.length <= 1 || bulkSubmitLoading}
                          className="text-gray-400 hover:text-red-600 disabled:opacity-40 p-1"
                          aria-label="Remove row"
                        >
                          <Trash2 size={14} />
                        </button>
                      </div>

                      <div className="space-y-1">
                        <label className="text-[10px] font-bold text-gray-500 uppercase tracking-widest">
                          Site Name <span className="text-red-500">*</span>
                        </label>
                        <input
                          type="text"
                          value={row.siteName}
                          onChange={(e) => updateBulkRow(row.id, { siteName: e.target.value })}
                          placeholder="Site name"
                          className="w-full border border-[#edeef0] rounded px-2 py-2 text-sm"
                          disabled={bulkSubmitLoading}
                        />
                      </div>

                      <div className="space-y-1">
                        <label className="text-[10px] font-bold text-gray-500 uppercase tracking-widest">Address</label>
                        <input
                          type="text"
                          value={row.address}
                          onChange={(e) => updateBulkRow(row.id, { address: e.target.value })}
                          placeholder="Address"
                          className="w-full border border-[#edeef0] rounded px-2 py-2 text-sm"
                          disabled={bulkSubmitLoading}
                        />
                      </div>

                      <div className="grid grid-cols-2 gap-2">
                        <div className="space-y-1">
                          <label className="text-[10px] font-bold text-gray-500 uppercase tracking-widest">$ Mon Rev</label>
                          <input
                            type="number"
                            min={0}
                            step={0.01}
                            value={row.monthlyRevenue === "" ? "" : row.monthlyRevenue}
                            onChange={(e) => {
                              const v = e.target.value;
                              updateBulkRow(row.id, { monthlyRevenue: v === "" ? "" : parseFloat(v) || 0 });
                            }}
                            className="w-full border border-[#edeef0] rounded px-2 py-2 text-sm"
                            disabled={bulkSubmitLoading}
                          />
                        </div>
                        <div className="space-y-1">
                          <label className="text-[10px] font-bold text-gray-500 uppercase tracking-widest">$ Fort Budget</label>
                          <input
                            type="number"
                            min={0}
                            step={0.01}
                            value={row.fortnightCostBudget === "" ? "" : row.fortnightCostBudget}
                            onChange={(e) => {
                              const v = e.target.value;
                              updateBulkRow(row.id, { fortnightCostBudget: v === "" ? "" : parseFloat(v) || 0 });
                            }}
                            className="w-full border border-[#edeef0] rounded px-2 py-2 text-sm"
                            disabled={bulkSubmitLoading}
                          />
                        </div>
                      </div>

                      <div className="grid grid-cols-2 gap-2">
                        <div className="space-y-1">
                          <label className="text-[10px] font-bold text-gray-500 uppercase tracking-widest">$/hr</label>
                          <input
                            type="number"
                            min={0}
                            step={0.01}
                            value={row.weekdayLabourRate === "" ? "" : row.weekdayLabourRate}
                            onChange={(e) => {
                              const v = e.target.value;
                              updateBulkRow(row.id, { weekdayLabourRate: v === "" ? "" : parseFloat(v) || 0 });
                            }}
                            placeholder="—"
                            className="w-full border border-[#edeef0] rounded px-2 py-2 text-sm"
                            disabled={bulkSubmitLoading}
                          />
                        </div>
                        <div className="space-y-1">
                          <label className="text-[10px] font-bold text-gray-500 uppercase tracking-widest">State</label>
                          <select
                            value={row.state}
                            onChange={(e) => updateBulkRow(row.id, { state: e.target.value })}
                            className="w-full border border-[#edeef0] rounded px-2 py-2 text-sm"
                            disabled={bulkSubmitLoading}
                          >
                            <option value="">—</option>
                            {AU_STATES.map((s) => (
                              <option key={s} value={s}>{s}</option>
                            ))}
                          </select>
                        </div>
                      </div>

                      <div className="grid grid-cols-2 gap-2 items-end">
                        <div className="space-y-1">
                          <label className="text-[10px] font-bold text-gray-500 uppercase tracking-widest">Visit freq</label>
                          <select
                            value={row.visitFrequency}
                            onChange={(e) => updateBulkRow(row.id, { visitFrequency: e.target.value as VisitFreq })}
                            className="w-full border border-[#edeef0] rounded px-2 py-2 text-sm"
                            disabled={bulkSubmitLoading}
                          >
                            <option value="Weekly">Weekly</option>
                            <option value="Fortnightly">Fortnightly</option>
                            <option value="Monthly">Monthly</option>
                          </select>
                        </div>
                        <label className="inline-flex items-center gap-2 pb-2">
                          <input
                            type="checkbox"
                            checked={row.active}
                            onChange={(e) => updateBulkRow(row.id, { active: e.target.checked })}
                            className="rounded border-gray-300"
                            disabled={bulkSubmitLoading}
                          />
                          <span className="text-sm text-gray-700">Active</span>
                        </label>
                      </div>

                      {row.visitFrequency === "Monthly" ? (
                        <div className="space-y-1">
                          <label className="text-[10px] font-bold text-gray-500 uppercase tracking-widest">Hours per visit</label>
                          <input
                            type="number"
                            min={0}
                            step={0.5}
                            value={row.hoursPerVisit}
                            onChange={(e) => {
                              const v = e.target.value;
                              updateBulkRow(row.id, { hoursPerVisit: v === "" ? "" : parseFloat(v) || 0 });
                            }}
                            className="w-full border border-[#edeef0] rounded px-2 py-2 text-sm"
                            disabled={bulkSubmitLoading}
                          />
                        </div>
                      ) : (
                        <div className="space-y-2">
                          <p className="text-[10px] font-bold text-gray-500 uppercase tracking-widest">Week 1 Daily Hours</p>
                          <div className="grid grid-cols-3 gap-2">
                            {DAY_KEYS.map((day) => (
                              <label key={day} className="space-y-1">
                                <span className="text-[10px] text-gray-500">{day}</span>
                                <input
                                  type="number"
                                  min={0}
                                  step={0.5}
                                  value={row.dailyHours[day]}
                                  onChange={(e) => {
                                    const v = e.target.value;
                                    updateBulkRow(row.id, {
                                      dailyHours: { ...row.dailyHours, [day]: v === "" ? "" : parseFloat(v) || 0 },
                                    });
                                  }}
                                  className="w-full border border-[#edeef0] rounded px-2 py-1.5 text-sm"
                                  disabled={bulkSubmitLoading}
                                />
                              </label>
                            ))}
                          </div>
                        </div>
                      )}

                      {row.visitFrequency === "Fortnightly" && (
                        <div className="space-y-2">
                          <p className="text-[10px] font-bold text-gray-500 uppercase tracking-widest">Week 2 Daily Hours</p>
                          <div className="grid grid-cols-3 gap-2">
                            {DAY_KEYS.map((day) => (
                              <label key={"mobile-w2-" + day} className="space-y-1">
                                <span className="text-[10px] text-gray-500">{day}</span>
                                <input
                                  type="number"
                                  min={0}
                                  step={0.5}
                                  value={row.dailyHoursWeek2[day]}
                                  onChange={(e) => {
                                    const v = e.target.value;
                                    updateBulkRow(row.id, {
                                      dailyHoursWeek2: { ...row.dailyHoursWeek2, [day]: v === "" ? "" : parseFloat(v) || 0 },
                                    });
                                  }}
                                  className="w-full border border-[#edeef0] rounded px-2 py-1.5 text-sm"
                                  disabled={bulkSubmitLoading}
                                />
                              </label>
                            ))}
                          </div>
                        </div>
                      )}
                    </div>
                  ))}
                </div>

                <div className="hidden md:block border border-[#edeef0] rounded-lg table-scroll-mobile">
                  <table className="w-full min-w-[1180px] text-left text-sm">
                    <thead>
                      <tr className="bg-gray-50 border-b border-[#edeef0] text-[10px] font-bold text-gray-500 uppercase tracking-widest">
                        <th className="py-2 px-1 w-8"></th>
                        <th className="py-2 px-1 min-w-[120px]">Site Name <span className="text-red-500">*</span></th>
                        <th className="py-2 px-1 min-w-[120px]">Address</th>
                        <th className="py-2 px-1 w-20">$ Mon Rev</th>
                        <th className="py-2 px-1 w-20">$ Fort Budget</th>
                        <th className="py-2 px-1 w-16">$/hr</th>
                        <th className="py-2 px-1 w-16">State</th>
                        <th className="py-2 px-1 w-12 text-center">Active</th>
                        <th className="py-2 px-1 w-20">Visit freq</th>
                        {DAY_KEYS.map((d) => (
                          <th key={d} className="py-2 px-0.5 w-10 text-center">{d}</th>
                        ))}
                        {DAY_KEYS.map((d) => (
                          <th key={"w2-" + d} className="py-2 px-0.5 w-10 text-center">{d}</th>
                        ))}
                        <th className="py-2 px-1 w-14 text-center">Hrs/visit</th>
                      </tr>
                    </thead>
                    <tbody>
                      {bulkRows.map((row) => (
                        <tr key={row.id} className="border-b border-[#edeef0] last:border-b-0 hover:bg-gray-50/50">
                          <td className="py-1 px-1">
                            <button
                              type="button"
                              onClick={() => removeBulkRow(row.id)}
                              disabled={bulkRows.length <= 1 || bulkSubmitLoading}
                              className="text-gray-400 hover:text-red-600 disabled:opacity-40 p-0.5"
                              aria-label="Remove row"
                            >
                              <Trash2 size={14} />
                            </button>
                          </td>
                          <td className="py-1 px-1">
                            <input
                              type="text"
                              value={row.siteName}
                              onChange={(e) => updateBulkRow(row.id, { siteName: e.target.value })}
                              placeholder="Site name"
                              className="w-full border border-[#edeef0] rounded px-1.5 py-1 text-sm min-w-0"
                              disabled={bulkSubmitLoading}
                            />
                          </td>
                          <td className="py-1 px-1">
                            <input
                              type="text"
                              value={row.address}
                              onChange={(e) => updateBulkRow(row.id, { address: e.target.value })}
                              placeholder="Address"
                              className="w-full border border-[#edeef0] rounded px-1.5 py-1 text-sm min-w-0"
                              disabled={bulkSubmitLoading}
                            />
                          </td>
                          <td className="py-1 px-1">
                            <input
                              type="number"
                              min={0}
                              step={0.01}
                              value={row.monthlyRevenue === "" ? "" : row.monthlyRevenue}
                              onChange={(e) => {
                                const v = e.target.value;
                                updateBulkRow(row.id, { monthlyRevenue: v === "" ? "" : parseFloat(v) || 0 });
                              }}
                              className="w-full border border-[#edeef0] rounded px-1 py-1 text-sm"
                              disabled={bulkSubmitLoading}
                            />
                          </td>
                          <td className="py-1 px-1">
                            <input
                              type="number"
                              min={0}
                              step={0.01}
                              value={row.fortnightCostBudget === "" ? "" : row.fortnightCostBudget}
                              onChange={(e) => {
                                const v = e.target.value;
                                updateBulkRow(row.id, { fortnightCostBudget: v === "" ? "" : parseFloat(v) || 0 });
                              }}
                              className="w-full border border-[#edeef0] rounded px-1 py-1 text-sm"
                              disabled={bulkSubmitLoading}
                            />
                          </td>
                          <td className="py-1 px-1">
                            <input
                              type="number"
                              min={0}
                              step={0.01}
                              value={row.weekdayLabourRate === "" ? "" : row.weekdayLabourRate}
                              onChange={(e) => {
                                const v = e.target.value;
                                updateBulkRow(row.id, { weekdayLabourRate: v === "" ? "" : parseFloat(v) || 0 });
                              }}
                              placeholder="—"
                              className="w-full border border-[#edeef0] rounded px-1 py-1 text-sm"
                              disabled={bulkSubmitLoading}
                            />
                          </td>
                          <td className="py-1 px-1">
                            <select
                              value={row.state}
                              onChange={(e) => updateBulkRow(row.id, { state: e.target.value })}
                              className="w-full border border-[#edeef0] rounded px-1 py-1 text-sm"
                              disabled={bulkSubmitLoading}
                            >
                              <option value="">—</option>
                              {AU_STATES.map((s) => (
                                <option key={s} value={s}>{s}</option>
                              ))}
                            </select>
                          </td>
                          <td className="py-1 px-1 text-center">
                            <input
                              type="checkbox"
                              checked={row.active}
                              onChange={(e) => updateBulkRow(row.id, { active: e.target.checked })}
                              className="rounded border-gray-300"
                              disabled={bulkSubmitLoading}
                            />
                          </td>
                          <td className="py-1 px-1">
                            <select
                              value={row.visitFrequency}
                              onChange={(e) => updateBulkRow(row.id, { visitFrequency: e.target.value as VisitFreq })}
                              className="w-full border border-[#edeef0] rounded px-1 py-1 text-sm"
                              disabled={bulkSubmitLoading}
                            >
                              <option value="Weekly">Weekly</option>
                              <option value="Fortnightly">Fortnightly</option>
                              <option value="Monthly">Monthly</option>
                            </select>
                          </td>
                          {DAY_KEYS.map((day) => (
                            <td key={day} className="py-1 px-0.5">
                              <input
                                type="number"
                                min={0}
                                step={0.5}
                                value={row.visitFrequency === "Monthly" ? "" : row.dailyHours[day]}
                                onChange={(e) => {
                                  const v = e.target.value;
                                  updateBulkRow(row.id, {
                                    dailyHours: { ...row.dailyHours, [day]: v === "" ? "" : parseFloat(v) || 0 },
                                  });
                                }}
                                disabled={bulkSubmitLoading || row.visitFrequency === "Monthly"}
                                className="w-full border border-[#edeef0] rounded px-0.5 py-1 text-xs text-center"
                              />
                            </td>
                          ))}
                          {DAY_KEYS.map((day) => (
                            <td key={"w2-" + day} className="py-1 px-0.5">
                              <input
                                type="number"
                                min={0}
                                step={0.5}
                                value={row.visitFrequency === "Fortnightly" ? row.dailyHoursWeek2[day] : ""}
                                onChange={(e) => {
                                  const v = e.target.value;
                                  updateBulkRow(row.id, {
                                    dailyHoursWeek2: { ...row.dailyHoursWeek2, [day]: v === "" ? "" : parseFloat(v) || 0 },
                                  });
                                }}
                                disabled={bulkSubmitLoading || row.visitFrequency !== "Fortnightly"}
                                className="w-full border border-[#edeef0] rounded px-0.5 py-1 text-xs text-center"
                              />
                            </td>
                          ))}
                          <td className="py-1 px-1">
                            <input
                              type="number"
                              min={0}
                              step={0.5}
                              value={row.visitFrequency === "Monthly" ? row.hoursPerVisit : ""}
                              onChange={(e) => {
                                const v = e.target.value;
                                updateBulkRow(row.id, { hoursPerVisit: v === "" ? "" : parseFloat(v) || 0 });
                              }}
                              disabled={bulkSubmitLoading || row.visitFrequency !== "Monthly"}
                              placeholder="—"
                              className="w-full border border-[#edeef0] rounded px-1 py-1 text-sm text-center"
                            />
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
                <button
                  type="button"
                  onClick={addBulkRow}
                  disabled={bulkSubmitLoading}
                  className="mt-2 flex items-center gap-2 text-sm font-medium text-gray-700 hover:text-gray-900"
                >
                  <Plus size={16} /> Add new line
                </button>
                <p className="text-[10px] text-gray-400 mt-1">
                  {bulkRows.filter((r) => r.siteName.trim()).length} site(s) will be created (rows with a name).
                </p>
              </div>

              <div>
                <h4 className="text-xs font-bold text-gray-500 uppercase tracking-widest mb-3">Assign managers to all sites</h4>
                <div className="space-y-2">
                  {managers.map((m) => (
                    <label key={m.id} className="flex items-center gap-2 cursor-pointer">
                      <input
                        type="checkbox"
                        checked={bulkManagerIds.includes(m.id)}
                        onChange={(e) => {
                          if (e.target.checked) setBulkManagerIds((prev) => [...prev, m.id]);
                          else setBulkManagerIds((prev) => prev.filter((id) => id !== m.id));
                        }}
                        className="rounded border-gray-300"
                        disabled={bulkSubmitLoading}
                      />
                      <span className="text-sm text-gray-800">{m.fullName}</span>
                    </label>
                  ))}
                  {managers.length === 0 && (
                    <p className="text-xs text-gray-400">No managers in CleanTrack Users.</p>
                  )}
                </div>
              </div>

              {bulkError && (
                <div className="bg-amber-50 border border-amber-200 text-amber-800 px-3 py-2 rounded-lg text-sm">
                  {bulkError}
                </div>
              )}

              {bulkProgress && (
                <div className="flex items-center gap-2 text-sm text-gray-600">
                  <Loader2 className="animate-spin" size={18} />
                  Adding site {bulkProgress.current} of {bulkProgress.total}…
                </div>
              )}
            </div>

            <div className="p-4 sm:p-6 pt-0 flex flex-col-reverse sm:flex-row sm:justify-end gap-2">
              <button
                onClick={() => !bulkSubmitLoading && setBulkModalOpen(false)}
                className="w-full sm:w-auto px-4 py-2 text-sm font-medium text-gray-600 hover:bg-gray-100 rounded-lg disabled:opacity-50"
                disabled={bulkSubmitLoading}
              >
                Cancel
              </button>
              <button
                onClick={handleBulkSubmit}
                disabled={bulkSubmitLoading || bulkRows.filter((r) => r.siteName.trim()).length === 0}
                className="w-full sm:w-auto so-btn-primary px-6 py-2.5 rounded-lg text-sm font-bold disabled:opacity-50 flex items-center justify-center gap-2"
              >
                {bulkSubmitLoading && <Loader2 className="animate-spin" size={14} />}
                Add {bulkRows.filter((r) => r.siteName.trim()).length || 0} site{bulkRows.filter((r) => r.siteName.trim()).length !== 1 ? "s" : ""}
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

export default SiteManager;
