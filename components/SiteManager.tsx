import React, { useState, useEffect, useCallback, useRef, useMemo } from "react";
import { Plus, Edit3, X, MapPin, Loader2, Trash2, Layers, Search, ChevronUp, ChevronDown, UserMinus, UserPlus } from "lucide-react";
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
} from "../repositories/sitesRepo";
import { getAssignedSiteIdsForManager, createSiteManagerAssignment, getAssignedManagersForSite, deleteSiteManagerAssignment, fetchSiteManagerAssignments, joinAssignmentsToSites } from "../repositories/siteManagersRepo";
import { getCleanTrackManagers } from "../repositories/usersRepo";
import { createSiteBudget, getSiteBudgets, updateSiteBudget, type SiteBudgetHours } from "../repositories/budgetsRepo";
import { getSiteCleanerAssignments } from "../repositories/assignedCleanersRepo";
import { normalizeListItemId } from "../lib/sharepoint";

const AU_STATES = ["ACT", "NSW", "NT", "QLD", "SA", "TAS", "VIC", "WA"];
const DAY_LABELS = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"] as const;
type DayKey = (typeof DAY_LABELS)[number];
/** Display order: Monday first, Sunday last. */
const DAY_KEYS: DayKey[] = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"];

type VisitFreq = "Weekly" | "Fortnightly" | "Monthly";

const DEFAULT_DAY_HOURS: Record<DayKey, number> = { Sun: 0, Mon: 0, Tue: 0, Wed: 0, Thu: 0, Fri: 0, Sat: 0 };

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
  dailyHours: Record<DayKey, number>;
  dailyHoursWeek2: Record<DayKey, number>;
  hoursPerVisit: number | "";
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
  const [dailyHours, setDailyHours] = useState<Record<DayKey, number>>({
    Sun: 0, Mon: 0, Tue: 0, Wed: 0, Thu: 0, Fri: 0, Sat: 0,
  });
  const [dailyHoursWeek2, setDailyHoursWeek2] = useState<Record<DayKey, number>>({
    Sun: 0, Mon: 0, Tue: 0, Wed: 0, Thu: 0, Fri: 0, Sat: 0,
  });
  const [fortnightCostBudget, setFortnightCostBudget] = useState<number | "">("");
  const [selectedManagerIds, setSelectedManagerIds] = useState<string[]>([]);
  const [visitFrequency, setVisitFrequency] = useState<VisitFreq>("Weekly");
  const [hoursPerVisit, setHoursPerVisit] = useState<number | "">("");
  const [weekdayLabourRate, setWeekdayLabourRate] = useState<number | "">("");
  const [saturdayLabourRate, setSaturdayLabourRate] = useState<number | "">("");
  const [sundayLabourRate, setSundayLabourRate] = useState<number | "">("");
  const [phLabourRate, setPhLabourRate] = useState<number | "">("");

  const [bulkModalOpen, setBulkModalOpen] = useState(false);
  const [bulkRows, setBulkRows] = useState<BulkSiteRow[]>([]);
  const [bulkManagerIds, setBulkManagerIds] = useState<string[]>([]);
  const [bulkSubmitLoading, setBulkSubmitLoading] = useState(false);
  const [bulkProgress, setBulkProgress] = useState<{ current: number; total: number } | null>(null);
  const [bulkError, setBulkError] = useState<string | null>(null);
  const [siteSearchQuery, setSiteSearchQuery] = useState("");
  type SiteSortKey = "name" | "address" | "state" | "fortnightlyCap";
  const [siteSortBy, setSiteSortBy] = useState<SiteSortKey>("name");
  const [siteSortDir, setSiteSortDir] = useState<"asc" | "desc">("asc");
  const [selectedSiteIds, setSelectedSiteIds] = useState<string[]>([]);
  const [bulkDeleteLoading, setBulkDeleteLoading] = useState(false);
  const editingSiteIdRef = useRef<string | null>(null);

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
      setSiteSortDir(key === "fortnightlyCap" ? "desc" : "asc");
    }
  };

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
    setDailyHours({ Sun: 0, Mon: 0, Tue: 0, Wed: 0, Thu: 0, Fri: 0, Sat: 0 });
    setDailyHoursWeek2({ Sun: 0, Mon: 0, Tue: 0, Wed: 0, Thu: 0, Fri: 0, Sat: 0 });
    setFortnightCostBudget("");
    setVisitFrequency("Weekly");
    setHoursPerVisit("");
    setWeekdayLabourRate("");
    setSaturdayLabourRate("");
    setSundayLabourRate("");
    setPhLabourRate("");
    setSelectedManagerEmails([]);
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
            Sun: budget.sunday,
            Mon: budget.monday,
            Tue: budget.tuesday,
            Wed: budget.wednesday,
            Thu: budget.thursday,
            Fri: budget.friday,
            Sat: budget.saturday,
          }
        : { Sun: 0, Mon: 0, Tue: 0, Wed: 0, Thu: 0, Fri: 0, Sat: 0 }
    );
    setDailyHoursWeek2(
      budget && budget.week2Sunday != null
        ? {
            Sun: budget.week2Sunday,
            Mon: budget.week2Monday ?? 0,
            Tue: budget.week2Tuesday ?? 0,
            Wed: budget.week2Wednesday ?? 0,
            Thu: budget.week2Thursday ?? 0,
            Fri: budget.week2Friday ?? 0,
            Sat: budget.week2Saturday ?? 0,
          }
        : { Sun: 0, Mon: 0, Tue: 0, Wed: 0, Thu: 0, Fri: 0, Sat: 0 }
    );
    setFortnightCostBudget(budget?.fortnightCostBudget != null ? budget.fortnightCostBudget : "");
    setSelectedManagerEmails([]);
    const freq = budget?.visitFrequency;
    const normalizedFreq: VisitFreq =
      freq === "Fortnightly" || freq === "Monthly" ? freq : "Weekly";
    setVisitFrequency(normalizedFreq);
    setHoursPerVisit(
      budget?.hoursPerVisit != null && budget.hoursPerVisit > 0 ? budget.hoursPerVisit : ""
    );
    setWeekdayLabourRate(
      budget?.weekdayLabourRate != null && budget.weekdayLabourRate >= 0 ? budget.weekdayLabourRate : (budget?.budgetLabourRate != null && budget.budgetLabourRate > 0 ? budget.budgetLabourRate : "")
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
    dailyHours: { ...DEFAULT_DAY_HOURS },
    dailyHoursWeek2: { ...DEFAULT_DAY_HOURS },
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
            sundayHours: isMonthly ? 0 : row.dailyHours.Sun,
            mondayHours: isMonthly ? 0 : row.dailyHours.Mon,
            tuesdayHours: isMonthly ? 0 : row.dailyHours.Tue,
            wednesdayHours: isMonthly ? 0 : row.dailyHours.Wed,
            thursdayHours: isMonthly ? 0 : row.dailyHours.Thu,
            fridayHours: isMonthly ? 0 : row.dailyHours.Fri,
            saturdayHours: isMonthly ? 0 : row.dailyHours.Sat,
            active: true,
            visitFrequency: row.visitFrequency,
            ...(isMonthly && row.hoursPerVisit !== "" && { hoursPerVisit: Number(row.hoursPerVisit) }),
            ...(isFortnightly && {
              week2SundayHours: row.dailyHoursWeek2.Sun,
              week2MondayHours: row.dailyHoursWeek2.Mon,
              week2TuesdayHours: row.dailyHoursWeek2.Tue,
              week2WednesdayHours: row.dailyHoursWeek2.Wed,
              week2ThursdayHours: row.dailyHoursWeek2.Thu,
              week2FridayHours: row.dailyHoursWeek2.Fri,
              week2SaturdayHours: row.dailyHoursWeek2.Sat,
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
        });
        const siteId = created.id;
        const isMonthly = visitFrequency === "Monthly";
        const isFortnightly = visitFrequency === "Fortnightly";
        try {
          await createSiteBudget(token, {
            budgetName: `${form.siteName.trim()} Budget`,
            siteListItemId: siteId,
            sundayHours: isMonthly ? 0 : dailyHours.Sun,
            mondayHours: isMonthly ? 0 : dailyHours.Mon,
            tuesdayHours: isMonthly ? 0 : dailyHours.Tue,
            wednesdayHours: isMonthly ? 0 : dailyHours.Wed,
            thursdayHours: isMonthly ? 0 : dailyHours.Thu,
            fridayHours: isMonthly ? 0 : dailyHours.Fri,
            saturdayHours: isMonthly ? 0 : dailyHours.Sat,
            active: true,
            visitFrequency,
            ...(isMonthly && hoursPerVisit !== "" && { hoursPerVisit: Number(hoursPerVisit) }),
            ...(isFortnightly && {
              week2SundayHours: dailyHoursWeek2.Sun,
              week2MondayHours: dailyHoursWeek2.Mon,
              week2TuesdayHours: dailyHoursWeek2.Tue,
              week2WednesdayHours: dailyHoursWeek2.Wed,
              week2ThursdayHours: dailyHoursWeek2.Thu,
              week2FridayHours: dailyHoursWeek2.Fri,
              week2SaturdayHours: dailyHoursWeek2.Sat,
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
        });
        const budget = budgetsBySiteId[String(editingSite.id)] ?? budgetsBySiteId["name:" + (editingSite.siteName.trim() + " Budget")];
        const isMonthly = visitFrequency === "Monthly";
        const isFortnightly = visitFrequency === "Fortnightly";
        const budgetPayload = {
          sundayHours: isMonthly ? 0 : dailyHours.Sun,
          mondayHours: isMonthly ? 0 : dailyHours.Mon,
          tuesdayHours: isMonthly ? 0 : dailyHours.Tue,
          wednesdayHours: isMonthly ? 0 : dailyHours.Wed,
          thursdayHours: isMonthly ? 0 : dailyHours.Thu,
          fridayHours: isMonthly ? 0 : dailyHours.Fri,
          saturdayHours: isMonthly ? 0 : dailyHours.Sat,
          active: true,
          siteListItemId: editingSite.id,
          visitFrequency,
          hoursPerVisit: isMonthly && hoursPerVisit !== "" ? Number(hoursPerVisit) : (isFortnightly || visitFrequency === "Weekly" ? 0 : undefined),
          week2SundayHours: isFortnightly ? dailyHoursWeek2.Sun : 0,
          week2MondayHours: isFortnightly ? dailyHoursWeek2.Mon : 0,
          week2TuesdayHours: isFortnightly ? dailyHoursWeek2.Tue : 0,
          week2WednesdayHours: isFortnightly ? dailyHoursWeek2.Wed : 0,
          week2ThursdayHours: isFortnightly ? dailyHoursWeek2.Thu : 0,
          week2FridayHours: isFortnightly ? dailyHoursWeek2.Fri : 0,
          week2SaturdayHours: isFortnightly ? dailyHoursWeek2.Sat : 0,
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
          <div className="flex items-center gap-2 flex-shrink-0">
            <button
              onClick={openBulkModal}
              className="so-btn-secondary px-4 py-2.5 text-sm font-medium flex items-center gap-2 touch-manipulation"
            >
              <Layers size={16} /> Bulk Add
            </button>
            <button
              onClick={openAdd}
              className="so-btn-primary px-4 py-2.5 text-sm font-medium flex items-center gap-2 touch-manipulation shadow-sm"
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
        <div className="sticky top-12 z-20 flex flex-wrap items-center gap-2 py-2 px-3 bg-amber-50 border border-amber-200 rounded-lg shadow-sm">
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
        <div className="so-table bg-white overflow-hidden table-scroll-mobile">
          <table className="w-full border-collapse text-left min-w-0 table-auto md:table-fixed md:min-w-[900px]">
            <colgroup className="hidden md:contents">
              {isAdmin && <col style={{ width: '4%' }} />}
              <col style={{ width: isAdmin ? '16%' : '18%' }} />
              <col style={{ width: isAdmin ? '16%' : '20%' }} />
              <col style={{ width: '6%' }} />
              <col style={{ width: '11%' }} />
              <col style={{ width: '10%' }} />
              <col style={{ width: isAdmin ? '19%' : '21%' }} />
              <col style={{ width: '8%' }} />
              <col style={{ width: isAdmin ? '6%' : '8%' }} />
              {isAdmin && <col style={{ width: '10%' }} />}
            </colgroup>
            <thead>
              <tr className="border-b border-[#edeef0]">
                {isAdmin && (
                  <th className="px-2 py-2 md:px-1.5 md:py-1.5 w-10">
                    <input
                      type="checkbox"
                      checked={sortedSites.length > 0 && selectedSet.size === sortedSites.length}
                      onChange={selectAllFiltered}
                      className="rounded border-gray-300"
                      aria-label="Select all"
                    />
                  </th>
                )}
                <th className="px-2 py-2 md:px-1.5 md:py-1.5 text-left">
                  <button
                    type="button"
                    onClick={() => handleSiteSort("name")}
                    className="text-[10px] font-semibold text-gray-700 uppercase tracking-widest flex items-center gap-0.5 hover:text-gray-900 focus:outline-none focus:ring-2 focus:ring-gray-900/20 rounded"
                  >
                    Site name
                    {siteSortBy === "name" && (siteSortDir === "asc" ? <ChevronUp size={12} /> : <ChevronDown size={12} />)}
                  </button>
                </th>
                <th className="hidden md:table-cell px-1.5 py-1.5 text-left">
                  <button
                    type="button"
                    onClick={() => handleSiteSort("address")}
                    className="text-[10px] font-semibold text-gray-700 uppercase tracking-widest flex items-center gap-0.5 hover:text-gray-900 focus:outline-none focus:ring-2 focus:ring-gray-900/20 rounded"
                  >
                    Address
                    {siteSortBy === "address" && (siteSortDir === "asc" ? <ChevronUp size={12} /> : <ChevronDown size={12} />)}
                  </button>
                </th>
                <th className="hidden md:table-cell px-1.5 py-1.5 text-left">
                  <button
                    type="button"
                    onClick={() => handleSiteSort("state")}
                    className="text-[10px] font-semibold text-gray-700 uppercase tracking-widest flex items-center gap-0.5 hover:text-gray-900 focus:outline-none focus:ring-2 focus:ring-gray-900/20 rounded"
                  >
                    State
                    {siteSortBy === "state" && (siteSortDir === "asc" ? <ChevronUp size={12} /> : <ChevronDown size={12} />)}
                  </button>
                </th>
                <th className="hidden md:table-cell px-1.5 py-1.5 text-[10px] font-semibold text-gray-700 uppercase tracking-widest">Assigned managers</th>
                <th className="hidden md:table-cell px-1.5 py-1.5 text-[10px] font-semibold text-gray-700 uppercase tracking-widest">Assigned cleaners</th>
                <th className="px-2 py-2 md:px-1.5 md:py-1.5 text-[10px] font-semibold text-gray-700 uppercase tracking-widest">Daily hours</th>
                <th className="hidden md:table-cell px-1.5 py-1.5 text-[10px] font-semibold text-gray-700 uppercase tracking-widest">Weekday rate</th>
                <th className="px-2 py-2 md:px-1.5 md:py-1.5">
                  <button
                    type="button"
                    onClick={() => handleSiteSort("fortnightlyCap")}
                    className="text-[10px] font-semibold text-gray-700 uppercase tracking-widest flex items-center gap-0.5 hover:text-gray-900 focus:outline-none focus:ring-2 focus:ring-gray-900/20 rounded"
                  >
                    Fortnight cap
                    {siteSortBy === "fortnightlyCap" && (siteSortDir === "asc" ? <ChevronUp size={12} /> : <ChevronDown size={12} />)}
                  </button>
                </th>
                {isAdmin && (
                  <th className="px-2 py-2 md:px-1.5 md:py-1.5 text-[10px] font-semibold text-gray-700 uppercase tracking-widest text-right">Actions</th>
                )}
              </tr>
            </thead>
            <tbody className="divide-y divide-[#edeef0]">
          {sortedSites.map((site) => {
            const budget = budgetsBySiteId[String(site.id)];
            const dayHours = budget
              ? [
                  { day: "Mon" as const, h: budget.monday },
                  { day: "Tue" as const, h: budget.tuesday },
                  { day: "Wed" as const, h: budget.wednesday },
                  { day: "Thu" as const, h: budget.thursday },
                  { day: "Fri" as const, h: budget.friday },
                  { day: "Sat" as const, h: budget.saturday },
                  { day: "Sun" as const, h: budget.sunday },
                ]
              : [];
            const fortnightCap = budget?.fortnightCap ?? 0;
            const budgetRate = budget?.weekdayLabourRate ?? budget?.budgetLabourRate; // budgetLabourRate = legacy column name

            return (
              <tr key={site.id} className="transition-colors">
                {isAdmin && (
                  <td className="px-2 py-2 md:px-1.5 md:py-1.5 align-top">
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
                  </td>
                )}
                <td className="px-2 py-2 md:px-1.5 md:py-1.5 align-top min-w-[120px]">
                  {onViewSite ? (
                    <button
                      type="button"
                      onClick={() => onViewSite(site.id)}
                      className="text-left text-sm md:text-xs font-bold text-gray-900 break-words hover:underline"
                    >
                      {site.siteName || "Unnamed site"}
                    </button>
                  ) : (
                    <span className="text-sm md:text-xs font-bold text-gray-900 break-words">
                      {site.siteName || "Unnamed site"}
                    </span>
                  )}
                </td>
                <td className="hidden md:table-cell px-1.5 py-1.5 align-top">
                  {site.address ? (
                    <span className="text-[11px] text-gray-600 flex items-center gap-0.5 break-words">
                      <MapPin size={10} className="text-gray-400 shrink-0 flex-shrink-0" />
                      {site.address}
                    </span>
                  ) : (
                    <span className="text-[11px] text-gray-400">—</span>
                  )}
                </td>
                <td className="hidden md:table-cell px-1.5 py-1.5 align-top whitespace-nowrap">
                  <span className="text-[11px] font-medium text-gray-700">{site.state || "—"}</span>
                </td>
                <td className="hidden md:table-cell px-1.5 py-1.5 align-top">
                  {(() => {
                    const { assignedManagers } = assignedManagersBySiteId[site.id] ?? { assignedManagers: [] };
                    if (assignedManagers.length === 0) {
                      return <span className="text-[10px] text-gray-400">No assigned managers</span>;
                    }
                    return (
                      <div className="flex flex-wrap gap-1 items-center">
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
                <td className="hidden md:table-cell px-1.5 py-1.5 align-top">
                  {(() => {
                    const cleanerNames = assignedCleanersBySiteId[normalizeListItemId(site.id)] ?? [];
                    if (cleanerNames.length === 0) {
                      return <span className="text-[10px] text-gray-400">—</span>;
                    }
                    const primaryName = cleanerNames[0]?.trim() || "";
                    const hasMore = cleanerNames.length > 1;
                    return (
                      <span className="text-[11px] font-medium text-gray-700 break-words" title={hasMore ? cleanerNames.join(", ") : undefined}>
                        {primaryName}
                        {hasMore ? (
                          <span className="text-[10px] text-gray-500">, +{cleanerNames.length - 1} more</span>
                        ) : null}
                      </span>
                    );
                  })()}
                </td>
                <td className="px-2 py-2 md:px-1.5 md:py-1.5 align-top">
                  {dayHours.length > 0 ? (
                    <div className="grid grid-cols-7 gap-1 md:gap-px bg-[#edeef0] rounded border border-[#edeef0] overflow-hidden max-w-full md:max-w-[200px]">
                      {dayHours.map(({ day, h }) => (
                        <div
                          key={day}
                          className={`bg-white px-1.5 py-1 md:px-0.5 md:py-0.5 text-center min-w-0 ${
                            (h ?? 0) > 0 ? "text-blue-700" : "text-gray-400"
                          }`}
                          title={`${day}: ${h ?? 0}h`}
                        >
                          <span className="block text-[9px] md:text-[8px] font-medium uppercase leading-tight text-gray-500">{day.length > 2 ? day.charAt(0) : day}</span>
                          <span className="block text-xs md:text-[10px] font-bold tabular-nums mt-0.5">{(h ?? 0) > 0 ? Number(h) : "0"}</span>
                        </div>
                      ))}
                    </div>
                  ) : (
                    <span className="text-[10px] text-gray-400">—</span>
                  )}
                </td>
                <td className="hidden md:table-cell px-1.5 py-1.5 align-top">
                  {budgetRate != null && budgetRate > 0 ? (
                    <span className="text-[11px] font-medium text-gray-700">${Number(budgetRate).toFixed(2)}/hr</span>
                  ) : (
                    <span className="text-[10px] text-gray-400">—</span>
                  )}
                </td>
                <td className="px-2 py-2 md:px-1.5 md:py-1.5 align-top whitespace-nowrap">
                  <span className="text-sm md:text-xs font-bold text-gray-900">{fortnightCap}h</span>
                </td>
                {isAdmin && (
                  <td className="px-2 py-2 md:px-1.5 md:py-1.5 align-top text-right whitespace-nowrap">
                    <div className="flex items-center justify-end gap-1 flex-wrap">
                      <button
                        onClick={() => openEdit(site)}
                        className="touch-target p-2.5 sm:p-1.5 rounded text-blue-600 hover:text-blue-800 hover:bg-blue-50 inline-flex items-center justify-center"
                        aria-label={`Edit ${site.siteName}`}
                        title="Edit"
                      >
                        <Edit3 size={18} className="sm:w-3.5 sm:h-3.5 w-[18px] h-[18px]" />
                      </button>
                      <button
                        onClick={() => handleSetActive(site, !site.active)}
                        className="touch-target p-2.5 sm:p-1.5 rounded text-gray-600 hover:text-gray-900 hover:bg-gray-100 inline-flex items-center justify-center"
                        aria-label={site.active ? `Deactivate ${site.siteName}` : `Activate ${site.siteName}`}
                        title={site.active ? "Deactivate" : "Activate"}
                      >
                        {site.active ? <UserMinus size={18} className="sm:w-3.5 sm:h-3.5 w-[18px] h-[18px]" /> : <UserPlus size={18} className="sm:w-3.5 sm:h-3.5 w-[18px] h-[18px]" />}
                      </button>
                      <button
                        onClick={() => handleDeleteSite(site)}
                        className="touch-target p-2.5 sm:p-1.5 rounded text-red-600 hover:text-red-800 hover:bg-red-50 inline-flex items-center justify-center"
                        aria-label={`Delete ${site.siteName}`}
                        title="Delete site"
                      >
                        <Trash2 size={18} className="sm:w-3.5 sm:h-3.5 w-[18px] h-[18px]" />
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
          className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 p-4"
          onClick={closeModal}
        >
          <div
            className="bg-white rounded-xl shadow-xl max-w-2xl w-full mx-auto overflow-y-auto max-h-[90vh]"
            onClick={(e) => e.stopPropagation()}
          >
            <div className="flex justify-between items-center p-6 border-b border-[#edeef0]">
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

            <div className="p-6 space-y-6">
              <div>
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
                  <div className="grid grid-cols-2 gap-4">
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

              <div>
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
                          const week2Sum = DAY_KEYS.reduce((s, d) => s + dailyHoursWeek2[d], 0);
                          if (week2Sum === 0) setDailyHoursWeek2({ ...dailyHours });
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
              <div>
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
                        onChange={(e) =>
                          setDailyHours((h) => ({ ...h, [day]: parseFloat(e.target.value) || 0 }))
                        }
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
                            onChange={(e) =>
                              setDailyHours((h) => ({ ...h, [day]: parseFloat(e.target.value) || 0 }))
                            }
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
                            onChange={(e) =>
                              setDailyHoursWeek2((h) => ({ ...h, [day]: parseFloat(e.target.value) || 0 }))
                            }
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

              <div>
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
              <div className="px-6 pb-4 space-y-2">
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

            <div className="p-6 pt-0 flex justify-end gap-2">
              <button
                onClick={closeModal}
                className="px-4 py-2 text-sm font-medium text-gray-600 hover:bg-gray-100 rounded-lg"
              >
                Cancel
              </button>
              <button
                onClick={handleSubmit}
                disabled={submitLoading}
                className="bg-gray-900 text-white px-6 py-2.5 rounded-lg text-sm font-bold hover:bg-gray-800 disabled:opacity-50 flex items-center gap-2"
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
            className="bg-white rounded-xl shadow-xl w-full max-w-[1600px] mx-auto overflow-y-auto max-h-[90vh]"
            onClick={(e) => e.stopPropagation()}
          >
            <div className="flex justify-between items-center p-6 border-b border-[#edeef0]">
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

            <div className="p-6 space-y-6">
              <div>
                <h4 className="text-xs font-bold text-gray-500 uppercase tracking-widest mb-2">Sites</h4>
                <div className="border border-[#edeef0] rounded-lg overflow-x-auto">
                  <table className="w-full text-left text-sm">
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
                                onChange={(e) =>
                                  updateBulkRow(row.id, {
                                    dailyHours: { ...row.dailyHours, [day]: parseFloat(e.target.value) || 0 },
                                  })
                                }
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
                                onChange={(e) =>
                                  updateBulkRow(row.id, {
                                    dailyHoursWeek2: { ...row.dailyHoursWeek2, [day]: parseFloat(e.target.value) || 0 },
                                  })
                                }
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

            <div className="p-6 pt-0 flex justify-end gap-2">
              <button
                onClick={() => !bulkSubmitLoading && setBulkModalOpen(false)}
                className="px-4 py-2 text-sm font-medium text-gray-600 hover:bg-gray-100 rounded-lg disabled:opacity-50"
                disabled={bulkSubmitLoading}
              >
                Cancel
              </button>
              <button
                onClick={handleBulkSubmit}
                disabled={bulkSubmitLoading || bulkRows.filter((r) => r.siteName.trim()).length === 0}
                className="bg-gray-900 text-white px-6 py-2.5 rounded-lg text-sm font-bold hover:bg-gray-800 disabled:opacity-50 flex items-center gap-2"
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
