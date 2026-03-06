import React, { useState, useEffect, useCallback } from "react";
import { Plus, Edit3, X, MapPin, Loader2, Trash2, Layers } from "lucide-react";
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
import { getAssignedSiteIdsForManager, createSiteManagerAssignment, getAssignedManagersForSite, deleteSiteManagerAssignment } from "../repositories/siteManagersRepo";
import { getCleanTrackManagers } from "../repositories/usersRepo";
import { createSiteBudget, getSiteBudgets, updateSiteBudget, type SiteBudgetHours } from "../repositories/budgetsRepo";
import { getAssignedCleanersBySite } from "../repositories/assignedCleanersRepo";

const AU_STATES = ["ACT", "NSW", "NT", "QLD", "SA", "TAS", "VIC", "WA"];
const DAY_LABELS = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"] as const;
type DayKey = (typeof DAY_LABELS)[number];
const DAY_KEYS: DayKey[] = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];

type VisitFreq = "Weekly" | "Fortnightly" | "Monthly";

const DEFAULT_DAY_HOURS: Record<DayKey, number> = { Sun: 0, Mon: 0, Tue: 0, Wed: 0, Thu: 0, Fri: 0, Sat: 0 };

interface BulkSiteRow {
  id: string;
  siteName: string;
  address: string;
  monthlyRevenue: number | "";
  fortnightCostBudget: number | "";
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
  onUpdateSite?: () => void;
}

const SiteManager: React.FC<SiteManagerProps> = ({ onUpdateSite }) => {
  const { isAdmin: isAdminFromRole } = useRole();
  const { user } = useAppAuth();
  const isAdmin = isAdminFromRole || user?.role === "Admin";

  const [sites, setSites] = useState<Site[]>([]);
  const [budgetsBySiteId, setBudgetsBySiteId] = useState<Record<string, SiteBudgetHours>>({});
  const [assignedCleanersBySiteId, setAssignedCleanersBySiteId] = useState<Record<string, { name: string; payRate: number }[]>>({});
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
  const [managers, setManagers] = useState<{ fullName: string; email: string }[]>([]);
  const [dailyHours, setDailyHours] = useState<Record<DayKey, number>>({
    Sun: 0, Mon: 0, Tue: 0, Wed: 0, Thu: 0, Fri: 0, Sat: 0,
  });
  const [dailyHoursWeek2, setDailyHoursWeek2] = useState<Record<DayKey, number>>({
    Sun: 0, Mon: 0, Tue: 0, Wed: 0, Thu: 0, Fri: 0, Sat: 0,
  });
  const [fortnightCostBudget, setFortnightCostBudget] = useState<number | "">("");
  const [selectedManagerEmails, setSelectedManagerEmails] = useState<string[]>([]);
  const [visitFrequency, setVisitFrequency] = useState<VisitFreq>("Weekly");
  const [hoursPerVisit, setHoursPerVisit] = useState<number | "">("");

  const [bulkModalOpen, setBulkModalOpen] = useState(false);
  const [bulkRows, setBulkRows] = useState<BulkSiteRow[]>([]);
  const [bulkManagerEmails, setBulkManagerEmails] = useState<string[]>([]);
  const [bulkSubmitLoading, setBulkSubmitLoading] = useState(false);
  const [bulkProgress, setBulkProgress] = useState<{ current: number; total: number } | null>(null);
  const [bulkError, setBulkError] = useState<string | null>(null);

  const showToast = useCallback((msg: string) => {
    setToast(msg);
    setTimeout(() => setToast(null), 4000);
  }, []);

  const loadSites = useCallback(async () => {
    const token = await getGraphAccessToken();
    if (!token) {
      setError("Sign in with Microsoft to view sites.");
      setSites([]);
      setLoading(false);
      return;
    }
    setLoading(true);
    setError(null);
    try {
      let data = await getSites(token);
      if (!isAdmin && user?.role === "Manager" && user?.email) {
        const assignedIds = await getAssignedSiteIdsForManager(token, user.email);
        data = assignedIds.length > 0 ? data.filter((s) => assignedIds.includes(s.id)) : [];
      }
      setSites(data);
      const [budgets, cleanersMap] = await Promise.all([
        getSiteBudgets(token).catch(() => ({})),
        getAssignedCleanersBySite(token).catch(() => ({})),
      ]);
      setBudgetsBySiteId(budgets);
      setAssignedCleanersBySiteId(cleanersMap);
    } catch (e) {
      const msg = e instanceof Error ? e.message : "Failed to load sites.";
      setError(msg);
      setSites([]);
    } finally {
      setLoading(false);
    }
  }, [isAdmin, user?.role, user?.email]);

  useEffect(() => {
    loadSites();
  }, [loadSites]);

  const openAdd = () => {
    setModalMode("add");
    setEditingSite(null);
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
    setFortnightCostBudget("");
    setSelectedManagerEmails([]);
    const freq = budget?.visitFrequency;
    const normalizedFreq: VisitFreq =
      freq === "Fortnightly" || freq === "Monthly" ? freq : "Weekly";
    setVisitFrequency(normalizedFreq);
    setHoursPerVisit(
      budget?.hoursPerVisit != null && budget.hoursPerVisit > 0 ? budget.hoursPerVisit : ""
    );
    setSubmitError(null);
    getGraphAccessToken().then(async (token) => {
      if (!token) return;
      const [managersList, assignments] = await Promise.all([
        getCleanTrackManagers(token).catch(() => []),
        getAssignedManagersForSite(token, site.id).catch(() => []),
      ]);
      setManagers(managersList);
      setSelectedManagerEmails(assignments.map((a) => a.email));
    });
  };

  const closeModal = () => {
    setModalMode(null);
    setEditingSite(null);
    setSubmitError(null);
  };

  /** Create a new empty bulk row. */
  const createEmptyBulkRow = useCallback((): BulkSiteRow => ({
    id: `bulk-${Date.now()}-${Math.random().toString(36).slice(2)}`,
    siteName: "",
    address: "",
    monthlyRevenue: "",
    fortnightCostBudget: "",
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
    setBulkManagerEmails([]);
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
          });
        } catch (budgetErr) {
          console.warn("Bulk: site budget create failed for", siteName, budgetErr);
        }
        for (const email of bulkManagerEmails) {
          try {
            await createSiteManagerAssignment(token, siteId, email, {
              assignmentName: `${siteName} - ${email}`,
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
    await loadSites();
    onUpdateSite?.();
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
          });
        } catch (budgetErr) {
          console.warn("Site budget create failed (site was created):", budgetErr);
        }
        for (const email of selectedManagerEmails) {
          try {
            await createSiteManagerAssignment(token, siteId, email, {
              assignmentName: `${form.siteName.trim()} - ${email}`,
            });
          } catch (assignErr) {
            console.warn("Manager assignment failed for", email, assignErr);
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
          });
        }
        const currentAssignments = await getAssignedManagersForSite(token, editingSite.id);
        const currentEmails = new Set(currentAssignments.map((a) => a.email.toLowerCase()));
        const selectedSet = new Set(selectedManagerEmails.map((e) => e.toLowerCase()));
        for (const a of currentAssignments) {
          if (!selectedSet.has(a.email.toLowerCase())) {
            try {
              await deleteSiteManagerAssignment(token, a.itemId);
            } catch (err) {
              console.warn("Remove manager assignment failed:", err);
            }
          }
        }
        for (const email of selectedManagerEmails) {
          if (!currentEmails.has(email.toLowerCase())) {
            try {
              await createSiteManagerAssignment(token, editingSite.id, email, {
                assignmentName: `${form.siteName.trim()} - ${email}`,
              });
            } catch (err) {
              console.warn("Add manager assignment failed:", err);
            }
          }
        }
        showToast("Site updated.");
      }
      closeModal();
      await loadSites();
      onUpdateSite?.();
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
      await loadSites();
      onUpdateSite?.();
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
      await loadSites();
      onUpdateSite?.();
    } catch (e) {
      const msg = e instanceof Error ? e.message : "Delete failed.";
      showToast(msg.includes("403") || msg.includes("401") ? `Permission missing. ${WRITE_PERMISSION_HINT}` : msg);
    }
  };

  return (
    <div className="space-y-8">
      <div className="flex justify-between items-end border-b border-[#edeef0] pb-4">
        <div>
          <h2 className="text-3xl font-bold text-gray-900">Sites</h2>
          <p className="text-gray-500 text-sm mt-1">
            Sites &amp; Budgets — Configure service windows and financial caps per site.
          </p>
        </div>
        {isAdmin && (
          <div className="flex items-center gap-2">
            <button
              onClick={openBulkModal}
              className="bg-white text-gray-900 border border-[#edeef0] px-4 py-2 rounded-lg text-sm font-bold hover:bg-gray-50 transition-colors flex items-center gap-2"
            >
              <Layers size={16} /> Bulk Add
            </button>
            <button
              onClick={openAdd}
              className="bg-gray-900 text-white px-4 py-2 rounded-lg text-sm font-bold hover:bg-gray-800 transition-colors flex items-center gap-2"
            >
              <Plus size={16} /> New Site
            </button>
          </div>
        )}
      </div>

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
      ) : sites.length === 0 ? (
        <div className="text-gray-500 py-8">No sites found.</div>
      ) : (
        <div className="space-y-3">
          {sites.map((site) => {
            const budget = budgetsBySiteId[String(site.id)];
            const cleaners = assignedCleanersBySiteId[site.id] ?? [];
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

            return (
              <div
                key={site.id}
                className="border border-[#edeef0] rounded-lg bg-white shadow-sm px-4 py-3 flex flex-wrap items-center justify-between gap-2"
              >
                <div className="flex-1 min-w-0 flex flex-wrap items-center gap-x-4 gap-y-1">
                  <h3 className="text-base font-bold text-gray-900">{site.siteName || "Unnamed site"}</h3>
                  {site.address && (
                    <p className="text-xs text-gray-600 flex items-center gap-1">
                      <MapPin size={12} className="text-gray-400 shrink-0" />
                      {site.address}
                    </p>
                  )}
                  <div className="flex flex-wrap gap-1.5">
                    {cleaners.length === 0 ? (
                      <span className="text-[11px] text-gray-400">No assigned cleaners</span>
                    ) : (
                      cleaners.map((c, i) => (
                        <span
                          key={i}
                          className="inline-flex items-center px-2 py-0.5 rounded-full text-[11px] font-medium bg-gray-100 text-gray-800"
                        >
                          {c.name} ${c.payRate}/h
                        </span>
                      ))
                    )}
                  </div>
                </div>
                <div className="flex flex-wrap items-center gap-2">
                  <div className="flex flex-wrap gap-1">
                    {dayHours.map(({ day, h }) => (
                      <span
                        key={day}
                        className={`inline-flex items-center px-1.5 py-0.5 rounded text-[11px] font-medium ${
                          (h ?? 0) > 0 ? "bg-blue-100 text-blue-800" : "bg-gray-50 text-gray-400"
                        }`}
                      >
                        {day} {h ?? 0}h
                      </span>
                    ))}
                  </div>
                  <div className="text-right shrink-0">
                    <p className="text-[9px] font-bold text-gray-500 uppercase tracking-wider">FORTNIGHTLY CAP</p>
                    <p className="text-lg font-bold text-gray-900 leading-tight">{fortnightCap}h</p>
                  </div>
                </div>
                {isAdmin && (
                  <div className="w-full flex items-center justify-end gap-2 pt-1.5 mt-1 border-t border-[#edeef0]">
                    <button
                      onClick={() => openEdit(site)}
                      className="text-blue-600 hover:text-blue-800 text-xs font-medium flex items-center gap-1"
                    >
                      <Edit3 size={12} /> Edit
                    </button>
                    <button
                      onClick={() => handleSetActive(site, !site.active)}
                      className="text-gray-600 hover:text-gray-900 text-xs font-medium"
                    >
                      {site.active ? "Deactivate" : "Activate"}
                    </button>
                    <button
                      onClick={() => handleDeleteSite(site)}
                      className="text-red-600 hover:text-red-800 text-xs font-medium flex items-center gap-1"
                      title="Delete site"
                    >
                      <Trash2 size={12} /> Delete
                    </button>
                  </div>
                )}
              </div>
            );
          })}
          <p className="text-gray-400 text-xs flex items-center gap-1">
            <span className="inline-block w-4 h-4 rounded bg-gray-200 flex items-center justify-center text-[10px]">i</span>
            Daily budgets are persistent across fortnight cycles.
          </p>
        </div>
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
                    {modalMode === "add" && (
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
                    )}
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
                  {managers.map((m) => (
                    <label key={m.email} className="flex items-center gap-2 cursor-pointer">
                      <input
                        type="checkbox"
                        checked={selectedManagerEmails.includes(m.email)}
                        onChange={(e) => {
                          if (e.target.checked) {
                            setSelectedManagerEmails((prev) => [...prev, m.email]);
                          } else {
                            setSelectedManagerEmails((prev) => prev.filter((x) => x !== m.email));
                          }
                        }}
                        className="rounded border-gray-300"
                      />
                      <span className="text-sm text-gray-800">{m.fullName}</span>
                    </label>
                  ))}
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
                        <th className="py-2 px-1 w-16">State</th>
                        <th className="py-2 px-1 w-12 text-center">Active</th>
                        <th className="py-2 px-1 w-20">Visit freq</th>
                        {DAY_KEYS.map((d) => (
                          <th key={d} className="py-2 px-0.5 w-10 text-center">{d}</th>
                        ))}
                        {DAY_KEYS.map((d) => (
                          <th key={"w2-" + d} className="py-2 px-0.5 w-10 text-center">W2 {d}</th>
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
                    <label key={m.email} className="flex items-center gap-2 cursor-pointer">
                      <input
                        type="checkbox"
                        checked={bulkManagerEmails.includes(m.email)}
                        onChange={(e) => {
                          if (e.target.checked) setBulkManagerEmails((prev) => [...prev, m.email]);
                          else setBulkManagerEmails((prev) => prev.filter((x) => x !== m.email));
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
