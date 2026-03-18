/**
 * Ad Hoc Jobs – list, filters (month, status, manager, site), create/edit form.
 * Admin: all jobs. Manager: jobs assigned to them.
 */
import React, { useState, useEffect, useCallback, useMemo } from "react";
import { AdHocJob } from "../types";
import { Plus, X, Loader2, Pencil, AlertCircle, Upload, FileText, Trash2, FileSpreadsheet } from "lucide-react";
import { useRole } from "../contexts/RoleContext";
import { useAppAuth } from "../contexts/AppAuthContext";
import { getGraphAccessToken } from "../lib/graph";
import { getAdHocJobs, createAdHocJob, updateAdHocJob, uploadAdHocJobAttachments, deleteAdHocJob, getAdHocJobAttachments, deleteAdHocJobAttachment, type AdHocJobPayload, type AdHocJobFilters, type AdHocAttachment } from "../repositories/adHocJobsRepo";
import { getCleanTrackUserByEmail, getCleanTrackManagers } from "../repositories/usersRepo";
import { getSites } from "../repositories/sitesRepo";
import type { Site } from "../repositories/sitesRepo";
import { endOfMonth, format } from "date-fns";
import { getPublicHolidaysInRange } from "../lib/publicHolidays";
import { generateAdHocOccurrencesForRange } from "../lib/adhocSchedule";
import { exportAdHocJobsToSpreadsheet } from "../services/exportService";

const STATUS_OPTIONS = ["Requested", "Approved", "Scheduled", "Completed", "Cancelled", "In Progress"];
// Reuse the existing "Job Type" column as schedule semantics (minimal schema approach).
const SCHEDULE_TYPE_OPTIONS = ["Once Off", "Recurring"];
const RECURRENCE_FREQUENCY_OPTIONS = ["Weekly", "Fortnightly", "Monthly"] as const;
const WEEK_OF_MONTH_OPTIONS = ["First", "Second", "Third", "Fourth", "Last"] as const;
const MONTHLY_MODE_OPTIONS = [
  { id: "day_of_month", label: "Day of Month" },
  { id: "nth_weekday", label: "Nth Weekday" },
] as const;
const WEEKDAY_OPTIONS = [
  { id: 1, label: "Mon" },
  { id: 2, label: "Tue" },
  { id: 3, label: "Wed" },
  { id: 4, label: "Thu" },
  { id: 5, label: "Fri" },
  { id: 6, label: "Sat" },
  { id: 0, label: "Sun" },
] as const;

function normalizeScheduleType(raw: string | undefined | null): "Once Off" | "Recurring" {
  const s = String(raw ?? "").trim().toLowerCase();
  if (s.includes("recurr")) return "Recurring";
  return "Once Off";
}

function scheduleTypeLabel(raw: string | undefined | null): string {
  const s = String(raw ?? "").trim().toLowerCase();
  if (s.includes("recurr")) return "Recurring";
  return "Once Off";
}

export default function AdHocJobsManager() {
  const { isAdmin: isAdminFromRole } = useRole();
  const { user } = useAppAuth();
  const isAdmin = isAdminFromRole || user?.role === "Admin";

  const [jobs, setJobs] = useState<AdHocJob[]>([]);
  const [sites, setSites] = useState<Site[]>([]);
  const [managers, setManagers] = useState<{ id: string; fullName: string; email: string }[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [toast, setToast] = useState<string | null>(null);

  const [filterMonth, setFilterMonth] = useState<string>(() => format(new Date(), "yyyy-MM"));
  const [filterStatus, setFilterStatus] = useState<string>("");
  const [filterManagerId, setFilterManagerId] = useState<string>("");
  const [filterSiteId, setFilterSiteId] = useState<string>("");

  const [modalOpen, setModalOpen] = useState(false);
  const [editingJob, setEditingJob] = useState<AdHocJob | null>(null);
  const [submitLoading, setSubmitLoading] = useState(false);
  const [currentUserId, setCurrentUserId] = useState<string | null>(null);
  const [rowDeletingId, setRowDeletingId] = useState<string | null>(null);

  const showToast = useCallback((msg: string) => {
    setToast(msg);
    setTimeout(() => setToast(null), 4000);
  }, []);

  const loadJobs = useCallback(async () => {
    const token = await getGraphAccessToken();
    if (!token) {
      setError("Sign in with Microsoft to view ad hoc jobs.");
      setJobs([]);
      setLoading(false);
      return;
    }
    setLoading(true);
    setError(null);
    try {
      const filters: AdHocJobFilters = { month: filterMonth };
      if (filterStatus) filters.status = filterStatus;
      if (filterManagerId) filters.assignedManagerId = filterManagerId;
      if (filterSiteId) filters.siteId = filterSiteId;
      if (!isAdmin && user?.email) {
        const ctUser = await getCleanTrackUserByEmail(token, user.email);
        if (ctUser?.id) filters.assignedManagerId = ctUser.id;
      }
      const data = await getAdHocJobs(token, filters);
      setJobs(data);
    } catch (e) {
      const msg = e instanceof Error ? e.message : "Failed to load ad hoc jobs.";
      setError(msg);
      setJobs([]);
    } finally {
      setLoading(false);
    }
  }, [isAdmin, user?.email, filterMonth, filterStatus, filterManagerId, filterSiteId]);

  useEffect(() => {
    loadJobs();
  }, [loadJobs]);

  useEffect(() => {
    let cancelled = false;
    getGraphAccessToken().then(async (token) => {
      if (!token || !user?.email || cancelled) return;
      const ctUser = await getCleanTrackUserByEmail(token, user.email);
      if (!cancelled && ctUser?.id) setCurrentUserId(ctUser.id);
    });
    return () => { cancelled = true; };
  }, [user?.email]);

  useEffect(() => {
    let cancelled = false;
    getGraphAccessToken().then(async (token) => {
      if (!token || cancelled) return;
      try {
        const [siteList, managerList] = await Promise.all([getSites(token), getCleanTrackManagers(token)]);
        if (!cancelled) {
          setSites(siteList);
          setManagers(managerList);
        }
      } catch {
        if (!cancelled) setManagers([]);
      }
    });
    return () => { cancelled = true; };
  }, []);

  const openCreate = () => {
    setEditingJob(null);
    setModalOpen(true);
  };

  const openEdit = (job: AdHocJob) => {
    setEditingJob(job);
    setModalOpen(true);
  };

  const handleDelete = useCallback(
    async (job: AdHocJob) => {
      if (!window.confirm(`Delete "${job.jobName || "this job"}"? This will remove it from the Ad Hoc Jobs list.`)) {
        return;
      }
      const token = await getGraphAccessToken();
      if (!token) return;
      setRowDeletingId(job.id);
      try {
        await deleteAdHocJob(token, job.id);
        await loadJobs();
        showToast("Ad hoc job deleted.");
      } catch (err) {
        alert(err instanceof Error ? err.message : "Delete failed.");
      } finally {
        setRowDeletingId(null);
      }
    },
    [loadJobs, showToast]
  );

  const closeModal = () => {
    setModalOpen(false);
    setEditingJob(null);
  };

  return (
    <div className="space-y-6">
      <p className="text-gray-500 text-sm">
        One-off and additional work: carpet cleans, emergency cleans, sporadic jobs. Create jobs, capture requester details and approval proof, then link timesheets.
      </p>

      <div className="flex flex-col sm:flex-row gap-4 items-start sm:items-center justify-between">
        <div className="flex flex-wrap gap-3 items-center">
          <label className="flex items-center gap-2">
            <span className="text-xs font-bold text-gray-500 uppercase">Month</span>
            <input
              type="month"
              value={filterMonth}
              onChange={(e) => setFilterMonth(e.target.value)}
              className="border border-[#edeef0] rounded-lg px-3 py-2 text-sm"
            />
          </label>
          <label className="flex items-center gap-2">
            <span className="text-xs font-bold text-gray-500 uppercase">Status</span>
            <select
              value={filterStatus}
              onChange={(e) => setFilterStatus(e.target.value)}
              className="border border-[#edeef0] rounded-lg px-3 py-2 text-sm"
            >
              <option value="">All</option>
              {STATUS_OPTIONS.map((s) => (
                <option key={s} value={s}>{s}</option>
              ))}
            </select>
          </label>
          {isAdmin && (
            <>
              <label className="flex items-center gap-2">
                <span className="text-xs font-bold text-gray-500 uppercase">Manager</span>
                <select
                  value={filterManagerId}
                  onChange={(e) => setFilterManagerId(e.target.value)}
                  className="border border-[#edeef0] rounded-lg px-3 py-2 text-sm min-w-[140px]"
                >
                  <option value="">All</option>
                  {managers.map((m) => (
                    <option key={m.id} value={m.id}>{m.fullName || m.email}</option>
                  ))}
                </select>
              </label>
              <label className="flex items-center gap-2">
                <span className="text-xs font-bold text-gray-500 uppercase">Site</span>
                <select
                  value={filterSiteId}
                  onChange={(e) => setFilterSiteId(e.target.value)}
                  className="border border-[#edeef0] rounded-lg px-3 py-2 text-sm min-w-[180px]"
                >
                  <option value="">All</option>
                  {sites.map((s) => (
                    <option key={s.id} value={s.id}>{s.siteName || s.address || s.id}</option>
                  ))}
                </select>
              </label>
            </>
          )}
        </div>
        <div className="flex items-center gap-2">
          <button
            type="button"
            onClick={() => exportAdHocJobsToSpreadsheet(jobs, filterMonth)}
            className="flex items-center gap-1.5 px-4 py-2.5 bg-green-50 text-green-700 border border-green-200 rounded-lg text-sm font-bold hover:bg-green-100 transition-colors"
          >
            <FileSpreadsheet size={18} />
            Export (XLSX)
          </button>
          <button
            type="button"
            onClick={openCreate}
            className="flex items-center gap-2 bg-gray-900 text-white px-4 py-2.5 rounded-lg text-sm font-bold hover:bg-gray-800"
          >
            <Plus size={18} />
            New Ad Hoc Job
          </button>
        </div>
      </div>

      {toast && (
        <div className="bg-green-50 border border-green-200 text-green-800 px-4 py-2 rounded-lg text-sm">
          {toast}
        </div>
      )}

      {error && (
        <div className="bg-amber-50 border border-amber-200 text-amber-800 px-4 py-3 rounded-lg text-sm">
          {error}
        </div>
      )}

      {loading ? (
        <div className="flex items-center gap-2 text-gray-500">
          <Loader2 className="animate-spin" size={20} />
          Loading ad hoc jobs…
        </div>
      ) : jobs.length === 0 ? (
        <div className="text-gray-500 py-8">
          No ad hoc jobs match the filters. Create one to get started.
        </div>
      ) : (
        <div className="border border-[#edeef0] rounded-lg bg-white shadow-sm overflow-x-auto">
          <table className="w-full border-collapse text-left table-auto min-w-[1320px]">
            <colgroup>
              <col className="min-w-[140px]" />
              <col className="min-w-[100px]" />
              <col className="min-w-[100px]" />
              <col className="min-w-[100px]" />
              <col className="min-w-[160px]" />
              <col className="min-w-[120px]" />
              <col className="min-w-[88px]" />
              <col className="min-w-[88px]" />
              <col className="min-w-[88px]" />
              <col className="min-w-[90px]" />
              <col className="min-w-[88px]" />
              <col className="min-w-[72px]" />
              <col className="min-w-[72px]" />
              <col className="min-w-[72px]" />
              <col className="min-w-[72px]" />
              <col className="w-20" />
            </colgroup>
            <thead>
              <tr className="bg-[#fcfcfb] border-b border-[#edeef0]">
                <th className="px-3 py-2 text-[9px] font-bold text-gray-500 uppercase tracking-widest text-left">Job Name</th>
                <th className="px-2 py-2 text-[9px] font-bold text-gray-500 uppercase tracking-widest hidden md:table-cell">Schedule Type</th>
                <th className="px-2 py-2 text-[9px] font-bold text-gray-500 uppercase tracking-widest hidden lg:table-cell">Company</th>
                <th className="px-2 py-2 text-[9px] font-bold text-gray-500 uppercase tracking-widest hidden lg:table-cell">Client</th>
                <th className="px-2 py-2 text-[9px] font-bold text-gray-500 uppercase tracking-widest hidden md:table-cell">Site</th>
                <th className="px-2 py-2 text-[9px] font-bold text-gray-500 uppercase tracking-widest hidden md:table-cell">Assigned Manager</th>
                <th className="px-2 py-2 text-[9px] font-bold text-gray-500 uppercase tracking-widest">Requested</th>
                <th className="px-2 py-2 text-[9px] font-bold text-gray-500 uppercase tracking-widest hidden lg:table-cell">Scheduled</th>
                <th className="px-2 py-2 text-[9px] font-bold text-gray-500 uppercase tracking-widest hidden lg:table-cell">Completed</th>
                <th className="px-2 py-2 text-[9px] font-bold text-gray-500 uppercase tracking-widest">Status</th>
                <th className="px-2 py-2 text-[9px] font-bold text-gray-500 uppercase tracking-widest">Budgeted Hrs</th>
                <th className="px-2 py-2 text-[9px] font-bold text-gray-500 uppercase tracking-widest hidden xl:table-cell">Charge</th>
                <th className="px-2 py-2 text-[9px] font-bold text-gray-500 uppercase tracking-widest hidden xl:table-cell">Cost</th>
                <th className="px-2 py-2 text-[9px] font-bold text-gray-500 uppercase tracking-widest hidden 2xl:table-cell">Gross Profit</th>
                <th className="px-2 py-2 text-[9px] font-bold text-gray-500 uppercase tracking-widest">Proof</th>
                <th className="px-2 py-2 text-[9px] font-bold text-gray-500 uppercase tracking-widest w-20">Actions</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-[#edeef0]">
              {jobs.map((j) => {
                const canDelete =
                  isAdmin ||
                  (currentUserId &&
                    j.assignedManagerId &&
                    j.assignedManagerId === currentUserId);
                const siteMatch = sites.find((s) => s.id === j.siteId);
                const siteLabel =
                  siteMatch?.siteName || siteMatch?.address || j.siteName || "—";
                return (
                <tr key={j.id} className="hover:bg-[#f7f6f3]">
                  <td className="px-3 py-2 max-w-[200px]">
                    <button
                      type="button"
                      onClick={() => openEdit(j)}
                      className="text-sm font-semibold text-gray-900 hover:underline text-left block truncate w-full"
                      title={j.jobName || "—"}
                    >
                      {j.jobName || "—"}
                    </button>
                  </td>
                  <td className="px-2 py-2 text-xs text-gray-600 hidden md:table-cell">
                    {scheduleTypeLabel(j.jobType)}
                  </td>
                  <td className="px-2 py-2 text-xs text-gray-600 hidden md:table-cell max-w-[120px]" title={j.companyName || "—"}>
                    <span className="block truncate">{j.companyName || "—"}</span>
                  </td>
                  <td className="px-2 py-2 text-xs text-gray-600 hidden md:table-cell max-w-[120px]" title={j.clientName || "—"}>
                    <span className="block truncate">{j.clientName || "—"}</span>
                  </td>
                  <td className="px-2 py-2 text-xs text-gray-600 hidden md:table-cell max-w-[180px]" title={siteLabel}>
                    <span className="block truncate">{siteLabel}</span>
                  </td>
                  <td className="px-2 py-2 text-xs text-gray-600 hidden md:table-cell max-w-[120px]" title={j.assignedManagerName || "—"}>
                    <span className="block truncate">{j.assignedManagerName || "—"}</span>
                  </td>
                  <td className="px-2 py-2 text-xs text-gray-700">{j.requestedDate ? format(new Date(j.requestedDate), "dd MMM yyyy") : "—"}</td>
                  <td className="px-2 py-2 text-xs text-gray-700 hidden md:table-cell">{j.scheduledDate ? format(new Date(j.scheduledDate), "dd MMM yyyy") : "—"}</td>
                  <td className="px-2 py-2 text-xs text-gray-700 hidden md:table-cell">{j.completedDate ? format(new Date(j.completedDate), "dd MMM yyyy") : "—"}</td>
                  <td className="px-2 py-2">
                    <span className={`text-xs font-bold px-1.5 py-0.5 rounded ${
                      j.status === "Completed" ? "bg-green-100 text-green-800" :
                      j.status === "Scheduled" || j.status === "Approved" ? "bg-blue-100 text-blue-800" :
                      "bg-gray-100 text-gray-800"
                    }`}>
                      {j.status || "Requested"}
                    </span>
                  </td>
                  <td className="px-2 py-2 text-xs font-medium text-gray-700">
                    {j.budgetedHours != null ? `${j.budgetedHours}h` : "—"}
                  </td>
                  <td className="px-2 py-2 text-xs text-gray-700 hidden md:table-cell">
                    {j.charge != null ? `$${Number(j.charge).toFixed(2)}` : "—"}
                  </td>
                  <td className="px-2 py-2 text-xs text-gray-700 hidden md:table-cell">
                    {j.cost != null ? `$${Number(j.cost).toFixed(2)}` : "—"}
                  </td>
                  <td className="px-2 py-2 text-xs text-gray-700 hidden md:table-cell">
                    {j.grossProfit != null ? `$${Number(j.grossProfit).toFixed(2)}` : "—"}
                  </td>
                  <td className="px-2 py-2">
                    {j.approvalProofRequired && !j.approvalProofUploaded ? (
                      <span className="inline-flex items-center gap-0.5 text-amber-600 text-xs" title="Approval proof missing">
                        <AlertCircle size={14} /> Missing
                      </span>
                    ) : j.approvalProofUploaded ? (
                      <span className="text-xs text-green-600">Yes</span>
                    ) : (
                      <span className="text-xs text-gray-400">—</span>
                    )}
                  </td>
                  <td className="px-2 py-2">
                    <div className="flex items-center gap-1">
                      <button
                        type="button"
                        onClick={() => openEdit(j)}
                        className="p-2 rounded text-blue-600 hover:bg-blue-50"
                        aria-label={`Edit ${j.jobName}`}
                      >
                        <Pencil size={16} />
                      </button>
                      {canDelete && (
                        <button
                          type="button"
                          onClick={() => handleDelete(j)}
                          className="p-2 rounded text-red-600 hover:bg-red-50 disabled:opacity-40"
                          aria-label={`Delete ${j.jobName}`}
                          disabled={rowDeletingId === j.id}
                        >
                          <Trash2 size={16} />
                        </button>
                      )}
                    </div>
                  </td>
                </tr>
              );})}
            </tbody>
          </table>
        </div>
      )}

      {modalOpen && (
        <AdHocJobFormModal
          job={editingJob}
          sites={sites}
          managers={managers}
          currentUserId={currentUserId}
          onClose={closeModal}
          onSaved={() => {
            loadJobs();
            closeModal();
            showToast(editingJob ? "Job updated." : "Ad hoc job created.");
          }}
          submitLoading={submitLoading}
          setSubmitLoading={setSubmitLoading}
          isAdmin={isAdmin}
        />
      )}
    </div>
  );
}

interface AdHocJobFormModalProps {
  job: AdHocJob | null;
  sites: Site[];
  managers: { id: string; fullName: string; email: string }[];
  currentUserId: string | null;
  onClose: () => void;
  onSaved: () => void;
  submitLoading: boolean;
  setSubmitLoading: (v: boolean) => void;
  isAdmin: boolean;
}

function AdHocJobFormModal({
  job,
  sites,
  managers,
  currentUserId,
  onClose,
  onSaved,
  submitLoading,
  setSubmitLoading,
  isAdmin,
}: AdHocJobFormModalProps) {
  const isEdit = !!job;
  const [proofFiles, setProofFiles] = useState<File[]>([]);
  const [form, setForm] = useState<AdHocJobPayload>({
    jobName: job?.jobName ?? "",
    jobType: job?.jobType ?? "",
    companyName: job?.companyName ?? "",
    clientName: job?.clientName ?? "",
    siteId: job?.siteId ?? null,
    manualSiteName: job?.manualSiteName ?? "",
    manualSiteAddress: (job as any)?.manualSiteAddress ?? "",
    assignedManagerId: job?.assignedManagerId ?? currentUserId ?? null,
    requestedByName: job?.requestedByName ?? "",
    requestedByEmail: job?.requestedByEmail ?? "",
    requestChannel: job?.requestChannel ?? "",
    requestedDate: job?.requestedDate ?? format(new Date(), "yyyy-MM-dd"),
    scheduledDate: job?.scheduledDate ?? null,
    completedDate: job?.completedDate ?? null,
    status: job?.status ?? "Requested",
    budgetedHours: job?.budgetedHours ?? null,
    recurrenceFrequency: job?.recurrenceFrequency ?? null,
    recurrenceEndDate: job?.recurrenceEndDate ?? null,
    hoursPerServiceDay: job?.hoursPerServiceDay ?? null,
    recurrenceWeekdays: job?.recurrenceWeekdays ?? null,
    weekdayHours: (job as any)?.weekdayHours ?? null,
    monthlyMode: job?.monthlyMode ?? null,
    monthlyDayOfMonth: job?.monthlyDayOfMonth ?? null,
    monthlyWeekOfMonth: job?.monthlyWeekOfMonth ?? null,
    monthlyWeekday: job?.monthlyWeekday ?? null,
    monthlyHours: (job as any)?.monthlyHours ?? null,
    weekdayChargeRateOverride: (job as any)?.weekdayChargeRateOverride ?? null,
    saturdayChargeRateOverride: (job as any)?.saturdayChargeRateOverride ?? null,
    sundayChargeRateOverride: (job as any)?.sundayChargeRateOverride ?? null,
    publicHolidayChargeRateOverride: (job as any)?.publicHolidayChargeRateOverride ?? null,
    weekdayCostRateOverride: (job as any)?.weekdayCostRateOverride ?? null,
    saturdayCostRateOverride: (job as any)?.saturdayCostRateOverride ?? null,
    sundayCostRateOverride: (job as any)?.sundayCostRateOverride ?? null,
    publicHolidayCostRateOverride: (job as any)?.publicHolidayCostRateOverride ?? null,
    actualHours: job?.actualHours ?? null,
    serviceProvider: job?.serviceProvider ?? "",
    chargeRatePerHour: job?.chargeRatePerHour ?? null,
    costRatePerHour: job?.costRatePerHour ?? null,
    charge: job?.charge ?? null,
    cost: job?.cost ?? null,
    grossProfit: job?.grossProfit ?? null,
    markupPercent: job?.markupPercent ?? null,
    gpPercent: job?.gpPercent ?? null,
    description: job?.description ?? "",
    approvalProofUploaded: job?.approvalProofUploaded ?? false,
    approvalReference: job?.approvalReference ?? "",
    notesForInformation: job?.notesForInformation ?? "",
    active: job?.active ?? true,
  });
  const [siteSearch, setSiteSearch] = useState<string>("");
  const [pasteMode, setPasteMode] = useState(false);
  const [pasteError, setPasteError] = useState<string | null>(null);
  const [existingAttachments, setExistingAttachments] = useState<AdHocAttachment[]>([]);
  const scheduleType = normalizeScheduleType(form.jobType);
  const isRecurring = scheduleType === "Recurring";
  const [siteMode, setSiteMode] = useState<"existing" | "new">(() =>
    form.siteId ? "existing" : "new"
  );

  const filteredSites = useMemo(() => {
    const q = siteSearch.trim().toLowerCase();
    if (!q) return sites;
    return sites.filter((s) => {
      const name = (s.siteName || "").toLowerCase();
      const address = (s.address || "").toLowerCase();
      return name.includes(q) || address.includes(q);
    });
  }, [sites, siteSearch]);

  useEffect(() => {
    if (!form.siteId) {
      setSiteSearch("");
      return;
    }
    const selected = sites.find((s) => s.id === form.siteId);
    if (selected) {
      const label = selected.siteName || selected.address || selected.id;
      setSiteSearch(label);
    }
  }, [form.siteId, sites]);

  useEffect(() => {
    if (job) {
      setForm({
        jobName: job.jobName ?? "",
        jobType: job.jobType ?? "",
        companyName: job.companyName ?? "",
        clientName: job.clientName ?? "",
        siteId: job.siteId ?? null,
        manualSiteName: job.manualSiteName ?? "",
        assignedManagerId: job.assignedManagerId ?? null,
        requestedByName: job.requestedByName ?? "",
        requestedByEmail: job.requestedByEmail ?? "",
        requestChannel: job.requestChannel ?? "",
        requestedDate: job.requestedDate ?? null,
        scheduledDate: job.scheduledDate ?? null,
        completedDate: job.completedDate ?? null,
        status: job.status ?? "Requested",
        budgetedHours: job.budgetedHours ?? null,
        recurrenceFrequency: job.recurrenceFrequency ?? null,
        recurrenceEndDate: job.recurrenceEndDate ?? null,
        hoursPerServiceDay: job.hoursPerServiceDay ?? null,
        recurrenceWeekdays: job.recurrenceWeekdays ?? null,
        weekdayHours: (job as any).weekdayHours ?? null,
        monthlyMode: job.monthlyMode ?? null,
        monthlyDayOfMonth: job.monthlyDayOfMonth ?? null,
        monthlyWeekOfMonth: job.monthlyWeekOfMonth ?? null,
        monthlyWeekday: job.monthlyWeekday ?? null,
        monthlyHours: (job as any).monthlyHours ?? null,
        weekdayChargeRateOverride: (job as any).weekdayChargeRateOverride ?? null,
        saturdayChargeRateOverride: (job as any).saturdayChargeRateOverride ?? null,
        sundayChargeRateOverride: (job as any).sundayChargeRateOverride ?? null,
        publicHolidayChargeRateOverride: (job as any).publicHolidayChargeRateOverride ?? null,
        weekdayCostRateOverride: (job as any).weekdayCostRateOverride ?? null,
        saturdayCostRateOverride: (job as any).saturdayCostRateOverride ?? null,
        sundayCostRateOverride: (job as any).sundayCostRateOverride ?? null,
        publicHolidayCostRateOverride: (job as any).publicHolidayCostRateOverride ?? null,
        actualHours: job.actualHours ?? null,
        serviceProvider: job.serviceProvider ?? "",
        chargeRatePerHour: job.chargeRatePerHour ?? null,
        costRatePerHour: job.costRatePerHour ?? null,
        charge: job.charge ?? null,
        cost: job.cost ?? null,
        grossProfit: job.grossProfit ?? null,
        markupPercent: job.markupPercent ?? null,
        gpPercent: job.gpPercent ?? null,
        description: job.description ?? "",
        approvalProofUploaded: job.approvalProofUploaded ?? false,
        approvalReference: job.approvalReference ?? "",
        notesForInformation: job.notesForInformation ?? "",
        active: job.active ?? true,
      });
      // load existing attachments for this job
      let cancelled = false;
      getGraphAccessToken().then(async (token) => {
        if (!token || cancelled) return;
        try {
          const list = await getAdHocJobAttachments(token, job.id);
          if (!cancelled) {
            setExistingAttachments(list);
            console.log("[AdHoc] existing attachments", { jobId: job.id, count: list.length });
          }
        } catch {
          if (!cancelled) setExistingAttachments([]);
        }
      });
      return () => {
        cancelled = true;
      };
    }
  }, [job]);

  useEffect(() => {
    if (!job) setProofFiles([]);
  }, [job]);

  // Build a minimal AdHocJob from form state for occurrence generation (preview / save totals).
  const formToSyntheticJob = useCallback((payload: AdHocJobPayload): AdHocJob => {
    const j = job ?? ({} as AdHocJob);
    return {
      ...j,
      id: j?.id ?? "preview",
      jobName: payload.jobName ?? "",
      jobType: payload.jobType ?? "",
      scheduledDate: payload.scheduledDate ?? null,
      recurrenceEndDate: payload.recurrenceEndDate ?? null,
      recurrenceFrequency: payload.recurrenceFrequency ?? null,
      hoursPerServiceDay: payload.hoursPerServiceDay ?? null,
      recurrenceWeekdays: payload.recurrenceWeekdays ?? null,
      weekdayHours: payload.weekdayHours ?? null,
      monthlyMode: payload.monthlyMode ?? null,
      monthlyDayOfMonth: payload.monthlyDayOfMonth ?? null,
      monthlyWeekOfMonth: payload.monthlyWeekOfMonth ?? null,
      monthlyWeekday: payload.monthlyWeekday ?? null,
      monthlyHours: (payload as any).monthlyHours ?? null,
      chargeRatePerHour: payload.chargeRatePerHour ?? null,
      costRatePerHour: payload.costRatePerHour ?? null,
      weekdayChargeRateOverride: (payload as any).weekdayChargeRateOverride ?? null,
      saturdayChargeRateOverride: (payload as any).saturdayChargeRateOverride ?? null,
      sundayChargeRateOverride: (payload as any).sundayChargeRateOverride ?? null,
      publicHolidayChargeRateOverride: (payload as any).publicHolidayChargeRateOverride ?? null,
      weekdayCostRateOverride: (payload as any).weekdayCostRateOverride ?? null,
      saturdayCostRateOverride: (payload as any).saturdayCostRateOverride ?? null,
      sundayCostRateOverride: (payload as any).sundayCostRateOverride ?? null,
      publicHolidayCostRateOverride: (payload as any).publicHolidayCostRateOverride ?? null,
      active: payload.active ?? true,
      status: payload.status ?? "Requested",
    } as AdHocJob;
  }, [job]);

  /** Recurring jobs: compute Charge/Cost/GP for the current billing month.
   *
   * Rules:
   * - Anchor is the job Start Date (`scheduledDate`).
   * - If an End Date is set, we cap the window at the earlier of End Date and end of the current month.
   * - If no End Date is set, we take the full current month window.
   * - We never generate occurrences before the later of Start Date and start of the current month.
   */
  const computeRecurringTotals = useCallback(
    (payload: AdHocJobPayload): { charge: number; cost: number; grossProfit: number; markupPercent: number | null; gpPercent: number | null } | null => {
      if (!payload.scheduledDate || !payload.recurrenceFrequency) return null;
      const startDate = new Date(payload.scheduledDate);
      if (isNaN(startDate.getTime())) return null;

      const today = new Date();
      const monthStart = new Date(today.getFullYear(), today.getMonth(), 1);
      const monthEnd = endOfMonth(today);

      const endDate = payload.recurrenceEndDate ? new Date(payload.recurrenceEndDate) : null;
      if (endDate && isNaN(endDate.getTime())) return null;

      // Clamp the preview window to the current month, and also to any explicit End Date.
      const previewStart = startDate > monthStart ? startDate : monthStart;
      const previewEnd = endDate && endDate < monthEnd ? endDate : monthEnd;

      if (previewStart > previewEnd) return null;
      const synthetic = formToSyntheticJob(payload);
      const phSet = getPublicHolidaysInRange(previewStart, previewEnd);
      const occurrences = generateAdHocOccurrencesForRange(synthetic, previewStart, previewEnd, phSet);
      const charge = occurrences.reduce((sum, o) => sum + o.chargeTotal, 0);
      const cost = occurrences.reduce((sum, o) => sum + o.costTotal, 0);
      const grossProfit = charge - cost;
      const markupPercent = cost > 0 ? (grossProfit / cost) * 100 : null;
      const gpPercent = charge > 0 ? (grossProfit / charge) * 100 : null;
      return { charge, cost, grossProfit, markupPercent, gpPercent };
    },
    [formToSyntheticJob]
  );

  const recurringTotals = useMemo(() => {
    if (!isRecurring) return null;
    return computeRecurringTotals(form);
  }, [isRecurring, form, computeRecurringTotals]);

  // Recalculate financial outputs when base inputs change.
  // For Once Off jobs we can safely compute Charge/Cost/GP from budgeted hours and base rates.
  // For Recurring jobs totals depend on the selected period, so we leave Charge/Cost/GP untouched.
  const recalcFinance = useCallback((base: AdHocJobPayload): AdHocJobPayload => {
    const schedule = normalizeScheduleType(base.jobType);
    if (schedule === "Recurring") {
      return base;
    }

    const hours = base.budgetedHours ?? 0;
    const chargeRate = base.chargeRatePerHour ?? 0;
    const costRate = base.costRatePerHour ?? 0;

    const charge = hours * chargeRate;
    const cost = hours * costRate;
    const grossProfit = charge - cost;

    const markupPercent =
      cost > 0 ? ((grossProfit / cost) * 100) : null;
    const gpPercent =
      charge > 0 ? ((grossProfit / charge) * 100) : null;

    const next: AdHocJobPayload = {
      ...base,
      charge,
      cost,
      grossProfit,
      markupPercent,
      gpPercent,
    };

    console.log("[AdHoc] recalc finance", {
      budgetedHours: hours,
      chargeRatePerHour: chargeRate,
      costRatePerHour: costRate,
      charge,
      cost,
      grossProfit,
      markupPercent,
      gpPercent,
    });

    return next;
  }, []);

  const fmt2 = (v: number | null | undefined): string => {
    if (v == null || Number.isNaN(v)) return "";
    return Number(v).toFixed(2);
  };

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!form.jobName?.trim()) return;
    // Conditional validation (keep minimal and predictable).
    if (!form.jobType?.trim()) {
      alert("Schedule Type is required.");
      return;
    }
    if (siteMode === "existing") {
      if (!form.siteId) {
        alert("Select an existing site, or switch to New adhoc site.");
        return;
      }
    } else {
      if (!form.manualSiteName?.trim()) {
        alert("Site name is required for a new adhoc site.");
        return;
      }
    }
    if (scheduleType === "Once Off") {
      if (!form.scheduledDate) {
        alert("Scheduled Date is required for Once Off jobs.");
        return;
      }
      if (form.budgetedHours == null) {
        alert("Budgeted Hours is required for Once Off jobs.");
        return;
      }
    } else {
      if (!form.scheduledDate) {
        alert("Start Date is required for Recurring jobs.");
        return;
      }
      if (!form.recurrenceFrequency) {
        alert("Frequency is required for Recurring jobs.");
        return;
      }
      if (form.recurrenceFrequency === "Weekly" || form.recurrenceFrequency === "Fortnightly") {
        const wh = (form as any).weekdayHours as Record<string, number> | null;
        const hasAny =
          wh &&
          Object.values(wh).some((v) => typeof v === "number" && v > 0);
        if (!hasAny) {
          alert("Enter hours for at least one weekday for Weekly/Fortnightly jobs.");
          return;
        }
      }
      if (form.recurrenceFrequency === "Monthly") {
        if (!form.monthlyMode) {
          alert("Select a Monthly recurrence mode.");
          return;
        }
        if (form.monthlyMode === "day_of_month") {
          if (form.monthlyDayOfMonth == null) {
            alert("Day of Month is required.");
            return;
          }
        } else {
          if (!form.monthlyWeekOfMonth) {
            alert("Week of Month is required.");
            return;
          }
          if (form.monthlyWeekday == null) {
            alert("Weekday is required.");
            return;
          }
        }
        if ((form as any).monthlyHours == null) {
          alert("Hours for the monthly occurrence is required.");
          return;
        }
      }
    }
    const token = await getGraphAccessToken();
    if (!token) return;
    setSubmitLoading(true);
    try {
      let payload: AdHocJobPayload = form;
      if (scheduleType === "Recurring") {
        const totals = computeRecurringTotals(form);
        if (totals) {
          payload = {
            ...form,
            charge: totals.charge,
            cost: totals.cost,
            grossProfit: totals.grossProfit,
            markupPercent: totals.markupPercent,
            gpPercent: totals.gpPercent,
          };
        }
      }
      let itemId: string;
      if (isEdit && job) {
        await updateAdHocJob(token, job.id, payload);
        itemId = job.id;
      } else {
        const created = await createAdHocJob(token, { ...payload, status: payload.status || "Requested" });
        itemId = created.id;
      }
      if (proofFiles.length > 0) {
        await uploadAdHocJobAttachments(token, itemId, proofFiles);
        await updateAdHocJob(token, itemId, { approvalProofUploaded: true });
      }
      onSaved();
    } catch (err) {
      alert(err instanceof Error ? err.message : "Save failed.");
    } finally {
      setSubmitLoading(false);
    }
  };

  const acceptProofFiles = (files: FileList | null) => {
    if (!files?.length) return;
    const allowed = Array.from(files).filter((f) => {
      const t = f.type.toLowerCase();
      const name = (f.name || "").toLowerCase();
      return t.startsWith("image/") || name.endsWith(".pdf") || t === "application/pdf";
    });
    setProofFiles((prev) => [...prev, ...allowed]);
  };

  const handleProofPaste = (e: React.ClipboardEvent<HTMLTextAreaElement>) => {
    const items = e.clipboardData?.items;
    if (!items || items.length === 0) return;
    const images: File[] = [];
    for (const it of Array.from(items)) {
      const item = it as DataTransferItem;
      if (item.kind !== "file") continue;
      const file = item.getAsFile();
      if (!file) continue;
      if (!file.type.toLowerCase().startsWith("image/")) continue;
      const ext = (file.type.split("/")[1] || "png").toLowerCase();
      const name = `pasted-screenshot-${Date.now()}.${ext}`;
      images.push(new File([file], name, { type: file.type }));
    }
    if (images.length === 0) {
      setPasteError(
        "No image detected on the clipboard. If you copied from email, try downloading the image or dragging the attachment instead."
      );
      return;
    }
    e.preventDefault();
    setPasteError(null);
    setProofFiles((prev) => [...prev, ...images]);
  };

  const [dragOver, setDragOver] = useState(false);

  const field = (label: string, id: string, children: React.ReactNode) => (
    <div className="mb-4">
      <label htmlFor={id} className="block text-xs font-bold text-gray-500 uppercase mb-1">{label}</label>
      {children}
    </div>
  );

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 p-4" onClick={onClose}>
      <div
        className="bg-white rounded-xl shadow-xl max-w-2xl w-full max-h-[90vh] overflow-y-auto"
        onClick={(e) => e.stopPropagation()}
      >
        <div className="flex justify-between items-center p-6 border-b border-[#edeef0]">
          <h3 className="text-lg font-bold text-gray-900">{isEdit ? "Edit Ad Hoc Job" : "New Ad Hoc Job"}</h3>
          <button type="button" onClick={onClose} className="p-2 rounded text-gray-400 hover:text-gray-600"><X size={20} /></button>
        </div>
        <form onSubmit={handleSubmit} className="p-6 space-y-4">
          {field("Job Name *", "jobName", (
            <input
              id="jobName"
              type="text"
              value={form.jobName}
              onChange={(e) => setForm((f) => ({ ...f, jobName: e.target.value }))}
              className="w-full border border-[#edeef0] rounded-lg px-3 py-2 text-sm"
              placeholder="e.g. Carpet clean – Building A"
              required
            />
          ))}
          {field("Schedule Type", "jobType", (
            <select
              id="jobType"
              value={form.jobType ?? ""}
              onChange={(e) => setForm((f) => ({ ...f, jobType: e.target.value || "" }))}
              className="w-full border border-[#edeef0] rounded-lg px-3 py-2 text-sm bg-white"
            >
              <option value="">Select schedule…</option>
              {SCHEDULE_TYPE_OPTIONS.map((opt) => (
                <option key={opt} value={opt}>
                  {opt}
                </option>
              ))}
            </select>
          ))}
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
            {field("Company Name", "companyName", (
              <input
                id="companyName"
                type="text"
                value={form.companyName ?? ""}
                onChange={(e) => setForm((f) => ({ ...f, companyName: e.target.value }))}
                className="w-full border border-[#edeef0] rounded-lg px-3 py-2 text-sm"
              />
            ))}
            {field("Client Name", "clientName", (
              <input
                id="clientName"
                type="text"
                value={form.clientName ?? ""}
                onChange={(e) => setForm((f) => ({ ...f, clientName: e.target.value }))}
                className="w-full border border-[#edeef0] rounded-lg px-3 py-2 text-sm"
              />
            ))}
          </div>
          {field("Site", "site", (
            <div>
              <div className="flex items-center gap-3 mb-2 text-xs">
                <label className="inline-flex items-center gap-1">
                  <input
                    type="radio"
                    className="h-3 w-3"
                    checked={siteMode === "existing"}
                    onChange={() => {
                      setSiteMode("existing");
                      setForm((f) => ({ ...f, manualSiteName: "", manualSiteAddress: f.manualSiteAddress }));
                    }}
                  />
                  <span className="text-gray-600">Existing site</span>
                </label>
                <label className="inline-flex items-center gap-1">
                  <input
                    type="radio"
                    className="h-3 w-3"
                    checked={siteMode === "new"}
                    onChange={() => {
                      setSiteMode("new");
                      setForm((f) => ({ ...f, siteId: null }));
                    }}
                  />
                  <span className="text-gray-600">New adhoc site</span>
                </label>
              </div>
              {siteMode === "existing" ? (
                <div className="relative">
                  <input
                    id="site"
                    type="text"
                    value={siteSearch}
                    onChange={(e) => {
                      const value = e.target.value;
                      setSiteSearch(value);
                      setForm((f) => ({ ...f, siteId: null }));
                    }}
                    placeholder="Search existing site…"
                    className="w-full border border-[#edeef0] rounded-lg px-3 py-2 text-sm"
                    autoComplete="off"
                  />
                  {siteSearch.trim() && filteredSites.length > 0 && !form.siteId && (
                    <div className="absolute z-10 mt-1 w-full max-h-56 overflow-auto rounded-lg border border-[#edeef0] bg-white shadow-lg">
                      {filteredSites.map((s) => {
                        const label = s.siteName || s.address || s.id;
                        return (
                          <button
                            key={s.id}
                            type="button"
                            onClick={() => {
                              setForm((f) => ({ ...f, siteId: s.id, manualSiteName: "", manualSiteAddress: "" }));
                              setSiteSearch(label);
                            }}
                            className="block w-full px-3 py-1.5 text-left text-sm text-gray-800 hover:bg-gray-100"
                          >
                            {label}
                          </button>
                        );
                      })}
                    </div>
                  )}
                </div>
              ) : (
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                  <div>
                    <input
                      id="manualSiteName"
                      type="text"
                      value={form.manualSiteName ?? ""}
                      onChange={(e) =>
                        setForm((f) => ({ ...f, manualSiteName: e.target.value }))
                      }
                      placeholder="Site name"
                      className="w-full border border-[#edeef0] rounded-lg px-3 py-2 text-sm"
                    />
                  </div>
                  <div>
                    <input
                      id="manualSiteAddress"
                      type="text"
                      value={(form as any).manualSiteAddress ?? ""}
                      onChange={(e) =>
                        setForm((f) => ({ ...(f as any), manualSiteAddress: e.target.value }))
                      }
                      placeholder="Site address"
                      className="w-full border border-[#edeef0] rounded-lg px-3 py-2 text-sm"
                    />
                  </div>
                </div>
              )}
            </div>
          ))}
          {field("Requested By Name", "reqName", (
            <input id="reqName" type="text" value={form.requestedByName ?? ""} onChange={(e) => setForm((f) => ({ ...f, requestedByName: e.target.value }))} className="w-full border border-[#edeef0] rounded-lg px-3 py-2 text-sm" />
          ))}
          {field("Requested By Email", "reqEmail", (
            <input id="reqEmail" type="email" value={form.requestedByEmail ?? ""} onChange={(e) => setForm((f) => ({ ...f, requestedByEmail: e.target.value }))} className="w-full border border-[#edeef0] rounded-lg px-3 py-2 text-sm" />
          ))}
          {field("Request Channel", "channel", (
            <input
              id="channel"
              type="text"
              value={form.requestChannel ?? ""}
              onChange={(e) =>
                setForm((f) => ({ ...f, requestChannel: e.target.value }))
              }
              className="w-full border border-[#edeef0] rounded-lg px-3 py-2 text-sm"
              placeholder="e.g. Email, Phone"
            />
          ))}
          {field("Service Provider", "serviceProvider", (
            <input
              id="serviceProvider"
              type="text"
              value={form.serviceProvider ?? ""}
              onChange={(e) =>
                setForm((f) => ({ ...f, serviceProvider: e.target.value }))
              }
              className="w-full border border-[#edeef0] rounded-lg px-3 py-2 text-sm"
              placeholder="e.g. Adept, subcontractor name"
            />
          ))}
          <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
            {field("Requested Date", "reqDate", (
              <input
                id="reqDate"
                type="date"
                value={form.requestedDate ?? ""}
                onChange={(e) =>
                  setForm((f) => ({ ...f, requestedDate: e.target.value || null }))
                }
                className="w-full border border-[#edeef0] rounded-lg px-3 py-2 text-sm"
              />
            ))}
            {field("Assigned Manager", "manager", (
              isAdmin ? (
                <select
                  id="manager"
                  value={form.assignedManagerId ?? ""}
                  onChange={(e) =>
                    setForm((f) => ({ ...f, assignedManagerId: e.target.value || null }))
                  }
                  className="w-full border border-[#edeef0] rounded-lg px-3 py-2 text-sm"
                >
                  <option value="">— None —</option>
                  {managers.map((m) => (
                    <option key={m.id} value={m.id}>
                      {m.fullName || m.email}
                    </option>
                  ))}
                </select>
              ) : (
                <input
                  id="manager"
                  type="text"
                  value={
                    managers.find((m) => m.id === (form.assignedManagerId || currentUserId))
                      ?.fullName || "Me"
                  }
                  readOnly
                  className="w-full border border-[#edeef0] rounded-lg px-3 py-2 text-sm bg-gray-50"
                />
              )
            ))}
            {field("Status", "status", (
              <select id="status" value={form.status ?? "Requested"} onChange={(e) => setForm((f) => ({ ...f, status: e.target.value }))} className="w-full border border-[#edeef0] rounded-lg px-3 py-2 text-sm">
                {STATUS_OPTIONS.map((s) => (
                  <option key={s} value={s}>{s}</option>
                ))}
              </select>
            ))}
          </div>
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
            {field(isRecurring ? "Start Date" : "Scheduled Date", "schedDate", (
              <input
                id="schedDate"
                type="date"
                value={form.scheduledDate ?? ""}
                onChange={(e) => setForm((f) => ({ ...f, scheduledDate: e.target.value || null }))}
                className="w-full border border-[#edeef0] rounded-lg px-3 py-2 text-sm"
              />
            ))}
            {field(isRecurring ? "End Date (optional)" : "Completed Date", "compDate", (
              <input
                id="compDate"
                type="date"
                value={isRecurring ? (form.recurrenceEndDate ?? "") : (form.completedDate ?? "")}
                onChange={(e) =>
                  setForm((f) =>
                    isRecurring
                      ? ({ ...f, recurrenceEndDate: e.target.value || null })
                      : ({ ...f, completedDate: e.target.value || null })
                  )
                }
                className="w-full border border-[#edeef0] rounded-lg px-3 py-2 text-sm"
              />
            ))}
          </div>

          {scheduleType === "Once Off" ? (
            <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
              {field("Budgeted Hours", "budgHrs", (
                <input
                  id="budgHrs"
                  type="number"
                  min={0}
                  step={0.5}
                  value={form.budgetedHours ?? ""}
                  onChange={(e) =>
                    setForm((prev) =>
                      recalcFinance({
                        ...prev,
                        budgetedHours:
                          e.target.value === "" ? null : parseFloat(e.target.value) || 0,
                      })
                    )
                  }
                  className="w-full border border-[#edeef0] rounded-lg px-3 py-2 text-sm"
                />
              ))}
              {field("Charge Rate ($/hr)", "chargeRate", (
                <input
                  id="chargeRate"
                  type="number"
                  min={0}
                  step={0.01}
                  value={form.chargeRatePerHour ?? ""}
                  onChange={(e) =>
                    setForm((prev) =>
                      recalcFinance({
                        ...prev,
                        chargeRatePerHour:
                          e.target.value === "" ? null : parseFloat(e.target.value) || 0,
                      })
                    )
                  }
                  className="w-full border border-[#edeef0] rounded-lg px-3 py-2 text-sm"
                />
              ))}
              {field("Cost Rate ($/hr)", "costRate", (
                <input
                  id="costRate"
                  type="number"
                  min={0}
                  step={0.01}
                  value={form.costRatePerHour ?? ""}
                  onChange={(e) =>
                    setForm((prev) =>
                      recalcFinance({
                        ...prev,
                        costRatePerHour:
                          e.target.value === "" ? null : parseFloat(e.target.value) || 0,
                      })
                    )
                  }
                  className="w-full border border-[#edeef0] rounded-lg px-3 py-2 text-sm"
                />
              ))}
            </div>
          ) : (
            <div className="border border-[#edeef0] rounded-xl p-4 bg-[#fcfcfb] space-y-3">
              <p className="text-xs font-bold text-gray-600 uppercase tracking-widest">
                Recurring schedule
              </p>
              <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
                {field("Frequency", "recFreq", (
                  <select
                    id="recFreq"
                    value={form.recurrenceFrequency ?? ""}
                    onChange={(e) => setForm((f) => ({ ...f, recurrenceFrequency: (e.target.value as any) || null }))}
                    className="w-full border border-[#edeef0] rounded-lg px-3 py-2 text-sm bg-white"
                  >
                    <option value="">Select…</option>
                    {RECURRENCE_FREQUENCY_OPTIONS.map((opt) => (
                      <option key={opt} value={opt}>{opt}</option>
                    ))}
                  </select>
                ))}
              </div>

              {(form.recurrenceFrequency === "Weekly" || form.recurrenceFrequency === "Fortnightly") && (
                <div>
                  <label className="block text-xs font-bold text-gray-500 uppercase mb-2">Weekday hours</label>
                  <div className="grid grid-cols-2 sm:grid-cols-4 gap-2">
                    {WEEKDAY_OPTIONS.map((d) => {
                      const wh = ((form as any).weekdayHours ?? {}) as Record<string, number>;
                      const current = wh[String(d.id)] ?? 0;
                      return (
                        <div key={d.id} className="flex items-center justify-between gap-2 border border-[#edeef0] rounded-lg bg-white px-3 py-2">
                          <span className="text-xs font-bold text-gray-700">{d.label}</span>
                          <input
                            type="number"
                            min={0}
                            step={0.1}
                            value={current === 0 ? "" : String(current)}
                            onChange={(e) => {
                              const v = e.target.value === "" ? 0 : parseFloat(e.target.value) || 0;
                              const next = { ...(wh || {}) };
                              next[String(d.id)] = v;
                              // Keep recurrenceWeekdays in sync for back-compat.
                              const selected = Object.entries(next)
                                .filter(([, hrs]) => typeof hrs === "number" && hrs > 0)
                                .map(([k]) => parseInt(k, 10))
                                .filter((n) => Number.isFinite(n));
                              setForm((f) => ({ ...(f as any), weekdayHours: next, recurrenceWeekdays: selected }));
                            }}
                            className="w-20 border border-[#edeef0] rounded-md px-2 py-1 text-sm text-right"
                            placeholder="0"
                          />
                        </div>
                      );
                    })}
                  </div>
                </div>
              )}

              {form.recurrenceFrequency === "Monthly" && (
                <div className="space-y-3">
                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                    {field("Monthly Mode", "monthlyMode", (
                      <select
                        id="monthlyMode"
                        value={form.monthlyMode ?? ""}
                        onChange={(e) => setForm((f) => ({ ...f, monthlyMode: (e.target.value as any) || null }))}
                        className="w-full border border-[#edeef0] rounded-lg px-3 py-2 text-sm bg-white"
                      >
                        <option value="">Select…</option>
                        {MONTHLY_MODE_OPTIONS.map((opt) => (
                          <option key={opt.id} value={opt.id}>{opt.label}</option>
                        ))}
                      </select>
                    ))}
                    {form.monthlyMode === "day_of_month" ? (
                      field("Day of Month (1–31)", "monthlyDom", (
                        <input
                          id="monthlyDom"
                          type="number"
                          min={1}
                          max={31}
                          step={1}
                          value={form.monthlyDayOfMonth ?? ""}
                          onChange={(e) => setForm((f) => ({ ...f, monthlyDayOfMonth: e.target.value === "" ? null : parseInt(e.target.value, 10) || 1 }))}
                          className="w-full border border-[#edeef0] rounded-lg px-3 py-2 text-sm"
                        />
                      ))
                    ) : form.monthlyMode === "nth_weekday" ? (
                      field("Week of Month", "monthlyWom", (
                        <select
                          id="monthlyWom"
                          value={form.monthlyWeekOfMonth ?? ""}
                          onChange={(e) => setForm((f) => ({ ...f, monthlyWeekOfMonth: (e.target.value as any) || null }))}
                          className="w-full border border-[#edeef0] rounded-lg px-3 py-2 text-sm bg-white"
                        >
                          <option value="">Select…</option>
                          {WEEK_OF_MONTH_OPTIONS.map((opt) => (
                            <option key={opt} value={opt}>{opt}</option>
                          ))}
                        </select>
                      ))
                    ) : (
                      <div />
                    )}
                  </div>

                  {form.monthlyMode === "nth_weekday" && (
                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                      {field("Weekday", "monthlyWd", (
                        <select
                          id="monthlyWd"
                          value={form.monthlyWeekday ?? ""}
                          onChange={(e) => setForm((f) => ({ ...f, monthlyWeekday: e.target.value === "" ? null : parseInt(e.target.value, 10) }))}
                          className="w-full border border-[#edeef0] rounded-lg px-3 py-2 text-sm bg-white"
                        >
                          <option value="">Select…</option>
                          {WEEKDAY_OPTIONS.map((d) => (
                            <option key={d.id} value={d.id}>{d.label}</option>
                          ))}
                        </select>
                      ))}
                    </div>
                  )}

                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                    {field("Hours (monthly occurrence)", "monthlyHours", (
                      <input
                        id="monthlyHours"
                        type="number"
                        min={0}
                        step={0.1}
                        value={(form as any).monthlyHours ?? ""}
                        onChange={(e) =>
                          setForm((f) => ({ ...(f as any), monthlyHours: e.target.value === "" ? null : parseFloat(e.target.value) || 0 }))
                        }
                        className="w-full border border-[#edeef0] rounded-lg px-3 py-2 text-sm"
                      />
                    ))}
                  </div>
                </div>
              )}

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                {field("Charge Rate ($/hr)", "chargeRate", (
                  <input
                    id="chargeRate"
                    type="number"
                    min={0}
                    step={0.01}
                    value={form.chargeRatePerHour ?? ""}
                    onChange={(e) =>
                      setForm((prev) =>
                        recalcFinance({
                          ...prev,
                          chargeRatePerHour:
                            e.target.value === "" ? null : parseFloat(e.target.value) || 0,
                        })
                      )
                    }
                    className="w-full border border-[#edeef0] rounded-lg px-3 py-2 text-sm"
                  />
                ))}
                {field("Cost Rate ($/hr)", "costRate", (
                  <input
                    id="costRate"
                    type="number"
                    min={0}
                    step={0.01}
                    value={form.costRatePerHour ?? ""}
                    onChange={(e) =>
                      setForm((prev) =>
                        recalcFinance({
                          ...prev,
                          costRatePerHour:
                            e.target.value === "" ? null : parseFloat(e.target.value) || 0,
                        })
                      )
                    }
                    className="w-full border border-[#edeef0] rounded-lg px-3 py-2 text-sm"
                  />
                ))}
              </div>

              <div className="border-t border-[#edeef0] pt-3">
                <p className="text-xs font-bold text-gray-600 uppercase tracking-widest mb-2">Rate overrides (optional)</p>
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                  {field("Weekday Charge Rate", "wdCharge", (
                    <input
                      id="wdCharge"
                      type="number"
                      min={0}
                      step={0.01}
                      value={(form as any).weekdayChargeRateOverride ?? ""}
                      onChange={(e) => setForm((f) => ({ ...(f as any), weekdayChargeRateOverride: e.target.value === "" ? null : parseFloat(e.target.value) || 0 }))}
                      className="w-full border border-[#edeef0] rounded-lg px-3 py-2 text-sm"
                    />
                  ))}
                  {field("Weekday Cost Rate", "wdCost", (
                    <input
                      id="wdCost"
                      type="number"
                      min={0}
                      step={0.01}
                      value={(form as any).weekdayCostRateOverride ?? ""}
                      onChange={(e) => setForm((f) => ({ ...(f as any), weekdayCostRateOverride: e.target.value === "" ? null : parseFloat(e.target.value) || 0 }))}
                      className="w-full border border-[#edeef0] rounded-lg px-3 py-2 text-sm"
                    />
                  ))}
                  {field("Saturday Charge Rate", "satCharge", (
                    <input
                      id="satCharge"
                      type="number"
                      min={0}
                      step={0.01}
                      value={(form as any).saturdayChargeRateOverride ?? ""}
                      onChange={(e) => setForm((f) => ({ ...(f as any), saturdayChargeRateOverride: e.target.value === "" ? null : parseFloat(e.target.value) || 0 }))}
                      className="w-full border border-[#edeef0] rounded-lg px-3 py-2 text-sm"
                    />
                  ))}
                  {field("Saturday Cost Rate", "satCost", (
                    <input
                      id="satCost"
                      type="number"
                      min={0}
                      step={0.01}
                      value={(form as any).saturdayCostRateOverride ?? ""}
                      onChange={(e) => setForm((f) => ({ ...(f as any), saturdayCostRateOverride: e.target.value === "" ? null : parseFloat(e.target.value) || 0 }))}
                      className="w-full border border-[#edeef0] rounded-lg px-3 py-2 text-sm"
                    />
                  ))}
                  {field("Sunday Charge Rate", "sunCharge", (
                    <input
                      id="sunCharge"
                      type="number"
                      min={0}
                      step={0.01}
                      value={(form as any).sundayChargeRateOverride ?? ""}
                      onChange={(e) => setForm((f) => ({ ...(f as any), sundayChargeRateOverride: e.target.value === "" ? null : parseFloat(e.target.value) || 0 }))}
                      className="w-full border border-[#edeef0] rounded-lg px-3 py-2 text-sm"
                    />
                  ))}
                  {field("Sunday Cost Rate", "sunCost", (
                    <input
                      id="sunCost"
                      type="number"
                      min={0}
                      step={0.01}
                      value={(form as any).sundayCostRateOverride ?? ""}
                      onChange={(e) => setForm((f) => ({ ...(f as any), sundayCostRateOverride: e.target.value === "" ? null : parseFloat(e.target.value) || 0 }))}
                      className="w-full border border-[#edeef0] rounded-lg px-3 py-2 text-sm"
                    />
                  ))}
                  {field("Public Holiday Charge Rate", "phCharge", (
                    <input
                      id="phCharge"
                      type="number"
                      min={0}
                      step={0.01}
                      value={(form as any).publicHolidayChargeRateOverride ?? ""}
                      onChange={(e) => setForm((f) => ({ ...(f as any), publicHolidayChargeRateOverride: e.target.value === "" ? null : parseFloat(e.target.value) || 0 }))}
                      className="w-full border border-[#edeef0] rounded-lg px-3 py-2 text-sm"
                    />
                  ))}
                  {field("Public Holiday Cost Rate", "phCost", (
                    <input
                      id="phCost"
                      type="number"
                      min={0}
                      step={0.01}
                      value={(form as any).publicHolidayCostRateOverride ?? ""}
                      onChange={(e) => setForm((f) => ({ ...(f as any), publicHolidayCostRateOverride: e.target.value === "" ? null : parseFloat(e.target.value) || 0 }))}
                      className="w-full border border-[#edeef0] rounded-lg px-3 py-2 text-sm"
                    />
                  ))}
                </div>
              </div>
            </div>
          )}
          <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
            {field("Charge ($)", "charge", (
              <input
                id="charge"
                type="text"
                inputMode="decimal"
                value={fmt2(isRecurring && recurringTotals ? recurringTotals.charge : form.charge)}
                readOnly
                className="w-full border border-[#edeef0] rounded-lg px-3 py-2 text-sm bg-gray-50"
              />
            ))}
            {field("Cost ($)", "cost", (
              <input
                id="cost"
                type="text"
                inputMode="decimal"
                value={fmt2(isRecurring && recurringTotals ? recurringTotals.cost : form.cost)}
                readOnly
                className="w-full border border-[#edeef0] rounded-lg px-3 py-2 text-sm bg-gray-50"
              />
            ))}
            {isAdmin ? (
              field("Gross Profit ($)", "gp", (
                <input
                  id="gp"
                  type="text"
                  inputMode="decimal"
                  value={fmt2(isRecurring && recurringTotals ? recurringTotals.grossProfit : form.grossProfit)}
                  readOnly
                  className="w-full border border-[#edeef0] rounded-lg px-3 py-2 text-sm bg-gray-50"
                />
              ))
            ) : (
              <div />
            )}
          </div>
          {isAdmin && (
            <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
              {field("Markup %", "markup", (
                <input
                  id="markup"
                  type="text"
                  inputMode="decimal"
                  value={fmt2(isRecurring && recurringTotals ? recurringTotals.markupPercent : form.markupPercent)}
                  readOnly
                  className="w-full border border-[#edeef0] rounded-lg px-3 py-2 text-sm bg-gray-50"
                />
              ))}
              {field("GP %", "gpPct", (
                <input
                  id="gpPct"
                  type="text"
                  inputMode="decimal"
                  value={fmt2(isRecurring && recurringTotals ? recurringTotals.gpPercent : form.gpPercent)}
                  readOnly
                  className="w-full border border-[#edeef0] rounded-lg px-3 py-2 text-sm bg-gray-50"
                />
              ))}
            </div>
          )}
          {field("Description", "desc", (
            <textarea id="desc" value={form.description ?? ""} onChange={(e) => setForm((f) => ({ ...f, description: e.target.value }))} className="w-full border border-[#edeef0] rounded-lg px-3 py-2 text-sm" rows={2} />
          ))}
          <div className="mb-4">
            <label className="block text-xs font-bold text-gray-500 uppercase mb-1">Approval proof (screenshot or PDF)</label>
            <p className="text-xs text-gray-500 mb-2">
              Upload proof of client confirmation. Drag and drop, click to add files, or use
              &nbsp;
              <button
                type="button"
                className="underline text-xs text-gray-700 hover:text-gray-900"
                onClick={() => {
                  setPasteMode(true);
                  setPasteError(null);
                }}
              >
                Paste Screenshot
              </button>
              .
            </p>
            <div
              className={`border-2 border-dashed rounded-lg p-6 text-center transition-colors ${
                dragOver ? "border-blue-500 bg-blue-50" : "border-[#edeef0] bg-gray-50/50"
              }`}
              onDragOver={(e) => { e.preventDefault(); e.stopPropagation(); setDragOver(true); }}
              onDragLeave={(e) => { e.preventDefault(); e.stopPropagation(); setDragOver(false); }}
              onDrop={(e) => {
                e.preventDefault();
                e.stopPropagation();
                setDragOver(false);
                acceptProofFiles(e.dataTransfer.files);
              }}
            >
              <input
                type="file"
                accept=".pdf,image/*"
                multiple
                className="hidden"
                id="proof-file-input"
                onChange={(e) => { acceptProofFiles(e.target.files); e.target.value = ""; }}
              />
              <label htmlFor="proof-file-input" className="cursor-pointer flex flex-col items-center gap-2">
                <Upload className="text-gray-400" size={28} />
                <span className="text-sm text-gray-600">Drop files here or click to browse</span>
                <span className="text-xs text-gray-400">PDF or images (screenshots)</span>
              </label>
            </div>
            {existingAttachments.length > 0 && (
              <div className="mt-2 flex flex-wrap gap-2">
                {existingAttachments.map((a) => (
                  <div
                    key={a.url}
                    className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full text-xs font-medium bg-green-50 text-green-800 border border-green-200 max-w-full"
                    title={a.fileName}
                  >
                    <button
                      type="button"
                      onClick={() => window.open(a.url, "_blank", "noopener")}
                      className="inline-flex items-center gap-1 mr-1"
                    >
                      <FileText size={12} className="text-green-600 flex-shrink-0" />
                      <span className="truncate max-w-[120px]">{a.fileName}</span>
                    </button>
                    <button
                      type="button"
                      onClick={async () => {
                        const token = await getGraphAccessToken();
                        if (!token || !job) return;
                        await deleteAdHocJobAttachment(token, job.id, a.fileName);
                        setExistingAttachments((prev) =>
                          prev.filter((x) => x.fileName !== a.fileName)
                        );
                      }}
                      className="text-gray-400 hover:text-red-600 flex-shrink-0"
                      aria-label={`Delete ${a.fileName}`}
                    >
                      <X size={12} />
                    </button>
                  </div>
                ))}
              </div>
            )}
            {(pasteMode || proofFiles.length > 0) && (
              <div className="mt-2">
                <p className="text-xs text-gray-600 mb-1">
                  Click in the field and press <span className="font-semibold">Ctrl+V</span> to paste a screenshot from your clipboard.
                </p>
                <div className="w-full border border-dashed border-[#cbd5e1] rounded-md px-2 py-1 text-xs text-gray-600 bg-white flex flex-wrap gap-2 items-center min-h-[40px]">
                  {proofFiles.map((f, i) => (
                    <div
                      key={i}
                      className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full text-xs font-medium bg-gray-100 text-gray-800 border border-gray-200 max-w-full"
                      title={f.name}
                    >
                      <FileText size={12} className="text-gray-500 flex-shrink-0" />
                      <span className="truncate max-w-[140px]">{f.name}</span>
                      <button
                        type="button"
                        onClick={() =>
                          setProofFiles((prev) => prev.filter((_, idx) => idx !== i))
                        }
                        className="text-gray-400 hover:text-red-600 flex-shrink-0"
                        aria-label={`Remove ${f.name}`}
                      >
                        <X size={12} />
                      </button>
                    </div>
                  ))}
                  <textarea
                    onPaste={handleProofPaste}
                    autoFocus
                    rows={1}
                    className="flex-1 min-w-[120px] border-none outline-none text-xs text-gray-500 bg-transparent resize-none"
                    placeholder={
                      proofFiles.length === 0
                        ? "Click here and press Ctrl+V to paste a screenshot…"
                        : "Press Ctrl+V to add another screenshot…"
                    }
                  />
                </div>
                {pasteError && (
                  <p className="mt-1 text-xs text-amber-600">
                    {pasteError}
                  </p>
                )}
              </div>
            )}
          </div>
          <div className="flex flex-wrap gap-6">
            <label className="flex items-center gap-2">
              <input type="checkbox" checked={form.approvalProofRequired ?? false} onChange={(e) => setForm((f) => ({ ...f, approvalProofRequired: e.target.checked }))} className="rounded border-gray-300" />
              <span className="text-sm">Approval proof required</span>
            </label>
            <label className="flex items-center gap-2">
              <input type="checkbox" checked={form.approvalProofUploaded ?? false} onChange={(e) => setForm((f) => ({ ...f, approvalProofUploaded: e.target.checked }))} className="rounded border-gray-300" />
              <span className="text-sm">Approval proof uploaded (or upload files above)</span>
            </label>
            <label className="flex items-center gap-2">
              <input type="checkbox" checked={form.active ?? true} onChange={(e) => setForm((f) => ({ ...f, active: e.target.checked }))} className="rounded border-gray-300" />
              <span className="text-sm">Active</span>
            </label>
          </div>
          {field("Approval reference notes", "proofNotes", (
            <textarea id="proofNotes" value={form.approvalReferenceNotes ?? ""} onChange={(e) => setForm((f) => ({ ...f, approvalReferenceNotes: e.target.value }))} className="w-full border border-[#edeef0] rounded-lg px-3 py-2 text-sm" rows={2} placeholder="Reference to email / ticket" />
          ))}
          <div className="flex justify-end gap-2 pt-4 border-t border-[#edeef0]">
            <button type="button" onClick={onClose} className="px-4 py-2 rounded-lg border border-[#edeef0] text-gray-700 hover:bg-gray-50">Cancel</button>
            <button type="submit" disabled={submitLoading || !form.jobName?.trim()} className="px-4 py-2 rounded-lg bg-gray-900 text-white font-medium hover:bg-gray-800 disabled:opacity-50 flex items-center gap-2">
              {submitLoading && <Loader2 className="animate-spin" size={18} />}
              {isEdit ? "Update" : "Create"}
            </button>
          </div>
        </form>
      </div>
    </div>
  );
}
