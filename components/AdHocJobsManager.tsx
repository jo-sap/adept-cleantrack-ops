/**
 * Ad Hoc Jobs – list, filters (month, status, manager, site), create/edit form.
 * Admin: all jobs. Manager: jobs assigned to them.
 */
import React, { useState, useEffect, useCallback } from "react";
import { AdHocJob } from "../types";
import { Plus, X, Loader2, Pencil, AlertCircle } from "lucide-react";
import { useRole } from "../contexts/RoleContext";
import { useAppAuth } from "../contexts/AppAuthContext";
import { getGraphAccessToken } from "../lib/graph";
import { getAdHocJobs, createAdHocJob, updateAdHocJob, type AdHocJobPayload, type AdHocJobFilters } from "../repositories/adHocJobsRepo";
import { getCleanTrackUserByEmail, getCleanTrackManagers } from "../repositories/usersRepo";
import { getSites } from "../repositories/sitesRepo";
import type { Site } from "../repositories/sitesRepo";
import { format } from "date-fns";

const STATUS_OPTIONS = ["Requested", "Approved", "Scheduled", "Completed"];
const JOB_TYPE_PLACEHOLDER = "e.g. Carpet clean, Emergency, One-off";

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
        <button
          type="button"
          onClick={openCreate}
          className="flex items-center gap-2 bg-gray-900 text-white px-4 py-2.5 rounded-lg text-sm font-bold hover:bg-gray-800"
        >
          <Plus size={18} />
          New Ad Hoc Job
        </button>
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
        <div className="border border-[#edeef0] rounded-lg bg-white shadow-sm overflow-hidden table-scroll-mobile">
          <table className="w-full border-collapse text-left min-w-0 table-auto md:table-fixed md:min-w-[800px]">
            <thead>
              <tr className="bg-[#fcfcfb] border-b border-[#edeef0]">
                <th className="px-2 py-2 text-[9px] font-bold text-gray-500 uppercase tracking-widest">Job Name</th>
                <th className="px-2 py-2 text-[9px] font-bold text-gray-500 uppercase tracking-widest hidden md:table-cell">Site</th>
                <th className="px-2 py-2 text-[9px] font-bold text-gray-500 uppercase tracking-widest hidden md:table-cell">Assigned Manager</th>
                <th className="px-2 py-2 text-[9px] font-bold text-gray-500 uppercase tracking-widest">Requested</th>
                <th className="px-2 py-2 text-[9px] font-bold text-gray-500 uppercase tracking-widest hidden md:table-cell">Scheduled</th>
                <th className="px-2 py-2 text-[9px] font-bold text-gray-500 uppercase tracking-widest hidden md:table-cell">Completed</th>
                <th className="px-2 py-2 text-[9px] font-bold text-gray-500 uppercase tracking-widest">Status</th>
                <th className="px-2 py-2 text-[9px] font-bold text-gray-500 uppercase tracking-widest">Budgeted Hrs</th>
                <th className="px-2 py-2 text-[9px] font-bold text-gray-500 uppercase tracking-widest hidden md:table-cell">Budgeted Revenue</th>
                <th className="px-2 py-2 text-[9px] font-bold text-gray-500 uppercase tracking-widest">Proof</th>
                <th className="px-2 py-2 text-[9px] font-bold text-gray-500 uppercase tracking-widest w-20">Actions</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-[#edeef0]">
              {jobs.map((j) => (
                <tr key={j.id} className="hover:bg-[#f7f6f3]">
                  <td className="px-2 py-2">
                    <span className="text-sm font-semibold text-gray-900">{j.jobName || "—"}</span>
                  </td>
                  <td className="px-2 py-2 text-xs text-gray-600 hidden md:table-cell">{j.siteName || "—"}</td>
                  <td className="px-2 py-2 text-xs text-gray-600 hidden md:table-cell">{j.assignedManagerName || "—"}</td>
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
                  <td className="px-2 py-2 text-xs font-medium text-gray-700">{j.budgetedHours != null ? `${j.budgetedHours}h` : "—"}</td>
                  <td className="px-2 py-2 text-xs text-gray-700 hidden md:table-cell">{j.budgetedRevenue != null ? `$${Number(j.budgetedRevenue).toFixed(2)}` : "—"}</td>
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
                    <button
                      type="button"
                      onClick={() => openEdit(j)}
                      className="p-2 rounded text-blue-600 hover:bg-blue-50"
                      aria-label={`Edit ${j.jobName}`}
                    >
                      <Pencil size={16} />
                    </button>
                  </td>
                </tr>
              ))}
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
          onSaved={() => { loadJobs(); closeModal(); showToast(editingJob ? "Job updated." : "Ad hoc job created."); }}
          submitLoading={submitLoading}
          setSubmitLoading={setSubmitLoading}
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
}: AdHocJobFormModalProps) {
  const isEdit = !!job;
  const [form, setForm] = useState<AdHocJobPayload>({
    jobName: job?.jobName ?? "",
    jobType: "",
    siteId: job?.siteId ?? null,
    assignedManagerId: job?.assignedManagerId ?? currentUserId ?? null,
    requestedByName: job?.requestedByName ?? "",
    requestedByEmail: job?.requestedByEmail ?? "",
    requestedByCompany: job?.requestedByCompany ?? "",
    requestChannel: job?.requestChannel ?? "",
    requestSummary: job?.requestSummary ?? "",
    requestedDate: job?.requestedDate ?? format(new Date(), "yyyy-MM-dd"),
    scheduledDate: job?.scheduledDate ?? null,
    completedDate: job?.completedDate ?? null,
    status: job?.status ?? "Requested",
    budgetedHours: job?.budgetedHours ?? null,
    budgetedLabourRate: job?.budgetedLabourRate ?? null,
    budgetedRevenue: job?.budgetedRevenue ?? null,
    description: job?.description ?? "",
    approvalProofRequired: job?.approvalProofRequired ?? false,
    approvalProofUploaded: job?.approvalProofUploaded ?? false,
    approvalReferenceNotes: job?.approvalReferenceNotes ?? "",
    active: job?.active ?? true,
  });

  useEffect(() => {
    if (job) {
      setForm({
        jobName: job.jobName ?? "",
        jobType: job.jobType ?? "",
        siteId: job.siteId ?? null,
        assignedManagerId: job.assignedManagerId ?? null,
        requestedByName: job.requestedByName ?? "",
        requestedByEmail: job.requestedByEmail ?? "",
        requestedByCompany: job.requestedByCompany ?? "",
        requestChannel: job.requestChannel ?? "",
        requestSummary: job.requestSummary ?? "",
        requestedDate: job.requestedDate ?? null,
        scheduledDate: job.scheduledDate ?? null,
        completedDate: job.completedDate ?? null,
        status: job.status ?? "Requested",
        budgetedHours: job.budgetedHours ?? null,
        budgetedLabourRate: job.budgetedLabourRate ?? null,
        budgetedRevenue: job.budgetedRevenue ?? null,
        description: job.description ?? "",
        approvalProofRequired: job.approvalProofRequired ?? false,
        approvalProofUploaded: job.approvalProofUploaded ?? false,
        approvalReferenceNotes: job.approvalReferenceNotes ?? "",
        active: job.active ?? true,
      });
    }
  }, [job]);

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!form.jobName?.trim()) return;
    const token = await getGraphAccessToken();
    if (!token) return;
    setSubmitLoading(true);
    try {
      if (isEdit && job) {
        await updateAdHocJob(token, job.id, form);
      } else {
        await createAdHocJob(token, { ...form, status: form.status || "Requested" });
      }
      onSaved();
    } catch (err) {
      alert(err instanceof Error ? err.message : "Save failed.");
    } finally {
      setSubmitLoading(false);
    }
  };

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
          {field("Job Type", "jobType", (
            <input
              id="jobType"
              type="text"
              value={form.jobType ?? ""}
              onChange={(e) => setForm((f) => ({ ...f, jobType: e.target.value }))}
              className="w-full border border-[#edeef0] rounded-lg px-3 py-2 text-sm"
              placeholder={JOB_TYPE_PLACEHOLDER}
            />
          ))}
          {field("Site", "site", (
            <select
              id="site"
              value={form.siteId ?? ""}
              onChange={(e) => setForm((f) => ({ ...f, siteId: e.target.value || null }))}
              className="w-full border border-[#edeef0] rounded-lg px-3 py-2 text-sm"
            >
              <option value="">— None —</option>
              {sites.map((s) => (
                <option key={s.id} value={s.id}>{s.siteName || s.address || s.id}</option>
              ))}
            </select>
          ))}
          {field("Requested By Name", "reqName", (
            <input id="reqName" type="text" value={form.requestedByName ?? ""} onChange={(e) => setForm((f) => ({ ...f, requestedByName: e.target.value }))} className="w-full border border-[#edeef0] rounded-lg px-3 py-2 text-sm" />
          ))}
          {field("Requested By Email", "reqEmail", (
            <input id="reqEmail" type="email" value={form.requestedByEmail ?? ""} onChange={(e) => setForm((f) => ({ ...f, requestedByEmail: e.target.value }))} className="w-full border border-[#edeef0] rounded-lg px-3 py-2 text-sm" />
          ))}
          {field("Requested By Company", "reqCompany", (
            <input id="reqCompany" type="text" value={form.requestedByCompany ?? ""} onChange={(e) => setForm((f) => ({ ...f, requestedByCompany: e.target.value }))} className="w-full border border-[#edeef0] rounded-lg px-3 py-2 text-sm" />
          ))}
          {field("Request Channel", "channel", (
            <input id="channel" type="text" value={form.requestChannel ?? ""} onChange={(e) => setForm((f) => ({ ...f, requestChannel: e.target.value }))} className="w-full border border-[#edeef0] rounded-lg px-3 py-2 text-sm" placeholder="e.g. Email, Phone" />
          ))}
          {field("Request Summary", "summary", (
            <textarea id="summary" value={form.requestSummary ?? ""} onChange={(e) => setForm((f) => ({ ...f, requestSummary: e.target.value }))} className="w-full border border-[#edeef0] rounded-lg px-3 py-2 text-sm" rows={2} />
          ))}
          <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
            {field("Requested Date", "reqDate", (
              <input id="reqDate" type="date" value={form.requestedDate ?? ""} onChange={(e) => setForm((f) => ({ ...f, requestedDate: e.target.value || null }))} className="w-full border border-[#edeef0] rounded-lg px-3 py-2 text-sm" />
            ))}
            {field("Assigned Manager", "manager", (
              <select id="manager" value={form.assignedManagerId ?? ""} onChange={(e) => setForm((f) => ({ ...f, assignedManagerId: e.target.value || null }))} className="w-full border border-[#edeef0] rounded-lg px-3 py-2 text-sm">
                <option value="">— None —</option>
                {managers.map((m) => (
                  <option key={m.id} value={m.id}>{m.fullName || m.email}</option>
                ))}
              </select>
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
            {field("Scheduled Date", "schedDate", (
              <input id="schedDate" type="date" value={form.scheduledDate ?? ""} onChange={(e) => setForm((f) => ({ ...f, scheduledDate: e.target.value || null }))} className="w-full border border-[#edeef0] rounded-lg px-3 py-2 text-sm" />
            ))}
            {field("Completed Date", "compDate", (
              <input id="compDate" type="date" value={form.completedDate ?? ""} onChange={(e) => setForm((f) => ({ ...f, completedDate: e.target.value || null }))} className="w-full border border-[#edeef0] rounded-lg px-3 py-2 text-sm" />
            ))}
          </div>
          <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
            {field("Budgeted Hours", "budgHrs", (
              <input id="budgHrs" type="number" min={0} step={0.5} value={form.budgetedHours ?? ""} onChange={(e) => setForm((f) => ({ ...f, budgetedHours: e.target.value === "" ? null : parseFloat(e.target.value) || 0 }))} className="w-full border border-[#edeef0] rounded-lg px-3 py-2 text-sm" />
            ))}
            {field("Budgeted Labour Rate ($/hr)", "budgRate", (
              <input id="budgRate" type="number" min={0} step={0.01} value={form.budgetedLabourRate ?? ""} onChange={(e) => setForm((f) => ({ ...f, budgetedLabourRate: e.target.value === "" ? null : parseFloat(e.target.value) || 0 }))} className="w-full border border-[#edeef0] rounded-lg px-3 py-2 text-sm" />
            ))}
            {field("Budgeted Revenue ($)", "budgRev", (
              <input id="budgRev" type="number" min={0} step={0.01} value={form.budgetedRevenue ?? ""} onChange={(e) => setForm((f) => ({ ...f, budgetedRevenue: e.target.value === "" ? null : parseFloat(e.target.value) || 0 }))} className="w-full border border-[#edeef0] rounded-lg px-3 py-2 text-sm" />
            ))}
          </div>
          {field("Description", "desc", (
            <textarea id="desc" value={form.description ?? ""} onChange={(e) => setForm((f) => ({ ...f, description: e.target.value }))} className="w-full border border-[#edeef0] rounded-lg px-3 py-2 text-sm" rows={2} />
          ))}
          <div className="flex flex-wrap gap-6">
            <label className="flex items-center gap-2">
              <input type="checkbox" checked={form.approvalProofRequired ?? false} onChange={(e) => setForm((f) => ({ ...f, approvalProofRequired: e.target.checked }))} className="rounded border-gray-300" />
              <span className="text-sm">Approval proof required</span>
            </label>
            <label className="flex items-center gap-2">
              <input type="checkbox" checked={form.approvalProofUploaded ?? false} onChange={(e) => setForm((f) => ({ ...f, approvalProofUploaded: e.target.checked }))} className="rounded border-gray-300" />
              <span className="text-sm">Approval proof uploaded (e.g. attachment in SharePoint)</span>
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
