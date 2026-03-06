import React, { useState, useEffect, useCallback } from "react";
import { Cleaner } from "../types";
import { Plus, CreditCard, X, UserCircle, Search, Loader2, Pencil, Layers, Trash2 } from "lucide-react";
import { useRole } from "../contexts/RoleContext";
import { useAppAuth } from "../contexts/AppAuthContext";
import { getGraphAccessToken } from "../lib/graph";
import {
  getCleaners,
  createCleaner,
  updateCleaner,
  deleteCleaner,
  type CleanerItem,
  type CleanerPayload,
} from "../repositories/cleanersRepo";

const WRITE_PERMISSION_HINT =
  "Admin consent may be required for Sites.ReadWrite.All or Lists.ReadWrite.All.";

/** Map SharePoint CleanerItem to app Cleaner type (for Dashboard/Timesheets compatibility). */
function toAppCleaner(c: CleanerItem): Cleaner {
  const parts = c.cleanerName.trim().split(/\s+/);
  const firstName = parts[0] ?? c.cleanerName;
  const lastName = parts.slice(1).join(" ") ?? "";
  return {
    id: c.id,
    firstName,
    lastName,
    email: "",
    phone: "",
    bankAccountName: c.accountName,
    bankBsb: c.bsb,
    bankAccountNumber: c.accountNumber,
    payRatePerHour: c.payRatePerHour,
  };
}

interface CleanerManagerProps {
  /** Called after cleaners list changes so parent can refresh its state (e.g. for Dashboard/Timesheets). */
  onCleanersRefresh?: () => void;
}

interface BulkCleanerRow {
  id: string;
  cleanerName: string;
  payRatePerHour: number | "";
  accountName: string;
  bsb: string;
  accountNumber: string;
  active: boolean;
}

const CleanerManager: React.FC<CleanerManagerProps> = ({ onCleanersRefresh }) => {
  const { isAdmin: isAdminFromRole } = useRole();
  const { user } = useAppAuth();
  const isAdmin = isAdminFromRole || user?.role === "Admin";

  const [cleaners, setCleaners] = useState<CleanerItem[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [toast, setToast] = useState<string | null>(null);
  const [isAdding, setIsAdding] = useState(false);
  const [searchQuery, setSearchQuery] = useState("");
  const [editingCleaner, setEditingCleaner] = useState<CleanerItem | null>(null);
  const [formData, setFormData] = useState<CleanerPayload>({
    cleanerName: "",
    payRatePerHour: 25,
    accountName: "",
    bsb: "",
    accountNumber: "",
    active: true,
  });
  const [submitLoading, setSubmitLoading] = useState(false);
  const [submitError, setSubmitError] = useState<string | null>(null);

  const [bulkModalOpen, setBulkModalOpen] = useState(false);
  const [bulkRows, setBulkRows] = useState<BulkCleanerRow[]>([]);
  const [bulkSubmitLoading, setBulkSubmitLoading] = useState(false);

  const showToast = useCallback((msg: string) => {
    setToast(msg);
    setTimeout(() => setToast(null), 4000);
  }, []);

  const loadCleaners = useCallback(async () => {
    const token = await getGraphAccessToken();
    if (!token) {
      setError("Sign in with Microsoft to view cleaners.");
      setCleaners([]);
      setLoading(false);
      return;
    }
    setLoading(true);
    setError(null);
    try {
      const data = await getCleaners(token);
      setCleaners(data);
    } catch (e) {
      const msg = e instanceof Error ? e.message : "Failed to load cleaners.";
      setError(msg);
      setCleaners([]);
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    loadCleaners();
  }, [loadCleaners]);

  const handleOpenAdd = () => {
    setIsAdding(true);
    setEditingCleaner(null);
    setFormData({
      cleanerName: "",
      payRatePerHour: 25,
      accountName: "",
      bsb: "",
      accountNumber: "",
      active: true,
    });
    setSubmitError(null);
  };

  const handleOpenBulk = () => {
    setBulkModalOpen(true);
    setBulkRows([
      {
        id: crypto.randomUUID(),
        cleanerName: "",
        payRatePerHour: 25,
        accountName: "",
        bsb: "",
        accountNumber: "",
        active: true,
      },
    ]);
  };

  const updateBulkRow = (id: string, patch: Partial<BulkCleanerRow>) => {
    setBulkRows((rows) => rows.map((r) => (r.id === id ? { ...r, ...patch } : r)));
  };

  const addBulkRow = () => {
    setBulkRows((rows) => [
      ...rows,
      {
        id: crypto.randomUUID(),
        cleanerName: "",
        payRatePerHour: 25,
        accountName: "",
        bsb: "",
        accountNumber: "",
        active: true,
      },
    ]);
  };

  const removeBulkRow = (id: string) => {
    setBulkRows((rows) => (rows.length <= 1 ? rows : rows.filter((r) => r.id !== id)));
  };

  const handleBulkSubmit = async () => {
    const rowsToCreate = bulkRows.filter((r) => r.cleanerName.trim() !== "");
    if (rowsToCreate.length === 0) {
      setBulkModalOpen(false);
      return;
    }
    const token = await getGraphAccessToken();
    if (!token) {
      setError("Not signed in. Sign in with Microsoft to add cleaners.");
      return;
    }
    setBulkSubmitLoading(true);
    try {
      for (const row of rowsToCreate) {
        await createCleaner(token, {
          cleanerName: row.cleanerName.trim(),
          payRatePerHour: row.payRatePerHour === "" ? 25 : (row.payRatePerHour ?? 25),
          accountName: row.accountName?.trim() ?? "",
          bsb: row.bsb?.trim() ?? "",
          accountNumber: row.accountNumber?.trim() ?? "",
          active: row.active !== false,
        });
      }
      showToast(`Added ${rowsToCreate.length} cleaner(s).`);
      setBulkModalOpen(false);
      await loadCleaners();
      onCleanersRefresh?.();
    } catch (e) {
      const msg = e instanceof Error ? e.message : "Bulk add failed.";
      setError(msg);
    } finally {
      setBulkSubmitLoading(false);
    }
  };

  const handleOpenEdit = (cleaner: CleanerItem) => {
    setEditingCleaner(cleaner);
    setIsAdding(false);
    setFormData({
      cleanerName: cleaner.cleanerName,
      payRatePerHour: cleaner.payRatePerHour ?? 25,
      accountName: cleaner.accountName ?? "",
      bsb: cleaner.bsb ?? "",
      accountNumber: cleaner.accountNumber ?? "",
      active: cleaner.active !== false,
    });
    setSubmitError(null);
  };

  const handleCloseEdit = () => {
    setEditingCleaner(null);
    setFormData({
      cleanerName: "",
      payRatePerHour: 25,
      accountName: "",
      bsb: "",
      accountNumber: "",
      active: true,
    });
  };

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!formData.cleanerName.trim()) {
      setSubmitError("Cleaner Name is required.");
      return;
    }
    const token = await getGraphAccessToken();
    if (!token) {
      setSubmitError("Not signed in. Sign in with Microsoft to add a cleaner.");
      return;
    }
    setSubmitLoading(true);
    setSubmitError(null);
    try {
      if (editingCleaner) {
        await updateCleaner(token, editingCleaner.id, {
          cleanerName: formData.cleanerName.trim(),
          payRatePerHour: formData.payRatePerHour ?? 25,
          accountName: formData.accountName?.trim() ?? "",
          bsb: formData.bsb?.trim() ?? "",
          accountNumber: formData.accountNumber?.trim() ?? "",
          active: formData.active !== false,
        });
        showToast("Cleaner updated.");
        setEditingCleaner(null);
      } else {
        await createCleaner(token, {
          cleanerName: formData.cleanerName.trim(),
          payRatePerHour: formData.payRatePerHour ?? 25,
          accountName: formData.accountName?.trim() ?? "",
          bsb: formData.bsb?.trim() ?? "",
          accountNumber: formData.accountNumber?.trim() ?? "",
          active: formData.active !== false,
        });
        showToast("Cleaner added.");
        setIsAdding(false);
      }
      await loadCleaners();
      onCleanersRefresh?.();
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

  const handleToggleActive = async (cleaner: CleanerItem) => {
    const token = await getGraphAccessToken();
    if (!token) {
      setError("Not signed in. Sign in with Microsoft to update cleaners.");
      return;
    }
    try {
      await updateCleaner(token, cleaner.id, { active: !cleaner.active });
      await loadCleaners();
      onCleanersRefresh?.();
      showToast(`Cleaner ${!cleaner.active ? "activated" : "deactivated"}.`);
    } catch (e) {
      const msg = e instanceof Error ? e.message : "Failed to update cleaner.";
      setError(msg);
    }
  };

  const handleDeleteCleaner = async (cleaner: CleanerItem) => {
    if (!window.confirm(`Delete cleaner "${cleaner.cleanerName}"? Existing timesheet entries will still reference this cleaner.`)) {
      return;
    }
    const token = await getGraphAccessToken();
    if (!token) {
      setError("Not signed in. Sign in with Microsoft to delete cleaners.");
      return;
    }
    try {
      await deleteCleaner(token, cleaner.id);
      await loadCleaners();
      onCleanersRefresh?.();
      showToast("Cleaner deleted.");
    } catch (e) {
      const msg = e instanceof Error ? e.message : "Failed to delete cleaner.";
      setError(msg);
    }
  };

  const filteredCleaners = cleaners.filter((c) =>
    c.cleanerName.toLowerCase().includes(searchQuery.toLowerCase().trim())
  );

  return (
    <div className="space-y-8 animate-fadeIn">
      <div className="flex flex-col md:flex-row justify-between items-start md:items-end gap-4 border-b border-[#edeef0] pb-6">
        <div>
          <h2 className="text-3xl font-bold text-gray-900">Cleaner Team</h2>
          <p className="text-gray-500 text-sm mt-1">Manage personnel, onboarding details, and banking records.</p>
        </div>
        <div className="flex gap-2 w-full md:w-auto">
          <div className="relative flex-1 md:w-64">
            <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" size={16} />
            <input
              type="text"
              placeholder="Search team..."
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              className="w-full pl-10 pr-4 py-2 bg-white border border-[#edeef0] rounded-md text-sm outline-none focus:ring-1 focus:ring-gray-900 transition-all"
            />
          </div>
          {isAdmin && (
            <>
              <button
                onClick={handleOpenBulk}
                className="bg-white text-gray-700 px-4 py-2 rounded-md text-sm font-semibold border border-[#edeef0] hover:bg-gray-50 transition-colors flex items-center gap-2"
              >
                <Layers size={16} /> Bulk Add
              </button>
              <button
                onClick={handleOpenAdd}
                className="bg-gray-900 text-white px-6 py-2 rounded-md text-sm font-semibold hover:bg-gray-800 transition-colors flex items-center gap-2"
              >
                <Plus size={18} /> New Cleaner
              </button>
            </>
          )}
        </div>
      </div>

      {toast && (
        <div className="bg-green-50 border border-green-200 text-green-800 px-4 py-2 rounded-lg text-sm">
          {toast}
        </div>
      )}

      {loading ? (
        <div className="flex items-center gap-2 text-gray-500">
          <Loader2 className="animate-spin" size={20} /> Loading cleaners…
        </div>
      ) : error ? (
        <div className="bg-amber-50 border border-amber-200 text-amber-800 px-4 py-3 rounded-lg">
          {error}
        </div>
      ) : filteredCleaners.length === 0 ? (
        <div className="text-gray-500 py-8">
          {searchQuery.trim() ? "No cleaners match your search." : "No cleaners yet. Add one to get started."}
        </div>
      ) : (
        <div className="border border-[#edeef0] rounded-lg overflow-hidden bg-white">
          <table className="w-full text-left">
            <thead>
              <tr className="bg-gray-50 border-b border-[#edeef0] text-[10px] font-bold text-gray-500 uppercase tracking-widest">
                <th className="py-3 px-4 w-12"></th>
                <th className="py-3 px-4">Name</th>
                <th className="py-3 px-4 w-20">Status</th>
                <th className="py-3 px-4 w-24">Rate</th>
                <th className="py-3 px-4">Account name</th>
                <th className="py-3 px-4 w-24">BSB</th>
                <th className="py-3 px-4 w-28">Account no.</th>
                {isAdmin && <th className="py-3 px-4 w-20 text-right">Actions</th>}
              </tr>
            </thead>
            <tbody>
              {filteredCleaners.map((cleaner) => {
                const appCleaner = toAppCleaner(cleaner);
                return (
                  <tr
                    key={cleaner.id}
                    className="border-b border-[#edeef0] last:border-b-0 hover:bg-gray-50/80 transition-colors"
                  >
                    <td className="py-2.5 px-4">
                      <div className="w-9 h-9 bg-gray-100 rounded-lg flex items-center justify-center text-gray-600 font-bold text-sm">
                        {appCleaner.firstName.charAt(0)}
                        {appCleaner.lastName ? appCleaner.lastName.charAt(0) : ""}
                      </div>
                    </td>
                    <td className="py-2.5 px-4">
                      <span className="font-semibold text-gray-900">{cleaner.cleanerName}</span>
                    </td>
                    <td className="py-2.5 px-4">
                      <span className={`text-[11px] font-bold uppercase ${cleaner.active ? "text-green-600" : "text-gray-400"}`}>
                        {cleaner.active ? "Active" : "Inactive"}
                      </span>
                    </td>
                    <td className="py-2.5 px-4">
                      <span className="text-sm font-semibold text-green-600">${cleaner.payRatePerHour || 0}/hr</span>
                    </td>
                    <td className="py-2.5 px-4 text-sm text-gray-600 truncate max-w-[160px]" title={cleaner.accountName || undefined}>
                      {cleaner.accountName || "—"}
                    </td>
                    <td className="py-2.5 px-4 text-sm text-gray-600 font-mono">
                      {cleaner.bsb || "—"}
                    </td>
                    <td className="py-2.5 px-4 text-sm text-gray-600 font-mono">
                      {cleaner.accountNumber || "—"}
                    </td>
                    {isAdmin && (
                      <td className="py-2.5 px-4 text-right">
                        <div className="flex justify-end gap-2">
                          <button
                            type="button"
                            onClick={() => handleToggleActive(cleaner)}
                            className="inline-flex items-center gap-1 px-2 py-1 text-xs font-medium rounded-md border border-[#edeef0] bg-white hover:bg-gray-50"
                            aria-label={cleaner.active ? `Deactivate ${cleaner.cleanerName}` : `Activate ${cleaner.cleanerName}`}
                          >
                            {cleaner.active ? "Deactivate" : "Activate"}
                          </button>
                          <button
                            type="button"
                            onClick={() => handleOpenEdit(cleaner)}
                            className="inline-flex items-center gap-1 px-2 py-1 text-xs font-medium text-gray-700 bg-white border border-[#edeef0] rounded-md hover:bg-gray-50 hover:border-gray-300 transition-colors"
                            aria-label={`Edit ${cleaner.cleanerName}`}
                          >
                            <Pencil size={14} />
                          </button>
                          <button
                            type="button"
                            onClick={() => handleDeleteCleaner(cleaner)}
                            className="inline-flex items-center gap-1 px-2 py-1 text-xs font-medium text-red-600 bg-white border border-red-100 rounded-md hover:bg-red-50 hover:border-red-300 transition-colors"
                            aria-label={`Delete ${cleaner.cleanerName}`}
                          >
                            <Trash2 size={14} />
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
      )}

      {(isAdding || editingCleaner) && (
        <div className="fixed inset-0 bg-gray-900/60 backdrop-blur-sm z-[100] flex items-center justify-center p-4">
          <div className="bg-white w-full max-w-xl rounded-2xl shadow-2xl overflow-hidden animate-slideUp">
            <div className="p-6 border-b border-[#edeef0] flex justify-between items-center bg-[#fcfcfb]">
              <div className="flex items-center gap-3">
                <div className="w-8 h-8 bg-gray-900 rounded-lg flex items-center justify-center text-white">
                  <UserCircle size={20} />
                </div>
                <h3 className="font-bold text-gray-900">{editingCleaner ? "Edit Cleaner" : "Add Cleaner"}</h3>
              </div>
              <button
                onClick={() => { setIsAdding(false); handleCloseEdit(); }}
                className="text-gray-400 hover:text-gray-900 transition-colors"
              >
                <X size={20} />
              </button>
            </div>

            <form onSubmit={handleSubmit} className="p-8 space-y-6">
              <div className="space-y-1">
                <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">
                  Cleaner Name <span className="text-red-500">*</span>
                </label>
                <input
                  required
                  type="text"
                  value={formData.cleanerName}
                  onChange={(e) => setFormData({ ...formData, cleanerName: e.target.value })}
                  className="w-full px-3 py-2 border border-[#edeef0] rounded-md text-sm outline-none focus:ring-1 focus:ring-gray-900"
                  placeholder="Full name"
                />
              </div>

              <div className="space-y-1">
                <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">
                  Pay Rate ($/hr)
                </label>
                <input
                  type="number"
                  step="0.5"
                  min={0}
                  value={formData.payRatePerHour ?? 25}
                  onChange={(e) =>
                    setFormData({ ...formData, payRatePerHour: parseFloat(e.target.value) || 0 })
                  }
                  className="w-full px-3 py-2 border border-[#edeef0] rounded-md text-sm outline-none focus:ring-1 focus:ring-gray-900"
                />
              </div>

              <div className="bg-gray-50 p-6 rounded-xl space-y-4 border border-[#edeef0]">
                <h4 className="text-[10px] font-bold text-gray-400 uppercase tracking-widest flex items-center gap-2 mb-2">
                  <CreditCard size={12} /> Payroll & Banking
                </h4>
                <div className="space-y-1">
                  <label className="text-[10px] font-bold text-gray-500 uppercase">Account Name</label>
                  <input
                    type="text"
                    value={formData.accountName ?? ""}
                    onChange={(e) => setFormData({ ...formData, accountName: e.target.value })}
                    className="w-full px-3 py-2 border border-[#edeef0] rounded-md text-sm outline-none focus:bg-white"
                  />
                </div>
                <div className="grid grid-cols-3 gap-4">
                  <div className="space-y-1 col-span-1">
                    <label className="text-[10px] font-bold text-gray-500 uppercase">BSB</label>
                    <input
                      type="text"
                      placeholder="000-000"
                      value={formData.bsb ?? ""}
                      onChange={(e) => setFormData({ ...formData, bsb: e.target.value })}
                      className="w-full px-3 py-2 border border-[#edeef0] rounded-md text-sm outline-none focus:bg-white"
                    />
                  </div>
                  <div className="space-y-1 col-span-2">
                    <label className="text-[10px] font-bold text-gray-500 uppercase">Account Number</label>
                    <input
                      type="text"
                      value={formData.accountNumber ?? ""}
                      onChange={(e) => setFormData({ ...formData, accountNumber: e.target.value })}
                      className="w-full px-3 py-2 border border-[#edeef0] rounded-md text-sm outline-none focus:bg-white"
                    />
                  </div>
                </div>
              </div>

              <div className="flex items-center gap-2">
                <input
                  type="checkbox"
                  id="active"
                  checked={formData.active !== false}
                  onChange={(e) => setFormData({ ...formData, active: e.target.checked })}
                  className="rounded border-gray-300"
                />
                <label htmlFor="active" className="text-sm text-gray-700">
                  Active
                </label>
              </div>

              {submitError && (
                <div className="bg-amber-50 border border-amber-200 text-amber-800 px-3 py-2 rounded-lg text-sm">
                  {submitError}
                </div>
              )}

              <button
                type="submit"
                disabled={submitLoading}
                className="w-full bg-gray-900 text-white py-3 rounded-xl font-bold hover:bg-gray-800 transition-all shadow-lg active:scale-[0.98] disabled:opacity-50 flex items-center justify-center gap-2"
              >
                {submitLoading && <Loader2 className="animate-spin" size={18} />}
                {editingCleaner ? "Save changes" : "Add Cleaner"}
              </button>
            </form>
          </div>
        </div>
      )}

      {bulkModalOpen && (
        <div className="fixed inset-0 bg-gray-900/60 backdrop-blur-sm z-[100] flex items-center justify-center p-4">
          <div className="bg-white w-full max-w-4xl rounded-2xl shadow-2xl overflow-hidden animate-slideUp">
            <div className="p-6 border-b border-[#edeef0] flex justify-between items-center bg-[#fcfcfb]">
              <div className="flex items-center gap-3">
                <div className="w-8 h-8 bg-gray-900 rounded-lg flex items-center justify-center text-white">
                  <Layers size={20} />
                </div>
                <h3 className="font-bold text-gray-900">Bulk Add Cleaners</h3>
              </div>
              <button
                onClick={() => !bulkSubmitLoading && setBulkModalOpen(false)}
                className="text-gray-400 hover:text-gray-900 transition-colors"
                disabled={bulkSubmitLoading}
              >
                <X size={20} />
              </button>
            </div>

            <div className="p-6 space-y-4">
              <p className="text-xs text-gray-500">
                Enter multiple cleaners at once. Only rows with a name will be created.
              </p>
              <div className="border border-[#edeef0] rounded-lg overflow-x-auto">
                <table className="w-full text-left text-sm">
                  <thead>
                    <tr className="bg-gray-50 border-b border-[#edeef0] text-[10px] font-bold text-gray-500 uppercase tracking-widest">
                      <th className="py-2 px-2 w-6"></th>
                      <th className="py-2 px-2 min-w-[140px]">Cleaner Name</th>
                      <th className="py-2 px-2 w-24">Rate ($/hr)</th>
                      <th className="py-2 px-2 min-w-[160px]">Account name</th>
                      <th className="py-2 px-2 w-24">BSB</th>
                      <th className="py-2 px-2 w-32">Account no.</th>
                      <th className="py-2 px-2 w-20 text-center">Active</th>
                    </tr>
                  </thead>
                  <tbody>
                    {bulkRows.map((row) => (
                      <tr key={row.id} className="border-b border-[#edeef0] last:border-b-0">
                        <td className="py-1 px-2">
                          <button
                            type="button"
                            onClick={() => removeBulkRow(row.id)}
                            disabled={bulkRows.length <= 1 || bulkSubmitLoading}
                            className="text-gray-400 hover:text-red-600 disabled:opacity-40 text-xs"
                          >
                            ✕
                          </button>
                        </td>
                        <td className="py-1 px-2">
                          <input
                            type="text"
                            value={row.cleanerName}
                            onChange={(e) => updateBulkRow(row.id, { cleanerName: e.target.value })}
                            className="w-full border border-[#edeef0] rounded px-2 py-1 text-sm"
                            placeholder="Name"
                            disabled={bulkSubmitLoading}
                          />
                        </td>
                        <td className="py-1 px-2">
                          <input
                            type="number"
                            min={0}
                            step={0.5}
                            value={row.payRatePerHour === "" ? "" : row.payRatePerHour}
                            onChange={(e) =>
                              updateBulkRow(row.id, {
                                payRatePerHour: e.target.value === "" ? "" : parseFloat(e.target.value) || 0,
                              })
                            }
                            className="w-full border border-[#edeef0] rounded px-2 py-1 text-sm"
                            disabled={bulkSubmitLoading}
                          />
                        </td>
                        <td className="py-1 px-2">
                          <input
                            type="text"
                            value={row.accountName}
                            onChange={(e) => updateBulkRow(row.id, { accountName: e.target.value })}
                            className="w-full border border-[#edeef0] rounded px-2 py-1 text-sm"
                            disabled={bulkSubmitLoading}
                          />
                        </td>
                        <td className="py-1 px-2">
                          <input
                            type="text"
                            value={row.bsb}
                            onChange={(e) => updateBulkRow(row.id, { bsb: e.target.value })}
                            className="w-full border border-[#edeef0] rounded px-2 py-1 text-sm"
                            disabled={bulkSubmitLoading}
                          />
                        </td>
                        <td className="py-1 px-2">
                          <input
                            type="text"
                            value={row.accountNumber}
                            onChange={(e) => updateBulkRow(row.id, { accountNumber: e.target.value })}
                            className="w-full border border-[#edeef0] rounded px-2 py-1 text-sm"
                            disabled={bulkSubmitLoading}
                          />
                        </td>
                        <td className="py-1 px-2 text-center">
                          <input
                            type="checkbox"
                            checked={row.active}
                            onChange={(e) => updateBulkRow(row.id, { active: e.target.checked })}
                            className="rounded border-gray-300"
                            disabled={bulkSubmitLoading}
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
                className="text-xs font-medium text-gray-700 hover:text-gray-900"
              >
                + Add new line
              </button>

              <div className="mt-4 flex justify-end gap-2">
                <button
                  type="button"
                  onClick={() => !bulkSubmitLoading && setBulkModalOpen(false)}
                  className="px-4 py-2 text-sm font-medium text-gray-600 hover:bg-gray-100 rounded-lg"
                  disabled={bulkSubmitLoading}
                >
                  Cancel
                </button>
                <button
                  type="button"
                  onClick={handleBulkSubmit}
                  disabled={bulkSubmitLoading}
                  className="bg-gray-900 text-white px-6 py-2.5 rounded-lg text-sm font-bold hover:bg-gray-800 disabled:opacity-50 flex items-center gap-2"
                >
                  {bulkSubmitLoading && <Loader2 className="animate-spin" size={14} />}
                  Add {bulkRows.filter((r) => r.cleanerName.trim() !== "").length} cleaners
                </button>
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

export default CleanerManager;
