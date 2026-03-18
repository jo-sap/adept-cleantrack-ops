import React, { useState, useEffect, useCallback, useMemo } from "react";
import { Cleaner } from "../types";
import { Plus, CreditCard, X, UserCircle, Search, Loader2, Layers, Trash2, Pencil, UserMinus, UserPlus, ChevronUp, ChevronDown } from "lucide-react";
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
import {
  getSiteCleanerAssignments,
  type SiteCleanerAssignment,
} from "../repositories/assignedCleanersRepo";
import { getSites, type Site } from "../repositories/sitesRepo";

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
  // Managers should have the same CRUD capabilities as Admin for cleaners onboarding.
  const canManageCleaners = isAdminFromRole || user?.role === "Admin" || user?.role === "Manager";

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
  const [selectedCleanerIds, setSelectedCleanerIds] = useState<string[]>([]);
  const [bulkDeleteLoading, setBulkDeleteLoading] = useState(false);

  const [assignmentsByCleanerId, setAssignmentsByCleanerId] = useState<Record<string, SiteCleanerAssignment[]>>({});
  const [hoverCleanerId, setHoverCleanerId] = useState<string | null>(null);
  const [hoverPopover, setHoverPopover] = useState<{
    cleanerId: string;
    left: number;
    top: number;
  } | null>(null);
  const [isHoverPopoverOver, setIsHoverPopoverOver] = useState(false);
  const [modalCleaner, setModalCleaner] = useState<CleanerItem | null>(null);
  const [modalAssignments, setModalAssignments] = useState<SiteCleanerAssignment[]>([]);
  const [sitesById, setSitesById] = useState<Record<string, Site>>({});

  type CleanerSortKey = "name" | "status" | "rate";
  const [cleanerSortBy, setCleanerSortBy] = useState<CleanerSortKey>("name");

  /** Resolve site display name from assignment (lookup) or fallback join to CleanTrack Sites. */
  const resolveSiteDisplayName = useCallback(
    (a: SiteCleanerAssignment, sites: Record<string, Site>): string => {
      if (a.siteName && String(a.siteName).trim()) return String(a.siteName).trim();
      const site = (a.siteId && sites[String(a.siteId)]) || (a.siteId && sites[a.siteId as unknown as string]);
      if (site?.siteName) return site.siteName;
      if (site?.address) return site.address;
      return "";
    },
    []
  );
  const [cleanerSortDir, setCleanerSortDir] = useState<"asc" | "desc">("asc");

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
      const [data, assignments, sites] = await Promise.all([
        getCleaners(token),
        getSiteCleanerAssignments(token, { activeOnly: true }),
        getSites(token).catch(() => [] as Site[]),
      ]);

      // Debug: raw assignments and grouped counts
      console.log("[CleanerManager] site-cleaner assignments (active only) sample:", {
        total: assignments.length,
        first: assignments[0],
      });

      const grouped: Record<string, SiteCleanerAssignment[]> = {};
      for (const a of assignments) {
        if (!a.cleanerId || !a.active) continue;
        if (!grouped[a.cleanerId]) grouped[a.cleanerId] = [];
        grouped[a.cleanerId].push(a);
      }
      console.log("[CleanerManager] assignments grouped by cleaner:", {
        cleanerCount: Object.keys(grouped).length,
      });

      if (sites.length > 0) {
        console.log("[CleanerManager] sample CleanTrack Sites for fallback join:", {
          count: sites.length,
          first: sites[0],
        });
      }

      setCleaners(data);
      setAssignmentsByCleanerId(grouped);
      setSitesById(
        sites.reduce<Record<string, Site>>((acc, s) => {
          acc[String(s.id)] = s;
          return acc;
        }, {})
      );
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

  const sortedCleaners = useMemo(() => {
    const list = [...filteredCleaners];
    list.sort((a, b) => {
      let cmp = 0;
      if (cleanerSortBy === "name") {
        const an = a.cleanerName.toLowerCase();
        const bn = b.cleanerName.toLowerCase();
        cmp = an.localeCompare(bn);
      } else if (cleanerSortBy === "status") {
        const av = a.active ? 1 : 0;
        const bv = b.active ? 1 : 0;
        cmp = av - bv;
      } else if (cleanerSortBy === "rate") {
        const ar = a.payRatePerHour ?? 0;
        const br = b.payRatePerHour ?? 0;
        cmp = ar - br;
      }
      return cleanerSortDir === "asc" ? cmp : -cmp;
    });
    return list;
  }, [filteredCleaners, cleanerSortBy, cleanerSortDir]);

  const handleCleanerSort = (key: CleanerSortKey) => {
    if (cleanerSortBy === key) {
      setCleanerSortDir((d) => (d === "asc" ? "desc" : "asc"));
    } else {
      setCleanerSortBy(key);
      setCleanerSortDir(key === "name" ? "asc" : "desc");
    }
  };

  const selectedSet = new Set(selectedCleanerIds);
  const toggleSelected = (id: string) => {
    setSelectedCleanerIds((prev) =>
      prev.includes(id) ? prev.filter((x) => x !== id) : [...prev, id]
    );
  };
  const selectAllFiltered = () => {
    const ids = filteredCleaners.map((c) => c.id);
    if (ids.length === 0) return;
    if (selectedSet.size === ids.length) {
      setSelectedCleanerIds([]);
    } else {
      setSelectedCleanerIds(ids);
    }
  };
  const handleBulkDelete = async () => {
    if (selectedCleanerIds.length === 0) return;
    if (
      !window.confirm(
        `Delete ${selectedCleanerIds.length} selected cleaner(s)? Existing timesheet entries will still reference them.`
      )
    ) {
      return;
    }
    const token = await getGraphAccessToken();
    if (!token) {
      setError("Not signed in. Sign in with Microsoft to delete cleaners.");
      return;
    }
    setBulkDeleteLoading(true);
    setError(null);
    let deleted = 0;
    let failed = 0;
    try {
      for (const id of selectedCleanerIds) {
        try {
          await deleteCleaner(token, id);
          deleted++;
        } catch {
          failed++;
        }
      }
      setSelectedCleanerIds([]);
      await loadCleaners();
      onCleanersRefresh?.();
      if (failed > 0) {
        showToast(`Deleted ${deleted} cleaner(s). ${failed} failed.`);
      } else {
        showToast(deleted === 1 ? "Cleaner deleted." : `Deleted ${deleted} cleaners.`);
      }
    } catch (e) {
      const msg = e instanceof Error ? e.message : "Bulk delete failed.";
      setError(msg);
    } finally {
      setBulkDeleteLoading(false);
    }
  };

  return (
    <div className="space-y-6 sm:space-y-8 animate-fadeIn">
      <div className="flex flex-col sm:flex-row sm:justify-between sm:items-end gap-4 border-b border-[#edeef0] pb-4">
        <div className="min-w-0">
          <h2 className="text-2xl sm:text-3xl font-bold text-gray-900">Cleaner Team</h2>
          <p className="text-gray-500 text-sm mt-1">Manage personnel, onboarding details, and banking records.</p>
        </div>
        {canManageCleaners && (
          <div className="flex items-center gap-2 flex-shrink-0">
            <button
              onClick={handleOpenBulk}
              className="bg-white text-gray-900 border border-[#edeef0] px-4 py-2 rounded-lg text-sm font-bold hover:bg-gray-50 transition-colors flex items-center gap-2"
            >
              <Layers size={16} /> Bulk Add
            </button>
            <button
              onClick={handleOpenAdd}
              className="bg-gray-900 text-white px-4 py-2 rounded-lg text-sm font-bold hover:bg-gray-800 transition-colors flex items-center gap-2"
            >
              <Plus size={16} /> New Cleaner
            </button>
          </div>
        )}
      </div>

      <div className="relative">
        <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" size={18} />
        <input
          type="search"
          placeholder="Search team…"
          value={searchQuery}
          onChange={(e) => setSearchQuery(e.target.value)}
          className="w-full pl-10 pr-4 py-2.5 border border-[#edeef0] rounded-lg text-sm text-gray-900 placeholder-gray-400 focus:outline-none focus:ring-2 focus:ring-gray-900/10 focus:border-gray-400"
          aria-label="Search cleaners"
        />
      </div>

      {canManageCleaners && selectedCleanerIds.length > 0 && (
        <div className="sticky top-12 z-20 flex flex-wrap items-center gap-2 py-2 px-3 bg-amber-50 border border-amber-200 rounded-lg shadow-sm">
          <span className="text-sm font-medium text-amber-800">
            {selectedCleanerIds.length} selected
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
            onClick={() => setSelectedCleanerIds([])}
            className="text-xs font-medium text-amber-800 hover:text-amber-900"
          >
            Clear selection
          </button>
        </div>
      )}

      {modalCleaner && (
        <div className="fixed inset-0 bg-gray-900/60 backdrop-blur-sm z-[120] flex items-center justify-center p-4">
          <div className="bg-white w-full max-w-lg rounded-2xl shadow-2xl overflow-hidden">
            <div className="p-5 border-b border-[#edeef0] flex justify-between items-center bg-[#fcfcfb]">
              <div className="min-w-0">
                <h3 className="font-semibold text-gray-900 truncate">
                  {modalCleaner.cleanerName} — Assigned Sites
                </h3>
                <p className="text-[12px] text-gray-500 mt-0.5">
                  {modalAssignments.length} active site
                  {modalAssignments.length === 1 ? "" : "s"}
                </p>
              </div>
              <button
                type="button"
                onClick={() => {
                  setModalCleaner(null);
                  setModalAssignments([]);
                }}
                className="text-gray-400 hover:text-gray-900 transition-colors"
                aria-label="Close assigned sites"
              >
                <X size={20} />
              </button>
            </div>
            <div className="p-5 max-h-[60vh] overflow-y-auto">
              {modalAssignments.length === 0 ? (
                <p className="text-sm text-gray-500">
                  No active site assignments for this cleaner.
                </p>
              ) : (
                <ul className="space-y-2 text-sm">
                  {modalAssignments.map((a) => (
                    <li
                      key={a.id}
                      className="flex items-start gap-2 border-b border-[#edeef0] last:border-b-0 pb-2"
                    >
                      <span className="mt-0.5 text-[10px] text-gray-400">•</span>
                      <div>
                        <p className="font-medium text-gray-900">
                          {resolveSiteDisplayName(a, sitesById) || "Unnamed site"}
                        </p>
                      </div>
                    </li>
                  ))}
                </ul>
              )}
            </div>
          </div>
        </div>
      )}

      {toast && (
        <div className="bg-green-50 border border-green-200 text-green-800 px-4 py-2 rounded-lg text-sm">
          {toast}
        </div>
      )}

      {hoverCleanerId && hoverPopover && (
        <div
          className="fixed z-[200] w-64 so-card bg-white p-3 text-xs shadow-lg"
          style={{ left: hoverPopover.left, top: hoverPopover.top }}
          onMouseEnter={() => setIsHoverPopoverOver(true)}
          onMouseLeave={() => {
            setIsHoverPopoverOver(false);
            setHoverCleanerId(null);
            setHoverPopover(null);
          }}
        >
          {(() => {
            const assignments = assignmentsByCleanerId[hoverCleanerId] ?? [];
            const names = assignments
              .map((a) => resolveSiteDisplayName(a, sitesById))
              .filter((n) => !!n);
            const preview = names.slice(0, 6);
            const remaining = Math.max(names.length - preview.length, 0);
            return (
              <>
                <p className="text-[11px] font-semibold text-gray-800 mb-1">
                  Assigned Sites
                </p>
                {assignments.length === 0 ? (
                  <p className="text-[11px] text-gray-500">
                    No active site assignments
                  </p>
                ) : preview.length === 0 ? (
                  <p className="text-[11px] text-gray-500">
                    Assigned sites found, but site names are missing from lookup fields.
                  </p>
                ) : (
                  <ul className="space-y-0.5 text-gray-700">
                    {preview.map((name) => (
                      <li key={name} className="truncate">
                        • {name}
                      </li>
                    ))}
                    {remaining > 0 && (
                      <li className="text-gray-500">+{remaining} more</li>
                    )}
                  </ul>
                )}
              </>
            );
          })()}
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
        <div className="border border-[#edeef0] rounded-lg bg-white shadow-sm overflow-hidden table-scroll-mobile">
          <table className="w-full border-collapse text-left table-fixed min-w-[760px]">
            <colgroup>
              {canManageCleaners && <col style={{ width: '4%' }} />}
              <col style={{ width: canManageCleaners ? '5%' : '6%' }} />
              <col style={{ width: canManageCleaners ? '22%' : '26%' }} />
              <col style={{ width: '8%' }} />
              <col style={{ width: '8%' }} />
              <col style={{ width: canManageCleaners ? '18%' : '22%' }} />
              <col style={{ width: '10%' }} />
              <col style={{ width: '8%' }} />
              <col style={{ width: '11%' }} />
              {canManageCleaners && <col style={{ width: '12%' }} />}
            </colgroup>
            <thead>
              <tr className="bg-[#fcfcfb] border-b border-[#edeef0]">
                {canManageCleaners && (
                  <th className="px-1.5 py-1.5 w-10">
                    <input
                      type="checkbox"
                      checked={filteredCleaners.length > 0 && selectedSet.size === filteredCleaners.length}
                      onChange={selectAllFiltered}
                      className="rounded border-gray-300"
                      aria-label="Select all"
                    />
                  </th>
                )}
                <th className="px-1.5 py-1.5 w-12"></th>
                <th className="px-1.5 py-1.5">
                  <button
                    type="button"
                    onClick={() => handleCleanerSort("name")}
                    className="text-[9px] font-bold text-gray-500 uppercase tracking-widest flex items-center gap-0.5 hover:text-gray-900 focus:outline-none focus:ring-2 focus:ring-gray-900/20 rounded"
                  >
                    Name
                    {cleanerSortBy === "name" &&
                      (cleanerSortDir === "asc" ? (
                        <ChevronUp size={10} />
                      ) : (
                        <ChevronDown size={10} />
                      ))}
                  </button>
                </th>
                <th className="px-1.5 py-1.5">
                  <button
                    type="button"
                    onClick={() => handleCleanerSort("status")}
                    className="text-[9px] font-bold text-gray-500 uppercase tracking-widest inline-flex items-center gap-0.5 hover:text-gray-900 focus:outline-none focus:ring-2 focus:ring-gray-900/20 rounded"
                  >
                    Status
                    {cleanerSortBy === "status" &&
                      (cleanerSortDir === "asc" ? (
                        <ChevronUp size={10} />
                      ) : (
                        <ChevronDown size={10} />
                      ))}
                  </button>
                </th>
                <th className="px-1.5 py-1.5">
                  <button
                    type="button"
                    onClick={() => handleCleanerSort("rate")}
                    className="text-[9px] font-bold text-gray-500 uppercase tracking-widest inline-flex items-center gap-0.5 hover:text-gray-900 focus:outline-none focus:ring-2 focus:ring-gray-900/20 rounded"
                  >
                    Rate
                    {cleanerSortBy === "rate" &&
                      (cleanerSortDir === "asc" ? (
                        <ChevronUp size={10} />
                      ) : (
                        <ChevronDown size={10} />
                      ))}
                  </button>
                </th>
                <th className="px-1.5 py-1.5 text-[9px] font-bold text-gray-500 uppercase tracking-widest">Account name</th>
                <th className="px-1.5 py-1.5 text-[9px] font-bold text-gray-500 uppercase tracking-widest">BSB</th>
                <th className="px-1.5 py-1.5 text-[9px] font-bold text-gray-500 uppercase tracking-widest">Account no.</th>
                <th className="px-1.5 py-1.5 text-[9px] font-bold text-gray-500 uppercase tracking-widest">Sites Assigned</th>
                {canManageCleaners && <th className="px-1.5 py-1.5 text-[9px] font-bold text-gray-500 uppercase tracking-widest text-right">Actions</th>}
              </tr>
            </thead>
            <tbody>
              {sortedCleaners.map((cleaner) => {
                const appCleaner = toAppCleaner(cleaner);
                const assignments = assignmentsByCleanerId[cleaner.id] ?? [];
                const assignedSiteCount = assignments.length;
                const assignedSiteNames = assignments
                  .map((a) => resolveSiteDisplayName(a, sitesById))
                  .filter((n) => !!n);
                const previewNames = assignedSiteNames.slice(0, 4);
                const remaining = Math.max(assignedSiteNames.length - previewNames.length, 0);
                return (
                  <tr
                    key={cleaner.id}
                    className="border-b border-[#edeef0] last:border-b-0 hover:bg-[#f7f6f3] transition-colors"
                  >
                    {canManageCleaners && (
                      <td className="px-1.5 py-1.5">
                        <input
                          type="checkbox"
                          checked={selectedSet.has(cleaner.id)}
                          onChange={() => toggleSelected(cleaner.id)}
                          className="rounded border-gray-300"
                          aria-label={`Select ${cleaner.cleanerName}`}
                        />
                      </td>
                    )}
                    <td className="px-1.5 py-1.5">
                      <div className="w-8 h-8 bg-gray-100 rounded-lg flex items-center justify-center text-gray-600 font-bold text-[11px]">
                        {appCleaner.firstName.charAt(0)}
                        {appCleaner.lastName ? appCleaner.lastName.charAt(0) : ""}
                      </div>
                    </td>
                    <td className="px-1.5 py-1.5">
                      <span className="text-xs font-semibold text-gray-900 break-words">{cleaner.cleanerName}</span>
                    </td>
                    <td className="px-1.5 py-1.5">
                      <span className={`text-[10px] font-bold uppercase ${cleaner.active ? "text-green-600" : "text-gray-400"}`}>
                        {cleaner.active ? "Active" : "Inactive"}
                      </span>
                    </td>
                    <td className="px-1.5 py-1.5">
                      <span className="text-[11px] font-semibold text-green-600">${Number(cleaner.payRatePerHour || 0).toFixed(2)}/hr</span>
                    </td>
                    <td className="px-1.5 py-1.5 text-[11px] text-gray-600 break-words" title={cleaner.accountName || undefined}>
                      {cleaner.accountName || "—"}
                    </td>
                    <td className="px-1.5 py-1.5 text-[11px] text-gray-600 font-mono">
                      {cleaner.bsb || "—"}
                    </td>
                    <td className="px-1.5 py-1.5 text-[11px] text-gray-600 font-mono">
                      {cleaner.accountNumber || "—"}
                    </td>
                    <td className="px-1.5 py-1.5">
                      <div className="relative inline-block">
                        <button
                          type="button"
                          onMouseEnter={(e) => {
                            setHoverCleanerId(cleaner.id);
                            const rect = (e.currentTarget as HTMLButtonElement).getBoundingClientRect();
                            setHoverPopover({
                              cleanerId: cleaner.id,
                              left: Math.round(rect.left),
                              top: Math.round(rect.bottom + 8),
                            });
                          }}
                          onMouseLeave={() => {
                            // If user is moving into the popover, don't immediately close
                            setTimeout(() => {
                              if (!isHoverPopoverOver) {
                                setHoverCleanerId(null);
                                setHoverPopover(null);
                              }
                            }, 40);
                          }}
                          onClick={() => {
                            setModalCleaner(cleaner);
                            setModalAssignments(assignments);
                            console.log("[CleanerManager] open sites modal payload:", {
                              cleaner,
                              assignments,
                            });
                          }}
                          className="text-[11px] font-medium px-2 py-1 rounded-full bg-[#ECF3F4] text-[#3E5F6A] border border-transparent hover:border-[#3E5F6A]/50 hover:bg-[#dde8ea] transition-colors"
                        >
                          {assignedSiteCount === 1
                            ? "1 site"
                            : `${assignedSiteCount} sites`}
                        </button>
                      </div>
                    </td>
                    {canManageCleaners && (
                      <td className="px-1.5 py-1.5 text-right">
                        <div className="flex justify-end gap-1 flex-wrap">
                          <button
                            type="button"
                            onClick={() => handleOpenEdit(cleaner)}
                            className="touch-target p-2.5 sm:p-1.5 rounded text-blue-600 hover:text-blue-800 hover:bg-blue-50 inline-flex items-center justify-center"
                            aria-label={`Edit ${cleaner.cleanerName}`}
                            title="Edit"
                          >
                            <Pencil size={18} className="sm:w-3.5 sm:h-3.5 w-[18px] h-[18px]" />
                          </button>
                          <button
                            type="button"
                            onClick={() => handleToggleActive(cleaner)}
                            className="touch-target p-2.5 sm:p-1.5 rounded text-gray-600 hover:text-gray-900 hover:bg-gray-100 inline-flex items-center justify-center"
                            aria-label={cleaner.active ? `Deactivate ${cleaner.cleanerName}` : `Activate ${cleaner.cleanerName}`}
                            title={cleaner.active ? "Deactivate" : "Activate"}
                          >
                            {cleaner.active ? <UserMinus size={18} className="sm:w-3.5 sm:h-3.5 w-[18px] h-[18px]" /> : <UserPlus size={18} className="sm:w-3.5 sm:h-3.5 w-[18px] h-[18px]" />}
                          </button>
                          <button
                            type="button"
                            onClick={() => handleDeleteCleaner(cleaner)}
                            className="touch-target p-2.5 sm:p-1.5 rounded text-red-600 hover:text-red-800 hover:bg-red-50 inline-flex items-center justify-center"
                            aria-label={`Delete ${cleaner.cleanerName}`}
                            title="Delete"
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
