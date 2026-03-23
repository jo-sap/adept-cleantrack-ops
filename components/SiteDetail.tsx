
import React, { useEffect, useMemo, useState } from 'react';
import { XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, BarChart, Bar, Cell } from 'recharts';
import { Site, TimeEntry, FortnightPeriod, Cleaner, SiteCleanerAssignment } from '../types';
import { format, addDays, getDay } from 'date-fns';
import { ArrowLeft, CalendarDays, Wallet, Percent, UserPlus, X } from 'lucide-react';
import { useRole } from '../contexts/RoleContext';
import { useAppAuth } from '../contexts/AppAuthContext';
import { getGraphAccessToken } from '../lib/graph';
import { createSiteCleanerAssignment, getSiteCleanerAssignments, updateSiteCleanerAssignment } from '../repositories/assignedCleanersRepo';

interface SiteDetailProps {
  site: Site;
  cleaners: Cleaner[];
  entries: TimeEntry[];
  currentPeriod: FortnightPeriod;
  onBack: () => void;
  /** Label for the back button, e.g. "Dashboard" or "Sites & Budgets". */
  backLabel?: string;
  onRefreshSites?: () => void;
}

const SiteDetail: React.FC<SiteDetailProps> = ({ site, cleaners, entries, currentPeriod, onBack, backLabel = 'Dashboard', onRefreshSites }) => {
  const { isAdmin: isAdminFromRole } = useRole();
  const { user } = useAppAuth();
  /** Strict: no GP/revenue visible to Managers — Admin only */
  const isAdmin = (isAdminFromRole || user?.role === 'Admin') && user?.role !== 'Manager';
  const canManageAssignments = isAdminFromRole || !!user;

  const [assignments, setAssignments] = useState<SiteCleanerAssignment[]>([]);
  const [assignmentsLoading, setAssignmentsLoading] = useState(false);
  const [assignmentsError, setAssignmentsError] = useState<string | null>(null);
  const [showAssignForm, setShowAssignForm] = useState(false);
  const [assignCleanerId, setAssignCleanerId] = useState<string>('');
  const [assignCleanerSearch, setAssignCleanerSearch] = useState<string>('');
  const [savingAssignment, setSavingAssignment] = useState(false);

  useEffect(() => {
    let cancelled = false;
    const loadAssignments = async () => {
      setAssignmentsLoading(true);
      setAssignmentsError(null);
      try {
        const token = await getGraphAccessToken();
        if (!token || cancelled) {
          if (!token) {
            setAssignmentsError('Cleaner assignments unavailable (no Graph token).');
          }
          setAssignmentsLoading(false);
          return;
        }
        const list = await getSiteCleanerAssignments(token, {
          activeOnly: true,
          siteId: site.id,
        });
        if (!cancelled) {
          console.log('[SiteDetail] active assignments for site', site.id, list);
          setAssignments(list);
        }
      } catch (err) {
        if (!cancelled) {
          setAssignmentsError('Failed to load cleaner assignments for this site.');
        }
      } finally {
        if (!cancelled) setAssignmentsLoading(false);
      }
    };
    loadAssignments();
    return () => {
      cancelled = true;
    };
  }, [site.id]);

  const periodEntries = useMemo(() => entries.filter(e => {
    const date = new Date(e.date);
    return date >= currentPeriod.startDate && date <= currentPeriod.endDate && e.siteId === site.id;
  }), [entries, currentPeriod, site.id]);

  const financialStats = useMemo(() => {
    let actualLaborCost = 0;
    periodEntries.forEach(e => {
      const cleaner = cleaners.find(c => c.id === e.cleanerId);
      const rate = e.pay_rate_snapshot ?? cleaner?.payRatePerHour ?? 0;
      actualLaborCost += e.hours * rate;
    });

    const fortnightRevenue = ((site as any).monthly_revenue ?? (site as any).monthlyRevenue ?? 0) * 12 / 26;
    const grossProfit = fortnightRevenue - actualLaborCost;
    const margin = fortnightRevenue > 0 ? (grossProfit / fortnightRevenue) * 100 : 0;

    return { actualLaborCost, fortnightRevenue, grossProfit, margin };
  }, [periodEntries, cleaners, site]);

  const budgetedHoursPerFortnight = (site as any).budgeted_hours_per_fortnight ?? (site as any).budgetedHoursPerFortnight ?? 0;
  const dailyBudgetsArr = (site as any).daily_budgets ?? (site as any).dailyBudgets ?? [];
  const dailyBudgetsWeek2Arr = (site as any).daily_budgets_week2 ?? (site as any).dailyBudgetsWeek2 ?? [];

  const getPlannedForIndex = useMemo(() => {
    const visit = String((site as any).visit_frequency ?? (site as any).visitFrequency ?? "").trim().toLowerCase();
    const hasWeek2 = Array.isArray(dailyBudgetsWeek2Arr) && dailyBudgetsWeek2Arr.length >= 7;
    return (dateIndex: number, dayOfWeek: number) => {
      if (visit === "monthly") return 0;
      const arr =
        visit === "fortnightly" && hasWeek2 && dateIndex >= 7 ? dailyBudgetsWeek2Arr : dailyBudgetsArr;
      return (arr?.[dayOfWeek] ?? 0) as number;
    };
  }, [site, dailyBudgetsArr, dailyBudgetsWeek2Arr]);

  const availableCleanersForAssignment = useMemo(() => {
    const activeAssignedIds = new Set(
      assignments.filter(a => a.active && a.cleanerId).map(a => a.cleanerId)
    );
    const options = cleaners.filter(c => !activeAssignedIds.has(c.id));
    console.log('[SiteDetail] assign-cleaner dropdown options', {
      siteId: site.id,
      availableCleanerIds: options.map(o => o.id),
    });
    return options;
  }, [assignments, cleaners, site.id]);

  const filteredCleanersForAssignment = useMemo(() => {
    const q = assignCleanerSearch.trim().toLowerCase();
    if (!q) return availableCleanersForAssignment;
    return availableCleanersForAssignment.filter((c) => {
      const first = (c.firstName || "").toLowerCase();
      const last = (c.lastName || "").toLowerCase();
      const email = (c.email || "").toLowerCase();
      return first.includes(q) || last.includes(q) || `${first} ${last}`.trim().includes(q) || email.includes(q);
    });
  }, [availableCleanersForAssignment, assignCleanerSearch]);

  const handleCreateAssignment = async () => {
    if (!assignCleanerId) return;
    setSavingAssignment(true);
    try {
      const token = await getGraphAccessToken();
      if (!token) {
        setAssignmentsError('Cannot assign cleaner – missing Microsoft Graph token.');
        return;
      }
      const cleaner = cleaners.find(c => c.id === assignCleanerId);
      if (!cleaner) return;
      const assignmentName = `${site.name} - ${cleaner.firstName} ${cleaner.lastName}`.trim();
      const created = await createSiteCleanerAssignment(token, {
        siteId: site.id,
        cleanerId: assignCleanerId,
        assignmentName,
        active: true,
      });
      setAssignments(prev => [...prev, created]);
      setShowAssignForm(false);
      setAssignCleanerId('');
      setAssignCleanerSearch('');
      if (onRefreshSites) {
        onRefreshSites();
      }
    } catch (err) {
      setAssignmentsError('Failed to create cleaner assignment.');
    } finally {
      setSavingAssignment(false);
    }
  };

  const handleDeactivateAssignment = async (assignment: SiteCleanerAssignment) => {
    try {
      const token = await getGraphAccessToken();
      if (!token) {
        setAssignmentsError('Cannot update assignment – missing Microsoft Graph token.');
        return;
      }
      await updateSiteCleanerAssignment(token, assignment.id, { active: false });
      setAssignments(prev =>
        prev.filter(a => a.id !== assignment.id)
      );
      if (onRefreshSites) {
        onRefreshSites();
      }
    } catch (err) {
      setAssignmentsError('Failed to update cleaner assignment.');
    }
  };

  const dailyStats = useMemo(() => {
    const stats = [];
    for (let i = 0; i < 14; i++) {
      const day = addDays(currentPeriod.startDate, i);
      const dayStr = format(day, 'yyyy-MM-dd');
      const dayOfWeek = getDay(day);
      const actual = periodEntries
        .filter(e => e.date === dayStr)
        .reduce((sum, e) => sum + e.hours, 0);

      const dayBudget = getPlannedForIndex(i, dayOfWeek);
      const isScheduled = dayBudget > 0;

      stats.push({
        date: dayStr,
        displayDate: format(day, 'EEE d'),
        actual,
        budget: dayBudget,
        isScheduled,
        variance: isScheduled ? (actual - dayBudget) : actual
      });
    }
    return stats;
  }, [site, periodEntries, currentPeriod, dailyBudgetsArr, getPlannedForIndex]);

  const totalActualHours = periodEntries.reduce((sum, e) => sum + e.hours, 0);
  const hourVariance = totalActualHours - budgetedHoursPerFortnight;

  return (
    <div className="space-y-6 sm:space-y-8 animate-fadeIn">
      <div className="flex flex-col sm:flex-row sm:justify-between sm:items-end gap-4 border-b border-[#edeef0] pb-4">
        <div className="min-w-0">
          <h2 className="text-2xl sm:text-3xl font-bold text-gray-900">Site Detail</h2>
          <p className="text-gray-500 text-sm mt-1">
            Performance and variance for {site.name || "this site"}.
          </p>
        </div>
        <div className="flex-shrink-0">
          <button
            type="button"
            onClick={onBack}
            className="flex items-center gap-2 bg-white text-gray-900 border border-[#edeef0] px-4 py-2.5 rounded-lg text-sm font-bold hover:bg-gray-50 transition-colors"
          >
            <ArrowLeft size={16} />
            Back to {backLabel}
          </button>
        </div>
      </div>

      <div className="flex flex-col md:flex-row justify-between items-start md:items-end gap-4 border-b border-[#edeef0] pb-4">
        <div>
          <h3 className="text-xl sm:text-2xl font-bold text-gray-900">{site.name}</h3>
          <p className="text-gray-500 text-sm mt-0.5">{site.address}</p>
        </div>
        <div className="flex gap-6">
          <div className="text-right">
            <p className="text-[9px] font-bold text-gray-500 uppercase tracking-widest">Fortnight hours</p>
            <p className={`text-xl font-bold ${hourVariance > 0 ? 'text-red-600' : 'text-green-600'}`}>{totalActualHours.toFixed(1)}h</p>
          </div>
          <div className="text-right">
            <p className="text-[9px] font-bold text-gray-500 uppercase tracking-widest">Hourly target</p>
            <p className="text-xl font-bold text-gray-900">{budgetedHoursPerFortnight}h</p>
          </div>
        </div>
      </div>

      <div className="border border-[#edeef0] rounded-lg bg-white shadow-sm p-6 space-y-4">
        <div className="flex items-center justify-between gap-2">
          <div className="flex items-center gap-2">
            <CalendarDays className="text-gray-400" size={18} />
            <h3 className="text-[9px] font-bold text-gray-500 uppercase tracking-widest">
              Assigned Cleaners
            </h3>
          </div>
          {canManageAssignments && (
            <button
              type="button"
              onClick={() => setShowAssignForm(true)}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-[11px] font-bold bg-gray-900 text-white hover:bg-black transition-colors"
            >
              <UserPlus size={14} />
              Assign cleaner
            </button>
          )}
        </div>

        {assignmentsError && (
          <p className="text-xs text-red-500">{assignmentsError}</p>
        )}

        {assignmentsLoading ? (
          <p className="text-xs text-gray-400">Loading current assignments…</p>
        ) : assignments.length === 0 ? (
          <p className="text-xs text-gray-400">
            No active cleaner assignments for this site.
          </p>
        ) : (
          <ul className="divide-y divide-[#edeef0]">
            {assignments.map((a) => (
              <li key={a.id} className="flex items-center justify-between py-2">
                <div className="flex flex-col">
                  {(() => {
                    let cleanerDisplayName =
                      a.cleanerName ||
                      (() => {
                        const match = cleaners.find((c) => c.id === a.cleanerId);
                        if (!match) return "";
                        const full = `${match.firstName} ${match.lastName}`.trim();
                        return full || match.firstName || match.lastName || "";
                      })();
                    if (!cleanerDisplayName && a.assignmentName) {
                      const parts = a.assignmentName.split(" - ");
                      const last = parts[parts.length - 1]?.trim();
                      cleanerDisplayName = last || a.assignmentName.trim();
                    }
                    const normalizedAssignment = (a.assignmentName || "").trim().toLowerCase();
                    const normalizedExpected = `${site.name || site.id} - ${cleanerDisplayName}`.trim().toLowerCase();
                    const showAssignmentName =
                      a.assignmentName &&
                      normalizedAssignment &&
                      normalizedAssignment !== normalizedExpected;

                    return (
                      <>
                        <span className="text-sm font-semibold text-gray-900">
                          {cleanerDisplayName || "Cleaner"}
                        </span>
                        {showAssignmentName && (
                          <span className="text-[11px] text-gray-400">
                            {a.assignmentName}
                          </span>
                        )}
                      </>
                    );
                  })()}
                </div>
                {canManageAssignments && (
                  <button
                    type="button"
                    onClick={() => handleDeactivateAssignment(a)}
                    className="text-[11px] font-bold text-red-500 hover:text-red-700"
                  >
                    Remove
                  </button>
                )}
              </li>
            ))}
          </ul>
        )}

        {showAssignForm && canManageAssignments && (
          <div className="border border-dashed border-[#edeef0] rounded-lg p-4 bg-gray-50 space-y-3">
            <div className="flex items-center justify-between">
              <p className="text-xs font-semibold text-gray-700">
                Assign cleaner to {site.name}
              </p>
              <button
                type="button"
                onClick={() => {
                  setShowAssignForm(false);
                  setAssignCleanerId('');
                }}
                className="text-gray-400 hover:text-gray-600"
              >
                <X size={14} />
              </button>
            </div>
            <div>
              <label className="block text-[10px] font-bold text-gray-400 uppercase mb-1">
                Cleaner
              </label>
              <div className="relative">
                <input
                  type="text"
                  value={assignCleanerSearch}
                  onChange={(e) => {
                    const value = e.target.value;
                    setAssignCleanerSearch(value);
                    setAssignCleanerId('');
                  }}
                  placeholder="Start typing cleaner name…"
                  className="w-full border border-[#edeef0] rounded-lg px-3 py-2 text-sm bg-white"
                  autoComplete="off"
                />
                {assignCleanerSearch.trim() && filteredCleanersForAssignment.length > 0 && !assignCleanerId && (
                  <div className="absolute z-10 mt-1 w-full max-h-56 overflow-auto rounded-lg border border-[#edeef0] bg-white shadow-lg">
                    {filteredCleanersForAssignment.map((c) => {
                      const label = `${c.firstName} ${c.lastName}`.trim() || c.email || c.id;
                      return (
                        <button
                          key={c.id}
                          type="button"
                          onClick={() => {
                            setAssignCleanerId(c.id);
                            setAssignCleanerSearch(label);
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
              {availableCleanersForAssignment.length === 0 && (
                <p className="mt-1 text-[11px] text-gray-400">
                  All cleaners are already assigned to this site.
                </p>
              )}
            </div>
            <div className="flex justify-end gap-2">
              <button
                type="button"
                onClick={() => {
                  setShowAssignForm(false);
                  setAssignCleanerId('');
                }}
                className="px-3 py-1.5 text-[11px] font-bold text-gray-500 hover:text-gray-700"
              >
                Cancel
              </button>
              <button
                type="button"
                disabled={!assignCleanerId || savingAssignment}
                onClick={handleCreateAssignment}
                className="px-3 py-1.5 text-[11px] font-bold rounded-lg bg-gray-900 text-white disabled:opacity-40"
              >
                {savingAssignment ? 'Saving…' : 'Save assignment'}
              </button>
            </div>
          </div>
        )}
      </div>

      {/* Admin Financial Audit Section */}
      {isAdmin && (
        <div className="border border-[#edeef0] rounded-lg bg-white shadow-sm p-6 space-y-4">
          <div className="flex items-center gap-2">
            <Wallet className="text-gray-400" size={18} />
            <h3 className="text-[9px] font-bold text-gray-500 uppercase tracking-widest">Financial Performance Audit</h3>
          </div>
          <div className="grid grid-cols-1 md:grid-cols-4 gap-4">
            <div className="p-3 rounded-lg border border-[#edeef0] bg-[#fcfcfb]">
              <p className="text-[9px] font-bold text-gray-500 uppercase tracking-widest mb-1">Contract revenue</p>
              <p className="text-lg font-bold text-gray-900">${financialStats.fortnightRevenue.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</p>
              <p className="text-[10px] text-gray-400 mt-0.5">(${((site as any).monthly_revenue ?? (site as any).monthlyRevenue ?? 0).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}/mo)</p>
            </div>
            <div className="p-3 rounded-lg border border-[#edeef0] bg-[#fcfcfb]">
              <p className="text-[9px] font-bold text-gray-500 uppercase tracking-widest mb-1">Actual labor cost</p>
              <p className="text-lg font-bold text-red-600">-${financialStats.actualLaborCost.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</p>
              <p className="text-[10px] text-gray-400 mt-0.5">Direct wages only</p>
            </div>
            <div className="p-3 rounded-lg border border-[#edeef0] bg-[#fcfcfb]">
              <p className="text-[9px] font-bold text-gray-500 uppercase tracking-widest mb-1">Gross profit</p>
              <p className="text-lg font-bold text-green-600">${financialStats.grossProfit.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</p>
              <p className="text-[10px] text-gray-400 mt-0.5">Before overheads</p>
            </div>
            <div className="p-3 rounded-lg border border-[#edeef0] bg-[#fcfcfb]">
              <p className="text-[9px] font-bold text-gray-500 uppercase tracking-widest mb-1">Profit margin</p>
              <div className="flex items-center gap-1.5">
                <p className={`text-lg font-bold ${financialStats.margin > 20 ? 'text-blue-600' : 'text-red-600'}`}>
                  {financialStats.margin.toFixed(1)}%
                </p>
                <Percent size={14} className="text-gray-300" />
              </div>
            </div>
          </div>
        </div>
      )}

      <div className="space-y-4">
        <div className="flex items-center gap-2">
          <CalendarDays className="text-gray-400" size={18} />
          <h3 className="text-[9px] font-bold text-gray-500 uppercase tracking-widest">Operational Variance Audit</h3>
        </div>

        <div className="grid grid-cols-2 sm:grid-cols-7 gap-2">
          {dailyStats.map((stat) => (
            <div key={stat.date} className="p-2.5 rounded-lg border border-[#edeef0] bg-white shadow-sm">
              <p className="text-[9px] font-bold text-gray-500 uppercase tracking-widest mb-1">{stat.displayDate}</p>
              <p className={`text-base font-bold ${stat.actual > stat.budget ? 'text-red-600' : 'text-gray-900'}`}>{stat.actual}h</p>
              <p className="text-[10px] text-gray-400 font-medium">Plan: {stat.budget}h</p>
            </div>
          ))}
        </div>

        <div className="h-[300px] w-full border border-[#edeef0] rounded-lg p-4 bg-white shadow-sm">
          <ResponsiveContainer width="100%" height="100%">
            <BarChart data={dailyStats}>
              <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f1f1" />
              <XAxis dataKey="displayDate" axisLine={false} tickLine={false} tick={{fill: '#a3a3a3', fontSize: 10}} />
              <YAxis axisLine={false} tickLine={false} tick={{fill: '#a3a3a3', fontSize: 10}} />
              <Tooltip 
                contentStyle={{ borderRadius: '8px', border: '1px solid #edeef0', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)', fontSize: '12px' }}
              />
              <Bar dataKey="actual" name="Actual Hours" radius={[4, 4, 0, 0]}>
                {dailyStats.map((entry, index) => (
                  <Cell key={`cell-${index}`} fill={entry.actual > entry.budget ? '#ef4444' : '#111827'} />
                ))}
              </Bar>
            </BarChart>
          </ResponsiveContainer>
        </div>
      </div>
    </div>
  );
};

export default SiteDetail;
