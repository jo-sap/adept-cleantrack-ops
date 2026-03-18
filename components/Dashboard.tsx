
import React, { useMemo, useState, useEffect } from 'react';
import { Site, TimeEntry, FortnightPeriod, Cleaner } from '../types';
import { ChevronRight, DollarSign, PieChart, Briefcase, ChevronUp, ChevronDown } from 'lucide-react';
import MiniComplianceGrid from './MiniComplianceGrid';
import { useRole } from '../contexts/RoleContext';
import { useAppAuth } from '../contexts/AppAuthContext';
import { getGraphAccessToken } from '../lib/graph';
import { getDashboardMetrics, type DashboardMetrics } from '../repositories/metricsRepo';
import { getAssignedSiteIdsForManager } from '../repositories/siteManagersRepo';
import { getAdHocJobs } from '../repositories/adHocJobsRepo';
import { getCleanTrackUserByEmail } from '../repositories/usersRepo';
import { DEV_BYPASS_LOGIN } from '../config/authFlags';
import { formatCurrencyAUD, formatCurrencyAUDSignedExpense, formatPercent } from '../utils';
import { computeBudgetedLabourCostForRange } from '../lib/budgetedLabourCost';
import { getPublicHolidaysInRange } from '../lib/publicHolidays';
import { format } from 'date-fns';

interface DashboardProps {
  sites: Site[];
  cleaners: Cleaner[];
  entries: TimeEntry[];
  currentPeriod: FortnightPeriod;
  onViewSite: (siteId: string) => void;
}

const Dashboard: React.FC<DashboardProps> = ({ sites, cleaners, entries, currentPeriod, onViewSite }) => {
  const { isAdmin: isAdminFromRole } = useRole();
  const { user } = useAppAuth();
  /** Strict: no GP/revenue visible to Managers — Admin only */
  const isAdmin = (isAdminFromRole || user?.role === "Admin") && user?.role !== "Manager";
  const [kpiMetrics, setKpiMetrics] = useState<DashboardMetrics | null>(null);
  const [kpiLoading, setKpiLoading] = useState(false);
  const [kpiError, setKpiError] = useState<string | null>(null);
  const [adHocStats, setAdHocStats] = useState<{
    total: number;
    completed: number;
    pending: number;
    budgetedHours: number;
    revenue: number;
    grossProfit: number;
  } | null>(null);

  useEffect(() => {
    if (!isAdmin) return;
    if (DEV_BYPASS_LOGIN) {
      setKpiMetrics({
        portfolioRevenue: 10000,
        laborExpenses: 7000,
        netGrossProfit: 3000,
        profitMargin: 0.3,
      });
      setKpiError(null);
      return;
    }
    let cancelled = false;
    setKpiLoading(true);
    setKpiError(null);
    getGraphAccessToken()
      .then((token) => {
        if (cancelled) return;
        if (!token) {
          setKpiError("Microsoft Graph not available.");
          setKpiMetrics({ portfolioRevenue: 0, laborExpenses: 0, netGrossProfit: 0, profitMargin: 0 });
          setKpiLoading(false);
          return;
        }
        const range = {
          start: currentPeriod.startDate,
          end: new Date(currentPeriod.endDate.getTime() + 24 * 60 * 60 * 1000),
        };
        const options =
          user?.role === "Manager" && user?.email && user.permissionScope?.trim().toLowerCase() !== "allsites"
            ? { assignedSiteIds: [] as string[] }
            : undefined;
        if (options && user?.email) {
          return getAssignedSiteIdsForManager(token, user.email).then((ids) => {
            if (cancelled) return;
            options.assignedSiteIds = ids;
            return getDashboardMetrics(token, range, options).then((res) => {
              if (cancelled) return;
              setKpiMetrics(res.metrics);
              setKpiError(res.error ?? null);
            });
          });
        }
        return getDashboardMetrics(token, range).then((res) => {
          if (cancelled) return;
          setKpiMetrics(res.metrics);
          setKpiError(res.error ?? null);
        });
      })
      .then(() => { if (!cancelled) setKpiLoading(false); })
      .catch((err) => {
        if (!cancelled) {
          setKpiError(err instanceof Error ? err.message : "Failed to load dashboard metrics.");
          setKpiMetrics(null);
          setKpiLoading(false);
        }
      });
    return () => { cancelled = true; };
  }, [isAdmin, currentPeriod.startDate.getTime(), currentPeriod.endDate.getTime(), user?.role, user?.email, user?.permissionScope]);

  useEffect(() => {
    let cancelled = false;
    const monthStr = format(currentPeriod.startDate, "yyyy-MM");
    getGraphAccessToken().then(async (token) => {
      if (!token || cancelled) return;
      try {
        const filters: { month: string; assignedManagerId?: string } = { month: monthStr };
        if (!isAdmin && user?.email) {
          const ctUser = await getCleanTrackUserByEmail(token, user.email);
          if (ctUser?.id) filters.assignedManagerId = ctUser.id;
        }
        const jobs = await getAdHocJobs(token, filters);
        if (cancelled) return;
        const completed = jobs.filter((j) => j.status === "Completed").length;
        const pending = jobs.filter((j) => j.status !== "Completed" && j.status !== "Cancelled").length;
        const budgetedHours = jobs.reduce((s, j) => s + (j.budgetedHours ?? 0), 0);
        const revenue = jobs.reduce((s, j) => s + (j.charge ?? 0), 0);
        const grossProfit = jobs.reduce((s, j) => {
          if (j.grossProfit != null) return s + j.grossProfit;
          const charge = j.charge ?? 0;
          const cost = j.cost ?? 0;
          return s + (charge - cost);
        }, 0);
        setAdHocStats({
          total: jobs.length,
          completed,
          pending,
          budgetedHours,
          revenue,
          grossProfit,
        });
      } catch {
        if (!cancelled) setAdHocStats(null);
      }
    });
    return () => { cancelled = true; };
  }, [currentPeriod.startDate, isAdmin, user?.email]);

  type RecapSortKey = 'name' | 'volume' | 'variance' | 'budget' | 'cleaner';
  const [recapSortBy, setRecapSortBy] = useState<RecapSortKey>('name');
  const [recapSortDir, setRecapSortDir] = useState<'asc' | 'desc'>('asc');

  // Map primary assigned cleaner name per site for quick lookup in the portfolio table
  const primaryCleanerNameBySiteId = useMemo(() => {
    const map: Record<string, string> = {};
    sites.forEach((site) => {
      const assignedIds = site.assigned_cleaner_ids ?? [];
      if (assignedIds.length === 0) return;
      const primaryId = assignedIds[0];
      const cleaner = cleaners.find((c) => c.id === primaryId);
      if (cleaner) {
        const fullName = `${cleaner.firstName} ${cleaner.lastName}`.trim();
        map[site.id] = fullName || cleaner.firstName || cleaner.lastName || "";
      }
    });
    return map;
  }, [sites, cleaners]);

  // Pre-process entries into a map for efficient lookups
  const siteDailyMap = useMemo(() => {
    const map: Record<string, Record<string, number>> = {};
    entries.forEach(e => {
      const date = new Date(e.date);
      if (date >= currentPeriod.startDate && date <= currentPeriod.endDate) {
        if (!map[e.siteId!]) map[e.siteId!] = {};
        map[e.siteId!][e.date] = (map[e.siteId!][e.date] || 0) + e.hours;
      }
    });
    return map;
  }, [entries, currentPeriod]);

  // Financial Metrics Pass
  const financialMetrics = useMemo(() => {
    let totalPortfolioRevenue = 0;
    let totalPortfolioLaborCost = 0;

    const recap = sites.map(site => {
      const dailyActuals = siteDailyMap[site.id] || {};
      let actualHoursTotal = 0;
      let laborCostTotal = 0;

      // Filter entries for this site and period
      const siteEntries = entries.filter(e => {
        const d = new Date(e.date);
        return e.siteId === site.id && d >= currentPeriod.startDate && d <= currentPeriod.endDate;
      });

      siteEntries.forEach(e => {
        actualHoursTotal += e.hours;
        const cleaner = cleaners.find(c => c.id === e.cleanerId);
        const rate = e.pay_rate_snapshot || cleaner?.payRatePerHour || 0;
        laborCostTotal += e.hours * rate;
      });

      const budgetHoursTotal = site.budgeted_hours_per_fortnight;
      const varianceHours = actualHoursTotal - budgetHoursTotal;
      
      // Fortnightly Revenue = (Monthly * 12) / 26
      const fortnightRevenue = (site.monthly_revenue * 12) / 26;
      const weekdayRate = site.budget_weekday_labour_rate ?? site.budget_labour_rate ?? 0;
      const saturdayRate = site.budget_saturday_labour_rate ?? weekdayRate;
      const sundayRate = site.budget_sunday_labour_rate ?? weekdayRate;
      const phRate = site.budget_ph_labour_rate ?? weekdayRate;
      const dailyBudgets = site.daily_budgets ?? [0, 0, 0, 0, 0, 0, 0]; // [Sun, Mon, ..., Sat]
      // Date-aware budget: each day in the period uses PH / Sat / Sun / Weekday rate so budget matches actuals
      const phInPeriod = getPublicHolidaysInRange(currentPeriod.startDate, currentPeriod.endDate);
      const budgetedLabourCost = computeBudgetedLabourCostForRange({
        startDate: currentPeriod.startDate,
        endDate: currentPeriod.endDate,
        dailyBudgets,
        weekdayRate,
        saturdayRate,
        sundayRate,
        phRate,
        publicHolidayDates: phInPeriod,
      });
      const grossProfit = fortnightRevenue - laborCostTotal;
      const margin = fortnightRevenue > 0 ? (grossProfit / fortnightRevenue) * 100 : 0;

      totalPortfolioRevenue += fortnightRevenue;
      totalPortfolioLaborCost += laborCostTotal;

      return {
        id: site.id,
        name: site.name,
        actualHoursTotal,
        budgetHoursTotal,
        varianceHours,
        fortnightRevenue,
        laborCostTotal,
        budgetedLabourCost,
        grossProfit,
        margin,
        dailyActuals
      };
    });

    return { 
      recap, 
      totalPortfolioRevenue, 
      totalPortfolioLaborCost,
      totalGrossProfit: totalPortfolioRevenue - totalPortfolioLaborCost,
      portfolioMargin: totalPortfolioRevenue > 0 ? ((totalPortfolioRevenue - totalPortfolioLaborCost) / totalPortfolioRevenue) * 100 : 0
    };
  }, [sites, cleaners, entries, currentPeriod, siteDailyMap]);

  // Use API KPIs when present and non-zero; otherwise use client-computed financial metrics (sites + entries)
  const displayMetrics = useMemo(() => {
    const fromApi = kpiMetrics && (kpiMetrics.portfolioRevenue > 0 || kpiMetrics.laborExpenses > 0);
    if (fromApi && kpiMetrics) {
      return {
        portfolioRevenue: kpiMetrics.portfolioRevenue,
        laborExpenses: kpiMetrics.laborExpenses,
        netGrossProfit: kpiMetrics.netGrossProfit,
        profitMargin: kpiMetrics.profitMargin,
      };
    }
    return {
      portfolioRevenue: financialMetrics.totalPortfolioRevenue,
      laborExpenses: financialMetrics.totalPortfolioLaborCost,
      netGrossProfit: financialMetrics.totalGrossProfit,
      profitMargin: financialMetrics.portfolioMargin / 100,
    };
  }, [kpiMetrics, financialMetrics]);

  const sortedRecap = useMemo(() => {
    const list = [...financialMetrics.recap];
    list.sort((a, b) => {
      let cmp = 0;
      if (recapSortBy === 'name') {
        const na = (a.name || '').toLowerCase();
        const nb = (b.name || '').toLowerCase();
        cmp = na.localeCompare(nb);
      } else if (recapSortBy === 'volume') {
        cmp = a.actualHoursTotal - b.actualHoursTotal;
      } else if (recapSortBy === 'variance') {
        cmp = a.varianceHours - b.varianceHours;
      } else if (recapSortBy === 'budget') {
        cmp = a.budgetedLabourCost - b.budgetedLabourCost;
      } else if (recapSortBy === 'cleaner') {
        const ca = (primaryCleanerNameBySiteId[a.id] || '').toLowerCase();
        const cb = (primaryCleanerNameBySiteId[b.id] || '').toLowerCase();
        cmp = ca.localeCompare(cb);
      }
      return recapSortDir === 'asc' ? cmp : -cmp;
    });
    return list;
  }, [financialMetrics.recap, recapSortBy, recapSortDir, primaryCleanerNameBySiteId]);

  const handleRecapSort = (key: RecapSortKey) => {
    if (recapSortBy === key) {
      setRecapSortDir((d) => (d === 'asc' ? 'desc' : 'asc'));
    } else {
      setRecapSortBy(key);
      setRecapSortDir(key === 'name' || key === 'cleaner' ? 'asc' : 'desc');
    }
  };

  return (
    <div className="space-y-6 sm:space-y-8 animate-fadeIn">
      {/* Financial KPIs (Admin only) */}
      {isAdmin && (
        <>
          {kpiError && (
            <p className="text-sm text-amber-700 bg-amber-50 border border-amber-200 rounded-lg px-4 py-2">
              {kpiError}
            </p>
          )}
          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
            <div className="so-card-soft px-5 py-5 flex items-start gap-3">
              <div className="mt-1 rounded-lg bg-[#ECF3F4] text-[#3E5F6A] p-2">
                <DollarSign size={16} />
              </div>
              <div className="space-y-1">
                <p className="text-[11px] font-semibold text-gray-500 uppercase tracking-[0.18em]">
                  Portfolio Revenue
                </p>
                {kpiLoading ? (
                  <h3 className="text-[24px] font-semibold text-gray-400 leading-tight">Loading…</h3>
                ) : (
                  <h3 className="text-[24px] font-semibold text-gray-900 leading-tight">
                    {formatCurrencyAUD(displayMetrics.portfolioRevenue)}
                  </h3>
                )}
              </div>
            </div>
            <div className="so-card-soft px-5 py-5 flex items-start gap-3">
              <div className="mt-1 rounded-lg bg-[#FEF3C7] text-[#92400E] p-2">
                <DollarSign size={16} />
              </div>
              <div className="space-y-1">
                <p className="text-[11px] font-semibold text-gray-500 uppercase tracking-[0.18em]">
                  Labor Expenses
                </p>
                {kpiLoading ? (
                  <h3 className="text-[24px] font-semibold text-gray-400 leading-tight">Loading…</h3>
                ) : (
                  <h3 className="text-[24px] font-semibold text-red-600 leading-tight">
                    {formatCurrencyAUDSignedExpense(displayMetrics.laborExpenses)}
                  </h3>
                )}
              </div>
            </div>
            <div className="so-card-soft px-5 py-5 flex items-start gap-3">
              <div className="mt-1 rounded-lg bg-[#DCFCE7] text-[#166534] p-2">
                <Briefcase size={16} />
              </div>
              <div className="space-y-1">
                <p className="text-[11px] font-semibold text-gray-500 uppercase tracking-[0.18em]">
                  Net Gross Profit
                </p>
                {kpiLoading ? (
                  <h3 className="text-[24px] font-semibold text-gray-400 leading-tight">Loading…</h3>
                ) : (
                  <h3 className="text-[24px] font-semibold text-emerald-700 leading-tight">
                    {formatCurrencyAUD(displayMetrics.netGrossProfit)}
                  </h3>
                )}
              </div>
            </div>
            <div className="so-card-soft px-5 py-5 flex items-start gap-3">
              <div className="mt-1 rounded-lg bg-[#E0F2FE] text-[#0369A1] p-2">
                <PieChart size={16} />
              </div>
              <div className="space-y-1">
                <p className="text-[11px] font-semibold text-gray-500 uppercase tracking-[0.18em]">
                  Profit Margin
                </p>
                {kpiLoading ? (
                  <h3 className="text-[24px] font-semibold text-gray-400 leading-tight">Loading…</h3>
                ) : (
                  <div className="flex items-center gap-2">
                    <h3 className="text-[24px] font-semibold text-gray-900 leading-tight">
                      {formatPercent(displayMetrics.profitMargin)}
                    </h3>
                    <PieChart size={18} className="text-gray-500" />
                  </div>
                )}
              </div>
            </div>
          </div>
        </>
      )}

      {/* Ad Hoc Jobs summary – visible to all */}
      {adHocStats != null && (
        <div className="space-y-2">
          <h3 className="text-lg sm:text-xl font-semibold text-gray-900 flex items-center gap-2">
            <Briefcase size={20} className="text-gray-500" />
            Ad Hoc Jobs this month
          </h3>
          <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
            <div className="so-card-soft p-4">
              <p className="text-[10px] font-semibold text-gray-500 uppercase tracking-[0.18em]">
                Total
              </p>
              <p className="text-xl font-semibold text-gray-900">
                {adHocStats.total}
              </p>
            </div>
            <div className="so-card-soft p-4">
              <p className="text-[10px] font-semibold text-gray-500 uppercase tracking-[0.18em]">
                Completed
              </p>
              <p className="text-xl font-semibold text-emerald-700">
                {adHocStats.completed}
              </p>
            </div>
            <div className="so-card-soft p-4">
              <p className="text-[10px] font-semibold text-gray-500 uppercase tracking-[0.18em]">
                Pending
              </p>
              <p className="text-xl font-semibold text-gray-700">
                {adHocStats.pending}
              </p>
            </div>
            <div className="so-card-soft p-4">
              {isAdmin ? (
                <>
                  <p className="text-[10px] font-semibold text-gray-500 uppercase tracking-[0.18em]">
                    Revenue
                  </p>
                  <p className="text-xl font-semibold text-gray-900">
                    {formatCurrencyAUD(adHocStats.revenue)}
                  </p>
                  <p className="mt-1 text-[10px] font-semibold text-gray-500 uppercase tracking-[0.18em]">
                    GP
                  </p>
                  <p className="text-sm font-semibold text-emerald-700">
                    {formatCurrencyAUD(adHocStats.grossProfit)}
                  </p>
                </>
              ) : (
                <>
                  <p className="text-[10px] font-semibold text-gray-500 uppercase tracking-[0.18em]">
                    Budgeted hrs
                  </p>
                  <p className="text-xl font-semibold text-gray-900">
                    {adHocStats.budgetedHours.toFixed(1)}h
                  </p>
                </>
              )}
            </div>
          </div>
        </div>
      )}

      {/* Site Audit Table - match Sites table styling */}
      <div className="space-y-4 sm:space-y-6">
        <h3 className="text-lg sm:text-xl font-semibold text-gray-900">
          Portfolio Compliance &amp; Performance
        </h3>
        <div className="so-table bg-white overflow-hidden table-scroll-mobile">
          <table className="w-full border-collapse text-left table-fixed min-w-[760px]">
            <colgroup>
              {isAdmin ? (
                <>
                  <col style={{ width: '23%' }} />
                  <col style={{ width: '11%' }} />
                  <col style={{ width: '13%' }} />
                  <col style={{ width: '17%' }} />
                  <col style={{ width: '18%' }} />
                  <col style={{ width: '13%' }} />
                  <col style={{ width: '5%' }} />
                </>
              ) : (
                <>
                  <col style={{ width: '28%' }} />
                  <col style={{ width: '12%' }} />
                  <col style={{ width: '14%' }} />
                  <col style={{ width: '28%' }} />
                  <col style={{ width: '18%' }} />
                  <col style={{ width: '5%' }} />
                </>
              )}
            </colgroup>
            <thead>
              <tr className="border-b border-[#edeef0]">
                <th className="text-left px-1.5 py-1.5">
                  <button
                    type="button"
                    onClick={() => handleRecapSort('name')}
                    className="text-[10px] font-semibold text-gray-700 uppercase tracking-widest flex items-center gap-0.5 hover:text-gray-900 focus:outline-none focus:ring-2 focus:ring-gray-900/20 rounded"
                  >
                    Site Name
                    {recapSortBy === 'name' && (recapSortDir === 'asc' ? <ChevronUp size={10} /> : <ChevronDown size={10} />)}
                  </button>
                </th>
                <th className="text-center px-1.5 py-1.5">
                  <button
                    type="button"
                    onClick={() => handleRecapSort('volume')}
                    className="text-[10px] font-semibold text-gray-700 uppercase tracking-widest inline-flex items-center gap-0.5 hover:text-gray-900 focus:outline-none focus:ring-2 focus:ring-gray-900/20 rounded"
                  >
                    Volume (Hrs)
                    {recapSortBy === 'volume' && (recapSortDir === 'asc' ? <ChevronUp size={10} /> : <ChevronDown size={10} />)}
                  </button>
                </th>
                <th className="text-center px-1.5 py-1.5">
                  <button
                    type="button"
                    onClick={() => handleRecapSort('variance')}
                    className="text-[10px] font-semibold text-gray-700 uppercase tracking-widest inline-flex items-center gap-0.5 hover:text-gray-900 focus:outline-none focus:ring-2 focus:ring-gray-900/20 rounded"
                  >
                    Variance
                    {recapSortBy === 'variance' && (recapSortDir === 'asc' ? <ChevronUp size={10} /> : <ChevronDown size={10} />)}
                  </button>
                </th>
                {isAdmin && (
                  <th className="text-center px-1.5 py-1.5">
                    <button
                      type="button"
                      onClick={() => handleRecapSort('budget')}
                      className="text-[10px] font-semibold text-gray-700 uppercase tracking-widest inline-flex items-center gap-0.5 hover:text-gray-900 focus:outline-none focus:ring-2 focus:ring-gray-900/20 rounded"
                    >
                      Budget labour
                      {recapSortBy === 'budget' && (recapSortDir === 'asc' ? <ChevronUp size={10} /> : <ChevronDown size={10} />)}
                    </button>
                  </th>
                )}
                <th className="text-center px-1.5 py-1.5 text-[10px] font-semibold text-gray-700 uppercase tracking-widest">14D Pattern</th>
                <th className="text-left px-1.5 py-1.5">
                  <button
                    type="button"
                    onClick={() => handleRecapSort('cleaner')}
                    className="text-[9px] font-bold text-gray-500 uppercase tracking-widest inline-flex items-center gap-0.5 hover:text-gray-900 focus:outline-none focus:ring-2 focus:ring-gray-900/20 rounded"
                  >
                    Cleaner Assigned
                    {recapSortBy === 'cleaner' && (recapSortDir === 'asc' ? <ChevronUp size={10} /> : <ChevronDown size={10} />)}
                  </button>
                </th>
                <th className="px-1 py-1.5 text-right"></th>
              </tr>
            </thead>
            <tbody className="divide-y divide-[#edeef0]">
              {sortedRecap.map((recap) => {
                const primaryCleanerName = primaryCleanerNameBySiteId[recap.id];
                const variancePositive = recap.varianceHours > 0.1;
                const varianceNegative = recap.varianceHours < -0.1;
                return (
                <tr 
                  key={recap.id} 
                  onClick={() => onViewSite(recap.id)}
                  className="transition-colors cursor-pointer group"
                >
                  <td className="px-1.5 py-1.5">
                    <span className="text-xs font-semibold text-gray-900 break-words">{recap.name}</span>
                  </td>
                  <td className="px-1.5 py-1.5 text-center">
                    <span className="text-[11px] font-medium text-gray-700">{recap.actualHoursTotal.toFixed(1)}h</span>
                  </td>
                  <td className="px-1.5 py-1.5 text-center">
                    <span
                      className={`inline-flex items-center justify-center px-2 py-0.5 rounded-full text-[10px] font-semibold ${
                        variancePositive
                          ? 'bg-red-50 text-red-700'
                          : varianceNegative
                          ? 'bg-emerald-50 text-emerald-700'
                          : 'bg-gray-50 text-gray-600'
                      }`}
                    >
                      {variancePositive ? '+' : ''}
                      {recap.varianceHours.toFixed(1)}h
                    </span>
                  </td>
                  {isAdmin && (
                    <td className="px-1.5 py-1.5 text-center">
                      <span className="text-[11px] font-medium text-gray-700">{formatCurrencyAUD(recap.budgetedLabourCost)}</span>
                    </td>
                  )}
                  <td className="px-1.5 py-1.5 flex justify-center">
                    <MiniComplianceGrid 
                      startDate={currentPeriod.startDate}
                      dailyBudgets={sites.find(s => s.id === recap.id)?.daily_budgets || []}
                      actualsByDate={recap.dailyActuals}
                    />
                  </td>
                  <td className="px-1.5 py-1.5">
                    {primaryCleanerName ? (
                      <span className="text-[11px] font-medium text-gray-700 break-words">
                        {primaryCleanerName}
                      </span>
                    ) : (
                      <span className="text-[10px] text-gray-400">—</span>
                    )}
                  </td>
                  <td className="px-1 py-1.5 text-right align-middle">
                    <ChevronRight size={12} className="text-gray-300 group-hover:text-gray-900 transition-colors inline-block" />
                  </td>
                </tr>
              ); })}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
};

export default Dashboard;
