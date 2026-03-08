
import React, { useMemo, useState, useEffect } from 'react';
import { Site, TimeEntry, FortnightPeriod, Cleaner } from '../types';
import { ChevronRight, DollarSign, PieChart, Briefcase } from 'lucide-react';
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
  const [adHocStats, setAdHocStats] = useState<{ total: number; completed: number; pending: number; budgetedHours: number; actualHours: number } | null>(null);

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
        setAdHocStats({
          total: jobs.length,
          completed,
          pending,
          budgetedHours,
          actualHours: 0,
        });
      } catch {
        if (!cancelled) setAdHocStats(null);
      }
    });
    return () => { cancelled = true; };
  }, [currentPeriod.startDate, isAdmin, user?.email]);

  const adHocActualHours = useMemo(() => {
    return entries
      .filter((e) => e.date >= format(currentPeriod.startDate, "yyyy-MM-dd") && e.date <= format(currentPeriod.endDate, "yyyy-MM-dd") && e.adhocJobId)
      .reduce((s, e) => s + e.hours, 0);
  }, [entries, currentPeriod]);

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
      const budgetLabourRate = site.budget_labour_rate ?? 0;
      const budgetedLabourCost = budgetHoursTotal * budgetLabourRate;
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
          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3 sm:gap-4">
            <div className="bg-[#f7f6f3] p-5 rounded-lg border border-[#edeef0]">
              <p className="text-[11px] font-bold text-gray-400 uppercase tracking-wider mb-1">Portfolio Revenue</p>
              {kpiLoading ? (
                <h3 className="text-2xl font-bold text-gray-400">Loading…</h3>
              ) : (
                <h3 className="text-2xl font-bold text-gray-900">
                  {formatCurrencyAUD(displayMetrics.portfolioRevenue)}
                </h3>
              )}
            </div>
            <div className="bg-[#f7f6f3] p-5 rounded-lg border border-[#edeef0]">
              <p className="text-[11px] font-bold text-gray-400 uppercase tracking-wider mb-1">Labor Expenses</p>
              {kpiLoading ? (
                <h3 className="text-2xl font-bold text-gray-400">Loading…</h3>
              ) : (
                <h3 className="text-2xl font-bold text-red-600">
                  {formatCurrencyAUDSignedExpense(displayMetrics.laborExpenses)}
                </h3>
              )}
            </div>
            <div className="bg-[#f7f6f3] p-5 rounded-lg border border-[#edeef0]">
              <p className="text-[11px] font-bold text-gray-400 uppercase tracking-wider mb-1">Net Gross Profit</p>
              {kpiLoading ? (
                <h3 className="text-2xl font-bold text-gray-400">Loading…</h3>
              ) : (
                <h3 className="text-2xl font-bold text-green-700">
                  {formatCurrencyAUD(displayMetrics.netGrossProfit)}
                </h3>
              )}
            </div>
            <div className="bg-[#f7f6f3] p-5 rounded-lg border border-[#edeef0]">
              <p className="text-[11px] font-bold text-gray-400 uppercase tracking-wider mb-1">Profit Margin</p>
              {kpiLoading ? (
                <h3 className="text-2xl font-bold text-gray-400">Loading…</h3>
              ) : (
                <div className="flex items-center gap-2">
                  <h3 className="text-2xl font-bold text-gray-900">
                    {formatPercent(displayMetrics.profitMargin)}
                  </h3>
                  <PieChart size={18} className="text-gray-500" />
                </div>
              )}
            </div>
          </div>
        </>
      )}

      {/* Ad Hoc Jobs summary – visible to all */}
      {adHocStats != null && (
        <div className="space-y-2">
          <h3 className="text-lg sm:text-xl font-bold text-gray-900 flex items-center gap-2">
            <Briefcase size={20} className="text-gray-500" />
            Ad Hoc Jobs this month
          </h3>
          <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
            <div className="bg-[#f7f6f3] p-4 rounded-lg border border-[#edeef0]">
              <p className="text-[10px] font-bold text-gray-400 uppercase tracking-wider">Total</p>
              <p className="text-xl font-bold text-gray-900">{adHocStats.total}</p>
            </div>
            <div className="bg-[#f7f6f3] p-4 rounded-lg border border-[#edeef0]">
              <p className="text-[10px] font-bold text-gray-400 uppercase tracking-wider">Completed</p>
              <p className="text-xl font-bold text-green-700">{adHocStats.completed}</p>
            </div>
            <div className="bg-[#f7f6f3] p-4 rounded-lg border border-[#edeef0]">
              <p className="text-[10px] font-bold text-gray-400 uppercase tracking-wider">Pending</p>
              <p className="text-xl font-bold text-gray-700">{adHocStats.pending}</p>
            </div>
            <div className="bg-[#f7f6f3] p-4 rounded-lg border border-[#edeef0]">
              <p className="text-[10px] font-bold text-gray-400 uppercase tracking-wider">Budgeted / Actual hrs</p>
              <p className="text-xl font-bold text-gray-900">{adHocStats.budgetedHours.toFixed(1)}h / {adHocActualHours.toFixed(1)}h</p>
            </div>
          </div>
        </div>
      )}

      {/* Site Audit Table - match Sites table styling */}
      <div className="space-y-4 sm:space-y-6">
        <h3 className="text-lg sm:text-xl font-bold text-gray-900">Portfolio Compliance & Performance</h3>
        <div className="border border-[#edeef0] rounded-lg bg-white shadow-sm overflow-hidden table-scroll-mobile">
          <table className="w-full border-collapse text-left table-fixed min-w-[700px]">
            <colgroup>
              {isAdmin ? (
                <>
                  <col style={{ width: '24%' }} />
                  <col style={{ width: '12%' }} />
                  <col style={{ width: '14%' }} />
                  <col style={{ width: '18%' }} />
                  <col style={{ width: '22%' }} />
                  <col style={{ width: '10%' }} />
                </>
              ) : (
                <>
                  <col style={{ width: '28%' }} />
                  <col style={{ width: '12%' }} />
                  <col style={{ width: '14%' }} />
                  <col style={{ width: '36%' }} />
                  <col style={{ width: '10%' }} />
                </>
              )}
            </colgroup>
            <thead>
              <tr className="bg-[#fcfcfb] border-b border-[#edeef0]">
                <th className="text-left px-1.5 py-1.5 text-[9px] font-bold text-gray-500 uppercase tracking-widest">Site Name</th>
                <th className="text-center px-1.5 py-1.5 text-[9px] font-bold text-gray-500 uppercase tracking-widest">Volume (Hrs)</th>
                <th className="text-center px-1.5 py-1.5 text-[9px] font-bold text-gray-500 uppercase tracking-widest">Variance</th>
                {isAdmin && <th className="text-center px-1.5 py-1.5 text-[9px] font-bold text-gray-500 uppercase tracking-widest">Budget labour</th>}
                <th className="text-center px-1.5 py-1.5 text-[9px] font-bold text-gray-500 uppercase tracking-widest">14D Pattern</th>
                <th className="px-1.5 py-1.5"></th>
              </tr>
            </thead>
            <tbody className="divide-y divide-[#edeef0]">
              {financialMetrics.recap.map((recap) => (
                <tr 
                  key={recap.id} 
                  onClick={() => onViewSite(recap.id)}
                  className="hover:bg-[#f7f6f3] transition-colors cursor-pointer group"
                >
                  <td className="px-1.5 py-1.5">
                    <span className="text-xs font-semibold text-gray-900 break-words">{recap.name}</span>
                  </td>
                  <td className="px-1.5 py-1.5 text-center">
                    <span className="text-[11px] font-medium text-gray-700">{recap.actualHoursTotal.toFixed(1)}h</span>
                  </td>
                  <td className="px-1.5 py-1.5 text-center">
                    <span className={`text-[11px] font-bold ${recap.varianceHours > 0.1 ? 'text-red-600' : 'text-green-600'}`}>
                      {recap.varianceHours > 0.1 ? '+' : ''}{recap.varianceHours.toFixed(1)}h
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
                  <td className="px-1.5 py-1.5 text-right">
                    <ChevronRight size={14} className="text-gray-300 group-hover:text-gray-900 transition-colors" />
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
};

export default Dashboard;
