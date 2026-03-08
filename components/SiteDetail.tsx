
import React, { useMemo } from 'react';
import { XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, BarChart, Bar, Cell } from 'recharts';
import { Site, TimeEntry, FortnightPeriod, Cleaner } from '../types';
import { format, addDays, getDay } from 'date-fns';
import { ArrowLeft, CalendarDays, Wallet, Percent } from 'lucide-react';
import { useRole } from '../contexts/RoleContext';
import { useAppAuth } from '../contexts/AppAuthContext';

interface SiteDetailProps {
  site: Site;
  cleaners: Cleaner[];
  entries: TimeEntry[];
  currentPeriod: FortnightPeriod;
  onBack: () => void;
}

const SiteDetail: React.FC<SiteDetailProps> = ({ site, cleaners, entries, currentPeriod, onBack }) => {
  const { isAdmin: isAdminFromRole } = useRole();
  const { user } = useAppAuth();
  /** Strict: no GP/revenue visible to Managers — Admin only */
  const isAdmin = (isAdminFromRole || user?.role === 'Admin') && user?.role !== 'Manager';

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

  const dailyStats = useMemo(() => {
    const stats = [];
    for (let i = 0; i < 14; i++) {
      const day = addDays(currentPeriod.startDate, i);
      const dayStr = format(day, 'yyyy-MM-dd');
      const dayOfWeek = getDay(day);
      const actual = periodEntries
        .filter(e => e.date === dayStr)
        .reduce((sum, e) => sum + e.hours, 0);

      const dayBudget = dailyBudgetsArr[dayOfWeek] ?? 0;
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
  }, [site, periodEntries, currentPeriod, dailyBudgetsArr]);

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
            Back to Dashboard
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
              <p className="text-lg font-bold text-gray-900">${financialStats.fortnightRevenue.toLocaleString(undefined, { minimumFractionDigits: 2 })}</p>
              <p className="text-[10px] text-gray-400 mt-0.5">(${((site as any).monthly_revenue ?? (site as any).monthlyRevenue ?? 0).toLocaleString()}/mo)</p>
            </div>
            <div className="p-3 rounded-lg border border-[#edeef0] bg-[#fcfcfb]">
              <p className="text-[9px] font-bold text-gray-500 uppercase tracking-widest mb-1">Actual labor cost</p>
              <p className="text-lg font-bold text-red-600">-${financialStats.actualLaborCost.toLocaleString(undefined, { minimumFractionDigits: 2 })}</p>
              <p className="text-[10px] text-gray-400 mt-0.5">Direct wages only</p>
            </div>
            <div className="p-3 rounded-lg border border-[#edeef0] bg-[#fcfcfb]">
              <p className="text-[9px] font-bold text-gray-500 uppercase tracking-widest mb-1">Gross profit</p>
              <p className="text-lg font-bold text-green-600">${financialStats.grossProfit.toLocaleString(undefined, { minimumFractionDigits: 2 })}</p>
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
