
import React, { useMemo } from 'react';
import { XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, BarChart, Bar, Cell } from 'recharts';
import { Site, TimeEntry, FortnightPeriod, Cleaner } from '../types';
import { format, addDays, getDay } from 'date-fns';
import { ArrowLeft, CalendarDays, TrendingUp, TrendingDown, DollarSign, Wallet, Percent } from 'lucide-react';
import { useRole } from '../contexts/RoleContext';

interface SiteDetailProps {
  site: Site;
  cleaners: Cleaner[];
  entries: TimeEntry[];
  currentPeriod: FortnightPeriod;
  onBack: () => void;
}

const SiteDetail: React.FC<SiteDetailProps> = ({ site, cleaners, entries, currentPeriod, onBack }) => {
  const { isAdmin } = useRole();

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
    <div className="space-y-8 animate-fadeIn">
      <button 
        onClick={onBack}
        className="flex items-center gap-2 text-sm text-gray-500 hover:text-gray-900 transition-colors mb-4 group"
      >
        <ArrowLeft size={16} className="group-hover:-translate-x-1 transition-transform" />
        Back to Dashboard Overview
      </button>

      <div className="flex flex-col md:flex-row justify-between items-start md:items-end gap-4 border-b border-[#edeef0] pb-6">
        <div>
          <h2 className="text-3xl font-bold text-gray-900">{site.name}</h2>
          <p className="text-gray-500 text-sm mt-1">{site.address}</p>
        </div>
        <div className="flex gap-4">
          <div className="text-right">
            <p className="text-[10px] font-bold text-gray-400 uppercase tracking-wider">Fortnight Hours</p>
            <p className={`text-2xl font-bold ${hourVariance > 0 ? 'text-red-600' : 'text-green-600'}`}>{totalActualHours.toFixed(1)}h</p>
          </div>
          <div className="text-right">
            <p className="text-[10px] font-bold text-gray-400 uppercase tracking-wider">Hourly Target</p>
            <p className="text-2xl font-bold text-gray-900">{budgetedHoursPerFortnight}h</p>
          </div>
        </div>
      </div>

      {/* Admin Financial Audit Section */}
      {isAdmin && (
        <div className="bg-[#fcfcfb] border border-[#edeef0] rounded-xl p-8 space-y-6">
          <div className="flex items-center gap-2 mb-2">
            <Wallet className="text-gray-400" size={20} />
            <h3 className="text-lg font-bold text-gray-900">Financial Performance Audit</h3>
          </div>
          <div className="grid grid-cols-1 md:grid-cols-4 gap-6">
            <div>
              <p className="text-[10px] font-bold text-gray-400 uppercase mb-1">Contract Revenue</p>
              <h4 className="text-xl font-bold text-gray-900">${financialStats.fortnightRevenue.toLocaleString(undefined, { minimumFractionDigits: 2 })}</h4>
              <p className="text-[10px] text-gray-400 mt-1">(${((site as any).monthly_revenue ?? (site as any).monthlyRevenue ?? 0).toLocaleString()}/mo)</p>
            </div>
            <div>
              <p className="text-[10px] font-bold text-gray-400 uppercase mb-1">Actual Labor Cost</p>
              <h4 className="text-xl font-bold text-red-600">-${financialStats.actualLaborCost.toLocaleString(undefined, { minimumFractionDigits: 2 })}</h4>
              <p className="text-[10px] text-gray-400 mt-1">Direct wages only</p>
            </div>
            <div>
              <p className="text-[10px] font-bold text-gray-400 uppercase mb-1">Gross Profit</p>
              <h4 className="text-xl font-bold text-green-600">${financialStats.grossProfit.toLocaleString(undefined, { minimumFractionDigits: 2 })}</h4>
              <p className="text-[10px] text-gray-400 mt-1">Before overheads</p>
            </div>
            <div className="bg-white p-4 rounded-lg border border-[#edeef0]">
              <p className="text-[10px] font-bold text-gray-400 uppercase mb-1">Profit Margin</p>
              <div className="flex items-center gap-2">
                <h4 className={`text-2xl font-bold ${financialStats.margin > 20 ? 'text-blue-600' : 'text-red-600'}`}>
                  {financialStats.margin.toFixed(1)}%
                </h4>
                <Percent size={16} className="text-gray-300" />
              </div>
            </div>
          </div>
        </div>
      )}

      <div className="space-y-6">
        <div className="flex items-center gap-2">
          <CalendarDays className="text-gray-400" size={20} />
          <h3 className="text-lg font-bold text-gray-900">Operational Variance Audit</h3>
        </div>

        <div className="grid grid-cols-2 sm:grid-cols-7 gap-2">
          {dailyStats.map((stat) => (
            <div key={stat.date} className="p-3 rounded-md border bg-white border-[#edeef0]">
              <p className="text-[9px] font-bold text-gray-400 uppercase mb-1">{stat.displayDate}</p>
              <p className={`text-lg font-bold ${stat.actual > stat.budget ? 'text-red-600' : 'text-gray-900'}`}>{stat.actual}h</p>
              <p className="text-[9px] text-gray-400 font-medium">Plan: {stat.budget}h</p>
            </div>
          ))}
        </div>

        <div className="h-[300px] w-full border border-[#edeef0] rounded-lg p-6 bg-white shadow-sm">
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
