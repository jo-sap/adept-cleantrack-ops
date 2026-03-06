
import React, { useState, useMemo, useEffect, useRef } from 'react';
import { Site, Cleaner, TimeEntry, FortnightPeriod } from '../types';
import { format, getDay, addDays } from 'date-fns';
import { useRole } from '../contexts/RoleContext';
import { getDayStatus } from '../utils';
import { 
  Save, Zap, Building, ArrowLeft, UserPlus, TrendingUp, TrendingDown, Check, Download, FileSpreadsheet, History
} from 'lucide-react';
import { exportFortnightTimesheets } from '../services/exportService';
import { supabase } from '../lib/supabase';

interface TimeEntryFormProps {
  sites: Site[];
  cleaners: Cleaner[];
  entries: TimeEntry[];
  currentPeriod: FortnightPeriod;
  onSaveBatch: (entries: Omit<TimeEntry, 'id'>[]) => void;
  onDeleteEntry: (id: string) => void;
  onUpdateSite?: (site: Site) => void;
}

const TimeEntryForm: React.FC<TimeEntryFormProps> = ({ 
  sites, cleaners, entries, currentPeriod, onSaveBatch, onUpdateSite
}) => {
  const { isAdmin, profile } = useRole();
  const [activeSiteId, setActiveSiteId] = useState<string | null>(null);
  const [selectedCleanerId, setSelectedCleanerId] = useState<string>('');
  const [draftHours, setDraftHours] = useState<Record<string, number>>({});
  const [isSaved, setIsSaved] = useState(false);
  const [hasUnsavedChanges, setHasUnsavedChanges] = useState(false);
  const [auditInfo, setAuditInfo] = useState<{name: string, time: string} | null>(null);
  const [cleanerSearch, setCleanerSearch] = useState('');
  
  const inputRefs = useRef<Record<string, HTMLInputElement | null>>({});

  const activeSite = sites.find(s => s.id === activeSiteId);
  const activeCleaner = cleaners.find(c => c.id === selectedCleanerId);

  const dates = useMemo(() => {
    const list = [];
    for (let i = 0; i < 14; i++) list.push(addDays(currentPeriod.startDate, i));
    return list;
  }, [currentPeriod]);

  const siteEntriesByDate = useMemo(() => {
    if (!activeSiteId) return {} as Record<string, Record<string, number>>;
    const map: Record<string, Record<string, number>> = {};
    entries.forEach(e => {
      if (e.siteId !== activeSiteId || !e.cleanerId) return;
      const key = e.date;
      if (!map[key]) map[key] = {};
      map[key][e.cleanerId] = (map[key][e.cleanerId] || 0) + e.hours;
    });
    return map;
  }, [entries, activeSiteId]);

  useEffect(() => {
    if (!activeSiteId || !activeSite) return;
    const personnelIds = activeSite.assigned_cleaner_ids?.length ? activeSite.assigned_cleaner_ids : cleaners.map(c => c.id);
    if (personnelIds.length === 0) return;
    const firstId = personnelIds[0];
    setSelectedCleanerId(prev => {
      if (prev && personnelIds.includes(prev)) return prev;
      return firstId;
    });
  }, [activeSiteId, activeSite, cleaners]);

  useEffect(() => {
    if (activeSiteId && selectedCleanerId) {
      const newDraft: Record<string, number> = {};
      dates.forEach(date => {
        const dateStr = format(date, 'yyyy-MM-dd');
        const existing = entries.find(e => e.siteId === activeSiteId && e.cleanerId === selectedCleanerId && e.date === dateStr);
        newDraft[dateStr] = existing ? existing.hours : 0;
      });
      setDraftHours(newDraft);
      setHasUnsavedChanges(false);
      fetchAuditTrail();
    }
  }, [activeSiteId, selectedCleanerId, currentPeriod.id]);

  const fetchAuditTrail = async () => {
    const { data } = await supabase.from('timesheet_batches').select('updated_at, profiles(full_name)')
      .match({ 
        site_id: activeSiteId, 
        cleaner_id: selectedCleanerId, 
        fortnight_start: format(currentPeriod.startDate, 'yyyy-MM-dd')
      }).single();
    if (data) setAuditInfo({ name: (data as any).profiles.full_name, time: format(new Date(data.updated_at), 'MMM d, h:mma') });
    else setAuditInfo(null);
  };

  const handleHourChange = (dateStr: string, val: string) => {
    const hours = val === '' ? 0 : Math.round(parseFloat(val) * 10) / 10;
    setDraftHours(prev => ({ ...prev, [dateStr]: hours }));
    setHasUnsavedChanges(true);
  };

  const handleSave = async () => {
    if (!activeSiteId || !selectedCleanerId || !activeCleaner || !activeSite) return;
    const batchData = (Object.entries(draftHours) as [string, number][]).map(([date, hours]) => ({
      siteId: activeSiteId,
      cleanerId: selectedCleanerId,
      date,
      hours,
      pay_rate_snapshot: activeSite.cleaner_rates[selectedCleanerId] || activeCleaner.payRatePerHour || 0
    }));
    await onSaveBatch(batchData as any);
    setHasUnsavedChanges(false);
    setIsSaved(true);
    fetchAuditTrail();
    setTimeout(() => setIsSaved(false), 2000);
  };

  const handleAutoFill = () => {
    if (!activeSite) return;
    const isPeriodBudget = activeSite.visit_frequency === "Monthly" || activeSite.visit_frequency === "Fortnightly";
    const periodCap = activeSite.budgeted_hours_per_fortnight ?? 0;
    const dailySum = dates.reduce((s, d) => s + (activeSite.daily_budgets[getDay(d)] || 0), 0);
    const usePeriodCap = isPeriodBudget && periodCap > 0 && dailySum === 0;

    const next: Record<string, number> = { ...draftHours };
    let changed = false;
    dates.forEach(date => {
      const dateStr = format(date, "yyyy-MM-dd");
      const current = next[dateStr] ?? 0;
      let planned: number;
      if (usePeriodCap) {
        planned = 0;
      } else {
        planned = activeSite.daily_budgets[getDay(date)] ?? 0;
      }
      if (planned > 0 && current === 0) {
        next[dateStr] = planned;
        changed = true;
      }
    });
    if (usePeriodCap && periodCap > 0) {
      const firstDateStr = format(dates[0], "yyyy-MM-dd");
      const currentTotal = dates.reduce((s, d) => s + (next[format(d, "yyyy-MM-dd")] ?? 0), 0);
      if (currentTotal === 0) {
        next[firstDateStr] = periodCap;
        changed = true;
      }
    }
    if (changed) {
      setDraftHours(next);
      setHasUnsavedChanges(true);
    }
  };

  const summary = useMemo(() => {
    if (!activeSite || !selectedCleanerId) return null;
    const isPeriodBudget =
      activeSite.visit_frequency === "Monthly" || activeSite.visit_frequency === "Fortnightly";
    const periodCap = activeSite.budgeted_hours_per_fortnight ?? 0;
    const dailyPlanSum = dates.reduce(
      (s, d) => s + (activeSite.daily_budgets[getDay(d)] || 0),
      0
    );
    const usePeriodCap = isPeriodBudget && periodCap > 0 && dailyPlanSum === 0;
    const budgetTotal = usePeriodCap
      ? periodCap
      : dates.reduce((s, date) => s + (activeSite.daily_budgets[getDay(date)] || 0), 0);

    let siteActualTotal = 0;
    let cleanerActualTotal = 0;
    let unplannedCount = 0;
    let missingCount = 0;

    dates.forEach((date) => {
      const dateStr = format(date, "yyyy-MM-dd");
      const sitePlanForDay = activeSite.daily_budgets[getDay(date)] || 0;
      const dayAssignments = siteEntriesByDate[dateStr] || {};
      const existingForCleaner = dayAssignments[selectedCleanerId] || 0;
      const existingDayTotal = Object.values(dayAssignments).reduce(
        (sum, h) => sum + h,
        0
      );
      const draftForCleaner =
        draftHours[dateStr] !== undefined ? draftHours[dateStr] : existingForCleaner;

      // Site-level actual after applying this draft
      const dayTotal = existingDayTotal - existingForCleaner + draftForCleaner;
      siteActualTotal += dayTotal;
      cleanerActualTotal += draftForCleaner;

      const planForDay = usePeriodCap ? 0 : sitePlanForDay;
      if (!usePeriodCap) {
        if (dayTotal > 0 && planForDay === 0) unplannedCount++;
        if (planForDay > 0 && dayTotal === 0) missingCount++;
      }
    });

    const rate =
      activeSite.cleaner_rates[selectedCleanerId] || (activeCleaner?.payRatePerHour || 0);
    return {
      budgetTotal,
      actualTotal: cleanerActualTotal,
      variance: siteActualTotal - budgetTotal,
      estPay: cleanerActualTotal * rate,
      unplannedCount,
      missingCount,
      isPeriodBudget: usePeriodCap,
    };
  }, [activeSite, selectedCleanerId, activeCleaner, dates, draftHours, siteEntriesByDate]);

  if (activeSiteId && activeSite) {
    return (
      <div className="space-y-6 animate-fadeIn pb-12">
        <div className="flex justify-between items-start">
          <div className="flex flex-col gap-3">
            <button onClick={() => setActiveSiteId(null)} className="flex items-center gap-2 text-xs text-gray-400 hover:text-gray-900 group w-fit"><ArrowLeft size={14} /> Back</button>
            <div className="flex items-center gap-3">
               <div className="bg-gray-100 p-2 rounded-lg"><Building className="text-gray-500" size={18} /></div>
               <div>
                  <h2 className="text-xl font-bold text-gray-900">{activeSite.name}</h2>
                  <p className="text-gray-500 text-[10px] font-medium uppercase tracking-widest">{activeSite.address}</p>
               </div>
            </div>
          </div>
          {auditInfo && (
            <div className="bg-gray-50 px-4 py-2 rounded-xl border border-gray-100 flex items-center gap-2">
              <History size={14} className="text-gray-400" />
              <div className="text-right">
                <p className="text-[9px] font-bold text-gray-400 uppercase leading-none mb-1">Last Edited</p>
                <p className="text-[10px] font-bold text-gray-700 leading-none">{auditInfo.name} • {auditInfo.time}</p>
              </div>
            </div>
          )}
        </div>

        {(() => {
          const personnelIds = activeSite.assigned_cleaner_ids?.length ? activeSite.assigned_cleaner_ids : cleaners.map(c => c.id);
          if (personnelIds.length === 0) {
            return <div className="py-2"><p className="text-xs text-gray-400">No cleaners in team. Add cleaners in Cleaner Team first.</p></div>;
          }
          const filter = cleanerSearch.trim().toLowerCase();
          const options = personnelIds
            .map(cid => cleaners.find(c => c.id === cid))
            .filter((c): c is Cleaner => !!c)
            .filter(c => {
              if (!filter) return true;
              const name = `${c.firstName} ${c.lastName}`.toLowerCase();
              return name.includes(filter);
            });
          return (
            <div className="flex flex-col md:flex-row md:items-end gap-3 border-b border-[#edeef0] pt-2 pb-2">
              <div className="flex-1">
                <label className="block text-[10px] font-bold text-gray-400 uppercase mb-1">
                  Personnel
                </label>
                <select
                  value={selectedCleanerId}
                  onChange={(e) => setSelectedCleanerId(e.target.value)}
                  className="w-full border border-[#edeef0] rounded-lg px-3 py-2 text-sm"
                >
                  {!selectedCleanerId && <option value="">Select cleaner…</option>}
                  {options.map(c => (
                    <option key={c.id} value={c.id}>
                      {c.firstName} {c.lastName}
                    </option>
                  ))}
                </select>
              </div>
              <div className="flex-1 max-w-xs">
                <label className="block text-[10px] font-bold text-gray-400 uppercase mb-1">
                  Search team
                </label>
                <input
                  type="text"
                  value={cleanerSearch}
                  onChange={(e) => setCleanerSearch(e.target.value)}
                  placeholder="Type name to filter…"
                  className="w-full border border-[#edeef0] rounded-lg px-3 py-2 text-sm"
                />
              </div>
            </div>
          );
        })()}

        {selectedCleanerId ? (
          <div className="space-y-6">
            <div className="sticky top-0 z-30 bg-white border border-[#edeef0] rounded-xl px-4 py-3 shadow-sm flex items-center justify-between h-20">
              <div className="grid grid-cols-4 divide-x divide-gray-100 flex-1">
                {['Budget', 'Actual', 'Variance', 'Est. Pay'].map((label, idx) => (
                  <div key={label} className="px-6 flex flex-col justify-center">
                    <span className="text-[10px] font-bold text-gray-400 uppercase tracking-widest mb-1.5">{label}</span>
                    <span className={`text-sm font-bold ${label === 'Variance' && summary && summary.variance > 0.1 ? 'text-red-600' : 'text-gray-900'}`}>
                      {idx === 0 ? summary?.budgetTotal.toFixed(1) + 'h' : idx === 1 ? summary?.actualTotal.toFixed(1) + 'h' : idx === 2 ? (summary!.variance > 0 ? '+' : '') + summary?.variance.toFixed(1) + 'h' : '$' + summary?.estPay.toLocaleString()}
                    </span>
                  </div>
                ))}
              </div>
              <div className="flex items-center gap-3 pr-2">
                <button type="button" onClick={handleAutoFill} className="flex flex-col items-center justify-center w-24 h-14 rounded-xl border border-[#edeef0] bg-white text-gray-700 hover:bg-gray-50 hover:border-gray-300 transition-all" title="Fill missing hours with planned values">
                  <Zap size={12} className="opacity-70 mb-0.5" /><span className="text-[10px] font-bold uppercase">Auto fill</span>
                </button>
                <button onClick={handleSave} disabled={!hasUnsavedChanges} className={`flex flex-col items-center justify-center w-24 h-14 rounded-xl transition-all ${isSaved ? 'bg-green-600 text-white' : 'bg-gray-900 text-white hover:bg-black disabled:opacity-40'}`}>
                   <Save size={12} className="opacity-70 mb-0.5" /><span className="text-[10px] font-bold uppercase">Save Batch</span>
                </button>
              </div>
            </div>
            {activeSite.visit_frequency === 'Monthly' && activeSite.budgeted_hours_per_fortnight > 0 && (
              <p className="text-xs text-gray-500 bg-gray-50 border border-gray-100 rounded-lg px-3 py-2">
                Monthly job – <strong>{activeSite.budgeted_hours_per_fortnight.toFixed(1)}h</strong> planned this period. Enter actual hours on the day(s) they worked and save.
              </p>
            )}
            <div className="grid grid-cols-7 gap-2">
              {dates.map((date, index) => {
                const dateStr = format(date, 'yyyy-MM-dd');
                const sitePlan = activeSite.daily_budgets[getDay(date)] || 0;
                const dayAssignments = siteEntriesByDate[dateStr] || {};
                const otherHours = Object.entries(dayAssignments)
                  .filter(([cid]) => cid !== selectedCleanerId)
                  .reduce((sum, [, h]) => sum + h, 0);
                const remainingPlan = Math.max(sitePlan - otherHours, 0);
                const planned = remainingPlan;
                const hours = draftHours[dateStr] || 0;
                const status = getDayStatus(planned, hours);
                const existingForThisCleaner = dayAssignments[selectedCleanerId] || 0;
                const fullyAllocatedToOthers =
                  sitePlan > 0 && remainingPlan <= 0 && existingForThisCleaner === 0;
                return (
                  <div key={dateStr} onClick={() => inputRefs.current[dateStr]?.focus()} className={`flex flex-col p-3 bg-white border rounded-xl cursor-pointer group relative overflow-hidden ${status.border} ${status.bg.replace('bg-', 'hover:bg-')}`}>
                    <div className={`absolute top-2 right-2 w-1.5 h-1.5 rounded-full ${status.dot}`} />
                    <div className="mb-2"><p className="text-[10px] font-bold text-gray-900 uppercase">{format(date, 'EEE')}</p><p className="text-[9px] font-medium text-gray-400">{format(date, 'MMM d')}</p></div>
                    <div className="mt-auto space-y-1.5">
                      <div className="flex justify-between items-center"><span className="text-[8px] font-bold text-gray-400">Plan</span><span className="text-[9px] font-bold text-gray-800">{planned.toFixed(1)}h</span></div>
                      <input
                        ref={el => inputRefs.current[dateStr] = el}
                        type="number"
                        step="0.1"
                        value={draftHours[dateStr] || ''}
                        onChange={e => handleHourChange(dateStr, e.target.value)}
                        className="w-full px-2 py-1.5 text-center text-xs font-bold bg-gray-50/50 border border-transparent rounded-lg focus:ring-1 focus:ring-gray-900 outline-none disabled:opacity-40 disabled:cursor-not-allowed"
                        disabled={fullyAllocatedToOthers}
                      />
                      <div className={`mt-1 text-center py-0.5 rounded-full text-[8px] font-black uppercase ${status.bg} ${status.color} border ${status.border}`}>{status.label}</div>
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        ) : <div className="py-20 text-center border-dashed border-2 border-gray-100 rounded-2xl text-gray-400 text-xs">Select personnel to audit.</div>}
      </div>
    );
  }

  return (
    <div className="space-y-10">
      <div className="flex items-center justify-between px-2">
        <h3 className="text-lg font-bold text-gray-900">Portfolio Audit Board</h3>
        <div className="flex gap-2">
          <button onClick={() => exportFortnightTimesheets(currentPeriod, sites, cleaners, entries as any, 'xlsx')} className="flex items-center gap-1.5 px-4 py-1.5 bg-green-50 text-green-700 border border-green-200 rounded-lg text-xs font-bold">Export (XLSX)</button>
        </div>
      </div>
      <div className="space-y-3">
        {sites.map(site => (
          <div key={site.id} onClick={() => setActiveSiteId(site.id)} className="group bg-white border border-[#edeef0] hover:border-gray-900 rounded-xl p-4 cursor-pointer flex items-center justify-between">
            <div className="flex items-center gap-4 w-1/3">
              <div className="w-10 h-10 bg-gray-50 group-hover:bg-white border border-[#edeef0] rounded-xl flex items-center justify-center text-gray-400"><Building size={20} /></div>
              <div className="min-w-0"><h4 className="text-sm font-bold text-gray-900 truncate">{site.name}</h4><p className="text-[10px] text-gray-400 truncate uppercase font-bold">{site.address}</p></div>
            </div>
            <div className="flex -space-x-2 w-32">
              {site.assigned_cleaner_ids.slice(0, 3).map(cid => <div key={cid} className="w-8 h-8 rounded-xl border-2 border-white bg-gray-100 flex items-center justify-center text-[10px] font-bold text-gray-500 shadow-sm">{cleaners.find(c => c.id === cid)?.firstName.charAt(0)}</div>)}
            </div>
            <div className="text-right border-l border-gray-100 pl-8">
              <p className="text-[9px] font-bold text-gray-400 uppercase">Budget</p>
              <p className="text-xs font-black text-gray-800">{site.budgeted_hours_per_fortnight.toFixed(0)}h</p>
            </div>
          </div>
        ))}
      </div>
    </div>
  );
};

export default TimeEntryForm;
