
import React, { useState, useMemo, useEffect, useRef } from 'react';
import { Site, Cleaner, TimeEntry, FortnightPeriod, AdHocJob } from '../types';
import { format, getDay, addDays } from 'date-fns';
import { useRole } from '../contexts/RoleContext';
import { getDayStatus } from '../utils';
import { getSiteRateForDate, computeBudgetedLabourCostForRange } from '../lib/budgetedLabourCost';
import { getPublicHolidaysInRange } from '../lib/publicHolidays';
import { Save, Zap, Building, ArrowLeft, UserPlus, TrendingUp, TrendingDown, Check, Download, FileSpreadsheet, History, Search, Loader2 } from 'lucide-react';
import { exportFortnightTimesheets } from '../services/exportService';
import { supabase } from '../lib/supabase';
import { getGraphAccessToken } from '../lib/graph';
import { getAdHocJobs } from '../repositories/adHocJobsRepo';
import { generateAdHocOccurrencesForRange, occurrencesToHoursByDate } from '../lib/adhocSchedule';

interface TimeEntryFormProps {
  sites: Site[];
  cleaners: Cleaner[];
  entries: TimeEntry[];
  currentPeriod: FortnightPeriod;
  onSaveBatch: (entries: Omit<TimeEntry, 'id'>[]) => void;
  onDeleteEntry: (id: string) => void;
  onUpdateSite?: (site: Site) => void;
}

type QueueKey = 'all' | 'needs-hours' | 'incomplete' | 'over-budget' | 'completed';

function normalizeScheduleType(raw: string | undefined | null): 'once_off' | 'recurring' {
  const s = String(raw ?? '').trim().toLowerCase();
  if (!s) return 'once_off';
  if (s.includes('recurr')) return 'recurring';
  return 'once_off';
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
  const [siteSearchQuery, setSiteSearchQuery] = useState('');
  const [saveLoading, setSaveLoading] = useState(false);
  const [adhocJobId, setAdhocJobId] = useState<string | null>(null);
  const [adHocJobsForSite, setAdHocJobsForSite] = useState<AdHocJob[]>([]);
  const [queueKey, setQueueKey] = useState<QueueKey>('all');
  
  const inputRefs = useRef<Record<string, HTMLInputElement | null>>({});

  const activeSite = sites.find(s => s.id === activeSiteId);
  const activeCleaner = cleaners.find(c => c.id === selectedCleanerId);

  const filteredSites = useMemo(() => {
    const q = siteSearchQuery.trim().toLowerCase();
    if (!q) return sites;
    return sites.filter(
      s =>
        (s.name?.toLowerCase().includes(q) ?? false) ||
        (s.address?.toLowerCase().includes(q) ?? false)
    );
  }, [sites, siteSearchQuery]);

  const dates = useMemo(() => {
    const list = [];
    for (let i = 0; i < 14; i++) list.push(addDays(currentPeriod.startDate, i));
    return list;
  }, [currentPeriod]);

  const siteSummaryById = useMemo(() => {
    const start = currentPeriod.startDate;
    // Use an exclusive end boundary to avoid timezone/midnight edge cases when parsing
    // ISO dates like "2026-03-22" (which can be interpreted as UTC and fall outside <= end).
    const endExclusive = addDays(currentPeriod.endDate, 1);
    const map: Record<
      string,
      { budgetTotal: number; actualTotal: number; variance: number }
    > = {};
    sites.forEach((site) => {
      const budgetTotal = site.budgeted_hours_per_fortnight ?? 0;
      let actualTotal = 0;
      entries.forEach((e) => {
        if (e.siteId !== site.id) return;
        // Contract totals only – Ad Hoc entries must not distort budget/variance.
        if (e.adhocJobId) return;
        const d = new Date(e.date);
        if (d >= start && d < endExclusive) {
          actualTotal += e.hours;
        }
      });
      const variance = actualTotal - budgetTotal;
      map[site.id] = { budgetTotal, actualTotal, variance };
    });
    return map;
  }, [sites, entries, currentPeriod]);

  const phInPeriod = useMemo(
    () => getPublicHolidaysInRange(currentPeriod.startDate, currentPeriod.endDate),
    [currentPeriod]
  );

  const activeAdHocJob = useMemo(
    () => (adhocJobId ? adHocJobsForSite.find((j) => j.id === adhocJobId) : undefined),
    [adhocJobId, adHocJobsForSite]
  );

  const activeAdHocOccurrences = useMemo(() => {
    if (!activeAdHocJob) return [];
    return generateAdHocOccurrencesForRange(
      activeAdHocJob,
      currentPeriod.startDate,
      currentPeriod.endDate,
      phInPeriod
    );
  }, [activeAdHocJob, currentPeriod.startDate, currentPeriod.endDate, phInPeriod]);

  const activeAdHocPlannedByDate = useMemo(() => {
    return occurrencesToHoursByDate(activeAdHocOccurrences);
  }, [activeAdHocOccurrences]);

  const activeAdHocTotals = useMemo(() => {
    return activeAdHocOccurrences.reduce(
      (acc, o) => {
        acc.hours += o.hours;
        acc.charge += o.chargeTotal;
        acc.cost += o.costTotal;
        return acc;
      },
      { hours: 0, charge: 0, cost: 0 }
    );
  }, [activeAdHocOccurrences]);

  const estimatedBudgetBySiteId = useMemo(() => {
    const map: Record<string, number> = {};
    const weekdayDefault = 0;
    sites.forEach((site) => {
      const weekdayRate = site.budget_weekday_labour_rate ?? site.budget_labour_rate ?? weekdayDefault;
      const saturdayRate = site.budget_saturday_labour_rate ?? weekdayRate;
      const sundayRate = site.budget_sunday_labour_rate ?? weekdayRate;
      const phRate = site.budget_ph_labour_rate ?? weekdayRate;
      const dailyBudgets = (site.daily_budgets?.length ?? 0) >= 7
        ? site.daily_budgets
        : [0, 0, 0, 0, 0, 0, 0];
      const cost = computeBudgetedLabourCostForRange({
        startDate: currentPeriod.startDate,
        endDate: currentPeriod.endDate,
        dailyBudgets,
        weekdayRate,
        saturdayRate,
        sundayRate,
        phRate,
        publicHolidayDates: phInPeriod,
      });
      map[site.id] = cost;
    });
    return map;
  }, [sites, currentPeriod, phInPeriod]);

  const totalEstimatedBudget = useMemo(() => {
    return filteredSites.reduce(
      (sum, site) => sum + (estimatedBudgetBySiteId[site.id] ?? 0),
      0
    );
  }, [filteredSites, estimatedBudgetBySiteId]);

  const queueCounts = useMemo(() => {
    let needs = 0;
    let incomplete = 0;
    let over = 0;
    let completed = 0;
    filteredSites.forEach((site) => {
      const summary = siteSummaryById[site.id];
      if (!summary) return;
      const logged = summary.actualTotal;
      const budget = summary.budgetTotal;
      if (logged === 0) {
        needs += 1;
      } else if (logged > 0 && logged < budget) {
        incomplete += 1;
      } else if (logged > budget) {
        over += 1;
      } else if (logged === budget) {
        completed += 1;
      }
    });
    return {
      all: filteredSites.length,
      needs,
      incomplete,
      over,
      completed,
    };
  }, [filteredSites, siteSummaryById]);

  const orderedSites = useMemo(() => {
    if (queueKey === 'all') return filteredSites;
    const classify = (siteId: string): QueueKey => {
      const summary = siteSummaryById[siteId];
      if (!summary) return 'all';
      const logged = summary.actualTotal;
      const budget = summary.budgetTotal;
      if (logged === 0) return 'needs-hours';
      if (logged > 0 && logged < budget) return 'incomplete';
      if (logged > budget) return 'over-budget';
      if (logged === budget) return 'completed';
      return 'all';
    };
    return [...filteredSites].sort((a, b) => {
      const aMatch = classify(a.id) === queueKey ? 0 : 1;
      const bMatch = classify(b.id) === queueKey ? 0 : 1;
      return aMatch - bMatch;
    });
  }, [filteredSites, queueKey, siteSummaryById]);

  const siteEntriesByDate = useMemo(() => {
    if (!activeSiteId) return {} as Record<string, Record<string, number>>;
    const map: Record<string, Record<string, number>> = {};
    entries.forEach(e => {
      if (e.siteId !== activeSiteId || !e.cleanerId) return;
      // This map is used for contract plan/variance logic – exclude Ad Hoc entries.
      if (e.adhocJobId) return;
      const key = e.date;
      if (!map[key]) map[key] = {};
      map[key][e.cleanerId] = (map[key][e.cleanerId] || 0) + e.hours;
    });
    return map;
  }, [entries, activeSiteId]);

  const adHocEntriesByDate = useMemo(() => {
    if (!activeSiteId || !adhocJobId) return {} as Record<string, Record<string, number>>;
    const map: Record<string, Record<string, number>> = {};
    entries.forEach((e) => {
      if (e.siteId !== activeSiteId || !e.cleanerId) return;
      if (!e.adhocJobId || e.adhocJobId !== adhocJobId) return;
      const key = e.date;
      if (!map[key]) map[key] = {};
      map[key][e.cleanerId] = (map[key][e.cleanerId] || 0) + e.hours;
    });
    return map;
  }, [entries, activeSiteId, adhocJobId]);

  useEffect(() => {
    if (!activeSiteId || !activeSite) {
      setSelectedCleanerId('');
      return;
    }
    const personnelIds = activeSite.assigned_cleaner_ids ?? [];
    if (personnelIds.length === 0) {
      setSelectedCleanerId('');
      return;
    }
    const firstId = personnelIds[0];
    setSelectedCleanerId(prev => {
      if (prev && personnelIds.includes(prev)) return prev;
      return firstId;
    });
  }, [activeSiteId, activeSite]);

  // When changing site/cleaner, default to contract entries when available; otherwise pick the first ad hoc job used.
  useEffect(() => {
    if (!activeSiteId || !selectedCleanerId) return;
    const dateKeys = new Set(dates.map((d) => format(d, 'yyyy-MM-dd')));
    const inPeriod = entries.filter(
      (e) =>
        e.siteId === activeSiteId &&
        e.cleanerId === selectedCleanerId &&
        dateKeys.has(e.date)
    );
    const contract = inPeriod.find((e) => !e.adhocJobId);
    if (contract) {
      setAdhocJobId(null);
      return;
    }
    const firstAdhoc = inPeriod.find((e) => !!e.adhocJobId);
    setAdhocJobId(firstAdhoc?.adhocJobId ?? null);
  }, [activeSiteId, selectedCleanerId, currentPeriod.id, entries, dates]);

  // Load draft hours for the current (site, cleaner, ad hoc context) selection.
  useEffect(() => {
    if (!activeSiteId || !selectedCleanerId) return;
    const newDraft: Record<string, number> = {};
    dates.forEach((date) => {
      const dateStr = format(date, 'yyyy-MM-dd');
      const existing = entries.find(
        (e) =>
          e.siteId === activeSiteId &&
          e.cleanerId === selectedCleanerId &&
          e.date === dateStr &&
          (adhocJobId ? e.adhocJobId === adhocJobId : !e.adhocJobId)
      );
      newDraft[dateStr] = existing ? existing.hours : 0;
    });
    setDraftHours(newDraft);
    setHasUnsavedChanges(false);
    fetchAuditTrail();
  }, [activeSiteId, selectedCleanerId, adhocJobId, currentPeriod.id, entries, dates]);

  useEffect(() => {
    if (!activeSiteId) {
      setAdHocJobsForSite([]);
      return;
    }
    let cancelled = false;
    getGraphAccessToken().then(async (token) => {
      if (!token || cancelled) return;
      try {
        const list = await getAdHocJobs(token, { siteId: activeSiteId });
        const eligible = list.filter((j) => {
          if (!j.active) return false;
          const status = String(j.status ?? "").trim().toLowerCase();
          if (status === "cancelled") return false;
          // Show if there is at least one generated occurrence in this selected period.
          const occ = generateAdHocOccurrencesForRange(j, currentPeriod.startDate, currentPeriod.endDate);
          return occ.length > 0;
        });

        if (!cancelled) setAdHocJobsForSite(eligible);
      } catch {
        if (!cancelled) setAdHocJobsForSite([]);
      }
    });
    return () => { cancelled = true; };
  }, [activeSiteId, currentPeriod.startDate.getTime(), currentPeriod.endDate.getTime()]);

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
    setSaveLoading(true);
    try {
      const batchData = (Object.entries(draftHours) as [string, number][]).map(([date, hours]) => ({
        siteId: activeSiteId,
        cleanerId: selectedCleanerId,
        date,
        hours,
        pay_rate_snapshot: activeSite.cleaner_rates[selectedCleanerId] || activeCleaner.payRatePerHour || 0,
        adhocJobId: adhocJobId || undefined
      }));
      await onSaveBatch(batchData as any);
      setHasUnsavedChanges(false);
      setIsSaved(true);
      await fetchAuditTrail();
      setTimeout(() => setIsSaved(false), 2000);
    } finally {
      setSaveLoading(false);
    }
  };

  const handleAutoFill = () => {
    if (!activeSite || !selectedCleanerId) return;
    if (adhocJobId) return;
    const isPeriodBudget = activeSite.visit_frequency === "Monthly" || activeSite.visit_frequency === "Fortnightly";
    const periodCap = activeSite.budgeted_hours_per_fortnight ?? 0;
    const dailySum = dates.reduce((s, d) => s + (activeSite.daily_budgets[getDay(d)] || 0), 0);
    const usePeriodCap = isPeriodBudget && periodCap > 0 && dailySum === 0;

    const next: Record<string, number> = { ...draftHours };
    let changed = false;
    dates.forEach(date => {
      const dateStr = format(date, "yyyy-MM-dd");
      const current = next[dateStr] ?? 0;
      const sitePlan = activeSite.daily_budgets[getDay(date)] ?? 0;
      // Remaining plan for this cleaner = site plan minus other cleaners' hours (same as "Plan" on the card)
      const dayAssignments = siteEntriesByDate[dateStr] || {};
      const otherHours = Object.entries(dayAssignments)
        .filter(([cid]) => cid !== selectedCleanerId)
        .reduce((sum, [, h]) => sum + h, 0);
      const remainingPlan = Math.max(sitePlan - otherHours, 0);

      let fillValue: number;
      if (usePeriodCap) {
        fillValue = 0;
      } else {
        fillValue = remainingPlan;
      }
      // Only fill days that have remaining plan and currently no hours (fill with "hours in plan")
      if (fillValue > 0 && current === 0) {
        next[dateStr] = fillValue;
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
    const isAdHoc = !!adhocJobId;
    const entriesByDate = isAdHoc ? adHocEntriesByDate : siteEntriesByDate;

    if (isAdHoc) {
      const budgetTotal = activeAdHocTotals.hours;
      let cleanerActualTotal = 0;
      let estPay = 0;

      // Site labour rates (Weekday, Sat, Sun, PH) — Est. Pay uses these, not the cleaner's personal rate
      const weekdayRate = activeSite.budget_weekday_labour_rate ?? activeSite.budget_labour_rate ?? 0;
      const saturdayRate = activeSite.budget_saturday_labour_rate ?? weekdayRate;
      const sundayRate = activeSite.budget_sunday_labour_rate ?? weekdayRate;
      const phRate = activeSite.budget_ph_labour_rate ?? weekdayRate;
      const phInPeriod = getPublicHolidaysInRange(currentPeriod.startDate, currentPeriod.endDate);

      dates.forEach((date) => {
        const dateStr = format(date, "yyyy-MM-dd");
        const dayAssignments = entriesByDate[dateStr] || {};
        const existingForCleaner = dayAssignments[selectedCleanerId] || 0;
        const draftForCleaner =
          draftHours[dateStr] !== undefined ? draftHours[dateStr] : existingForCleaner;
        cleanerActualTotal += draftForCleaner;
        if (draftForCleaner > 0) {
          const rate = getSiteRateForDate(date, weekdayRate, saturdayRate, sundayRate, phRate, phInPeriod);
          estPay += draftForCleaner * rate;
        }
      });

      return {
        budgetTotal,
        budgetDisplay: budgetTotal,
        actualTotal: cleanerActualTotal,
        variance: cleanerActualTotal - budgetTotal,
        estPay,
        adhocChargeTotal: activeAdHocTotals.charge,
        adhocCostTotal: activeAdHocTotals.cost,
        unplannedCount: 0,
        missingCount: 0,
        isPeriodBudget: false,
        periodCap: 0,
      };
    }

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
    let estPay = 0;
    let unplannedCount = 0;
    let missingCount = 0;

    // Site labour rates (Weekday, Sat, Sun, PH) — Est. Pay uses these, not the cleaner's personal rate
    const weekdayRate = activeSite.budget_weekday_labour_rate ?? activeSite.budget_labour_rate ?? 0;
    const saturdayRate = activeSite.budget_saturday_labour_rate ?? weekdayRate;
    const sundayRate = activeSite.budget_sunday_labour_rate ?? weekdayRate;
    const phRate = activeSite.budget_ph_labour_rate ?? weekdayRate;
    const phInPeriod = getPublicHolidaysInRange(currentPeriod.startDate, currentPeriod.endDate);

    dates.forEach((date) => {
      const dateStr = format(date, "yyyy-MM-dd");
      const sitePlanForDay = activeSite.daily_budgets[getDay(date)] || 0;
      const dayAssignments = entriesByDate[dateStr] || {};
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

      // Est. Pay: sum (hours × site rate for that day) — weekday / Sat / Sun / PH
      if (draftForCleaner > 0) {
        const rate = getSiteRateForDate(date, weekdayRate, saturdayRate, sundayRate, phRate, phInPeriod);
        estPay += draftForCleaner * rate;
      }

      const planForDay = usePeriodCap ? 0 : sitePlanForDay;
      if (!usePeriodCap) {
        if (dayTotal > 0 && planForDay === 0) unplannedCount++;
        if (planForDay > 0 && dayTotal === 0) missingCount++;
      }
    });

    // In period-cap mode, show 0h for this cleaner when they have no hours yet (e.g. newly assigned),
    // so we don't appear to "retain" the previous cleaner's budget. Site cap remains in variance logic.
    const budgetDisplay =
      usePeriodCap && cleanerActualTotal === 0 ? 0 : budgetTotal;

    return {
      budgetTotal,
      budgetDisplay,
      actualTotal: cleanerActualTotal,
      variance: siteActualTotal - budgetTotal,
      estPay,
      unplannedCount,
      missingCount,
      isPeriodBudget: usePeriodCap,
      periodCap,
    };
  }, [activeSite, selectedCleanerId, dates, draftHours, siteEntriesByDate, adHocEntriesByDate, adhocJobId, activeAdHocJob, currentPeriod]);

  if (activeSiteId && activeSite) {
    return (
      <div className="space-y-6 animate-fadeIn pb-12">
        <div className="flex justify-between items-start">
          <div className="flex flex-col gap-3">
            <button
              onClick={() => setActiveSiteId(null)}
              className="flex items-center gap-2 text-xs text-gray-500 hover:text-gray-900 group w-fit"
            >
              <ArrowLeft size={14} /> Back
            </button>
            <div className="flex items-center gap-3">
               <div className="bg-gray-100 p-2 rounded-xl border border-gray-200">
                 <Building className="text-gray-500" size={18} />
               </div>
               <div>
                  <h2 className="text-lg font-semibold text-gray-900">
                    {activeSite.name}
                  </h2>
                  <p className="text-gray-500 text-[11px]">
                    {activeSite.address}
                  </p>
               </div>
            </div>
          </div>
          {auditInfo && (
            <div className="bg-gray-50 px-4 py-2 rounded-xl border border-gray-200 flex items-center gap-2">
              <History size={14} className="text-gray-400" />
              <div className="text-right">
                <p className="text-[9px] font-bold text-gray-400 uppercase leading-none mb-1">Last Edited</p>
                <p className="text-[10px] font-bold text-gray-700 leading-none">{auditInfo.name} • {auditInfo.time}</p>
              </div>
            </div>
          )}
        </div>

        {(() => {
          const personnelIds = activeSite.assigned_cleaner_ids ?? [];
          if (personnelIds.length === 0) {
            return (
              <div className="py-2">
                <p className="text-xs text-gray-400">
                  No assigned cleaners for this site.
                </p>
              </div>
            );
          }
          const options = personnelIds
            .map(cid => cleaners.find(c => c.id === cid))
            .filter((c): c is Cleaner => !!c);
          console.log('[Timesheets] cleaner dropdown options', {
            siteId: activeSite.id,
            assignedCleanerIds: personnelIds,
            optionIds: options.map(o => o.id),
          });
          return (
            <div className="flex flex-col md:flex-row md:items-end gap-3 border-b border-[#edeef0] pt-2 pb-2">
              <div className="flex-1 max-w-xs">
                <label className="block text-[10px] font-bold text-gray-400 uppercase mb-1">
                  Personnel
                </label>
                <select
                  value={selectedCleanerId}
                  onChange={(e) => setSelectedCleanerId(e.target.value)}
                  className="w-full so-input bg-white px-3 py-2 text-sm"
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
                  Ad Hoc Job <span className="text-gray-400 font-normal">(optional)</span>
                </label>
                <select
                  value={adhocJobId ?? ""}
                  onChange={(e) => { setAdhocJobId(e.target.value || null); }}
                  className="w-full so-input bg-white px-3 py-2 text-sm"
                >
                  <option value="">Contract / standard work</option>
                  {adHocJobsForSite.map((j) => {
                    const schedule = normalizeScheduleType(j.jobType);
                    const tag = schedule === 'recurring' ? 'RECURRING' : 'ONCE OFF';
                    return (
                      <option key={j.id} value={j.id}>
                        {j.jobName}{j.scheduledDate ? ` • ${j.scheduledDate}` : ''} • {tag}
                      </option>
                    );
                  })}
                </select>
                {activeAdHocJob && (
                  <div className="mt-1 flex items-center gap-2">
                    <span className="inline-flex items-center px-2 py-0.5 rounded-full text-[10px] font-bold uppercase bg-gray-900 text-white">
                      ADHOC
                    </span>
                    <span className="inline-flex items-center px-2 py-0.5 rounded-full text-[10px] font-bold uppercase bg-gray-100 text-gray-700">
                      {normalizeScheduleType(activeAdHocJob.jobType) === 'recurring' ? 'RECURRING' : 'ONCE OFF'}
                    </span>
                  </div>
                )}
                {!activeAdHocJob && adHocJobsForSite.length > 0 && (
                  <div className="mt-1 space-y-1">
                    <p className="text-[11px] text-gray-500">
                      {adHocJobsForSite.length} ad hoc job{adHocJobsForSite.length === 1 ? "" : "s"} scheduled for this pay period.
                    </p>
                    <div className="flex flex-wrap gap-1.5">
                      {adHocJobsForSite.slice(0, 3).map((j) => (
                        <button
                          key={j.id}
                          type="button"
                          onClick={() => setAdhocJobId(j.id)}
                          className="px-2 py-1 rounded-full text-[10px] font-bold uppercase bg-gray-100 text-gray-700 hover:bg-gray-200"
                          title={j.description || j.jobName}
                        >
                          {normalizeScheduleType(j.jobType) === 'recurring' ? 'RECURRING' : 'ONCE OFF'} • {j.jobName}
                        </button>
                      ))}
                      {adHocJobsForSite.length > 3 && (
                        <span className="text-[10px] font-bold uppercase text-gray-400 px-2 py-1">
                          +{adHocJobsForSite.length - 3} more
                        </span>
                      )}
                    </div>
                  </div>
                )}
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
                      {idx === 0 ? (summary?.budgetDisplay ?? summary?.budgetTotal ?? 0).toFixed(1) + 'h' : idx === 1 ? summary?.actualTotal.toFixed(1) + 'h' : idx === 2 ? (summary!.variance > 0 ? '+' : '') + summary?.variance.toFixed(1) + 'h' : '$' + (summary?.estPay ?? 0).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                    </span>
                    {label === 'Budget' && summary?.isPeriodBudget && summary.periodCap > 0 && (
                      <span className="text-[9px] text-gray-400 mt-0.5">Site cap: {summary.periodCap.toFixed(1)}h</span>
                    )}
                    {label === 'Budget' && adhocJobId && activeAdHocJob && (
                      <span className="text-[9px] text-gray-400 mt-0.5">
                        {normalizeScheduleType(activeAdHocJob.jobType) === 'recurring' ? 'Recurring' : 'Once off'} • {activeAdHocOccurrences.length} occ
                      </span>
                    )}
                    {label === 'Est. Pay' && adhocJobId && summary && (summary as any).adhocChargeTotal != null && (
                      <span className="text-[9px] text-gray-400 mt-0.5">
                        Charge: {new Intl.NumberFormat('en-AU', { style: 'currency', currency: 'AUD', minimumFractionDigits: 0, maximumFractionDigits: 0 }).format((summary as any).adhocChargeTotal)} • Cost: {new Intl.NumberFormat('en-AU', { style: 'currency', currency: 'AUD', minimumFractionDigits: 0, maximumFractionDigits: 0 }).format((summary as any).adhocCostTotal)}
                      </span>
                    )}
                  </div>
                ))}
              </div>
              <div className="flex items-center gap-3 pr-2">
                <button
                  type="button"
                  onClick={handleAutoFill}
                  disabled={!!adhocJobId}
                  className="flex flex-col items-center justify-center w-24 h-14 rounded-xl border border-[#edeef0] bg-white text-gray-700 hover:bg-gray-50 hover:border-gray-300 transition-all"
                  title={adhocJobId ? "Auto fill is not available for Ad Hoc jobs" : "Fill empty days with remaining plan hours"}
                >
                  <Zap size={12} className="opacity-70 mb-0.5" /><span className="text-[10px] font-bold uppercase">Auto fill</span>
                </button>
                {adhocJobId && (
                  <button
                    type="button"
                    onClick={() => {
                      const next: Record<string, number> = { ...draftHours };
                      let changed = false;
                      for (const [dateStr, planned] of Object.entries(activeAdHocPlannedByDate)) {
                        if ((next[dateStr] ?? 0) === 0 && planned > 0) {
                          next[dateStr] = planned;
                          changed = true;
                        }
                      }
                      if (changed) {
                        setDraftHours(next);
                        setHasUnsavedChanges(true);
                      }
                    }}
                    className="flex flex-col items-center justify-center w-24 h-14 rounded-xl border border-[#edeef0] bg-white text-gray-700 hover:bg-gray-50 hover:border-gray-300 transition-all"
                    title="Fill empty days with scheduled hours"
                  >
                    <Zap size={12} className="opacity-70 mb-0.5" /><span className="text-[10px] font-bold uppercase">Fill</span>
                  </button>
                )}
                <button
                  onClick={handleSave}
                  disabled={!hasUnsavedChanges || saveLoading}
                  className={`flex flex-col items-center justify-center w-24 h-14 rounded-xl transition-all ${
                    isSaved
                      ? 'bg-green-600 text-white'
                      : 'so-btn-primary disabled:opacity-40'
                  }`}
                >
                  {saveLoading ? <Loader2 size={18} className="animate-spin mb-0.5" /> : <Save size={12} className="opacity-70 mb-0.5" />}
                  <span className="text-[10px] font-bold uppercase">{saveLoading ? 'Saving…' : 'Save Batch'}</span>
                </button>
              </div>
            </div>
            {!adhocJobId && activeSite.visit_frequency === 'Monthly' && activeSite.budgeted_hours_per_fortnight > 0 && (
              <p className="text-xs text-gray-500 bg-gray-50 border border-gray-100 rounded-lg px-3 py-2">
                Monthly job – <strong>{activeSite.budgeted_hours_per_fortnight.toFixed(1)}h</strong> planned this period. Enter actual hours on the day(s) they worked and save.
              </p>
            )}
            <div className="grid grid-cols-7 gap-2">
              {dates.map((date, index) => {
                const dateStr = format(date, 'yyyy-MM-dd');
                const isAdHoc = !!adhocJobId;
                const dayAssignments = (isAdHoc ? adHocEntriesByDate : siteEntriesByDate)[dateStr] || {};
                const sitePlan = isAdHoc ? 0 : (activeSite.daily_budgets[getDay(date)] || 0);
                const otherHours = Object.entries(dayAssignments)
                  .filter(([cid]) => cid !== selectedCleanerId)
                  .reduce((sum, [, h]) => sum + h, 0);
                const remainingPlan = Math.max(sitePlan - otherHours, 0);
                const planned = isAdHoc ? (activeAdHocPlannedByDate[dateStr] ?? 0) : remainingPlan;
                const hours = draftHours[dateStr] || 0;
                const status = isAdHoc
                  ? { border: "border-[#edeef0]", bg: "bg-gray-50", dot: "bg-gray-300", color: "text-gray-600", label: "Adhoc" }
                  : getDayStatus(planned, hours);
                const existingForThisCleaner = dayAssignments[selectedCleanerId] || 0;
                const fullyAllocatedToOthers =
                  !isAdHoc && sitePlan > 0 && remainingPlan <= 0 && existingForThisCleaner === 0;
                return (
                  <div key={dateStr} onClick={() => inputRefs.current[dateStr]?.focus()} className={`flex flex-col p-3 bg-white border rounded-xl cursor-pointer group relative overflow-hidden ${status.border} ${status.bg.replace('bg-', 'hover:bg-')}`}>
                    <div className={`absolute top-2 right-2 w-1.5 h-1.5 rounded-full ${status.dot}`} />
                    <div className="mb-2"><p className="text-[10px] font-bold text-gray-900 uppercase">{format(date, 'EEE')}</p><p className="text-[9px] font-medium text-gray-400">{format(date, 'MMM d')}</p></div>
                    <div className="mt-auto space-y-2">
                      <div className="flex justify-between items-center"><span className="text-[8px] font-bold text-gray-400">{isAdHoc ? "Scheduled" : "Plan"}</span><span className="text-[9px] font-bold text-gray-800">{planned.toFixed(1)}h</span></div>
                      <input
                        ref={el => inputRefs.current[dateStr] = el}
                        type="number"
                        step="0.1"
                        value={draftHours[dateStr] || ''}
                        onChange={e => handleHourChange(dateStr, e.target.value)}
                        className="w-full px-2.5 py-2 text-center text-sm font-semibold bg-white border border-gray-300 rounded-lg focus:ring-2 focus:ring-gray-900/15 focus:border-gray-900 shadow-[inset_0_0_0_1px_rgba(15,23,42,0.02)] placeholder-gray-400 disabled:opacity-40 disabled:cursor-not-allowed"
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
    <div className="space-y-6 sm:space-y-8">
      <div className="flex flex-col sm:flex-row sm:justify-between sm:items-center gap-4">
        <div className="min-w-0">
          <p className="text-[13px] text-gray-500">
            Portfolio audit board — enter and review hours by site and cleaner.
          </p>
        </div>
        <div className="flex gap-2 flex-shrink-0">
          <button
            onClick={() =>
              exportFortnightTimesheets(
                currentPeriod,
                sites,
                cleaners,
                entries as any,
                'xlsx'
              )
            }
            className="flex items-center gap-1.5 px-4 py-2 so-btn-secondary text-xs font-semibold"
          >
            <FileSpreadsheet size={14} />
            Export (XLSX)
          </button>
        </div>
      </div>

      <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-6 gap-3">
        {([
          { key: 'all' as QueueKey, label: 'All Sites', count: queueCounts.all },
          { key: 'needs-hours' as QueueKey, label: 'Sites needing hours', count: queueCounts.needs },
          { key: 'incomplete' as QueueKey, label: 'Incomplete sites', count: queueCounts.incomplete },
          { key: 'over-budget' as QueueKey, label: 'Over budget sites', count: queueCounts.over },
          { key: 'completed' as QueueKey, label: 'Completed sites', count: queueCounts.completed },
        ]).map((card) => {
          const active = queueKey === card.key;
          return (
            <button
              key={card.key}
              type="button"
              onClick={() => setQueueKey(card.key)}
              className={`text-left so-card-soft px-3.5 py-3 rounded-xl cursor-pointer focus:outline-none focus:ring-2 focus:ring-[#3E5F6A]/40 ${
                active
                  ? 'border-[#3E5F6A] bg-[#ECF3F4]'
                  : 'border-transparent hover:border-[#3E5F6A]/50'
              }`}
            >
              <p className="text-[11px] font-medium text-gray-500 uppercase tracking-[0.18em] mb-1">
                {card.label}
              </p>
              <p className={`text-[20px] font-semibold ${active ? 'text-[#3E5F6A]' : 'text-gray-900'}`}>
                {card.count}
              </p>
            </button>
          );
        })}
        <div className="text-left so-card-soft px-3.5 py-3 rounded-xl border-transparent">
          <p className="text-[11px] font-medium text-gray-500 uppercase tracking-[0.18em] mb-1">
            Estimated budget
          </p>
          <p className="text-[20px] font-semibold text-gray-900">
            {new Intl.NumberFormat('en-AU', { style: 'currency', currency: 'AUD', minimumFractionDigits: 0, maximumFractionDigits: 0 }).format(totalEstimatedBudget)}
          </p>
        </div>
      </div>
      <div className="relative">
        <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" size={18} />
        <input
          type="search"
          placeholder="Search sites by name or address…"
          value={siteSearchQuery}
          onChange={(e) => setSiteSearchQuery(e.target.value)}
          className="w-full pl-10 pr-4 py-2.5 so-input text-sm text-gray-900 placeholder-gray-400 bg-white"
          aria-label="Search sites"
        />
      </div>
      <div className="space-y-3">
        {orderedSites.map(site => {
          const summary = siteSummaryById[site.id];
          const budget = summary?.budgetTotal ?? 0;
          const actual = summary?.actualTotal ?? 0;
          const variance = summary?.variance ?? 0;
          const isBalanced = Math.abs(variance) < 0.05 && budget > 0;
          const over = variance > 0.05;
          const under = variance < -0.05;

          const cardClasses = [
            "group bg-white border border-[#edeef0] rounded-lg p-4 cursor-pointer flex items-center justify-between transition-colors",
          ];
          if (isBalanced) {
            cardClasses.push("border-green-500 bg-green-50/70");
          } else {
            cardClasses.push("hover:border-gray-900");
          }

          return (
          <div
            key={site.id}
            onClick={() => setActiveSiteId(site.id)}
            className={cardClasses.join(" ")}
          >
            <div className="flex items-center gap-4 w-1/3">
              <div className={`w-10 h-10 border border-[#edeef0] rounded-xl flex items-center justify-center text-gray-400 ${isBalanced ? "bg-green-100" : "bg-gray-50 group-hover:bg-white"}`}>
                <Building size={20} />
              </div>
              <div className="min-w-0">
                <h4 className="text-sm font-bold text-gray-900 truncate">
                  {site.name}
                </h4>
                <p className="text-[10px] text-gray-400 truncate uppercase font-bold">
                  {site.address}
                </p>
              </div>
            </div>
            <div className="flex -space-x-2 w-32">
              {site.assigned_cleaner_ids.slice(0, 3).map(cid => (
                <div
                  key={cid}
                  className={`w-8 h-8 rounded-xl border-2 border-white flex items-center justify-center text-[10px] font-bold shadow-sm ${
                    isBalanced ? "bg-green-100 text-green-800" : "bg-gray-100 text-gray-500"
                  }`}
                >
                  {cleaners.find(c => c.id === cid)?.firstName.charAt(0)}
                </div>
              ))}
            </div>
            <div className="text-right border-l border-gray-100 pl-8 min-w-[120px] space-y-2">
              <div>
                <p className="text-[9px] font-bold text-gray-400 uppercase">
                  Assigned budget
                </p>
                <p className="text-xs font-black text-gray-800">
                  {budget.toFixed(1)}h
                </p>
                <p className="text-xs font-black text-gray-800">
                  {new Intl.NumberFormat('en-AU', { style: 'currency', currency: 'AUD', minimumFractionDigits: 0, maximumFractionDigits: 0 }).format(
                    budget * (site.budget_weekday_labour_rate ?? site.budget_labour_rate ?? 0)
                  )}
                </p>
              </div>
              <div>
                <p className="text-[9px] font-bold text-gray-400 uppercase">
                  Est. budget
                </p>
                <p className="text-xs font-black text-gray-800">
                  {new Intl.NumberFormat('en-AU', { style: 'currency', currency: 'AUD', minimumFractionDigits: 0, maximumFractionDigits: 0 }).format(estimatedBudgetBySiteId[site.id] ?? 0)}
                </p>
              </div>
              {summary && (
                <div className="mt-1">
                  {isBalanced ? (
                    <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[10px] font-bold uppercase bg-green-600 text-white">
                      <Check size={10} />
                      Balanced
                    </span>
                  ) : (
                    <span
                      className={`inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[10px] font-bold uppercase ${
                        over
                          ? "bg-red-50 text-red-700"
                          : under
                          ? "bg-amber-50 text-amber-700"
                          : "bg-gray-50 text-gray-600"
                      }`}
                    >
                      {over ? (
                        <TrendingUp size={10} />
                      ) : under ? (
                        <TrendingDown size={10} />
                      ) : null}
                      {variance > 0 ? "+" : ""}
                      {variance.toFixed(1)}h
                    </span>
                  )}
                </div>
              )}
            </div>
          </div>
        );})}
        {siteSearchQuery.trim() && filteredSites.length === 0 && (
          <p className="text-sm text-gray-500 py-6 text-center">No sites match your search.</p>
        )}
      </div>
    </div>
  );
};

export default TimeEntryForm;
