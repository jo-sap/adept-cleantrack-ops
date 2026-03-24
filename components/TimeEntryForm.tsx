
import React, { useState, useMemo, useEffect, useRef, useCallback } from 'react';
import { Site, Cleaner, TimeEntry, FortnightPeriod, AdHocJob } from '../types';
import { format, getDay, addDays } from 'date-fns';
import { useRole } from '../contexts/RoleContext';
import { useAppAuth } from '../contexts/AppAuthContext';
import { getCleanTrackUserByEmail } from '../repositories/usersRepo';
import { getDayStatus } from '../utils';
import { getSiteRateForDate, computeBudgetedLabourCostForRange } from '../lib/budgetedLabourCost';
import { getPublicHolidaysInRange } from '../lib/publicHolidays';
import { getNoServicePeriodForDate } from '../lib/noServicePeriods';
import { Save, Zap, Building, ArrowLeft, UserPlus, TrendingUp, TrendingDown, Check, Download, FileSpreadsheet, Search, Loader2, Eraser, Copy, Briefcase } from 'lucide-react';
import { exportFortnightTimesheets } from '../services/exportService';
import { getGraphAccessToken } from '../lib/graph';
import { getAdHocJobs } from '../repositories/adHocJobsRepo';
import {
  listAllTimesheetPeriodNotes,
  upsertTimesheetPeriodNote,
  pickSiteNoteForPeriod,
  buildSiteNotesExportLookup,
  normalizeSiteLabelForNotes,
  comparablePeriodYmd,
  type SiteNotesExportLookup,
} from '../repositories/timesheetNotesRepo';
import { generateAdHocOccurrencesForRange, occurrencesToHoursByDate, adHocJobHasPlannedWorkInRange } from '../lib/adhocSchedule';
import TimesheetPeriodNotesPanel from './TimesheetPeriodNotesPanel';

interface TimeEntryFormProps {
  sites: Site[];
  cleaners: Cleaner[];
  entries: TimeEntry[];
  currentPeriod: FortnightPeriod;
  onSaveBatch: (entries: Omit<TimeEntry, 'id'>[]) => void;
  onDeleteEntry: (id: string) => void;
  onUpdateSite?: (site: Site) => void;
}

type QueueKey = 'all' | 'needs-hours' | 'incomplete' | 'over-budget' | 'completed' | 'adhoc';

function normalizeScheduleType(raw: string | undefined | null): 'once_off' | 'recurring' {
  const s = String(raw ?? '').trim().toLowerCase();
  if (!s) return 'once_off';
  if (s.includes('recurr')) return 'recurring';
  return 'once_off';
}

const TimeEntryForm: React.FC<TimeEntryFormProps> = ({ 
  sites, cleaners, entries, currentPeriod, onSaveBatch, onUpdateSite
}) => {
  const { isAdmin, isManager } = useRole();
  const { user: authUser } = useAppAuth();
  const [activeSiteId, setActiveSiteId] = useState<string | null>(null);
  const [selectedCleanerId, setSelectedCleanerId] = useState<string>('');
  const [draftHours, setDraftHours] = useState<Record<string, number>>({});
  const [isSaved, setIsSaved] = useState(false);
  const [hasUnsavedChanges, setHasUnsavedChanges] = useState(false);
  const [siteSearchQuery, setSiteSearchQuery] = useState('');
  const [saveLoading, setSaveLoading] = useState(false);
  const [adhocJobId, setAdhocJobId] = useState<string | null>(null);
  const [activeAdHocCardJobId, setActiveAdHocCardJobId] = useState<string | null>(null);
  const [adHocJobsForSite, setAdHocJobsForSite] = useState<AdHocJob[]>([]);
  const [queueKey, setQueueKey] = useState<QueueKey>('all');
  const [fortnightAdHocJobs, setFortnightAdHocJobs] = useState<AdHocJob[]>([]);
  const [managerNoteBody, setManagerNoteBody] = useState('');
  const [managerNoteSavedBody, setManagerNoteSavedBody] = useState('');
  const [notesListLoading, setNotesListLoading] = useState(false);
  const [notesListMissing, setNotesListMissing] = useState(false);
  const [notesListSchemaError, setNotesListSchemaError] = useState<string | null>(null);

  const inputRefs = useRef<Record<string, HTMLInputElement | null>>({});

  const isVirtualAdHocContext = !!activeSiteId && activeSiteId.startsWith("adhoc:");
  const virtualAdHocSite = useMemo(() => {
    if (!activeSiteId || !activeSiteId.startsWith("adhoc:")) return null;
    const jobId = activeSiteId.slice("adhoc:".length);
    const job =
      fortnightAdHocJobs.find((j) => j.id === jobId) ??
      adHocJobsForSite.find((j) => j.id === jobId);
    if (!job) return null;
    const address = [job.manualSiteAddress?.trim(), job.manualSiteState?.trim()]
      .filter(Boolean)
      .join(", ");
    return {
      id: activeSiteId,
      name: job.manualSiteName?.trim() || job.jobName || "Ad Hoc Site",
      address,
      is_active: true,
      budgeted_hours_per_fortnight: 0,
      daily_budgets: [0, 0, 0, 0, 0, 0, 0],
      assigned_cleaner_ids: cleaners.map((c) => c.id),
      monthly_revenue: 0,
      financial_budget: 0,
      cleaner_rates: {},
      visit_frequency: "Ad Hoc",
    } as Site;
  }, [activeSiteId, fortnightAdHocJobs, adHocJobsForSite, cleaners]);
  const activeSite = sites.find((s) => s.id === activeSiteId) ?? virtualAdHocSite ?? undefined;
  const activeCleaner = cleaners.find(c => c.id === selectedCleanerId);

  const managerNoteDirty = useMemo(
    () =>
      (isAdmin || isManager) &&
      managerNoteBody.trim() !== managerNoteSavedBody.trim(),
    [isAdmin, isManager, managerNoteBody, managerNoteSavedBody]
  );
  const canSaveAnything = hasUnsavedChanges || managerNoteDirty;

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

  const getPlannedHoursForDate = useCallback(
    (site: Site, dateIndex: number, date: Date): number => {
      if (getNoServicePeriodForDate(date, site.no_service_periods)) return 0;
      const visit = String(site.visit_frequency ?? "").trim().toLowerCase();
      // Monthly/fortnightly "period cap" mode: no per-day plan (handled elsewhere)
      if (visit === "monthly") return 0;
      const week2 = (site as any).daily_budgets_week2 as number[] | undefined;
      const budgets =
        visit === "fortnightly" && week2 && week2.length >= 7 && dateIndex >= 7
          ? week2
          : site.daily_budgets;
      return budgets?.[getDay(date)] ?? 0;
    },
    []
  );

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
      const isPeriodBudget =
        site.visit_frequency === "Monthly" || site.visit_frequency === "Fortnightly";
      const periodCap = site.budgeted_hours_per_fortnight ?? 0;
      const dailyPlanSum = dates.reduce(
        (s, d, idx) => s + getPlannedHoursForDate(site, idx, d),
        0
      );
      const usePeriodCap = isPeriodBudget && periodCap > 0 && dailyPlanSum === 0;
      const budgetTotal = usePeriodCap ? periodCap : dailyPlanSum;
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
  }, [sites, entries, currentPeriod, dates, getPlannedHoursForDate]);

  const phInPeriod = useMemo(
    () => getPublicHolidaysInRange(currentPeriod.startDate, currentPeriod.endDate),
    [currentPeriod]
  );

  useEffect(() => {
    let cancelled = false;
    getGraphAccessToken().then(async (token) => {
      if (!token || cancelled) return;
      try {
        const filters: { assignedManagerId?: string } = {};
        if (isManager && authUser?.email) {
          const ctUser = await getCleanTrackUserByEmail(token, authUser.email);
          if (ctUser?.id) filters.assignedManagerId = ctUser.id;
        }
        const all = await getAdHocJobs(
          token,
          Object.keys(filters).length > 0 ? filters : undefined
        );
        if (cancelled) return;
        const eligible = all.filter((j) => {
          if (j.timesheetApplicable === false) return false;
          return adHocJobHasPlannedWorkInRange(j, currentPeriod.startDate, currentPeriod.endDate, phInPeriod);
        });
        eligible.sort((a, b) =>
          (a.jobName || "").localeCompare(b.jobName || "", undefined, { sensitivity: "base" })
        );
        setFortnightAdHocJobs(eligible);
      } catch {
        if (!cancelled) setFortnightAdHocJobs([]);
      }
    });
    return () => {
      cancelled = true;
    };
  }, [
    currentPeriod.startDate,
    currentPeriod.endDate,
    phInPeriod,
    isManager,
    authUser?.email,
  ]);

  const activeAdHocJob = useMemo(
    () =>
      adhocJobId
        ? adHocJobsForSite.find((j) => j.id === adhocJobId) ??
          fortnightAdHocJobs.find((j) => j.id === adhocJobId)
        : undefined,
    [adhocJobId, adHocJobsForSite, fortnightAdHocJobs]
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
      const visit = String(site.visit_frequency ?? "").trim().toLowerCase();
      // For Fortnightly schedules with Week 2 budgets, compute expected cost per date so week2 is respected.
      // Otherwise keep the existing 7-day budget calculation.
      let cost = 0;
      if (visit === "fortnightly" && (site as any).daily_budgets_week2) {
        dates.forEach((d, idx) => {
          const planned = getPlannedHoursForDate(site, idx, d);
          if (planned <= 0) return;
          const rate = getSiteRateForDate(d, weekdayRate, saturdayRate, sundayRate, phRate, phInPeriod);
          cost += planned * rate;
        });
      } else {
        const dailyBudgets = (site.daily_budgets?.length ?? 0) >= 7
          ? site.daily_budgets
          : [0, 0, 0, 0, 0, 0, 0];
        cost = computeBudgetedLabourCostForRange({
          startDate: currentPeriod.startDate,
          endDate: currentPeriod.endDate,
          dailyBudgets,
          weekdayRate,
          saturdayRate,
          sundayRate,
          phRate,
          publicHolidayDates: phInPeriod,
        });
      }
      map[site.id] = cost;
    });
    return map;
  }, [sites, currentPeriod, phInPeriod, dates, getPlannedHoursForDate]);

  const adHocSummaryById = useMemo(() => {
    const map: Record<
      string,
      { plannedHours: number; plannedCost: number; loggedHours: number; loggedCost: number; variance: number }
    > = {};
    const start = currentPeriod.startDate;
    const endExclusive = addDays(currentPeriod.endDate, 1);

    fortnightAdHocJobs.forEach((job) => {
      const occ = generateAdHocOccurrencesForRange(job, currentPeriod.startDate, currentPeriod.endDate, phInPeriod);
      const plannedHours = occ.reduce((s, o) => s + o.hours, 0);
      const plannedCost = occ.reduce((s, o) => s + o.costTotal, 0);

      let loggedHours = 0;
      let loggedCost = 0;
      entries.forEach((e) => {
        if (e.adhocJobId !== job.id) return;
        const d = new Date(e.date);
        if (d < start || d >= endExclusive) return;
        loggedHours += e.hours;
        loggedCost += e.hours * (e.pay_rate_snapshot ?? 0);
      });

      map[job.id] = {
        plannedHours,
        plannedCost,
        loggedHours,
        loggedCost,
        variance: loggedHours - plannedHours,
      };
    });

    return map;
  }, [fortnightAdHocJobs, currentPeriod.startDate, currentPeriod.endDate, phInPeriod, entries]);

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
      adhoc: fortnightAdHocJobs.length,
    };
  }, [filteredSites, siteSummaryById, fortnightAdHocJobs.length]);

  const orderedSites = useMemo(() => {
    if (queueKey === 'adhoc') {
      const siteHasAdHoc = (siteId: string) =>
        fortnightAdHocJobs.some(
          (j) => j.siteId && String(j.siteId).trim() === String(siteId).trim()
        );
      return [...filteredSites].sort((a, b) => {
        const ah = siteHasAdHoc(a.id) ? 0 : 1;
        const bh = siteHasAdHoc(b.id) ? 0 : 1;
        if (ah !== bh) return ah - bh;
        return (a.name || "").localeCompare(b.name || "", undefined, { sensitivity: "base" });
      });
    }
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
  }, [filteredSites, queueKey, siteSummaryById, fortnightAdHocJobs]);

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
      if (!e.cleanerId) return;
      if (!isVirtualAdHocContext && e.siteId !== activeSiteId) return;
      if (!e.adhocJobId || e.adhocJobId !== adhocJobId) return;
      const key = e.date;
      if (!map[key]) map[key] = {};
      map[key][e.cleanerId] = (map[key][e.cleanerId] || 0) + e.hours;
    });
    return map;
  }, [entries, activeSiteId, adhocJobId, isVirtualAdHocContext]);

  useEffect(() => {
    if (!activeSiteId || !activeSite) {
      setSelectedCleanerId('');
      return;
    }
    const adhocMode = !!activeAdHocCardJobId;
    const personnelIds = adhocMode
      ? cleaners.map((c) => c.id)
      : (activeSite.assigned_cleaner_ids ?? []);
    if (personnelIds.length === 0) {
      setSelectedCleanerId('');
      return;
    }
    const firstId = personnelIds[0];
    setSelectedCleanerId(prev => {
      if (prev && personnelIds.includes(prev)) return prev;
      return firstId;
    });
  }, [activeSiteId, activeSite, activeAdHocCardJobId, cleaners]);

  // For standard site timesheets, always stay on contract work.
  // Ad hoc context is only entered from the dedicated ad hoc job cards.
  useEffect(() => {
    if (!activeSiteId || !selectedCleanerId) return;
    if (isVirtualAdHocContext) {
      setAdhocJobId(activeAdHocCardJobId ?? adhocJobId);
      return;
    }
    if (!activeAdHocCardJobId) {
      setAdhocJobId(null);
      return;
    }
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
  }, [activeSiteId, selectedCleanerId, currentPeriod.id, entries, dates, isVirtualAdHocContext, activeAdHocCardJobId, adhocJobId]);

  // Load draft hours for the current (site, cleaner, ad hoc context) selection.
  useEffect(() => {
    if (!activeSiteId || !selectedCleanerId) return;
    const newDraft: Record<string, number> = {};
    dates.forEach((date) => {
      const dateStr = format(date, 'yyyy-MM-dd');
      const existing = entries.find(
        (e) =>
          (isVirtualAdHocContext ? true : e.siteId === activeSiteId) &&
          e.cleanerId === selectedCleanerId &&
          e.date === dateStr &&
          (adhocJobId ? e.adhocJobId === adhocJobId : !e.adhocJobId)
      );
      newDraft[dateStr] = existing ? existing.hours : 0;
    });
    setDraftHours(newDraft);
    setHasUnsavedChanges(false);
  }, [activeSiteId, selectedCleanerId, adhocJobId, currentPeriod.id, entries, dates, isVirtualAdHocContext]);

  useEffect(() => {
    if (!activeSiteId) {
      setAdHocJobsForSite([]);
      return;
    }
    if (isVirtualAdHocContext) {
      const jobId = activeSiteId.slice("adhoc:".length);
      const picked = fortnightAdHocJobs.find((j) => j.id === jobId);
      setAdHocJobsForSite(picked ? [picked] : []);
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
  }, [activeSiteId, currentPeriod.startDate.getTime(), currentPeriod.endDate.getTime(), isVirtualAdHocContext, fortnightAdHocJobs]);

  useEffect(() => {
    if (!(isAdmin || isManager) || !activeSiteId || !activeSite) {
      setManagerNoteBody("");
      setManagerNoteSavedBody("");
      setNotesListMissing(false);
      setNotesListSchemaError(null);
      setNotesListLoading(false);
      return;
    }
    let cancelled = false;
    setNotesListLoading(true);
    getGraphAccessToken().then(async (token) => {
      if (!token || cancelled) {
        if (!cancelled) setNotesListLoading(false);
        return;
      }
      try {
        const { notes, listExists, listSchemaError: schemaErr } =
          await listAllTimesheetPeriodNotes(token);
        if (cancelled) return;
        const periodYmd = format(currentPeriod.startDate, "yyyy-MM-dd");
        const adhocTag = adhocJobId ? `adhocJob:${adhocJobId}` : "";
        const adhocNameNorm = normalizeSiteLabelForNotes(activeAdHocJob?.jobName ?? "");
        const picked = adhocTag
          ? notes.find((n) =>
              !n.cleanerId &&
              comparablePeriodYmd(n.periodStartYmd) === periodYmd &&
              (
                (n.tags ?? []).some((t) => t.trim().toLowerCase() === adhocTag.toLowerCase()) ||
                (!!adhocNameNorm && normalizeSiteLabelForNotes(n.siteLookupName) === adhocNameNorm)
              )
            )
          : pickSiteNoteForPeriod(notes, activeSiteId, periodYmd, activeSite?.name);
        const body = picked?.noteBody ?? "";
        setManagerNoteBody(body);
        setManagerNoteSavedBody(body);
        setNotesListMissing(!listExists);
        setNotesListSchemaError(schemaErr ?? null);
      } catch {
        if (!cancelled) {
          setNotesListMissing(true);
          setNotesListSchemaError(null);
        }
      } finally {
        if (!cancelled) setNotesListLoading(false);
      }
    });
    return () => {
      cancelled = true;
    };
  }, [isAdmin, isManager, activeSiteId, activeSite, currentPeriod.id, currentPeriod.startDate, adhocJobId]);

  const handleHourChange = (dateStr: string, val: string) => {
    // Allow quarter‑hour (and other 2‑decimal) precision like 4.75
    const hours = val === '' ? 0 : Math.round(parseFloat(val) * 100) / 100;
    setDraftHours(prev => ({ ...prev, [dateStr]: hours }));
    setHasUnsavedChanges(true);
  };

  const handleSave = async () => {
    if (!activeSiteId || !selectedCleanerId || !activeCleaner || !activeSite) return;
    if (!hasUnsavedChanges && !managerNoteDirty) return;
    if (adhocJobId && managerNoteBody.trim() === "") {
      alert("Manager Notes are required for Ad Hoc timesheets.");
      return;
    }
    const hoursWereDirty = hasUnsavedChanges;
    setSaveLoading(true);
    try {
      let anySuccess = false;
      if (hoursWereDirty) {
        const batchData = (Object.entries(draftHours) as [string, number][]).map(([date, hours]) => ({
          siteId: isVirtualAdHocContext ? undefined : activeSiteId,
          cleanerId: selectedCleanerId,
          date,
          hours,
          pay_rate_snapshot: activeSite.cleaner_rates[selectedCleanerId] || activeCleaner.payRatePerHour || 0,
          adhocJobId: adhocJobId || undefined
        }));
        await onSaveBatch(batchData as any);
        setHasUnsavedChanges(false);
        anySuccess = true;
      }

      const canPersistNote =
        (isAdmin || isManager) && !notesListMissing && !notesListSchemaError;
      if (canPersistNote && managerNoteDirty) {
        const token = await getGraphAccessToken();
        if (!token) {
          alert("Sign in with Microsoft to save the manager note.");
        } else {
          const periodYmd = format(currentPeriod.startDate, "yyyy-MM-dd");
          const adhocTag = adhocJobId ? `adhocJob:${adhocJobId}` : "";
          const noteResult = await upsertTimesheetPeriodNote(token, {
            siteId: isVirtualAdHocContext ? "" : activeSiteId,
            siteName: activeAdHocJob?.jobName ?? activeSite.name ?? "",
            periodStartYmd: periodYmd,
            cleanerId: null,
            tags: adhocTag ? [adhocTag] : [],
            noteBody: managerNoteBody,
          });
          if (!noteResult.ok) {
            alert(
              hoursWereDirty
                ? `Timesheet saved, but the manager note failed: ${noteResult.error}`
                : noteResult.error
            );
          } else {
            setManagerNoteSavedBody(managerNoteBody.trim());
            anySuccess = true;
          }
        }
      }

      if (anySuccess) {
        setIsSaved(true);
        setTimeout(() => setIsSaved(false), 2000);
      }
    } finally {
      setSaveLoading(false);
    }
  };

  const handleClearAll = () => {
    if (!activeSiteId || !selectedCleanerId) return;
    const label = adhocJobId ? "this Ad Hoc job" : "contract / standard work";
    const ok = window.confirm(
      `Clear all hours for ${label} in this fortnight? This will set all 14 days to 0.0 hours (not saved until you click Save).`
    );
    if (!ok) return;
    const next: Record<string, number> = {};
    dates.forEach((d) => {
      next[format(d, "yyyy-MM-dd")] = 0;
    });
    setDraftHours(next);
    setHasUnsavedChanges(true);
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
    dates.forEach((date, idx) => {
      const dateStr = format(date, "yyyy-MM-dd");
      const current = next[dateStr] ?? 0;
      const sitePlan = getPlannedHoursForDate(activeSite, idx, date);
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

  /** Map each day in the current grid to the same calendar position in the prior pay fortnight (site + selected cleaner + contract vs ad hoc). */
  const handleCopyPreviousFortnight = () => {
    if (!activeSiteId || !selectedCleanerId) return;
    const matchesContext = (e: TimeEntry) =>
      (isVirtualAdHocContext ? true : e.siteId === activeSiteId) &&
      e.cleanerId === selectedCleanerId &&
      (adhocJobId ? e.adhocJobId === adhocJobId : !e.adhocJobId);

    const prevDateStrs = dates.map((d) => format(addDays(d, -14), "yyyy-MM-dd"));
    const hasAnyPriorSaved = prevDateStrs.some((prevStr) =>
      entries.some((e) => e.date === prevStr && matchesContext(e))
    );
    if (!hasAnyPriorSaved) {
      window.alert(
        "No saved timesheet found in the previous fortnight for this site, cleaner, and job type."
      );
      return;
    }
    const ok = window.confirm(
      "Replace all 14 days in this view with hours from the previous fortnight for this cleaner? Unsaved changes will be overwritten. Click Save to persist."
    );
    if (!ok) return;

    const next: Record<string, number> = {};
    dates.forEach((date) => {
      const curStr = format(date, "yyyy-MM-dd");
      const prevStr = format(addDays(date, -14), "yyyy-MM-dd");
      const hours = entries
        .filter((e) => e.date === prevStr && matchesContext(e))
        .reduce((s, e) => s + e.hours, 0);
      next[curStr] = hours;
    });
    setDraftHours(next);
    setHasUnsavedChanges(true);
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
      (s, d, idx) => s + getPlannedHoursForDate(activeSite, idx, d),
      0
    );
    const usePeriodCap = isPeriodBudget && periodCap > 0 && dailyPlanSum === 0;
    const budgetTotal = usePeriodCap
      ? periodCap
      : dates.reduce((s, date, idx) => s + getPlannedHoursForDate(activeSite, idx, date), 0);

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

    dates.forEach((date, idx) => {
      const dateStr = format(date, "yyyy-MM-dd");
      const sitePlanForDay = getPlannedHoursForDate(activeSite, idx, date);
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
              onClick={() => {
                setActiveSiteId(null);
                setActiveAdHocCardJobId(null);
                setAdhocJobId(null);
              }}
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
          {/* Supabase-based "Last Edited" audit removed (SharePoint-only). */}
        </div>

        {(() => {
          const adhocMode = !!activeAdHocCardJobId;
          const adHocLockedMode = !!activeAdHocCardJobId;
          const personnelIds = adhocMode
            ? cleaners.map((c) => c.id)
            : (activeSite.assigned_cleaner_ids ?? []);
          if (!adhocMode && personnelIds.length === 0) {
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
                  {adhocMode ? "Cleaner" : "Personnel"}
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
                <div className="mb-1 flex items-center justify-between gap-2">
                  <label className="block text-[10px] font-bold text-gray-400 uppercase">
                    {adHocLockedMode ? "Ad Hoc Job" : "Type"}
                  </label>
                  {activeAdHocJob && (
                    <span className="text-[10px] font-semibold uppercase tracking-wide text-gray-500">
                      {normalizeScheduleType(activeAdHocJob.jobType) === 'recurring' ? 'Recurring' : 'Once off'}
                    </span>
                  )}
                </div>
                {adHocLockedMode ? (
                  <div className="w-full so-input bg-gray-50 px-3 py-2 text-sm text-gray-800 border border-[#edeef0] rounded-lg">
                    {activeAdHocJob?.jobName ?? "Ad hoc job"}
                    {activeAdHocJob?.scheduledDate ? ` • ${activeAdHocJob.scheduledDate}` : ""}
                  </div>
                ) : (
                  <div className="w-full so-input bg-gray-50 px-3 py-2 text-sm text-gray-800 border border-[#edeef0] rounded-lg">
                    Contract Work
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
                {[(adhocJobId ? 'Scheduled' : 'Budget'), 'Actual', 'Variance', 'Est. Pay'].map((label, idx) => (
                  <div key={label} className="px-6 flex flex-col justify-center">
                    <span className="text-[10px] font-bold text-gray-400 uppercase tracking-widest mb-1.5">{label}</span>
                    <span className={`text-sm font-bold ${label === 'Variance' && summary && summary.variance > 0.1 ? 'text-red-600' : 'text-gray-900'}`}>
                      {idx === 0 ? (summary?.budgetDisplay ?? summary?.budgetTotal ?? 0).toFixed(1) + 'h' : idx === 1 ? summary?.actualTotal.toFixed(1) + 'h' : idx === 2 ? (summary!.variance > 0 ? '+' : '') + summary?.variance.toFixed(1) + 'h' : '$' + (summary?.estPay ?? 0).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                    </span>
                    {label === 'Budget' && summary?.isPeriodBudget && summary.periodCap > 0 && (
                      <span className="text-[9px] text-gray-400 mt-0.5 leading-tight">Site cap: {summary.periodCap.toFixed(1)}h</span>
                    )}
                  </div>
                ))}
              </div>
              <div className="flex items-center gap-3 pr-2">
                {!adhocJobId && (
                  <button
                    type="button"
                    onClick={handleAutoFill}
                    className="flex flex-col items-center justify-center w-24 h-14 rounded-xl border border-[#edeef0] bg-white text-gray-700 hover:bg-gray-50 hover:border-gray-300 transition-all"
                    title="Fill empty days with remaining plan hours"
                  >
                    <Zap size={12} className="opacity-70 mb-0.5" /><span className="text-[10px] font-bold uppercase">Auto fill</span>
                  </button>
                )}
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
                  type="button"
                  onClick={handleCopyPreviousFortnight}
                  disabled={saveLoading}
                  className="flex flex-col items-center justify-center w-24 h-14 rounded-xl border border-[#edeef0] bg-white text-gray-700 hover:bg-gray-50 hover:border-gray-300 transition-all disabled:opacity-40"
                  title="Copy saved hours from the previous pay fortnight for this site and cleaner (same job type). Save to persist."
                >
                  <Copy size={12} className="opacity-70 mb-0.5" />
                  <span className="text-[10px] font-bold uppercase leading-tight text-center px-0.5">
                    Last FN
                  </span>
                </button>
                <button
                  type="button"
                  onClick={handleClearAll}
                  disabled={saveLoading}
                  className="flex flex-col items-center justify-center w-24 h-14 rounded-xl border border-[#edeef0] bg-white text-gray-700 hover:bg-gray-50 hover:border-gray-300 transition-all disabled:opacity-40"
                  title="Clear all hours for this fortnight (not saved until Save)"
                >
                  <Eraser size={12} className="opacity-70 mb-0.5" />
                  <span className="text-[10px] font-bold uppercase">Clear</span>
                </button>
                <button
                  onClick={handleSave}
                  disabled={!canSaveAnything || saveLoading}
                  className={`flex flex-col items-center justify-center w-24 h-14 rounded-xl transition-all ${
                    isSaved
                      ? 'bg-green-600 text-white'
                      : 'so-btn-primary disabled:opacity-40'
                  }`}
                >
                  {saveLoading ? <Loader2 size={18} className="animate-spin mb-0.5" /> : <Save size={12} className="opacity-70 mb-0.5" />}
                  <span className="text-[10px] font-bold uppercase">{saveLoading ? 'Saving…' : 'Save'}</span>
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
                const noServicePeriod = isAdHoc
                  ? undefined
                  : getNoServicePeriodForDate(date, activeSite.no_service_periods);
                const dayAssignments = (isAdHoc ? adHocEntriesByDate : siteEntriesByDate)[dateStr] || {};
                const sitePlan = isAdHoc ? 0 : getPlannedHoursForDate(activeSite, index, date);
                const otherHours = Object.entries(dayAssignments)
                  .filter(([cid]) => cid !== selectedCleanerId)
                  .reduce((sum, [, h]) => sum + h, 0);
                const remainingPlan = Math.max(sitePlan - otherHours, 0);
                const planned = isAdHoc ? (activeAdHocPlannedByDate[dateStr] ?? 0) : remainingPlan;
                const hours = draftHours[dateStr] || 0;
                const status = isAdHoc
                  ? { border: "border-[#edeef0]", bg: "bg-gray-50", dot: "bg-gray-300", color: "text-gray-600", label: "Adhoc" }
                  : noServicePeriod
                    ? { border: "border-green-200", bg: "bg-green-50", dot: "bg-green-500", color: "text-green-600", label: "On target" }
                  : getDayStatus(planned, hours);
                const existingForThisCleaner = dayAssignments[selectedCleanerId] || 0;
                const fullyAllocatedToOthers =
                  !isAdHoc && sitePlan > 0 && remainingPlan <= 0 && existingForThisCleaner === 0;
                const noServiceLabel = noServicePeriod?.label?.trim() || noServicePeriod?.reason?.trim() || "No Service";
                return (
                  <div key={dateStr} onClick={() => inputRefs.current[dateStr]?.focus()} className={`flex flex-col p-3 bg-white border rounded-xl cursor-pointer group relative overflow-hidden ${status.border} ${status.bg.replace('bg-', 'hover:bg-')}`}>
                    <div className={`absolute top-2 right-2 w-1.5 h-1.5 rounded-full ${status.dot}`} />
                    <div className="mb-2"><p className="text-[10px] font-bold text-gray-900 uppercase">{format(date, 'EEE')}</p><p className="text-[9px] font-medium text-gray-400">{format(date, 'MMM d')}</p></div>
                    <div className="mt-auto space-y-2">
                      <div className="flex justify-between items-center"><span className="text-[8px] font-bold text-gray-400">{isAdHoc ? "Scheduled" : "Plan"}</span><span className="text-[9px] font-bold text-gray-800">{planned.toFixed(1)}h</span></div>
                      <input
                        ref={el => inputRefs.current[dateStr] = el}
                        type="number"
                        step="0.01"
                        value={draftHours[dateStr] || ''}
                        onChange={e => handleHourChange(dateStr, e.target.value)}
                        className="w-full px-2.5 py-2 text-center text-sm font-semibold bg-white border border-gray-300 rounded-lg focus:ring-2 focus:ring-gray-900/15 focus:border-gray-900 shadow-[inset_0_0_0_1px_rgba(15,23,42,0.02)] placeholder-gray-400 disabled:opacity-40 disabled:cursor-not-allowed"
                        disabled={fullyAllocatedToOthers}
                      />
                      {noServicePeriod && (
                        <div className="text-center text-[8px] font-bold uppercase text-gray-500">
                          {noServiceLabel}
                        </div>
                      )}
                      <div className={`mt-1 text-center py-0.5 rounded-full text-[8px] font-black uppercase ${status.bg} ${status.color} border ${status.border}`}>{status.label}</div>
                    </div>
                  </div>
                );
              })}
            </div>
            {(isAdmin || isManager) && (
              <TimesheetPeriodNotesPanel
                currentPeriod={currentPeriod}
                canEdit={isAdmin || isManager}
                noteBody={managerNoteBody}
                onNoteChange={setManagerNoteBody}
                notesLoading={notesListLoading}
                listMissing={notesListMissing}
                listSchemaError={notesListSchemaError}
              />
            )}
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
            type="button"
            onClick={async () => {
              try {
                const periodYmd = format(currentPeriod.startDate, 'yyyy-MM-dd');
                let siteNotesLookup: SiteNotesExportLookup | undefined;
                let adHocJobsForExport: AdHocJob[] = [];
                const token = await getGraphAccessToken();
                if (token) {
                  try {
                    const { notes } = await listAllTimesheetPeriodNotes(token);
                    siteNotesLookup = buildSiteNotesExportLookup(notes, periodYmd);
                  } catch {
                    /* export without notes if list unavailable */
                  }
                  try {
                    adHocJobsForExport = await getAdHocJobs(token);
                  } catch {
                    /* export still works without ad hoc metadata fallback */
                  }
                }
                exportFortnightTimesheets(
                  currentPeriod,
                  sites,
                  cleaners,
                  entries as any,
                  'xlsx',
                  siteNotesLookup,
                  adHocJobsForExport
                );
              } catch (e) {
                console.error('Export (XLSX) failed', e);
                alert(e instanceof Error ? e.message : 'Export failed. Check the browser console.');
              }
            }}
            className="flex items-center gap-1.5 px-4 py-2 so-btn-secondary text-xs font-semibold"
          >
            <FileSpreadsheet size={14} />
            Export (XLSX)
          </button>
        </div>
      </div>

      <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-4 xl:grid-cols-6 gap-3">
        {([
          { key: 'all' as QueueKey, label: 'All Sites', count: queueCounts.all },
          { key: 'needs-hours' as QueueKey, label: 'Sites needing hours', count: queueCounts.needs },
          { key: 'incomplete' as QueueKey, label: 'Incomplete sites', count: queueCounts.incomplete },
          { key: 'over-budget' as QueueKey, label: 'Over budget sites', count: queueCounts.over },
          { key: 'completed' as QueueKey, label: 'Completed sites', count: queueCounts.completed },
          { key: 'adhoc' as QueueKey, label: 'Adhoc jobs', count: queueCounts.adhoc },
        ]).map((card) => {
          const active = queueKey === card.key;
          return (
            <button
              key={card.key}
              type="button"
              onClick={() => setQueueKey(card.key)}
              className={`h-full text-left so-card-soft px-3.5 py-3 rounded-xl cursor-pointer focus:outline-none focus:ring-2 focus:ring-[#3E5F6A]/40 ${
                active
                  ? 'border-[#3E5F6A] bg-[#ECF3F4]'
                  : 'border-transparent hover:border-[#3E5F6A]/50'
              }`}
            >
              <div className="flex h-full flex-col">
              <p className="min-h-[2.1rem] text-[11px] font-medium text-gray-500 uppercase tracking-[0.18em] leading-snug">
                {card.label}
              </p>
              <p className={`mt-auto text-[20px] font-semibold leading-none tabular-nums ${active ? 'text-[#3E5F6A]' : 'text-gray-900'}`}>
                {card.count}
              </p>
              </div>
            </button>
          );
        })}
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
        {queueKey === 'adhoc' && fortnightAdHocJobs.length > 0 && (
          <div className="space-y-2 mb-1">
            <p className="text-[10px] font-bold text-gray-500 uppercase tracking-widest px-1">
              Ad hoc jobs this fortnight — select to enter timesheet hours
            </p>
            {fortnightAdHocJobs.map((job) => {
              const linkedSite = job.siteId ? sites.find((s) => String(s.id) === String(job.siteId)) : undefined;
              const targetSite = linkedSite;
              const siteLabel =
                targetSite?.name?.trim() ||
                job.manualSiteName?.trim() ||
                "Ad hoc site";
              const selected = adhocJobId === job.id;
              const sums = adHocSummaryById[job.id] ?? {
                plannedHours: 0,
                plannedCost: 0,
                loggedHours: 0,
                loggedCost: 0,
                variance: 0,
              };
              const variancePositive = sums.variance > 0.05;
              const varianceNegative = sums.variance < -0.05;
              return (
                <div
                  key={job.id}
                  role="button"
                  tabIndex={0}
                  onClick={() => {
                    const contextSiteId = targetSite ? String(targetSite.id) : `adhoc:${job.id}`;
                    setActiveSiteId(contextSiteId);
                    setAdhocJobId(job.id);
                    setActiveAdHocCardJobId(job.id);
                  }}
                  onKeyDown={(e) => {
                    if (e.key === "Enter" || e.key === " ") {
                      e.preventDefault();
                      (e.currentTarget as HTMLDivElement).click();
                    }
                  }}
                  className={`group border rounded-lg p-4 flex items-center justify-between transition-colors ${
                    selected
                      ? "border-amber-500 bg-amber-50 cursor-pointer"
                      : "border-amber-200/80 bg-amber-50/40 cursor-pointer hover:border-amber-400"
                  }`}
                >
                  <div className="flex items-center gap-4 min-w-0">
                    <div className="w-10 h-10 border border-amber-200 rounded-xl flex items-center justify-center text-amber-700 bg-amber-100/80 shrink-0">
                      <Briefcase size={20} />
                    </div>
                    <div className="min-w-0">
                      <h4 className="text-sm font-bold text-gray-900 truncate">{job.jobName}</h4>
                      <p className="text-[10px] text-gray-500 truncate uppercase font-bold">{siteLabel}</p>
                    </div>
                  </div>
                  <div className="text-right border-l border-amber-100 pl-6 min-w-[132px] space-y-1.5 shrink-0">
                    <div>
                      <p className="text-[9px] font-bold text-gray-400 uppercase">Assigned budget</p>
                      <p className="text-xs font-black text-gray-800">{sums.plannedHours.toFixed(1)}h</p>
                    </div>
                    <div>
                      <p className="text-[9px] font-bold text-gray-400 uppercase">Est. budget</p>
                      <p className="text-xs font-black text-gray-800">
                        {new Intl.NumberFormat('en-AU', { style: 'currency', currency: 'AUD', minimumFractionDigits: 0, maximumFractionDigits: 0 }).format(sums.plannedCost)}
                      </p>
                    </div>
                    <div>
                      <p
                        className={`text-[10px] font-black ${
                          variancePositive
                            ? 'text-red-700'
                            : varianceNegative
                            ? 'text-amber-700'
                            : 'text-green-700'
                        }`}
                      >
                        {sums.variance > 0 ? '+' : ''}
                        {sums.variance.toFixed(1)}h
                      </p>
                    </div>
                  </div>
                </div>
              );
            })}
          </div>
        )}
        {queueKey === 'adhoc' && fortnightAdHocJobs.length === 0 && (
          <p className="text-sm text-amber-800/90 bg-amber-50 border border-amber-100 rounded-lg px-4 py-3">
            No ad hoc jobs have planned work in this fortnight.
          </p>
        )}
        {orderedSites.map(site => {
          const summary = siteSummaryById[site.id];
          const budget = summary?.budgetTotal ?? 0;
          const actual = summary?.actualTotal ?? 0;
          const variance = summary?.variance ?? 0;
          const isBalanced = Math.abs(variance) < 0.05 && budget > 0;
          const over = variance > 0.05;
          const under = variance < -0.05;

          const cardClasses = [
            "group bg-white border border-[#edeef0] rounded-lg p-3 sm:p-4 cursor-pointer flex items-center justify-between gap-2 sm:gap-3 transition-colors",
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
            <div className="flex items-center gap-2 sm:gap-4 min-w-0 flex-1">
              <div className={`w-8 h-8 sm:w-10 sm:h-10 border border-[#edeef0] rounded-xl flex items-center justify-center text-gray-400 shrink-0 ${isBalanced ? "bg-green-100" : "bg-gray-50 group-hover:bg-white"}`}>
                <Building size={18} className="sm:w-5 sm:h-5" />
              </div>
              <div className="min-w-0">
                <h4 className="text-sm font-bold text-gray-900 leading-tight break-words">
                  {site.name}
                </h4>
                <p className="hidden sm:block text-[10px] text-gray-400 truncate uppercase font-bold">
                  {site.address}
                </p>
              </div>
            </div>
            <div className="hidden sm:flex -space-x-2 w-32">
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
            <div className="text-right border-l border-gray-100 pl-3 sm:pl-8 min-w-[96px] sm:min-w-[120px] space-y-1.5 sm:space-y-2 shrink-0">
              <div>
                <p className="text-[9px] font-bold text-gray-400 uppercase">
                  Assigned budget
                </p>
                <p className="text-[11px] sm:text-xs font-black text-gray-800">
                  {budget.toFixed(1)}h
                </p>
                <p className="text-[11px] sm:text-xs font-black text-gray-800">
                  {new Intl.NumberFormat('en-AU', { style: 'currency', currency: 'AUD', minimumFractionDigits: 0, maximumFractionDigits: 0 }).format(
                    budget * (site.budget_weekday_labour_rate ?? site.budget_labour_rate ?? 0)
                  )}
                </p>
              </div>
              <div>
                <p className="text-[9px] font-bold text-gray-400 uppercase">
                  Est. budget
                </p>
                <p className="text-[11px] sm:text-xs font-black text-gray-800">
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
