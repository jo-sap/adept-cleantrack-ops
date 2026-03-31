import * as XLSX from 'xlsx-js-style';
import { format, addDays, endOfMonth } from 'date-fns';
import { Site, Cleaner, TimeEntry, FortnightPeriod, AdHocJob } from '../types';
import * as sharepoint from '../lib/sharepoint';
import { normalizeSiteLabelForNotes } from '../lib/siteNotesLabel';
import { resolveAdHocJobNameTemplate } from '../lib/adhocPlaceholders';
import type { SiteNotesExportLookup } from '../repositories/timesheetNotesRepo';
import { getSiteRateForDate } from '../lib/budgetedLabourCost';
import { getPublicHolidaysInRange } from '../lib/publicHolidays';
import { generateAdHocOccurrencesForRange } from '../lib/adhocSchedule';
import { AU_STATES } from '../lib/auStates';

function managerNoteForExport(
  lookup: SiteNotesExportLookup | undefined,
  cleanerId?: string,
  site?: Site,
  siteName?: string,
  adHocJobIds: string[] = [],
  adHocJobNames: string[] = []
): string {
  if (!lookup) return "";
  const cleanerKey = sharepoint.normalizeListItemId(cleanerId);
  const cleanerPrefix = cleanerKey ? `${cleanerKey}|` : "";
  const sid = site ? sharepoint.normalizeListItemId(site.id) : "";
  const byId = sid ? lookup.bySiteId[sid]?.trim() || "" : "";
  const byCleanerId =
    cleanerPrefix && sid ? lookup.byCleanerSiteId[`${cleanerPrefix}${sid}`]?.trim() || "" : "";
  const normalizedName = normalizeSiteLabelForNotes(siteName || site?.name || "");
  const byName = normalizedName ? lookup.bySiteNameLower[normalizedName]?.trim() || "" : "";
  const byCleanerName =
    cleanerPrefix && normalizedName
      ? lookup.byCleanerSiteNameLower[`${cleanerPrefix}${normalizedName}`]?.trim() || ""
      : "";
  let fuzzySiteNote = "";
  if (!byId && !byName && !byCleanerId && !byCleanerName && normalizedName) {
    // Handle label drift like "Blackwoods" vs "Blackwoods Macquarie Park".
    const fuzzy = Object.entries(lookup.bySiteNameLower).find(([k]) =>
      k.includes(normalizedName) || normalizedName.includes(k)
    );
    fuzzySiteNote = fuzzy?.[1]?.trim() || "";
  }
  const siteNote = byCleanerId || byCleanerName || byId || byName || fuzzySiteNote;
  const adHocNotes = adHocJobIds
    .map((id) => {
      const tag = `adhocjob:${String(id).trim().toLowerCase()}`;
      const cleanerTagged = cleanerPrefix
        ? lookup.byCleanerAdhocTag[`${cleanerPrefix}${tag}`]?.trim() || ""
        : "";
      if (cleanerTagged) return cleanerTagged;
      return lookup.byAdhocTag[tag]?.trim() || "";
    })
    .filter((v, idx, arr) => !!v && arr.indexOf(v) === idx);
  const adHocNameNotes = adHocJobNames
    .map((name) => {
      const key = normalizeSiteLabelForNotes(name);
      if (!key) return "";
      const exact = lookup.bySiteNameLower[key]?.trim() || "";
      if (exact) return exact;
      const fuzzy = Object.entries(lookup.bySiteNameLower).find(([k]) =>
        k.includes(key) || key.includes(k)
      );
      return fuzzy?.[1]?.trim() || "";
    })
    .filter((v, idx, arr) => !!v && arr.indexOf(v) === idx);
  const allAdHocNotes = [...adHocNotes, ...adHocNameNotes].filter(
    (v, idx, arr) => !!v && arr.indexOf(v) === idx
  );
  const finalNotes = [siteNote, ...allAdHocNotes].filter(
    (v, idx, arr) => !!v && arr.indexOf(v) === idx
  );
  return finalNotes.join(" | ");
}

/** Light yellow fill for cells with a manager note (Excel RGB without #). */
const NOTE_FILL_RGB = 'FFFFEB9C';

function applyNoteColumnHighlight(
  worksheet: XLSX.WorkSheet,
  noteColumnIndex: number,
  totalRows: number
): void {
  for (let r = 1; r < totalRows; r++) {
    const addr = XLSX.utils.encode_cell({ r, c: noteColumnIndex });
    const cell = worksheet[addr] as { v?: unknown; w?: string; s?: unknown } | undefined;
    if (!cell) continue;
    const raw = cell.v ?? cell.w ?? '';
    if (String(raw).trim() === '') continue;
    cell.s = {
      fill: { fgColor: { rgb: NOTE_FILL_RGB } },
      alignment: { wrapText: true, vertical: 'top' },
    };
  }
}

function getAdHocCostRateForDate(
  job: AdHocJob,
  date: Date,
  publicHolidayDates: Set<string>
): number {
  const base = job.costRatePerHour ?? 0;
  const ymd = format(date, "yyyy-MM-dd");
  if (publicHolidayDates.has(ymd) && job.publicHolidayCostRateOverride != null) {
    return job.publicHolidayCostRateOverride ?? base;
  }
  const dow = date.getDay(); // 0=Sun..6=Sat
  if (dow === 6 && job.saturdayCostRateOverride != null) return job.saturdayCostRateOverride ?? base;
  if (dow === 0 && job.sundayCostRateOverride != null) return job.sundayCostRateOverride ?? base;
  if (job.weekdayCostRateOverride != null) return job.weekdayCostRateOverride ?? base;
  return base;
}

function extractAustralianStateCode(value: string | null | undefined): string {
  const raw = String(value ?? "").trim();
  if (!raw) return "";
  const upper = raw.toUpperCase();
  const lettersOnly = upper.replace(/[^A-Z]/g, "");

  for (const state of AU_STATES) {
    if (lettersOnly === state) return state;
  }

  if (upper.includes("NEW SOUTH WALES")) return "NSW";
  if (upper.includes("VICTORIA")) return "VIC";
  if (upper.includes("QUEENSLAND")) return "QLD";
  if (upper.includes("SOUTH AUSTRALIA")) return "SA";
  if (upper.includes("WESTERN AUSTRALIA")) return "WA";
  if (upper.includes("TASMANIA")) return "TAS";
  if (upper.includes("NORTHERN TERRITORY")) return "NT";
  if (upper.includes("AUSTRALIAN CAPITAL TERRITORY")) return "ACT";
  return "";
}

/**
 * Hours-based units for the export month: once-off uses budgeted hours; recurring sums
 * scheduled occurrence hours in the filtered month (same window as the Ad Hoc modal).
 */
function getAdHocJobUnitsForExportMonth(j: AdHocJob, monthContextDate: Date): number | null {
  const schedule = String(j.jobType ?? '')
    .trim()
    .toLowerCase();
  if (!schedule.includes('recurr')) {
    return j.budgetedHours != null ? Number(j.budgetedHours) : null;
  }
  if (!j.scheduledDate || !j.recurrenceFrequency) {
    return j.budgetedHours != null ? Number(j.budgetedHours) : null;
  }
  const startDate = new Date(j.scheduledDate);
  if (isNaN(startDate.getTime())) return j.budgetedHours != null ? Number(j.budgetedHours) : null;

  const monthStart = new Date(monthContextDate.getFullYear(), monthContextDate.getMonth(), 1);
  const monthEnd = endOfMonth(monthContextDate);
  const endDate = j.recurrenceEndDate ? new Date(j.recurrenceEndDate) : null;
  if (endDate && isNaN(endDate.getTime())) {
    return j.budgetedHours != null ? Number(j.budgetedHours) : null;
  }

  const previewStart = startDate > monthStart ? startDate : monthStart;
  const previewEnd = endDate && endDate < monthEnd ? endDate : monthEnd;
  if (previewStart > previewEnd) {
    return j.budgetedHours != null ? Number(j.budgetedHours) : null;
  }

  const phSet = getPublicHolidaysInRange(previewStart, previewEnd);
  const occurrences = generateAdHocOccurrencesForRange(j, previewStart, previewEnd, phSet);
  const units = occurrences.reduce((sum, o) => sum + o.hours, 0);
  return units;
}

/** Export current Ad Hoc jobs to XLSX. Uses the same job list as the UI (respects month/status/manager/site filters). */
export const exportAdHocJobsToSpreadsheet = (jobs: AdHocJob[], monthLabel: string) => {
  const headers = [
    'Job Name',
    'Schedule',
    'Recurrence',
    'Company',
    'Site',
    'Manual Site',
    'Manual State',
    'Assigned Manager',
    'Requested',
    'Scheduled',
    'Completed',
    'Status',
    'Budgeted',
    'Units',
    'Charge per hour',
    'Charge',
    'Cost',
    'Gross Profit',
    'Approval',
    'Requested',
    'Description',
  ];

  const scheduleTypeLabel = (raw: string | undefined | null) => {
    const s = String(raw ?? '').trim().toLowerCase();
    if (s.includes('recurr')) return 'Recurring';
    return 'Once Off';
  };

  const recurrenceSummary = (j: AdHocJob) => {
    if (scheduleTypeLabel(j.jobType) !== 'Recurring') return '';
    const freq = j.recurrenceFrequency ?? '';
    if (!freq) return 'Recurring';
    if (freq === 'Weekly' || freq === 'Fortnightly') {
      const map: Record<number, string> = { 0: 'Sun', 1: 'Mon', 2: 'Tue', 3: 'Wed', 4: 'Thu', 5: 'Fri', 6: 'Sat' };
      const wh = (j as any).weekdayHours as Record<string, number> | null;
      const parts =
        wh && Object.keys(wh).length
          ? Object.entries(wh)
              .map(([k, v]) => ({ d: parseInt(k, 10), h: Number(v) }))
              .filter((x) => Number.isFinite(x.d) && Number.isFinite(x.h) && x.h > 0)
              .sort((a, b) => a.d - b.d)
              .map((x) => `${map[x.d] ?? String(x.d)} ${x.h}h`)
          : (j.recurrenceWeekdays ?? []).map((d) => map[d] ?? String(d));
      return parts.length ? `${freq} • ${parts.join(', ')}` : `${freq}`;
    }
    if (freq === 'Monthly') {
      const mh = (j as any).monthlyHours != null ? ` • ${(j as any).monthlyHours}h` : '';
      if (j.monthlyMode === 'day_of_month') return `Monthly • Day ${j.monthlyDayOfMonth ?? ''}${mh}`.trim();
      if (j.monthlyMode === 'nth_weekday') {
        const map: Record<number, string> = { 0: 'Sun', 1: 'Mon', 2: 'Tue', 3: 'Wed', 4: 'Thu', 5: 'Fri', 6: 'Sat' };
        const wd = j.monthlyWeekday != null ? (map[j.monthlyWeekday] ?? String(j.monthlyWeekday)) : '';
        return `Monthly • ${j.monthlyWeekOfMonth ?? ''} ${wd}${mh}`.trim();
      }
      return 'Monthly';
    }
    return String(freq);
  };

  const [y, m] = monthLabel.split('-').map(Number);
  const monthContextDate = new Date(y, (m || 1) - 1, 1);

  const rows: any[][] = jobs.map((j) => {
    const units = getAdHocJobUnitsForExportMonth(j, monthContextDate);
    return [
      resolveAdHocJobNameTemplate(j.jobName, monthContextDate) || j.jobName || '',
      scheduleTypeLabel(j.jobType),
      recurrenceSummary(j),
      j.companyName ?? '',
      j.siteName || j.manualSiteName || '',
      j.manualSiteAddress ?? '',
      j.manualSiteState ?? '',
      j.assignedManagerName ?? '',
      j.requestedDate ? format(new Date(j.requestedDate), 'dd MMM yyyy') : '',
      j.scheduledDate ? format(new Date(j.scheduledDate), 'dd MMM yyyy') : '',
      j.completedDate ? format(new Date(j.completedDate), 'dd MMM yyyy') : '',
      j.status ?? '',
      j.budgetedHours != null ? j.budgetedHours : '',
      units != null ? units : '',
      j.chargeRatePerHour != null ? Number(j.chargeRatePerHour) : '',
      j.charge != null ? Number(j.charge) : '',
      j.cost != null ? Number(j.cost) : '',
      j.grossProfit != null ? Number(j.grossProfit) : '',
      j.approvalProofUploaded ? 'Yes' : 'No',
      j.requestedByName ?? '',
      j.description ?? '',
    ];
  });

  if (rows.length === 0) {
    alert('No ad hoc jobs to export for the selected filters.');
    return;
  }

  const worksheet = XLSX.utils.aoa_to_sheet([headers, ...rows]);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Ad Hoc Jobs');

  const date = new Date(y, (m || 1) - 1, 1);
  const safeMonth = format(date, 'MMMM_yyyy').replace(/\s+/g, '_');
  const fileName = `CleanTrack_AdHocJobs_${safeMonth}`;
  XLSX.writeFile(workbook, `${fileName}.xlsx`);
};

/**
 * @param siteNotesLookup Optional maps from CleanTrack Timesheet Period Notes (by site id and by site name).
 */
export const exportFortnightTimesheets = (
  period: FortnightPeriod,
  sites: Site[],
  cleaners: Cleaner[],
  entries: TimeEntry[],
  formatType: 'xlsx' | 'csv' = 'xlsx',
  siteNotesLookup?: SiteNotesExportLookup,
  adHocJobs: AdHocJob[] = []
) => {
  const dates: Date[] = [];
  for (let i = 0; i < 14; i++) {
    dates.push(addDays(period.startDate, i));
  }

  // Used for PH ($/hr) selection. Other rate types are purely weekday-based.
  const phInPeriod = getPublicHolidaysInRange(period.startDate, period.endDate);

  const dateHeaders = dates.map(d => format(d, 'EEE d MMM'));
  const headers = [
    'Site',
    'Site Address',
    'Cleaner Name',
    ...dateHeaders,
    'Total Hours',
    'Contract Hours',
    'Ad Hoc Hours',
    'Ad Hoc Jobs',
    'Hourly rate',
    'Total Payment',
    'Account Name',
    'BSB',
    'Account Number',
    'Manager note',
  ];

  const rows: any[][] = [];
  const periodStartStr = format(period.startDate, 'yyyy-MM-dd');
  const periodEndStr = format(period.endDate, 'yyyy-MM-dd');
  const periodEntries = entries.filter(
    (e) =>
      !!e.cleanerId &&
      e.date >= periodStartStr &&
      e.date <= periodEndStr
  );
  const siteById = new Map(sites.map((s) => [s.id, s] as const));
  const cleanerById = new Map(cleaners.map((c) => [c.id, c] as const));
  const adHocById = new Map(adHocJobs.map((j) => [String(j.id), j] as const));
  const grouped = new Map<string, TimeEntry[]>();
  periodEntries.forEach((e) => {
    const siteKey = e.siteId ? String(e.siteId) : '__NO_SITE__';
    // Important: always separate Ad Hoc jobs into their own rows, even when the Ad Hoc entry is linked to a Site.
    // Otherwise Ad Hoc entries can merge into the same row as contract/standard work (and rates differ).
    const adhocKey = e.adhocJobId ? String(e.adhocJobId) : '__CONTRACT__';
    const key = `${siteKey}|${String(e.cleanerId)}|${adhocKey}`;
    if (!grouped.has(key)) grouped.set(key, []);
    grouped.get(key)!.push(e);
  });

  const groupKeys = Array.from(grouped.keys()).sort((a, b) => a.localeCompare(b));
  groupKeys.forEach((key) => {
    const group = grouped.get(key) ?? [];
    if (group.length === 0) return;
    const [siteId, cleanerId] = key.split('|');
    const cleaner = cleanerById.get(cleanerId);
    if (!cleaner) return;
    const site = siteId && siteId !== '__NO_SITE__' ? siteById.get(siteId) : undefined;
    const firstAdHocId = group.find((e) => !!e.adhocJobId)?.adhocJobId;
    const adHocJob = firstAdHocId ? adHocById.get(String(firstAdHocId)) : undefined;
    const derivedSiteName =
      site?.name ??
      adHocJob?.manualSiteName?.trim() ??
      adHocJob?.siteName?.trim() ??
      adHocJob?.jobName?.trim() ??
      'Ad Hoc (Unlinked Site)';
    const derivedSiteAddress =
      site?.address ??
      [adHocJob?.manualSiteAddress?.trim(), adHocJob?.manualSiteState?.trim()]
        .filter(Boolean)
        .join(', ');

    const weekdayRate = site?.budget_weekday_labour_rate ?? site?.budget_labour_rate ?? 0;
    const saturdayRate = site?.budget_saturday_labour_rate ?? weekdayRate;
    const sundayRate = site?.budget_sunday_labour_rate ?? weekdayRate;
    const phRate = site?.budget_ph_labour_rate ?? weekdayRate;
    const siteStateCode = extractAustralianStateCode(site?.state ?? site?.address);
    const phForSiteInPeriod = getPublicHolidaysInRange(
      period.startDate,
      period.endDate,
      siteStateCode || undefined
    );

    const row: any[] = [
      derivedSiteName,
      derivedSiteAddress,
      `${cleaner.firstName} ${cleaner.lastName}`
    ];

    let totalHours = 0;
    let contractHours = 0;
    let adHocHours = 0;
    let totalPayment = 0;
    const adHocJobNames = new Set<string>();
    const adHocJobIdsForRow = new Set<string>();

    dates.forEach((date) => {
      const dateStr = format(date, 'yyyy-MM-dd');
      const dayEntries = group.filter((e) => e.date === dateStr);
      const dayHours = dayEntries.reduce((sum, e) => sum + e.hours, 0);

      row.push(dayHours > 0 ? dayHours : '');
      totalHours += dayHours;

      dayEntries.forEach((e) => {
        if (e.adhocJobId) {
          adHocHours += e.hours;
          adHocJobIdsForRow.add(String(e.adhocJobId));
          const label =
            e.adhocJobName?.trim() ||
            (e.adhocJobId ? `Ad Hoc Job #${e.adhocJobId}` : '');
          if (label) adHocJobNames.add(label);
        } else {
          contractHours += e.hours;
        }
        // Payment is driven by the same logic as the UI:
        // - Contract/standard work: use site budget $/hr (weekday / Sat / Sun / PH)
        // - Ad Hoc: use the Ad Hoc job's cost $/hr (weekday / Sat / Sun / PH), which should match Est. Pay.
        const rate = e.adhocJobId
          ? (() => {
              const job = adHocById.get(String(e.adhocJobId));
              return job ? getAdHocCostRateForDate(job, date, phInPeriod) : 0;
            })()
          : getSiteRateForDate(date, weekdayRate, saturdayRate, sundayRate, phRate, phForSiteInPeriod);
        totalPayment += e.hours * rate;
      });
    });

    row.push(totalHours);
    row.push(contractHours > 0 ? contractHours : '');
    row.push(adHocHours > 0 ? adHocHours : '');
    row.push(adHocJobNames.size ? Array.from(adHocJobNames).join(', ') : '');
    row.push(totalHours > 0 ? `$${(totalPayment / totalHours).toFixed(2)}` : '');
    row.push(`$${totalPayment.toFixed(2)}`);
    row.push(cleaner.bankAccountName || '');
    row.push(cleaner.bankBsb || '');
    row.push(cleaner.bankAccountNumber || '');
    const adHocNameHints = new Set<string>(adHocJobNames);
    for (const id of Array.from(adHocJobIdsForRow)) {
      const job = adHocById.get(String(id));
      if (!job) continue;
      if (job.jobName?.trim()) adHocNameHints.add(job.jobName.trim());
      if (job.siteName?.trim()) adHocNameHints.add(job.siteName.trim());
      if (job.manualSiteName?.trim()) adHocNameHints.add(job.manualSiteName.trim());
    }
    row.push(
      managerNoteForExport(
        siteNotesLookup,
        cleanerId,
        site,
        derivedSiteName,
        Array.from(adHocJobIdsForRow),
        Array.from(adHocNameHints)
      )
    );
    rows.push(row);
  });

  if (rows.length === 0) {
    alert("No data found for the selected fortnight.");
    return;
  }

  const aoa = [headers, ...rows];
  const worksheet = XLSX.utils.aoa_to_sheet(aoa);
  const noteColIdx = headers.length - 1;

  if (formatType === 'xlsx') {
    applyNoteColumnHighlight(worksheet, noteColIdx, aoa.length);
    const cols = worksheet['!cols'] ?? [];
    cols[noteColIdx] = { wch: 48 };
    worksheet['!cols'] = cols;
  }

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Timesheets");

  const fileName = `CleanTrack_Timesheets_${format(period.startDate, 'yyyy-MM-dd')}_to_${format(period.endDate, 'yyyy-MM-dd')}`;

  if (formatType === 'xlsx') {
    XLSX.writeFile(workbook, `${fileName}.xlsx`);
  } else {
    XLSX.writeFile(workbook, `${fileName}.csv`, { bookType: 'csv' });
  }
};
