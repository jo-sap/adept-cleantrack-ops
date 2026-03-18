
import * as XLSX from 'xlsx';
import { format, addDays } from 'date-fns';
import { Site, Cleaner, TimeEntry, FortnightPeriod, AdHocJob } from '../types';

/** Export current Ad Hoc jobs to XLSX. Uses the same job list as the UI (respects month/status/manager/site filters). */
export const exportAdHocJobsToSpreadsheet = (
  jobs: AdHocJob[],
  monthLabel: string
) => {
  const headers = [
    'Job Name',
    'Schedule Type',
    'Recurrence',
    'Company',
    'Client',
    'Site',
    'Assigned Manager',
    'Requested',
    'Scheduled',
    'Completed',
    'Status',
    'Budgeted Hrs',
    'Charge',
    'Cost',
    'Gross Profit',
    'Approval Proof',
    'Requested By',
    'Requested By Email',
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

  const rows: any[][] = jobs.map((j) => [
    j.jobName ?? '',
    scheduleTypeLabel(j.jobType),
    recurrenceSummary(j),
    j.companyName ?? '',
    j.clientName ?? '',
    j.siteName || j.manualSiteName || '',
    j.assignedManagerName ?? '',
    j.requestedDate ? format(new Date(j.requestedDate), 'dd MMM yyyy') : '',
    j.scheduledDate ? format(new Date(j.scheduledDate), 'dd MMM yyyy') : '',
    j.completedDate ? format(new Date(j.completedDate), 'dd MMM yyyy') : '',
    j.status ?? '',
    j.budgetedHours != null ? j.budgetedHours : '',
    j.charge != null ? Number(j.charge) : '',
    j.cost != null ? Number(j.cost) : '',
    j.grossProfit != null ? Number(j.grossProfit) : '',
    j.approvalProofUploaded ? 'Yes' : 'No',
    j.requestedByName ?? '',
    j.requestedByEmail ?? '',
    j.description ?? '',
  ]);

  if (rows.length === 0) {
    alert('No ad hoc jobs to export for the selected filters.');
    return;
  }

  const worksheet = XLSX.utils.aoa_to_sheet([headers, ...rows]);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Ad Hoc Jobs');

  const [y, m] = monthLabel.split('-').map(Number);
  const date = new Date(y, (m || 1) - 1, 1);
  const safeMonth = format(date, 'MMMM_yyyy').replace(/\s+/g, '_');
  const fileName = `CleanTrack_AdHocJobs_${safeMonth}`;
  XLSX.writeFile(workbook, `${fileName}.xlsx`);
};

export const exportFortnightTimesheets = (
  period: FortnightPeriod,
  sites: Site[],
  cleaners: Cleaner[],
  entries: TimeEntry[],
  formatType: 'xlsx' | 'csv' = 'xlsx'
) => {
  const dates: Date[] = [];
  for (let i = 0; i < 14; i++) {
    dates.push(addDays(period.startDate, i));
  }

  const dateHeaders = dates.map(d => format(d, 'EEE d MMM'));
  const headers = [
    'Site',
    'Site Address',
    'Cleaner Name',
    ...dateHeaders,
    'Total Hours',
    'Total Payment',
    'Account Name',
    'BSB',
    'Account Number'
  ];

  const rows: any[][] = [];

  // Iterate through all sites and assigned cleaners
  sites.forEach(site => {
    // Correctly reference snake_case property
    site.assigned_cleaner_ids.forEach(cleanerId => {
      const cleaner = cleaners.find(c => c.id === cleanerId);
      if (!cleaner) return;

      // siteId and cleanerId are augmented onto TimeEntry in App.tsx and types.ts
      const cleanerEntries = entries.filter(e => e.siteId === site.id && e.cleanerId === cleanerId);
      
      // Skip if no work logged in this fortnight at this site
      if (cleanerEntries.length === 0) return;

      const row: any[] = [
        site.name,
        site.address,
        `${cleaner.firstName} ${cleaner.lastName}`
      ];

      let totalHours = 0;
      let totalPayment = 0;

      // Add daily hours
      dates.forEach(date => {
        const dateStr = format(date, 'yyyy-MM-dd');
        const dayEntries = cleanerEntries.filter(e => e.date === dateStr);
        const dayHours = dayEntries.reduce((sum, e) => sum + e.hours, 0);
        
        row.push(dayHours > 0 ? dayHours : '');
        totalHours += dayHours;

        // Calculate payment for this specific day using snapshot or fallback
        dayEntries.forEach(e => {
          // Use correct snake_case properties
          const rate = e.pay_rate_snapshot || site.cleaner_rates[cleanerId] || cleaner.payRatePerHour || 0;
          totalPayment += e.hours * rate;
        });
      });

      row.push(totalHours);
      row.push(`$${totalPayment.toFixed(2)}`);
      row.push(cleaner.bankAccountName || '');
      row.push(cleaner.bankBsb || '');
      row.push(cleaner.bankAccountNumber || '');

      rows.push(row);
    });
  });

  if (rows.length === 0) {
    alert("No data found for the selected fortnight.");
    return;
  }

  const worksheet = XLSX.utils.aoa_to_sheet([headers, ...rows]);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Timesheets");

  const fileName = `CleanTrack_Timesheets_${format(period.startDate, 'yyyy-MM-dd')}_to_${format(period.endDate, 'yyyy-MM-dd')}`;

  if (formatType === 'xlsx') {
    XLSX.writeFile(workbook, `${fileName}.xlsx`);
  } else {
    XLSX.writeFile(workbook, `${fileName}.csv`, { bookType: 'csv' });
  }
};