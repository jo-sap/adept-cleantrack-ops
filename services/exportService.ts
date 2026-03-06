
import * as XLSX from 'xlsx';
import { format, addDays } from 'date-fns';
import { Site, Cleaner, TimeEntry, FortnightPeriod } from '../types';

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