import { format, startOfYear, addDays, differenceInDays, startOfDay, endOfDay, isWithinInterval, endOfMonth, getDay } from 'date-fns';

// Reference start date for fortnight cycles: first Monday of 2024
const REFERENCE_START = new Date(2024, 0, 1);

export const getFortnightForDate = (date: Date) => {
  const diff = differenceInDays(startOfDay(date), startOfDay(REFERENCE_START));
  const fortnightIndex = Math.floor(diff / 14);
  
  const startDate = addDays(REFERENCE_START, fortnightIndex * 14);
  const endDate = addDays(startDate, 13);
  
  return {
    startDate,
    endDate,
    id: `${format(startDate, 'yyyy-MM-dd')}_${format(endDate, 'yyyy-MM-dd')}`
  };
};

/**
 * Default fortnight used for timesheet completion.
 *
 * Operationally, timesheets are completed on the Wednesday after the fortnight ends.
 * That means during the first week of a new fortnight (Mon–Sun, i.e. until its second-week Monday),
 * to the previous fortnight to prevent accidental entry into the not-yet-completed one.
 */
export const getFortnightForTimesheetCompletion = (date: Date) => {
  const current = getFortnightForDate(date);
  const today = startOfDay(date);

  // Cutover from the new fortnight happens on Monday of its second week (startDate + 7).
  const cutoff = addDays(startOfDay(current.startDate), 7);
  if (today < cutoff) {
    // Previous fortnight ends the day before the current fortnight starts.
    return getFortnightForDate(addDays(current.startDate, -1));
  }
  return current;
};

/**
 * Viewing lock rule:
 * - A fortnight is editable/viewable starting from its "second week" (startDate + 7 days).
 * - Fully completed fortnights remain unlocked.
 */
export const isFortnightLockedForView = (now: Date, startDate: Date, endDate: Date): boolean => {
  const n = startOfDay(now);
  const start = startOfDay(startDate);
  const end = startOfDay(endDate);

  // If the fortnight has ended, always allow view.
  if (n > end) return false;

  const unlockAt = addDays(start, 7);
  return n < unlockAt;
};

export const getAllFortnightsInRange = (start: Date, end: Date) => {
  const periods = [];
  let current = getFortnightForDate(start).startDate;
  
  while (current <= end) {
    periods.push(getFortnightForDate(current));
    current = addDays(current, 14);
  }
  return periods;
};

export const getStatusColor = (current: number, budget: number) => {
  if (current > budget) return 'text-red-600';
  if (current < budget * 0.8) return 'text-amber-500';
  return 'text-green-600';
};

export const getStatusBg = (current: number, budget: number) => {
  if (current > budget) return 'bg-red-50 text-red-700 border-red-200';
  if (current < budget * 0.8) return 'bg-amber-50 text-amber-700 border-amber-200';
  return 'bg-green-50 text-green-700 border-green-200';
};

/**
 * nth weekday within a given month (0=Sun..6=Sat).
 * Example: 3rd Thursday => which="Third", weekday=4.
 */
export function nthWeekdayOfMonth(
  year: number,
  monthIndex0: number,
  weekday: number,
  which: string
): Date | null {
  const first = new Date(year, monthIndex0, 1);
  const last = endOfMonth(first);
  const whichLower = String(which ?? '').toLowerCase();
  if (whichLower === "last") {
    // Walk backwards from last day until weekday matches.
    let d = startOfDay(last);
    while (getDay(d) !== weekday) d = addDays(d, -1);
    return d;
  }
  const order = ["first", "second", "third", "fourth"];
  const idx = order.indexOf(whichLower);
  if (idx < 0) return null;
  // Find first occurrence of weekday in month.
  let d = startOfDay(first);
  while (getDay(d) !== weekday) d = addDays(d, 1);
  d = addDays(d, idx * 7);
  // If overflowed to next month, invalid.
  if (d.getMonth() !== monthIndex0) return null;
  return d;
}

/** Round a number to 2 decimal places for currency storage (avoids floating-point drift). */
export function roundCurrency(value: number): number {
  return Math.round(value * 100) / 100;
}

/** Format number as AUD currency. Handles negatives with minus sign. Always 2 decimal places. */
export function formatCurrencyAUD(value: number): string {
  const abs = Math.abs(value);
  const formatted = abs.toLocaleString("en-AU", {
    style: "currency",
    currency: "AUD",
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
  return value < 0 ? `-${formatted.replace(/^-/, "")}` : formatted;
}

/** Format 0..1 as percentage with one decimal. */
export function formatPercent(value: number): string {
  const pct = value * 100;
  return `${pct.toFixed(1)}%`;
}

/** Format positive expense as negative currency for display (e.g. labor). */
export function formatCurrencyAUDSignedExpense(value: number): string {
  const abs = Math.abs(value);
  const formatted = abs.toLocaleString("en-AU", {
    style: "currency",
    currency: "AUD",
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
  return value <= 0 ? formatted : `-${formatted.replace(/^-/, "")}`;
}

/**
 * Standard Operational Audit Logic for Shift Compliance
 */
export const getDayStatus = (planned: number, actual: number) => {
  // Unplanned: Work done on a day where no hours were budgeted
  if (actual > 0 && planned === 0) {
    return { label: 'Unplanned', color: 'text-amber-600', bg: 'bg-amber-50', border: 'border-amber-200', dot: 'bg-amber-400' };
  }
  // Missing: No work logged on a day where hours were budgeted
  if (planned > 0 && actual === 0) {
    return { label: 'Missing', color: 'text-gray-500', bg: 'bg-gray-50', border: 'border-gray-200', dot: 'bg-gray-300' };
  }
  // On target: Logged hours match planned hours within 0.1h tolerance
  if (Math.abs(actual - planned) <= 0.1) {
    return { label: 'On target', color: 'text-green-600', bg: 'bg-green-50', border: 'border-green-200', dot: 'bg-green-500' };
  }
  // Over: Logged hours exceed planned hours
  if (actual > planned + 0.1) {
    return { label: 'Over', color: 'text-red-600', bg: 'bg-red-50', border: 'border-red-200', dot: 'bg-red-500' };
  }
  // Under: Logged hours are less than planned hours
  if (planned > 0 && actual < planned - 0.1) {
    return { label: 'Under', color: 'text-blue-600', bg: 'bg-blue-50', border: 'border-blue-200', dot: 'bg-blue-500' };
  }
  return { label: 'Unknown', color: 'text-gray-400', bg: 'bg-white', border: 'border-gray-100', dot: 'bg-gray-100' };
};