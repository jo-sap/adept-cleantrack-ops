/**
 * Date-aware budgeted labour cost: uses Public Holiday, Saturday, Sunday, and Weekday rates
 * for each day in the period so the budget matches actual rate rules.
 */

import { addDays, startOfDay } from "date-fns";
import { getPublicHolidaysInRange } from "./publicHolidays";

/** daily_budgets[0]=Sun, [1]=Mon, ..., [6]=Sat (hours per day in the weekly pattern). */
export interface BudgetedLabourCostInput {
  startDate: Date;
  endDate: Date;
  /** [Sun, Mon, Tue, Wed, Thu, Fri, Sat] — hours budgeted for each day of the week. */
  dailyBudgets: number[];
  weekdayRate: number;
  saturdayRate: number;
  sundayRate: number;
  /** Rate for public holidays (applies regardless of day of week). */
  phRate: number;
  /** Optional: set of PH date keys (YYYY-MM-DD). If not provided, uses default NSW calendar. */
  publicHolidayDates?: Set<string>;
}

function toDateKey(d: Date): string {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
}

/**
 * Return the site labour rate ($/hr) for a given date: PH rate if public holiday,
 * otherwise Sunday / Saturday / Weekday rate by day of week. Use for Est. Pay on timesheets.
 */
export function getSiteRateForDate(
  date: Date,
  weekdayRate: number,
  saturdayRate: number,
  sundayRate: number,
  phRate: number,
  publicHolidayDates: Set<string>
): number {
  const key = toDateKey(date);
  if (publicHolidayDates.has(key)) return phRate;
  const day = date.getDay(); // 0 = Sun, 6 = Sat
  if (day === 0) return sundayRate;
  if (day === 6) return saturdayRate;
  return weekdayRate;
}

/**
 * Compute budgeted labour cost for a date range by iterating each day and applying
 * the correct rate (PH, Sunday, Saturday, or Weekday). Uses the site's weekly hour pattern.
 */
export function computeBudgetedLabourCostForRange(input: BudgetedLabourCostInput): number {
  const {
    startDate,
    endDate,
    dailyBudgets,
    weekdayRate,
    saturdayRate,
    sundayRate,
    phRate,
    publicHolidayDates,
  } = input;

  const start = startOfDay(startDate);
  const end = startOfDay(endDate);
  const phSet = publicHolidayDates ?? getPublicHolidaysInRange(start, end);
  const budgets = dailyBudgets.length >= 7 ? dailyBudgets : [0, 0, 0, 0, 0, 0, 0];

  let total = 0;
  let d = new Date(start);

  while (d <= end) {
    const dayOfWeek = d.getDay(); // 0 = Sun, 6 = Sat
    const hours = budgets[dayOfWeek] ?? 0;
    if (hours > 0) {
      const key = toDateKey(d);
      const isPH = phSet.has(key);
      const rate = isPH
        ? phRate
        : dayOfWeek === 0
          ? sundayRate
          : dayOfWeek === 6
            ? saturdayRate
            : weekdayRate;
      total += hours * rate;
    }
    d = addDays(d, 1);
  }

  return total;
}
