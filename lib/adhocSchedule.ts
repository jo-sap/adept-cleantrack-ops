import {
  addDays,
  differenceInCalendarWeeks,
  endOfMonth,
  getDay,
  startOfDay,
} from "date-fns";
import type { AdHocJob } from "../types";

export type AdHocDayType = "public_holiday" | "saturday" | "sunday" | "weekday";

export type AdHocOccurrence = {
  adhocJobId: string;
  date: string; // yyyy-MM-dd
  hours: number;
  dayType: AdHocDayType;
  chargeRate: number;
  costRate: number;
  chargeTotal: number;
  costTotal: number;
};

function toDateKey(d: Date): string {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
}

function normalizeScheduleType(raw: string | undefined | null): "once_off" | "recurring" {
  const s = String(raw ?? "").trim().toLowerCase();
  if (s.includes("recurr")) return "recurring";
  return "once_off";
}

function clampRange(start: Date, end: Date, min: Date, max: Date): { start: Date; end: Date } | null {
  const s = start > min ? start : min;
  const e = end < max ? end : max;
  if (s > e) return null;
  return { start: s, end: e };
}

function nthWeekdayOfMonth(year: number, monthIndex0: number, weekday: number, which: string): Date | null {
  // weekday: 0=Sun..6=Sat
  const first = new Date(year, monthIndex0, 1);
  const last = endOfMonth(first);
  const whichLower = which.toLowerCase();
  if (whichLower === "last") {
    // walk backward from last day to the weekday
    let d = startOfDay(last);
    while (getDay(d) !== weekday) d = addDays(d, -1);
    return d;
  }
  const order = ["first", "second", "third", "fourth"];
  const idx = order.indexOf(whichLower);
  if (idx < 0) return null;
  // find first occurrence of weekday in month
  let d = startOfDay(first);
  while (getDay(d) !== weekday) d = addDays(d, 1);
  d = addDays(d, idx * 7);
  // if overflowed to next month, invalid
  if (d.getMonth() !== monthIndex0) return null;
  return d;
}

function monthIndexOf(d: Date): number {
  return d.getFullYear() * 12 + d.getMonth();
}

function getDayType(date: Date, publicHolidayDates?: Set<string>): AdHocDayType {
  const key = toDateKey(date);
  if (publicHolidayDates?.has(key)) return "public_holiday";
  const dow = getDay(date);
  if (dow === 6) return "saturday";
  if (dow === 0) return "sunday";
  return "weekday";
}

function getApplicableRate(
  job: AdHocJob,
  date: Date,
  kind: "charge" | "cost",
  publicHolidayDates?: Set<string>
): number {
  const base = kind === "charge" ? (job.chargeRatePerHour ?? 0) : (job.costRatePerHour ?? 0);
  const dayType = getDayType(date, publicHolidayDates);
  if (kind === "charge") {
    if (dayType === "public_holiday" && job.publicHolidayChargeRateOverride != null) return job.publicHolidayChargeRateOverride;
    if (dayType === "saturday" && job.saturdayChargeRateOverride != null) return job.saturdayChargeRateOverride;
    if (dayType === "sunday" && job.sundayChargeRateOverride != null) return job.sundayChargeRateOverride;
    if (dayType === "weekday" && job.weekdayChargeRateOverride != null) return job.weekdayChargeRateOverride;
    return base;
  }
  if (dayType === "public_holiday" && job.publicHolidayCostRateOverride != null) return job.publicHolidayCostRateOverride;
  if (dayType === "saturday" && job.saturdayCostRateOverride != null) return job.saturdayCostRateOverride;
  if (dayType === "sunday" && job.sundayCostRateOverride != null) return job.sundayCostRateOverride;
  if (dayType === "weekday" && job.weekdayCostRateOverride != null) return job.weekdayCostRateOverride;
  return base;
}

function deriveWeekdayHours(job: AdHocJob): Record<number, number> {
  const out: Record<number, number> = {};
  // Preferred: per-weekday hours map.
  const raw = job.weekdayHours ?? null;
  if (raw) {
    for (const [k, v] of Object.entries(raw)) {
      const idx = parseInt(k, 10);
      const n = typeof v === "number" ? v : parseFloat(String(v));
      if (Number.isFinite(idx) && idx >= 0 && idx <= 6 && Number.isFinite(n) && n > 0) out[idx] = n;
    }
  }
  // Back-compat: selected weekdays + single hours per day.
  if (Object.keys(out).length === 0 && (job.recurrenceWeekdays?.length ?? 0) > 0) {
    const hours = job.hoursPerServiceDay ?? 0;
    for (const d of job.recurrenceWeekdays ?? []) {
      if (hours > 0) out[d] = hours;
    }
  }
  return out;
}

function normalizeRecurrenceFrequency(
  raw: string | null | undefined
): "Weekly" | "Fortnightly" | "Monthly" | "Quarterly" | "Half Yearly" | "Annually" | null {
  const v = String(raw ?? "").trim().toLowerCase();
  if (!v) return null;
  if (v.startsWith("week")) return "Weekly";
  if (v.startsWith("fort")) return "Fortnightly";
  if (v.startsWith("quart")) return "Quarterly";
  if (v.startsWith("half")) return "Half Yearly";
  if (v.startsWith("ann")) return "Annually";
  if (v.startsWith("month")) return "Monthly";
  return null;
}

function normalizeMonthlyMode(raw: string | null | undefined): "day_of_month" | "nth_weekday" | null {
  const v = String(raw ?? "").trim().toLowerCase().replace(/[\s-]+/g, "_");
  if (!v) return null;
  if (v === "day_of_month" || v === "dayofmonth") return "day_of_month";
  if (v === "nth_weekday" || v === "nthweekday") return "nth_weekday";
  return null;
}

export function generateAdHocOccurrencesForRange(
  job: AdHocJob,
  rangeStart: Date,
  rangeEnd: Date,
  publicHolidayDates?: Set<string>
): AdHocOccurrence[] {
  if (!job.active) return [];
  const status = String(job.status ?? "").trim().toLowerCase();
  if (status === "cancelled") return [];

  const scheduleType = normalizeScheduleType(job.jobType);
  const chargeRate = job.chargeRatePerHour ?? 0;
  const costRate = job.costRatePerHour ?? 0;

  const start = startOfDay(rangeStart);
  const end = startOfDay(rangeEnd);

  if (scheduleType === "once_off") {
    const sched = job.scheduledDate ? new Date(job.scheduledDate) : null;
    if (!sched || isNaN(sched.getTime())) return [];
    const d = startOfDay(sched);
    if (d < start || d > end) return [];
    const hours = job.budgetedHours ?? 0;
    const dayType = getDayType(d, publicHolidayDates);
    const chargeRateApplied = getApplicableRate(job, d, "charge", publicHolidayDates);
    const costRateApplied = getApplicableRate(job, d, "cost", publicHolidayDates);
    return [
      {
        adhocJobId: job.id,
        date: toDateKey(d),
        hours,
        dayType,
        chargeRate: chargeRateApplied,
        costRate: costRateApplied,
        chargeTotal: hours * chargeRateApplied,
        costTotal: hours * costRateApplied,
      },
    ];
  }

  // recurring
  // Start Date is the recurrence anchor (we reuse `scheduledDate` as the Start Date).
  // Never generate occurrences before Start Date.
  const anchor = job.scheduledDate ? new Date(job.scheduledDate) : null;
  if (!anchor || isNaN(anchor.getTime())) return [];
  const recurrenceStart = startOfDay(anchor);

  // Respect End Date if present; additionally, if status is Completed, do not generate future occurrences
  // after the completed date (but keep historical occurrences when viewing past periods).
  const explicitEnd = job.recurrenceEndDate ? startOfDay(new Date(job.recurrenceEndDate)) : null;
  const completedEnd = job.completedDate ? startOfDay(new Date(job.completedDate)) : null;
  const recurrenceEnd =
    explicitEnd && completedEnd
      ? (explicitEnd < completedEnd ? explicitEnd : completedEnd)
      : (explicitEnd ?? completedEnd);

  const effective = clampRange(start, end, recurrenceStart, recurrenceEnd ?? end);
  if (!effective) return [];

  const weekdayHours = deriveWeekdayHours(job);
  const occurrences: AdHocOccurrence[] = [];

  const freq = normalizeRecurrenceFrequency(job.recurrenceFrequency);
  if (!freq) return [];

  if (freq === "Weekly" || freq === "Fortnightly") {
    if (Object.keys(weekdayHours).length === 0) return [];

    // iterate each day in clamped range
    let d = new Date(effective.start);
    while (d <= effective.end) {
      const dow = getDay(d);
      const hours = weekdayHours[dow] ?? 0;
      if (hours > 0) {
        let ok = true;
        if (freq === "Fortnightly") {
          // anchor week parity from recurrenceStart
          const weeks = differenceInCalendarWeeks(d, recurrenceStart, { weekStartsOn: 1 });
          ok = weeks % 2 === 0;
        }
        if (ok) {
          const dayType = getDayType(d, publicHolidayDates);
          const chargeRateApplied = getApplicableRate(job, d, "charge", publicHolidayDates);
          const costRateApplied = getApplicableRate(job, d, "cost", publicHolidayDates);
          occurrences.push({
            adhocJobId: job.id,
            date: toDateKey(d),
            hours,
            dayType,
            chargeRate: chargeRateApplied,
            costRate: costRateApplied,
            chargeTotal: hours * chargeRateApplied,
            costTotal: hours * costRateApplied,
          });
        }
      }
      d = addDays(d, 1);
    }
    return occurrences;
  }

  // Monthly-family cadence (monthly, quarterly, half-yearly, annually)
  if (
    freq === "Monthly" ||
    freq === "Quarterly" ||
    freq === "Half Yearly" ||
    freq === "Annually"
  ) {
    const mode = normalizeMonthlyMode(job.monthlyMode);
    if (!mode) return [];
    // Walk month by month across range.
    let cursor = new Date(effective.start.getFullYear(), effective.start.getMonth(), 1);
    const lastMonth = new Date(effective.end.getFullYear(), effective.end.getMonth(), 1);
    const anchorMonthIdx = monthIndexOf(recurrenceStart);
    while (cursor <= lastMonth) {
      let monthStep = 1;
      if (freq === "Quarterly") monthStep = 3;
      else if (freq === "Half Yearly") monthStep = 6;
      else if (freq === "Annually") monthStep = 12;

      if (monthStep > 1) {
        const diffMonths = monthIndexOf(cursor) - anchorMonthIdx;
        if (diffMonths < 0 || diffMonths % monthStep !== 0) {
          cursor = new Date(cursor.getFullYear(), cursor.getMonth() + 1, 1);
          continue;
        }
      }
      let occ: Date | null = null;
      if (mode === "day_of_month") {
        const dom = job.monthlyDayOfMonth ?? null;
        if (dom == null) return [];
        // If month doesn't have that day, skip.
        const candidate = new Date(cursor.getFullYear(), cursor.getMonth(), dom);
        if (candidate.getMonth() === cursor.getMonth()) occ = startOfDay(candidate);
      } else {
        const which = job.monthlyWeekOfMonth ?? null;
        const wd = job.monthlyWeekday ?? null;
        if (!which || wd == null) return [];
        occ = nthWeekdayOfMonth(cursor.getFullYear(), cursor.getMonth(), wd, which);
      }
      if (occ && occ >= effective.start && occ <= effective.end) {
        const hours = job.monthlyHours ?? job.hoursPerServiceDay ?? 0;
        const dayType = getDayType(occ, publicHolidayDates);
        const chargeRateApplied = getApplicableRate(job, occ, "charge", publicHolidayDates);
        const costRateApplied = getApplicableRate(job, occ, "cost", publicHolidayDates);
        occurrences.push({
          adhocJobId: job.id,
          date: toDateKey(occ),
          hours,
          dayType,
          chargeRate: chargeRateApplied,
          costRate: costRateApplied,
          chargeTotal: hours * chargeRateApplied,
          costTotal: hours * costRateApplied,
        });
      }
      cursor = new Date(cursor.getFullYear(), cursor.getMonth() + 1, 1);
    }
    return occurrences;
  }

  return [];
}

export function occurrencesToHoursByDate(occ: AdHocOccurrence[]): Record<string, number> {
  const map: Record<string, number> = {};
  for (const o of occ) {
    map[o.date] = (map[o.date] ?? 0) + o.hours;
  }
  return map;
}

/** True if the job has at least one planned occurrence in the inclusive date range (once-off or recurring). */
export function adHocJobHasPlannedWorkInRange(
  job: AdHocJob,
  rangeStart: Date,
  rangeEnd: Date,
  publicHolidayDates?: Set<string>
): boolean {
  return (
    generateAdHocOccurrencesForRange(job, rangeStart, rangeEnd, publicHolidayDates).length > 0
  );
}

