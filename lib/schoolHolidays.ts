import type { SiteNoServicePeriod } from "../repositories/sitesRepo";
import { NSW_SCHOOL_BREAKS_BY_YEAR, QLD_SCHOOL_BREAKS_BY_YEAR, type SchoolBreak } from "./schoolHolidaysData";

export type SchoolHolidayPeriodsResult =
  | { ok: true; periods: SiteNoServicePeriod[] }
  | { ok: false; error: string };

const REASON = "School holidays";

function normalizeState(state: string): string {
  return String(state ?? "")
    .trim()
    .toUpperCase();
}

function breaksForStateYear(state: string, year: number): SchoolBreak[] | null {
  if (state === "QLD") {
    const rows = QLD_SCHOOL_BREAKS_BY_YEAR[year];
    return rows ?? null;
  }
  if (state === "NSW") {
    const rows = NSW_SCHOOL_BREAKS_BY_YEAR[year];
    return rows ?? null;
  }
  return null;
}

/**
 * Returns no-service period objects for the given Australian state and calendar year.
 * Replace or extend `schoolHolidaysData.ts` or swap this function to call an external API.
 */
export function getSchoolHolidayPeriods(stateRaw: string, year: number): SchoolHolidayPeriodsResult {
  const state = normalizeState(stateRaw);
  if (!state) {
    return { ok: false, error: "Site State is required to populate school holidays." };
  }
  if (!Number.isFinite(year) || year < 2000 || year > 2100) {
    return { ok: false, error: "Choose a valid year." };
  }

  const breaks = breaksForStateYear(state, year);
  if (!breaks) {
    return {
      ok: false,
      error: `School holiday auto is not available for "${state}" yet. Use Manual mode or add calendars in schoolHolidaysData.ts. Supported: QLD, NSW.`,
    };
  }

  const periods: SiteNoServicePeriod[] = breaks.map((b) => ({
    label: b.label,
    start_date: b.start,
    end_date: b.end,
    reason: REASON,
    source: "school_holidays_auto",
    state,
    year,
  }));

  return { ok: true, periods };
}

/** For UI: years that have embedded data for at least one supported state. */
export function getSupportedSchoolHolidayYears(): number[] {
  return Object.keys(QLD_SCHOOL_BREAKS_BY_YEAR)
    .map(Number)
    .sort((a, b) => a - b);
}
