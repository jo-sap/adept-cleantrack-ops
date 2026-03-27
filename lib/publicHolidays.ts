/**
 * Public holiday calendar for budgeted labour cost (PH rate).
 * Default: NSW, Australia. Extend PUBLIC_HOLIDAYS or replace with SharePoint/API later.
 */

import { addDays, startOfDay } from "date-fns";

/** ISO date strings (YYYY-MM-DD), grouped by AU state. Extend as needed. */
const PUBLIC_HOLIDAYS_BY_STATE: Record<string, string[]> = {
  NSW: [
    // 2024 NSW
    "2024-01-01", "2024-01-26", "2024-03-29", "2024-03-30", "2024-03-31", "2024-04-01",
    "2024-04-25", "2024-06-10", "2024-08-05", "2024-10-07", "2024-12-25", "2024-12-26",
    // 2025 NSW
    "2025-01-01", "2025-01-27", "2025-04-18", "2025-04-19", "2025-04-20", "2025-04-21",
    "2025-04-25", "2025-06-09", "2025-08-04", "2025-10-06", "2025-12-25", "2025-12-26",
    // 2026 NSW
    "2026-01-01", "2026-01-26", "2026-04-03", "2026-04-04", "2026-04-05", "2026-04-06",
    "2026-04-25", "2026-06-08", "2026-08-03", "2026-10-05", "2026-12-25", "2026-12-26",
    "2026-12-28",
  ],
  VIC: [
    // 2024 VIC
    "2024-01-01", "2024-01-26", "2024-03-11", "2024-03-29", "2024-03-30", "2024-03-31",
    "2024-04-01", "2024-04-25", "2024-06-10", "2024-11-05", "2024-12-25", "2024-12-26",
    // 2025 VIC
    "2025-01-01", "2025-01-27", "2025-03-10", "2025-04-18", "2025-04-19", "2025-04-20",
    "2025-04-21", "2025-04-25", "2025-06-09", "2025-11-04", "2025-12-25", "2025-12-26",
    // 2026 VIC
    "2026-01-01", "2026-01-26", "2026-03-09", "2026-04-03", "2026-04-04", "2026-04-05",
    "2026-04-06", "2026-04-25", "2026-06-08", "2026-11-03", "2026-12-25", "2026-12-26",
    "2026-12-28",
  ],
};

const phSetByState = new Map<string, Set<string>>();

function normalizeState(state: string | null | undefined): string {
  const s = String(state ?? "").trim().toUpperCase();
  return s || "NSW";
}

function getPHSet(state?: string): Set<string> {
  const normalizedState = normalizeState(state);
  const key = PUBLIC_HOLIDAYS_BY_STATE[normalizedState] ? normalizedState : "NSW";
  const cached = phSetByState.get(key);
  if (cached) return cached;
  const set = new Set(PUBLIC_HOLIDAYS_BY_STATE[key] ?? []);
  phSetByState.set(key, set);
  return set;
}

/** Format date as YYYY-MM-DD for lookup. */
function toKey(d: Date): string {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
}

/** Returns true if the given date is a public holiday (default NSW calendar). */
export function isPublicHoliday(date: Date, state?: string): boolean {
  return getPHSet(state).has(toKey(date));
}

/** Returns the set of public holiday date keys (YYYY-MM-DD) within the range (inclusive). */
export function getPublicHolidaysInRange(startDate: Date, endDate: Date, state?: string): Set<string> {
  const set = new Set<string>();
  const ph = getPHSet(state);
  const start = startOfDay(startDate);
  const end = startOfDay(endDate);
  let d = new Date(start);
  while (d <= end) {
    if (ph.has(toKey(d))) set.add(toKey(d));
    d = addDays(d, 1);
  }
  return set;
}

/** Add more dates (e.g. from SharePoint or config). Call with ISO date strings. */
export function addPublicHolidays(dates: string[]): void {
  const set = getPHSet("NSW");
  dates.forEach((k) => set.add(k));
}
