/**
 * Embedded state-school holiday breaks (Australia) for offline / provider-free operation.
 * Dates follow typical state education department term breaks; verify annually against official calendars.
 * @see https://education.qld.gov.au/about-us/calendar/term-dates (QLD)
 */

export type SchoolBreak = { label: string; start: string; end: string };

/** QLD state schools — term breaks only (not pupil-free days). */
export const QLD_SCHOOL_BREAKS_BY_YEAR: Record<number, SchoolBreak[]> = {
  2024: [
    { label: "Term 1 holidays", start: "2024-03-30", end: "2024-04-14" },
    { label: "Term 2 holidays", start: "2024-06-29", end: "2024-07-14" },
    { label: "Term 3 holidays", start: "2024-09-21", end: "2024-10-06" },
    { label: "Summer holidays", start: "2024-12-14", end: "2025-01-26" },
  ],
  2025: [
    { label: "Term 1 holidays", start: "2025-04-05", end: "2025-04-20" },
    { label: "Term 2 holidays", start: "2025-06-28", end: "2025-07-13" },
    { label: "Term 3 holidays", start: "2025-09-20", end: "2025-10-05" },
    { label: "Summer holidays", start: "2025-12-13", end: "2026-01-25" },
  ],
  2026: [
    { label: "Term 1 holidays", start: "2026-04-04", end: "2026-04-19" },
    { label: "Term 2 holidays", start: "2026-06-27", end: "2026-07-12" },
    { label: "Term 3 holidays", start: "2026-09-19", end: "2026-10-04" },
    { label: "Summer holidays", start: "2026-12-12", end: "2027-01-24" },
  ],
  2027: [
    { label: "Term 1 holidays", start: "2027-04-03", end: "2027-04-18" },
    { label: "Term 2 holidays", start: "2027-06-26", end: "2027-07-11" },
    { label: "Term 3 holidays", start: "2027-09-18", end: "2027-10-03" },
    { label: "Summer holidays", start: "2027-12-11", end: "2028-01-23" },
  ],
  2028: [
    { label: "Term 1 holidays", start: "2028-04-01", end: "2028-04-16" },
    { label: "Term 2 holidays", start: "2028-06-24", end: "2028-07-09" },
    { label: "Term 3 holidays", start: "2028-09-16", end: "2028-10-01" },
    { label: "Summer holidays", start: "2028-12-09", end: "2029-01-21" },
  ],
};

/** NSW public schools — indicative term breaks (verify against NSW DoE each year). */
export const NSW_SCHOOL_BREAKS_BY_YEAR: Record<number, SchoolBreak[]> = {
  2024: [
    { label: "Term 1 holidays", start: "2024-04-13", end: "2024-04-28" },
    { label: "Term 2 holidays", start: "2024-07-08", end: "2024-07-21" },
    { label: "Term 3 holidays", start: "2024-09-30", end: "2024-10-11" },
    { label: "Summer holidays", start: "2024-12-21", end: "2025-02-03" },
  ],
  2025: [
    { label: "Term 1 holidays", start: "2025-04-12", end: "2025-04-27" },
    { label: "Term 2 holidays", start: "2025-07-07", end: "2025-07-20" },
    { label: "Term 3 holidays", start: "2025-09-29", end: "2025-10-10" },
    { label: "Summer holidays", start: "2025-12-20", end: "2026-02-02" },
  ],
  2026: [
    { label: "Term 1 holidays", start: "2026-04-11", end: "2026-04-26" },
    { label: "Term 2 holidays", start: "2026-07-06", end: "2026-07-19" },
    { label: "Term 3 holidays", start: "2026-09-28", end: "2026-10-09" },
    { label: "Summer holidays", start: "2026-12-19", end: "2027-02-01" },
  ],
  2027: [
    { label: "Term 1 holidays", start: "2027-04-10", end: "2027-04-25" },
    { label: "Term 2 holidays", start: "2027-07-05", end: "2027-07-18" },
    { label: "Term 3 holidays", start: "2027-09-27", end: "2027-10-08" },
    { label: "Summer holidays", start: "2027-12-18", end: "2028-01-31" },
  ],
  2028: [
    { label: "Term 1 holidays", start: "2028-04-08", end: "2028-04-23" },
    { label: "Term 2 holidays", start: "2028-07-03", end: "2028-07-16" },
    { label: "Term 3 holidays", start: "2028-09-25", end: "2028-10-06" },
    { label: "Summer holidays", start: "2028-12-16", end: "2029-01-29" },
  ],
};
