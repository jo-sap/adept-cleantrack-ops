
export type Role = 'admin' | 'manager';

export interface Profile {
  id: string;
  full_name: string;
  role: Role;
  region?: string;
}

export interface Site {
  id: string;
  name: string;
  address: string;
  /** Australian state/territory code (e.g. NSW). */
  state?: string;
  is_active: boolean;
  budgeted_hours_per_fortnight: number;
  daily_budgets: number[];
  /** Fortnightly budgets: optional Week 2 day hours (Sun..Sat). */
  daily_budgets_week2?: number[];
  /**
   * Contract Monthly recurrence (mirrors Ad Hoc monthly configuration).
   * When set, the app plans hours on specific monthly dates (e.g. 3rd Thu).
   * When not set, current behavior remains: Monthly uses a period cap (Hours per Visit ÷ 2).
   */
  monthlyMode?: "day_of_month" | "nth_weekday" | null;
  monthlyWeekOfMonth?: "First" | "Second" | "Third" | "Fourth" | "Last" | null;
  monthlyWeekday?: number | null; // 0=Sun..6=Sat
  monthlyDayOfMonth?: number | null; // 1..31

  /**
   * Monthly exception (delta) rule for contract Sites.
   * Adds extra planned hours on a single monthly occurrence date (e.g. +2 on 3rd Thu),
   * on top of the site's normal weekly/fortnightly plan.
   *
   * Enable/disable by setting hours delta <= 0 (or null).
   */
  monthlyExceptionHoursDelta?: number | null;
  monthlyExceptionMode?: "day_of_month" | "nth_weekday" | null;
  monthlyExceptionWeekOfMonth?: "First" | "Second" | "Third" | "Fourth" | "Last" | null;
  monthlyExceptionWeekday?: number | null; // 0=Sun..6=Sat
  monthlyExceptionDayOfMonth?: number | null; // 1..31
  assigned_cleaner_ids: string[];
  monthly_revenue: number;
  financial_budget: number;
  cleaner_rates: Record<string, number>;
  /** Weekday (Mon–Fri) rate ($/hr). From "Weekday Labour Rate" (was Budget Labour Rate). Backward compat. */
  budget_labour_rate?: number;
  /** Labour rates by day type ($/hr). From CleanTrack Site Budgets: Weekday Labour Rate, Saturday Labour Rate, Sunday Labour Rate, PH Labour Rate. */
  budget_weekday_labour_rate?: number;
  budget_saturday_labour_rate?: number;
  budget_sunday_labour_rate?: number;
  budget_ph_labour_rate?: number;
  managers?: Profile[]; // Joined data
  /** Weekly | Fortnightly | Monthly – from site budget */
  visit_frequency?: string;
  /** Site-level date windows where planned service is intentionally paused. */
  no_service_periods?: NoServicePeriod[];
}

export interface NoServicePeriod {
  label?: string;
  start_date: string; // yyyy-MM-dd
  end_date: string; // yyyy-MM-dd
  reason?: string;
  source?: string;
  state?: string;
  year?: number;
}

export interface Cleaner {
  id: string;
  firstName: string;
  lastName: string;
  email: string;
  phone: string;
  bankAccountName: string;
  bankBsb: string;
  bankAccountNumber: string;
  payRatePerHour: number;
  /** Workforce classification. */
  type?: "cleaner" | "contractor";
}

export interface TimeEntry {
  id: string;
  batch_id: string;
  date: string;
  hours: number;
  pay_rate_snapshot?: number;
  // Joined fields for operational context
  siteId?: string;
  cleanerId?: string;
  /** Optional link to Ad Hoc Job (CleanTrack Ad Hoc Jobs list item id). */
  adhocJobId?: string;
  /** Display name of linked Ad Hoc Job. */
  adhocJobName?: string;
}

/** Ad Hoc Job from CleanTrack Ad Hoc Jobs list (new schema). */
export interface AdHocJob {
  id: string;
  /** SharePoint Title, surfaced as Job Name. */
  jobName: string;
  /**
   * Stored in SharePoint "Job Type" for backward compatibility.
   * App semantics: schedule type = "Once Off" | "Recurring".
   */
  jobType: string;
  companyName: string;
  clientName: string;
  siteId: string | null;
  siteName: string;
  /** When no linked site exists, allow a manual/unlisted name. */
  manualSiteName?: string;
  /** Optional address for manual/unlisted site. */
  manualSiteAddress?: string;
  /** Australian state/territory code for manual/unlisted site (e.g. NSW). */
  manualSiteState?: string;
  description: string;
  requestedByName: string;
  requestedByEmail: string;
  requestChannel: string;
  requestedDate: string | null;
  assignedManagerId: string | null;
  assignedManagerName: string;
  scheduledDate: string | null;
  completedDate: string | null;
  status: string;
  budgetedHours: number | null;
  actualHours: number | null;
  serviceProvider: string;
  chargeRatePerHour: number | null;
  costRatePerHour: number | null;
  charge: number | null;
  cost: number | null;
  grossProfit: number | null;
  markupPercent: number | null;
  gpPercent: number | null;
  /** Optional flag indicating evidence is required before approval. */
  approvalProofRequired?: boolean;
  approvalProofUploaded: boolean;
  /** Optional workflow/method label from SharePoint (if configured). */
  approvalMethod?: string;
  approvalReference: string;
  notesForInformation: string;
  active: boolean;
  /** Whether this ad hoc job should appear in Timesheets for payroll entry/approval. */
  timesheetApplicable?: boolean;

  /** Recurring schedule fields (optional; only used when Schedule Type = Recurring). */
  recurrenceFrequency?: 'Weekly' | 'Fortnightly' | 'Monthly' | 'Quarterly' | null;
  /** Commencement/start date (yyyy-MM-dd). We reuse `scheduledDate` as the start date for recurring jobs. */
  recurrenceStartDate?: string | null;
  recurrenceEndDate?: string | null;
  /** Legacy: single hours-per-day (kept for backward compatibility). */
  hoursPerServiceDay?: number | null;
  /** Selected weekdays for weekly/fortnightly recurrence. 0=Sun..6=Sat. */
  recurrenceWeekdays?: number[] | null;
  /** Per-weekday hours for weekly/fortnightly. Keys are day indexes 0..6. */
  weekdayHours?: Record<string, number> | null;
  /** Monthly configuration. */
  monthlyMode?: 'day_of_month' | 'nth_weekday' | null;
  monthlyDayOfMonth?: number | null;
  monthlyWeekOfMonth?: 'First' | 'Second' | 'Third' | 'Fourth' | 'Last' | null;
  monthlyWeekday?: number | null; // 0=Sun..6=Sat
  /** Hours for monthly occurrence (day-of-month or nth-weekday). */
  monthlyHours?: number | null;

  /** Optional day-type rate overrides (fixed rates, not multipliers). */
  weekdayChargeRateOverride?: number | null;
  saturdayChargeRateOverride?: number | null;
  sundayChargeRateOverride?: number | null;
  publicHolidayChargeRateOverride?: number | null;
  weekdayCostRateOverride?: number | null;
  saturdayCostRateOverride?: number | null;
  sundayCostRateOverride?: number | null;
  publicHolidayCostRateOverride?: number | null;
  /** Placeholder for future attachment support. */
  attachmentCount?: number;
}

export interface FortnightPeriod {
  startDate: Date;
  endDate: Date;
  id: string;
}

/** Fixed vocabulary for timesheet period notes (managers tag context). */
export const TIMESHEET_NOTE_TAG_OPTIONS = [
  "Incomplete / under-delivered hours",
  "Site ending / no service next period",
  "Cleaner last period",
  "Cleaner change / handover",
  "Schedule / plan change",
  "Other",
] as const;

export type TimesheetNoteTag = (typeof TIMESHEET_NOTE_TAG_OPTIONS)[number];

/** One row in CleanTrack Timesheet Period Notes (site + fortnight, optional cleaner). */
export interface TimesheetPeriodNote {
  id: string;
  siteId: string;
  /** Display name from Site lookup when Graph returns it (export / fallback join). */
  siteLookupName?: string;
  /** yyyy-MM-dd, first day of the pay fortnight */
  periodStartYmd: string;
  cleanerId: string | null;
  tags: string[];
  noteBody: string;
}

export function serializeTags(tags: string[]): string {
  return tags.map((t) => t.trim()).filter(Boolean).join("; ");
}

export function deserializeTags(raw: string | null | undefined): string[] {
  if (raw == null || !String(raw).trim()) return [];
  return String(raw)
    .split(";")
    .map((x) => x.trim())
    .filter(Boolean);
}

/** Assignment of a cleaner to a site (from CleanTrack Site Cleaners). */
export interface SiteCleanerAssignment {
  id: string;
  assignmentName: string;
  siteId: string;
  siteName: string;
  cleanerId: string;
  cleanerName: string;
  active: boolean;
}

export type ViewType =
  | 'dashboard'
  | 'sites'
  | 'team'
  | 'cleaners'
  | 'timesheets'
  | 'adhoc-jobs'
  | 'site-detail';