
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
  is_active: boolean;
  budgeted_hours_per_fortnight: number;
  daily_budgets: number[];
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

export interface TimeBatch {
  id: string;
  site_id: string;
  cleaner_id: string;
  fortnight_start: string;
  fortnight_end: string;
  status: 'open' | 'locked';
  updated_at: string;
  updated_by: string;
  editor_name?: string; // Joined profile name
  entries?: TimeEntry[];
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
  approvalProofUploaded: boolean;
  approvalReference: string;
  notesForInformation: string;
  active: boolean;

  /** Recurring schedule fields (optional; only used when Schedule Type = Recurring). */
  recurrenceFrequency?: 'Weekly' | 'Fortnightly' | 'Monthly' | null;
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
  | 'insights'
  | 'adhoc-jobs'
  | 'site-detail'
  | 'auth-test';