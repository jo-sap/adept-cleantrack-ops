
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
  /** Hourly rate for budgeted labour at this site ($/hr). From CleanTrack Site Budgets. */
  budget_labour_rate?: number;
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

/** Ad Hoc Job from CleanTrack Ad Hoc Jobs list. */
export interface AdHocJob {
  id: string;
  jobName: string;
  jobType: string;
  siteId: string | null;
  siteName: string;
  requestedByName: string;
  requestedByEmail: string;
  requestedByCompany: string;
  requestChannel: string;
  requestSummary: string;
  requestedDate: string | null;
  assignedManagerId: string | null;
  assignedManagerName: string;
  scheduledDate: string | null;
  completedDate: string | null;
  status: string;
  budgetedHours: number | null;
  budgetedLabourRate: number | null;
  budgetedRevenue: number | null;
  description: string;
  approvalProofRequired: boolean;
  approvalProofUploaded: boolean;
  approvalReferenceNotes: string;
  active: boolean;
  /** Placeholder for future attachment support. */
  attachmentCount?: number;
}

export interface FortnightPeriod {
  startDate: Date;
  endDate: Date;
  id: string;
}

export type ViewType = 'dashboard' | 'sites' | 'team' | 'cleaners' | 'timesheets' | 'insights' | 'adhoc-jobs' | 'site-detail' | 'auth-test';