
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

export interface FortnightPeriod {
  startDate: Date;
  endDate: Date;
  id: string;
}

export type ViewType = 'dashboard' | 'sites' | 'team' | 'cleaners' | 'timesheets' | 'insights' | 'site-detail' | 'auth-test';