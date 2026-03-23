import { format } from "date-fns";
import type { NoServicePeriod } from "../types";

function normalizePeriod(raw: NoServicePeriod | undefined | null): NoServicePeriod | null {
  if (!raw) return null;
  const start = String(raw.start_date ?? "").trim();
  const end = String(raw.end_date ?? "").trim();
  if (!start || !end) return null;
  const label = String(raw.label ?? "").trim();
  const reason = String(raw.reason ?? "").trim();
  return {
    start_date: start,
    end_date: end,
    ...(label ? { label } : {}),
    ...(reason ? { reason } : {}),
  };
}

export function getNoServicePeriodForDate(
  date: Date,
  periods: NoServicePeriod[] | undefined | null
): NoServicePeriod | undefined {
  if (!periods || periods.length === 0) return undefined;
  const key = format(date, "yyyy-MM-dd");
  for (const p of periods) {
    const n = normalizePeriod(p);
    if (!n) continue;
    if (key >= n.start_date && key <= n.end_date) return n;
  }
  return undefined;
}

export function isDateInNoServicePeriod(
  date: Date,
  periods: NoServicePeriod[] | undefined | null
): boolean {
  return !!getNoServicePeriodForDate(date, periods);
}
