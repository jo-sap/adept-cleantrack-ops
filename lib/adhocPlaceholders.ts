import { format } from "date-fns";

/** Tokens shown as insert buttons in the ad hoc job form (must stay aligned with the resolver below). */
export const ADHOC_JOB_NAME_PLACEHOLDER_PILLS = [
  { token: "{{month}}", label: "Month" },
  { token: "{{month_short}}", label: "Short month" },
  { token: "{{year}}", label: "Year" },
  { token: "{{month_year}}", label: "Month & year" },
] as const;

/**
 * Resolve lightweight template placeholders for Ad Hoc job names.
 * Supported tokens:
 * - {{month}} => March
 * - {{month_short}} => Mar
 * - {{year}} => 2026
 * - {{month_year}} => March 2026
 */
export function resolveAdHocJobNameTemplate(
  template: string | null | undefined,
  contextDate: Date
): string {
  const raw = String(template ?? "").trim();
  if (!raw) return "";
  return raw.replace(/\{\{\s*([a-z_]+)\s*\}\}/gi, (_match, tokenRaw: string) => {
    const token = String(tokenRaw ?? "").trim().toLowerCase();
    if (token === "month") return format(contextDate, "MMMM");
    if (token === "month_short") return format(contextDate, "MMM");
    if (token === "year") return format(contextDate, "yyyy");
    if (token === "month_year") return format(contextDate, "MMMM yyyy");
    return _match;
  });
}

