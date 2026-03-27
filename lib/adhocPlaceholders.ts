import { format } from "date-fns";

/** Tokens shown as insert buttons in the ad hoc job form (must stay aligned with the resolver below). */
export const ADHOC_JOB_NAME_PLACEHOLDER_PILLS = [
  { token: "{{month}}", label: "Month" },
  { token: "{{month_short}}", label: "Short month" },
  { token: "{{year}}", label: "Year" },
  { token: "{{month_year}}", label: "Month & year" },
  { token: "{{date_of_service}}", label: "Date of service" },
] as const;

/**
 * Resolve lightweight template placeholders for Ad Hoc job names.
 * Supported tokens:
 * - {{month}} => March
 * - {{month_short}} => Mar
 * - {{year}} => 2026
 * - {{month_year}} => March 2026
 * - {{date_of_service}} => 09 Mar 2026
 */
export function resolveAdHocJobNameTemplate(
  template: string | null | undefined,
  contextDate: Date
): string {
  const raw = String(template ?? "").trim();
  if (!raw) return "";
  return raw.replace(
    /\{\s*\{?\s*([a-z_\-\s]+)\s*\}?\s*\}/gi,
    (_match, tokenRaw: string) => {
      const token = String(tokenRaw ?? "")
        .trim()
        .toLowerCase()
        .replace(/[\s\-]+/g, "_");
      if (token === "month") return format(contextDate, "MMMM");
      if (token === "month_short") return format(contextDate, "MMM");
      if (token === "year") return format(contextDate, "yyyy");
      if (token === "month_year") return format(contextDate, "MMMM yyyy");
      if (token === "date_of_service" || token === "service_date") {
        return format(contextDate, "dd MMM yyyy");
      }
      return _match;
    }
  );
}

