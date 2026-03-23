/** Normalize site labels for joining note lookup display names to app `site.name`. */
export function normalizeSiteLabelForNotes(s: string | undefined | null): string {
  return String(s ?? "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");
}
