import * as sharepoint from "../lib/sharepoint";
import { normalizeSiteLabelForNotes } from "../lib/siteNotesLabel";
import type { TimesheetPeriodNote } from "../types";
import { deserializeTags, serializeTags } from "../types";

export { normalizeSiteLabelForNotes };

/**
 * SharePoint list **CleanTrack Timesheet Period Notes**
 *
 * Supported layouts:
 * - Minimal: **Site** (must be **Look up** → list **CleanTrack Sites** — not text/hyperlink), **Period Start** (date / date & time), **Note** (multi-line).
 * - Optional: **Title** (filled by app), **Tags**, **Cleaner** (lookup) — used only if those columns exist.
 *
 * Period Start: store the first day of the pay fortnight (yyyy-MM-dd). Date-only or date-without-time display in SP is fine.
 */

const LIST_NAME = "CleanTrack Timesheet Period Notes";

function pad2(n: number): string {
  return String(n).padStart(2, "0");
}

function toLocalYmd(d: Date): string {
  return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
}

/** Graph dateTime columns sometimes return `{ DateTime, TimeZone }` instead of a plain string. */
function fieldToYmd(v: unknown): string | null {
  if (v == null) return null;
  if (typeof v === "string") {
    if (/^\d{4}-\d{2}-\d{2}/.test(v)) return v.slice(0, 10);
    const d = new Date(v);
    return isNaN(d.getTime()) ? null : toLocalYmd(d);
  }
  if (typeof v === "object" && v !== null) {
    if ("Value" in v) return fieldToYmd((v as { Value?: unknown }).Value);
    if ("DateTime" in v) return fieldToYmd((v as { DateTime?: unknown }).DateTime);
  }
  return null;
}

function coerceIdFromLookupCell(v: unknown, depth = 0): string {
  if (depth > 8) return "";
  if (v == null) return "";
  if (typeof v === "number" && Number.isFinite(v)) {
    if (v === 0) return "";
    return String(Math.trunc(v));
  }
  if (typeof v === "string") {
    const t = v.trim();
    if (!t || t === "0") return "";
    if (/^\d+$/.test(t)) return t;
    return "";
  }
  if (typeof v === "object") {
    const o = v as Record<string, unknown>;
    for (const k of ["LookupId", "lookupId", "Id", "id"]) {
      const inner = coerceIdFromLookupCell(o[k], depth + 1);
      if (inner) return inner;
    }
  }
  return "";
}

function getLookupId(f: Record<string, unknown>, columnBase: string): string {
  const idKey =
    Object.keys(f).find((k) => k === `${columnBase}LookupId` || k === `${columnBase}Id`) ??
    Object.keys(f).find((k) => k.toLowerCase() === `${columnBase.toLowerCase()}lookupid`);
  if (idKey) {
    const raw = f[idKey];
    if (raw == null || raw === "") return "";
    const s = String(raw).trim();
    if (s === "0") return "";
    return s;
  }
  // Graph often omits *LookupId on GET but still returns the lookup cell (id, object, or display-only string).
  return coerceIdFromLookupCell(f[columnBase]);
}

/** Title is filled as `{siteName} — {yyyy-MM-dd}`; use when Period/Site fields are incomplete in Graph. */
function parseSiteAndPeriodFromAppTitle(title: string): { siteName?: string; periodYmd?: string } {
  const t = title.trim();
  if (!t) return {};
  const periodMatch = t.match(/(\d{4}-\d{2}-\d{2})\s*$/);
  const periodYmd = periodMatch?.[1];
  let siteName: string | undefined;
  if (periodMatch && periodMatch.index != null && periodMatch.index > 0) {
    siteName = t
      .slice(0, periodMatch.index)
      .replace(/\s*[—–-]\s*$/, "")
      .trim();
  }
  if (!siteName) {
    const first = t.split(/\s*[—–-]\s*/)[0]?.trim();
    if (first && first !== t) siteName = first;
  }
  return {
    ...(siteName ? { siteName } : {}),
    ...(periodYmd ? { periodYmd } : {}),
  };
}

function buildColumnMap(columns: Array<{ name: string; displayName: string }>): Record<string, string> {
  const map: Record<string, string> = {};
  for (const c of columns) {
    if (c.displayName) map[c.displayName] = c.name;
    if (c.name) map[c.name] = c.name;
  }
  return map;
}

interface ResolvedNoteKeys {
  /** Internal field for Title; omitted on write if null */
  titleKey: string | null;
  /** Internal name of the Site lookup (for reading *LookupId from item fields) */
  siteFieldBase: string;
  /** Graph fields payload key, e.g. Site0LookupId — always `${siteFieldBase}LookupId` */
  siteLookupKey: string;
  periodStartKey: string;
  noteKey: string;
  tagsKey: string | null;
  cleanerLookupKey: string | null;
  hasCleanerColumn: boolean;
  /** Set when Site exists but is not a Sites lookup (common misconfiguration). */
  listSchemaError: string | null;
}

/**
 * Resolve Graph field names. Site must come from a real **lookup** column named Site — not the
 * default Title field if its display name was changed to "Site" (internal names like LinkTitle would
 * wrongly produce LinkTitleLookupId).
 */
async function resolveNoteListKeys(
  accessToken: string,
  spSiteId: string,
  listId: string
): Promise<ResolvedNoteKeys> {
  const [columns, defs] = await Promise.all([
    sharepoint.getListColumns(accessToken, spSiteId, listId),
    sharepoint.getListColumnDefinitions(accessToken, spSiteId, listId),
  ]);
  const map = buildColumnMap(columns);

  const isTitleLikeName = (internal: string) => {
    const x = internal.trim().toLowerCase();
    return x === "title" || x.startsWith("linktitle");
  };

  /** Columns whose display or internal name is "Site" */
  const siteNameCols = defs.filter((c) => {
    const dn = c.displayName?.trim().toLowerCase() ?? "";
    const n = c.name?.trim().toLowerCase() ?? "";
    return dn === "site" || n === "site";
  });

  const siteLookupCol = siteNameCols.find((c) => c.lookup != null);

  let listSchemaError: string | null = null;
  let siteFieldBase: string;

  if (siteLookupCol?.name?.trim()) {
    siteFieldBase = siteLookupCol.name.trim();
  } else if (
    siteNameCols.length === 1 &&
    siteNameCols[0].name?.trim() &&
    !isTitleLikeName(siteNameCols[0].name)
  ) {
    const sole = siteNameCols[0];
    if (sole.lookup == null && sole.text != null) {
      listSchemaError =
        'SharePoint: "Site" is a text column. Delete it and add **Look up** → **CleanTrack Sites** so the app can store the site link.';
      siteFieldBase = "Site";
    } else {
      // Graph sometimes omits `lookup` on real lookup columns; treat as Sites lookup.
      siteFieldBase = sole.name.trim();
    }
  } else if (siteNameCols.length > 0) {
    listSchemaError =
      'SharePoint: the "Site" column must be a Lookup to the list **CleanTrack Sites** (Add column → Look up → target that list). Single line of text, Hyperlink, or a renamed Title field will not work — remove the wrong column and add a proper Site lookup.';
    siteFieldBase = "Site";
  } else {
    const fromMap = map["Site"];
    if (fromMap && !isTitleLikeName(fromMap)) {
      siteFieldBase = fromMap;
    } else {
      siteFieldBase = "Site";
    }
  }

  const siteLookupKey = `${siteFieldBase}LookupId`;

  const titleCol = defs.find((c) => {
    if (c.lookup != null || c.personOrGroup != null) return false;
    const dn = c.displayName?.trim().toLowerCase() ?? "";
    if (dn === "site") return false;
    const name = c.name?.trim() ?? "";
    if (name === "Title") return true;
    if (dn === "title") return true;
    return false;
  });
  const titleKey = titleCol?.name ?? null;

  const periodStartKey =
    map["Period Start"] ??
    map["PeriodStart"] ??
    Object.keys(map).find((k) => k.toLowerCase().replace(/\s/g, "") === "periodstart") ??
    "Period_x0020_Start";

  const noteKey =
    map["Note"] ?? map["Note Body"] ?? map["NoteBody"] ?? "Note";

  const tagsKey = map["Tags"] ?? map["Tag"] ?? null;

  const cleanerInternal = map["Cleaner"];
  const hasCleanerColumn = !!cleanerInternal;
  const cleanerLookupKey = hasCleanerColumn
    ? cleanerInternal === "Cleaner"
      ? "CleanerLookupId"
      : `${cleanerInternal}LookupId`
    : null;

  return {
    titleKey,
    siteFieldBase,
    siteLookupKey,
    periodStartKey,
    noteKey,
    tagsKey,
    cleanerLookupKey,
    hasCleanerColumn,
    listSchemaError,
  };
}

function readLookupDisplayName(f: Record<string, unknown>, fieldInternalName: string): string {
  const v = f[fieldInternalName];
  if (v == null) return "";
  if (typeof v === "string") return v.trim();
  if (typeof v === "object" && v !== null && "LookupValue" in v) {
    return String((v as { LookupValue?: unknown }).LookupValue ?? "").trim();
  }
  return "";
}

/** Compare Period Start values from SharePoint vs app (yyyy-MM-dd, tolerate T… suffix). */
export function comparablePeriodYmd(raw: string | undefined | null): string {
  if (raw == null) return "";
  const s = String(raw).trim();
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
  const d = new Date(s);
  if (isNaN(d.getTime())) return "";
  return toLocalYmd(d);
}

function parseItem(item: sharepoint.GraphListItem, keys: ResolvedNoteKeys): TimesheetPeriodNote | null {
  const f = (item.fields ?? {}) as Record<string, unknown>;
  const siteRaw = getLookupId(f, keys.siteFieldBase);
  let siteLookupName = readLookupDisplayName(f, keys.siteFieldBase);
  let periodYmd = fieldToYmd(f[keys.periodStartKey]);

  if (keys.titleKey) {
    const fromTitle = parseSiteAndPeriodFromAppTitle(String(f[keys.titleKey] ?? ""));
    if (!periodYmd && fromTitle.periodYmd) periodYmd = fromTitle.periodYmd;
    if (!siteLookupName.trim() && fromTitle.siteName) siteLookupName = fromTitle.siteName;
  }

  if (!periodYmd) return null;
  const hasSiteId = !!siteRaw && String(siteRaw).trim() !== "" && String(siteRaw).trim() !== "0";
  if (!hasSiteId && !siteLookupName.trim()) return null;

  let cleanerRaw = "";
  if (keys.hasCleanerColumn) {
    cleanerRaw = getLookupId(f, "Cleaner");
  }

  const tagsRaw =
    keys.tagsKey && f[keys.tagsKey] != null ? String(f[keys.tagsKey]) : "";

  return {
    id: sharepoint.normalizeListItemId(item.id),
    siteId: hasSiteId ? sharepoint.normalizeListItemId(siteRaw) : "",
    ...(siteLookupName.trim() ? { siteLookupName: siteLookupName.trim() } : {}),
    periodStartYmd: periodYmd,
    cleanerId: cleanerRaw ? sharepoint.normalizeListItemId(cleanerRaw) : null,
    tags: deserializeTags(tagsRaw),
    noteBody: String(f[keys.noteKey] ?? "").trim(),
  };
}

function lookupWriteValue(listItemId: string): string | number {
  const n = parseInt(listItemId, 10);
  return Number.isNaN(n) ? listItemId : n;
}

export interface ListTimesheetPeriodNotesResult {
  notes: TimesheetPeriodNote[];
  listExists: boolean;
  /** Explains list misconfiguration (e.g. Site is text, not lookup). */
  listSchemaError?: string | null;
}

/** Site-wide note row for a pay fortnight (no cleaner id). */
export function pickSiteNoteForPeriod(
  notes: TimesheetPeriodNote[],
  siteId: string,
  periodStartYmd: string,
  /** When Graph omits lookup id on read, match using Site display / Title site name. */
  siteDisplayName?: string
): TimesheetPeriodNote | undefined {
  const s = sharepoint.normalizeListItemId(siteId);
  const p = comparablePeriodYmd(periodStartYmd);
  const nameNorm = normalizeSiteLabelForNotes(siteDisplayName);
  return notes.find((n) => {
    const nSid = n.siteId ? sharepoint.normalizeListItemId(n.siteId) : "";
    const idOk = !!s && nSid === s;
    const nameOk = !!nameNorm && normalizeSiteLabelForNotes(n.siteLookupName) === nameNorm;
    if (!idOk && !(nameOk && !nSid)) return false;
    if (comparablePeriodYmd(n.periodStartYmd) !== p) return false;
    return n.cleanerId == null || String(n.cleanerId).trim() === "";
  });
}

export type SiteNotesExportLookup = {
  bySiteId: Record<string, string>;
  bySiteNameLower: Record<string, string>;
};

/** For XLSX export: map by normalized site id and by lookup display name (ids can differ from app sites). */
export function buildSiteNotesExportLookup(
  notes: TimesheetPeriodNote[],
  periodStartYmd: string
): SiteNotesExportLookup {
  const p = comparablePeriodYmd(periodStartYmd);
  const bySiteId: Record<string, string> = {};
  const bySiteNameLower: Record<string, string> = {};

  for (const n of notes) {
    if (n.cleanerId != null && String(n.cleanerId).trim() !== "") continue;
    if (comparablePeriodYmd(n.periodStartYmd) !== p) continue;
    const body = n.noteBody?.trim();
    if (!body) continue;
    const sid = sharepoint.normalizeListItemId(n.siteId);
    if (sid) bySiteId[sid] = body;
    const nm = normalizeSiteLabelForNotes(n.siteLookupName);
    if (nm) bySiteNameLower[nm] = body;
  }
  return { bySiteId, bySiteNameLower };
}

export async function listAllTimesheetPeriodNotes(
  accessToken: string
): Promise<ListTimesheetPeriodNotesResult> {
  const spSiteId = await sharepoint.getSiteId(accessToken);
  const listId = await sharepoint.getListIdByName(accessToken, spSiteId, LIST_NAME);
  if (!listId) return { notes: [], listExists: false };

  const keys = await resolveNoteListKeys(accessToken, spSiteId, listId);

  const items = await sharepoint.getListItems(accessToken, spSiteId, listId);
  const out: TimesheetPeriodNote[] = [];
  for (const item of items) {
    const n = parseItem(item, keys);
    if (n) out.push(n);
  }
  return {
    notes: out,
    listExists: true,
    listSchemaError: keys.listSchemaError,
  };
}

export interface SaveTimesheetPeriodNoteInput {
  siteId: string;
  siteName: string;
  periodStartYmd: string;
  cleanerId: string | null;
  tags: string[];
  noteBody: string;
}

/**
 * Create, update, or delete (empty note body and no tags when Tags column exists) for (site, period, optional cleaner).
 */
export async function upsertTimesheetPeriodNote(
  accessToken: string,
  input: SaveTimesheetPeriodNoteInput
): Promise<{ ok: true } | { ok: false; error: string }> {
  const spSiteId = await sharepoint.getSiteId(accessToken);
  const listId = await sharepoint.getListIdByName(accessToken, spSiteId, LIST_NAME);
  if (!listId) {
    return { ok: false, error: `SharePoint list "${LIST_NAME}" was not found.` };
  }

  const keys = await resolveNoteListKeys(accessToken, spSiteId, listId);
  if (keys.listSchemaError) {
    return { ok: false, error: keys.listSchemaError };
  }

  const normSite = sharepoint.normalizeListItemId(input.siteId);
  const normCleaner = input.cleanerId ? sharepoint.normalizeListItemId(input.cleanerId) : null;
  const inputSiteLabel = normalizeSiteLabelForNotes(input.siteName);
  const inputPeriod = comparablePeriodYmd(input.periodStartYmd);

  const all = await sharepoint.getListItems(accessToken, spSiteId, listId);
  const matches: sharepoint.GraphListItem[] = [];
  for (const item of all) {
    const n = parseItem(item, keys);
    if (!n) continue;
    const nSid = n.siteId ? sharepoint.normalizeListItemId(n.siteId) : "";
    const idMatch = !!normSite && !!nSid && nSid === normSite;
    const nameOnlyMatch =
      !nSid &&
      !!inputSiteLabel &&
      normalizeSiteLabelForNotes(n.siteLookupName) === inputSiteLabel;
    if (!idMatch && !nameOnlyMatch) continue;
    if (comparablePeriodYmd(n.periodStartYmd) !== inputPeriod) continue;
    const itemCleaner = n.cleanerId ?? null;
    if (itemCleaner !== normCleaner) continue;
    matches.push(item);
  }

  const tagsStr = keys.tagsKey ? serializeTags(input.tags) : "";
  const bodyStr = input.noteBody.trim();
  const isEmpty =
    bodyStr === "" && (!keys.tagsKey || tagsStr === "");

  if (isEmpty) {
    for (const m of matches) {
      try {
        await sharepoint.deleteListItem(accessToken, spSiteId, listId, m.id);
      } catch (e) {
        const msg = e instanceof Error ? e.message : String(e);
        return { ok: false, error: msg };
      }
    }
    return { ok: true };
  }

  const titleBase = `${input.siteName} — ${input.periodStartYmd}${normCleaner ? " (cleaner)" : ""}`;
  const title = titleBase.length > 255 ? titleBase.slice(0, 252) + "…" : titleBase;

  const siteVal = lookupWriteValue(input.siteId);

  const baseFields: Record<string, unknown> = {
    [keys.siteLookupKey]: siteVal,
    [keys.periodStartKey]: input.periodStartYmd,
    [keys.noteKey]: bodyStr,
  };

  if (keys.titleKey) {
    baseFields[keys.titleKey] = title;
  }
  if (keys.tagsKey) {
    baseFields[keys.tagsKey] = tagsStr;
  }

  try {
    if (matches.length > 0) {
      const keep = matches[0];
      for (let i = 1; i < matches.length; i++) {
        await sharepoint.deleteListItem(accessToken, spSiteId, listId, matches[i].id);
      }
      const patch: Record<string, unknown> = { ...baseFields };
      if (keys.cleanerLookupKey) {
        patch[keys.cleanerLookupKey] = normCleaner ? lookupWriteValue(normCleaner) : null;
      }
      await sharepoint.updateListItem(accessToken, spSiteId, listId, keep.id, patch);
    } else {
      const createFields: Record<string, unknown> = { ...baseFields };
      if (keys.cleanerLookupKey && normCleaner) {
        createFields[keys.cleanerLookupKey] = lookupWriteValue(normCleaner);
      }
      await sharepoint.createListItem(accessToken, spSiteId, listId, createFields);
    }
    return { ok: true };
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e);
    return { ok: false, error: msg };
  }
}
