import * as sharepoint from "../lib/sharepoint";
import { getCleaners } from "./cleanersRepo";

const TIMESHEETS_LIST_NAME = "CleanTrack Timesheet Entries";

export interface AssignedCleaner {
  name: string;
  payRate: number;
}

function findFieldKey(fields: Record<string, unknown>, ...candidates: string[]): string | undefined {
  const keys = Object.keys(fields);
  for (const c of candidates) {
    const k = keys.find(
      (x) => x === c || x.toLowerCase() === c.toLowerCase()
    );
    if (k) return k;
  }
  return undefined;
}

function getLookupId(fields: Record<string, unknown>, baseName: string): string | null {
  const idKey =
    findFieldKey(fields, `${baseName}LookupId`, `${baseName}Id`) ??
    findFieldKey(fields, `${baseName}LookupId`);
  if (!idKey) return null;
  const v = (fields as Record<string, unknown>)[idKey];
  if (v == null) return null;
  return String(v).trim() || null;
}

/**
 * Derive assigned cleaners per site from CleanTrack Timesheet Entries (cleaners who have
 * entries for that site) and CleanTrack Cleaners (name + pay rate).
 */
export async function getAssignedCleanersBySite(
  accessToken: string
): Promise<Record<string, AssignedCleaner[]>> {
  const siteId = await sharepoint.getSiteId(accessToken);
  const timesheetsListId = await sharepoint.getListIdByName(
    accessToken,
    siteId,
    TIMESHEETS_LIST_NAME
  );
  if (!timesheetsListId) return {};

  const [timesheetItems, cleaners] = await Promise.all([
    sharepoint.getListItems(accessToken, siteId, timesheetsListId),
    getCleaners(accessToken),
  ]);

  const cleanerById = new Map<string | number, AssignedCleaner>();
  for (const c of cleaners) {
    cleanerById.set(c.id, { name: c.cleanerName, payRate: c.payRatePerHour });
  }

  const siteToCleanerIds = new Map<string, Set<string>>();
  for (const item of timesheetItems) {
    const f = (item.fields ?? {}) as Record<string, unknown>;
    const siteIdVal = getLookupId(f, "Site");
    const cleanerIdVal = getLookupId(f, "Cleaner");
    if (!siteIdVal || !cleanerIdVal) continue;
    let set = siteToCleanerIds.get(siteIdVal);
    if (!set) {
      set = new Set<string>();
      siteToCleanerIds.set(siteIdVal, set);
    }
    set.add(cleanerIdVal);
  }

  const result: Record<string, AssignedCleaner[]> = {};
  for (const [sid, ids] of siteToCleanerIds) {
    const list: AssignedCleaner[] = [];
    const seen = new Set<string>();
    for (const id of ids) {
      const info = cleanerById.get(id) ?? cleanerById.get(parseInt(id, 10));
      if (!info) continue;
      const key = `${info.name}|${info.payRate}`;
      if (seen.has(key)) continue;
      seen.add(key);
      list.push(info);
    }
    result[sid] = list;
  }
  return result;
}
