import * as sharepoint from "../lib/sharepoint";
import { getCleanTrackUserByEmail, getCleanTrackUserIdToNameMap } from "./usersRepo";

const SITE_MANAGERS_LIST_NAME = "CleanTrack Site Managers";

/** Temporary debug logs for verifying joins. Set to false to disable. */
const DEBUG_SITE_MANAGERS = true;

/**
 * One row in "CleanTrack Site Managers" = one site–manager assignment.
 * Site = lookup to CleanTrack Sites; Manager = lookup to CleanTrack Users (display = Full Name).
 * Assignment Name is display-only; all logic uses lookup IDs.
 */
export interface SiteManagerAssignment {
  id: string;
  assignmentName: string;
  siteId: string;
  siteName: string;
  managerId: string;
  managerName: string;
  active: boolean;
}

/** Extract lookup id from a field value (number, string, or object with LookupId). */
function extractLookupId(val: unknown): string {
  if (val == null) return "";
  if (typeof val === "number" && !Number.isNaN(val)) return String(val);
  if (typeof val === "object" && val !== null && "LookupId" in (val as object))
    return String((val as { LookupId?: number }).LookupId ?? "");
  return String(val).trim();
}

/** Extract display value from a lookup (LookupValue or string). */
function extractLookupValue(val: unknown): string {
  if (val == null) return "";
  if (typeof val === "object" && val !== null && "LookupValue" in (val as object))
    return String((val as { LookupValue?: string }).LookupValue ?? "").trim();
  return String(val).trim();
}

/** Normalize list item id so join works when Graph returns path-style ids (e.g. "sites/.../items/5") or "5". */
export const normalizeListItemId = sharepoint.normalizeListItemId;

/**
 * Fetch active manager assignments from "CleanTrack Site Managers".
 * Expands Site and Manager lookups; Manager display value = Full Name from CleanTrack Users.
 * Maps using lookup IDs only. Filters to Active = true.
 */
export async function fetchSiteManagerAssignments(
  accessToken: string
): Promise<SiteManagerAssignment[]> {
  const siteId = await sharepoint.getSiteId(accessToken);
  const listId = await sharepoint.getListIdByName(accessToken, siteId, SITE_MANAGERS_LIST_NAME);
  if (!listId) return [];

  const columns = await sharepoint.getListColumns(accessToken, siteId, listId);
  const displayToInternal: Record<string, string> = {};
  for (const c of columns) {
    if (c.displayName) displayToInternal[c.displayName] = c.name;
  }
  let siteColumnInternal = displayToInternal["Site"];
  if (!siteColumnInternal) {
    const siteCol = columns.find(
      (c) =>
        (c.displayName && /site/i.test(c.displayName)) ||
        (c.name && /site/i.test(c.name))
    );
    siteColumnInternal = siteCol?.name ?? "Site";
  }
  const siteLookupIdKey = `${siteColumnInternal}LookupId`;
  const siteKey = siteColumnInternal;

  let managerColumnInternal = displayToInternal["Manager"];
  if (!managerColumnInternal) {
    const managerCol = columns.find(
      (c) =>
        (c.displayName && /manager/i.test(c.displayName)) ||
        (c.name && /manager/i.test(c.name))
    );
    managerColumnInternal = managerCol?.name ?? "Manager";
  }
  const managerLookupIdKey = `${managerColumnInternal}LookupId`;
  const managerKey = managerColumnInternal;

  const activeInternal = displayToInternal["Active"] ?? "Active";
  const assignmentNameInternal =
    displayToInternal["Assignment Name"] ?? "Title";
  const fieldsToSelect = [
    siteLookupIdKey,
    siteKey,
    managerLookupIdKey,
    managerKey,
    activeInternal,
    assignmentNameInternal,
  ].filter((v, i, a) => a.indexOf(v) === i);
  const items = await sharepoint.getListItems(accessToken, siteId, listId, fieldsToSelect);

  if (DEBUG_SITE_MANAGERS) {
    console.log("[CleanTrack Site Managers] raw record count:", items.length);
    if (items.length > 0) {
      console.log("[CleanTrack Site Managers] sample raw record:", JSON.stringify(items[0], null, 2));
      console.log("[CleanTrack Site Managers] field keys:", Object.keys((items[0].fields ?? {}) as object));
    }
  }

  const result: SiteManagerAssignment[] = [];
  for (const item of items) {
    const f = (item.fields ?? {}) as Record<string, unknown>;
    const activeKey = Object.keys(f).find((k) => k === "Active" || k.toLowerCase() === "active");
    const active =
      activeKey === undefined
        ? true
        : f[activeKey] === true || f[activeKey] === "Yes" || String(f[activeKey]).toLowerCase() === "yes";
    if (!active) continue;

    const siteVal = f[siteLookupIdKey] ?? f["SiteLookupId"] ?? f["SiteId"] ?? f[siteKey] ?? f["Site"];
    const siteIdStr = extractLookupId(siteVal);
    if (!siteIdStr) {
      if (DEBUG_SITE_MANAGERS)
        console.warn("[CleanTrack Site Managers] skip: no site id", item.id, "keys tried", [siteLookupIdKey, "SiteLookupId", siteKey]);
      continue;
    }
    const siteNameStr = extractLookupValue(f[siteKey] ?? f["Site"] ?? siteVal);

    const managerRaw = f[managerLookupIdKey] ?? f[managerKey] ?? f["Manager"] ?? f["ManagerLookupId"];
    const managerIdStr =
      typeof managerRaw === "object" && managerRaw !== null && "LookupId" in (managerRaw as object)
        ? String((managerRaw as { LookupId?: number }).LookupId ?? "")
        : extractLookupId(managerRaw);
    const managerNameStr =
      typeof managerRaw === "object" && managerRaw !== null && "LookupValue" in (managerRaw as object)
        ? String((managerRaw as { LookupValue?: string }).LookupValue ?? "").trim()
        : extractLookupValue(f[managerKey] ?? f["Manager"] ?? managerRaw);

    const assignmentNameKey = Object.keys(f).find(
      (k) => k === "Title" || k === "Assignment Name" || (k && k.toLowerCase().includes("assignment"))
    );
    const assignmentName = assignmentNameKey ? String(f[assignmentNameKey] ?? "").trim() : "";

    if (item.id && siteIdStr && (managerIdStr || managerNameStr)) {
      result.push({
        id: item.id,
        assignmentName: assignmentName || `${siteNameStr || siteIdStr} - ${managerNameStr || managerIdStr}`,
        siteId: siteIdStr,
        siteName: siteNameStr,
        managerId: managerIdStr,
        managerName: managerNameStr || managerIdStr,
        active: true,
      });
    }
  }

  const userIdToName = await getCleanTrackUserIdToNameMap(accessToken);
  for (const a of result) {
    const normId = normalizeListItemId(a.managerId);
    if (normId && userIdToName[normId]) {
      a.managerName = userIdToName[normId];
    }
  }

  if (DEBUG_SITE_MANAGERS) {
    console.log("[CleanTrack Site Managers] mapped SiteManagerAssignment count:", result.length);
    console.log("[CleanTrack Site Managers] mapped assignments (sample):", result.slice(0, 3));
  }
  return result;
}

/**
 * Join active assignments to sites by siteId.
 * Each site gets assignedManagers = array of assignments for that site.
 * No primary manager logic; Is Primary is ignored.
 */
export function joinAssignmentsToSites<T extends { id: string }>(
  sites: T[],
  assignments: SiteManagerAssignment[]
): Record<string, { assignedManagers: SiteManagerAssignment[] }> {
  const bySite: Record<string, { assignedManagers: SiteManagerAssignment[] }> = {};
  for (const site of sites) {
    const siteIdNorm = normalizeListItemId(site.id);
    const forSite = assignments.filter((a) => normalizeListItemId(a.siteId) === siteIdNorm);
    bySite[site.id] = { assignedManagers: forSite };
    if (DEBUG_SITE_MANAGERS) {
      console.log(
        "[CleanTrack Site Managers] site id:",
        site.id,
        "normalized:",
        siteIdNorm,
        "assignedManagers:",
        forSite.map((a) => a.managerName)
      );
    }
  }
  return bySite;
}

/**
 * Get CleanTrack Sites list item IDs that a manager is assigned to (for Manager role scoping).
 * Uses Manager lookup ID: resolve managerEmail to CleanTrack Users list item id, then filter by that id.
 */
export async function getAssignedSiteIdsForManager(
  accessToken: string,
  managerEmail: string
): Promise<string[]> {
  const user = await getCleanTrackUserByEmail(accessToken, managerEmail);
  if (!user?.id) return [];
  const assignments = await fetchSiteManagerAssignments(accessToken);
  const managerIdNorm = normalizeListItemId(user.id);
  const siteIds = assignments
    .filter((a) => normalizeListItemId(a.managerId) === managerIdNorm)
    .map((a) => normalizeListItemId(a.siteId));
  return [...new Set(siteIds)];
}

/** Get managers assigned to a site (for Edit Site modal). Returns managerId, managerName, and assignment item id for delete. */
export async function getAssignedManagersForSite(
  accessToken: string,
  siteListItemId: string
): Promise<{ managerId: string; managerName: string; itemId: string }[]> {
  const assignments = await fetchSiteManagerAssignments(accessToken);
  const siteIdNorm = normalizeListItemId(siteListItemId);
  return assignments
    .filter((a) => normalizeListItemId(a.siteId) === siteIdNorm)
    .map((a) => ({ managerId: a.managerId, managerName: a.managerName, itemId: a.id }));
}

/** Remove a site–manager assignment by its list item id. */
export async function deleteSiteManagerAssignment(
  accessToken: string,
  assignmentListItemId: string
): Promise<void> {
  const siteId = await sharepoint.getSiteId(accessToken);
  const listId = await sharepoint.getListIdByName(accessToken, siteId, SITE_MANAGERS_LIST_NAME);
  if (!listId) throw new Error(`List "${SITE_MANAGERS_LIST_NAME}" not found.`);
  await sharepoint.deleteListItem(accessToken, siteId, listId, assignmentListItemId);
}

/**
 * Create a site–manager assignment.
 * Saves Site and Manager as SharePoint lookup IDs only (no email, no plain text names).
 * Assignment Name is auto-generated as "[Site Name] - [Manager Full Name]".
 * Is Primary is not set (column left untouched per requirement).
 */
export async function createSiteManagerAssignment(
  accessToken: string,
  siteListItemId: string,
  managerListItemId: string,
  options: { siteName: string; managerFullName: string }
): Promise<void> {
  const siteId = await sharepoint.getSiteId(accessToken);
  const listId = await sharepoint.getListIdByName(accessToken, siteId, SITE_MANAGERS_LIST_NAME);
  if (!listId) throw new Error(`List "${SITE_MANAGERS_LIST_NAME}" not found.`);
  const columns = await sharepoint.getListColumns(accessToken, siteId, listId);
  const map: Record<string, string> = {};
  for (const c of columns) {
    if (c.displayName) map[c.displayName] = c.name;
  }
  const siteInternal = map["Site"] ?? "Site";
  const siteLookupKey = siteInternal.endsWith("LookupId") ? siteInternal : `${siteInternal}LookupId`;
  const managerInternal = map["Manager"] ?? "Manager";
  const managerLookupKey = managerInternal.endsWith("LookupId") ? managerInternal : `${managerInternal}LookupId`;
  const assignmentNameInternal =
    map["Assignment Name"] === "LinkTitle" ? "Title" : (map["Assignment Name"] ?? "Title");
  const activeInternal = map["Active"] ?? "Active";

  const siteIdNorm = normalizeListItemId(siteListItemId);
  const siteNum = parseInt(siteIdNorm, 10);
  const siteLookupValue = Number.isNaN(siteNum) ? siteIdNorm : siteNum;

  const managerIdNorm = normalizeListItemId(managerListItemId);
  const managerNum = parseInt(managerIdNorm, 10);
  const managerLookupValue =
    !managerIdNorm || (managerNum === 0 && managerIdNorm === "0")
      ? null
      : Number.isNaN(managerNum)
        ? managerIdNorm
        : managerNum;

  const assignmentName = `${(options.siteName || "Site").trim()} - ${(options.managerFullName || "Manager").trim()}`;

  const fields: Record<string, unknown> = {
    [assignmentNameInternal]: assignmentName,
    [activeInternal]: true,
    [siteLookupKey]: siteLookupValue,
  };
  if (managerLookupValue !== null && managerLookupValue !== 0) {
    fields[managerLookupKey] = managerLookupValue;
  } else {
    throw new Error(
      "Manager lookup id is missing or zero. Ensure the selected manager comes from CleanTrack Users and has a valid list item id."
    );
  }

  await sharepoint.createListItem(accessToken, siteId, listId, fields);
}
