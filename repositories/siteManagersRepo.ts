import * as sharepoint from "../lib/sharepoint";

const SITE_MANAGERS_LIST_NAME = "CleanTrack Site Managers";

/**
 * Get the list of CleanTrack Sites list item IDs that a manager is assigned to.
 * Used for Manager role scoping: only these sites should be visible.
 * Returns empty array if no assignments or list not found.
 */
export async function getAssignedSiteIdsForManager(
  accessToken: string,
  managerEmail: string
): Promise<string[]> {
  const siteId = await sharepoint.getSiteId(accessToken);
  const listId = await sharepoint.getListIdByName(accessToken, siteId, SITE_MANAGERS_LIST_NAME);
  if (!listId) return [];
  const items = await sharepoint.getListItems(accessToken, siteId, listId);
  const email = managerEmail.trim().toLowerCase();
  const siteIds: string[] = [];
  for (const item of items) {
    const f = item.fields ?? {};
    const activeKey = Object.keys(f).find((k) => k === "Active" || k.toLowerCase() === "active");
    if (activeKey) {
      const v = (f as Record<string, unknown>)[activeKey];
      if (v === false || v === "No" || String(v).toLowerCase() === "no") continue;
    }
    const managerRaw =
      (f as Record<string, unknown>)["Manager"] ??
      (f as Record<string, unknown>)["ManagerEmail"] ??
      (f as Record<string, unknown>)["ManagerId"];
    const managerStr =
      typeof managerRaw === "object" && managerRaw !== null && "Email" in managerRaw
        ? String((managerRaw as { Email?: string }).Email ?? "")
        : String(managerRaw ?? "").trim();
    if (managerStr.toLowerCase() !== email) continue;

    const siteIdVal =
      (f as Record<string, unknown>)["SiteLookupId"] ??
      (f as Record<string, unknown>)["SiteId"] ??
      (f as Record<string, unknown>)["Site"];
    if (siteIdVal != null && siteIdVal !== "") {
      const id = typeof siteIdVal === "number" ? String(siteIdVal) : String(siteIdVal).trim();
      if (id) siteIds.push(id);
    }
  }
  return [...new Set(siteIds)];
}

export interface SiteManagerAssignment {
  email: string;
  itemId: string;
}

/** Get managers assigned to a site (for Edit Site modal). */
export async function getAssignedManagersForSite(
  accessToken: string,
  siteListItemId: string
): Promise<SiteManagerAssignment[]> {
  const siteId = await sharepoint.getSiteId(accessToken);
  const listId = await sharepoint.getListIdByName(accessToken, siteId, SITE_MANAGERS_LIST_NAME);
  if (!listId) return [];
  const items = await sharepoint.getListItems(accessToken, siteId, listId);
  const siteIdStr = siteListItemId.trim();
  const result: SiteManagerAssignment[] = [];
  for (const item of items) {
    const f = item.fields ?? {};
    const activeKey = Object.keys(f).find((k) => k === "Active" || k.toLowerCase() === "active");
    if (activeKey) {
      const v = (f as Record<string, unknown>)[activeKey];
      if (v === false || v === "No" || String(v).toLowerCase() === "no") continue;
    }
    const siteVal =
      (f as Record<string, unknown>)["SiteLookupId"] ??
      (f as Record<string, unknown>)["SiteId"] ??
      (f as Record<string, unknown>)["Site"];
    const sid = siteVal != null ? String(siteVal).trim() : "";
    if (sid !== siteIdStr) continue;
    const managerRaw =
      (f as Record<string, unknown>)["Manager"] ??
      (f as Record<string, unknown>)["ManagerEmail"] ??
      (f as Record<string, unknown>)["ManagerId"];
    const email =
      typeof managerRaw === "object" && managerRaw !== null && "Email" in managerRaw
        ? String((managerRaw as { Email?: string }).Email ?? "").trim()
        : String(managerRaw ?? "").trim();
    if (item.id && email) result.push({ email, itemId: item.id });
  }
  return result;
}

/** Remove a site-manager assignment by its list item id. */
export async function deleteSiteManagerAssignment(
  accessToken: string,
  assignmentListItemId: string
): Promise<void> {
  const siteId = await sharepoint.getSiteId(accessToken);
  const listId = await sharepoint.getListIdByName(accessToken, siteId, SITE_MANAGERS_LIST_NAME);
  if (!listId) throw new Error(`List "${SITE_MANAGERS_LIST_NAME}" not found.`);
  await sharepoint.deleteListItem(accessToken, siteId, listId, assignmentListItemId);
}

/** Create a Site Manager assignment: link a manager (by email) to a site. */
export async function createSiteManagerAssignment(
  accessToken: string,
  siteListItemId: string,
  managerEmail: string,
  options?: { assignmentName?: string; isPrimary?: boolean }
): Promise<void> {
  const siteId = await sharepoint.getSiteId(accessToken);
  const listId = await sharepoint.getListIdByName(accessToken, siteId, SITE_MANAGERS_LIST_NAME);
  if (!listId) throw new Error(`List "${SITE_MANAGERS_LIST_NAME}" not found.`);
  const columns = await sharepoint.getListColumns(accessToken, siteId, listId);
  const map: Record<string, string> = {};
  for (const c of columns) {
    if (c.displayName) map[c.displayName] = c.name;
  }
  const siteInternal = map["Site"] ?? "SiteLookupId";
  const managerInternal = map["Manager"] ?? "Manager";
  const assignmentNameInternal = map["Assignment Name"] ?? "Title";
  const activeInternal = map["Active"] ?? "Active";
  const isPrimaryInternal = map["Is Primary"] ?? "IsPrimary";

  const fields: Record<string, unknown> = {
    [assignmentNameInternal]: options?.assignmentName ?? `Site - ${managerEmail}`,
    [activeInternal]: true,
    [isPrimaryInternal]: options?.isPrimary ?? false,
    [managerInternal]: managerEmail,
  };
  const num = parseInt(siteListItemId, 10);
  fields[siteInternal] = Number.isNaN(num) ? siteListItemId : num;

  await sharepoint.createListItem(accessToken, siteId, listId, fields);
}
