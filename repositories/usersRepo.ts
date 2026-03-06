import * as sharepoint from "../lib/sharepoint";

const CLEANTRACK_USERS_LIST_NAME = "CleanTrack Users";

export interface CleanTrackUser {
  fullName: string;
  email: string;
  role: string;
  active: boolean;
  permissionScope: string | null;
}

function isActive(value: unknown): boolean {
  if (value === true) return true;
  if (typeof value === "string") return value.trim().toLowerCase() === "yes";
  return false;
}

function normalizeRole(value: unknown): string {
  const s = typeof value === "string" ? value.trim() : "";
  const lower = s.toLowerCase();
  if (lower === "admin") return "Admin";
  if (lower === "manager") return "Manager";
  return s || "Manager";
}

/**
 * Look up a user in the CleanTrack Users list by email (case-insensitive).
 * Returns null if not found.
 */
export async function getCleanTrackUserByEmail(
  accessToken: string,
  email: string
): Promise<CleanTrackUser | null> {
  const siteId = await sharepoint.getSiteId(accessToken);
  const listId = await sharepoint.getListIdByName(accessToken, siteId, CLEANTRACK_USERS_LIST_NAME);
  if (!listId) return null;
  const items = await sharepoint.getListItems(accessToken, siteId, listId);
  const needle = email.trim().toLowerCase();
  const row = items.find((item) => {
    const e = (item.fields?.Email as string) ?? "";
    return e.trim().toLowerCase() === needle;
  });
  if (!row?.fields) return null;
  const fields = row.fields;
  return {
    fullName: String(fields.Title ?? ""),
    email: String(fields.Email ?? ""),
    role: normalizeRole(fields.Role),
    active: isActive(fields.Active),
    permissionScope: fields.PermissionScope != null ? String(fields.PermissionScope) : null,
  };
}

/** Manager summary for assignment UI. */
export interface ManagerOption {
  fullName: string;
  email: string;
}

/** List all users in CleanTrack Users with Role = Manager and Active = true. */
export async function getCleanTrackManagers(
  accessToken: string
): Promise<ManagerOption[]> {
  const siteId = await sharepoint.getSiteId(accessToken);
  const listId = await sharepoint.getListIdByName(accessToken, siteId, CLEANTRACK_USERS_LIST_NAME);
  if (!listId) return [];
  const items = await sharepoint.getListItems(accessToken, siteId, listId);
  const result: ManagerOption[] = [];
  for (const item of items) {
    const fields = item.fields ?? {};
    if (!isActive(fields.Active)) continue;
    const role = normalizeRole(fields.Role);
    if (role !== "Manager") continue;
    const fullName = String(fields.Title ?? "").trim();
    const email = String(fields.Email ?? "").trim();
    if (email) result.push({ fullName: fullName || email, email });
  }
  return result;
}

export interface CleanTrackUserRow {
  id: string;
  fullName: string;
  email: string;
  role: string;
  active: boolean;
}

/** List all users from CleanTrack Users (any role). */
export async function getCleanTrackUsers(
  accessToken: string
): Promise<CleanTrackUserRow[]> {
  const siteId = await sharepoint.getSiteId(accessToken);
  const listId = await sharepoint.getListIdByName(accessToken, siteId, CLEANTRACK_USERS_LIST_NAME);
  if (!listId) return [];
  const items = await sharepoint.getListItems(accessToken, siteId, listId);
  const result: CleanTrackUserRow[] = [];
  for (const item of items) {
    const fields = item.fields ?? {};
    const fullName = String((fields as any).Title ?? "").trim();
    const email = String((fields as any).Email ?? "").trim();
    const role = normalizeRole((fields as any).Role);
    const active = isActive((fields as any).Active);
    if (!email) continue;
    result.push({
      id: item.id!,
      fullName: fullName || email,
      email,
      role,
      active,
    });
  }
  return result;
}

export interface ManagerUpsertPayload {
  fullName: string;
  email: string;
  permissionScope?: string | null;
}

/** Create or update a Manager row in CleanTrack Users (Role = Manager, Active = Yes). */
export async function upsertManagerUser(
  accessToken: string,
  payload: ManagerUpsertPayload
): Promise<void> {
  const siteId = await sharepoint.getSiteId(accessToken);
  const listId = await sharepoint.getListIdByName(accessToken, siteId, CLEANTRACK_USERS_LIST_NAME);
  if (!listId) throw new Error(`List "${CLEANTRACK_USERS_LIST_NAME}" not found.`);

  const [items, columns] = await Promise.all([
    sharepoint.getListItems(accessToken, siteId, listId),
    sharepoint.getListColumns(accessToken, siteId, listId),
  ]);

  const needle = payload.email.trim().toLowerCase();
  const existing = items.find((item) => {
    const e = String((item.fields as any)?.Email ?? "").trim().toLowerCase();
    return e === needle;
  });

  const map: Record<string, string> = {};
  for (const c of columns) {
    if (c.displayName) map[c.displayName] = c.name;
  }

  const nameKey = map["Title"] ?? map["Full Name"] ?? "Title";
  const emailKey = map["Email"] ?? "Email";
  const roleKey = map["Role"] ?? "Role";
  const activeKey = map["Active"] ?? "Active";
  const permScopeKey = map["Permission Scope"] ?? "PermissionScope";

  const fields: Record<string, unknown> = {
    [nameKey]: payload.fullName || payload.email,
    [emailKey]: payload.email,
    [roleKey]: "Manager",
    [activeKey]: true,
  };
  if (payload.permissionScope !== undefined) {
    fields[permScopeKey] = payload.permissionScope ?? null;
  }

  if (existing?.id) {
    await sharepoint.updateListItem(accessToken, siteId, listId, existing.id, fields);
  } else {
    await sharepoint.createListItem(accessToken, siteId, listId, fields);
  }
}

export interface UserUpsertPayload {
  fullName: string;
  email: string;
  role: "Admin" | "Manager";
  active: boolean;
  permissionScope?: string | null;
}

/** Create or update any CleanTrack User (Admin/Manager). */
export async function upsertUser(
  accessToken: string,
  payload: UserUpsertPayload
): Promise<void> {
  const siteId = await sharepoint.getSiteId(accessToken);
  const listId = await sharepoint.getListIdByName(accessToken, siteId, CLEANTRACK_USERS_LIST_NAME);
  if (!listId) throw new Error(`List "${CLEANTRACK_USERS_LIST_NAME}" not found.`);

  const [items, columns] = await Promise.all([
    sharepoint.getListItems(accessToken, siteId, listId),
    sharepoint.getListColumns(accessToken, siteId, listId),
  ]);

  const needle = payload.email.trim().toLowerCase();
  const existing = items.find((item) => {
    const e = String((item.fields as any)?.Email ?? "").trim().toLowerCase();
    return e === needle;
  });

  const map: Record<string, string> = {};
  for (const c of columns) {
    if (c.displayName) map[c.displayName] = c.name;
  }

  const nameKey = map["Title"] ?? map["Full Name"] ?? "Title";
  const emailKey = map["Email"] ?? "Email";
  const roleKey = map["Role"] ?? "Role";
  const activeKey = map["Active"] ?? "Active";
  const permScopeKey = map["Permission Scope"] ?? "PermissionScope";

  const fields: Record<string, unknown> = {
    [nameKey]: payload.fullName || payload.email,
    [emailKey]: payload.email,
    [roleKey]: payload.role,
    [activeKey]: payload.active,
  };
  if (payload.permissionScope !== undefined) {
    fields[permScopeKey] = payload.permissionScope ?? null;
  }

  if (existing?.id) {
    await sharepoint.updateListItem(accessToken, siteId, listId, existing.id, fields);
  } else {
    await sharepoint.createListItem(accessToken, siteId, listId, fields);
  }
}
