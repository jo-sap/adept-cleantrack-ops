/**
 * SharePoint read-only helpers via Microsoft Graph.
 * Requires a valid Graph access token (Sites.Read.All or Sites.ReadWrite.All).
 * Authorization header must use the ACCESS TOKEN from getGraphAccessToken() (result.accessToken), not idToken.
 */

import { decodeToken } from "./graph";

const GRAPH_BASE = "https://graph.microsoft.com/v1.0";
const SHAREPOINT_SITE_PATH = "adeptservicesaustralia.sharepoint.com:/sites/cleantrack";

/** Microsoft Graph access token audience (aud claim). */
const GRAPH_AUD = "https://graph.microsoft.com";
const GRAPH_AUD_GUID = "00000003-0000-0000-c000-000000000000";

const IS_DEV =
  typeof import.meta !== "undefined" && import.meta.env && (import.meta.env as { DEV?: boolean }).DEV === true;

function assertGraphAccessToken(accessToken: string): void {
  if (!IS_DEV) return;
  const { aud, scp } = decodeToken(accessToken);
  console.log("POST/PATCH token aud:", aud, "scp:", scp);
  if (aud && aud !== GRAPH_AUD && aud !== GRAPH_AUD_GUID) {
    throw new Error("Wrong token audience; expected Microsoft Graph access token (aud: https://graph.microsoft.com). Got aud: " + aud);
  }
}

async function graphFetch<T>(
  accessToken: string,
  url: string,
  options?: { method?: string; body?: string }
): Promise<T> {
  const res = await fetch(url, {
    method: options?.method ?? "GET",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      ...(options?.body ? { "Content-Type": "application/json" } : {}),
    },
    ...(options?.body ? { body: options.body } : {}),
  });
  const text = await res.text();
  if (!res.ok) {
    console.error("[SP] Graph error", res.status, text);
    throw new Error(`Graph ${res.status}: ${text}`);
  }
  let data: unknown;
  try {
    data = text ? JSON.parse(text) : {};
  } catch {
    throw new Error(`Graph request failed: ${res.status} ${res.statusText}`);
  }
  return data as T;
}

export interface GraphSiteResponse {
  id: string;
  displayName?: string;
  webUrl?: string;
}

/** Resolve the CleanTrack SharePoint site ID. Uses EXACT endpoint: GET .../sites/adeptservicesaustralia.sharepoint.com:/sites/cleantrack */
export async function getSiteId(accessToken: string): Promise<string> {
  const url = `${GRAPH_BASE}/sites/${SHAREPOINT_SITE_PATH}`;
  const site = await graphFetch<GraphSiteResponse>(accessToken, url);
  if (!site?.id) throw new Error("Site response missing id");
  const siteId = site.id;
  if (IS_DEV) console.log("[SP] resolved siteId:", siteId);
  return siteId;
}

/** Get the site's web URL (for SharePoint REST _api calls). Same token as Graph. */
export async function getSiteWebUrl(accessToken: string): Promise<string> {
  const siteId = await getSiteId(accessToken);
  const url = `${GRAPH_BASE}/sites/${siteId}?$select=webUrl`;
  const site = await graphFetch<GraphSiteResponse>(accessToken, url);
  const webUrl = site?.webUrl?.trim();
  if (!webUrl) throw new Error("Site webUrl not found");
  return webUrl.replace(/\/$/, "");
}

export interface GraphListEntry {
  id: string;
  displayName: string;
}

export interface GraphListsResponse {
  value?: GraphListEntry[];
}

/** List all lists in the site. */
export async function getLists(
  accessToken: string,
  siteId: string
): Promise<Array<{ id: string; displayName: string }>> {
  const url = `${GRAPH_BASE}/sites/${siteId}/lists`;
  const data = await graphFetch<GraphListsResponse>(accessToken, url);
  const value = data.value ?? [];
  return value.map((l) => ({ id: l.id, displayName: l.displayName ?? "" }));
}

/** Find list id by display name (case-insensitive exact). */
export async function getListIdByName(
  accessToken: string,
  siteId: string,
  displayName: string
): Promise<string | null> {
  const lists = await getLists(accessToken, siteId);
  const name = displayName.trim().toLowerCase();
  const found = lists.find((l) => l.displayName.trim().toLowerCase() === name);
  return found?.id ?? null;
}

/** Get the site's User Information List id (for Person columns). Uses $filter so hidden list is found. */
export async function getSiteUserInformationListId(
  accessToken: string,
  siteId: string
): Promise<string | null> {
  const names = ["User Information List", "Users", "UserInfo"];
  for (const displayName of names) {
    const encoded = encodeURIComponent(`displayName eq '${displayName.replace(/'/g, "''")}'`);
    const url = `${GRAPH_BASE}/sites/${siteId}/lists?$filter=${encoded}`;
    try {
      const data = await graphFetch<GraphListsResponse>(accessToken, url);
      const list = data.value?.[0];
      if (list?.id) return list.id;
    } catch {
      // try next name
    }
  }
  return getListIdByName(accessToken, siteId, "User Information List");
}

export interface GraphListColumn {
  name: string;
  displayName: string;
}

export interface GraphListColumnsResponse {
  value?: GraphListColumn[];
}

/** Get list columns (internal name and display name). GET /sites/{siteId}/lists/{listId}/columns */
export async function getListColumns(
  accessToken: string,
  siteId: string,
  listId: string
): Promise<Array<{ name: string; displayName: string }>> {
  const url = `${GRAPH_BASE}/sites/${siteId}/lists/${listId}/columns`;
  const data = await graphFetch<GraphListColumnsResponse>(accessToken, url);
  const value = data.value ?? [];
  return value.map((c) => ({ name: c.name ?? "", displayName: c.displayName ?? "" }));
}

/** Full column definition including type (lookup vs personOrGroup). Used to detect Manager column type. */
export interface GraphColumnDefinition {
  name: string;
  displayName: string;
  lookup?: unknown;
  personOrGroup?: unknown;
}

interface GraphColumnDefinitionsResponse {
  value?: GraphColumnDefinition[];
}

/** Get list column definitions including type (lookup / personOrGroup) so we can set Manager correctly. */
export async function getListColumnDefinitions(
  accessToken: string,
  siteId: string,
  listId: string
): Promise<GraphColumnDefinition[]> {
  const url = `${GRAPH_BASE}/sites/${siteId}/lists/${listId}/columns`;
  const data = await graphFetch<GraphColumnDefinitionsResponse>(accessToken, url);
  const value = data.value ?? [];
  return value.map((c) => ({
    name: c.name ?? "",
    displayName: c.displayName ?? "",
    lookup: c.lookup,
    personOrGroup: c.personOrGroup,
  }));
}

/** Get the site's User Information List lookup id for a user by email (for Person/Group columns). Returns null if not found. */
export async function getSiteUserLookupIdByEmail(
  accessToken: string,
  siteId: string,
  email: string
): Promise<string | null> {
  const listId = await getSiteUserInformationListId(accessToken, siteId);
  if (!listId) return null;
  const needle = email.trim().toLowerCase();
  if (!needle) return null;
  const tryFilter = async (fieldName: string): Promise<GraphListItem[]> => {
    try {
      const filter = `fields/${fieldName} eq '${needle.replace(/'/g, "''")}'`;
      return await getListItemsByFilter(accessToken, siteId, listId, filter);
    } catch {
      return [];
    }
  };
  for (const fieldName of ["EMail", "Email", "email"]) {
    const items = await tryFilter(fieldName);
    const match = items.find((item) => {
      const e = (item.fields?.EMail ?? item.fields?.Email ?? item.fields?.email) as string | undefined;
      return e && String(e).trim().toLowerCase() === needle;
    });
    if (match?.id) return normalizeListItemId(match.id);
    if (items.length > 0 && items[0].id) return normalizeListItemId(items[0].id);
  }
  try {
    const all = await getListItems(accessToken, siteId, listId);
    const match = all.find((item) => {
      const e = (item.fields?.EMail ?? item.fields?.Email ?? item.fields?.email) as string | undefined;
      return e && String(e).trim().toLowerCase() === needle;
    });
    if (match?.id) return normalizeListItemId(match.id);
  } catch {
    // ignore
  }
  return null;
}

export interface GraphListItem {
  id: string;
  fields?: Record<string, unknown>;
}

/** Normalize list item id for consistent join (Graph may return "5" or path-style). */
export function normalizeListItemId(id: string): string {
  const s = String(id ?? "").trim();
  if (!s) return s;
  const parts = s.split("/");
  const last = parts[parts.length - 1];
  return last ?? s;
}

export interface GraphListItemsResponse {
  value?: GraphListItem[];
}

/** Get list items with fields expanded. Optionally request specific fields (e.g. for lookup columns). */
export async function getListItems(
  accessToken: string,
  siteId: string,
  listId: string,
  fieldsSelect?: string[]
): Promise<GraphListItem[]> {
  let url = `${GRAPH_BASE}/sites/${siteId}/lists/${listId}/items?expand=fields`;
  if (fieldsSelect && fieldsSelect.length > 0) {
    const select = fieldsSelect.join(",");
    url += `($select=${select})`;
  }
  const data = await graphFetch<GraphListItemsResponse>(accessToken, url);
  const value = data.value ?? [];
  return value.map((item) => ({
    id: item.id,
    fields: item.fields ?? {},
  }));
}

/**
 * Create a list item. POST https://graph.microsoft.com/v1.0/sites/{siteId}/lists/{listId}/items
 * accessToken must be the Graph ACCESS TOKEN from getGraphAccessToken() (not idToken).
 * body: { fields: { "ColumnInternalName": value, ... } }
 */
export async function createListItem(
  accessToken: string,
  siteId: string,
  listId: string,
  fields: Record<string, unknown>
): Promise<GraphListItem> {
  assertGraphAccessToken(accessToken);
  const url = `${GRAPH_BASE}/sites/${siteId}/lists/${listId}/items`;
  if (IS_DEV) {
    console.log("[SP] POST URL:", url);
    console.log("[SP] listId:", listId);
    console.log("[SP] fields keys:", Object.keys(fields));
  }
  const body = JSON.stringify({ fields });
  const data = await graphFetch<GraphListItem & { fields?: Record<string, unknown> }>(
    accessToken,
    url,
    { method: "POST", body }
  );
  return { id: data.id, fields: data.fields ?? {} };
}

/**
 * Update list item fields. PATCH https://graph.microsoft.com/v1.0/sites/{siteId}/lists/{listId}/items/{itemId}/fields
 * accessToken must be the Graph ACCESS TOKEN from getGraphAccessToken() (not idToken).
 */
export async function updateListItem(
  accessToken: string,
  siteId: string,
  listId: string,
  itemId: string,
  fields: Record<string, unknown>
): Promise<void> {
  assertGraphAccessToken(accessToken);
  const url = `${GRAPH_BASE}/sites/${siteId}/lists/${listId}/items/${itemId}/fields`;
  if (IS_DEV) {
    console.log("[SP] PATCH URL:", url);
    console.log("[SP] listId:", listId);
    console.log("[SP] fields keys:", Object.keys(fields));
  }
  await graphFetch<unknown>(accessToken, url, {
    method: "PATCH",
    body: JSON.stringify(fields),
  });
}

/**
 * Delete a list item. DELETE .../sites/{siteId}/lists/{listId}/items/{itemId}
 * Returns 204 No Content on success.
 */
export async function deleteListItem(
  accessToken: string,
  siteId: string,
  listId: string,
  itemId: string
): Promise<void> {
  assertGraphAccessToken(accessToken);
  const url = `${GRAPH_BASE}/sites/${siteId}/lists/${listId}/items/${itemId}`;
  const res = await fetch(url, {
    method: "DELETE",
    headers: { Authorization: `Bearer ${accessToken}` },
  });
  if (!res.ok) {
    const text = await res.text();
    console.error("[SP] Graph DELETE error", res.status, text);
    throw new Error(`Graph ${res.status}: ${text}`);
  }
}

/**
 * Get list items optionally filtered. Tries Graph $filter first; on error fetches all and filters in JS.
 * filterQuery should be OData filter on fields, e.g. "fields/Email eq 'user@domain.com'".
 */
export async function getListItemsByFilter(
  accessToken: string,
  siteId: string,
  listId: string,
  filterQuery: string
): Promise<GraphListItem[]> {
  const encoded = encodeURIComponent(filterQuery);
  const url = `${GRAPH_BASE}/sites/${siteId}/lists/${listId}/items?expand=fields&$filter=${encoded}`;
  try {
    const data = await graphFetch<GraphListItemsResponse>(accessToken, url);
    const value = data.value ?? [];
    return value.map((item) => ({
      id: item.id,
      fields: item.fields ?? {},
    }));
  } catch {
    const all = await getListItems(accessToken, siteId, listId);
    return all;
  }
}

/** SharePoint REST: get form digest for POST requests. POST {webUrl}/_api/contextinfo */
export async function getSharePointFormDigest(
  webUrl: string,
  accessToken: string
): Promise<string> {
  const url = `${webUrl}/_api/contextinfo`;
  const res = await fetch(url, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      Accept: "application/json;odata=verbose",
    },
  });
  const text = await res.text();
  if (!res.ok) {
    console.error("[SP REST] contextinfo error", res.status, text);
    throw new Error(`SharePoint contextinfo ${res.status}: ${text}`);
  }
  const data = text ? JSON.parse(text) : {};
  const digest =
    (data as { d?: { GetContextWebInformation?: { FormDigestValue?: string } } }).d?.GetContextWebInformation
      ?.FormDigestValue;
  if (!digest) throw new Error("Form digest not in response");
  return digest;
}

/**
 * SharePoint REST: add an attachment to a list item.
 * POST {webUrl}/_api/web/lists/GetByTitle('ListTitle')/items(itemId)/AttachmentFiles/add(FileName='name')
 * List must have attachments enabled. Graph API does not support list item attachments; REST does.
 */
export async function addListItemAttachment(
  webUrl: string,
  listTitle: string,
  itemId: string,
  fileName: string,
  fileBody: ArrayBuffer,
  accessToken: string,
  formDigest: string
): Promise<void> {
  const escapeForOData = (s: string) => s.replace(/'/g, "''");
  const safeTitle = escapeForOData(listTitle.trim());
  const safeName = escapeForOData(fileName.trim()) || "attachment";
  const itemIdNum = normalizeListItemId(itemId);
  const url = `${webUrl}/_api/web/lists/GetByTitle('${safeTitle}')/items(${itemIdNum})/AttachmentFiles/add(FileName='${safeName}')`;
  const res = await fetch(url, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "X-RequestDigest": formDigest,
      "Content-Type": "application/octet-stream",
    },
    body: fileBody,
  });
  if (!res.ok) {
    const errText = await res.text();
    console.error("[SP REST] add attachment error", res.status, errText);
    throw new Error(`Upload attachment failed: ${res.status} ${errText}`);
  }
}
