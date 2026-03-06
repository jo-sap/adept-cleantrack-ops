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

export interface GraphListItem {
  id: string;
  fields?: Record<string, unknown>;
}

export interface GraphListItemsResponse {
  value?: GraphListItem[];
}

/** Get list items with fields expanded. */
export async function getListItems(
  accessToken: string,
  siteId: string,
  listId: string
): Promise<GraphListItem[]> {
  const url = `${GRAPH_BASE}/sites/${siteId}/lists/${listId}/items?expand=fields`;
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
