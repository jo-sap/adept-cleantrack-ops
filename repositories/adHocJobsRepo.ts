/**
 * CleanTrack Ad Hoc Jobs – fetch, create, update.
 * List: "CleanTrack Ad Hoc Jobs"
 * Job Name = SharePoint Title. Site = lookup to CleanTrack Sites. Assigned Manager = lookup to CleanTrack Users.
 * All relationship keys use lookup IDs only.
 */

import * as sharepoint from "../lib/sharepoint";
import { getCleanTrackUserIdToNameMap } from "./usersRepo";
import type { AdHocJob } from "../types";

const AD_HOC_JOBS_LIST_NAME = "CleanTrack Ad Hoc Jobs";

/** Temporary debug logs. Set to false to disable. */
const DEBUG_ADHOC = true;

export interface AdHocJobFilters {
  /** Filter by requested date month (year-month string "YYYY-MM"). */
  month?: string;
  status?: string;
  assignedManagerId?: string;
  siteId?: string;
}

function toDateStr(v: unknown): string | null {
  if (v == null) return null;
  if (typeof v === "string") {
    const s = v.trim();
    if (!s) return null;
    const d = new Date(s);
    return isNaN(d.getTime()) ? null : d.toISOString().slice(0, 10);
  }
  return null;
}

function toNum(v: unknown): number | null {
  if (v == null || v === "") return null;
  if (typeof v === "number" && !Number.isNaN(v)) return v;
  const n = parseFloat(String(v).replace(/[^0-9.-]/g, ""));
  return Number.isNaN(n) ? null : n;
}

function toBool(v: unknown): boolean {
  if (v === true || v === "Yes") return true;
  if (typeof v === "string") return v.trim().toLowerCase() === "yes";
  return false;
}

function getStr(f: Record<string, unknown>, ...keys: string[]): string {
  for (const k of keys) {
    const v = f[k];
    if (v != null && v !== "") return String(v).trim();
  }
  return "";
}

function getLookupId(f: Record<string, unknown>, baseKey: string): string {
  const idKey = Object.keys(f).find(
    (k) => k === `${baseKey}LookupId` || k === `${baseKey}Id` || k.toLowerCase() === `${baseKey.toLowerCase()}lookupid`
  );
  if (!idKey) return "";
  const v = f[idKey];
  if (v == null) return "";
  if (typeof v === "number") return String(v);
  return String(v).trim();
}

function getLookupValue(f: Record<string, unknown>, baseKey: string): string {
  const nameKey = Object.keys(f).find(
    (k) => k === baseKey || k.toLowerCase() === baseKey.toLowerCase()
  );
  if (!nameKey) return "";
  const v = f[nameKey];
  if (v == null) return "";
  if (typeof v === "object" && v !== null && "LookupValue" in (v as object))
    return String((v as { LookupValue?: string }).LookupValue ?? "").trim();
  return String(v).trim();
}

/** Build display name -> internal name map from list columns. */
async function getFieldMap(
  accessToken: string,
  siteId: string,
  listId: string
): Promise<Record<string, string>> {
  const columns = await sharepoint.getListColumns(accessToken, siteId, listId);
  const map: Record<string, string> = {};
  for (const c of columns) {
    if (c.displayName) map[c.displayName] = c.name;
  }
  return map;
}

/** Map SharePoint list item to AdHocJob. Resolve manager name from CleanTrack Users when possible. */
function itemToAdHocJob(
  item: sharepoint.GraphListItem,
  map: Record<string, string>,
  managerIdToName: Record<string, string>
): AdHocJob {
  const f = (item.fields ?? {}) as Record<string, unknown>;
  const titleKey = map["Job Name"] ?? "Title";
  const jobName = getStr(f, titleKey, "Title");

  const siteId = getLookupId(f, "Site");
  const siteName = getLookupValue(f, "Site");
  const assignedManagerId = getLookupId(f, "Assigned Manager");
  let assignedManagerName = getLookupValue(f, "Assigned Manager");
  if (!assignedManagerName && assignedManagerId) {
    const norm = sharepoint.normalizeListItemId(assignedManagerId);
    assignedManagerName = managerIdToName[norm] ?? managerIdToName[assignedManagerId] ?? "";
  }

  const reqNameKey = map["Requested By Name"] ?? "Requested_x0020_By_x0020_Name";
  const reqEmailKey = map["Requested By Email"] ?? "Requested_x0020_By_x0020_Email";
  const reqCompanyKey = map["Requested By Company"] ?? "Requested_x0020_By_x0020_Company";
  const channelKey = map["Request Channel"] ?? "Request_x0020_Channel";
  const summaryKey = map["Request Summary"] ?? "Request_x0020_Summary";
  const reqDateKey = map["Requested Date"] ?? "Requested_x0020_Date";
  const schedKey = map["Scheduled Date"] ?? "Scheduled_x0020_Date";
  const completedKey = map["Completed Date"] ?? "Completed_x0020_Date";
  const statusKey = map["Status"] ?? "Status";
  const jobTypeKey = map["Job Type"] ?? "Job_x0020_Type";
  const budgetHrsKey = map["Budgeted Hours"] ?? "Budgeted_x0020_Hours";
  const budgetRateKey = map["Budgeted Labour Rate"] ?? "Budgeted_x0020_Labour_x0020_Rate";
  const budgetRevKey = map["Budgeted Revenue"] ?? "Budgeted_x0020_Revenue";
  const descKey = map["Description"] ?? "Description";
  const proofReqKey = map["Approval Proof Required"] ?? "Approval_x0020_Proof_x0020_Required";
  const proofUpKey = map["Approval Proof Uploaded"] ?? "Approval_x0020_Proof_x0020_Uploaded";
  const proofNotesKey = map["Approval Reference Notes"] ?? "Approval_x0020_Reference_x0020_Notes";
  const activeKey = map["Active"] ?? "Active";

  return {
    id: sharepoint.normalizeListItemId(item.id),
    jobName,
    jobType: getStr(f, jobTypeKey, "Job Type"),
    siteId: siteId || null,
    siteName: siteName || "",
    requestedByName: getStr(f, reqNameKey),
    requestedByEmail: getStr(f, reqEmailKey),
    requestedByCompany: getStr(f, reqCompanyKey),
    requestChannel: getStr(f, channelKey),
    requestSummary: getStr(f, summaryKey),
    requestedDate: toDateStr(f[reqDateKey] ?? f["RequestedDate"]),
    assignedManagerId: assignedManagerId || null,
    assignedManagerName,
    scheduledDate: toDateStr(f[schedKey] ?? f["ScheduledDate"]),
    completedDate: toDateStr(f[completedKey] ?? f["CompletedDate"]),
    status: getStr(f, statusKey, "Status"),
    budgetedHours: toNum(f[budgetHrsKey] ?? f["BudgetedHours"]),
    budgetedLabourRate: toNum(f[budgetRateKey] ?? f["BudgetedLabourRate"]),
    budgetedRevenue: toNum(f[budgetRevKey] ?? f["BudgetedRevenue"]),
    description: getStr(f, descKey),
    approvalProofRequired: toBool(f[proofReqKey] ?? f["ApprovalProofRequired"]),
    approvalProofUploaded: toBool(f[proofUpKey] ?? f["ApprovalProofUploaded"]),
    approvalReferenceNotes: getStr(f, proofNotesKey),
    active: toBool(f[activeKey] ?? f["Active"]) !== false,
  };
}

/**
 * Fetch Ad Hoc Jobs with optional filters. Manager role: pass assignedManagerId to restrict to their jobs.
 */
export async function getAdHocJobs(
  accessToken: string,
  filters?: AdHocJobFilters
): Promise<AdHocJob[]> {
  const siteId = await sharepoint.getSiteId(accessToken);
  const listId = await sharepoint.getListIdByName(accessToken, siteId, AD_HOC_JOBS_LIST_NAME);
  if (!listId) return [];

  const [items, map, managerIdToName] = await Promise.all([
    sharepoint.getListItems(accessToken, siteId, listId),
    getFieldMap(accessToken, siteId, listId),
    getCleanTrackUserIdToNameMap(accessToken),
  ]);

  if (DEBUG_ADHOC && items.length > 0) {
    console.log("[CleanTrack Ad Hoc Jobs] raw record count:", items.length);
    console.log("[CleanTrack Ad Hoc Jobs] sample raw:", JSON.stringify(items[0], null, 2));
    console.log("[CleanTrack Ad Hoc Jobs] field map keys:", Object.keys(map));
  }

  let list: AdHocJob[] = items.map((item) =>
    itemToAdHocJob(item, map, managerIdToName)
  );

  if (DEBUG_ADHOC && list.length > 0) {
    console.log("[CleanTrack Ad Hoc Jobs] mapped sample:", JSON.stringify(list[0], null, 2));
  }

  if (filters?.month) {
    const [y, m] = filters.month.split("-").map(Number);
    list = list.filter((j) => {
      const d = j.requestedDate;
      if (!d) return false;
      const [ey, em] = d.split("-").map(Number);
      return ey === y && em === m;
    });
    if (DEBUG_ADHOC) console.log("[CleanTrack Ad Hoc Jobs] after month filter:", filters.month, "count:", list.length);
  }
  if (filters?.status) {
    list = list.filter((j) => j.status.toLowerCase() === filters!.status!.toLowerCase());
  }
  if (filters?.assignedManagerId) {
    const norm = sharepoint.normalizeListItemId(filters.assignedManagerId);
    list = list.filter((j) => j.assignedManagerId && sharepoint.normalizeListItemId(j.assignedManagerId) === norm);
  }
  if (filters?.siteId) {
    const norm = sharepoint.normalizeListItemId(filters.siteId);
    list = list.filter((j) => j.siteId && sharepoint.normalizeListItemId(j.siteId) === norm);
  }

  return list;
}

/** Payload for create/update. Use lookup IDs for Site and Assigned Manager. */
export interface AdHocJobPayload {
  jobName: string;
  jobType?: string;
  siteId?: string | null;
  assignedManagerId?: string | null;
  requestedByName?: string;
  requestedByEmail?: string;
  requestedByCompany?: string;
  requestChannel?: string;
  requestSummary?: string;
  requestedDate?: string | null;
  scheduledDate?: string | null;
  completedDate?: string | null;
  status?: string;
  budgetedHours?: number | null;
  budgetedLabourRate?: number | null;
  budgetedRevenue?: number | null;
  description?: string;
  approvalProofRequired?: boolean;
  approvalProofUploaded?: boolean;
  approvalReferenceNotes?: string;
  active?: boolean;
}

function payloadToFields(
  payload: AdHocJobPayload,
  map: Record<string, string>
): Record<string, unknown> {
  const titleKey = map["Job Name"] ?? "Title";
  const fields: Record<string, unknown> = {};
  if (payload.jobName !== undefined) fields[titleKey === "LinkTitle" ? "Title" : titleKey] = payload.jobName;
  if (payload.jobType !== undefined) fields[map["Job Type"] ?? "Job_x0020_Type"] = payload.jobType ?? "";

  const siteInternal = map["Site"] ?? "Site";
  const siteLookupKey = siteInternal === "Site" ? "SiteLookupId" : `${siteInternal}LookupId`;
  if (payload.siteId !== undefined && payload.siteId != null && payload.siteId !== "") {
    const num = /^\d+$/.test(payload.siteId) ? parseInt(payload.siteId, 10) : payload.siteId;
    fields[siteLookupKey] = num;
  }

  const mgrInternal = map["Assigned Manager"] ?? "Assigned_x0020_Manager";
  const mgrLookupKey = mgrInternal === "Assigned Manager" ? "Assigned_x0020_ManagerLookupId" : `${mgrInternal}LookupId`;
  if (payload.assignedManagerId !== undefined && payload.assignedManagerId != null && payload.assignedManagerId !== "") {
    const num = /^\d+$/.test(payload.assignedManagerId) ? parseInt(payload.assignedManagerId, 10) : payload.assignedManagerId;
    fields[mgrLookupKey] = num;
  }

  const set = (displayName: string, internalFallback: string, value: unknown) => {
    const k = map[displayName] ?? internalFallback;
    if (k && value !== undefined) fields[k] = value;
  };
  set("Requested By Name", "Requested_x0020_By_x0020_Name", payload.requestedByName ?? "");
  set("Requested By Email", "Requested_x0020_By_x0020_Email", payload.requestedByEmail ?? "");
  set("Requested By Company", "Requested_x0020_By_x0020_Company", payload.requestedByCompany ?? "");
  set("Request Channel", "Request_x0020_Channel", payload.requestChannel ?? "");
  set("Request Summary", "Request_x0020_Summary", payload.requestSummary ?? "");
  set("Requested Date", "Requested_x0020_Date", payload.requestedDate ?? null);
  set("Scheduled Date", "Scheduled_x0020_Date", payload.scheduledDate ?? null);
  set("Completed Date", "Completed_x0020_Date", payload.completedDate ?? null);
  set("Status", "Status", payload.status ?? "Requested");
  set("Budgeted Hours", "Budgeted_x0020_Hours", payload.budgetedHours ?? null);
  set("Budgeted Labour Rate", "Budgeted_x0020_Labour_x0020_Rate", payload.budgetedLabourRate ?? null);
  set("Budgeted Revenue", "Budgeted_x0020_Revenue", payload.budgetedRevenue ?? null);
  set("Description", "Description", payload.description ?? "");
  set("Approval Proof Required", "Approval_x0020_Proof_x0020_Required", payload.approvalProofRequired ?? false);
  set("Approval Proof Uploaded", "Approval_x0020_Proof_x0020_Uploaded", payload.approvalProofUploaded ?? false);
  set("Approval Reference Notes", "Approval_x0020_Reference_x0020_Notes", payload.approvalReferenceNotes ?? "");
  set("Active", "Active", payload.active !== false);
  return fields;
}

/** Resolve internal column names for Ad Hoc Jobs list (for create/update). */
async function getAdHocFieldMap(accessToken: string): Promise<Record<string, string>> {
  const siteId = await sharepoint.getSiteId(accessToken);
  const listId = await sharepoint.getListIdByName(accessToken, siteId, AD_HOC_JOBS_LIST_NAME);
  if (!listId) throw new Error(`List "${AD_HOC_JOBS_LIST_NAME}" not found.`);
  return getFieldMap(accessToken, siteId, listId);
}

export async function createAdHocJob(
  accessToken: string,
  payload: AdHocJobPayload
): Promise<AdHocJob> {
  const siteId = await sharepoint.getSiteId(accessToken);
  const listId = await sharepoint.getListIdByName(accessToken, siteId, AD_HOC_JOBS_LIST_NAME);
  if (!listId) throw new Error(`List "${AD_HOC_JOBS_LIST_NAME}" not found.`);

  const map = await getAdHocFieldMap(accessToken);
  const full: AdHocJobPayload = {
    status: "Requested",
    active: true,
    ...payload,
  };
  const fields = payloadToFields(full, map);

  if (DEBUG_ADHOC) {
    console.log("[CleanTrack Ad Hoc Jobs] create payload fields:", JSON.stringify(fields, null, 2));
  }

  const created = await sharepoint.createListItem(accessToken, siteId, listId, fields);
  const managerIdToName = await getCleanTrackUserIdToNameMap(accessToken);
  return itemToAdHocJob(created, map, managerIdToName);
}

export async function updateAdHocJob(
  accessToken: string,
  itemId: string,
  payload: Partial<AdHocJobPayload>
): Promise<void> {
  const siteId = await sharepoint.getSiteId(accessToken);
  const listId = await sharepoint.getListIdByName(accessToken, siteId, AD_HOC_JOBS_LIST_NAME);
  if (!listId) throw new Error(`List "${AD_HOC_JOBS_LIST_NAME}" not found.`);

  const map = await getAdHocFieldMap(accessToken);
  const fields = payloadToFields(payload as AdHocJobPayload, map);
  if (Object.keys(fields).length === 0) return;

  if (DEBUG_ADHOC) {
    console.log("[CleanTrack Ad Hoc Jobs] update itemId:", itemId, "fields:", JSON.stringify(fields, null, 2));
  }

  await sharepoint.updateListItem(accessToken, siteId, listId, itemId, fields);
}
