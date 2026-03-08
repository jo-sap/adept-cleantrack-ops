# Ad Hoc Jobs – Implementation Summary

## Files changed / created

| File | Change |
|------|--------|
| `types.ts` | Added `AdHocJob` interface; added `adhocJobId?`, `adhocJobName?` to `TimeEntry`; added `'adhoc-jobs'` to `ViewType`. |
| `repositories/adHocJobsRepo.ts` | **New.** Fetch, create, update CleanTrack Ad Hoc Jobs; field map; filters (month, status, manager, site). |
| `repositories/metricsRepo.ts` | `TimesheetEntryFlat` and fetch: added `adhocJobId`, `adhocJobName`; resolve Ad Hoc Job column by list columns; debug log. `TimesheetEntryPayload`: added `adhocJobId?`; save create/update: include Ad Hoc Job lookup. |
| `components/AdHocJobsManager.tsx` | **New.** Ad Hoc Jobs page: list, month/status/manager/site filters, create/edit modal with full form. |
| `components/Sidebar.tsx` | Added "Ad Hoc Jobs" nav item (Briefcase icon). |
| `App.tsx` | Import `AdHocJobsManager`; route `case 'adhoc-jobs'`; `max-w-7xl` for adhoc-jobs; graph entries map: include `adhocJobId`, `adhocJobName`; save payload: include `adhocJobId`. |
| `components/TimeEntryForm.tsx` | Optional "Ad Hoc Job" dropdown filtered by selected site; state `adhocJobId`, `adHocJobsForSite`; load jobs when site changes; include `adhocJobId` in batch save payload. |
| `components/Dashboard.tsx` | Ad Hoc Jobs summary block: total, completed, pending, budgeted vs actual hours for current month; fetch ad hoc jobs by month (and manager when not admin). |

---

## SharePoint field mapping (CleanTrack Ad Hoc Jobs)

Assumptions: list display name is **"CleanTrack Ad Hoc Jobs"**. Internal names are resolved at runtime via `getListColumns`; below are fallbacks used when display name → internal name map is missing.

| Display name | Internal name (fallback) | Notes |
|--------------|--------------------------|--------|
| Job Name | Title | SharePoint Title; create/update use Title when not LinkTitle. |
| Job Type | Job_x0020_Type | |
| Site | Site | Lookup to CleanTrack Sites; id sent as `SiteLookupId` (numeric). |
| Requested By Name | Requested_x0020_By_x0020_Name | |
| Requested By Email | Requested_x0020_By_x0020_Email | |
| Requested By Company | Requested_x0020_By_x0020_Company | |
| Request Channel | Request_x0020_Channel | |
| Request Summary | Request_x0020_Summary | |
| Requested Date | Requested_x0020_Date | |
| Assigned Manager | Assigned_x0020_Manager | Lookup to CleanTrack Users; id sent as `Assigned_x0020_ManagerLookupId` (numeric). |
| Scheduled Date | Scheduled_x0020_Date | |
| Completed Date | Completed_x0020_Date | |
| Status | Status | |
| Budgeted Hours | Budgeted_x0020_Hours | |
| Budgeted Labour Rate | Budgeted_x0020_Labour_x0020_Rate | |
| Budgeted Revenue | Budgeted_x0020_Revenue | |
| Description | Description | |
| Approval Proof Required | Approval_x0020_Proof_x0020_Required | |
| Approval Proof Uploaded | Approval_x0020_Proof_x0020_Uploaded | |
| Approval Reference Notes | Approval_x0020_Reference_x0020_Notes | |
| Active | Active | |

---

## Create Ad Hoc Job – payload structure

Example payload sent to `sharepoint.createListItem` (field keys are internal names from column map):

```json
{
  "Title": "Carpet clean – Building A",
  "Job_x0020_Type": "Carpet clean",
  "SiteLookupId": 5,
  "Assigned_x0020_ManagerLookupId": 12,
  "Requested_x0020_By_x0020_Name": "Jane Smith",
  "Requested_x0020_By_x0020_Email": "jane@example.com",
  "Requested_x0020_By_x0020_Company": "Acme",
  "Request_x0020_Channel": "Email",
  "Request_x0020_Summary": "Deep clean foyer",
  "Requested_x0020_Date": "2025-03-01",
  "Scheduled_x0020_Date": null,
  "Completed_x0020_Date": null,
  "Status": "Requested",
  "Budgeted_x0020_Hours": 4,
  "Budgeted_x0020_Labour_x0020_Rate": 28,
  "Budgeted_x0020_Revenue": 150,
  "Description": "",
  "Approval_x0020_Proof_x0020_Required": true,
  "Approval_x0020_Proof_x0020_Uploaded": false,
  "Approval_x0020_Reference_x0020_Notes": "",
  "Active": true
}
```

- **Title** = Job Name (required).
- **SiteLookupId** = CleanTrack Sites list item id (number).
- **Assigned_x0020_ManagerLookupId** = CleanTrack Users list item id (number).

---

## Create / update Timesheet Entry with Ad Hoc Job – payload structure

**Create** (new item):

- `Work Date`, `Hours`, `SiteLookupId`, `CleanerLookupId` as today.
- When an Ad Hoc Job is selected: `Ad_x0020_Hoc_x0020_JobLookupId` (or resolved from list columns) = Ad Hoc Jobs list item id (number).

**Update** (existing item):

- `Hours` always; when payload has `adhocJobId`, the Ad Hoc Job lookup field is also sent (or cleared if null).

Example create fields:

```json
{
  "WorkDate": "2025-03-05",
  "Hours": 3,
  "SiteLookupId": 5,
  "CleanerLookupId": 8,
  "Ad_x0020_Hoc_x0020_JobLookupId": 2
}
```

---

## Role-based visibility

- **Admin:** All ad hoc jobs; filters (month, status, manager, site) apply to full list.
- **Manager:** Only jobs where `Assigned Manager` = current user (CleanTrack Users list item id from `getCleanTrackUserByEmail`). Month filter still applies; manager/site filters are not shown in UI (they only see their own jobs).

---

## Debug logs

- **adHocJobsRepo:** `DEBUG_ADHOC` (default true): raw record count, sample raw record, field map keys, mapped sample, create/update payload fields, month filter count.
- **metricsRepo:** In development, logs timesheet range count, count with Ad Hoc Job, and sample entry.

Set `DEBUG_ADHOC = false` in `repositories/adHocJobsRepo.ts` to turn off ad hoc logs.

---

## TODOs / limitations

1. **Attachments:** List supports attachments for approval proof; app does not yet upload or list attachments. Form has "Approval proof uploaded" checkbox; actual file upload would require Graph drive/item attachment API.
2. **Month filter:** Ad Hoc Jobs list filter uses **Requested Date** month. No extra SharePoint column for month.
3. **Timesheet key:** Existing timesheet match is still `siteId|cleanerId|date`. Ad Hoc Job is an optional attribute on the same row; changing Ad Hoc Job on an existing row is done by updating that row with the new lookup (or null).
4. **Dashboard actual hours:** "Actual hrs" for ad hoc is sum of entries in the **current period** that have `adhocJobId` set. Period is the fortnight; ad hoc stats are by **month** (requested date). So "Budgeted / Actual hrs" compares month job budgeted hours to period-linked actual hours for simplicity.
5. **Approval proof:** No hard block; missing proof is indicated with an "Missing" badge and AlertCircle icon when Approval Proof Required is true and Approval Proof Uploaded is false.

---

## Recurring site work unchanged

- Normal site budgets, recurring hours, and timesheets work as before.
- Ad Hoc Job on timesheet is optional; leave "Ad Hoc Job" as "None – recurring work" for normal work.
- No refactors to existing site/cleaner/timesheet logic beyond adding optional `adhocJobId` to payload and response.
