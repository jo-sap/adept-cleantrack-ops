# Adept CleanTrack Ops – Concept & Architecture

## What It Is

Adept CleanTrack Ops is an internal operations platform that:

- **Centralises** the cleaning portfolio (sites, managers, cleaners)
- **Tracks** time worked (timesheets) and budget vs actual
- **Produces** executive KPIs (revenue, labour cost, gross profit, margin)
- **Controls** access by role (Admin vs Manager) via Microsoft login + CleanTrack Users list
- **Uses** Microsoft 365 (Entra ID + SharePoint Lists) as the single source of truth

---

## Access & Roles

### Login

- **Microsoft (Entra ID) only** – no in-app email/password.
- After login, the app checks the **CleanTrack Users** SharePoint list:
  - **Full Name**, **Email**, **Role** (Admin/Manager), **Active**, **PermissionScope**, etc.
- Only users in the list with **Active = true** can use the app (allow-list).

### Admin

- Full portfolio: add/edit sites, add/edit cleaners, view all timesheets, full KPI dashboard.
- Can assign managers to sites (CleanTrack Site Managers list).

### Manager

- Sees **only sites assigned to them** (CleanTrack Site Managers).
- Dashboard KPIs and site list are scoped to those sites.
- Cannot add/edit sites or cleaners; read-only on portfolio, can use timesheets for their sites.

---

## SharePoint Lists (Data Model)

| List | Purpose |
|------|--------|
| **CleanTrack Users** | Allow-list + role (Admin/Manager). Checked after Microsoft login. |
| **CleanTrack Sites** | Master sites: Site Name, Address, State, Active, Monthly Revenue. |
| **CleanTrack Cleaners** | Cleaner directory: Cleaner Name, Pay Rate, Account Name, BSB, Account Number, Active. |
| **CleanTrack Timesheet Entries** | Actual hours: Entry Name, Site (lookup), Cleaner (lookup), Work Date, Hours, Fortnight Start, Notes. |
| **CleanTrack Timesheet Period Notes** | Manager note per site + fortnight: **Site** = lookup column → **CleanTrack Sites** (not plain text), **Period Start** (date), **Note** (multi-line). Optional: Title, Tags, Cleaner. |
| **CleanTrack Site Budgets** | Planned hours per site: Budget Name, Site, Mon–Sun Hours, Active. |
| **CleanTrack Site Managers** | Assignments: Assignment Name, Site (lookup), Manager (email/lookup), Active, Is Primary. |

---

## What’s Wired in the App

### Auth & Scoping

- **Microsoft login** → **getGraphAccessToken()** with `User.Read` + `Sites.ReadWrite.All`.
- **CleanTrack Users** lookup by email → role (Admin/Manager) and Active.
- **Manager scoping**: **getAssignedSiteIdsForManager(token, userEmail)** from CleanTrack Site Managers → filter sites, dashboard KPIs, and Sites & Budgets view to assigned sites only.

### Sites & Budgets

- **List sites** from CleanTrack Sites (SharePoint).
- **Add / Edit / Activate–Deactivate** sites (Admin only); writes use internal column names (getListColumns).
- **Managers**: read-only list of assigned sites only.

### Cleaner Team

- **List cleaners** from CleanTrack Cleaners.
- **Add cleaner** (Admin only); writes to SharePoint with correct column mapping (e.g. Title for name, not LinkTitle).

### Dashboard

- **Four KPIs** from **getDashboardMetrics** (SharePoint):
  - Portfolio Revenue (pro-rated Monthly Revenue from Sites)
  - Labor Expenses (Timesheet Entries × Cleaner Pay Rate)
  - Net Gross Profit, Profit Margin
- **Admin**: all sites. **Manager**: only assigned sites (optional `assignedSiteIds` in metrics).
- Site table and “View site” driven by **sites** from App (already filtered for Manager).

### Timesheets

- Load **CleanTrack Timesheet Entries** from SharePoint by fortnight; save via Graph (`saveTimesheetEntriesToSharePoint`).
- **Period notes** (managers/admins): **CleanTrack Timesheet Period Notes** — one note per site per fortnight on the timesheet screen (`timesheetNotesRepo`).

### Budget vs Actual (Planned vs Timesheet)

- **CleanTrack Site Budgets** holds planned hours per weekday per site.
- **Variance** = actual hours (from timesheets) − planned hours; can be shown per site on dashboard or Sites & Budgets (to be fully wired in UI).

---

## Key Repos & Flows

- **usersRepo** – getCleanTrackUserByEmail (CleanTrack Users).
- **sitesRepo** – getSites (optional assignedSiteIds), createSite, updateSite, setSiteActive, toAppSite.
- **siteManagersRepo** – getAssignedSiteIdsForManager (CleanTrack Site Managers).
- **cleanersRepo** – getCleaners, createCleaner.
- **metricsRepo** – getDashboardMetrics (optional assignedSiteIds for Manager).
- **sharepoint** – getSiteId, getLists, getListIdByName, getListItems, getListColumns, createListItem, updateListItem.

---

## Not in Scope (Yet)

- Payroll system; accounting system; full rostering.
- AI Insights (anomaly detection, recommendations).
- Approve/lock timesheet entries; compliance docs for cleaners.
- CleanTrack Users CRUD in-app (currently read-only for auth).

This doc reflects the intended design and what is implemented so the whole platform can be built out and wired end-to-end.
