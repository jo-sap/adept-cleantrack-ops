# CleanTrack Site Budgets – SharePoint list setup

Use this when creating or updating the **CleanTrack Site Budgets** list so it stays tidy and legible, and matches what the app expects.

---

## 1. Recommended column order (for a tidy list)

Order columns in the list like this. Grouping makes it easy to scan and to build views.

| Order | Column display name        | Type        | Used when        |
|-------|----------------------------|------------|------------------|
| 1     | Budget Name                | Single line| Always            |
| 2     | Site                       | Lookup     | Always            |
| 3     | Active                     | Yes/No     | Always            |
| 4     | **Visit Frequency**        | **Choice** | Always            |
| 5     | **Hours per Visit**         | **Number** | Monthly only      |
| 6     | Monday Hours               | Number     | Weekly / Fortnightly (Week 1) |
| 7     | Tuesday Hours              | Number     | "                 |
| 8     | Wednesday Hours            | Number     | "                 |
| 9     | Thursday Hours             | Number     | "                 |
| 10    | Friday Hours               | Number     | "                 |
| 11    | Saturday Hours             | Number     | "                 |
| 12    | Sunday Hours               | Number     | "                 |
| 13    | Week 2 Monday Hours        | Number     | Fortnightly only  |
| 14    | Week 2 Tuesday Hours       | Number     | "                 |
| 15    | Week 2 Wednesday Hours     | Number     | "                 |
| 16    | Week 2 Thursday Hours      | Number     | "                 |
| 17    | Week 2 Friday Hours        | Number     | "                 |
| 18    | Week 2 Saturday Hours      | Number     | "                 |
| 19    | Week 2 Sunday Hours        | Number     | "                 |

The app matches columns by **display name**. Use the names in the table exactly (including spaces and “Week 2”) so the app can read and write them.

---

## 2. Visit Frequency choice values

Create the **Visit Frequency** column as a **Choice** column with exactly these values (no extra spaces):

- `Weekly`
- `Fortnightly`
- `Monthly`

Default can be **Weekly**.

---

## 3. Number columns

- **Hours per Visit**, and all **Hours** columns: use **Number** with 2 decimal places (or more if you need it).
- Allow zero; no need to require a value (the app will write 0 when not used).

---

## 4. List views (for legibility)

Create separate views so each view only shows what’s relevant:

- **All columns**  
  Use when you need to see or edit everything (e.g. in list settings or for support).

- **Weekly budgets**  
  Filter: `Visit Frequency = Weekly`.  
  Show: Budget Name, Site, Active, Visit Frequency, Monday Hours … Sunday Hours.  
  Hide: Hours per Visit, all Week 2 columns.

- **Fortnightly budgets**  
  Filter: `Visit Frequency = Fortnightly`.  
  Show: Budget Name, Site, Active, Visit Frequency, Week 1 (Mon–Sun), Week 2 (Mon–Sun).  
  Hide: Hours per Visit.

- **Monthly budgets**  
  Filter: `Visit Frequency = Monthly`.  
  Show: Budget Name, Site, Active, Visit Frequency, Hours per Visit.  
  Hide: All day columns (Monday Hours … Sunday Hours and all Week 2 columns).

That way the list stays tidy and legible no matter which frequency you’re looking at.

---

## 5. Optional: “Week 1” prefix

If you want the list to read even more clearly, you can rename the first set of days to “Week 1 …”:

- Week 1 Monday Hours  
- Week 1 Tuesday Hours  
- …  
- Week 1 Sunday Hours  

The app currently expects **Monday Hours** … **Sunday Hours** (no “Week 1”). If you rename them, the app would need a small update to also look for “Week 1 Monday Hours” etc. as fallbacks. If you keep the existing names, no change is needed.

---

## 6. Summary

- Use the **exact display names** in the table so the app can read/write the list.
- Put **Visit Frequency** and **Hours per Visit** near the top, then Week 1 days, then Week 2 days.
- Use **list views** to show only the columns that matter for Weekly / Fortnightly / Monthly.
- Use the **Visit Frequency** choice values **Weekly**, **Fortnightly**, **Monthly** for filtering and consistency.

---

## 7. Behaviour when changing Visit Frequency (Edit Site)

When you change the **Visit frequency** dropdown and save, the app does the following so the list stays consistent:

| From → To    | What the app does |
|-------------|-------------------|
| **→ Weekly**   | Saves Week 1 hours only. Saves **0** for all Week 2 columns and **0** for Hours per Visit so old values are cleared. |
| **→ Fortnightly** | Saves Week 1 and Week 2 hours. If Week 2 was empty when you switched, the UI copies Week 1 into Week 2 so you can adjust. Saves **0** for Hours per Visit. |
| **→ Monthly**   | Saves **Visit Frequency** = Monthly and **Hours per Visit** (e.g. 25). Saves **0** for all Week 1 and Week 2 day columns so the list doesn’t keep old weekly/fortnightly hours. Fortnight cap = Hours per Visit ÷ 2. |

So after save, the list only has non-zero values in the columns that apply to the selected frequency.

---

## 8. Assigning a cleaner to a monthly (or any) job

Cleaners are **assigned to a site by recording time** for that site:

1. Go to **Timesheets**.
2. Click the **site** (e.g. Steel Builders Pty Ltd).
3. If the site has no cleaners yet, the **personnel list shows all cleaners** – choose the one doing the monthly job (e.g. Joseph Sapio).
4. Enter **actual hours** on the day(s) they worked (e.g. 25 on one day for a monthly visit).
5. Click **Save Batch**.

After the first save, that cleaner is treated as assigned to the site and will appear in the site’s personnel list. For **monthly** sites, the timesheet uses the budget’s **Hours per Visit** (e.g. 25) to derive the fortnight plan (25 ÷ 2 = 12.5h per period); the **Budget** total reflects that, and you can enter actuals on any day(s).
