# Accurate budget: Public holidays, Saturday & Sunday rates

The app uses a **date-aware** budgeted labour cost so the “budget” figure matches how you actually pay: different rates for weekdays, Saturday, Sunday, and public holidays.

---

## How it works

1. **Rates per site (CleanTrack Site Budgets)**  
   Each site has four optional labour rates ($/hr):
   - **Weekday Labour Rate** — Mon–Fri
   - **Saturday Labour Rate** — Saturday
   - **Sunday Labour Rate** — Sunday  
   - **PH Labour Rate** — Public holidays (overrides the day-of-week rate when the date is a PH)

2. **Weekly hour pattern**  
   The budget’s **Daily Service Hours** (Sun, Mon, …, Sat) define how many hours are planned for each day of the week. That pattern is repeated for every week.

3. **Public holiday calendar**  
   A list of public holiday **dates** is used to decide when to apply the **PH rate** instead of the weekday/Sat/Sun rate.  
   Default: **NSW, Australia** for 2024–2026 (see `lib/publicHolidays.ts`).

4. **Date-aware calculation (Dashboard)**  
   For the selected **fortnight**:
   - For **each date** in the period we:
     - Take the **hours** from the site’s pattern for that day of week (e.g. Monday hours).
     - If the date is a **public holiday** → use **PH Labour Rate**.
     - Else if **Sunday** → use **Sunday Labour Rate**.
     - Else if **Saturday** → use **Saturday Labour Rate**.
     - Else → use **Weekday Labour Rate**.
   - **Budgeted labour cost** = sum over all days of `(hours × rate)`.

So the budget reflects exactly which days fall in the period and whether they are PH / Sat / Sun / weekday.

---

## Files involved

| File | Role |
|------|------|
| `lib/publicHolidays.ts` | List of PH dates (default NSW), `isPublicHoliday(date)`, `getPublicHolidaysInRange(start, end)`. |
| `lib/budgetedLabourCost.ts` | `computeBudgetedLabourCostForRange({ startDate, endDate, dailyBudgets, weekdayRate, saturdayRate, sundayRate, phRate, publicHolidayDates? })` — returns total $ for the period. |
| `components/Dashboard.tsx` | Uses the above for each site’s “Budgeted labour cost” in the recap and totals. |

---

## Extending the system

### Add more years or regions

- Edit **`lib/publicHolidays.ts`** and add more ISO date strings (`YYYY-MM-DD`) to `PUBLIC_HOLIDAYS`, or call `addPublicHolidays([...])` at app init if you load dates from elsewhere.

### Use a SharePoint “Public Holidays” list

1. Create a list with at least a **Date** column (and optionally Name, Region).
2. Add a repo (e.g. `repositories/publicHolidaysRepo.ts`) that fetches items and returns `Set<string>` of `YYYY-MM-DD` for a given range.
3. In the Dashboard (or a context/store), fetch PH dates for `currentPeriod` and pass them into `computeBudgetedLabourCostForRange` as `publicHolidayDates`.

### Use an external API

- Call a public-holiday API (e.g. by state/country) for the period, convert to `Set<string>` of date keys, and pass as `publicHolidayDates` into `computeBudgetedLabourCostForRange`.

### Show budgeted vs actual on Site Detail

- In `SiteDetail.tsx`, use `computeBudgetedLabourCostForRange` with `currentPeriod` and the site’s `daily_budgets` and four rates, then display “Budgeted labour” alongside “Actual labour cost” so users can compare per site.

---

## Fallbacks

- If a site has no **Saturday** or **Sunday** or **PH** rate set, the app falls back to the **Weekday** rate for that day type.
- If no PH dates are loaded, the calendar in `publicHolidays.ts` is used (NSW 2024–2026). If you never add dates and the period is outside that range, no days will be treated as PH.
