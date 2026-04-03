import React from "react";
import type { Cleaner, Site } from "../types";
import { normalizeListItemId } from "../lib/sharepoint";
import { addDays, addMonths, endOfMonth, format, startOfMonth } from "date-fns";
import { getPublicHolidaysInRange } from "../lib/publicHolidays";

interface ContractorFinanceProps {
  sites: Site[];
  cleaners: Cleaner[];
}

const MONTHS_PER_YEAR = 12;
const FORTNIGHTS_PER_YEAR = 26;

const toCurrency = (n: number): string =>
  n.toLocaleString("en-AU", {
    style: "currency",
    currency: "AUD",
    maximumFractionDigits: 2,
  });

function toDateKey(d: Date): string {
  return format(d, "yyyy-MM-dd");
}

function getPlannedHoursForDate(site: Site, day: Date): number {
  const dayOfWeek = day.getDay(); // 0=Sun..6=Sat
  const week1 = site.daily_budgets ?? [0, 0, 0, 0, 0, 0, 0];
  const week2 = site.daily_budgets_week2;
  const visit = String(site.visit_frequency ?? "").trim().toLowerCase();

  // For fortnightly sites with week2 budgets, use average day profile over a 2-week cycle
  // so monthly estimates remain stable without requiring a cycle anchor date.
  if (visit === "fortnightly" && Array.isArray(week2) && week2.length >= 7) {
    return ((week1[dayOfWeek] ?? 0) + (week2[dayOfWeek] ?? 0)) / 2;
  }
  return week1[dayOfWeek] ?? 0;
}

const ContractorFinance: React.FC<ContractorFinanceProps> = ({ sites, cleaners }) => {
  const [selectedMonthStart, setSelectedMonthStart] = React.useState<Date>(() => startOfMonth(new Date()));

  const contractors = React.useMemo(
    () =>
      new Set(
        cleaners
          .filter((c) => (c.type ?? "cleaner") === "contractor")
          .map((c) => normalizeListItemId(String(c.id)))
      ),
    [cleaners]
  );

  const rows = React.useMemo(() => {
    const monthStart = startOfMonth(selectedMonthStart);
    const monthEnd = endOfMonth(selectedMonthStart);

    return sites
      .filter((site) =>
        (site.assigned_cleaner_ids ?? []).some((id) =>
          contractors.has(normalizeListItemId(String(id)))
        )
      )
      .map((site) => {
        const phSet = getPublicHolidaysInRange(monthStart, monthEnd, site.state);
        const weekdayRate = Number(site.budget_weekday_labour_rate ?? site.budget_labour_rate ?? 0);
        const saturdayRate = Number(site.budget_saturday_labour_rate ?? weekdayRate);
        const sundayRate = Number(site.budget_sunday_labour_rate ?? weekdayRate);

        let monthlyPlannedHours = 0;
        let monthlyPlannedHoursExcludingPH = 0;
        let monthlyContractorBudget = 0;
        let d = new Date(monthStart);
        while (d <= monthEnd) {
          const dayKey = toDateKey(d);
          const plannedHours = getPlannedHoursForDate(site, d);
          monthlyPlannedHours += plannedHours;
          if (!phSet.has(dayKey)) {
            monthlyPlannedHoursExcludingPH += plannedHours;
            const day = d.getDay(); // 0=Sun..6=Sat
            const applicableRate = day === 0 ? sundayRate : day === 6 ? saturdayRate : weekdayRate;
            monthlyContractorBudget += plannedHours * applicableRate;
          }
          d = addDays(d, 1);
        }
        const monthlyRevenue = Number(site.monthly_revenue ?? 0);
        const variance = monthlyRevenue - monthlyContractorBudget;
        return {
          id: site.id,
          siteName: site.name,
          weekdayRate,
          saturdayRate,
          sundayRate,
          monthlyPlannedHours,
          monthlyPlannedHoursExcludingPH,
          monthlyRevenue,
          monthlyContractorBudget,
          variance,
        };
      })
      .sort((a, b) => a.siteName.localeCompare(b.siteName));
  }, [sites, contractors, selectedMonthStart]);

  const totals = React.useMemo(() => {
    return rows.reduce(
      (acc, row) => {
        acc.revenue += row.monthlyRevenue;
        acc.budget += row.monthlyContractorBudget;
        acc.variance += row.variance;
        return acc;
      },
      { revenue: 0, budget: 0, variance: 0 }
    );
  }, [rows]);

  return (
    <div className="space-y-5 animate-fadeIn">
      <div className="bg-white border border-[#edeef0] rounded-lg shadow-sm p-4">
        <h2 className="text-lg font-semibold text-gray-900">Contractor Sites Financial View</h2>
        <div className="mt-3 flex flex-wrap items-center gap-2">
          <button
            type="button"
            onClick={() => setSelectedMonthStart((d) => startOfMonth(addMonths(d, -1)))}
            className="px-2.5 py-1.5 text-xs font-semibold rounded border border-[#d9dce1] text-gray-700 hover:bg-gray-50"
          >
            Previous month
          </button>
          <div className="px-3 py-1.5 rounded bg-[#f8f9fb] border border-[#edeef0] text-sm font-medium text-gray-800">
            {format(selectedMonthStart, "MMMM yyyy")}
          </div>
          <button
            type="button"
            onClick={() => setSelectedMonthStart((d) => startOfMonth(addMonths(d, 1)))}
            className="px-2.5 py-1.5 text-xs font-semibold rounded border border-[#d9dce1] text-gray-700 hover:bg-gray-50"
          >
            Next month
          </button>
        </div>
        <p className="text-sm text-gray-600 mt-3">
          Monthly contractor budget uses <span className="font-medium">Rate x Planned Hours</span> for the selected
          month, with <span className="font-medium">public holidays excluded</span> by each site's state calendar.
        </p>
      </div>

      <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
        <div className="bg-white border border-[#edeef0] rounded-lg shadow-sm p-4">
          <p className="text-xs uppercase tracking-wide text-gray-500">Total Monthly Revenue</p>
          <p className="text-xl font-semibold text-gray-900">{toCurrency(totals.revenue)}</p>
        </div>
        <div className="bg-white border border-[#edeef0] rounded-lg shadow-sm p-4">
          <p className="text-xs uppercase tracking-wide text-gray-500">Total Monthly Contractor Budget</p>
          <p className="text-xl font-semibold text-gray-900">{toCurrency(totals.budget)}</p>
        </div>
        <div className="bg-white border border-[#edeef0] rounded-lg shadow-sm p-4">
          <p className="text-xs uppercase tracking-wide text-gray-500">Total Variance</p>
          <p className={`text-xl font-semibold ${totals.variance < 0 ? "text-rose-600" : "text-emerald-700"}`}>
            {toCurrency(totals.variance)}
          </p>
        </div>
      </div>

      <div className="border border-[#edeef0] rounded-lg bg-white shadow-sm overflow-hidden">
        <table className="w-full border-collapse text-left">
          <thead className="bg-[#fcfcfb] border-b border-[#edeef0]">
            <tr className="text-[11px] font-semibold text-gray-600 uppercase tracking-wide">
              <th className="py-3 px-4">Site</th>
              <th className="py-3 px-4">Rates (Wk/Sat/Sun)</th>
              <th className="py-3 px-4">Planned Hrs</th>
              <th className="py-3 px-4">Billable Hrs (Ex PH)</th>
              <th className="py-3 px-4">Monthly Revenue</th>
              <th className="py-3 px-4">Monthly Contractor Budget</th>
              <th className="py-3 px-4">Variance</th>
            </tr>
          </thead>
          <tbody>
            {rows.length === 0 ? (
              <tr>
                <td colSpan={7} className="py-8 px-4 text-sm text-gray-500">
                  No active contractor-assigned sites found.
                </td>
              </tr>
            ) : (
              rows.map((row) => (
                <tr key={row.id} className="border-b border-[#edeef0] last:border-b-0">
                  <td className="py-3 px-4 text-sm text-gray-900">{row.siteName}</td>
                  <td className="py-3 px-4 text-sm text-gray-700">
                    {toCurrency(row.weekdayRate)} / {toCurrency(row.saturdayRate)} / {toCurrency(row.sundayRate)}
                  </td>
                  <td className="py-3 px-4 text-sm text-gray-700">{row.monthlyPlannedHours.toFixed(1)}h</td>
                  <td className="py-3 px-4 text-sm text-gray-700">{row.monthlyPlannedHoursExcludingPH.toFixed(1)}h</td>
                  <td className="py-3 px-4 text-sm text-gray-700">{toCurrency(row.monthlyRevenue)}</td>
                  <td className="py-3 px-4 text-sm text-gray-700">{toCurrency(row.monthlyContractorBudget)}</td>
                  <td className={`py-3 px-4 text-sm font-medium ${row.variance < 0 ? "text-rose-600" : "text-emerald-700"}`}>
                    {toCurrency(row.variance)}
                  </td>
                </tr>
              ))
            )}
          </tbody>
        </table>
      </div>
    </div>
  );
};

export default ContractorFinance;
