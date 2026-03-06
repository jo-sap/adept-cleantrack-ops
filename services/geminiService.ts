
import { GoogleGenAI, GenerateContentResponse } from "@google/genai";
import { Site, TimeEntry, Cleaner } from "../types";
import { getDay } from "date-fns";

export const getOperationalInsights = async (
  sites: Site[], 
  cleaners: Cleaner[], 
  entries: TimeEntry[], 
  periodLabel: string
): Promise<string> => {
  // Initialize the Gemini API client using the required pattern
  const ai = new GoogleGenAI({ apiKey: process.env.API_KEY });
  
  const dataContext = {
    period: periodLabel,
    sites: sites.map(s => {
      // siteId is now part of the augmented TimeEntry interface
      const siteEntries = entries.filter(e => e.siteId === s.id);
      const actual = siteEntries.reduce((acc, curr) => acc + curr.hours, 0);
      
      return {
        name: s.name,
        // Use correct snake_case properties from the Site interface
        fortnightBudget: s.budgeted_hours_per_fortnight,
        dailyBudgets: s.daily_budgets, // Array of 7 numbers (Sun-Sat)
        totalActual: actual,
        logs: siteEntries.map(e => {
          const dayIdx = getDay(new Date(e.date));
          const target = s.daily_budgets[dayIdx];
          return { 
            date: e.date, 
            dayOfWeek: formatDay(dayIdx),
            targetForThisDay: target,
            isScheduled: target > 0,
            hours: e.hours 
          };
        })
      };
    })
  };

  const prompt = `
    As a Clean Operations Auditor, analyze the following timesheet data for ${periodLabel}.
    
    The manager has configured granular budgets for each day of the week (e.g., some days might be 3h, others 2h, others 0h).
    
    Audit Focus:
    1. "Target Drift": Identify sites where logs consistently exceed the specific daily target (e.g., logging 4h when 3h was planned).
    2. "Schedule Compliance": Flag work done on unbudgeted days (leakage).
    3. "Operational Efficiency": Compare the total fortnightly actuals against the calculated budget.
    4. "Clean Team Productivity": Are specific sites consistently running over budget due to workload issues?

    Context Data:
    ${JSON.stringify(dataContext, null, 2)}

    Format: Return the analysis in professional Markdown with clear headings and bold highlights.
  `;

  try {
    // Audit analysis requires advanced reasoning, hence gemini-3-pro-preview is preferred
    const response: GenerateContentResponse = await ai.models.generateContent({
      model: 'gemini-3-pro-preview',
      contents: prompt,
    });
    
    // Direct access to text property per API guidelines
    return response.text || "Operational analysis unavailable.";
  } catch (error) {
    console.error("AI Insights Error:", error);
    return "Operational Auditor is currently offline. Please check your data and try again.";
  }
};

const formatDay = (idx: number) => ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'][idx];