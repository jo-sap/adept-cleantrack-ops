import React from "react";
import { format } from "date-fns";
import { Loader2, StickyNote } from "lucide-react";
import type { FortnightPeriod } from "../types";

interface TimesheetPeriodNotesPanelProps {
  currentPeriod: FortnightPeriod;
  canEdit: boolean;
  noteBody: string;
  onNoteChange: (value: string) => void;
  notesLoading: boolean;
  listMissing: boolean;
  listSchemaError: string | null;
}

const TimesheetPeriodNotesPanel: React.FC<TimesheetPeriodNotesPanelProps> = ({
  currentPeriod,
  canEdit,
  noteBody,
  onNoteChange,
  notesLoading,
  listMissing,
  listSchemaError,
}) => {
  const periodLabel = `${format(currentPeriod.startDate, "MMM d")} — ${format(currentPeriod.endDate, "MMM d")}`;

  if (!canEdit) return null;

  return (
    <div className="border border-[#edeef0] rounded-xl bg-[#fafafb] px-4 py-3 space-y-3">
      <div className="flex items-start gap-2">
        <StickyNote className="text-gray-500 shrink-0 mt-0.5" size={16} />
        <div className="min-w-0 flex-1">
          <h3 className="text-xs font-bold text-gray-900 uppercase tracking-wide">
            Manager notes
          </h3>
          <p className="text-[11px] text-gray-500 mt-0.5">
            For this site and fortnight ({periodLabel}). Saved together with <strong>Save</strong> above.
          </p>
        </div>
        {notesLoading && <Loader2 className="animate-spin text-gray-400 shrink-0" size={16} />}
      </div>

      {listMissing && !notesLoading && (
        <p className="text-[11px] text-amber-800 bg-amber-50 border border-amber-100 rounded-lg px-2 py-1.5">
          Create the list <strong>CleanTrack Timesheet Period Notes</strong> on the CleanTrack site with{" "}
          <strong>Site</strong> (lookup to CleanTrack Sites), <strong>Period Start</strong> (date), and{" "}
          <strong>Note</strong> (multi-line).
        </p>
      )}

      {listSchemaError && !notesLoading && (
        <p className="text-[11px] text-red-900 bg-red-50 border border-red-100 rounded-lg px-2 py-1.5">
          {listSchemaError}
        </p>
      )}

      <div className="bg-white border border-gray-200 rounded-lg p-3">
        <textarea
          value={noteBody}
          onChange={(e) => onNoteChange(e.target.value)}
          rows={4}
          placeholder="Optional context while completing this timesheet (e.g. hours short, site ending, cleaner change)."
          className="w-full so-input px-3 py-2 text-sm resize-y min-h-[96px] placeholder:text-gray-400"
          disabled={!!listSchemaError}
        />
      </div>
    </div>
  );
};

export default TimesheetPeriodNotesPanel;
