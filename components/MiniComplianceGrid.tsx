import React from 'react';
import { format, addDays, getDay } from 'date-fns';
import { getDayStatus } from '../utils';

interface MiniComplianceGridProps {
  startDate: Date;
  dailyBudgets: number[];
  actualsByDate: Record<string, number>;
}

const MiniComplianceGrid: React.FC<MiniComplianceGridProps> = ({ startDate, dailyBudgets, actualsByDate }) => {
  const dates = Array.from({ length: 14 }, (_, i) => addDays(startDate, i));

  return (
    <div className="grid grid-cols-7 grid-rows-2 gap-[2px] w-fit">
      {dates.map((date) => {
        const dateStr = format(date, 'yyyy-MM-dd');
        const dayOfWeek = getDay(date);
        const budget = dailyBudgets[dayOfWeek] || 0;
        const actual = actualsByDate[dateStr] || 0;
        const { dot, label } = getDayStatus(budget, actual);
        const displayDate = format(date, 'MMM d');
        
        return (
          <div
            key={date.toISOString()}
            title={`${displayDate}: ${label}`}
            className={`w-2.5 h-2.5 rounded-[1px] ${dot.replace('text-', 'bg-')} transition-opacity hover:opacity-75 cursor-help`}
          />
        );
      })}
    </div>
  );
};

export default MiniComplianceGrid;