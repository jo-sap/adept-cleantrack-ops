
import React from 'react';
import { LayoutDashboard, Building2, Users, Clock, BrainCircuit, ShieldCheck, ChevronRight, KeyRound } from 'lucide-react';
import { ViewType } from '../types';
import { useAppAuth } from '../contexts/AppAuthContext';

interface SidebarProps {
  currentView: ViewType;
  onViewChange: (view: ViewType) => void;
}

const Sidebar: React.FC<SidebarProps> = ({ currentView, onViewChange }) => {
  const { user } = useAppAuth();
  const displayRole = user?.role ?? "—";
  
  const menuItems = [
    { id: 'dashboard', label: 'Dashboard', icon: LayoutDashboard },
    { id: 'sites', label: 'Sites & Budgets', icon: Building2 },
    ...(user?.role === 'Admin'
      ? [{ id: 'team', label: 'Team', icon: Users }]
      : []),
    { id: 'cleaners', label: 'Cleaner Team', icon: Users },
    { id: 'timesheets', label: 'Timesheets', icon: Clock },
    { id: 'insights', label: 'AI Insights', icon: BrainCircuit },
    { id: 'auth-test', label: 'Auth Test', icon: KeyRound },
  ];

  return (
    <aside className="w-60 bg-[#f7f6f3] border-r border-[#edeef0] h-screen sticky top-0 flex flex-col select-none">
      <div className="p-4 mb-2 flex flex-col gap-4">
        <div className="flex items-center gap-2 hover:bg-black/5 cursor-pointer transition-colors duration-150 rounded-md p-2">
          <div className="bg-gray-800 p-1 rounded shadow-sm">
            <ShieldCheck className="text-white" size={16} />
          </div>
          <h1 className="font-semibold text-[14px] text-gray-700 truncate">CleanTrack Ops</h1>
        </div>

        <div className="px-2">
          <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest mb-1 block">Role</label>
          <p className="text-[12px] font-medium text-gray-700">Role: {displayRole}</p>
        </div>
      </div>
      
      <nav className="flex-1 px-2 space-y-[2px]">
        {menuItems.map((item) => {
          const Icon = item.icon;
          const isActive = currentView === item.id;
          return (
            <button
              key={item.id}
              onClick={() => onViewChange(item.id as ViewType)}
              className={`w-full flex items-center gap-2 px-3 py-1.5 rounded-md transition-colors duration-150 group text-[14px] ${
                isActive 
                  ? 'bg-white shadow-sm text-gray-900 border border-[#edeef0]' 
                  : 'text-gray-500 hover:bg-black/5 hover:text-gray-900'
              }`}
            >
              <Icon size={16} className={isActive ? 'text-gray-800' : 'text-gray-400 group-hover:text-gray-600'} />
              <span className={isActive ? 'font-medium' : 'font-normal'}>{item.label}</span>
            </button>
          );
        })}
      </nav>

      <div className="p-4 mt-auto">
        <div className="flex items-center gap-2 px-2 py-2 hover:bg-black/5 rounded-md cursor-pointer transition-colors">
          <div className="w-6 h-6 rounded-full bg-orange-100 flex items-center justify-center text-orange-700 font-bold text-[10px]">
            {displayRole === 'Admin' ? 'AD' : displayRole === 'Manager' ? 'OM' : '—'}
          </div>
          <div className="flex-1 min-w-0">
            <p className="text-[12px] font-medium text-gray-700 truncate">{displayRole === 'Admin' ? 'System Admin' : displayRole === 'Manager' ? 'Ops Manager' : displayRole}</p>
          </div>
          <ChevronRight size={14} className="text-gray-400" />
        </div>
      </div>
    </aside>
  );
};

export default Sidebar;
