
import React from 'react';
import { LayoutDashboard, Building2, Users, Clock, BrainCircuit, ShieldCheck, ChevronRight, KeyRound, X, Briefcase } from 'lucide-react';
import { ViewType } from '../types';
import { useAppAuth } from '../contexts/AppAuthContext';

interface SidebarProps {
  currentView: ViewType;
  onViewChange: (view: ViewType) => void;
  /** On mobile: drawer open state. When true, sidebar is shown as overlay. */
  mobileOpen?: boolean;
  /** Called when user requests to close the mobile drawer (e.g. backdrop click or after nav). */
  onMobileClose?: () => void;
}

const Sidebar: React.FC<SidebarProps> = ({ currentView, onViewChange, mobileOpen, onMobileClose }) => {
  const { user } = useAppAuth();
  const displayRole = user?.role ?? "—";

  const handleNav = (view: ViewType) => {
    onViewChange(view);
    onMobileClose?.();
  };

  const menuItems = [
    { id: 'dashboard', label: 'Dashboard', icon: LayoutDashboard },
    { id: 'sites', label: 'Sites & Budgets', icon: Building2 },
    ...(user?.role === 'Admin'
      ? [{ id: 'team', label: 'Team', icon: Users }]
      : []),
    { id: 'cleaners', label: 'Cleaner Team', icon: Users },
    { id: 'timesheets', label: 'Timesheets', icon: Clock },
    { id: 'adhoc-jobs', label: 'Ad Hoc Jobs', icon: Briefcase },
    { id: 'insights', label: 'AI Insights', icon: BrainCircuit },
    { id: 'auth-test', label: 'Auth Test', icon: KeyRound },
  ];

  const panel = (
    <>
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

      <nav className="flex-1 px-2 space-y-[2px] overflow-y-auto">
        {menuItems.map((item) => {
          const Icon = item.icon;
          const isActive = currentView === item.id;
          return (
            <button
              key={item.id}
              onClick={() => handleNav(item.id as ViewType)}
              className={`w-full flex items-center gap-2 px-3 py-3 rounded-md transition-colors duration-150 group text-[14px] min-h-[44px] ${
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
        <div className="flex items-center gap-2 px-2 py-2 hover:bg-black/5 rounded-md cursor-pointer transition-colors min-h-[44px]">
          <div className="w-6 h-6 rounded-full bg-orange-100 flex items-center justify-center text-orange-700 font-bold text-[10px]">
            {displayRole === 'Admin' ? 'AD' : displayRole === 'Manager' ? 'OM' : '—'}
          </div>
          <div className="flex-1 min-w-0">
            <p className="text-[12px] font-medium text-gray-700 truncate">{displayRole === 'Admin' ? 'System Admin' : displayRole === 'Manager' ? 'Ops Manager' : displayRole}</p>
          </div>
          <ChevronRight size={14} className="text-gray-400" />
        </div>
      </div>
    </>
  );

  return (
    <>
      {/* Desktop: fixed sidebar so it stays visible when scrolling */}
      <aside className="hidden lg:flex fixed left-0 top-0 z-30 w-60 h-screen bg-[#f7f6f3] border-r border-[#edeef0] flex-col select-none">
        {panel}
      </aside>

      {/* Mobile: overlay drawer */}
      {mobileOpen && (
        <div className="fixed inset-0 z-50 lg:hidden" aria-modal="true" role="dialog">
          <div
            className="absolute inset-0 bg-black/40"
            onClick={onMobileClose}
            aria-hidden="true"
          />
          <div className="absolute left-0 top-0 bottom-0 w-72 max-w-[85vw] bg-[#f7f6f3] border-r border-[#edeef0] flex flex-col shadow-xl">
            <div className="flex items-center justify-between p-4 border-b border-[#edeef0]">
              <span className="font-semibold text-gray-800">Menu</span>
              <button
                type="button"
                onClick={onMobileClose}
                className="p-2 -m-2 rounded-md text-gray-500 hover:bg-black/10 hover:text-gray-700 touch-manipulation"
                aria-label="Close menu"
              >
                <X size={20} />
              </button>
            </div>
            <div className="flex-1 flex flex-col min-h-0 overflow-hidden">
              {panel}
            </div>
          </div>
        </div>
      )}
    </>
  );
};

export default Sidebar;
