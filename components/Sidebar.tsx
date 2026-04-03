
import React from 'react';
import { LayoutDashboard, Building2, Users, Clock, ChevronRight, X, Briefcase, ReceiptText } from 'lucide-react';
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
    ...(user?.role === 'Accounts'
      ? [{ id: 'contractor-finance', label: 'Contractor Finance', icon: ReceiptText }]
      : [
          { id: 'dashboard', label: 'Dashboard', icon: LayoutDashboard },
          { id: 'sites', label: 'Sites & Budgets', icon: Building2 },
          ...(user?.role === 'Admin'
            ? [{ id: 'team', label: 'Team', icon: Users }]
            : []),
          { id: 'cleaners', label: 'Workforce', icon: Users },
          { id: 'timesheets', label: 'Timesheets', icon: Clock },
          { id: 'adhoc-jobs', label: 'Ad Hoc Jobs', icon: Briefcase },
          ...(user?.role === 'Admin'
            ? [{ id: 'contractor-finance', label: 'Contractor Finance', icon: ReceiptText }]
            : []),
        ]),
  ];

  const panel = (
    <>
      <div className="px-5 py-5 mb-2 flex flex-col gap-4 so-sidebar-header border-b border-white/5">
        <div className="flex items-center gap-3 rounded-lg px-1.5 py-1.5 cursor-default">
          <div className="h-8 w-8 rounded-md bg-white/90 p-1.5 shadow-sm">
            <img
              src="/images/a_adept.png"
              alt="Adept Timesheet Ops logo"
              className="h-full w-full object-contain"
            />
          </div>
          <div className="min-w-0">
            <h1 className="font-semibold text-[14px] text-slate-50 truncate">Adept Timesheet Ops</h1>
            <p className="text-[11px] text-slate-300 truncate">Operations workspace</p>
          </div>
        </div>

        <div className="px-2">
          <label className="text-[10px] font-semibold text-slate-400 uppercase tracking-[0.18em] mb-1 block">
            Role
          </label>
          <p className="text-[12px] font-medium text-slate-100">Role: {displayRole}</p>
        </div>
      </div>

      <nav className="flex-1 px-3 space-y-1.5 overflow-y-auto">
        {menuItems.map((item) => {
          const Icon = item.icon;
          const isActive = currentView === item.id;
          return (
            <button
              key={item.id}
              onClick={() => handleNav(item.id as ViewType)}
              className={`w-full flex items-center gap-3 px-3.5 py-3 rounded-md so-sidebar-nav-item transition-colors duration-150 group text-[13px] min-h-[42px] ${
                isActive
                  ? 'so-sidebar-nav-item--active text-slate-50'
                  : 'text-slate-300 hover:bg-slate-700/70 hover:text-white'
              }`}
            >
              <Icon
                size={16}
                className={
                  isActive
                    ? 'text-teal-200'
                    : 'text-slate-400 group-hover:text-slate-200'
                }
              />
              <span className={isActive ? 'font-medium' : 'font-normal'}>{item.label}</span>
            </button>
          );
        })}
      </nav>

      <div className="px-5 py-4 mt-auto border-t border-white/5">
        <div className="flex items-center gap-3 px-2 py-2 rounded-lg transition-colors min-h-[44px] hover:bg-slate-700/70 cursor-default">
          <div className="w-7 h-7 rounded-full bg-emerald-500/10 flex items-center justify-center text-emerald-300 font-semibold text-[10px] border border-emerald-400/30">
            {displayRole === 'Admin' ? 'AD' : displayRole === 'Manager' ? 'OM' : displayRole === 'Accounts' ? 'AC' : '—'}
          </div>
          <div className="flex-1 min-w-0">
            <p className="text-[12px] font-medium text-slate-50 truncate">
              {displayRole === 'Admin'
                ? 'System Admin'
                : displayRole === 'Manager'
                ? 'Ops Manager'
                : displayRole === 'Accounts'
                ? 'Financial Controller'
                : displayRole}
            </p>
            <p className="text-[11px] text-slate-400 truncate">
              Microsoft 365 · Internal
            </p>
          </div>
          <ChevronRight size={14} className="text-slate-500" />
        </div>
      </div>
    </>
  );

  return (
    <>
      {/* Desktop: fixed sidebar so it stays visible when scrolling */}
      <aside className="hidden lg:flex fixed left-0 top-0 z-30 w-60 h-screen so-sidebar border-r border-slate-800 flex-col select-none">
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
          <div className="absolute left-0 top-0 bottom-0 w-72 max-w-[85vw] so-sidebar border-r border-slate-800 flex flex-col shadow-xl">
            <div className="flex items-center justify-between p-4 border-b border-slate-800">
              <span className="font-semibold text-slate-100">Navigation</span>
              <button
                type="button"
                onClick={onMobileClose}
                className="p-2 -m-2 rounded-md text-slate-400 hover:bg-slate-700/70 hover:text-white touch-manipulation"
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
