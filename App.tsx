
import React, { useState, useEffect, FC } from 'react';
import Sidebar from './components/Sidebar';
import Dashboard from './components/Dashboard';
import SiteManager from './components/SiteManager';
import AdHocJobsManager from './components/AdHocJobsManager';
import TimeEntryForm from './components/TimeEntryForm';
import SiteDetail from './components/SiteDetail';
import CleanerManager from './components/CleanerManager';
import TeamManager from './components/TeamManager';
import AuthTest from './components/AuthTest';
import SignInScreen from './components/SignInScreen';
import UnauthorizedScreen from './components/UnauthorizedScreen';
import { DevBypassBanner } from './components/DevBypassBanner';
import { Site, Cleaner, ViewType, FortnightPeriod, TimeBatch, TimeEntry } from './types';
import { getFortnightForDate } from './utils';
import { ChevronLeft, ChevronRight, Loader2, Menu } from 'lucide-react';
import { format } from 'date-fns';
import { RoleProvider, useRole } from './contexts/RoleContext';
import { AppAuthProvider, useAppAuth } from './contexts/AppAuthContext';
import { DEV_BYPASS_LOGIN } from './config/authFlags';
import { supabase } from './lib/supabase';
import { getGraphAccessToken } from './lib/graph';
import { getCleaners } from './repositories/cleanersRepo';
import { getSites, toAppSite } from './repositories/sitesRepo';
import { getSiteBudgets } from './repositories/budgetsRepo';
import { getTimesheetEntriesForRange, saveTimesheetEntriesToSharePoint } from './repositories/metricsRepo';
import { getAssignedSiteIdsForManager } from './repositories/siteManagersRepo';

const AppGate: FC = () => {
  const { authStatus, signOut } = useAppAuth();
  const canEnter = DEV_BYPASS_LOGIN || authStatus === 'authenticated';
  if (authStatus === 'loading' && !DEV_BYPASS_LOGIN) {
    return <div className="h-screen flex items-center justify-center bg-gray-50"><Loader2 className="animate-spin text-gray-400" /></div>;
  }
  if (authStatus === 'authorizing') {
    return <div className="h-screen flex items-center justify-center bg-gray-50"><Loader2 className="animate-spin text-gray-400" /></div>;
  }
  if (authStatus === 'unauthorized') {
    return <UnauthorizedScreen onSignOut={signOut} />;
  }
  if (!canEnter) return <SignInScreen />;
  return (
    <RoleProvider>
      <AppContent />
    </RoleProvider>
  );
};

const AppContent: FC = () => {
  const { role, loading, logout } = useRole();
  const { authStatus, signOut: signOutApp, user } = useAppAuth();
  const isAppAuthenticated = authStatus === 'authenticated';
  const [currentView, setCurrentView] = useState<ViewType>('dashboard');
  const [selectedSiteId, setSelectedSiteId] = useState<string | null>(null);
  const [sites, setSites] = useState<Site[]>([]);
  const [cleaners, setCleaners] = useState<Cleaner[]>([]);
  const [batches, setBatches] = useState<TimeBatch[]>([]);
  const [graphEntries, setGraphEntries] = useState<TimeEntry[]>([]);
  const [graphEntriesLoaded, setGraphEntriesLoaded] = useState(false);
  const [currentPeriod, setCurrentPeriod] = useState<FortnightPeriod>(getFortnightForDate(new Date()));
  const [dataLoading, setDataLoading] = useState(true);
  const [sidebarMobileOpen, setSidebarMobileOpen] = useState(false);

  const fetchSites = async () => {
    setDataLoading(true);
    const token = await getGraphAccessToken();
    if (token) {
      try {
        let list = await getSites(token);
        if (user?.role === 'Manager' && user?.email) {
          const isAllSites = user.permissionScope?.trim().toLowerCase() === 'allsites';
          if (!isAllSites) {
            const assignedIds = await getAssignedSiteIdsForManager(token, user.email);
            if (assignedIds.length > 0) list = list.filter((s) => assignedIds.includes(s.id));
          }
        }
        const budgets = await getSiteBudgets(token).catch(() => ({}));
        setSites(list.map((s) => {
          const budget = budgets[String(s.id)] ?? budgets["name:" + (s.siteName.trim() + " Budget")];
          return toAppSite(s, budget) as Site;
        }));
        setDataLoading(false);
        return;
      } catch {
        // fallback to Supabase
      }
    }
    const { data, error } = await supabase.from('sites').select(`
      *,
      managers:site_managers(active, profiles(*))
    `);
    if (!error) setSites(data as any[]);
    setDataLoading(false);
  };

  useEffect(() => {
    if (role || DEV_BYPASS_LOGIN || isAppAuthenticated) {
      fetchSites();
      fetchCleaners();
      fetchBatches();
    }
  }, [role, currentPeriod.id, isAppAuthenticated, user?.role, user?.email]);

  useEffect(() => {
    let cancelled = false;
    setGraphEntriesLoaded(false);
    getGraphAccessToken().then((token) => {
      if (!token || cancelled) {
        if (!cancelled) setGraphEntries([]);
        return;
      }
      const range = {
        start: currentPeriod.startDate,
        end: new Date(currentPeriod.endDate.getTime() + 24 * 60 * 60 * 1000),
      };
      getTimesheetEntriesForRange(token, range).then((entries) => {
        if (cancelled) return;
        setGraphEntries(entries.map((e) => ({
          id: e.id,
          batch_id: e.id,
          date: e.date,
          hours: e.hours,
          pay_rate_snapshot: e.pay_rate_snapshot,
          siteId: e.siteId,
          cleanerId: e.cleanerId,
          adhocJobId: e.adhocJobId,
          adhocJobName: e.adhocJobName,
        } as TimeEntry)));
        setGraphEntriesLoaded(true);
      }).catch(() => { if (!cancelled) setGraphEntriesLoaded(false); });
    });
    return () => { cancelled = true; };
  }, [currentPeriod.startDate.getTime(), currentPeriod.endDate.getTime()]);

  const fetchCleaners = async () => {
    const token = await getGraphAccessToken();
    if (token) {
      try {
        const items = await getCleaners(token);
        setCleaners(
          items.map((c) => {
            const parts = c.cleanerName.trim().split(/\s+/);
            const firstName = parts[0] ?? c.cleanerName;
            const lastName = parts.slice(1).join(" ") ?? "";
            return {
              id: c.id,
              firstName,
              lastName,
              email: "",
              phone: "",
              bankAccountName: c.accountName,
              bankBsb: c.bsb,
              bankAccountNumber: c.accountNumber,
              payRatePerHour: c.payRatePerHour,
            } as Cleaner;
          })
        );
        return;
      } catch {
        // fallback to Supabase if Graph fails
      }
    }
    const { data, error } = await supabase.from('cleaners').select('*');
    if (!error) setCleaners(data.map(c => ({
      ...c,
      firstName: c.first_name,
      lastName: c.last_name,
      bankAccountName: c.bank_account_name,
      bankBsb: c.bank_bsb,
      bankAccountNumber: c.bank_account_number,
      payRatePerHour: c.pay_rate_per_hour
    })));
  };

  const fetchBatches = async () => {
    const { data, error } = await supabase.from('timesheet_batches').select(`
      *,
      editor:profiles(full_name),
      entries:timesheet_entries(*)
    `)
    .eq('fortnight_start', format(currentPeriod.startDate, 'yyyy-MM-dd'));

    if (!error) setBatches(data.map(b => ({
      ...b,
      editor_name: b.editor?.full_name
    })));
  };

  if (loading && !DEV_BYPASS_LOGIN && !isAppAuthenticated) return <div className="h-screen flex items-center justify-center bg-gray-50"><Loader2 className="animate-spin text-gray-400" /></div>;

  const handleSaveBatchEntries = async (newEntries: Omit<TimeEntry, 'id'>[]) => {
    if (newEntries.length === 0) return;
    const token = await getGraphAccessToken();
    const range = {
      start: currentPeriod.startDate,
      end: new Date(currentPeriod.endDate.getTime() + 24 * 60 * 60 * 1000),
    };

    if (token) {
      const payload = newEntries.map((ne) => ({
        siteId: (ne as any).siteId,
        cleanerId: (ne as any).cleanerId,
        date: ne.date,
        hours: ne.hours,
        adhocJobId: (ne as any).adhocJobId ?? null,
      }));
      const result = await saveTimesheetEntriesToSharePoint(token, range, payload);
      if (result.error) {
        alert(result.error);
        return;
      }
      const entries = await getTimesheetEntriesForRange(token, range);
      setGraphEntries(entries.map((e) => ({
        id: e.id,
        batch_id: e.id,
        date: e.date,
        hours: e.hours,
        pay_rate_snapshot: e.pay_rate_snapshot,
        siteId: e.siteId,
        cleanerId: e.cleanerId,
        adhocJobId: e.adhocJobId,
        adhocJobName: e.adhocJobName,
      } as TimeEntry)));
      setGraphEntriesLoaded(true);
      return;
    }

    const { siteId, cleanerId } = newEntries[0] as any;
    const { data: userData } = await supabase.auth.getUser();
    const userId = userData.user?.id;
    
    let batch = batches.find(b => b.site_id === siteId && b.cleaner_id === cleanerId);
    
    if (!batch) {
      const { data, error } = await supabase.from('timesheet_batches').insert({
        site_id: siteId,
        cleaner_id: cleanerId,
        fortnight_start: format(currentPeriod.startDate, 'yyyy-MM-dd'),
        fortnight_end: format(currentPeriod.endDate, 'yyyy-MM-dd'),
        updated_by: userId
      }).select().single();
      if (error) {
        alert('Save failed. Sign in with Microsoft to save timesheets, or configure Supabase.');
        return;
      }
      batch = data;
    }

    const entriesToUpsert = newEntries.map(ne => ({
      batch_id: batch!.id,
      date: ne.date,
      hours: ne.hours,
      pay_rate_snapshot: ne.pay_rate_snapshot,
      updated_by: userId
    }));

    const { error: upsertError } = await supabase.from('timesheet_entries').upsert(entriesToUpsert, {
      onConflict: 'batch_id, date'
    });

    if (upsertError) alert(upsertError.message);
    fetchBatches();
  };

  const renderContent = () => {
    const flatEntries = batches.flatMap(b => b.entries || []).map(e => ({
      ...e,
      siteId: batches.find(b => b.id === e.batch_id)?.site_id,
      cleanerId: batches.find(b => b.id === e.batch_id)?.cleaner_id
    }));
    const entries = graphEntriesLoaded ? graphEntries : flatEntries;

    switch (currentView) {
      case 'dashboard':
        return <Dashboard sites={sites} cleaners={cleaners} entries={entries as any} currentPeriod={currentPeriod} onViewSite={id => { setSelectedSiteId(id); setCurrentView('site-detail'); }} />;
      case 'timesheets':
        return <TimeEntryForm sites={sites} cleaners={cleaners} entries={entries as any} currentPeriod={currentPeriod} onSaveBatch={handleSaveBatchEntries} onDeleteEntry={() => {}} onUpdateSite={fetchSites} />;
      case 'sites':
        return <SiteManager onUpdateSite={fetchSites} />;
      case 'team':
        return <TeamManager />;
      case 'cleaners':
        return <CleanerManager onCleanersRefresh={fetchCleaners} />;
      case 'adhoc-jobs':
        return <AdHocJobsManager />;
      case 'site-detail':
        const site = sites.find(s => s.id === selectedSiteId);
        if (!site) return null;
        return <SiteDetail site={site} cleaners={cleaners} entries={entries as any} currentPeriod={currentPeriod} onBack={() => setCurrentView('dashboard')} />;
      case 'auth-test':
        return <AuthTest />;
      default: return <div>Coming Soon</div>;
    }
  };

  return (
    <div className="flex min-h-screen bg-white min-w-0">
      <Sidebar
        currentView={currentView === 'site-detail' ? 'dashboard' : currentView}
        onViewChange={setCurrentView}
        mobileOpen={sidebarMobileOpen}
        onMobileClose={() => setSidebarMobileOpen(false)}
      />
      <main className="flex-1 overflow-x-hidden flex flex-col min-w-0 lg:pl-60">
        <nav className="min-h-12 border-b border-[#edeef0] flex flex-wrap items-center justify-between gap-2 px-3 sm:px-6 py-2 sticky top-0 bg-white/95 backdrop-blur-md z-30">
          <div className="flex items-center gap-2 min-w-0">
            <button
              type="button"
              onClick={() => setSidebarMobileOpen(true)}
              className="lg:hidden p-2 -ml-2 rounded-md text-gray-600 hover:bg-black/5 touch-manipulation"
              aria-label="Open menu"
            >
              <Menu size={20} />
            </button>
            <div className="flex items-center gap-2 text-sm text-gray-500 min-w-0 truncate">
              <span className="truncate">CleanTrack</span>
              <span>/</span>
              <span className="text-gray-900 font-medium capitalize truncate">{currentView.replace('-', ' ')}</span>
            </div>
          </div>
          <div className="flex items-center gap-2 sm:gap-4 flex-shrink-0">
            <DevBypassBanner />
            <div className="flex items-center gap-1 text-[11px] sm:text-[12px] text-gray-400 font-medium">
              <button onClick={() => setCurrentPeriod(getFortnightForDate(new Date(currentPeriod.startDate.getTime() - 14 * 24 * 60 * 60 * 1000)))} className="p-1.5 sm:p-1 hover:bg-black/5 rounded touch-manipulation" aria-label="Previous period"><ChevronLeft size={16} /></button>
              <span className="px-1 sm:px-2 whitespace-nowrap">{format(currentPeriod.startDate, 'MMM d')} — {format(currentPeriod.endDate, 'MMM d')}</span>
              <button onClick={() => setCurrentPeriod(getFortnightForDate(new Date(currentPeriod.startDate.getTime() + 14 * 24 * 60 * 60 * 1000)))} className="p-1.5 sm:p-1 hover:bg-black/5 rounded touch-manipulation" aria-label="Next period"><ChevronRight size={16} /></button>
            </div>
            <button onClick={() => { logout(); signOutApp(); }} className="text-[10px] font-bold text-red-500 hover:text-red-700 uppercase tracking-widest py-2 px-2 touch-manipulation">Logout</button>
          </div>
        </nav>
        <div className={`flex-1 w-full mx-auto p-4 sm:p-6 lg:p-12 box-border ${['sites', 'dashboard', 'cleaners', 'team', 'adhoc-jobs'].includes(currentView) ? 'max-w-7xl' : 'max-w-5xl'}`}>
          <header className="mb-6 sm:mb-12">
             <h1 className="text-2xl sm:text-4xl font-bold text-gray-900 tracking-tight capitalize mb-2">{currentView.replace('-', ' ')}</h1>
             <div className="w-full h-[1px] bg-[#edeef0] mt-4 sm:mt-8"></div>
          </header>
          {dataLoading ? <Loader2 className="animate-spin text-gray-200" /> : renderContent()}
        </div>
      </main>
    </div>
  );
};

const App: FC = () => (
  <AppAuthProvider>
    <AppGate />
  </AppAuthProvider>
);

export default App;