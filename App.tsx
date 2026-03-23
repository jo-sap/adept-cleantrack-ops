
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
import { Site, Cleaner, ViewType, FortnightPeriod, TimeEntry } from './types';
import { getFortnightForDate } from './utils';
import { ChevronLeft, ChevronRight, Loader2, Menu } from 'lucide-react';
import { format, addDays } from 'date-fns';
import { RoleProvider, useRole } from './contexts/RoleContext';
import { AppAuthProvider, useAppAuth } from './contexts/AppAuthContext';
import { DEV_BYPASS_LOGIN } from './config/authFlags';
import { getGraphAccessToken } from './lib/graph';
import { getCleaners } from './repositories/cleanersRepo';
import { getSites, toAppSite } from './repositories/sitesRepo';
import { getSiteBudgets } from './repositories/budgetsRepo';
import { getTimesheetEntriesForRange, saveTimesheetEntriesToSharePoint, type TimesheetEntryFlat } from './repositories/metricsRepo';
import { getAssignedSiteIdsForManager } from './repositories/siteManagersRepo';
import { getActiveCleanerIdsBySite } from './repositories/assignedCleanersRepo';

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
  /** View to return to when leaving Site Detail (e.g. 'dashboard' or 'sites'). */
  const [viewBeforeSiteDetail, setViewBeforeSiteDetail] = useState<ViewType>('dashboard');
  /** Increment when switching to Sites so SiteManager refetches assignments (e.g. after assigning on Site Detail). */
  const [sitesRefreshTrigger, setSitesRefreshTrigger] = useState(0);
  const prevViewRef = React.useRef<ViewType | null>(null);
  React.useEffect(() => {
    if (currentView === 'sites' && prevViewRef.current !== 'sites') {
      setSitesRefreshTrigger((t) => t + 1);
    }
    prevViewRef.current = currentView;
  }, [currentView]);
  const [sites, setSites] = useState<Site[]>([]);
  const [cleaners, setCleaners] = useState<Cleaner[]>([]);
  const [graphEntries, setGraphEntries] = useState<TimeEntry[]>([]);
  const [graphEntriesLoaded, setGraphEntriesLoaded] = useState(false);
  const [currentPeriod, setCurrentPeriod] = useState<FortnightPeriod>(getFortnightForDate(new Date()));
  const [dataLoading, setDataLoading] = useState(true);
  const [sidebarMobileOpen, setSidebarMobileOpen] = useState(false);

  /** Normalise a raw SharePoint timesheet entry so its site/cleaner ids line up with the current app sites/cleaners.
   *  This bridges cases where list item ids drift (e.g. sites or cleaners recreated) but names stay the same. */
  const mapFlatEntryToAppEntry = React.useCallback(
    (e: TimesheetEntryFlat): TimeEntry => {
      let siteId = e.siteId;
      let cleanerId = e.cleanerId;

      const siteById = new Map(sites.map((s) => [s.id, s]));
      const siteByName = new Map(
        sites.map((s) => [s.name?.trim().toLowerCase() ?? '', s] as const)
      ) as Map<string, Site>;
      const cleanerById = new Map(cleaners.map((c) => [c.id, c]));
      const cleanerByName = new Map(
        cleaners.map((c) => [`${c.firstName} ${c.lastName}`.trim().toLowerCase(), c] as const)
      ) as Map<string, Cleaner>;

      if (!siteById.has(siteId) && (e.siteName ?? '').trim()) {
        const s = siteByName.get((e.siteName ?? '').trim().toLowerCase());
        if (s) siteId = s.id;
      }
      if (!cleanerById.has(cleanerId) && (e.cleanerName ?? '').trim()) {
        const c = cleanerByName.get((e.cleanerName ?? '').trim().toLowerCase());
        if (c) cleanerId = c.id;
      }

      return {
        id: e.id,
        batch_id: e.id,
        date: e.date,
        hours: e.hours,
        pay_rate_snapshot: e.pay_rate_snapshot,
        siteId,
        cleanerId,
        adhocJobId: e.adhocJobId,
        adhocJobName: e.adhocJobName,
      } as TimeEntry;
    },
    [sites, cleaners]
  );

  const fetchSites = async (silent = false) => {
    if (!silent) setDataLoading(true);
    try {
      const token = await getGraphAccessToken();
      if (!token) {
        setSites([]);
        return;
      }
      let list = await getSites(token);
      if (user?.role === 'Manager' && user?.email) {
        const isAllSites = user.permissionScope?.trim().toLowerCase() === 'allsites';
        if (!isAllSites) {
          const assignedIds = await getAssignedSiteIdsForManager(token, user.email);
          if (assignedIds.length > 0) list = list.filter((s) => assignedIds.includes(s.id));
        }
      }
      const budgets = await getSiteBudgets(token).catch(() => ({}));
      const activeAssignmentsBySite = await getActiveCleanerIdsBySite(token).catch(() => ({} as Record<string, string[]>));
      setSites(
        list.map((s) => {
          const budget =
            budgets[String(s.id)] ??
            budgets["name:" + (s.siteName.trim() + " Budget")];
          const appSite = toAppSite(s, budget) as Site;
          const assignedIds = activeAssignmentsBySite[s.id] ?? [];
          const withAssignments: Site = {
            ...appSite,
            assigned_cleaner_ids: assignedIds,
          };
          return withAssignments;
        })
      );
    } catch (e) {
      console.error("fetchSites failed", e);
      setSites([]);
    } finally {
      if (!silent) setDataLoading(false);
    }
  };

  // Do not depend on currentPeriod here: refetching sites sets dataLoading and unmounts the main view,
  // which resets Timesheets drill-in state (active site) when the user changes fortnight with the header chevrons.
  useEffect(() => {
    if (role || DEV_BYPASS_LOGIN || isAppAuthenticated) {
      fetchSites();
      fetchCleaners();
    }
  }, [role, isAppAuthenticated, user?.role, user?.email]);

  /** Load current + prior fortnight so Timesheets can copy from the previous period without a second request. */
  const timesheetFetchRange = React.useMemo(
    () => ({
      start: addDays(currentPeriod.startDate, -14),
      end: new Date(currentPeriod.endDate.getTime() + 24 * 60 * 60 * 1000),
    }),
    [currentPeriod.startDate.getTime(), currentPeriod.endDate.getTime()]
  );

  useEffect(() => {
    let cancelled = false;
    setGraphEntriesLoaded(false);
    getGraphAccessToken().then((token) => {
      if (!token || cancelled) {
        if (!cancelled) setGraphEntries([]);
        return;
      }
      getTimesheetEntriesForRange(token, timesheetFetchRange).then((entries) => {
        if (cancelled) return;
        setGraphEntries(entries.map((e) => mapFlatEntryToAppEntry(e)));
        setGraphEntriesLoaded(true);
      }).catch(() => { if (!cancelled) setGraphEntriesLoaded(false); });
    });
    return () => { cancelled = true; };
  }, [timesheetFetchRange.start.getTime(), timesheetFetchRange.end.getTime(), mapFlatEntryToAppEntry]);

  const fetchCleaners = async () => {
    const token = await getGraphAccessToken();
    if (!token) {
      setCleaners([]);
      return;
    }
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
      // 1. Optimistic update: merge the batch we just saved into state so the UI shows saved hours immediately
      //    (avoids refetch returning before SharePoint has the new items and wiping the form to zeros)
      const optimisticEntries: TimeEntry[] = payload.map((p, i) => ({
        id: `opt-${p.siteId}-${p.cleanerId}-${p.date}`,
        batch_id: `opt-${p.siteId}-${p.cleanerId}-${p.date}`,
        date: p.date,
        hours: p.hours,
        siteId: p.siteId,
        cleanerId: p.cleanerId,
        adhocJobId: (p as any).adhocJobId ?? undefined,
        pay_rate_snapshot: (newEntries[i] as any).pay_rate_snapshot,
      } as TimeEntry));
      setGraphEntries((prev) => {
        const byKey = new Map<string, TimeEntry>();
        prev.forEach((e) => byKey.set(`${e.siteId}|${e.cleanerId}|${e.date}`, e));
        optimisticEntries.forEach((e) => byKey.set(`${e.siteId}|${e.cleanerId}|${e.date}`, e));
        return Array.from(byKey.values());
      });
      setGraphEntriesLoaded(true);

      // 2. Refetch from SharePoint and merge server data (server wins per key) so we get real IDs
      getTimesheetEntriesForRange(token, timesheetFetchRange)
        .then((entriesFromServer) => {
          const mapped = entriesFromServer.map((e) => mapFlatEntryToAppEntry(e));
          setGraphEntries((prev) => {
            const byKey = new Map<string, TimeEntry>();
            prev.forEach((e) => byKey.set(`${e.siteId}|${e.cleanerId}|${e.date}`, e));
            mapped.forEach((e) => byKey.set(`${e.siteId}|${e.cleanerId}|${e.date}`, e));
            return Array.from(byKey.values());
          });
        })
        .catch(() => {});
      return;
    }

    alert('Save failed. Sign in with Microsoft to save timesheets.');
  };

  const renderContent = () => {
    const entries = graphEntriesLoaded ? graphEntries : [];

    switch (currentView) {
      case 'dashboard':
        return <Dashboard sites={sites} cleaners={cleaners} entries={entries as any} currentPeriod={currentPeriod} onViewSite={id => { setViewBeforeSiteDetail('dashboard'); setSelectedSiteId(id); setCurrentView('site-detail'); }} />;
      case 'timesheets':
        return <TimeEntryForm sites={sites} cleaners={cleaners} entries={entries as any} currentPeriod={currentPeriod} onSaveBatch={handleSaveBatchEntries} onDeleteEntry={() => {}} onUpdateSite={fetchSites} />;
      case 'sites':
        return (
          <SiteManager
            onUpdateSite={(silent?: boolean) => fetchSites(silent === true)}
            onViewSite={(id) => {
              setViewBeforeSiteDetail('sites');
              setSelectedSiteId(id);
              setCurrentView('site-detail');
            }}
            refreshTrigger={sitesRefreshTrigger}
          />
        );
      case 'team':
        return <TeamManager />;
      case 'cleaners':
        return <CleanerManager onCleanersRefresh={fetchCleaners} />;
      case 'adhoc-jobs':
        return <AdHocJobsManager />;
      case 'site-detail':
        const site = sites.find(s => s.id === selectedSiteId);
        if (!site) return null;
        const backLabel = viewBeforeSiteDetail === 'sites' ? 'Sites & Budgets' : viewBeforeSiteDetail === 'dashboard' ? 'Dashboard' : viewBeforeSiteDetail.replace(/-/g, ' ');
        return (
          <SiteDetail
            site={site}
            cleaners={cleaners}
            entries={entries as any}
            currentPeriod={currentPeriod}
            onBack={() => setCurrentView(viewBeforeSiteDetail)}
            backLabel={backLabel}
            onRefreshSites={fetchSites}
          />
        );
      case 'auth-test':
        return <AuthTest />;
      default: return <div>Coming Soon</div>;
    }
  };

  return (
    <div className="flex min-h-screen so-app-shell min-w-0">
      <Sidebar
        currentView={currentView === 'site-detail' ? 'dashboard' : currentView}
        onViewChange={setCurrentView}
        mobileOpen={sidebarMobileOpen}
        onMobileClose={() => setSidebarMobileOpen(false)}
      />
      <main className="flex-1 flex flex-col min-w-0 lg:pl-60">
        {/* Sticky nav must not sit inside an overflow-x-hidden parent (breaks position:sticky). */}
        <nav className="min-h-12 shrink-0 border-b border-[#E5E7EB] flex flex-wrap items-center justify-between gap-2 px-3 sm:px-6 py-2 sticky top-0 bg-[#FCFCFD]/95 backdrop-blur-md z-40 shadow-[0_1px_0_rgba(0,0,0,0.04)]">
          <div className="flex items-center gap-2 min-w-0">
            <button
              type="button"
              onClick={() => setSidebarMobileOpen(true)}
              className="lg:hidden p-2 -ml-2 rounded-md text-gray-600 hover:bg-black/5 touch-manipulation"
              aria-label="Open menu"
            >
              <Menu size={20} />
            </button>
            <div className="flex items-center gap-2 text-xs sm:text-sm text-gray-500 min-w-0 truncate">
              <span className="truncate text-gray-400">Adept Timesheet Ops</span>
              <span className="text-gray-300">/</span>
              <span className="text-gray-900 font-medium capitalize truncate">
                {currentView.replace('-', ' ')}
              </span>
            </div>
          </div>
          <div className="flex items-center gap-2 sm:gap-4 flex-shrink-0">
            <DevBypassBanner />
            <div className="flex items-center gap-1 text-[11px] sm:text-[12px] text-gray-500 font-medium">
              <button
                onClick={() =>
                  setCurrentPeriod(
                    getFortnightForDate(
                      new Date(
                        currentPeriod.startDate.getTime() -
                          14 * 24 * 60 * 60 * 1000
                      )
                    )
                  )
                }
                className="p-1.5 sm:p-1 hover:bg-gray-100 rounded-md touch-manipulation"
                aria-label="Previous period"
              >
                <ChevronLeft size={16} />
              </button>
              <span className="px-1 sm:px-2 whitespace-nowrap text-gray-700">
                {format(currentPeriod.startDate, 'MMM d')} —{' '}
                {format(currentPeriod.endDate, 'MMM d')}
              </span>
              <button
                onClick={() =>
                  setCurrentPeriod(
                    getFortnightForDate(
                      new Date(
                        currentPeriod.startDate.getTime() +
                          14 * 24 * 60 * 60 * 1000
                      )
                    )
                  )
                }
                className="p-1.5 sm:p-1 hover:bg-gray-100 rounded-md touch-manipulation"
                aria-label="Next period"
              >
                <ChevronRight size={16} />
              </button>
            </div>
            <button
              onClick={() => {
                logout();
                signOutApp();
              }}
              className="text-[11px] font-medium text-gray-500 hover:text-gray-900 px-2 py-1.5 rounded-md hover:bg-gray-100 touch-manipulation"
            >
              Logout
            </button>
          </div>
        </nav>
        <div className="flex-1 min-w-0 overflow-x-hidden">
          <div
            className={`flex-1 w-full mx-auto p-4 sm:p-6 lg:p-10 box-border ${
              ['sites', 'dashboard', 'cleaners', 'team', 'adhoc-jobs'].includes(
                currentView
              )
                ? 'max-w-7xl'
                : 'max-w-5xl'
            }`}
          >
            <header className="so-page-header">
              <h1 className="text-[24px] sm:text-[30px] font-semibold text-gray-900 tracking-tight capitalize">
                {currentView.replace('-', ' ')}
              </h1>
            </header>
            {dataLoading ? (
              <Loader2 className="animate-spin text-gray-400" aria-label="Loading" />
            ) : (
              renderContent()
            )}
          </div>
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