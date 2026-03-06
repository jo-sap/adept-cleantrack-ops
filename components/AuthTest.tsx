import React, { useState } from "react";
import { InteractionRequiredAuthError } from "@azure/msal-browser";
import { msalInstance } from "../lib/msal";
import { GRAPH_SCOPES } from "../src/auth/graphScopes";
import { graphGet, getGraphAccessToken } from "../lib/graph";
import * as sharepoint from "../lib/sharepoint";
import { useAppAuth } from "../contexts/AppAuthContext";
import { DEV_BYPASS_LOGIN, isMicrosoftAuthConfigured } from "../config/authFlags";

interface MeResponse {
  displayName?: string;
  mail?: string;
  userPrincipalName?: string;
}

const CLEANTRACK_SITES_LIST_NAME = "CleanTrack Sites";
const SITE_FIELDS = ["Title", "SiteName", "Client", "Address", "Suburb", "State", "Postcode", "Frequency", "Notes"];

const AuthTest: React.FC = () => {
  const { authStatus, user } = useAppAuth();
  const [account, setAccount] = useState(msalInstance.getAllAccounts()[0] ?? null);
  const [graphResult, setGraphResult] = useState<{ ok: boolean; status: number; data?: MeResponse; text?: string } | null>(null);
  const [loading, setLoading] = useState(false);

  const [siteId, setSiteId] = useState<string | null>(null);
  const [lists, setLists] = useState<Array<{ id: string; displayName: string }>>([]);
  const [listError, setListError] = useState<string | null>(null);
  const [sitesLoading, setSitesLoading] = useState(false);
  const [sitesItems, setSitesItems] = useState<sharepoint.GraphListItem[]>([]);
  const [sitesError, setSitesError] = useState<string | null>(null);
  const [availableListNames, setAvailableListNames] = useState<string[]>([]);
  const [expandedRowId, setExpandedRowId] = useState<string | null>(null);
  const [tokenScopes, setTokenScopes] = useState<string[] | null>(null);
  const [tokenScopesError, setTokenScopesError] = useState<string | null>(null);

  const handleShowTokenScopes = async () => {
    setTokenScopesError(null);
    setTokenScopes(null);
    const acc = msalInstance.getAllAccounts()[0];
    if (!acc) {
      setTokenScopesError("No account. Sign in first.");
      return;
    }
    try {
      const response = await msalInstance.acquireTokenSilent({ account: acc, scopes: GRAPH_SCOPES });
      setTokenScopes(response.scopes ?? []);
    } catch (err) {
      if (err instanceof InteractionRequiredAuthError) {
        try {
          const response = await msalInstance.acquireTokenPopup({ scopes: GRAPH_SCOPES });
          setTokenScopes(response.scopes ?? []);
        } catch (e) {
          setTokenScopesError(e instanceof Error ? e.message : "Failed to get token");
        }
      } else {
        setTokenScopesError(err instanceof Error ? err.message : "Failed to get token");
      }
    }
  };

  const hasSitesReadWriteAll =
    tokenScopes != null &&
    tokenScopes.some(
      (s) =>
        s === "Sites.ReadWrite.All" || s.endsWith("/Sites.ReadWrite.All")
    );

  const handleSignIn = async () => {
    setGraphResult(null);
    setLoading(true);
    try {
      await msalInstance.loginPopup({ scopes: GRAPH_SCOPES });
      setAccount(msalInstance.getAllAccounts()[0] ?? null);
    } catch (err) {
      setGraphResult({
        ok: false,
        status: 0,
        text: err instanceof Error ? err.message : "Sign-in failed",
      });
    } finally {
      setLoading(false);
    }
  };

  const handleSignOut = async () => {
    setGraphResult(null);
    setLoading(true);
    try {
      await msalInstance.logoutPopup();
      setAccount(null);
    } catch (err) {
      setGraphResult({
        ok: false,
        status: 0,
        text: err instanceof Error ? err.message : "Sign-out failed",
      });
    } finally {
      setLoading(false);
    }
  };

  const handleTestGraph = async () => {
    setGraphResult(null);
    setLoading(true);
    try {
      const result = await graphGet<MeResponse>("https://graph.microsoft.com/v1.0/me", GRAPH_SCOPES);
      setGraphResult(result);
    } catch (err) {
      setGraphResult({
        ok: false,
        status: 0,
        text: err instanceof Error ? err.message : "Graph request failed",
      });
    } finally {
      setLoading(false);
    }
  };

  const noTokenMessage = DEV_BYPASS_LOGIN
    ? "DEV bypass enabled — Graph calls disabled (no token)."
    : "No Graph access token. Sign in with Microsoft (non-bypass) and ensure Graph scopes are enabled.";

  const handleLoadSiteAndLists = async () => {
    setListError(null);
    setSiteId(null);
    setLists([]);
    const token = await getGraphAccessToken();
    if (!token) {
      setListError(noTokenMessage);
      return;
    }
    setLoading(true);
    try {
      const id = await sharepoint.getSiteId(token);
      setSiteId(id);
      const listArr = await sharepoint.getLists(token, id);
      setLists(listArr);
    } catch (err) {
      setListError(err instanceof Error ? err.message : "Failed to load site/lists");
    } finally {
      setLoading(false);
    }
  };

  const handleLoadSites = async () => {
    setSitesError(null);
    setSitesItems([]);
    setAvailableListNames([]);
    const token = await getGraphAccessToken();
    if (!token) {
      setSitesError(noTokenMessage);
      return;
    }
    setSitesLoading(true);
    try {
      let sid = siteId ?? null;
      if (!sid) {
        sid = await sharepoint.getSiteId(token);
        setSiteId(sid);
      }
      const listArr = await sharepoint.getLists(token, sid);
      setAvailableListNames(listArr.map((l) => l.displayName));
      const listId = await sharepoint.getListIdByName(token, sid, CLEANTRACK_SITES_LIST_NAME);
      if (listId == null) {
        setSitesError(`List not found: "${CLEANTRACK_SITES_LIST_NAME}". Available lists: ${listArr.map((l) => l.displayName).join(", ") || "none"}`);
        return;
      }
      const items = await sharepoint.getListItems(token, sid, listId);
      setSitesItems(items);
    } catch (err) {
      setSitesError(err instanceof Error ? err.message : "Failed to load CleanTrack Sites");
    } finally {
      setSitesLoading(false);
    }
  };

  return (
    <div className="max-w-lg space-y-6">
      <div className="rounded-xl border border-[#edeef0] bg-white shadow-sm overflow-hidden">
        <div className="px-5 py-4 border-b border-[#edeef0]">
          <h2 className="text-lg font-semibold text-gray-900">Auth state</h2>
        </div>
        <div className="p-5 space-y-2 text-sm font-mono">
          <p><span className="text-gray-500">DEV_BYPASS_LOGIN:</span> {String(DEV_BYPASS_LOGIN)}</p>
          <p><span className="text-gray-500">authStatus:</span> {authStatus}</p>
          <p><span className="text-gray-500">user:</span> {user ? `${user.name} (${user.email})` : "null"}</p>
          <p><span className="text-gray-500">Microsoft env configured:</span> {String(isMicrosoftAuthConfigured)}</p>
        </div>
      </div>
      <div className="rounded-xl border border-[#edeef0] bg-white shadow-sm overflow-hidden">
        <div className="px-5 py-4 border-b border-[#edeef0]">
          <h2 className="text-lg font-semibold text-gray-900">Microsoft Login</h2>
        </div>
        <div className="p-5 space-y-4">
          {!account ? (
            <>
              <p className="text-sm text-gray-500">Sign in with your Microsoft 365 account to test Graph API access.</p>
              <button
                type="button"
                onClick={handleSignIn}
                disabled={loading}
                className="px-4 py-2 bg-[#238636] hover:bg-[#2ea043] text-white text-sm font-medium rounded-lg transition-colors disabled:opacity-50"
              >
                {loading ? "Signing in…" : "Sign in with Microsoft"}
              </button>
            </>
          ) : (
            <>
              <p className="text-sm text-gray-700">
                Signed in as: <span className="font-medium text-gray-900">{account.username}</span>
              </p>
              <div className="flex flex-wrap gap-2">
                <button
                  type="button"
                  onClick={handleTestGraph}
                  disabled={loading}
                  className="px-4 py-2 bg-gray-100 hover:bg-gray-200 text-gray-800 text-sm font-medium rounded-lg transition-colors disabled:opacity-50"
                >
                  {loading ? "Loading…" : "Test Graph /me"}
                </button>
                <button
                  type="button"
                  onClick={handleSignOut}
                  disabled={loading}
                  className="px-4 py-2 text-red-600 hover:bg-red-50 text-sm font-medium rounded-lg transition-colors disabled:opacity-50"
                >
                  Sign out
                </button>
              </div>
              <div className="pt-2 border-t border-gray-200">
                <p className="text-xs font-medium text-gray-500 mb-1">Token scopes (debug)</p>
                <button
                  type="button"
                  onClick={handleShowTokenScopes}
                  disabled={loading}
                  className="px-3 py-1.5 bg-gray-100 hover:bg-gray-200 text-gray-700 text-xs font-medium rounded transition-colors disabled:opacity-50"
                >
                  Show token scopes
                </button>
                {tokenScopesError && (
                  <p className="text-xs text-red-600 mt-1">{tokenScopesError}</p>
                )}
                {tokenScopes != null && (
                  <div className="mt-2">
                    <div className="flex items-center gap-2 flex-wrap">
                      {tokenScopes.length === 0 && (
                        <span className="text-xs text-gray-500">(none)</span>
                      )}
                      {tokenScopes.map((s) => (
                        <span
                          key={s}
                          className="text-xs font-mono text-gray-600 bg-gray-100 px-1.5 py-0.5 rounded"
                        >
                          {s}
                        </span>
                      ))}
                      {hasSitesReadWriteAll && (
                        <span className="inline-flex items-center px-2 py-0.5 rounded text-xs font-medium bg-green-100 text-green-800">
                          Sites.ReadWrite.All present
                        </span>
                      )}
                    </div>
                  </div>
                )}
              </div>
              {graphResult?.ok && graphResult.data && (
                <div className="rounded-lg bg-[#f0fdf4] border border-[#bbf7d0] p-3 text-sm">
                  <p className="font-medium text-gray-900">{graphResult.data.displayName ?? "—"}</p>
                  <p className="text-gray-600">{graphResult.data.mail ?? graphResult.data.userPrincipalName ?? "—"}</p>
                </div>
              )}
              {graphResult && !graphResult.ok && (
                <div className="rounded-lg bg-red-50 border border-red-200 p-3 text-sm">
                  <p className="font-medium text-red-800">Error {graphResult.status}</p>
                  <p className="text-red-700 mt-1">{graphResult.text ?? "Unknown error"}</p>
                </div>
              )}
            </>
          )}
        </div>
      </div>

      <div className="rounded-xl border border-[#edeef0] bg-white shadow-sm overflow-hidden">
        <div className="px-5 py-4 border-b border-[#edeef0]">
          <h2 className="text-lg font-semibold text-gray-900">SharePoint (read-only)</h2>
        </div>
        <div className="p-5 space-y-4">
          {DEV_BYPASS_LOGIN && (
            <p className="text-sm text-amber-700 bg-amber-50 border border-amber-200 rounded-lg px-3 py-2">
              {noTokenMessage}
            </p>
          )}
          <div className="flex flex-wrap gap-2">
            <button
              type="button"
              onClick={handleLoadSiteAndLists}
              disabled={loading}
              className="px-4 py-2 bg-gray-100 hover:bg-gray-200 text-gray-800 text-sm font-medium rounded-lg transition-colors disabled:opacity-50"
            >
              {loading ? "Loading…" : "Load Site + Lists"}
            </button>
            <button
              type="button"
              onClick={handleLoadSites}
              disabled={sitesLoading}
              className="px-4 py-2 bg-gray-100 hover:bg-gray-200 text-gray-800 text-sm font-medium rounded-lg transition-colors disabled:opacity-50"
            >
              {sitesLoading ? "Loading…" : "Load Sites"}
            </button>
          </div>
          {listError && (
            <div className="rounded-lg bg-red-50 border border-red-200 p-3 text-sm text-red-800">{listError}</div>
          )}
          {siteId && (
            <div className="text-sm">
              <p className="font-medium text-gray-700">Site ID</p>
              <p className="font-mono text-xs text-gray-600 break-all mt-1">{siteId}</p>
            </div>
          )}
          {lists.length > 0 && (
            <div className="text-sm">
              <p className="font-medium text-gray-700 mb-1">Lists</p>
              <ul className="list-disc list-inside space-y-0.5 font-mono text-xs text-gray-600">
                {lists.map((l) => (
                  <li key={l.id}>{l.displayName} <span className="text-gray-400">({l.id})</span></li>
                ))}
              </ul>
            </div>
          )}
          {sitesError && (
            <div className="rounded-lg bg-red-50 border border-red-200 p-3 text-sm text-red-800">{sitesError}</div>
          )}
          {sitesItems.length > 0 && (
            <div className="overflow-x-auto">
              <p className="font-medium text-gray-700 mb-2">CleanTrack Sites list</p>
              <table className="w-full text-sm border border-[#edeef0] rounded-lg overflow-hidden">
                <thead>
                  <tr className="bg-gray-50 border-b border-[#edeef0]">
                    {SITE_FIELDS.map((f) => (
                      <th key={f} className="text-left px-3 py-2 font-medium text-gray-700">{f}</th>
                    ))}
                    <th className="w-10" />
                  </tr>
                </thead>
                <tbody>
                  {sitesItems.map((item) => {
                    const fields = item.fields ?? {};
                    const expanded = expandedRowId === item.id;
                    return (
                      <React.Fragment key={item.id}>
                        <tr className="border-b border-[#edeef0] hover:bg-gray-50/50">
                          {SITE_FIELDS.map((key) => (
                            <td key={key} className="px-3 py-2 text-gray-800 max-w-[200px] truncate" title={String(fields[key] ?? "")}>
                              {String(fields[key] ?? "—")}
                            </td>
                          ))}
                          <td>
                            <button
                              type="button"
                              onClick={() => setExpandedRowId(expanded ? null : item.id)}
                              className="text-[10px] text-gray-500 hover:text-gray-700"
                            >
                              {expanded ? "Hide" : "Raw"}
                            </button>
                          </td>
                        </tr>
                        {expanded && (
                          <tr className="bg-gray-50 border-b border-[#edeef0]">
                            <td colSpan={SITE_FIELDS.length + 1} className="px-3 py-2">
                              <pre className="text-xs text-gray-600 overflow-auto max-h-40 whitespace-pre-wrap break-words">
                                {JSON.stringify(fields, null, 2)}
                              </pre>
                            </td>
                          </tr>
                        )}
                      </React.Fragment>
                    );
                  })}
                </tbody>
              </table>
            </div>
          )}
        </div>
      </div>
    </div>
  );
};

export default AuthTest;
