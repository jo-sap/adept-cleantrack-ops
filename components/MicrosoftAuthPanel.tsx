
import React, { useState } from 'react';
import { useMsal, AuthenticatedTemplate, UnauthenticatedTemplate } from "@azure/msal-react";
import { GRAPH_SCOPES } from "../src/auth/graphScopes";
import { graphGet, GraphResponse } from "../lib/graph";
import { Network, LogIn, LogOut, CheckCircle2, AlertCircle, Loader2 } from 'lucide-react';

const MicrosoftAuthPanel: React.FC = () => {
  const { instance, accounts } = useMsal();
  const [testResult, setTestResult] = useState<GraphResponse | null>(null);
  const [loading, setLoading] = useState(false);

  const handleLogin = () => {
    instance.loginPopup({ scopes: GRAPH_SCOPES }).catch(e => console.error(e));
  };

  const handleLogout = () => {
    instance.logoutPopup().catch(e => console.error(e));
  };

  const handleTestConnection = async () => {
    setLoading(true);
    try {
      const result = await graphGet("https://graph.microsoft.com/v1.0/me", GRAPH_SCOPES);
      setTestResult(result);
    } catch (err) {
      console.error("Test connection failed", err);
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="bg-white border border-[#edeef0] rounded-xl p-6 shadow-sm mb-8 animate-fadeIn">
      <div className="flex items-center justify-between mb-4">
        <div className="flex items-center gap-2">
          <div className="bg-blue-50 p-2 rounded-lg text-blue-600">
            <Network size={20} />
          </div>
          <h3 className="text-sm font-bold text-gray-900">Microsoft Integration</h3>
        </div>
        
        <div className="flex gap-2">
          <UnauthenticatedTemplate>
            <button 
              onClick={handleLogin}
              className="flex items-center gap-2 bg-white border border-[#edeef0] hover:bg-gray-50 text-gray-700 px-4 py-2 rounded-lg text-xs font-bold transition-all shadow-sm"
            >
              <LogIn size={14} /> Sign in with Microsoft
            </button>
          </UnauthenticatedTemplate>
          
          <AuthenticatedTemplate>
            <button 
              onClick={handleLogout}
              className="flex items-center gap-2 bg-white border border-red-50 hover:bg-red-50 text-red-600 px-4 py-2 rounded-lg text-xs font-bold transition-all shadow-sm"
            >
              <LogOut size={14} /> Sign out
            </button>
          </AuthenticatedTemplate>
        </div>
      </div>

      <AuthenticatedTemplate>
        <div className="space-y-4">
          <div className="bg-[#fcfcfb] p-3 rounded-lg border border-[#edeef0]">
            <p className="text-[10px] font-bold text-gray-400 uppercase tracking-widest mb-1">Signed in as</p>
            <p className="text-xs font-semibold text-gray-700">{accounts[0]?.username || 'Unknown'}</p>
          </div>

          <button 
            onClick={handleTestConnection}
            disabled={loading}
            className="w-full flex items-center justify-center gap-2 bg-gray-900 hover:bg-black text-white px-4 py-2 rounded-lg text-xs font-bold transition-all disabled:opacity-50 shadow-md"
          >
            {loading ? <Loader2 size={14} className="animate-spin" /> : <Network size={14} />}
            Test Microsoft Connection
          </button>

          {testResult && (
            <div className={`p-4 rounded-lg border animate-slideDown ${testResult.ok ? 'bg-green-50 border-green-100' : 'bg-red-50 border-red-100'}`}>
              <div className="flex items-center gap-2 mb-2">
                {testResult.ok ? (
                  <CheckCircle2 size={16} className="text-green-600" />
                ) : (
                  <AlertCircle size={16} className="text-red-600" />
                )}
                <span className={`text-[11px] font-bold uppercase ${testResult.ok ? 'text-green-700' : 'text-red-700'}`}>
                  {testResult.ok ? 'Connection Successful' : `Error (${testResult.status})`}
                </span>
              </div>
              
              {testResult.json ? (
                <div className="grid grid-cols-2 gap-2 mt-2">
                  <div className="bg-white/50 p-2 rounded border border-green-200/50">
                    <p className="text-[9px] font-bold text-gray-400 uppercase">Name</p>
                    <p className="text-[11px] font-bold text-gray-800">{testResult.json.displayName}</p>
                  </div>
                  <div className="bg-white/50 p-2 rounded border border-green-200/50">
                    <p className="text-[9px] font-bold text-gray-400 uppercase">UPN</p>
                    <p className="text-[11px] font-bold text-gray-800 truncate">{testResult.json.userPrincipalName || testResult.json.mail}</p>
                  </div>
                </div>
              ) : (
                <pre className="text-[10px] bg-white/50 p-2 rounded mt-2 overflow-x-auto text-gray-600 font-mono">
                  {testResult.text}
                </pre>
              )}
            </div>
          )}
        </div>
      </AuthenticatedTemplate>

      <UnauthenticatedTemplate>
        <p className="text-xs text-gray-500 bg-gray-50 p-3 rounded-lg border border-dashed border-[#edeef0]">
          Sign in with your organization's Microsoft account to enable SharePoint synchronization and test connectivity.
        </p>
      </UnauthenticatedTemplate>
    </div>
  );
};

export default MicrosoftAuthPanel;
