import React from "react";
import { useAppAuth } from "../contexts/AppAuthContext";
import { DEV_SHOW_LEGACY_LOGIN } from "../config/authFlags";

/** Microsoft logo (4 quadrants) for the sign-in button */
const MicrosoftLogo: React.FC<{ className?: string }> = ({ className }) => (
  <svg className={className} viewBox="0 0 21 21" width="20" height="20" aria-hidden>
    <rect x="0" y="0" width="10" height="10" fill="#F25022" />
    <rect x="11" y="0" width="10" height="10" fill="#7FBA00" />
    <rect x="0" y="11" width="10" height="10" fill="#00A4EF" />
    <rect x="11" y="11" width="10" height="10" fill="#FFB900" />
  </svg>
);

const SignInScreen: React.FC = () => {
  const { signInWithMicrosoft, isMicrosoftAuthConfigured } = useAppAuth();
  // Legacy Supabase login has been removed (SharePoint-only).

  return (
    <div className="min-h-screen flex bg-[#F6F7F8]">
      {/* Left: sign-in form — wider column, content vertically centered */}
      <div className="flex flex-col min-h-screen w-full lg:w-[55%] lg:max-w-[720px] px-8 md:px-12 py-8">
        <div className="flex-1 flex flex-col justify-center items-center">
          <div className="w-full max-w-sm">
            <div className="flex items-center gap-3 mb-10">
              <div className="w-9 h-9 rounded-xl bg-white p-1.5 shadow-sm">
                <img
                  src="/images/a_adept.png"
                  alt="Adept Timesheet Ops logo"
                  className="w-full h-full object-contain"
                />
              </div>
              <div className="flex flex-col">
                <span className="text-[18px] font-semibold text-gray-900">
                  Adept Timesheet Ops
                </span>
                <span className="text-[13px] text-gray-500">
                  Operational control centre
                </span>
              </div>
            </div>

            <h1 className="text-[26px] font-semibold text-gray-900 mb-2">
              Sign in to workspace
            </h1>
            <p className="text-[14px] text-gray-500 mb-8">
              Use your organisation’s Microsoft account to access timesheets, sites
              and audits.
            </p>

            {isMicrosoftAuthConfigured ? (
              <button
                type="button"
                onClick={() => signInWithMicrosoft()}
                className="w-full flex items-center justify-center gap-3 py-3 px-4 so-btn-primary text-[14px] font-medium shadow-sm active:scale-[0.99] transition-transform"
              >
                <MicrosoftLogo />
                Log in with Microsoft
              </button>
            ) : (
              <div className="rounded-xl bg-amber-50 border border-amber-200 p-4 text-sm text-amber-800">
                Microsoft sign-in is not configured. Set VITE_AZURE_AD_CLIENT_ID and
                VITE_AZURE_AD_TENANT_ID to enable log in.
              </div>
            )}

          {DEV_SHOW_LEGACY_LOGIN && (
            <p className="mt-4 text-xs text-gray-500">
              Legacy email/password login has been removed.
            </p>
          )}
          </div>
        </div>

        <p className="text-[11px] text-gray-400 py-4">
          © {new Date().getFullYear()} Adept Services · Internal operations platform
        </p>
      </div>

      {/* Right: brand panel */}
      <div className="hidden lg:flex flex-1 relative bg-[#111827] overflow-hidden">
        <div className="absolute inset-0 opacity-[0.65]" aria-hidden>
          <div
            className="w-full h-full"
            style={{
              backgroundImage:
                'radial-gradient(circle at 0% 0%, rgba(148,163,184,0.18) 0, transparent 45%), radial-gradient(circle at 80% 20%, rgba(56,189,248,0.18) 0, transparent 40%), linear-gradient(135deg, #020617 0%, #0f172a 40%, #020617 100%)',
            }}
          />
        </div>
        <div className="absolute inset-0 flex items-center justify-center p-12 z-10">
          <div className="max-w-xs text-center space-y-3">
            <p className="text-slate-100 font-medium text-lg">
              Operational control for cleaning, timesheets and compliance.
            </p>
            <p className="text-[12px] text-slate-300">
              Built for managers and audits — not another generic portal.
            </p>
          </div>
        </div>
      </div>
    </div>
  );
};

export default SignInScreen;
