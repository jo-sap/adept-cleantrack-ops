import React from "react";
import { useAppAuth } from "../contexts/AppAuthContext";
import { DEV_SHOW_LEGACY_LOGIN } from "../config/authFlags";
import Login from "./Login";

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
  const [showLegacy, setShowLegacy] = React.useState(false);

  if (DEV_SHOW_LEGACY_LOGIN && showLegacy) {
    return <Login />;
  }

  return (
    <div className="min-h-screen flex bg-white">
      {/* Left: sign-in form — wider column, content vertically centered */}
      <div className="flex flex-col min-h-screen w-full lg:w-[55%] lg:max-w-[720px] p-8 md:p-12">
        <div className="flex-1 flex flex-col justify-center items-center">
          <div className="w-full max-w-sm">
            <div className="flex items-center gap-2 mb-10">
              <div className="w-9 h-9 rounded-lg bg-gray-900 flex items-center justify-center text-white font-bold text-lg">
                CT
              </div>
              <span className="text-xl font-semibold text-gray-900">CleanTrack Ops</span>
            </div>

          <h1 className="text-2xl font-bold text-gray-900 mb-1">Log in</h1>
          <p className="text-sm text-gray-500 mb-8">
            Sign in with your organisation’s Microsoft account to continue.
          </p>

          {isMicrosoftAuthConfigured ? (
            <button
              type="button"
              onClick={() => signInWithMicrosoft()}
              className="w-full flex items-center justify-center gap-3 py-3 px-4 rounded-xl border-2 border-gray-200 bg-white text-gray-900 font-semibold hover:bg-gray-50 hover:border-gray-300 transition-colors active:scale-[0.99]"
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
            <button
              type="button"
              onClick={() => setShowLegacy(true)}
              className="mt-4 text-xs text-gray-500 hover:text-gray-700"
            >
              Use email / password (legacy)
            </button>
          )}
          </div>
        </div>

        <p className="text-xs text-gray-400 py-4">
          © {new Date().getFullYear()} All Rights Reserved. Privacy and Terms.
        </p>
      </div>

      {/* Right: decorative panel — charcoal, centered tagline, "a" scaled down, right-aligned so front of letter shows */}
      <div className="hidden lg:flex flex-1 relative bg-[#2C2C2C] overflow-hidden">
        <div className="absolute inset-0 flex items-center justify-center p-12 z-10">
          <p className="text-gray-300 font-medium text-lg text-center max-w-xs">
            Enterprise cleaning audit and timesheet management
          </p>
        </div>
        <div className="absolute -bottom-[1%] -right-[10%] w-[100vmin] h-[100vmin]">
          <img
            src="/logo-a.png"
            alt=""
            className="w-full h-full object-contain object-bottom object-right select-none pointer-events-none mix-blend-screen"
            style={{ opacity: 0.2 }}
            aria-hidden
          />
        </div>
      </div>
    </div>
  );
};

export default SignInScreen;
