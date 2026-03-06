import React from "react";
import { ShieldCheck } from "lucide-react";
import { useAppAuth } from "../contexts/AppAuthContext";
import { DEV_SHOW_LEGACY_LOGIN } from "../config/authFlags";
import Login from "./Login";

const SignInScreen: React.FC = () => {
  const { signInWithMicrosoft, isMicrosoftAuthConfigured } = useAppAuth();
  const [showLegacy, setShowLegacy] = React.useState(false);

  if (DEV_SHOW_LEGACY_LOGIN && showLegacy) {
    return <Login />;
  }

  return (
    <div className="min-h-screen flex items-center justify-center bg-gray-50 p-6">
      <div className="w-full max-w-md bg-white rounded-3xl shadow-xl border border-gray-100 overflow-hidden">
        <div className="p-8 text-center bg-gray-900 text-white">
          <div className="bg-white/10 w-12 h-12 rounded-xl flex items-center justify-center mx-auto mb-4">
            <ShieldCheck size={24} />
          </div>
          <h1 className="text-2xl font-bold">CleanTrack Ops</h1>
          <p className="text-gray-400 text-sm mt-1">Sign in to continue</p>
        </div>

        <div className="p-8 space-y-6">
          <h2 className="text-lg font-semibold text-gray-900">Sign in</h2>

          {isMicrosoftAuthConfigured ? (
            <button
              type="button"
              onClick={() => signInWithMicrosoft()}
              className="w-full bg-[#238636] hover:bg-[#2ea043] text-white py-3 rounded-xl font-bold transition-all shadow-lg active:scale-[0.98] flex items-center justify-center gap-2"
            >
              Sign in with Microsoft
            </button>
          ) : (
            <div className="rounded-xl bg-amber-50 border border-amber-200 p-4 text-sm text-amber-800">
              Microsoft auth not configured. Set VITE_AZURE_AD_CLIENT_ID and VITE_AZURE_AD_TENANT_ID to enable sign-in.
            </div>
          )}

          {DEV_SHOW_LEGACY_LOGIN && (
            <button
              type="button"
              onClick={() => setShowLegacy(true)}
              className="w-full text-[12px] text-gray-500 hover:text-gray-700"
            >
              Use email / password (legacy)
            </button>
          )}
        </div>
      </div>
    </div>
  );
};

export default SignInScreen;
