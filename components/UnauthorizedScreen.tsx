import React from "react";
import { ShieldX } from "lucide-react";

interface UnauthorizedScreenProps {
  onSignOut: () => void;
}

const UnauthorizedScreen: React.FC<UnauthorizedScreenProps> = ({ onSignOut }) => (
  <div className="min-h-screen flex items-center justify-center bg-gray-50 p-6">
    <div className="w-full max-w-md bg-white rounded-3xl shadow-xl border border-gray-100 overflow-hidden text-center">
      <div className="p-8">
        <div className="bg-red-100 w-12 h-12 rounded-xl flex items-center justify-center mx-auto mb-4">
          <ShieldX className="text-red-600" size={24} />
        </div>
        <h1 className="text-xl font-bold text-gray-900">Access denied</h1>
        <p className="text-gray-600 text-sm mt-2">
          Your account is not authorised to access CleanTrack Ops. Please contact your administrator.
        </p>
        <button
          type="button"
          onClick={onSignOut}
          className="mt-6 px-6 py-2.5 bg-gray-900 text-white text-sm font-medium rounded-xl hover:bg-black transition-colors"
        >
          Sign out
        </button>
      </div>
    </div>
  </div>
);

export default UnauthorizedScreen;
