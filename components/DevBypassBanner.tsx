import React from "react";
import { DEV_BYPASS_LOGIN } from "../config/authFlags";

export const DevBypassBanner: React.FC = () => {
  if (!DEV_BYPASS_LOGIN) return null;
  return (
    <span className="text-[11px] font-medium text-amber-700 bg-amber-100 px-2 py-1 rounded">
      DEV MODE bypass enabled
    </span>
  );
};

export default DevBypassBanner;
