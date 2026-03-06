/**
 * Auth feature flags and Microsoft auth config detection.
 *
 * Env keys (Vite — use in .env or .env.local):
 *   VITE_DEV_BYPASS_LOGIN     — "true" to skip login and enter app (dev only)
 *   VITE_DEV_SHOW_LEGACY_LOGIN — "true" to show email/password login option
 *   VITE_AZURE_AD_CLIENT_ID   — Azure AD app (client) ID for Microsoft login
 *   VITE_AZURE_AD_TENANT_ID   — Azure AD tenant ID (or use "common")
 *
 * If VITE_AZURE_AD_* are not set, the app shows "Microsoft auth not configured"
 * but does not crash; DEV_BYPASS_LOGIN still works.
 */
const env = typeof import.meta !== "undefined" && import.meta.env ? import.meta.env : {} as Record<string, string | undefined>;

export const DEV_BYPASS_LOGIN = env.VITE_DEV_BYPASS_LOGIN === "true";
export const DEV_SHOW_LEGACY_LOGIN = env.VITE_DEV_SHOW_LEGACY_LOGIN === "true";

export const isMicrosoftAuthConfigured =
  !!(env.VITE_AZURE_AD_CLIENT_ID && env.VITE_AZURE_AD_TENANT_ID);
