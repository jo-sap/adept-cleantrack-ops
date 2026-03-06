
import { PublicClientApplication } from "@azure/msal-browser";

// Env (Vite): VITE_AZURE_AD_CLIENT_ID, VITE_AZURE_AD_TENANT_ID. Fallback for dev.
const env = typeof import.meta !== "undefined" && import.meta.env ? import.meta.env : {} as Record<string, string | undefined>;
const tenantId = env.VITE_AZURE_AD_TENANT_ID || "3aaa69e7-cd58-4195-83d5-f631ad9eea58";
const clientId = env.VITE_AZURE_AD_CLIENT_ID || "84deae7d-9f8c-4a27-9df0-e364ecc3dfcd";

export const msalConfig = {
  auth: {
    clientId: clientId,
    authority: `https://login.microsoftonline.com/${tenantId}`,
    redirectUri: window.location.origin,
  },
  cache: {
    cacheLocation: "localStorage",
    storeAuthStateInCookie: false,
  },
};

export const msalInstance = new PublicClientApplication(msalConfig);
