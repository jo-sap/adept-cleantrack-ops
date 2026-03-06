import { msalInstance } from "./msal";
import { InteractionRequiredAuthError } from "@azure/msal-browser";
import { GRAPH_SCOPES } from "../src/auth/graphScopes";

/** Decode JWT payload without verification (for dev diagnostics). Returns aud and scp. */
export function decodeToken(accessToken: string): { aud?: string; scp?: string } {
  try {
    const parts = accessToken.split(".");
    if (parts.length < 2) return {};
    const payload = parts[1];
    const base64 = payload.replace(/-/g, "+").replace(/_/g, "/");
    const padded = base64 + "==".slice(0, (4 - (base64.length % 4)) % 4);
    const json = atob(padded);
    const decoded = JSON.parse(json) as Record<string, unknown>;
    return {
      aud: typeof decoded.aud === "string" ? decoded.aud : undefined,
      scp: typeof decoded.scp === "string" ? decoded.scp : undefined,
    };
  } catch {
    return {};
  }
}

const IS_DEV =
  typeof import.meta !== "undefined" && import.meta.env && (import.meta.env as { DEV?: boolean }).DEV === true;

/** Returns a Graph access token or null if not signed in / not configured. Does not throw. Returns ONLY accessToken (never idToken). */
export async function getGraphAccessToken(): Promise<string | null> {
  const account = msalInstance.getAllAccounts()[0];
  if (!account) return null;
  try {
    const response = await msalInstance.acquireTokenSilent({ account, scopes: GRAPH_SCOPES });
    if (IS_DEV && response.accessToken) {
      const { aud, scp } = decodeToken(response.accessToken);
      console.log("GRAPH token aud:", aud, "scp:", scp);
    }
    console.log("GRAPH token scopes:", response.scopes);
    return response.accessToken;
  } catch (error) {
    if (error instanceof InteractionRequiredAuthError) {
      try {
        const response = await msalInstance.acquireTokenPopup({ scopes: GRAPH_SCOPES });
        if (IS_DEV && response.accessToken) {
          const { aud, scp } = decodeToken(response.accessToken);
          console.log("GRAPH token aud:", aud, "scp:", scp);
        }
        console.log("GRAPH token scopes:", response.scopes);
        return response.accessToken;
      } catch {
        return null;
      }
    }
    return null;
  }
}

export async function acquireGraphToken(scopes: string[]): Promise<string> {
  const account = msalInstance.getAllAccounts()[0];
  if (!account) {
    throw new Error("Not signed in");
  }
  try {
    const response = await msalInstance.acquireTokenSilent({ account, scopes });
    return response.accessToken;
  } catch (error) {
    if (error instanceof InteractionRequiredAuthError) {
      const response = await msalInstance.acquireTokenPopup({ scopes });
      return response.accessToken;
    }
    throw error;
  }
}

export async function graphGet<T = unknown>(
  url: string,
  scopes: string[]
): Promise<{ ok: boolean; status: number; data?: T; text?: string }> {
  try {
    const token = await acquireGraphToken(scopes);
    const response = await fetch(url, {
      headers: { Authorization: `Bearer ${token}` },
    });
    const text = await response.text();
    try {
      const data = JSON.parse(text) as T;
      return { ok: response.ok, status: response.status, data, text };
    } catch {
      return { ok: response.ok, status: response.status, text };
    }
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : "Request failed";
    return { ok: false, status: 0, text: message };
  }
}
