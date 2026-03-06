import React, { createContext, useContext, useState, useEffect, useCallback, FC, ReactNode } from "react";
import { DEV_BYPASS_LOGIN } from "../config/authFlags";
import { isMicrosoftAuthConfigured } from "../config/authFlags";
import { msalInstance } from "../lib/msal";
import { GRAPH_SCOPES } from "../src/auth/graphScopes";
import { getGraphAccessToken } from "../lib/graph";
import { getCleanTrackUserByEmail } from "../repositories/usersRepo";

export type AuthStatus = "loading" | "unauthenticated" | "authorizing" | "authenticated" | "unauthorized";

export interface AuthUser {
  id: string;
  name: string;
  email: string;
  role?: string;
  permissionScope?: string | null;
}

interface AppAuthContextValue {
  authStatus: AuthStatus;
  user: AuthUser | null;
  isMicrosoftAuthConfigured: boolean;
  signInWithMicrosoft: () => Promise<void>;
  signOut: () => Promise<void>;
}

const AppAuthContext = createContext<AppAuthContextValue | undefined>(undefined);

const DEV_USER: AuthUser = { id: "dev-user", name: "Dev User", email: "dev@local", role: "Admin" };

async function authorizeWithSharePoint(
  email: string,
  setAuthStatus: (s: AuthStatus) => void,
  setUser: (u: AuthUser | null) => void
): Promise<void> {
  setAuthStatus("authorizing");
  const token = await getGraphAccessToken();
  if (!token) {
    setAuthStatus("unauthorized");
    setUser(null);
    return;
  }
  try {
    const ctUser = await getCleanTrackUserByEmail(token, email);
    if (!ctUser || !ctUser.active) {
      setAuthStatus("unauthorized");
      setUser(null);
      return;
    }
    setUser({
      id: email,
      name: ctUser.fullName || email,
      email: ctUser.email,
      role: ctUser.role,
      permissionScope: ctUser.permissionScope ?? null,
    });
    setAuthStatus("authenticated");
  } catch {
    setAuthStatus("unauthorized");
    setUser(null);
  }
}

export const AppAuthProvider: FC<{ children: ReactNode }> = ({ children }) => {
  const [authStatus, setAuthStatus] = useState<AuthStatus>("loading");
  const [user, setUser] = useState<AuthUser | null>(null);

  const signOut = useCallback(async () => {
    if (DEV_BYPASS_LOGIN) return;
    try {
      await msalInstance.logoutPopup();
    } finally {
      setUser(null);
      setAuthStatus("unauthenticated");
    }
  }, []);

  const syncMsalAccounts = useCallback(() => {
    if (DEV_BYPASS_LOGIN) {
      setAuthStatus("authenticated");
      setUser(DEV_USER);
      return;
    }
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length === 0) {
      setUser(null);
      setAuthStatus("unauthenticated");
      return;
    }
    const a = accounts[0];
    const email = a.username ?? "";
    setUser({
      id: a.homeAccountId,
      name: a.name ?? email,
      email,
    });
    authorizeWithSharePoint(email, setAuthStatus, setUser);
  }, []);

  useEffect(() => {
    if (DEV_BYPASS_LOGIN) {
      setAuthStatus("authenticated");
      setUser(DEV_USER);
      return;
    }
    msalInstance.initialize().then(() => {
      syncMsalAccounts();
    }).catch(() => {
      setAuthStatus("unauthenticated");
      setUser(null);
    });
  }, [syncMsalAccounts]);

  const signInWithMicrosoft = useCallback(async () => {
    if (!isMicrosoftAuthConfigured) return;
    try {
      await msalInstance.loginPopup({ scopes: GRAPH_SCOPES });
      syncMsalAccounts();
    } catch {
      syncMsalAccounts();
    }
  }, [syncMsalAccounts]);

  const value: AppAuthContextValue = {
    authStatus,
    user,
    isMicrosoftAuthConfigured,
    signInWithMicrosoft,
    signOut,
  };

  return (
    <AppAuthContext.Provider value={value}>
      {children}
    </AppAuthContext.Provider>
  );
};

export const useAppAuth = () => {
  const ctx = useContext(AppAuthContext);
  if (!ctx) throw new Error("useAppAuth must be used within AppAuthProvider");
  return ctx;
};
