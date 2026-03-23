
import React, { createContext, useContext, FC, ReactNode, useMemo } from "react";
import type { Role, Profile } from "../types";
import { useAppAuth } from "./AppAuthContext";

interface AuthContextType {
  role: Role | null;
  profile: Profile | null;
  isAdmin: boolean;
  isManager: boolean;
  loading: boolean;
  logout: () => Promise<void>;
  setRole: (role: Role) => void;
}

const AuthContext = createContext<AuthContextType | undefined>(undefined);

export const RoleProvider: FC<{ children: ReactNode }> = ({ children }) => {
  const { authStatus, user, signOut } = useAppAuth();
  const loading = authStatus === "loading" || authStatus === "authorizing";

  const role: Role | null = useMemo(() => {
    const r = String(user?.role ?? "").trim().toLowerCase();
    if (r === "admin") return "admin";
    if (r === "manager") return "manager";
    return null;
  }, [user?.role]);

  const profile: Profile | null = useMemo(() => {
    if (!user) return null;
    if (!role) return null;
    return {
      id: user.id,
      full_name: user.name,
      role,
    };
  }, [user, role]);

  const logout = async () => {
    await signOut();
  };

  // Role is sourced from CleanTrack Users (SharePoint). Mutating it locally is not supported.
  const setRole = () => {};

  const isAdmin = role === "admin";
  const isManager = role === "manager";

  return (
    <AuthContext.Provider value={{ 
      role, 
      profile, 
      isAdmin, 
      isManager, 
      loading,
      logout,
      setRole
    }}>
      {children}
    </AuthContext.Provider>
  );
};

export const useRole = () => {
  const context = useContext(AuthContext);
  if (!context) throw new Error('useRole must be used within a RoleProvider');
  return context;
};