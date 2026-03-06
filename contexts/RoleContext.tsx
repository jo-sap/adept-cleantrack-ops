
import React, { createContext, useContext, useState, useEffect, FC, ReactNode } from 'react';
import { Role, Profile } from '../types';
import { supabase } from '../lib/supabase';
import { Session } from '@supabase/supabase-js';

interface AuthContextType {
  role: Role | null;
  profile: Profile | null;
  session: Session | null;
  isAdmin: boolean;
  isManager: boolean;
  loading: boolean;
  logout: () => Promise<void>;
  setRole: (role: Role) => void;
}

const AuthContext = createContext<AuthContextType | undefined>(undefined);

export const RoleProvider: FC<{ children: ReactNode }> = ({ children }) => {
  const [session, setSession] = useState<Session | null>(null);
  const [profile, setProfile] = useState<Profile | null>(null);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    supabase.auth.getSession().then(({ data: { session } }) => {
      setSession(session);
      if (session) fetchProfile(session.user.id);
      else setLoading(false);
    });

    const { data: { subscription } } = supabase.auth.onAuthStateChange((_event, session) => {
      setSession(session);
      if (session) fetchProfile(session.user.id);
      else {
        setProfile(null);
        setLoading(false);
      }
    });

    return () => subscription.unsubscribe();
  }, []);

  const fetchProfile = async (uid: string) => {
    const { data, error } = await supabase
      .from('profiles')
      .select('*')
      .eq('id', uid)
      .single();

    if (error && error.code === 'PGRST116') {
      const { data: user } = await supabase.auth.getUser();
      const newProfile = {
        id: uid,
        full_name: user.user?.user_metadata.full_name || 'New User',
        role: 'manager' as Role
      };
      await supabase.from('profiles').insert(newProfile);
      setProfile(newProfile);
    } else {
      setProfile(data);
    }
    setLoading(false);
  };

  const logout = async () => {
    await supabase.auth.signOut();
  };

  const setRole = (role: Role) => {
    if (profile) setProfile({ ...profile, role });
  };

  const isAdmin = profile?.role === 'admin';
  const isManager = profile?.role === 'manager';

  return (
    <AuthContext.Provider value={{ 
      role: profile?.role || null, 
      profile, 
      session, 
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