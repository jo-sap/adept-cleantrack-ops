
import { createClient } from '@supabase/supabase-js';

// Access environment variables. 
// We provide default values that satisfy the URL/Key format requirements of the Supabase SDK 
// to prevent "supabaseUrl is required" or "supabaseKey is required" errors during module initialization.
// Real functionality requires these environment variables to be correctly set in the deployment environment.
const supabaseUrl = process.env.SUPABASE_URL || 'https://placeholder.supabase.co';
const supabaseAnonKey = process.env.SUPABASE_ANON_KEY || 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.placeholder';

export const supabase = createClient(supabaseUrl, supabaseAnonKey);
