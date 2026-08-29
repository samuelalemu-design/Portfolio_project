import { createClient } from '@supabase/supabase-js';

const supabaseUrl = (typeof process !== 'undefined' && process.env && process.env.NEXT_PUBLIC_SUPABASE_URL) 
  ? process.env.NEXT_PUBLIC_SUPABASE_URL 
  : 'https://yudcsarufgftnncyzuat.supabase.co';

const supabaseAnonKey = (typeof process !== 'undefined' && process.env && process.env.NEXT_PUBLIC_SUPABASE_ANON_KEY)
  ? process.env.NEXT_PUBLIC_SUPABASE_ANON_KEY
  : 'sb_publishable_J6d3i52t2h3Pm2L5cSneZw_MInpF0Y1';

export const supabase = createClient(supabaseUrl, supabaseAnonKey);

if (typeof window !== 'undefined') {
  window.supabaseClient = supabase;
  if (!window.supabase || typeof window.supabase.from !== 'function') {
    window.supabase = supabase;
  }
}
