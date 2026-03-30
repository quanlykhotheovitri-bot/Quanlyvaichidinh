import { createClient } from '@supabase/supabase-js';

const getValidUrl = (url: any): string => {
  const placeholder = 'https://placeholder-project.supabase.co';
  if (typeof url !== 'string') return placeholder;
  try {
    const parsed = new URL(url);
    if (parsed.protocol === 'http:' || parsed.protocol === 'https:') {
      return url;
    }
  } catch {
    // Not a valid URL
  }
  return placeholder;
};

const supabaseUrl = getValidUrl(import.meta.env.VITE_SUPABASE_URL);
const supabaseAnonKey = import.meta.env.VITE_SUPABASE_ANON_KEY || 'placeholder-key';

if (!import.meta.env.VITE_SUPABASE_URL || !import.meta.env.VITE_SUPABASE_ANON_KEY || supabaseUrl.includes('placeholder')) {
  console.warn('Supabase URL or Anon Key is missing or invalid. Please set VITE_SUPABASE_URL and VITE_SUPABASE_ANON_KEY in your environment variables.');
}

export const supabase = createClient(supabaseUrl, supabaseAnonKey);
