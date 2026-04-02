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

export const isSupabaseConfigured = !!import.meta.env.VITE_SUPABASE_URL && 
                                    !!import.meta.env.VITE_SUPABASE_ANON_KEY && 
                                    !supabaseUrl.includes('placeholder');

export const supabase = createClient(supabaseUrl, supabaseAnonKey);

export const testSupabaseConnection = async () => {
  if (!isSupabaseConfigured) return { success: false, message: 'Supabase credentials missing' };
  try {
    const { data, error } = await supabase.from('products').select('count').limit(1);
    if (error) throw error;
    return { success: true, message: 'Connected successfully' };
  } catch (error: any) {
    console.error('Supabase connection test failed:', error);
    return { success: false, message: error.message || 'Connection failed' };
  }
};
