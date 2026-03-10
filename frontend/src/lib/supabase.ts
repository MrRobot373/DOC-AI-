import { createClient } from '@supabase/supabase-js'

const supabaseUrl = import.meta.env.VITE_SUPABASE_URL || ''
const supabaseAnonKey = import.meta.env.VITE_SUPABASE_ANON_KEY || ''

if (!supabaseUrl || !supabaseAnonKey) {
    console.error('Supabase environment variables are missing! Authentication and Database features will not work.')
}

// Fallback to dummy values to prevent app crash if vars are not set during Vercel build
export const supabase = createClient(
    supabaseUrl || 'https://placeholder.supabase.co',
    supabaseAnonKey || 'placeholder_key'
)
