-- Run this in the Supabase SQL Editor to create the required tables

-- 1. Table for User Settings (API Keys)
CREATE TABLE IF NOT EXISTS public.user_settings (
  user_id UUID REFERENCES auth.users(id) PRIMARY KEY,
  ollama_api_key TEXT,
  ollama_host_url TEXT DEFAULT 'https://ollama.com',
  created_at TIMESTAMP WITH TIME ZONE DEFAULT TIMEZONE('utc'::text, NOW()) NOT NULL,
  updated_at TIMESTAMP WITH TIME ZONE DEFAULT TIMEZONE('utc'::text, NOW()) NOT NULL
);

-- Note: In a production app you'd want Row Level Security (RLS) on user_settings:
ALTER TABLE public.user_settings ENABLE ROW LEVEL SECURITY;
CREATE POLICY "Users can manage their own settings" 
  ON public.user_settings FOR ALL 
  USING (auth.uid() = user_id);

-- 2. Table for Review History
CREATE TABLE IF NOT EXISTS public.review_history (
  id UUID DEFAULT gen_random_uuid() PRIMARY KEY,
  user_id UUID REFERENCES auth.users(id) NOT NULL,
  document_name TEXT NOT NULL,
  report_filename TEXT NOT NULL,
  created_at TIMESTAMP WITH TIME ZONE DEFAULT TIMEZONE('utc'::text, NOW()) NOT NULL
);

-- Basic RLS for history
ALTER TABLE public.review_history ENABLE ROW LEVEL SECURITY;
CREATE POLICY "Users can view their own history" 
  ON public.review_history FOR SELECT 
  USING (auth.uid() = user_id);
  
CREATE POLICY "Users can insert their own history" 
  ON public.review_history FOR INSERT 
  WITH CHECK (auth.uid() = user_id);
