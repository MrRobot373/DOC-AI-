-- ============================================================
-- DOC-AI Supabase Schema — run this in the Supabase SQL Editor
-- (safe to re-run: all statements use IF NOT EXISTS / OR REPLACE)
-- ============================================================

-- ── 1. User Settings ────────────────────────────────────────
CREATE TABLE IF NOT EXISTS public.user_settings (
  user_id          UUID REFERENCES auth.users(id) ON DELETE CASCADE PRIMARY KEY,
  ollama_api_key   TEXT,
  ollama_host_url  TEXT DEFAULT 'http://localhost:11434',
  ollama_runtime   TEXT DEFAULT 'local',          -- 'local' | 'cloud'
  selected_model   TEXT,
  vision_model     TEXT,
  created_at       TIMESTAMPTZ DEFAULT NOW() NOT NULL,
  updated_at       TIMESTAMPTZ DEFAULT NOW() NOT NULL
);

ALTER TABLE public.user_settings ENABLE ROW LEVEL SECURITY;

DROP POLICY IF EXISTS "Users manage own settings" ON public.user_settings;
CREATE POLICY "Users manage own settings"
  ON public.user_settings FOR ALL
  USING (auth.uid() = user_id)
  WITH CHECK (auth.uid() = user_id);


-- ── 2. Reviews ──────────────────────────────────────────────
-- One row per review; findings stored as JSONB so status
-- updates are atomic (replaces the whole-file blob approach).
CREATE TABLE IF NOT EXISTS public.reviews (
  id               TEXT PRIMARY KEY,              -- 8-char hex review_id
  user_id          UUID REFERENCES auth.users(id) ON DELETE CASCADE NOT NULL,
  status           TEXT DEFAULT 'starting',       -- starting|parsing|reviewing|done|error
  progress         INT DEFAULT 0,
  message          TEXT,
  review_mode      TEXT DEFAULT 'pro',
  document_info    JSONB,
  findings         JSONB,
  summary          JSONB,
  report_filename  TEXT,
  pdf_filename     TEXT,
  engine_status    JSONB,
  error            TEXT,
  original_filepath TEXT,
  created_at       TIMESTAMPTZ DEFAULT NOW() NOT NULL,
  updated_at       TIMESTAMPTZ DEFAULT NOW() NOT NULL
);

ALTER TABLE public.reviews ENABLE ROW LEVEL SECURITY;

DROP POLICY IF EXISTS "Users manage own reviews" ON public.reviews;
CREATE POLICY "Users manage own reviews"
  ON public.reviews FOR ALL
  USING (auth.uid() = user_id)
  WITH CHECK (auth.uid() = user_id);

-- Auto-update updated_at
CREATE OR REPLACE FUNCTION public.set_updated_at()
RETURNS TRIGGER LANGUAGE plpgsql AS $$
BEGIN NEW.updated_at = NOW(); RETURN NEW; END;
$$;

DROP TRIGGER IF EXISTS reviews_updated_at ON public.reviews;
CREATE TRIGGER reviews_updated_at
  BEFORE UPDATE ON public.reviews
  FOR EACH ROW EXECUTE FUNCTION public.set_updated_at();

DROP TRIGGER IF EXISTS user_settings_updated_at ON public.user_settings;
CREATE TRIGGER user_settings_updated_at
  BEFORE UPDATE ON public.user_settings
  FOR EACH ROW EXECUTE FUNCTION public.set_updated_at();


-- ── 3. Review History (lightweight list for the dashboard) ──
CREATE TABLE IF NOT EXISTS public.review_history (
  id               UUID DEFAULT gen_random_uuid() PRIMARY KEY,
  user_id          UUID REFERENCES auth.users(id) ON DELETE CASCADE NOT NULL,
  document_name    TEXT NOT NULL,
  report_filename  TEXT NOT NULL,
  review_id        TEXT REFERENCES public.reviews(id) ON DELETE SET NULL,
  created_at       TIMESTAMPTZ DEFAULT NOW() NOT NULL
);

ALTER TABLE public.review_history ENABLE ROW LEVEL SECURITY;

DROP POLICY IF EXISTS "Users view own history" ON public.review_history;
CREATE POLICY "Users view own history"
  ON public.review_history FOR SELECT
  USING (auth.uid() = user_id);

DROP POLICY IF EXISTS "Users insert own history" ON public.review_history;
CREATE POLICY "Users insert own history"
  ON public.review_history FOR INSERT
  WITH CHECK (auth.uid() = user_id);


-- ── 4. Feedback ─────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS public.feedback (
  id         UUID DEFAULT gen_random_uuid() PRIMARY KEY,
  user_email TEXT,
  type       TEXT,
  message    TEXT,
  image_url  TEXT,
  created_at TIMESTAMPTZ DEFAULT NOW() NOT NULL
);
-- Feedback is admin-read-only; backend uses the service role key to insert.
