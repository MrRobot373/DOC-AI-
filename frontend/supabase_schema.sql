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


-- ── 5. Admin users ──────────────────────────────────────────
CREATE TABLE IF NOT EXISTS public.admin_users (
  user_id UUID REFERENCES auth.users(id) ON DELETE CASCADE PRIMARY KEY,
  created_at TIMESTAMPTZ DEFAULT NOW() NOT NULL
);
-- Seed the platform admin (runs once the user has signed up):
INSERT INTO public.admin_users (user_id)
SELECT id FROM auth.users WHERE email = 'yash.badgujar@getmysolutions.in'
ON CONFLICT DO NOTHING;

ALTER TABLE public.admin_users ENABLE ROW LEVEL SECURITY;
DROP POLICY IF EXISTS "Admins read admin list" ON public.admin_users;
CREATE POLICY "Admins read admin list" ON public.admin_users FOR SELECT
  USING ((SELECT COUNT(*) FROM public.admin_users a WHERE a.user_id = auth.uid()) > 0);


-- ── 6. Shared LLM key pool (managed by admin) ───────────────
CREATE TABLE IF NOT EXISTS public.llm_pool_keys (
  id         UUID DEFAULT gen_random_uuid() PRIMARY KEY,
  label      TEXT,
  provider   TEXT NOT NULL,            -- 'ollama_cloud' | 'freellmapi' | 'openai_compat'
  host_url   TEXT NOT NULL,
  api_key    TEXT NOT NULL,
  model_hint TEXT,                     -- default model name for this key
  vision_model_hint TEXT,
  priority   INT DEFAULT 0,            -- lower = tried first
  active     BOOLEAN DEFAULT true,
  created_at TIMESTAMPTZ DEFAULT NOW() NOT NULL
);
ALTER TABLE public.llm_pool_keys ENABLE ROW LEVEL SECURITY;
-- Only admins can see/manage pool keys via the client; the backend reads them
-- with the service-role key (bypasses RLS) when running a pooled review.
DROP POLICY IF EXISTS "Admins manage pool keys" ON public.llm_pool_keys;
CREATE POLICY "Admins manage pool keys" ON public.llm_pool_keys FOR ALL
  USING ((SELECT COUNT(*) FROM public.admin_users a WHERE a.user_id = auth.uid()) > 0)
  WITH CHECK ((SELECT COUNT(*) FROM public.admin_users a WHERE a.user_id = auth.uid()) > 0);


-- ── 7. user_settings: opt into the shared pool ("Auto" mode) ─
ALTER TABLE public.user_settings ADD COLUMN IF NOT EXISTS use_pool BOOLEAN DEFAULT false;
ALTER TABLE public.user_settings ADD COLUMN IF NOT EXISTS glossary_json JSONB;
ALTER TABLE public.user_settings ADD COLUMN IF NOT EXISTS notify_email BOOLEAN DEFAULT false;


-- ── 8. Usage / audit log ────────────────────────────────────
CREATE TABLE IF NOT EXISTS public.audit_log (
  id         UUID DEFAULT gen_random_uuid() PRIMARY KEY,
  user_id    UUID,
  user_email TEXT,
  action     TEXT,                     -- review_start | review_done | download | fix_applied
  review_id  TEXT,
  metadata   JSONB,
  created_at TIMESTAMPTZ DEFAULT NOW() NOT NULL
);
ALTER TABLE public.audit_log ENABLE ROW LEVEL SECURITY;
DROP POLICY IF EXISTS "Own or admin reads audit" ON public.audit_log;
CREATE POLICY "Own or admin reads audit" ON public.audit_log FOR SELECT
  USING (auth.uid() = user_id
    OR (SELECT COUNT(*) FROM public.admin_users a WHERE a.user_id = auth.uid()) > 0);
