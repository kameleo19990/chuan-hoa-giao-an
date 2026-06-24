-- ============================================================
-- CHẠY FILE NÀY TRONG: Supabase Dashboard → SQL Editor
-- Xóa bảng cũ (nếu có) rồi tạo lại sạch
-- ============================================================

DROP TABLE IF EXISTS public.usage_logs CASCADE;
DROP TABLE IF EXISTS public.profiles   CASCADE;

-- ── Bảng profiles ─────────────────────────────────────────────
CREATE TABLE public.profiles (
  id          UUID    REFERENCES auth.users(id) ON DELETE CASCADE PRIMARY KEY,
  email       TEXT,
  used_quota  INT     NOT NULL DEFAULT 0,
  is_pro      BOOLEAN NOT NULL DEFAULT FALSE,
  quota_month TEXT    DEFAULT TO_CHAR(NOW(), 'YYYY-MM'),
  created_at  TIMESTAMPTZ DEFAULT NOW()
);

-- Bật Row Level Security
ALTER TABLE public.profiles ENABLE ROW LEVEL SECURITY;

-- User chỉ đọc/sửa profile của chính mình
CREATE POLICY "User đọc profile của mình"
  ON public.profiles FOR SELECT USING (auth.uid() = id);

CREATE POLICY "User cập nhật profile của mình"
  ON public.profiles FOR UPDATE USING (auth.uid() = id);

-- Backend dùng service key → bypass RLS hoàn toàn
CREATE POLICY "Service role toàn quyền"
  ON public.profiles FOR ALL USING (true) WITH CHECK (true);

-- ── Trigger: tự tạo profile khi user đăng ký ─────────────────
CREATE OR REPLACE FUNCTION public.handle_new_user()
RETURNS TRIGGER AS $$
BEGIN
  INSERT INTO public.profiles (id, email)
  VALUES (NEW.id, NEW.email)
  ON CONFLICT (id) DO NOTHING;
  RETURN NEW;
END;
$$ LANGUAGE plpgsql SECURITY DEFINER;

DROP TRIGGER IF EXISTS on_auth_user_created ON auth.users;
CREATE TRIGGER on_auth_user_created
  AFTER INSERT ON auth.users
  FOR EACH ROW EXECUTE FUNCTION public.handle_new_user();
