# Migrasi ke Supabase + Vercel

Aplikasi tetap **Vite SPA**. Selama env Supabase kosong, app berjalan seperti
biasa memakai **localStorage** (tanpa perubahan perilaku). Begitu env terisi,
data lembur otomatis pindah ke **Supabase**.

---

## 1. Supabase

1. Buat project di https://supabase.com → **New project**.
2. Buka **SQL Editor** → New query → tempel isi [`supabase/schema.sql`](supabase/schema.sql) → **Run**.
3. Query baru lagi → tempel [`supabase/seed.sql`](supabase/seed.sql) → **Run** (isi roster karyawan WKI).
4. **Project Settings → API**, salin:
   - `Project URL`  → `VITE_SUPABASE_URL`
   - `anon public`  → `VITE_SUPABASE_ANON_KEY`
5. Isi ke file `.env` (lokal):

   ```
   VITE_SUPABASE_URL=https://xxxx.supabase.co
   VITE_SUPABASE_ANON_KEY=eyJhbGciOi...
   ```

6. Restart dev server (`npm run dev`). Cek console: kalau tidak ada error Supabase,
   entri baru akan tersimpan ke tabel `public.lembur`.

> **RLS**: kebijakan bawaan di `schema.sql` masih **permisif** (cocok untuk mode
> roster tanpa password). Perketat saat sudah pakai Supabase Auth (lihat §3).

---

## 2. Vercel

1. Push repo ke GitHub.
2. https://vercel.com → **Add New → Project** → import repo.
3. Framework autodetect **Vite** (sudah ada [`vercel.json`](vercel.json)).
4. **Environment Variables** → tambahkan `VITE_SUPABASE_URL` & `VITE_SUPABASE_ANON_KEY`
   (nilai sama dengan `.env`).
5. **Deploy**. SPA rewrite sudah diatur agar semua route jatuh ke `index.html`.

CLI alternatif:

```bash
npm i -g vercel
vercel            # preview
vercel --prod     # production
```

---

## 3. (Opsional) Auth email + password

Saat ini login = pilih dari roster (tanpa password). Untuk auth asli:

1. Supabase **Authentication → Providers → Email** (enable).
2. Buat akun untuk tiap karyawan/atasan/admin (atau pakai invite).
3. Hubungkan `auth.users.id` ke `public.profil.auth_uid`.
4. Ganti login roster dengan `supabase.auth.signInWithPassword(...)`.
5. Perketat RLS, contoh entri hanya bisa dilihat pemiliknya + atasan/admin:

   ```sql
   -- contoh; sesuaikan dgn desain profil/peran
   drop policy if exists lembur_read on public.lembur;
   create policy lembur_read on public.lembur for select
     using ( auth.role() = 'authenticated' );
   ```

Helper `supabase.auth` sudah tersedia di [`src/lib/supabase.js`](src/lib/supabase.js).

---

## Arsitektur data

| Lapisan | File | Peran |
|---|---|---|
| Client | `src/lib/supabase.js` | Inisialisasi client; `isSupabaseEnabled` |
| Repo | `src/lib/repo.js` | 1 API async; auto pilih Supabase / localStorage + mapping kolom |
| UI | `src/LemburApp.jsx` | Mutator write-through ke repo saat Supabase aktif |

Bila `.env` kosong → semua lewat localStorage (fallback), jadi aman untuk
dikembangkan bertahap.
