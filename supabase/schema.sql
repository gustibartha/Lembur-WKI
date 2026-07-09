-- ============================================================
-- Lembur.id — Supabase schema
-- Jalankan di Supabase Studio → SQL Editor → New query → Run
-- ============================================================

-- Extensions
create extension if not exists "pgcrypto";

-- ------------------------------------------------------------
-- 1. KARYAWAN (master data Office WKI)
-- ------------------------------------------------------------
create table if not exists public.karyawan (
  id            uuid primary key default gen_random_uuid(),
  no            int,
  nama          text not null,
  nomor_induk   text unique not null,
  ttl           text,
  tahun_masuk   text,
  jabatan       text,
  kategori      text not null default 'SPL',      -- SPL | WaterIntake | Pengemudi
  boleh_lembur  boolean not null default true,
  gaji_pokok    numeric default 5729876,
  created_at    timestamptz not null default now()
);

-- ------------------------------------------------------------
-- 2. PROFIL (peran aplikasi; opsional dihubungkan ke auth.users)
-- ------------------------------------------------------------
-- peran: Pekerja | Atasan | Admin
create table if not exists public.profil (
  id            uuid primary key default gen_random_uuid(),
  auth_uid      uuid references auth.users(id) on delete cascade,
  nama          text not null,
  email         text,
  peran         text not null default 'Pekerja',
  karyawan_id   uuid references public.karyawan(id) on delete set null,
  created_at    timestamptz not null default now()
);

-- ------------------------------------------------------------
-- 3. LEMBUR (entri lembur / piket)
-- ------------------------------------------------------------
create table if not exists public.lembur (
  id                uuid primary key default gen_random_uuid(),
  karyawan_id       uuid references public.karyawan(id) on delete set null,
  user_nama         text not null,
  kategori          text not null default 'SPL',    -- SPL | WaterIntake | Pengemudi
  tipe              text not null default 'Lembur',  -- Lembur | Piket
  bagian            text,
  pekerjaan         text,
  tujuan            text,
  tanggal_mulai     date,
  jam_mulai         text,
  tanggal_selesai   date,
  jam_selesai       text,
  jenis             text default 'Hari Kerja',       -- Hari Kerja | Hari Libur
  durasi_jam        numeric default 0,
  estimasi_gaji     numeric default 0,
  hasil             text,
  keterangan        text,
  penanggung_jawab  text,
  status            text not null default 'Draft',   -- Draft | Terkirim | Disetujui
  evidence          jsonb,                           -- { name, type, size, dataUrl }
  ttd_pekerja       text,                            -- data URL PNG
  ttd_atasan        text,
  atasan_nama       text,
  catatan_atasan    text,
  sent_at           timestamptz,
  approved_at       timestamptz,
  created_at        timestamptz not null default now(),
  updated_at        timestamptz not null default now()
);

create index if not exists lembur_karyawan_idx on public.lembur (karyawan_id);
create index if not exists lembur_status_idx   on public.lembur (status);
create index if not exists lembur_tanggal_idx  on public.lembur (tanggal_mulai);

-- Auto-update updated_at
create or replace function public.touch_updated_at()
returns trigger language plpgsql as $$
begin
  new.updated_at = now();
  return new;
end $$;

drop trigger if exists lembur_touch on public.lembur;
create trigger lembur_touch before update on public.lembur
  for each row execute function public.touch_updated_at();

-- ------------------------------------------------------------
-- 4. PENGATURAN (single-row konfigurasi perusahaan)
-- ------------------------------------------------------------
create table if not exists public.pengaturan (
  id          int primary key default 1,
  data        jsonb not null default '{}'::jsonb,
  updated_at  timestamptz not null default now(),
  constraint pengaturan_singleton check (id = 1)
);

-- ============================================================
-- Row Level Security
-- ============================================================
-- Catatan: kebijakan di bawah PERMISIF (untuk mode roster/anon).
-- Perketat saat sudah pakai Supabase Auth email+password.
alter table public.karyawan   enable row level security;
alter table public.profil     enable row level security;
alter table public.lembur     enable row level security;
alter table public.pengaturan enable row level security;

-- Baca untuk semua (anon + authenticated)
drop policy if exists karyawan_read on public.karyawan;
create policy karyawan_read on public.karyawan for select using (true);

drop policy if exists lembur_read on public.lembur;
create policy lembur_read on public.lembur for select using (true);

drop policy if exists lembur_write on public.lembur;
create policy lembur_write on public.lembur for all using (true) with check (true);

drop policy if exists pengaturan_read on public.pengaturan;
create policy pengaturan_read on public.pengaturan for select using (true);

drop policy if exists pengaturan_write on public.pengaturan;
create policy pengaturan_write on public.pengaturan for all using (true) with check (true);

drop policy if exists profil_all on public.profil;
create policy profil_all on public.profil for all using (true) with check (true);
