-- ============================================================
-- Seed master data karyawan Office WKI
-- Jalankan SETELAH schema.sql
-- ============================================================
insert into public.karyawan (no, nama, nomor_induk, ttl, tahun_masuk, jabatan, kategori, boleh_lembur)
values
  (1, 'PADMONO',              '6525255WKI', 'Jakarta, 7 Desember 1965',      '24 Maret 2025', 'Komisaris',                'SPL',       false),
  (2, 'M. JUFRI ARFAH',       '6025254WKI', 'Ujung Pandang, 24 Mei 1960',    '24 Maret 2025', 'Direktur',                 'SPL',       false),
  (3, 'SUHERMAN',             '7422018WKI', 'Jakarta, 15 Juni 1976',         '01 Juli 2022',  'Staff Marketing',          'SPL',       true),
  (4, 'EDDY DJAELANI',        '7322020WKI', 'Bogor, 5 Mei 1972',             '01 Juli 2022',  'Staff Marketing',          'SPL',       true),
  (5, 'BOBBY ANUGRAH SIGIRO', '9822198WKI', 'Tangerang, 9 Mei 1998',         '01 April 2022', 'Staff Marketing',          'SPL',       true),
  (6, 'RISKA RAMADHANTI',     '9825257WKI', 'Jakarta, 23 Januari 1998',      '14 Mei 2025',   'Staff Keuangan dan Pajak', 'SPL',       true),
  (7, 'YESI BUDI INDRIANI',   '9926263WKI', 'Cilacap, 01 Maret 1999',        '29 April 2026', 'Staff SDM dan Umum',       'SPL',       true),
  (8, 'UJANG NURYADIN',       '7422023WKI', 'Lubuk Linggau, 08 Maret 1974',  '01 Juli 2022',  'Pengemudi',                'Pengemudi', true)
on conflict (nomor_induk) do nothing;
