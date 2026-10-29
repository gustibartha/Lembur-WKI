import React, { useState, useMemo, useEffect, useRef } from "react";
import QRCode from "qrcode";
import * as XLSX from "xlsx";
import {
  isSupabaseEnabled,
  listEntries as repoListEntries,
  createEntry as repoCreateEntry,
  updateEntry as repoUpdateEntry,
  deleteEntry as repoDeleteEntry,
} from "./lib/repo";
import {
  Clock,
  Plus,
  Trash2,
  Send,
  LayoutDashboard,
  History,
  User,
  Sun,
  Moon,
  CircleDashed,
  CheckCheck,
  LogOut,
  ArrowLeft,
  Calendar,
  Wallet,
  Settings as SettingsIcon,
  Paperclip,
  FileText,
  ImageIcon,
  X,
  Users,
  Download,
  ClipboardList,
  ChevronRight,
  Check,
  Printer,
  PenLine,
  QrCode,
  Eraser,
  Droplets,
  ArrowUpRight,
  Sparkles,
  Car,
  Building2,
  CalendarClock,
  FileSpreadsheet,
  Filter,
} from "lucide-react";

// ---------- Storage ----------
const STORAGE = {
  user: "lembur:user",
  entries: "lembur:entries:v6",
  tarif: "lembur:tarif",
  gajiPokok: "lembur:gajiPokok",
  theme: "lembur:theme",
  company: "lembur:company:v2",
};

const COMPANY_DEFAULT = {
  nama: "PT. WIRA KARYA INDONESIA",
  inisial: "WKI",
  kota: "Jakarta",
  // Water Intake unit (PLTGU Blok 1 Muara Karang) config
  waterIntake: {
    judul: "DAFTAR LEMBUR PETUGAS KEBERSIHAN",
    subjudul: "WATER INTAKE PLTGU BLOK 1 MUARA KARANG",
    kota: "Jakarta",
    instansi: "PT. PLN NP UP Muara Karang",
    koordinatorRegu: "Earlan S.",
    admSdm: "",
    mengetahui: "Roshie",
  },
};

const KATEGORI = {
  SPL: {
    label: "Surat Perintah Lembur",
    short: "SPL Staff",
    formula: "office",
  },
  WaterIntake: {
    label: "Daftar Lembur Petugas Kebersihan",
    short: "Water Intake",
    formula: "tad",
  },
  Pengemudi: {
    label: "Daftar Lembur Pengemudi",
    short: "Pengemudi",
    formula: "tad",
  },
};

const TIPE_ENTRI = {
  Lembur: { label: "Lembur", icon: Clock },
  Piket: { label: "Piket", icon: CalendarClock },
};

// Gaji pokok default dari Excel "PERHITUNGAN LEMBUR TAD & OFFICE"
const GAJI_POKOK_DEFAULT = 5_729_876;
const PEMBAGI = 173; // jam kerja per bulan (UU Ketenagakerjaan)
// Konstanta dari Excel: 927378.774566474 (libur > 6 jam offset)
const LIBUR_OFFSET_CONST = 927_378.774566474;

const HARI_ID = [
  "Minggu",
  "Senin",
  "Selasa",
  "Rabu",
  "Kamis",
  "Jumat",
  "Sabtu",
];
const namaHari = (iso) => {
  if (!iso) return "";
  return HARI_ID[new Date(iso).getDay()] || "";
};

const load = (k, fallback) => {
  try {
    const v = localStorage.getItem(k);
    return v ? JSON.parse(v) : fallback;
  } catch {
    return fallback;
  }
};
const save = (k, v) => {
  try {
    localStorage.setItem(k, JSON.stringify(v));
  } catch (err) {
    console.warn("storage error", err);
    alert("Gagal menyimpan: kemungkinan kuota browser penuh (bukti terlalu besar).");
  }
};

// ---------- Demo data ----------
const DEMO_USERS = [
  {
    id: "u_andi",
    nama: "Andi Saputra",
    email: "andi@kerja.id",
    peran: "Pekerja",
    bagian: "IT Support",
    pekerjaan: "Network Engineer",
    penanggungJawab: "Rudi (Supervisor IT)",
  },
  {
    id: "u_rina",
    nama: "Rina Marlina",
    email: "rina@kerja.id",
    peran: "Pekerja",
    bagian: "Operasional",
    pekerjaan: "Operations Analyst",
    penanggungJawab: "Yusuf (Head Ops)",
  },
  {
    id: "u_budi",
    nama: "Budi Hartono",
    email: "budi@kerja.id",
    peran: "Pekerja",
    bagian: "Gudang",
    pekerjaan: "Staff Gudang",
    penanggungJawab: "Hasan (Kepala Gudang)",
  },
  {
    id: "u_warno",
    nama: "Warno",
    email: "warno@kerja.id",
    peran: "Pekerja",
    bagian: "Petugas Kebersihan",
    pekerjaan: "Petugas Water Intake",
    penanggungJawab: "Earlan S. (Koordinator Regu)",
    kategoriDefault: "WaterIntake",
    gajiPokok: 5_729_876,
  },
  {
    id: "u_dedy",
    nama: "Dedy Krismulyana",
    email: "dedy@kerja.id",
    peran: "Pekerja",
    bagian: "Pengemudi",
    pekerjaan: "Pengemudi TAD",
    penanggungJawab: "Wahyudi (Kepala Pengemudi)",
    kategoriDefault: "Pengemudi",
    gajiPokok: 5_729_876,
  },
];
const DEMO_ATASAN = {
  id: "u_admin",
  nama: "Sari Wulandari",
  email: "sari@kerja.id",
  peran: "Atasan",
  bagian: "HR",
  pekerjaan: "HR Manager",
};
const DEMO_ADMIN = {
  id: "u_payroll",
  nama: "Ani Pratiwi",
  email: "ani@kerja.id",
  peran: "Admin",
  bagian: "Payroll",
  pekerjaan: "Admin Payroll",
};

// ---------- Master data karyawan Office WKI ----------
// bolehLembur: Komisaris & Direktur (no.1-2) TIDAK dapat lembur.
const KARYAWAN_WKI = [
  {
    no: 1,
    nama: "PADMONO",
    nomorInduk: "6525255WKI",
    ttl: "Jakarta, 7 Desember 1965",
    tahunMasuk: "24 Maret 2025",
    jabatan: "Komisaris",
    bolehLembur: false,
    kategori: "SPL",
  },
  {
    no: 2,
    nama: "M. JUFRI ARFAH",
    nomorInduk: "6025254WKI",
    ttl: "Ujung Pandang, 24 Mei 1960",
    tahunMasuk: "24 Maret 2025",
    jabatan: "Direktur",
    bolehLembur: false,
    kategori: "SPL",
  },
  {
    no: 3,
    nama: "SUHERMAN",
    nomorInduk: "7422018WKI",
    ttl: "Jakarta, 15 Juni 1976",
    tahunMasuk: "01 Juli 2022",
    jabatan: "Staff Marketing",
    bolehLembur: true,
    kategori: "SPL",
  },
  {
    no: 4,
    nama: "EDDY DJAELANI",
    nomorInduk: "7322020WKI",
    ttl: "Bogor, 5 Mei 1972",
    tahunMasuk: "01 Juli 2022",
    jabatan: "Staff Marketing",
    bolehLembur: true,
    kategori: "SPL",
  },
  {
    no: 5,
    nama: "BOBBY ANUGRAH SIGIRO",
    nomorInduk: "9822198WKI",
    ttl: "Tangerang, 9 Mei 1998",
    tahunMasuk: "01 April 2022",
    jabatan: "Staff Marketing",
    bolehLembur: true,
    kategori: "SPL",
  },
  {
    no: 6,
    nama: "RISKA RAMADHANTI",
    nomorInduk: "9825257WKI",
    ttl: "Jakarta, 23 Januari 1998",
    tahunMasuk: "14 Mei 2025",
    jabatan: "Staff Keuangan dan Pajak",
    bolehLembur: true,
    kategori: "SPL",
  },
  {
    no: 7,
    nama: "YESI BUDI INDRIANI",
    nomorInduk: "9926263WKI",
    ttl: "Cilacap, 01 Maret 1999",
    tahunMasuk: "29 April 2026",
    jabatan: "Staff SDM dan Umum",
    bolehLembur: true,
    kategori: "SPL",
  },
  {
    no: 8,
    nama: "UJANG NURYADIN",
    nomorInduk: "7422023WKI",
    ttl: "Lubuk Linggau, 08 Maret 1974",
    tahunMasuk: "01 Juli 2022",
    jabatan: "Pengemudi",
    bolehLembur: true,
    kategori: "Pengemudi",
  },
];

// ---------- Domain ----------
// Multiplier table (display saja — perhitungan riil pakai formula per-kategori di hitungGaji)
const MULTIPLIER = { "Hari Kerja": 1.5, "Hari Libur": 2 };

// Hitung upah lembur sesuai Excel "PERHITUNGAN LEMBUR TAD & OFFICE":
// - kategori "office" (SPL Staff): flat per jam, max 8 jam/hari
// - kategori "tad" (Water Intake / Pengemudi): formula bertingkat
function hitungGaji(jam, jenis, kategori = "SPL", gajiPokok = GAJI_POKOK_DEFAULT) {
  const j = Number(jam || 0);
  if (j <= 0) return 0;
  const gp = Number(gajiPokok || GAJI_POKOK_DEFAULT);
  const base = gp / PEMBAGI;
  const formula = KATEGORI[kategori]?.formula || "office";

  if (formula === "office") {
    // TUL Office: per jam × base, capped 8 jam (kolom F:M di Excel)
    return Math.round(Math.min(j, 8) * base);
  }

  // TAD tier formula (Water Intake / Pengemudi)
  const B8 = 1.5 * base;
  const B9 = 2.0 * base;
  const B10 = 4.0 * base;
  const B11 = B9 - B8;
  const B12 = LIBUR_OFFSET_CONST - B8 * 5 - B9 - B10;

  if (jenis === "Hari Libur") {
    if (j < 6) return Math.round(j * B8);
    if (j === 6) return Math.round(B9 + B8 * 5);
    return Math.round(j * B10 - B12);
  }
  // Hari Kerja
  if (j < 2) return Math.round(j * B8);
  return Math.round(j * B9 - B11);
}

// Estimasi tarif/jam (untuk display di dashboard)
const tarifPerJam = (gajiPokok = GAJI_POKOK_DEFAULT) =>
  Math.round((Number(gajiPokok) || GAJI_POKOK_DEFAULT) / PEMBAGI);

const formatRp = (n) =>
  "Rp " + Number(n || 0).toLocaleString("id-ID", { maximumFractionDigits: 0 });

const formatTanggal = (iso) => {
  if (!iso) return "";
  const d = new Date(iso);
  return d.toLocaleDateString("id-ID", { day: "numeric", month: "short", year: "numeric" });
};

const formatTanggalPanjang = (iso) => {
  if (!iso) return "";
  const d = new Date(iso);
  return d.toLocaleDateString("id-ID", {
    day: "numeric",
    month: "long",
    year: "numeric",
  });
};

// Hitung durasi jam dari (tanggal+jam mulai) → (tanggal+jam selesai)
const computeDuration = (tglMulai, jamMulai, tglSelesai, jamSelesai) => {
  if (!tglMulai || !jamMulai || !tglSelesai || !jamSelesai) return 0;
  const start = new Date(`${tglMulai}T${jamMulai}:00`);
  const end = new Date(`${tglSelesai}T${jamSelesai}:00`);
  if (isNaN(start) || isNaN(end) || end <= start) return 0;
  return Math.round(((end - start) / 3_600_000) * 100) / 100;
};

const startOfWeek = () => {
  const d = new Date();
  const day = (d.getDay() + 6) % 7;
  d.setDate(d.getDate() - day);
  d.setHours(0, 0, 0, 0);
  return d;
};
const startOfMonth = () => {
  const d = new Date();
  d.setDate(1);
  d.setHours(0, 0, 0, 0);
  return d;
};

const STATUS = {
  Draft: {
    label: "Draft",
    cls: "bg-zinc-200 text-zinc-700 dark:bg-zinc-800 dark:text-zinc-300",
    icon: CircleDashed,
  },
  Terkirim: {
    label: "Terkirim",
    cls: "bg-amber-100 text-amber-800 dark:bg-amber-400/15 dark:text-amber-300",
    icon: Send,
  },
  Disetujui: {
    label: "Disetujui",
    cls: "bg-emerald-100 text-emerald-800 dark:bg-emerald-400/15 dark:text-emerald-300",
    icon: CheckCheck,
  },
};

const mkEntry = (o) => {
  const durasi = computeDuration(
    o.tanggalMulai,
    o.jamMulai,
    o.tanggalSelesai,
    o.jamSelesai
  );
  return {
    kategori: "SPL",
    tipe: "Lembur",
    durasiJam: durasi,
    estimasiGaji: hitungGaji(durasi, o.jenis, o.kategori || "SPL"),
    status: "Draft",
    evidence: null,
    ...o,
  };
};

const seedEntries = () => [
  mkEntry({
    id: 1,
    userId: "u_andi",
    userNama: "Andi Saputra",
    bagian: "IT Support",
    pekerjaan: "Maintenance jaringan kantor",
    tujuan: "Memastikan koneksi internet stabil sebelum jam kerja",
    tanggalMulai: "2026-05-20",
    jamMulai: "18:00",
    tanggalSelesai: "2026-05-20",
    jamSelesai: "21:30",
    jenis: "Hari Kerja",
    hasil: "Switch lantai 3 diganti, koneksi normal",
    keterangan: "Closing laporan Q1",
    penanggungJawab: "Rudi (Supervisor IT)",
    status: "Draft",
  }),
  mkEntry({
    id: 2,
    userId: "u_rina",
    userNama: "Rina Marlina",
    bagian: "Operasional",
    pekerjaan: "Migrasi server produksi",
    tujuan: "Pindah ke datacenter baru tanpa downtime",
    tanggalMulai: "2026-05-17",
    jamMulai: "08:00",
    tanggalSelesai: "2026-05-17",
    jamSelesai: "14:00",
    jenis: "Hari Libur",
    hasil: "8 service termigrasi, downtime 12 menit",
    keterangan: "Tim 3 orang, koordinasi via WA",
    penanggungJawab: "Yusuf (Head Ops)",
    status: "Terkirim",
  }),
  mkEntry({
    id: 3,
    userId: "u_budi",
    userNama: "Budi Hartono",
    bagian: "Gudang",
    pekerjaan: "Stock opname tahunan",
    tujuan: "Mencocokkan stok fisik dengan sistem",
    tanggalMulai: "2026-05-15",
    jamMulai: "17:00",
    tanggalSelesai: "2026-05-15",
    jamSelesai: "19:00",
    jenis: "Hari Kerja",
    hasil: "Selisih 0,3% (dalam toleransi)",
    keterangan: "",
    penanggungJawab: "Hasan (Kepala Gudang)",
    status: "Disetujui",
  }),
  mkEntry({
    id: 4,
    userId: "u_andi",
    userNama: "Andi Saputra",
    bagian: "IT Support",
    pekerjaan: "Pemasangan kabel jaringan lantai 3",
    tujuan: "Persiapan ruang meeting baru",
    tanggalMulai: "2026-05-14",
    jamMulai: "18:00",
    tanggalSelesai: "2026-05-14",
    jamSelesai: "22:00",
    jenis: "Hari Kerja",
    hasil: "12 titik aktif, rapi",
    keterangan: "",
    penanggungJawab: "Rudi (Supervisor IT)",
    status: "Disetujui",
  }),
  mkEntry({
    id: 5,
    userId: "u_rina",
    userNama: "Rina Marlina",
    bagian: "Operasional",
    pekerjaan: "Audit data inventory",
    tujuan: "Verifikasi data sebelum tutup buku",
    tanggalMulai: "2026-05-12",
    jamMulai: "19:00",
    tanggalSelesai: "2026-05-12",
    jamSelesai: "21:30",
    jenis: "Hari Kerja",
    hasil: "Selesai 4 dari 5 segmen",
    keterangan: "Sisanya dilanjutkan besok pagi",
    penanggungJawab: "Yusuf (Head Ops)",
    status: "Terkirim",
  }),
  // ----- Water Intake entries -----
  mkEntry({
    id: 6,
    kategori: "WaterIntake",
    userId: "u_warno",
    userNama: "Warno",
    pekerjaan: "Ngantiin cuti Aries Munadar",
    tanggalMulai: "2026-04-20",
    jamMulai: "15:00",
    tanggalSelesai: "2026-04-20",
    jamSelesai: "22:00",
    jenis: "Hari Kerja",
    keterangan: "OK",
    status: "Terkirim",
  }),
  mkEntry({
    id: 7,
    kategori: "WaterIntake",
    userId: "u_warno",
    userNama: "Warno",
    pekerjaan: "Ngantiin cuti Aries Munadar",
    tanggalMulai: "2026-04-21",
    jamMulai: "15:00",
    tanggalSelesai: "2026-04-21",
    jamSelesai: "22:00",
    jenis: "Hari Kerja",
    keterangan: "OK",
    status: "Terkirim",
  }),
  // ----- Pengemudi entries -----
  mkEntry({
    id: 8,
    kategori: "Pengemudi",
    userId: "u_dedy",
    userNama: "Dedy Krismulyana",
    pekerjaan: "Antar dinas direktur ke Bandung",
    tanggalMulai: "2026-05-22",
    jamMulai: "17:00",
    tanggalSelesai: "2026-05-22",
    jamSelesai: "22:00",
    jenis: "Hari Kerja",
    keterangan: "OK",
    penanggungJawab: "Wahyudi",
    status: "Terkirim",
  }),
  mkEntry({
    id: 9,
    kategori: "Pengemudi",
    userId: "u_dedy",
    userNama: "Dedy Krismulyana",
    pekerjaan: "Standby tim teknis ke Cilegon",
    tanggalMulai: "2026-05-24",
    jamMulai: "06:00",
    tanggalSelesai: "2026-05-24",
    jamSelesai: "14:00",
    jenis: "Hari Libur",
    keterangan: "OK",
    penanggungJawab: "Wahyudi",
    status: "Terkirim",
  }),
  // ----- Piket Pengemudi entries -----
  mkEntry({
    id: 10,
    kategori: "Pengemudi",
    tipe: "Piket",
    userId: "u_dedy",
    userNama: "Dedy Krismulyana",
    pekerjaan: "Piket malam — kantor pusat",
    tanggalMulai: "2026-05-23",
    jamMulai: "22:00",
    tanggalSelesai: "2026-05-24",
    jamSelesai: "06:00",
    jenis: "Hari Libur",
    keterangan: "Tidak ada penugasan, standby",
    penanggungJawab: "Wahyudi",
    status: "Disetujui",
  }),
  mkEntry({
    id: 11,
    kategori: "Pengemudi",
    tipe: "Piket",
    userId: "u_dedy",
    userNama: "Dedy Krismulyana",
    pekerjaan: "Piket akhir pekan",
    tanggalMulai: "2026-05-18",
    jamMulai: "08:00",
    tanggalSelesai: "2026-05-18",
    jamSelesai: "17:00",
    jenis: "Hari Libur",
    keterangan: "Antar 2x ke bandara",
    penanggungJawab: "Wahyudi",
    status: "Draft",
  }),
];

// ---------- Root ----------
export default function LemburApp() {
  const [user, setUser] = useState(() => load(STORAGE.user, null));
  const [theme, setTheme] = useState(
    () => localStorage.getItem(STORAGE.theme) || "light"
  );
  const [entries, setEntries] = useState(() => load(STORAGE.entries, seedEntries()));
  const [tarif, setTarif] = useState(() => load(STORAGE.tarif, 35000));
  const [gajiPokok, setGajiPokok] = useState(() =>
    load(STORAGE.gajiPokok, GAJI_POKOK_DEFAULT)
  );
  const [company, setCompany] = useState(() => load(STORAGE.company, COMPANY_DEFAULT));
  const [view, setView] = useState("dashboard");
  const [viewerEvidence, setViewerEvidence] = useState(null);
  const [suratEntry, setSuratEntry] = useState(null);
  const [approveTarget, setApproveTarget] = useState(null);

  useEffect(() => {
    const link = document.createElement("link");
    link.href =
      "https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,300..900&family=Geist:wght@300;400;500;600&family=Geist+Mono:wght@400;500&display=swap";
    link.rel = "stylesheet";
    document.head.appendChild(link);
    return () => document.head.removeChild(link);
  }, []);

  // Supabase: muat entri dari server saat mount (menggantikan seed lokal)
  useEffect(() => {
    if (!isSupabaseEnabled) return;
    let alive = true;
    repoListEntries(seedEntries())
      .then((rows) => {
        if (alive) setEntries(rows);
      })
      .catch((err) => console.error("Gagal memuat entri Supabase:", err));
    return () => {
      alive = false;
    };
  }, []);

  // Persist ke localStorage HANYA di mode lokal (Supabase pakai write-through)
  useEffect(() => {
    if (!isSupabaseEnabled) save(STORAGE.entries, entries);
  }, [entries]);
  useEffect(() => save(STORAGE.tarif, tarif), [tarif]);
  useEffect(() => save(STORAGE.gajiPokok, gajiPokok), [gajiPokok]);
  useEffect(() => save(STORAGE.company, company), [company]);
  useEffect(() => {
    localStorage.setItem(STORAGE.theme, theme);
    document.documentElement.classList.toggle("dark", theme === "dark");
  }, [theme]);
  useEffect(() => {
    if (user) save(STORAGE.user, user);
    else localStorage.removeItem(STORAGE.user);
  }, [user]);

  const serif = { fontFamily: "'Fraunces', serif" };
  const sans = { fontFamily: "'Geist', system-ui, sans-serif" };
  const mono = { fontFamily: "'Geist Mono', monospace" };

  const isAdmin = user?.peran === "Atasan";
  const isPayroll = user?.peran === "Admin";

  // Mutators
  const addEntry = (data) => {
    const durasi = computeDuration(
      data.tanggalMulai,
      data.jamMulai,
      data.tanggalSelesai,
      data.jamSelesai
    );
    const userGp = user.gajiPokok || gajiPokok;
    const entry = {
      id: Date.now(),
      userId: user.id,
      userNama: user.nama,
      karyawanId: user.karyawanId ?? null,
      ...data,
      durasiJam: durasi,
      estimasiGaji: hitungGaji(durasi, data.jenis, data.kategori, userGp),
      gajiPokokSaat: userGp,
      status: "Draft",
    };
    setEntries((arr) => [entry, ...arr]);
    setView("riwayat");
    if (isSupabaseEnabled) {
      repoCreateEntry(entry)
        .then((saved) =>
          // ganti id sementara dgn id (uuid) dari server
          setEntries((arr) => arr.map((e) => (e.id === entry.id ? saved : e)))
        )
        .catch((err) => console.error("Gagal simpan entri:", err));
    }
  };
  const deleteEntry = (id) => {
    setEntries((arr) => arr.filter((e) => e.id !== id));
    if (isSupabaseEnabled)
      repoDeleteEntry(id).catch((err) => console.error("Gagal hapus:", err));
  };
  const kirimSemuaDraft = () => {
    const patch = { status: "Terkirim", sentAt: new Date().toISOString() };
    const targets = entries.filter(
      (e) => e.userId === user.id && e.status === "Draft"
    );
    setEntries((arr) =>
      arr.map((e) =>
        e.userId === user.id && e.status === "Draft" ? { ...e, ...patch } : e
      )
    );
    if (isSupabaseEnabled)
      targets.forEach((e) =>
        repoUpdateEntry(e.id, patch).catch((err) =>
          console.error("Gagal kirim:", err)
        )
      );
  };

  const requestApprove = (entry) => setApproveTarget({ kind: "single", entry });
  const requestBatchApprove = (ids) => {
    if (!ids || ids.length === 0) return;
    setApproveTarget({ kind: "batch", ids });
  };

  const finalizeApproval = ({ ttdAtasan, catatanAtasan }) => {
    const stamp = {
      status: "Disetujui",
      ttdAtasan,
      catatanAtasan,
      atasanNama: user.nama,
      approvedAt: new Date().toISOString(),
    };
    const ids =
      approveTarget.kind === "single"
        ? [approveTarget.entry.id]
        : approveTarget.ids;
    setEntries((arr) =>
      arr.map((e) => (ids.includes(e.id) ? { ...e, ...stamp } : e))
    );
    if (suratEntry && ids.includes(suratEntry.id)) {
      setSuratEntry({ ...suratEntry, ...stamp });
    }
    if (isSupabaseEnabled)
      ids.forEach((id) =>
        repoUpdateEntry(id, stamp).catch((err) =>
          console.error("Gagal setujui:", err)
        )
      );
    setApproveTarget(null);
  };

  // Reset view when role context changes
  useEffect(() => {
    if (!user) return;
    setView(isAdmin ? "dashboard" : "dashboard");
  }, [user?.id]);

  // ---------- LOGIN ----------
  if (!user) {
    return (
      <Shell theme={theme} sans={sans}>
        <Login onLogin={setUser} serif={serif} mono={mono} />
      </Shell>
    );
  }

  // ---------- APP ----------
  const ownEntries = entries.filter((e) => e.userId === user.id);

  return (
    <Shell theme={theme} sans={sans}>
      <div className="mx-auto max-w-md min-h-screen flex flex-col">
        <TopBar
          user={user}
          theme={theme}
          onToggleTheme={() => setTheme(theme === "dark" ? "light" : "dark")}
          mono={mono}
          serif={serif}
        />

        <main className="flex-1 px-5 pb-28 pt-2">
          {isPayroll ? (
            <>
              {view === "dashboard" && (
                <PayrollDashboard
                  user={user}
                  entries={entries}
                  setView={setView}
                  serif={serif}
                  mono={mono}
                />
              )}
              {view === "rekap" && (
                <PayrollRekap
                  entries={entries}
                  gajiPokok={gajiPokok}
                  onViewSurat={setSuratEntry}
                  onViewEvidence={setViewerEvidence}
                  serif={serif}
                  mono={mono}
                />
              )}
              {view === "profil" && (
                <Profil
                  user={user}
                  setUser={setUser}
                  tarif={tarif}
                  setTarif={setTarif}
                  gajiPokok={gajiPokok}
                  setGajiPokok={setGajiPokok}
                  company={company}
                  setCompany={setCompany}
                  onLogout={() => setUser(null)}
                  serif={serif}
                  mono={mono}
                />
              )}
            </>
          ) : isAdmin ? (
            <>
              {view === "dashboard" && (
                <AdminDashboard
                  user={user}
                  entries={entries}
                  setView={setView}
                  onViewSurat={setSuratEntry}
                  serif={serif}
                  mono={mono}
                />
              )}
              {view === "rekap" && (
                <AdminRekap
                  entries={entries}
                  onApprove={requestApprove}
                  onApproveBatch={requestBatchApprove}
                  onViewEvidence={setViewerEvidence}
                  onViewSurat={setSuratEntry}
                  serif={serif}
                  mono={mono}
                />
              )}
              {view === "profil" && (
                <Profil
                  user={user}
                  setUser={setUser}
                  tarif={tarif}
                  setTarif={setTarif}
                  gajiPokok={gajiPokok}
                  setGajiPokok={setGajiPokok}
                  company={company}
                  setCompany={setCompany}
                  onLogout={() => setUser(null)}
                  serif={serif}
                  mono={mono}
                />
              )}
            </>
          ) : (
            <>
              {view === "dashboard" && (
                <Dashboard
                  user={user}
                  entries={ownEntries}
                  gajiPokok={user.gajiPokok || gajiPokok}
                  setView={setView}
                  onViewSurat={setSuratEntry}
                  serif={serif}
                  mono={mono}
                />
              )}
              {view === "input" && (
                <InputForm
                  user={user}
                  gajiPokok={user.gajiPokok || gajiPokok}
                  onSave={addEntry}
                  onCancel={() => setView("dashboard")}
                  serif={serif}
                  mono={mono}
                />
              )}
              {view === "riwayat" && (
                <Riwayat
                  entries={ownEntries}
                  onDelete={deleteEntry}
                  onKirim={kirimSemuaDraft}
                  onViewEvidence={setViewerEvidence}
                  onViewSurat={setSuratEntry}
                  serif={serif}
                  mono={mono}
                />
              )}
              {view === "profil" && (
                <Profil
                  user={user}
                  setUser={setUser}
                  tarif={tarif}
                  setTarif={setTarif}
                  gajiPokok={gajiPokok}
                  setGajiPokok={setGajiPokok}
                  company={company}
                  setCompany={setCompany}
                  onLogout={() => setUser(null)}
                  serif={serif}
                  mono={mono}
                />
              )}
            </>
          )}
        </main>

        <BottomNav
          view={view}
          setView={setView}
          isAdmin={isAdmin}
          isPayroll={isPayroll}
          mono={mono}
        />

        {viewerEvidence && (
          <EvidenceViewer
            evidence={viewerEvidence}
            onClose={() => setViewerEvidence(null)}
            mono={mono}
            serif={serif}
          />
        )}

        {suratEntry &&
          (suratEntry.kategori === "WaterIntake" ||
            suratEntry.kategori === "Pengemudi") && (
          <DaftarLemburView
            anchorEntry={suratEntry}
            allEntries={entries}
            company={company}
            currentUser={user}
            onApproveBatch={
              isAdmin
                ? (ids) => {
                    setApproveTarget({ kind: "batch", ids });
                  }
                : null
            }
            onClose={() => setSuratEntry(null)}
            mono={mono}
            serif={serif}
          />
        )}
        {suratEntry &&
          suratEntry.kategori !== "WaterIntake" &&
          suratEntry.kategori !== "Pengemudi" && (
          <SuratView
            entry={suratEntry}
            company={company}
            currentUser={user}
            onApprove={
              isAdmin && suratEntry.status === "Terkirim"
                ? () => requestApprove(suratEntry)
                : null
            }
            onClose={() => setSuratEntry(null)}
            mono={mono}
            serif={serif}
          />
        )}

        {approveTarget && (
          <ApprovalDialog
            target={approveTarget}
            currentUser={user}
            onConfirm={finalizeApproval}
            onCancel={() => setApproveTarget(null)}
            mono={mono}
            serif={serif}
          />
        )}
      </div>
    </Shell>
  );
}

// ---------- Shell ----------
function Shell({ children, theme, sans }) {
  return (
    <div
      className="min-h-screen w-full overflow-x-hidden transition-colors text-zinc-900 dark:text-zinc-100"
      style={{
        ...sans,
        background:
          theme === "dark"
            ? "linear-gradient(180deg, #2e2e38 0%, #25252e 100%)"
            : "#ffffff",
      }}
    >
      {/* grid texture */}
      <div
        className="pointer-events-none fixed inset-0 z-0 opacity-[0.04] dark:opacity-[0.07]"
        style={{
          backgroundImage:
            "linear-gradient(currentColor 1px, transparent 1px), linear-gradient(90deg, currentColor 1px, transparent 1px)",
          backgroundSize: "56px 56px",
        }}
      />
      {/* animated gradient blobs - lebih cerah */}
      <div
        className="pointer-events-none fixed inset-0 z-0 overflow-hidden"
        aria-hidden
      >
        <div
          className="absolute animate-float"
          style={{
            top: "-20%",
            left: "-10%",
            width: "70vw",
            height: "70vw",
            maxWidth: 900,
            maxHeight: 900,
            background:
              theme === "dark"
                ? "radial-gradient(circle, rgba(251,191,36,0.55), transparent 60%)"
                : "radial-gradient(circle, rgba(251,191,36,0.22), transparent 60%)",
            filter: "blur(60px)",
          }}
        />
        <div
          className="absolute animate-float-delay"
          style={{
            bottom: "-20%",
            right: "-15%",
            width: "60vw",
            height: "60vw",
            maxWidth: 700,
            maxHeight: 700,
            background:
              theme === "dark"
                ? "radial-gradient(circle, rgba(244,114,182,0.35), transparent 60%)"
                : "radial-gradient(circle, rgba(244,114,182,0.18), transparent 60%)",
            filter: "blur(80px)",
          }}
        />
        <div
          className="absolute animate-float"
          style={{
            top: "30%",
            right: "10%",
            width: "45vw",
            height: "45vw",
            maxWidth: 500,
            maxHeight: 500,
            background:
              theme === "dark"
                ? "radial-gradient(circle, rgba(56,189,248,0.30), transparent 60%)"
                : "radial-gradient(circle, rgba(56,189,248,0.15), transparent 60%)",
            filter: "blur(70px)",
            animationDelay: "-9s",
          }}
        />
        <div
          className="absolute animate-float-delay"
          style={{
            top: "55%",
            left: "20%",
            width: "40vw",
            height: "40vw",
            maxWidth: 450,
            maxHeight: 450,
            background:
              theme === "dark"
                ? "radial-gradient(circle, rgba(134,239,172,0.25), transparent 60%)"
                : "radial-gradient(circle, rgba(134,239,172,0.14), transparent 60%)",
            filter: "blur(75px)",
            animationDelay: "-14s",
          }}
        />
      </div>
      <div className="relative z-10">{children}</div>
    </div>
  );
}

// ---------- Login ----------
function Login({ onLogin, serif, mono }) {
  const [nama, setNama] = useState("");
  const [email, setEmail] = useState("");
  const [profil, setProfil] = useState("Organik"); // Organik | WaterIntake | Pengemudi | Atasan | Admin
  const [nip, setNip] = useState(""); // nomor induk karyawan WKI terpilih

  // Roster berbasis profil (Organik = staff SPL, Pengemudi = kategori Pengemudi)
  const roster = useMemo(() => {
    if (profil === "Organik")
      return KARYAWAN_WKI.filter((k) => k.kategori === "SPL");
    if (profil === "Pengemudi")
      return KARYAWAN_WKI.filter((k) => k.kategori === "Pengemudi");
    return [];
  }, [profil]);
  const usesRoster = profil === "Organik" || profil === "Pengemudi";
  const selectedKaryawan = useMemo(
    () => KARYAWAN_WKI.find((k) => k.nomorInduk === nip) || null,
    [nip]
  );

  const submit = () => {
    const peranMap = { Atasan: "Atasan", Admin: "Admin" };
    const peran = peranMap[profil] || "Pekerja";

    if (usesRoster) {
      // wajib pilih karyawan yang boleh lembur
      if (!selectedKaryawan || !selectedKaryawan.bolehLembur) return;
      onLogin({
        id: "u_wki_" + selectedKaryawan.nomorInduk,
        nama: selectedKaryawan.nama,
        email:
          email.trim() ||
          `${selectedKaryawan.nama.toLowerCase().replace(/\s+/g, ".")}@wki.co.id`,
        peran: "Pekerja",
        kategoriDefault: selectedKaryawan.kategori,
        nomorInduk: selectedKaryawan.nomorInduk,
        jabatan: selectedKaryawan.jabatan,
        bagian: selectedKaryawan.jabatan,
        pekerjaan: selectedKaryawan.jabatan,
        ttl: selectedKaryawan.ttl,
        tahunMasuk: selectedKaryawan.tahunMasuk,
      });
      return;
    }

    if (!nama.trim()) return;
    onLogin({
      id: peran === "Atasan" ? "u_admin_custom" : "u_payroll_custom",
      nama: nama.trim(),
      email:
        email.trim() ||
        `${nama.trim().toLowerCase().replace(/\s+/g, ".")}@kerja.id`,
      peran,
      kategoriDefault: null,
    });
  };

  const canSubmit = usesRoster
    ? selectedKaryawan && selectedKaryawan.bolehLembur
    : nama.trim().length > 0;

  const demoOrganik = () => onLogin(DEMO_USERS[0]); // Andi
  const demoWater = () => onLogin(DEMO_USERS.find((u) => u.id === "u_warno"));
  const demoSopir = () => onLogin(DEMO_USERS.find((u) => u.id === "u_dedy"));
  const demoAtasan = () => onLogin(DEMO_ATASAN);
  const demoAdmin = () => onLogin(DEMO_ADMIN);

  const PROFIL_OPTIONS = [
    { k: "Organik", l: "Organik WKI", icon: Building2 },
    { k: "WaterIntake", l: "Water Intake", icon: Droplets },
    { k: "Pengemudi", l: "Pengemudi", icon: Car },
    { k: "Atasan", l: "Atasan", icon: User },
    { k: "Admin", l: "Admin Payroll", icon: FileSpreadsheet },
  ];

  return (
    <div className="mx-auto max-w-md min-h-screen flex flex-col justify-center px-6 py-10 animate-fade-in">
      <div className="mb-8 animate-slide-up">
        <div className="flex items-center gap-2.5 mb-10">
          <div className="relative">
            <div className="absolute inset-0 bg-amber-400 rounded-xl blur-md opacity-50 animate-pulse-glow" />
            <div className="relative w-11 h-11 rounded-xl bg-gradient-to-br from-amber-300 to-amber-500 flex items-center justify-center shadow-lg shadow-amber-500/30">
              <Clock className="w-5 h-5 text-zinc-950" strokeWidth={2.5} />
            </div>
          </div>
          <span className="text-2xl tracking-tight" style={serif}>
            Lembur<span className="italic text-amber-500 dark:text-amber-400">.id</span>
          </span>
          <span
            className="ml-auto text-[9px] uppercase tracking-widest text-zinc-500 border border-zinc-300 dark:border-zinc-600 px-2 py-1 rounded-full"
            style={mono}
          >
            v2.6
          </span>
        </div>
        <h1
          className="text-5xl leading-[0.95] tracking-tight mb-4"
          style={serif}
        >
          Catat lemburmu,
          <br />
          <span className="italic font-light bg-gradient-to-r from-amber-500 to-amber-300 dark:from-amber-400 dark:to-amber-200 bg-clip-text text-transparent">
            kirim ke atasan.
          </span>
        </h1>
        <p className="text-sm text-zinc-600 dark:text-zinc-400 leading-relaxed">
          Bikin Surat Perintah Lembur dengan tanda tangan digital & QR.
          Approval cepat, rekap rapi.
        </p>
      </div>

      <div className="space-y-4">
        <Field label="Jenis Pekerjaan" mono={mono}>
          <div className="grid grid-cols-2 gap-2">
            {PROFIL_OPTIONS.map((p, i) => {
              const Icon = p.icon;
              const active = profil === p.k;
              const isLastSolo =
                PROFIL_OPTIONS.length % 2 === 1 &&
                i === PROFIL_OPTIONS.length - 1;
              return (
                <button
                  key={p.k}
                  type="button"
                  onClick={() => {
                    setProfil(p.k);
                    setNip("");
                  }}
                  className={`py-3 px-3 text-sm rounded-lg border transition flex items-center gap-2 ${
                    isLastSolo ? "col-span-2 justify-center" : ""
                  } ${
                    active
                      ? "bg-amber-400 text-zinc-950 border-amber-400 font-medium shadow-md shadow-amber-500/20"
                      : "bg-white dark:bg-zinc-800 border-zinc-300 dark:border-zinc-600 text-zinc-600 dark:text-zinc-400 hover:border-amber-400/60"
                  }`}
                >
                  <Icon className="w-4 h-4 shrink-0" strokeWidth={active ? 2.5 : 2} />
                  <span className="truncate">{p.l}</span>
                </button>
              );
            })}
          </div>
        </Field>

        {usesRoster ? (
          <>
            <Field label="Pilih Karyawan" mono={mono} required>
              <select
                value={nip}
                onChange={(e) => setNip(e.target.value)}
                className="w-full bg-white dark:bg-zinc-800 border border-zinc-300 dark:border-zinc-600 rounded-lg px-4 py-3 text-sm focus:outline-none focus:border-amber-400 transition"
              >
                <option value="">— pilih nama —</option>
                {roster.map((k) => (
                  <option
                    key={k.nomorInduk}
                    value={k.nomorInduk}
                    disabled={!k.bolehLembur}
                  >
                    {k.no}. {k.nama} — {k.jabatan}
                    {!k.bolehLembur ? " (tidak dapat lembur)" : ""}
                  </option>
                ))}
              </select>
            </Field>

            {selectedKaryawan && (
              <div className="rounded-lg border border-zinc-200 dark:border-zinc-600 bg-stone-50 dark:bg-zinc-800/60 p-3 text-xs space-y-1 animate-slide-down">
                <div className="flex justify-between gap-2">
                  <span className="text-zinc-500" style={mono}>NOMOR INDUK</span>
                  <span className="font-medium" style={mono}>{selectedKaryawan.nomorInduk}</span>
                </div>
                <div className="flex justify-between gap-2">
                  <span className="text-zinc-500" style={mono}>JABATAN</span>
                  <span className="font-medium text-right">{selectedKaryawan.jabatan}</span>
                </div>
                <div className="flex justify-between gap-2">
                  <span className="text-zinc-500" style={mono}>TTL</span>
                  <span className="font-medium text-right">{selectedKaryawan.ttl}</span>
                </div>
                {!selectedKaryawan.bolehLembur && (
                  <div className="text-red-500 pt-1 border-t border-zinc-200 dark:border-zinc-700 mt-1">
                    ⚠ {selectedKaryawan.jabatan} tidak berhak mengajukan lembur.
                  </div>
                )}
              </div>
            )}
          </>
        ) : (
          <Field label="Nama" mono={mono} required>
            <input
              value={nama}
              onChange={(e) => setNama(e.target.value)}
              placeholder="cth. Sari Wulandari"
              className="w-full bg-white dark:bg-zinc-800 border border-zinc-300 dark:border-zinc-600 rounded-lg px-4 py-3 text-sm focus:outline-none focus:border-amber-400 transition"
            />
          </Field>
        )}

        <Field label="Email (opsional)" mono={mono}>
          <input
            value={email}
            onChange={(e) => setEmail(e.target.value)}
            placeholder="nama@wki.co.id"
            type="email"
            className="w-full bg-white dark:bg-zinc-800 border border-zinc-300 dark:border-zinc-600 rounded-lg px-4 py-3 text-sm focus:outline-none focus:border-amber-400 transition"
          />
        </Field>

        <button
          onClick={submit}
          disabled={!canSubmit}
          className="w-full bg-gradient-to-r from-amber-400 to-amber-500 text-zinc-950 font-medium py-3.5 rounded-xl hover:from-amber-300 hover:to-amber-400 transition-all mt-2 glow-amber disabled:opacity-40 disabled:cursor-not-allowed"
        >
          Masuk
        </button>

        <div className="flex items-center gap-3 my-2">
          <div className="flex-1 h-px bg-zinc-300 dark:bg-zinc-800" />
          <span className="text-[10px] uppercase tracking-widest text-zinc-500" style={mono}>
            atau demo cepat
          </span>
          <div className="flex-1 h-px bg-zinc-300 dark:bg-zinc-800" />
        </div>

        <div className="grid grid-cols-2 gap-2">
          {[
            { onClick: demoOrganik, label: "Organik WKI", name: "Andi Saputra", icon: Building2 },
            { onClick: demoWater, label: "Water Intake", name: "Warno", icon: Droplets },
            { onClick: demoSopir, label: "Pengemudi", name: "Dedy K.", icon: Car },
            { onClick: demoAtasan, label: "Atasan", name: "Sari W.", icon: User },
            { onClick: demoAdmin, label: "Admin Payroll", name: "Ani Pratiwi", icon: FileSpreadsheet, solo: true },
          ].map((d, i) => {
            const Icon = d.icon;
            return (
              <button
                key={i}
                onClick={d.onClick}
                className={`py-2.5 px-3 text-xs rounded-lg border border-zinc-300 dark:border-zinc-600 hover:border-amber-400 hover:bg-amber-400/5 transition text-left flex items-center gap-2 ${
                  d.solo ? "col-span-2" : ""
                }`}
              >
                <Icon className="w-3.5 h-3.5 text-amber-500 dark:text-amber-400 shrink-0" />
                <div className="min-w-0">
                  <span
                    className="block text-zinc-500 text-[9px] uppercase tracking-widest"
                    style={mono}
                  >
                    {d.label}
                  </span>
                  <span className="truncate block">{d.name}</span>
                </div>
              </button>
            );
          })}
        </div>

        <p className="text-[11px] text-center text-zinc-500" style={mono}>
          DEMO — DATA DISIMPAN DI BROWSER ANDA
        </p>
      </div>
    </div>
  );
}

// ---------- Top Bar ----------
function TopBar({ user, theme, onToggleTheme, mono, serif }) {
  return (
    <header className="flex items-center justify-between px-5 pt-5 pb-3 animate-fade-in">
      <div className="flex items-center gap-2.5">
        <div className="relative">
          <div className="w-9 h-9 rounded-lg bg-gradient-to-br from-amber-300 to-amber-500 flex items-center justify-center shadow-md shadow-amber-500/30">
            <Clock className="w-4 h-4 text-zinc-950" strokeWidth={2.5} />
          </div>
        </div>
        <div>
          <div className="text-base leading-tight" style={serif}>
            Lembur<span className="italic text-amber-500 dark:text-amber-400">.id</span>
          </div>
          <div className="text-[10px] uppercase tracking-widest text-zinc-500 flex items-center gap-1.5" style={mono}>
            <span>{user.nama}</span>
            <span className="w-1 h-1 rounded-full bg-zinc-400" />
            <span className="text-amber-600 dark:text-amber-400">{user.peran}</span>
          </div>
        </div>
      </div>
      <button
        onClick={onToggleTheme}
        aria-label="Toggle theme"
        className="w-10 h-10 rounded-full border border-zinc-200 dark:border-zinc-600 bg-white/60 dark:bg-zinc-700/50 backdrop-blur flex items-center justify-center text-zinc-600 dark:text-zinc-400 hover:border-amber-400 hover:text-amber-500 hover:rotate-12 transition-all duration-300"
      >
        {theme === "dark" ? <Sun className="w-4 h-4" /> : <Moon className="w-4 h-4" />}
      </button>
    </header>
  );
}

// ---------- Pekerja: Dashboard ----------
function Dashboard({ user, entries, gajiPokok, setView, onViewSurat, serif, mono }) {
  const m = useMemo(() => {
    const monthStart = startOfMonth();
    const inMonth = entries.filter((e) => new Date(e.tanggal) >= monthStart);
    return {
      inMonth,
      totalJam: inMonth.reduce((s, e) => s + Number(e.durasiJam || 0), 0),
      totalLibur: inMonth
        .filter((e) => e.jenis === "Hari Libur")
        .reduce((s, e) => s + Number(e.durasiJam || 0), 0),
      draft: inMonth.filter((e) => e.status === "Draft").length,
    };
  }, [entries]);

  const recent = entries.slice(0, 3);

  return (
    <div className="space-y-5 stagger">
      <div className="pt-2">
        <div
          className="text-xs uppercase tracking-[0.2em] text-amber-500 dark:text-amber-400/80"
          style={mono}
        >
          Halo, {user.nama.split(" ")[0]} 👋
        </div>
        <h2 className="text-4xl mt-1 leading-[0.95] tracking-tight" style={serif}>
          Rekap <span className="italic font-light">bulan ini</span>
        </h2>
      </div>

      <div className="relative rounded-2xl overflow-hidden bg-gradient-to-br from-amber-300 via-amber-400 to-amber-500 text-zinc-950 p-6 shadow-xl shadow-amber-500/20">
        <div
          className="absolute inset-0 opacity-10"
          style={{
            backgroundImage:
              "linear-gradient(rgba(0,0,0,.4) 1px, transparent 1px), linear-gradient(90deg, rgba(0,0,0,.4) 1px, transparent 1px)",
            backgroundSize: "32px 32px",
          }}
        />
        <div className="absolute inset-0 animate-shimmer pointer-events-none" />
        <div className="absolute -top-12 -right-12 w-40 h-40 rounded-full bg-amber-200 opacity-40 blur-2xl" />
        <div className="relative">
          <div className="flex items-center gap-1.5 text-[10px] uppercase tracking-widest mb-3" style={mono}>
            <Clock className="w-3 h-3" />
            Total Jam Lembur
          </div>
          <div className="text-5xl leading-none tracking-tight" style={serif}>
            {m.totalJam.toFixed(1)} <span className="text-2xl">jam</span>
          </div>
          <div className="flex items-center justify-between mt-6 pt-4 border-t border-zinc-900/15 text-sm">
            <div>
              <div className="text-[9px] uppercase tracking-widest opacity-60" style={mono}>ENTRI</div>
              <div className="text-lg" style={serif}>{m.inMonth.length}</div>
            </div>
            <div className="h-8 w-px bg-zinc-900/15" />
            <div>
              <div className="text-[9px] uppercase tracking-widest opacity-60" style={mono}>DRAFT</div>
              <div className="text-lg" style={serif}>{m.draft}</div>
            </div>
            <div className="h-8 w-px bg-zinc-900/15" />
            <div>
              <div className="text-[9px] uppercase tracking-widest opacity-60" style={mono}>LIBUR</div>
              <div className="text-lg" style={serif}>{(m.totalLibur ?? 0).toFixed(1)}</div>
            </div>
          </div>
        </div>
      </div>

      <button
        onClick={() => setView("input")}
        className="w-full bg-white/80 dark:bg-zinc-700/55 backdrop-blur text-zinc-900 dark:text-zinc-100 border border-zinc-200 dark:border-zinc-600 rounded-2xl p-5 flex items-center justify-between hover:border-amber-400 hover:shadow-lg hover:shadow-amber-500/10 lift group"
      >
        <div className="flex items-center gap-3">
          <div className="w-11 h-11 rounded-xl bg-gradient-to-br from-amber-400/20 to-amber-500/10 text-amber-500 dark:text-amber-400 flex items-center justify-center group-hover:scale-105 transition-transform">
            <Plus className="w-5 h-5" strokeWidth={2.5} />
          </div>
          <div className="text-left">
            <div className="text-sm font-medium">Tambah Lembur</div>
            <div className="text-xs text-zinc-500">Catat jam lembur + upload bukti</div>
          </div>
        </div>
        <div
          className="text-amber-500 dark:text-amber-400 text-xs uppercase tracking-widest flex items-center gap-1 group-hover:gap-2 transition-all"
          style={mono}
        >
          mulai
          <ArrowUpRight className="w-3.5 h-3.5 group-hover:translate-x-0.5 group-hover:-translate-y-0.5 transition-transform" />
        </div>
      </button>

      <div className="grid grid-cols-2 gap-3">
        <InfoChip
          icon={Clock}
          label="Jam libur"
          value={`${m.totalLibur.toFixed(1)} jam`}
          mono={mono}
        />
        <InfoChip
          icon={Calendar}
          label="Periode mulai"
          value={formatTanggal(startOfMonth().toISOString())}
          mono={mono}
        />
      </div>

      <div>
        <div className="flex items-center justify-between mb-3">
          <h3 className="text-lg" style={serif}>Terbaru</h3>
          <button
            onClick={() => setView("riwayat")}
            className="text-xs text-amber-500 dark:text-amber-400 uppercase tracking-widest"
            style={mono}
          >
            Lihat semua →
          </button>
        </div>
        <div className="space-y-2">
          {recent.length === 0 && (
            <div className="text-center py-10 text-sm text-zinc-500 border border-dashed border-zinc-300 dark:border-zinc-600 rounded-xl">
              Belum ada entri. Tekan "Tambah Lembur".
            </div>
          )}
          {recent.map((e) => (
            <EntryRow
              key={e.id}
              entry={e}
              onViewSurat={onViewSurat}
              mono={mono}
              serif={serif}
              compact
            />
          ))}
        </div>
      </div>
    </div>
  );
}

function InfoChip({ icon: Icon, label, value, mono }) {
  return (
    <div className="bg-white dark:bg-zinc-700/50 border border-zinc-200 dark:border-zinc-600 rounded-xl p-4">
      <div className="flex items-center gap-2 text-zinc-500">
        <Icon className="w-3.5 h-3.5" />
        <span className="text-[10px] uppercase tracking-widest" style={mono}>
          {label}
        </span>
      </div>
      <div className="mt-1.5 text-sm font-medium">{value}</div>
    </div>
  );
}

// ---------- Pekerja: Input (Surat Perintah Lembur) ----------
function InputForm({ user, gajiPokok, onSave, onCancel, serif, mono }) {
  const today = new Date().toISOString().slice(0, 10);
  const [form, setForm] = useState({
    kategori: user.kategoriDefault || "SPL",
    tipe: "Lembur",
    bagian: user.bagian || "",
    pekerjaan: "",
    tujuan: "",
    tanggalMulai: today,
    jamMulai: "18:00",
    tanggalSelesai: today,
    jamSelesai: "20:00",
    jenis: "Hari Kerja",
    hasil: "",
    keterangan: "",
    penanggungJawab: user.penanggungJawab || "",
    evidence: null,
    ttdPekerja: null,
  });
  const [showQR, setShowQR] = useState(false);
  const [showDetail, setShowDetail] = useState(false);
  const [lewatTengahMalam, setLewatTengahMalam] = useState(false);
  const isWater = form.kategori === "WaterIntake";
  const isPengemudi = form.kategori === "Pengemudi";
  const isTAD = isWater || isPengemudi;

  // Jaga tanggal selesai = tanggal mulai kalau tidak lewat tengah malam
  useEffect(() => {
    if (!lewatTengahMalam) {
      setForm((f) =>
        f.tanggalSelesai === f.tanggalMulai
          ? f
          : { ...f, tanggalSelesai: f.tanggalMulai }
      );
    }
  }, [lewatTengahMalam, form.tanggalMulai]);

  const durasi = useMemo(
    () =>
      computeDuration(
        form.tanggalMulai,
        form.jamMulai,
        form.tanggalSelesai,
        form.jamSelesai
      ),
    [form.tanggalMulai, form.jamMulai, form.tanggalSelesai, form.jamSelesai]
  );

  const preview = useMemo(
    () => hitungGaji(durasi, form.jenis, form.kategori, gajiPokok),
    [durasi, form.jenis, form.kategori, gajiPokok]
  );
  const basePerJam = useMemo(() => tarifPerJam(gajiPokok), [gajiPokok]);

  const valid =
    form.tanggalMulai &&
    form.jamMulai &&
    form.tanggalSelesai &&
    form.jamSelesai &&
    durasi > 0 &&
    form.pekerjaan.trim() &&
    (isTAD || form.tujuan.trim());

  const set = (k, v) => setForm((f) => ({ ...f, [k]: v }));

  // Helper: tambah jam ke "HH:MM", info kalau lewat hari
  const tambahJam = (jamStr, hours) => {
    const [h, m] = (jamStr || "00:00").split(":").map(Number);
    const total = h * 60 + m + hours * 60;
    const lewatHari = total >= 24 * 60;
    const mod = ((total % 1440) + 1440) % 1440;
    const nh = String(Math.floor(mod / 60)).padStart(2, "0");
    const nm = String(mod % 60).padStart(2, "0");
    return { jam: `${nh}:${nm}`, lewatHari };
  };
  const tambahHari = (iso, n) => {
    const d = new Date(`${iso}T00:00:00`);
    d.setDate(d.getDate() + n);
    return d.toISOString().slice(0, 10);
  };

  // Set jam selesai berdasar durasi cepat (dari jam mulai)
  const setDurasiCepat = (hours) => {
    const { jam, lewatHari } = tambahJam(form.jamMulai, hours);
    if (lewatHari) setLewatTengahMalam(true);
    setForm((f) => ({
      ...f,
      jamSelesai: jam,
      tanggalSelesai: lewatHari ? tambahHari(f.tanggalMulai, 1) : f.tanggalMulai,
    }));
  };

  const PRESET_MULAI = ["08:00", "17:00", "18:00", "22:00"];
  const PRESET_DURASI = [1, 2, 3, 4];

  return (
    <div className="space-y-5">
      <div className="flex items-center gap-3 pt-2">
        <button
          onClick={onCancel}
          className="w-9 h-9 rounded-full border border-zinc-300 dark:border-zinc-600 flex items-center justify-center text-zinc-500 hover:text-amber-500 hover:border-amber-400 transition"
        >
          <ArrowLeft className="w-4 h-4" />
        </button>
        <div>
          <h2 className="text-2xl leading-tight" style={serif}>
            {isPengemudi
              ? `Daftar ${form.tipe}`
              : isTAD
              ? "Daftar Lembur"
              : "Surat Perintah Lembur"}
          </h2>
          <div className="text-[10px] uppercase tracking-widest text-zinc-500 mt-0.5" style={mono}>
            {isPengemudi
              ? `Pengemudi · ${form.tipe === "Piket" ? "Shift Jaga" : "TAD"}`
              : isWater
              ? "Petugas Kebersihan · Water Intake"
              : "Penugasan Kerja Lebih"}
          </div>
        </div>
      </div>

      <div className="text-[11px] text-zinc-400 flex items-center gap-1" style={mono}>
        <span className="text-amber-500">*</span>
        <span>wajib diisi</span>
      </div>

      {/* Tipe Entri — khusus Pengemudi */}
      {isPengemudi && (
        <FormSection title="Tipe Entri" mono={mono} serif={serif}>
          <div className="grid grid-cols-2 gap-2">
            {Object.entries(TIPE_ENTRI).map(([k, v]) => {
              const Icon = v.icon;
              const active = form.tipe === k;
              return (
                <button
                  key={k}
                  type="button"
                  onClick={() => set("tipe", k)}
                  className={`py-3 px-3 text-sm rounded-lg border transition flex items-center justify-center gap-2 ${
                    active
                      ? "bg-amber-400 text-zinc-950 border-amber-400 font-medium shadow-md shadow-amber-500/20"
                      : "bg-stone-50 dark:bg-zinc-800 border-zinc-200 dark:border-zinc-600 text-zinc-600 dark:text-zinc-400 hover:border-amber-400/60"
                  }`}
                >
                  <Icon className="w-4 h-4" strokeWidth={active ? 2.5 : 2} />
                  <span>{v.label}</span>
                </button>
              );
            })}
          </div>
          <div className="text-[10px] text-zinc-500 leading-relaxed" style={mono}>
            <strong>Lembur</strong>: jam ekstra di luar jam kerja regular.{" "}
            <strong>Piket</strong>: shift jaga (mis. piket malam / akhir pekan).
          </div>
        </FormSection>
      )}

      {/* Identitas ringkas */}
      <div className="flex items-center gap-3 rounded-3xl bg-gradient-to-br from-amber-50 to-white dark:from-amber-400/10 dark:to-zinc-800/30 ring-1 ring-amber-200/50 dark:ring-white/5 p-4">
        <div
          className="w-12 h-12 rounded-2xl bg-gradient-to-br from-amber-300 to-amber-500 text-zinc-950 flex items-center justify-center text-xl shadow-md shadow-amber-500/25 shrink-0"
          style={serif}
        >
          {user.nama.charAt(0).toUpperCase()}
        </div>
        <div className="min-w-0">
          <div className="font-semibold truncate">{user.nama}</div>
          <div className="text-xs text-zinc-500 truncate" style={mono}>
            {user.jabatan || user.pekerjaan || "Pekerja"}
            {user.nomorInduk ? ` · ${user.nomorInduk}` : ""}
          </div>
        </div>
      </div>

      {/* Pekerjaan */}
      <FormSection title="Pekerjaan Lembur" mono={mono} serif={serif}>
        <Field label="Uraian pekerjaan" mono={mono} required>
          <input
            value={form.pekerjaan}
            onChange={(e) => set("pekerjaan", e.target.value)}
            placeholder={
              isPengemudi
                ? "cth. Antar dinas ke Bandung"
                : isWater
                ? "cth. Ngantiin cuti Aries Munadar"
                : "cth. Pemasangan kabel jaringan lantai 3"
            }
            className={inputCls}
          />
        </Field>
        {!isTAD && (
          <Field label="Tujuan" mono={mono} required>
            <textarea
              value={form.tujuan}
              onChange={(e) => set("tujuan", e.target.value)}
              placeholder="cth. Persiapan ruang meeting baru"
              rows={2}
              className={inputCls + " resize-none"}
            />
          </Field>
        )}
      </FormSection>

      {/* Waktu Lembur */}
      <FormSection title="Waktu Lembur" mono={mono} serif={serif}>
        <Field label="Tanggal" mono={mono} required>
          <input
            type="date"
            value={form.tanggalMulai}
            onChange={(e) => set("tanggalMulai", e.target.value)}
            className={inputCls}
            style={mono}
          />
        </Field>

        {/* Preset jam mulai */}
        <div>
          <div className="block text-[12px] font-medium tracking-wide text-zinc-400 dark:text-zinc-500 mb-1.5">
            Jam mulai cepat
          </div>
          <div className="grid grid-cols-4 gap-2">
            {PRESET_MULAI.map((t) => (
              <button
                key={t}
                type="button"
                onClick={() => set("jamMulai", t)}
                className={`py-2.5 text-sm rounded-xl transition ${
                  form.jamMulai === t
                    ? "bg-gradient-to-br from-amber-300 to-amber-500 text-zinc-950 font-semibold shadow-sm shadow-amber-500/20"
                    : "bg-zinc-50 dark:bg-zinc-900/40 ring-1 ring-zinc-200/70 dark:ring-white/5 text-zinc-500 hover:ring-amber-400/50"
                }`}
                style={mono}
              >
                {t}
              </button>
            ))}
          </div>
        </div>

        <div className="grid grid-cols-2 gap-3">
          <Field label="Jam Mulai" mono={mono} required>
            <input
              type="time"
              value={form.jamMulai}
              onChange={(e) => set("jamMulai", e.target.value)}
              className={inputCls}
              style={mono}
            />
          </Field>
          <Field label="Jam Selesai" mono={mono} required>
            <input
              type="time"
              value={form.jamSelesai}
              onChange={(e) => set("jamSelesai", e.target.value)}
              className={inputCls}
              style={mono}
            />
          </Field>
        </div>

        {/* Durasi cepat */}
        <div>
          <div className="block text-[12px] font-medium tracking-wide text-zinc-400 dark:text-zinc-500 mb-1.5">
            Durasi cepat <span className="text-zinc-300 dark:text-zinc-600">(dari jam mulai)</span>
          </div>
          <div className="grid grid-cols-4 gap-2">
            {PRESET_DURASI.map((h) => {
              const aktif = Math.round(durasi) === h && durasi > 0;
              return (
                <button
                  key={h}
                  type="button"
                  onClick={() => setDurasiCepat(h)}
                  className={`py-2.5 text-sm rounded-xl transition flex items-center justify-center gap-1 ${
                    aktif
                      ? "bg-gradient-to-br from-amber-300 to-amber-500 text-zinc-950 font-semibold shadow-sm shadow-amber-500/20"
                      : "bg-zinc-50 dark:bg-zinc-900/40 ring-1 ring-zinc-200/70 dark:ring-white/5 text-zinc-500 hover:ring-amber-400/50"
                  }`}
                  style={mono}
                >
                  +{h}j
                </button>
              );
            })}
          </div>
        </div>

        {/* Toggle lewat tengah malam */}
        <button
          type="button"
          onClick={() => setLewatTengahMalam((v) => !v)}
          className="flex items-center gap-2 text-xs text-zinc-500 hover:text-amber-500 transition"
        >
          <span
            className={`w-9 h-5 rounded-full transition relative ${
              lewatTengahMalam ? "bg-amber-400" : "bg-zinc-300 dark:bg-zinc-600"
            }`}
          >
            <span
              className={`absolute top-0.5 w-4 h-4 bg-white rounded-full transition-all ${
                lewatTengahMalam ? "left-4.5" : "left-0.5"
              }`}
              style={{ left: lewatTengahMalam ? "18px" : "2px" }}
            />
          </span>
          <span className="text-[12px] font-medium tracking-wide">
            Lewat tengah malam (selesai beda hari)
          </span>
        </button>

        {lewatTengahMalam && (
          <Field label="Tanggal Selesai" mono={mono} required>
            <input
              type="date"
              value={form.tanggalSelesai}
              onChange={(e) => set("tanggalSelesai", e.target.value)}
              className={inputCls}
              style={mono}
            />
          </Field>
        )}

        <Field label="Jenis Hari" mono={mono}>
          <div className="grid grid-cols-2 gap-2">
            {["Hari Kerja", "Hari Libur"].map((j) => (
              <button
                key={j}
                type="button"
                onClick={() => set("jenis", j)}
                className={`py-3.5 text-sm rounded-xl transition ${
                  form.jenis === j
                    ? "bg-gradient-to-br from-amber-300 to-amber-500 text-zinc-950 font-semibold shadow-sm shadow-amber-500/20"
                    : "bg-zinc-50 dark:bg-zinc-900/40 ring-1 ring-zinc-200/70 dark:ring-white/5 text-zinc-500"
                }`}
              >
                {j}
                <div
                  className={`text-[10px] mt-0.5 ${
                    form.jenis === j ? "text-zinc-900/70" : "text-zinc-400"
                  }`}
                  style={mono}
                >
                  ×{MULTIPLIER[j]}
                </div>
              </button>
            ))}
          </div>
        </Field>

        <div className="bg-zinc-50 dark:bg-zinc-900/40 rounded-xl p-3.5 flex items-center justify-between">
          <span className="text-[12px] font-medium tracking-wide text-zinc-400 dark:text-zinc-500">
            Total durasi
          </span>
          <span className="text-lg" style={serif}>
            {durasi > 0 ? (
              <>
                {durasi.toFixed(2)}
                <span className="text-[10px] text-zinc-500 ml-1" style={mono}>
                  jam
                </span>
              </>
            ) : (
              <span className="text-xs text-red-500" style={mono}>
                JAM SELESAI HARUS &gt; JAM MULAI
              </span>
            )}
          </span>
        </div>
      </FormSection>

      {/* Total Jam — hero */}
      <div className="relative rounded-3xl overflow-hidden bg-gradient-to-br from-amber-300 via-amber-400 to-amber-500 text-zinc-950 p-5 shadow-lg shadow-amber-500/25">
        <div className="absolute -top-10 -right-10 w-32 h-32 rounded-full bg-amber-200 opacity-40 blur-2xl" />
        <div className="relative">
          <div className="flex items-center gap-1.5 text-[11px] font-medium tracking-wide opacity-80" style={mono}>
            <Clock className="w-3.5 h-3.5" />
            Total Jam Lembur
          </div>
          <div className="text-4xl mt-1 tracking-tight" style={serif}>
            {durasi.toFixed(2)} <span className="text-xl">jam</span>
          </div>
          <div className="text-[11px] mt-2 opacity-70" style={mono}>
            {KATEGORI[form.kategori]?.short} · {form.jenis}
          </div>
        </div>
      </div>

      {/* Detail tambahan — collapsible */}
      <div className="bg-white dark:bg-zinc-700/50 border border-zinc-200 dark:border-zinc-600 rounded-2xl overflow-hidden">
        <button
          type="button"
          onClick={() => setShowDetail((v) => !v)}
          className="w-full p-4 flex items-center justify-between hover:bg-amber-400/5 transition"
        >
          <div className="flex items-center gap-2">
            <h3 className="text-base" style={serif}>Detail Tambahan</h3>
            <span className="text-[10px] uppercase tracking-widest text-zinc-500" style={mono}>
              opsional
            </span>
          </div>
          <ChevronRight
            className={`w-4 h-4 text-zinc-400 transition-transform ${
              showDetail ? "rotate-90" : ""
            }`}
          />
        </button>

        {showDetail && (
          <div className="px-5 pb-5 space-y-4 border-t border-zinc-200 dark:border-zinc-600 pt-4">
            {!isTAD && (
              <Field label="Bagian" mono={mono}>
                <input
                  value={form.bagian}
                  onChange={(e) => set("bagian", e.target.value)}
                  placeholder="cth. IT Support"
                  className={inputCls}
                />
              </Field>
            )}
            {!isTAD && (
              <Field label="Hasil" mono={mono}>
                <textarea
                  value={form.hasil}
                  onChange={(e) => set("hasil", e.target.value)}
                  placeholder="cth. 12 titik kabel terpasang, semua aktif"
                  rows={2}
                  className={inputCls + " resize-none"}
                />
              </Field>
            )}
            <Field label="Keterangan" mono={mono}>
              <textarea
                value={form.keterangan}
                onChange={(e) => set("keterangan", e.target.value)}
                placeholder={isTAD ? "cth. OK / izin / sakit" : "catatan tambahan"}
                rows={2}
                className={inputCls + " resize-none"}
              />
            </Field>
            <Field
              label={isTAD ? "Koordinator Regu" : "Penanggung Jawab"}
              mono={mono}
            >
              <input
                value={form.penanggungJawab}
                onChange={(e) => set("penanggungJawab", e.target.value)}
                placeholder={
                  isPengemudi
                    ? "cth. Wahyudi (Kepala Pengemudi)"
                    : isWater
                    ? "cth. Earlan S."
                    : "cth. Rudi (Supervisor IT)"
                }
                className={inputCls}
              />
            </Field>
            <Field label="Bukti (foto / PDF)" mono={mono}>
              <EvidenceUploader
                value={form.evidence}
                onChange={(ev) => set("evidence", ev)}
                mono={mono}
              />
            </Field>

            <div className="pt-1">
              <div className="flex items-center justify-between mb-1.5">
                <div
                  className="text-[10px] uppercase tracking-widest text-zinc-500 flex items-center gap-1.5"
                  style={mono}
                >
                  <PenLine className="w-3 h-3 text-amber-500" />
                  Tanda Tangan Pekerja
                </div>
                <button
                  type="button"
                  onClick={() => setShowQR((v) => !v)}
                  className="text-[10px] uppercase tracking-widest flex items-center gap-1 text-zinc-500 hover:text-amber-500 transition"
                  style={mono}
                >
                  <QrCode className="w-3 h-3" />
                  {showQR ? "Sembunyikan QR" : "Lihat QR"}
                </button>
              </div>
              <SignaturePad
                value={form.ttdPekerja}
                onChange={(d) => set("ttdPekerja", d)}
                mono={mono}
              />
            </div>

            {showQR && (
              <div className="flex items-center gap-3 bg-stone-50 dark:bg-zinc-800 border border-zinc-200 dark:border-zinc-600 rounded-lg p-3">
                <div className="bg-white p-2 rounded shrink-0">
                  <QRCodeView
                    data={buildQRPayload(
                      {
                        id: "preview-" + Date.now().toString(36),
                        userNama: user.nama,
                        ...form,
                        durasiJam: durasi,
                        estimasiGaji: preview,
                        status: "Draft",
                      },
                      { inisial: "WKI" }
                    )}
                    size={90}
                  />
                </div>
                <div className="text-[11px] text-zinc-500 leading-relaxed" style={mono}>
                  QR tercetak pada Surat Perintah berisi ringkasan data lembur
                  untuk verifikasi cepat oleh atasan.
                </div>
              </div>
            )}
          </div>
        )}
      </div>

      <button
        disabled={!valid}
        onClick={() => onSave(form)}
        className="w-full bg-amber-400 text-zinc-950 font-medium py-3.5 rounded-lg hover:bg-amber-300 disabled:opacity-40 disabled:cursor-not-allowed transition flex items-center justify-center gap-2 sticky bottom-24 shadow-lg shadow-amber-500/20"
      >
        <Plus className="w-4 h-4" strokeWidth={2.5} />
        Simpan Surat Perintah
      </button>
    </div>
  );
}

const inputCls =
  "w-full bg-zinc-50 dark:bg-zinc-900/40 border border-transparent ring-1 ring-zinc-200/70 dark:ring-white/5 rounded-xl px-4 py-3.5 text-[15px] placeholder:text-zinc-400 focus:outline-none focus:ring-amber-400 focus:border-amber-400 focus:bg-white dark:focus:bg-zinc-900 transition";

function FormSection({ title, children, mono, serif }) {
  return (
    <div className="bg-white dark:bg-zinc-800/40 ring-1 ring-zinc-100 dark:ring-white/5 shadow-sm shadow-zinc-200/50 dark:shadow-black/20 rounded-3xl p-5 space-y-4">
      {title && (
        <h3 className="text-[13px] font-semibold tracking-wide text-zinc-400 dark:text-zinc-500">
          {title}
        </h3>
      )}
      {children}
    </div>
  );
}

function ReadOnlyValue({ value, mono }) {
  return (
    <div
      className="w-full bg-stone-100 dark:bg-zinc-800 border border-zinc-200 dark:border-zinc-600 rounded-lg px-4 py-3 text-sm text-zinc-600 dark:text-zinc-400"
      style={mono}
    >
      {value}
    </div>
  );
}

// ---------- Evidence Uploader ----------
const MAX_EVIDENCE_BYTES = 1_500_000; // ~1.5MB

function EvidenceUploader({ value, onChange, mono }) {
  const inputRef = useRef(null);
  const [busy, setBusy] = useState(false);

  const pick = () => inputRef.current?.click();

  const handle = (e) => {
    const file = e.target.files?.[0];
    e.target.value = "";
    if (!file) return;
    if (file.size > MAX_EVIDENCE_BYTES) {
      alert(
        `Berkas terlalu besar (${(file.size / 1024 / 1024).toFixed(
          1
        )}MB). Maks 1.5MB. Kompres foto dulu, ya.`
      );
      return;
    }
    setBusy(true);
    const reader = new FileReader();
    reader.onload = () => {
      onChange({
        name: file.name,
        type: file.type,
        size: file.size,
        dataUrl: reader.result,
      });
      setBusy(false);
    };
    reader.onerror = () => {
      alert("Gagal membaca berkas.");
      setBusy(false);
    };
    reader.readAsDataURL(file);
  };

  const remove = () => onChange(null);

  if (value) {
    const isImg = value.type?.startsWith("image/");
    return (
      <div className="flex items-center gap-3 bg-stone-50 dark:bg-zinc-800 border border-zinc-200 dark:border-zinc-600 rounded-lg p-3">
        <div className="w-12 h-12 rounded-md overflow-hidden bg-zinc-200 dark:bg-zinc-800 flex items-center justify-center shrink-0">
          {isImg ? (
            <img src={value.dataUrl} alt="" className="w-full h-full object-cover" />
          ) : (
            <FileText className="w-5 h-5 text-zinc-500" />
          )}
        </div>
        <div className="flex-1 min-w-0">
          <div className="text-sm truncate">{value.name}</div>
          <div className="text-[10px] text-zinc-500" style={mono}>
            {(value.size / 1024).toFixed(0)} KB · {value.type || "file"}
          </div>
        </div>
        <button
          onClick={remove}
          className="p-1.5 text-zinc-500 hover:text-red-500 transition"
          aria-label="Hapus bukti"
        >
          <X className="w-4 h-4" />
        </button>
      </div>
    );
  }

  return (
    <>
      <input
        ref={inputRef}
        type="file"
        accept="image/*,application/pdf"
        onChange={handle}
        className="hidden"
      />
      <button
        type="button"
        onClick={pick}
        disabled={busy}
        className="w-full flex items-center justify-center gap-2 py-3 rounded-lg border-2 border-dashed border-zinc-300 dark:border-zinc-600 text-zinc-500 hover:border-amber-400 hover:text-amber-500 transition"
      >
        <Paperclip className="w-4 h-4" />
        {busy ? "Memuat…" : "Pilih foto / PDF (maks 1.5 MB)"}
      </button>
    </>
  );
}

// ---------- Evidence Viewer ----------
function EvidenceViewer({ evidence, onClose, mono, serif }) {
  const isImg = evidence.type?.startsWith("image/");
  return (
    <div
      className="fixed inset-0 z-50 bg-black/80 backdrop-blur-sm flex items-center justify-center p-4"
      onClick={onClose}
    >
      <div
        className="bg-zinc-900 border border-zinc-800 rounded-2xl max-w-md w-full max-h-[90vh] overflow-hidden flex flex-col"
        onClick={(e) => e.stopPropagation()}
      >
        <div className="flex items-center justify-between p-4 border-b border-zinc-800">
          <div className="min-w-0">
            <div className="text-sm text-zinc-100 truncate" style={serif}>
              {evidence.name}
            </div>
            <div className="text-[10px] uppercase tracking-widest text-zinc-500" style={mono}>
              Bukti Lembur · {(evidence.size / 1024).toFixed(0)} KB
            </div>
          </div>
          <button
            onClick={onClose}
            className="w-8 h-8 rounded-full border border-zinc-700 flex items-center justify-center text-zinc-400 hover:text-amber-400 hover:border-amber-400 transition shrink-0"
          >
            <X className="w-4 h-4" />
          </button>
        </div>
        <div className="flex-1 overflow-auto bg-zinc-950 flex items-center justify-center p-4">
          {isImg ? (
            <img src={evidence.dataUrl} alt={evidence.name} className="max-w-full max-h-full" />
          ) : (
            <div className="text-center text-zinc-400">
              <FileText className="w-12 h-12 mx-auto mb-3 text-zinc-600" />
              <div className="text-sm mb-4">Pratinjau tidak tersedia</div>
              <a
                href={evidence.dataUrl}
                download={evidence.name}
                className="inline-flex items-center gap-2 bg-amber-400 text-zinc-950 px-4 py-2.5 rounded-lg font-medium hover:bg-amber-300 transition"
              >
                <Download className="w-4 h-4" />
                Unduh berkas
              </a>
            </div>
          )}
        </div>
        {isImg && (
          <div className="p-3 border-t border-zinc-800 text-center">
            <a
              href={evidence.dataUrl}
              download={evidence.name}
              className="inline-flex items-center gap-2 text-xs text-amber-400 hover:underline"
            >
              <Download className="w-3.5 h-3.5" />
              Unduh
            </a>
          </div>
        )}
      </div>
    </div>
  );
}

// ---------- Pekerja: Riwayat ----------
function Riwayat({ entries, onDelete, onKirim, onViewEvidence, onViewSurat, serif, mono }) {
  const [filter, setFilter] = useState("bulan");
  const [tipeFilter, setTipeFilter] = useState("Semua"); // Semua | Lembur | Piket

  const hasPiket = useMemo(
    () => entries.some((e) => e.tipe === "Piket"),
    [entries]
  );

  const filtered = useMemo(() => {
    let list = entries;
    if (filter !== "semua") {
      const cutoff = filter === "minggu" ? startOfWeek() : startOfMonth();
      list = list.filter((e) => new Date(e.tanggalMulai || e.tanggal) >= cutoff);
    }
    if (tipeFilter !== "Semua") {
      list = list.filter((e) => (e.tipe || "Lembur") === tipeFilter);
    }
    return list;
  }, [entries, filter, tipeFilter]);

  const totals = useMemo(
    () => ({
      jam: filtered.reduce((s, e) => s + Number(e.durasiJam || 0), 0),
      entri: filtered.length,
      draft: filtered.filter((e) => e.status === "Draft").length,
    }),
    [filtered]
  );

  return (
    <div className="space-y-5">
      <div className="pt-2">
        <h2 className="text-3xl leading-tight" style={serif}>Riwayat</h2>
        <p className="text-sm text-zinc-500 mt-1">Pantau dan kirim lembur ke atasan.</p>
      </div>

      <div className="flex gap-2 text-xs" style={mono}>
        {[
          { k: "minggu", l: "MINGGU INI" },
          { k: "bulan", l: "BULAN INI" },
          { k: "semua", l: "SEMUA" },
        ].map((f) => (
          <button
            key={f.k}
            onClick={() => setFilter(f.k)}
            className={`px-3 py-1.5 rounded-full uppercase tracking-widest border transition ${
              filter === f.k
                ? "bg-amber-400 text-zinc-950 border-amber-400"
                : "border-zinc-300 dark:border-zinc-600 text-zinc-500"
            }`}
          >
            {f.l}
          </button>
        ))}
      </div>

      {hasPiket && (
        <div className="flex gap-2 text-xs" style={mono}>
          {["Semua", "Lembur", "Piket"].map((t) => {
            const Icon =
              t === "Lembur" ? Clock : t === "Piket" ? CalendarClock : null;
            return (
              <button
                key={t}
                onClick={() => setTipeFilter(t)}
                className={`px-3 py-1.5 rounded-full uppercase tracking-widest border transition flex items-center gap-1.5 ${
                  tipeFilter === t
                    ? t === "Piket"
                      ? "bg-rose-500 text-white border-rose-500"
                      : "bg-zinc-900 dark:bg-zinc-100 text-amber-400 dark:text-zinc-900 border-zinc-900 dark:border-zinc-100"
                    : "border-zinc-300 dark:border-zinc-600 text-zinc-500"
                }`}
              >
                {Icon && <Icon className="w-3 h-3" />}
                {t}
              </button>
            );
          })}
        </div>
      )}

      <div className="grid grid-cols-3 gap-2">
        <Stat label="Total Jam" value={totals.jam.toFixed(1)} mono={mono} serif={serif} />
        <Stat label="Entri" value={totals.entri} mono={mono} serif={serif} small />
        <Stat label="Draft" value={totals.draft} mono={mono} serif={serif} highlight />
      </div>

      {totals.draft > 0 && (
        <button
          onClick={onKirim}
          className="w-full bg-zinc-900 dark:bg-zinc-100 text-zinc-100 dark:text-zinc-900 font-medium py-3.5 rounded-lg hover:opacity-90 transition flex items-center justify-center gap-2"
        >
          <Send className="w-4 h-4" />
          Kirim {totals.draft} Draft ke Atasan
        </button>
      )}

      <div className="space-y-2">
        {filtered.length === 0 && (
          <div className="text-center py-12 text-sm text-zinc-500 border border-dashed border-zinc-300 dark:border-zinc-600 rounded-xl">
            Tidak ada entri pada periode ini.
          </div>
        )}
        {filtered.map((e) => (
          <EntryRow
            key={e.id}
            entry={e}
            onDelete={() => onDelete(e.id)}
            onViewEvidence={onViewEvidence}
            onViewSurat={onViewSurat}
            mono={mono}
            serif={serif}
          />
        ))}
      </div>
    </div>
  );
}

function Stat({ label, value, mono, serif, highlight, small }) {
  return (
    <div
      className={`rounded-xl p-3 border ${
        highlight
          ? "bg-amber-400/10 border-amber-400/30"
          : "bg-white dark:bg-zinc-700/50 border-zinc-200 dark:border-zinc-600"
      }`}
    >
      <div className="text-[9px] uppercase tracking-widest text-zinc-500" style={mono}>
        {label}
      </div>
      <div
        className={`${small ? "text-base" : "text-xl"} mt-0.5 ${
          highlight ? "text-amber-600 dark:text-amber-400" : ""
        }`}
        style={serif}
      >
        {value}
      </div>
    </div>
  );
}

function EntryRow({
  entry,
  onDelete,
  onViewEvidence,
  onViewSurat,
  mono,
  serif,
  compact,
  showUser,
  onApprove,
}) {
  const s = STATUS[entry.status] || STATUS.Draft;
  const Icon = s.icon;
  const ev = entry.evidence;
  const tanggalTampil = entry.tanggalMulai || entry.tanggal;
  return (
    <div className="flex items-center gap-3 p-3 rounded-xl bg-white dark:bg-zinc-700/50 border border-zinc-200 dark:border-zinc-600">
      <button
        onClick={onViewSurat ? () => onViewSurat(entry) : undefined}
        className={`flex-1 min-w-0 flex items-center gap-3 text-left ${
          onViewSurat ? "cursor-pointer" : "cursor-default"
        }`}
        disabled={!onViewSurat}
      >
        <div
          className={`w-1 h-12 rounded-full shrink-0 ${
            entry.jenis === "Hari Libur" ? "bg-amber-400" : "bg-zinc-400 dark:bg-zinc-700"
          }`}
        />
        <div className="flex-1 min-w-0">
          <div className="flex items-center gap-2 flex-wrap">
            {showUser && (
              <span className="text-xs font-medium truncate">{entry.userNama}</span>
            )}
            <span className="text-[11px] text-zinc-500" style={mono}>
              {formatTanggal(tanggalTampil)}
            </span>
            {entry.jamMulai && entry.jamSelesai && (
              <span className="text-[11px] text-zinc-500" style={mono}>
                {entry.jamMulai}–{entry.jamSelesai}
              </span>
            )}
            <span
              className={`text-[9px] px-1.5 py-0.5 rounded uppercase tracking-widest flex items-center gap-1 ${s.cls}`}
              style={mono}
            >
              <Icon className="w-2.5 h-2.5" />
              {s.label}
            </span>
            {entry.kategori === "WaterIntake" && (
              <span
                className="text-[9px] px-1.5 py-0.5 rounded uppercase tracking-widest flex items-center gap-1 bg-sky-400/15 text-sky-600 dark:text-sky-300"
                style={mono}
              >
                <Droplets className="w-2.5 h-2.5" />
                water intake
              </span>
            )}
            {entry.kategori === "Pengemudi" && (
              <span
                className="text-[9px] px-1.5 py-0.5 rounded uppercase tracking-widest flex items-center gap-1 bg-violet-400/15 text-violet-600 dark:text-violet-300"
                style={mono}
              >
                <Car className="w-2.5 h-2.5" />
                pengemudi
              </span>
            )}
            {entry.tipe === "Piket" && (
              <span
                className="text-[9px] px-1.5 py-0.5 rounded uppercase tracking-widest flex items-center gap-1 bg-rose-400/15 text-rose-600 dark:text-rose-300 font-medium"
                style={mono}
              >
                <CalendarClock className="w-2.5 h-2.5" />
                piket
              </span>
            )}
            {ev && (
              <span
                className="text-[9px] px-1.5 py-0.5 rounded uppercase tracking-widest flex items-center gap-1 bg-amber-400/15 text-amber-600 dark:text-amber-300"
                style={mono}
              >
                {ev.type?.startsWith("image/") ? (
                  <ImageIcon className="w-2.5 h-2.5" />
                ) : (
                  <FileText className="w-2.5 h-2.5" />
                )}
                bukti
              </span>
            )}
          </div>
          {(entry.pekerjaan || entry.deskripsi) && (
            <div className="text-sm truncate mt-0.5">
              {entry.pekerjaan || entry.deskripsi}
            </div>
          )}
          {!compact && entry.bagian && (
            <div className="text-[11px] text-zinc-500 mt-0.5" style={mono}>
              {entry.bagian.toUpperCase()} · {entry.jenis.toUpperCase()} · ×
              {MULTIPLIER[entry.jenis]}
            </div>
          )}
        </div>
        <div className="text-right shrink-0">
          <div className="text-lg text-amber-600 dark:text-amber-400" style={serif}>
            {Number(entry.durasiJam).toFixed(1)}
            <span className="text-[10px] text-zinc-500 ml-0.5" style={mono}>
              jam
            </span>
          </div>
          <div className="text-[9px] uppercase tracking-widest text-zinc-400" style={mono}>
            {entry.jenis === "Hari Libur" ? "libur" : "kerja"}
          </div>
        </div>
      </button>

      <div className="flex items-center shrink-0">
        {ev && onViewEvidence && (
          <button
            onClick={() => onViewEvidence(ev)}
            className="p-1.5 text-zinc-400 hover:text-amber-500 transition"
            aria-label="Lihat bukti"
            title="Lihat bukti"
          >
            <Paperclip className="w-3.5 h-3.5" />
          </button>
        )}
        {onApprove && entry.status === "Terkirim" && (
          <button
            onClick={() => onApprove(entry)}
            className="p-1.5 text-emerald-500 hover:bg-emerald-400/10 rounded transition"
            aria-label="Setujui"
            title="Setujui"
          >
            <Check className="w-4 h-4" strokeWidth={2.5} />
          </button>
        )}
        {onDelete && entry.status === "Draft" && (
          <button
            onClick={onDelete}
            className="p-1.5 text-zinc-400 hover:text-red-500 transition"
            aria-label="Hapus"
          >
            <Trash2 className="w-3.5 h-3.5" />
          </button>
        )}
      </div>
    </div>
  );
}

// ---------- Admin: Dashboard ----------
function AdminDashboard({ user, entries, setView, onViewSurat, serif, mono }) {
  const m = useMemo(() => {
    const monthStart = startOfMonth();
    const inMonth = entries.filter(
      (e) => new Date(e.tanggalMulai || e.tanggal) >= monthStart
    );
    const totalJam = inMonth.reduce((s, e) => s + Number(e.durasiJam || 0), 0);
    const totalGaji = inMonth.reduce((s, e) => s + Number(e.estimasiGaji || 0), 0);
    const pending = inMonth.filter((e) => e.status === "Terkirim").length;
    const employees = new Set(inMonth.map((e) => e.userId)).size;
    return { inMonth, totalJam, totalGaji, pending, employees };
  }, [entries]);

  const pendingRecent = useMemo(
    () =>
      entries
        .filter((e) => e.status === "Terkirim")
        .sort(
          (a, b) =>
            new Date(b.tanggalMulai || b.tanggal) -
            new Date(a.tanggalMulai || a.tanggal)
        )
        .slice(0, 3),
    [entries]
  );

  return (
    <div className="space-y-5">
      <div className="pt-2">
        <div
          className="text-xs uppercase tracking-[0.2em] text-amber-500 dark:text-amber-400/80"
          style={mono}
        >
          Atasan · {user.nama.split(" ")[0]}
        </div>
        <h2 className="text-3xl mt-1 leading-tight" style={serif}>
          Rekap tim bulan ini
        </h2>
      </div>

      <div className="relative rounded-2xl overflow-hidden bg-gradient-to-br from-amber-300 via-amber-400 to-amber-500 text-zinc-950 p-6 shadow-xl shadow-amber-500/20">
        <div
          className="absolute inset-0 opacity-10"
          style={{
            backgroundImage:
              "linear-gradient(rgba(0,0,0,.4) 1px, transparent 1px), linear-gradient(90deg, rgba(0,0,0,.4) 1px, transparent 1px)",
            backgroundSize: "32px 32px",
          }}
        />
        <div className="absolute inset-0 animate-shimmer pointer-events-none" />
        <div className="absolute -top-12 -right-12 w-40 h-40 rounded-full bg-amber-200 opacity-40 blur-2xl" />
        <div className="relative">
          <div className="flex items-center gap-1.5 text-[10px] uppercase tracking-widest mb-3" style={mono}>
            <Users className="w-3 h-3" />
            Total Jam Lembur Tim
          </div>
          <div className="text-5xl leading-none tracking-tight" style={serif}>
            {m.totalJam.toFixed(1)} <span className="text-2xl">jam</span>
          </div>
          <div className="grid grid-cols-3 gap-2 mt-6 pt-4 border-t border-zinc-900/15 text-sm">
            <div>
              <div className="opacity-60 text-[9px] uppercase tracking-widest" style={mono}>ENTRI</div>
              <div className="text-base" style={serif}>{m.inMonth.length}</div>
            </div>
            <div>
              <div className="opacity-60 text-[9px] uppercase tracking-widest" style={mono}>ORANG</div>
              <div className="text-base" style={serif}>{m.employees}</div>
            </div>
            <div>
              <div className="opacity-60 text-[9px] uppercase tracking-widest" style={mono}>PENDING</div>
              <div className="text-base" style={serif}>{m.pending}</div>
            </div>
          </div>
        </div>
      </div>

      <button
        onClick={() => setView("rekap")}
        className="w-full bg-white dark:bg-zinc-800 border border-zinc-200 dark:border-zinc-600 rounded-2xl p-5 flex items-center justify-between hover:border-amber-400 transition"
      >
        <div className="flex items-center gap-3">
          <div className="w-10 h-10 rounded-lg bg-amber-400/15 text-amber-500 dark:text-amber-400 flex items-center justify-center">
            <ClipboardList className="w-5 h-5" strokeWidth={2.5} />
          </div>
          <div className="text-left">
            <div className="text-sm font-medium">Compile Rekap Karyawan</div>
            <div className="text-xs text-zinc-500">
              Filter periode, setujui, dan export CSV
            </div>
          </div>
        </div>
        <ChevronRight className="w-5 h-5 text-zinc-400" />
      </button>

      <div>
        <div className="flex items-center justify-between mb-3">
          <h3 className="text-lg" style={serif}>
            Menunggu persetujuan
          </h3>
          <button
            onClick={() => setView("rekap")}
            className="text-xs text-amber-500 dark:text-amber-400 uppercase tracking-widest"
            style={mono}
          >
            Lihat semua →
          </button>
        </div>
        <div className="space-y-2">
          {pendingRecent.length === 0 && (
            <div className="text-center py-10 text-sm text-zinc-500 border border-dashed border-zinc-300 dark:border-zinc-600 rounded-xl">
              Tidak ada entri menunggu persetujuan.
            </div>
          )}
          {pendingRecent.map((e) => (
            <EntryRow
              key={e.id}
              entry={e}
              onViewSurat={onViewSurat}
              mono={mono}
              serif={serif}
              compact
              showUser
            />
          ))}
        </div>
      </div>
    </div>
  );
}

// ---------- Admin: Rekap ----------
function AdminRekap({
  entries,
  onApprove,
  onApproveBatch,
  onViewEvidence,
  onViewSurat,
  serif,
  mono,
}) {
  const [filter, setFilter] = useState("bulan");
  const [statusFilter, setStatusFilter] = useState("Semua");
  const [expanded, setExpanded] = useState(null);

  const filtered = useMemo(() => {
    let list = entries;
    if (filter !== "semua") {
      const cutoff = filter === "minggu" ? startOfWeek() : startOfMonth();
      list = list.filter(
        (e) => new Date(e.tanggalMulai || e.tanggal) >= cutoff
      );
    }
    if (statusFilter !== "Semua") {
      list = list.filter((e) => e.status === statusFilter);
    }
    return list;
  }, [entries, filter, statusFilter]);

  // Group by user
  const grouped = useMemo(() => {
    const map = new Map();
    for (const e of filtered) {
      const k = e.userId;
      if (!map.has(k)) {
        map.set(k, {
          userId: e.userId,
          userNama: e.userNama,
          jam: 0,
          gaji: 0,
          entri: [],
          draft: 0,
          terkirim: 0,
          disetujui: 0,
        });
      }
      const g = map.get(k);
      g.jam += Number(e.durasiJam || 0);
      g.gaji += Number(e.estimasiGaji || 0);
      g.entri.push(e);
      if (e.status === "Draft") g.draft++;
      if (e.status === "Terkirim") g.terkirim++;
      if (e.status === "Disetujui") g.disetujui++;
    }
    for (const g of map.values()) {
      g.entri.sort(
        (a, b) =>
          new Date(b.tanggalMulai || b.tanggal) -
          new Date(a.tanggalMulai || a.tanggal)
      );
    }
    return [...map.values()].sort((a, b) => b.gaji - a.gaji);
  }, [filtered]);

  const totals = useMemo(
    () => ({
      jam: filtered.reduce((s, e) => s + Number(e.durasiJam || 0), 0),
      gaji: filtered.reduce((s, e) => s + Number(e.estimasiGaji || 0), 0),
      pending: filtered.filter((e) => e.status === "Terkirim").length,
      orang: grouped.length,
    }),
    [filtered, grouped]
  );

  const exportCSV = () => {
    const headers = [
      "Kategori",
      "Tipe",
      "Nama",
      "Bagian",
      "Pekerjaan",
      "Tujuan",
      "Tanggal Mulai",
      "Jam Mulai",
      "Tanggal Selesai",
      "Jam Selesai",
      "Durasi Jam",
      "Jenis",
      "Hasil",
      "Keterangan",
      "Penanggung Jawab",
      "Status",
      "Bukti",
    ];
    const clean = (v) => String(v ?? "").replace(/\n/g, " ");
    const rows = filtered
      .slice()
      .sort((a, b) => {
        if (a.userNama !== b.userNama) return a.userNama.localeCompare(b.userNama);
        return (
          new Date(a.tanggalMulai || a.tanggal) -
          new Date(b.tanggalMulai || b.tanggal)
        );
      })
      .map((e) => [
        KATEGORI[e.kategori]?.short || "SPL Staff",
        e.tipe || "Lembur",
        e.userNama,
        e.bagian || "",
        clean(e.pekerjaan || e.deskripsi),
        clean(e.tujuan),
        e.tanggalMulai || e.tanggal || "",
        e.jamMulai || "",
        e.tanggalSelesai || e.tanggal || "",
        e.jamSelesai || "",
        e.durasiJam,
        e.jenis,
        clean(e.hasil),
        clean(e.keterangan),
        e.penanggungJawab || "",
        e.status,
        e.evidence ? e.evidence.name : "",
      ]);
    const escape = (v) => {
      const s = String(v ?? "");
      return /[",\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
    };
    const csv = [headers, ...rows].map((r) => r.map(escape).join(",")).join("\n");
    const blob = new Blob(["﻿" + csv], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    const periodLabel = filter === "semua" ? "semua" : filter;
    a.download = `rekap-lembur-${periodLabel}-${new Date()
      .toISOString()
      .slice(0, 10)}.csv`;
    a.click();
    URL.revokeObjectURL(url);
  };

  const approveAllPending = () => {
    const ids = filtered.filter((e) => e.status === "Terkirim").map((e) => e.id);
    if (ids.length === 0) return;
    onApproveBatch(ids);
  };

  return (
    <div className="space-y-5">
      <div className="pt-2 flex items-start justify-between gap-3">
        <div>
          <h2 className="text-3xl leading-tight" style={serif}>Rekap Tim</h2>
          <p className="text-sm text-zinc-500 mt-1">
            Compile lembur seluruh karyawan.
          </p>
        </div>
        <button
          onClick={exportCSV}
          disabled={filtered.length === 0}
          className="shrink-0 flex items-center gap-1.5 bg-amber-400 text-zinc-950 text-xs font-medium px-3 py-2 rounded-lg hover:bg-amber-300 disabled:opacity-40 transition"
        >
          <Download className="w-3.5 h-3.5" />
          CSV
        </button>
      </div>

      {/* Period filter */}
      <div className="flex gap-2 text-xs" style={mono}>
        {[
          { k: "minggu", l: "MINGGU" },
          { k: "bulan", l: "BULAN" },
          { k: "semua", l: "SEMUA" },
        ].map((f) => (
          <button
            key={f.k}
            onClick={() => setFilter(f.k)}
            className={`px-3 py-1.5 rounded-full uppercase tracking-widest border transition ${
              filter === f.k
                ? "bg-amber-400 text-zinc-950 border-amber-400"
                : "border-zinc-300 dark:border-zinc-600 text-zinc-500"
            }`}
          >
            {f.l}
          </button>
        ))}
      </div>

      {/* Status filter */}
      <div className="flex gap-2 text-xs flex-wrap" style={mono}>
        {["Semua", "Draft", "Terkirim", "Disetujui"].map((s) => (
          <button
            key={s}
            onClick={() => setStatusFilter(s)}
            className={`px-3 py-1.5 rounded-full uppercase tracking-widest border transition ${
              statusFilter === s
                ? "bg-zinc-900 text-amber-400 border-zinc-900 dark:bg-zinc-100 dark:text-zinc-900 dark:border-zinc-100"
                : "border-zinc-300 dark:border-zinc-600 text-zinc-500"
            }`}
          >
            {s}
          </button>
        ))}
      </div>

      {/* Totals */}
      <div className="grid grid-cols-3 gap-2">
        <Stat label="Orang" value={totals.orang} mono={mono} serif={serif} />
        <Stat label="Total Jam" value={totals.jam.toFixed(1)} mono={mono} serif={serif} />
        <Stat
          label="Pending"
          value={totals.pending}
          mono={mono}
          serif={serif}
          highlight
        />
      </div>

      {totals.pending > 0 && (
        <button
          onClick={approveAllPending}
          className="w-full bg-emerald-500 text-white font-medium py-3.5 rounded-lg hover:bg-emerald-400 transition flex items-center justify-center gap-2"
        >
          <CheckCheck className="w-4 h-4" />
          Setujui {totals.pending} Pending
        </button>
      )}

      {/* Per-employee groups */}
      <div className="space-y-2">
        {grouped.length === 0 && (
          <div className="text-center py-12 text-sm text-zinc-500 border border-dashed border-zinc-300 dark:border-zinc-600 rounded-xl">
            Tidak ada data pada filter ini.
          </div>
        )}
        {grouped.map((g) => {
          const open = expanded === g.userId;
          return (
            <div
              key={g.userId}
              className="bg-white dark:bg-zinc-700/50 border border-zinc-200 dark:border-zinc-600 rounded-xl overflow-hidden"
            >
              <button
                onClick={() => setExpanded(open ? null : g.userId)}
                className="w-full p-4 flex items-center gap-3 hover:bg-stone-50 dark:hover:bg-zinc-800/70 transition"
              >
                <div
                  className="w-10 h-10 rounded-full bg-amber-400/20 text-amber-600 dark:text-amber-400 flex items-center justify-center text-sm shrink-0"
                  style={serif}
                >
                  {g.userNama.charAt(0).toUpperCase()}
                </div>
                <div className="flex-1 min-w-0 text-left">
                  <div className="text-sm font-medium truncate">{g.userNama}</div>
                  <div className="text-[11px] text-zinc-500 mt-0.5" style={mono}>
                    {g.entri.length} ENTRI · {g.jam.toFixed(1)} JAM
                    {g.terkirim > 0 && (
                      <span className="text-amber-600 dark:text-amber-400">
                        {" "}
                        · {g.terkirim} PENDING
                      </span>
                    )}
                  </div>
                </div>
                <div className="text-right shrink-0">
                  <div
                    className="text-base text-amber-600 dark:text-amber-400"
                    style={serif}
                  >
                    {g.jam.toFixed(1)}
                    <span className="text-[10px] text-zinc-500 ml-0.5" style={mono}>jam</span>
                  </div>
                </div>
                <ChevronRight
                  className={`w-4 h-4 text-zinc-400 transition-transform shrink-0 ${
                    open ? "rotate-90" : ""
                  }`}
                />
              </button>
              {open && (
                <div className="border-t border-zinc-200 dark:border-zinc-600 p-3 space-y-2 bg-stone-50/50 dark:bg-zinc-900/40">
                  {g.entri.map((e) => (
                    <EntryRow
                      key={e.id}
                      entry={e}
                      onViewEvidence={onViewEvidence}
                      onViewSurat={onViewSurat}
                      onApprove={onApprove}
                      mono={mono}
                      serif={serif}
                    />
                  ))}
                </div>
              )}
            </div>
          );
        })}
      </div>
    </div>
  );
}

// ---------- Payroll: Dashboard ----------
function PayrollDashboard({ user, entries, setView, serif, mono }) {
  const stats = useMemo(() => {
    const monthStart = startOfMonth();
    const inMonth = entries.filter(
      (e) => new Date(e.tanggalMulai || e.tanggal) >= monthStart
    );
    const lemburEntries = inMonth.filter((e) => (e.tipe || "Lembur") === "Lembur");
    const piketEntries = inMonth.filter((e) => e.tipe === "Piket");
    const totalJamLembur = lemburEntries.reduce(
      (s, e) => s + Number(e.durasiJam || 0),
      0
    );
    const totalUpahLembur = lemburEntries.reduce(
      (s, e) => s + Number(e.estimasiGaji || 0),
      0
    );
    const totalJamPiket = piketEntries.reduce(
      (s, e) => s + Number(e.durasiJam || 0),
      0
    );
    const totalUpahPiket = piketEntries.reduce(
      (s, e) => s + Number(e.estimasiGaji || 0),
      0
    );
    const orang = new Set(inMonth.map((e) => e.userId)).size;
    return {
      lemburEntries,
      piketEntries,
      totalJamLembur,
      totalUpahLembur,
      totalJamPiket,
      totalUpahPiket,
      orang,
      total: totalUpahLembur + totalUpahPiket,
    };
  }, [entries]);

  return (
    <div className="space-y-5 stagger">
      <div className="pt-2">
        <div
          className="text-xs uppercase tracking-[0.2em] text-amber-500 dark:text-amber-400/80"
          style={mono}
        >
          Admin Payroll · {user.nama.split(" ")[0]}
        </div>
        <h2 className="text-4xl mt-1 leading-[0.95] tracking-tight" style={serif}>
          Rekap <span className="italic font-light">payroll</span>
        </h2>
      </div>

      <div className="relative rounded-2xl overflow-hidden bg-gradient-to-br from-amber-300 via-amber-400 to-amber-500 text-zinc-950 p-6 shadow-xl shadow-amber-500/20">
        <div
          className="absolute inset-0 opacity-10"
          style={{
            backgroundImage:
              "linear-gradient(rgba(0,0,0,.4) 1px, transparent 1px), linear-gradient(90deg, rgba(0,0,0,.4) 1px, transparent 1px)",
            backgroundSize: "32px 32px",
          }}
        />
        <div className="absolute inset-0 animate-shimmer pointer-events-none" />
        <div className="absolute -top-12 -right-12 w-40 h-40 rounded-full bg-amber-200 opacity-40 blur-2xl" />
        <div className="relative">
          <div className="flex items-center gap-1.5 text-[10px] uppercase tracking-widest mb-3" style={mono}>
            <Clock className="w-3 h-3" />
            Total Jam Bulan Ini
          </div>
          <div className="text-5xl leading-none tracking-tight" style={serif}>
            {(stats.totalJamLembur + stats.totalJamPiket).toFixed(1)}{" "}
            <span className="text-2xl">jam</span>
          </div>
          <div className="grid grid-cols-3 gap-2 mt-6 pt-4 border-t border-zinc-900/15 text-sm">
            <div>
              <div className="opacity-60 text-[9px] uppercase tracking-widest" style={mono}>ENTRI</div>
              <div className="text-base" style={serif}>
                {stats.lemburEntries.length + stats.piketEntries.length}
              </div>
            </div>
            <div>
              <div className="opacity-60 text-[9px] uppercase tracking-widest" style={mono}>ORANG</div>
              <div className="text-base" style={serif}>{stats.orang}</div>
            </div>
            <div>
              <div className="opacity-60 text-[9px] uppercase tracking-widest" style={mono}>PIKET</div>
              <div className="text-base" style={serif}>{stats.totalJamPiket.toFixed(1)}</div>
            </div>
          </div>
        </div>
      </div>

      {/* Breakdown Lembur vs Piket */}
      <div className="grid grid-cols-2 gap-3">
        <div className="bg-white/80 dark:bg-zinc-700/50 backdrop-blur border border-zinc-200 dark:border-zinc-600 rounded-2xl p-4">
          <div className="flex items-center gap-1.5 text-[10px] uppercase tracking-widest text-zinc-500 mb-2" style={mono}>
            <Clock className="w-3 h-3" />
            Lembur
          </div>
          <div className="text-2xl text-amber-600 dark:text-amber-400" style={serif}>
            {stats.totalJamLembur.toFixed(1)} <span className="text-sm">jam</span>
          </div>
          <div className="text-[11px] text-zinc-500 mt-0.5" style={mono}>
            {stats.lemburEntries.length} ENTRI
          </div>
        </div>
        <div className="bg-white/80 dark:bg-zinc-700/50 backdrop-blur border border-zinc-200 dark:border-zinc-600 rounded-2xl p-4">
          <div className="flex items-center gap-1.5 text-[10px] uppercase tracking-widest text-rose-500 dark:text-rose-300 mb-2" style={mono}>
            <CalendarClock className="w-3 h-3" />
            Piket
          </div>
          <div className="text-2xl text-rose-600 dark:text-rose-400" style={serif}>
            {stats.totalJamPiket.toFixed(1)} <span className="text-sm">jam</span>
          </div>
          <div className="text-[11px] text-zinc-500 mt-0.5" style={mono}>
            {stats.piketEntries.length} ENTRI
          </div>
        </div>
      </div>

      <button
        onClick={() => setView("rekap")}
        className="w-full bg-white/80 dark:bg-zinc-700/50 backdrop-blur border border-zinc-200 dark:border-zinc-600 rounded-2xl p-5 flex items-center justify-between hover:border-amber-400 hover:shadow-lg hover:shadow-amber-500/10 lift group"
      >
        <div className="flex items-center gap-3">
          <div className="w-11 h-11 rounded-xl bg-gradient-to-br from-amber-400/20 to-amber-500/10 text-amber-500 dark:text-amber-400 flex items-center justify-center group-hover:scale-105 transition-transform">
            <FileSpreadsheet className="w-5 h-5" strokeWidth={2.5} />
          </div>
          <div className="text-left">
            <div className="text-sm font-medium">Rekap & Download Excel</div>
            <div className="text-xs text-zinc-500">
              Per karyawan · total jam · filter periode
            </div>
          </div>
        </div>
        <ArrowUpRight className="w-4 h-4 text-amber-500 dark:text-amber-400 group-hover:translate-x-0.5 group-hover:-translate-y-0.5 transition-transform" />
      </button>
    </div>
  );
}

// ---------- Payroll: Rekap ----------
function PayrollRekap({ entries, gajiPokok, onViewSurat, onViewEvidence, serif, mono }) {
  const [period, setPeriod] = useState("bulan"); // minggu | bulan | semua
  const [tipeFilter, setTipeFilter] = useState("Semua"); // Semua | Lembur | Piket
  const [kategoriFilter, setKategoriFilter] = useState("Semua"); // Semua | SPL | WaterIntake | Pengemudi
  const [expanded, setExpanded] = useState(null);

  const filtered = useMemo(() => {
    let list = entries;
    if (period !== "semua") {
      const cutoff = period === "minggu" ? startOfWeek() : startOfMonth();
      list = list.filter((e) => new Date(e.tanggalMulai || e.tanggal) >= cutoff);
    }
    if (tipeFilter !== "Semua") {
      list = list.filter((e) => (e.tipe || "Lembur") === tipeFilter);
    }
    if (kategoriFilter !== "Semua") {
      list = list.filter((e) => e.kategori === kategoriFilter);
    }
    return list;
  }, [entries, period, tipeFilter, kategoriFilter]);

  const totals = useMemo(() => {
    const lembur = filtered.filter((e) => (e.tipe || "Lembur") === "Lembur");
    const piket = filtered.filter((e) => e.tipe === "Piket");
    return {
      lemburJam: lembur.reduce((s, e) => s + Number(e.durasiJam || 0), 0),
      lemburUpah: lembur.reduce((s, e) => s + Number(e.estimasiGaji || 0), 0),
      piketJam: piket.reduce((s, e) => s + Number(e.durasiJam || 0), 0),
      piketUpah: piket.reduce((s, e) => s + Number(e.estimasiGaji || 0), 0),
      totalJam: filtered.reduce((s, e) => s + Number(e.durasiJam || 0), 0),
      totalUpah: filtered.reduce((s, e) => s + Number(e.estimasiGaji || 0), 0),
      orang: new Set(filtered.map((e) => e.userId)).size,
    };
  }, [filtered]);

  // Group per employee with lembur/piket breakdown
  const perEmployee = useMemo(() => {
    const map = new Map();
    for (const e of filtered) {
      if (!map.has(e.userId)) {
        map.set(e.userId, {
          userId: e.userId,
          userNama: e.userNama,
          kategori: e.kategori,
          lemburJam: 0,
          lemburUpah: 0,
          piketJam: 0,
          piketUpah: 0,
          entri: [],
        });
      }
      const g = map.get(e.userId);
      g.entri.push(e);
      const isPiket = e.tipe === "Piket";
      if (isPiket) {
        g.piketJam += Number(e.durasiJam || 0);
        g.piketUpah += Number(e.estimasiGaji || 0);
      } else {
        g.lemburJam += Number(e.durasiJam || 0);
        g.lemburUpah += Number(e.estimasiGaji || 0);
      }
    }
    for (const g of map.values()) {
      g.totalJam = g.lemburJam + g.piketJam;
      g.totalUpah = g.lemburUpah + g.piketUpah;
      g.entri.sort(
        (a, b) =>
          new Date(b.tanggalMulai || b.tanggal) -
          new Date(a.tanggalMulai || a.tanggal)
      );
    }
    return [...map.values()].sort((a, b) => b.totalUpah - a.totalUpah);
  }, [filtered]);

  const downloadExcel = () => {
    if (filtered.length === 0) return;
    const wb = XLSX.utils.book_new();

    // Sheet 1: Ringkasan per karyawan (jam saja)
    const summaryRows = [
      [
        "REKAP JAM LEMBUR & PIKET",
      ],
      [
        `Periode: ${
          period === "minggu" ? "Minggu Ini" : period === "bulan" ? "Bulan Ini" : "Semua"
        }`,
        `Tipe: ${tipeFilter}`,
        `Kategori: ${kategoriFilter}`,
        `Diunduh: ${new Date().toLocaleString("id-ID")}`,
      ],
      [],
      [
        "No",
        "Nama",
        "Kategori",
        "Jam Lembur",
        "Jam Piket",
        "Total Jam",
        "Jumlah Entri",
      ],
      ...perEmployee.map((g, i) => [
        i + 1,
        g.userNama,
        KATEGORI[g.kategori]?.short || "SPL",
        Number(g.lemburJam.toFixed(2)),
        Number(g.piketJam.toFixed(2)),
        Number(g.totalJam.toFixed(2)),
        g.entri.length,
      ]),
      [],
      [
        "TOTAL",
        "",
        "",
        Number(totals.lemburJam.toFixed(2)),
        Number(totals.piketJam.toFixed(2)),
        Number(totals.totalJam.toFixed(2)),
        filtered.length,
      ],
    ];
    const wsSummary = XLSX.utils.aoa_to_sheet(summaryRows);
    wsSummary["!cols"] = [
      { wch: 5 },
      { wch: 22 },
      { wch: 14 },
      { wch: 12 },
      { wch: 12 },
      { wch: 12 },
      { wch: 12 },
    ];
    // Merge title row
    wsSummary["!merges"] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 6 } }];
    XLSX.utils.book_append_sheet(wb, wsSummary, "Ringkasan");

    // Sheet 2: Detail entri (jam saja)
    const detailRows = [
      [
        "No",
        "Nama",
        "Kategori",
        "Tipe",
        "Tanggal",
        "Jam Mulai",
        "Jam Selesai",
        "Durasi (jam)",
        "Jenis Hari",
        "Pekerjaan",
        "Keterangan",
        "Status",
      ],
      ...filtered
        .slice()
        .sort((a, b) => {
          if (a.userNama !== b.userNama)
            return a.userNama.localeCompare(b.userNama);
          return (
            new Date(a.tanggalMulai || a.tanggal) -
            new Date(b.tanggalMulai || b.tanggal)
          );
        })
        .map((e, i) => [
          i + 1,
          e.userNama,
          KATEGORI[e.kategori]?.short || "SPL",
          e.tipe || "Lembur",
          e.tanggalMulai || e.tanggal || "",
          e.jamMulai || "",
          e.jamSelesai || "",
          Number(Number(e.durasiJam || 0).toFixed(2)),
          e.jenis || "",
          (e.pekerjaan || e.deskripsi || "").toString(),
          (e.keterangan || "").toString(),
          e.status,
        ]),
    ];
    const wsDetail = XLSX.utils.aoa_to_sheet(detailRows);
    wsDetail["!cols"] = [
      { wch: 5 },
      { wch: 22 },
      { wch: 14 },
      { wch: 8 },
      { wch: 12 },
      { wch: 10 },
      { wch: 10 },
      { wch: 11 },
      { wch: 12 },
      { wch: 32 },
      { wch: 28 },
      { wch: 12 },
    ];
    XLSX.utils.book_append_sheet(wb, wsDetail, "Detail");

    const periodLabel = period === "semua" ? "semua" : period;
    const datestamp = new Date().toISOString().slice(0, 10);
    XLSX.writeFile(wb, `rekap-jam-${periodLabel}-${datestamp}.xlsx`);
  };

  return (
    <div className="space-y-5">
      <div className="pt-2 flex items-start justify-between gap-3">
        <div>
          <h2 className="text-3xl leading-tight" style={serif}>
            Rekap Payroll
          </h2>
          <p className="text-sm text-zinc-500 mt-1">
            Total jam · per karyawan · download Excel.
          </p>
        </div>
        <button
          onClick={downloadExcel}
          disabled={filtered.length === 0}
          className="shrink-0 flex items-center gap-1.5 bg-gradient-to-r from-emerald-500 to-emerald-600 text-white text-xs font-medium px-3 py-2 rounded-lg hover:from-emerald-400 hover:to-emerald-500 disabled:opacity-40 disabled:cursor-not-allowed transition shadow-md shadow-emerald-500/20"
        >
          <FileSpreadsheet className="w-3.5 h-3.5" />
          Excel
        </button>
      </div>

      {/* Period filter */}
      <div className="flex gap-2 text-xs" style={mono}>
        {[
          { k: "minggu", l: "MINGGU" },
          { k: "bulan", l: "BULAN" },
          { k: "semua", l: "SEMUA" },
        ].map((f) => (
          <button
            key={f.k}
            onClick={() => setPeriod(f.k)}
            className={`px-3 py-1.5 rounded-full uppercase tracking-widest border transition ${
              period === f.k
                ? "bg-amber-400 text-zinc-950 border-amber-400"
                : "border-zinc-300 dark:border-zinc-600 text-zinc-500"
            }`}
          >
            {f.l}
          </button>
        ))}
      </div>

      {/* Tipe filter */}
      <div className="flex gap-2 text-xs flex-wrap" style={mono}>
        {["Semua", "Lembur", "Piket"].map((t) => {
          const Icon =
            t === "Lembur" ? Clock : t === "Piket" ? CalendarClock : Filter;
          return (
            <button
              key={t}
              onClick={() => setTipeFilter(t)}
              className={`px-3 py-1.5 rounded-full uppercase tracking-widest border transition flex items-center gap-1.5 ${
                tipeFilter === t
                  ? t === "Piket"
                    ? "bg-rose-500 text-white border-rose-500"
                    : "bg-zinc-900 dark:bg-zinc-100 text-amber-400 dark:text-zinc-900 border-zinc-900 dark:border-zinc-100"
                  : "border-zinc-300 dark:border-zinc-600 text-zinc-500"
              }`}
            >
              <Icon className="w-3 h-3" />
              {t}
            </button>
          );
        })}
      </div>

      {/* Kategori filter */}
      <div className="flex gap-2 text-xs flex-wrap" style={mono}>
        {["Semua", "SPL", "WaterIntake", "Pengemudi"].map((k) => (
          <button
            key={k}
            onClick={() => setKategoriFilter(k)}
            className={`px-3 py-1.5 rounded-full uppercase tracking-widest border transition ${
              kategoriFilter === k
                ? "bg-amber-400 text-zinc-950 border-amber-400"
                : "border-zinc-300 dark:border-zinc-600 text-zinc-500"
            }`}
          >
            {k === "Semua" ? "SEMUA" : KATEGORI[k]?.short || k}
          </button>
        ))}
      </div>

      {/* Stats */}
      <div className="grid grid-cols-2 gap-3">
        <div className="bg-white/80 dark:bg-zinc-700/50 border border-zinc-200 dark:border-zinc-600 rounded-2xl p-4">
          <div
            className="flex items-center gap-1.5 text-[10px] uppercase tracking-widest text-amber-600 dark:text-amber-400 mb-2"
            style={mono}
          >
            <Clock className="w-3 h-3" />
            Lembur
          </div>
          <div className="text-xl text-amber-600 dark:text-amber-400" style={serif}>
            {totals.lemburJam.toFixed(1)} <span className="text-sm">jam</span>
          </div>
        </div>
        <div className="bg-white/80 dark:bg-zinc-700/50 border border-zinc-200 dark:border-zinc-600 rounded-2xl p-4">
          <div
            className="flex items-center gap-1.5 text-[10px] uppercase tracking-widest text-rose-500 dark:text-rose-300 mb-2"
            style={mono}
          >
            <CalendarClock className="w-3 h-3" />
            Piket
          </div>
          <div className="text-xl text-rose-600 dark:text-rose-400" style={serif}>
            {totals.piketJam.toFixed(1)} <span className="text-sm">jam</span>
          </div>
        </div>
      </div>

      <div className="bg-gradient-to-br from-amber-400/10 to-amber-500/5 border border-amber-400/30 rounded-2xl p-4 flex items-center justify-between">
        <div>
          <div className="text-[10px] uppercase tracking-widest text-amber-600 dark:text-amber-400" style={mono}>
            Total Jam Keseluruhan
          </div>
          <div className="text-2xl text-amber-600 dark:text-amber-400" style={serif}>
            {totals.totalJam.toFixed(1)} <span className="text-base">jam</span>
          </div>
        </div>
        <div className="text-right">
          <div className="text-[10px] uppercase tracking-widest text-zinc-500" style={mono}>
            {totals.orang} ORANG
          </div>
          <div className="text-[10px] uppercase tracking-widest text-zinc-500 mt-1" style={mono}>
            {filtered.length} ENTRI
          </div>
        </div>
      </div>

      {/* Per-employee list */}
      <div className="space-y-2">
        {perEmployee.length === 0 && (
          <div className="text-center py-12 text-sm text-zinc-500 border border-dashed border-zinc-300 dark:border-zinc-600 rounded-xl">
            Tidak ada data pada filter ini.
          </div>
        )}
        {perEmployee.map((g) => {
          const open = expanded === g.userId;
          return (
            <div
              key={g.userId}
              className="bg-white/80 dark:bg-zinc-700/50 border border-zinc-200 dark:border-zinc-600 rounded-xl overflow-hidden"
            >
              <button
                onClick={() => setExpanded(open ? null : g.userId)}
                className="w-full p-4 flex items-start gap-3 hover:bg-amber-400/5 transition text-left"
              >
                <div
                  className="w-10 h-10 rounded-full bg-amber-400/20 text-amber-600 dark:text-amber-400 flex items-center justify-center text-sm shrink-0"
                  style={serif}
                >
                  {g.userNama.charAt(0).toUpperCase()}
                </div>
                <div className="flex-1 min-w-0">
                  <div className="flex items-center gap-2 flex-wrap">
                    <span className="text-sm font-medium truncate">{g.userNama}</span>
                    <span
                      className="text-[9px] uppercase tracking-widest text-zinc-500"
                      style={mono}
                    >
                      {KATEGORI[g.kategori]?.short || "SPL"}
                    </span>
                  </div>
                  <div className="grid grid-cols-2 gap-x-3 mt-2 text-[11px]" style={mono}>
                    <div>
                      <span className="text-amber-600 dark:text-amber-400">LEMBUR</span>{" "}
                      <span style={serif} className="text-zinc-700 dark:text-zinc-200">
                        {g.lemburJam.toFixed(1)} jam
                      </span>
                    </div>
                    <div>
                      <span className="text-rose-500 dark:text-rose-300">PIKET</span>{" "}
                      <span style={serif} className="text-zinc-700 dark:text-zinc-200">
                        {g.piketJam.toFixed(1)} jam
                      </span>
                    </div>
                  </div>
                </div>
                <div className="text-right shrink-0">
                  <div
                    className="text-base text-amber-600 dark:text-amber-400"
                    style={serif}
                  >
                    {g.totalJam.toFixed(1)}
                    <span className="text-[10px] text-zinc-500 ml-0.5" style={mono}>jam</span>
                  </div>
                  <div
                    className="text-[10px] text-zinc-500 mt-0.5"
                    style={mono}
                  >
                    {g.entri.length} ENTRI
                  </div>
                </div>
                <ChevronRight
                  className={`w-4 h-4 text-zinc-400 transition-transform shrink-0 mt-2 ${
                    open ? "rotate-90" : ""
                  }`}
                />
              </button>
              {open && (
                <div className="border-t border-zinc-200 dark:border-zinc-600 p-3 space-y-2 bg-stone-50/50 dark:bg-zinc-800/40">
                  {g.entri.map((e) => (
                    <EntryRow
                      key={e.id}
                      entry={e}
                      onViewSurat={onViewSurat}
                      onViewEvidence={onViewEvidence}
                      mono={mono}
                      serif={serif}
                    />
                  ))}
                </div>
              )}
            </div>
          );
        })}
      </div>

      <div className="text-[10px] text-zinc-500 leading-relaxed text-center pt-2" style={mono}>
        FILE EXCEL 2 SHEET: RINGKASAN (PER ORANG) + DETAIL (PER ENTRI)
      </div>
    </div>
  );
}

// ---------- Profil ----------
function Profil({
  user,
  setUser,
  tarif,
  setTarif,
  gajiPokok,
  setGajiPokok,
  company,
  setCompany,
  onLogout,
  serif,
  mono,
}) {
  const [gpGlobalDraft, setGpGlobalDraft] = useState(String(gajiPokok));
  const [gpUserDraft, setGpUserDraft] = useState(
    String(user.gajiPokok || gajiPokok)
  );
  const [bagian, setBagian] = useState(user.bagian || "");
  const [pekerjaan, setPekerjaan] = useState(user.pekerjaan || "");
  const [pj, setPj] = useState(user.penanggungJawab || "");

  const saveProfile = () =>
    setUser({ ...user, bagian, pekerjaan, penanggungJawab: pj });
  const saveUserGp = () => {
    const v = Number(gpUserDraft) || GAJI_POKOK_DEFAULT;
    setUser({ ...user, gajiPokok: v });
  };
  const effectiveGp = user.gajiPokok || gajiPokok;

  return (
    <div className="space-y-5">
      <div className="pt-2">
        <h2 className="text-3xl leading-tight" style={serif}>Profil</h2>
      </div>

      <div className="bg-white dark:bg-zinc-700/50 border border-zinc-200 dark:border-zinc-600 rounded-2xl p-5">
        <div className="flex items-center gap-3">
          <div
            className="w-12 h-12 rounded-full bg-amber-400/20 text-amber-600 dark:text-amber-400 flex items-center justify-center text-lg"
            style={serif}
          >
            {user.nama.charAt(0).toUpperCase()}
          </div>
          <div className="min-w-0">
            <div className="font-medium truncate">{user.nama}</div>
            <div className="text-xs text-zinc-500 truncate" style={mono}>
              {user.email}
            </div>
            <div
              className="text-[10px] uppercase tracking-widest text-amber-600 dark:text-amber-400 mt-0.5"
              style={mono}
            >
              {user.peran}
            </div>
          </div>
        </div>
      </div>

      <FormSection title="Data Pekerja" mono={mono} serif={serif}>
        <Field label="Bagian" mono={mono}>
          <input
            value={bagian}
            onChange={(e) => setBagian(e.target.value)}
            onBlur={saveProfile}
            placeholder="cth. IT Support"
            className={inputCls}
          />
        </Field>
        <Field label="Pekerjaan" mono={mono}>
          <input
            value={pekerjaan}
            onChange={(e) => setPekerjaan(e.target.value)}
            onBlur={saveProfile}
            placeholder="cth. Network Engineer"
            className={inputCls}
          />
        </Field>
        <Field label="Penanggung Jawab Default" mono={mono}>
          <input
            value={pj}
            onChange={(e) => setPj(e.target.value)}
            onBlur={saveProfile}
            placeholder="cth. Rudi (Supervisor IT)"
            className={inputCls}
          />
        </Field>
        <div className="text-[11px] text-zinc-500 leading-relaxed">
          Data ini otomatis terisi setiap kali Anda membuat Surat Perintah Lembur.
        </div>
      </FormSection>

      <FormSection title="Perusahaan" mono={mono} serif={serif}>
        <Field label="Nama Perusahaan" mono={mono}>
          <input
            value={company.nama}
            onChange={(e) => setCompany({ ...company, nama: e.target.value })}
            placeholder="PT. WIRA KARYA INDONESIA"
            className={inputCls}
          />
        </Field>
        <div className="grid grid-cols-2 gap-3">
          <Field label="Inisial / Logo" mono={mono}>
            <input
              value={company.inisial}
              onChange={(e) => setCompany({ ...company, inisial: e.target.value })}
              placeholder="WKI"
              maxLength={5}
              className={inputCls}
              style={mono}
            />
          </Field>
          <Field label="Kota" mono={mono}>
            <input
              value={company.kota}
              onChange={(e) => setCompany({ ...company, kota: e.target.value })}
              placeholder="Jakarta"
              className={inputCls}
            />
          </Field>
        </div>
      </FormSection>

      {user.peran === "Atasan" && (
        <FormSection
          title="Template Water Intake"
          mono={mono}
          serif={serif}
        >
          <Field label="Judul" mono={mono}>
            <input
              value={company.waterIntake?.judul || ""}
              onChange={(e) =>
                setCompany({
                  ...company,
                  waterIntake: { ...company.waterIntake, judul: e.target.value },
                })
              }
              placeholder="DAFTAR LEMBUR PETUGAS KEBERSIHAN"
              className={inputCls}
            />
          </Field>
          <Field label="Sub-judul" mono={mono}>
            <input
              value={company.waterIntake?.subjudul || ""}
              onChange={(e) =>
                setCompany({
                  ...company,
                  waterIntake: { ...company.waterIntake, subjudul: e.target.value },
                })
              }
              placeholder="WATER INTAKE PLTGU BLOK 1 MUARA KARANG"
              className={inputCls}
            />
          </Field>
          <Field label="Instansi" mono={mono}>
            <input
              value={company.waterIntake?.instansi || ""}
              onChange={(e) =>
                setCompany({
                  ...company,
                  waterIntake: { ...company.waterIntake, instansi: e.target.value },
                })
              }
              placeholder="PT. PLN NP UP Muara Karang"
              className={inputCls}
            />
          </Field>
          <div className="grid grid-cols-1 gap-3">
            <Field label="Koordinator Regu" mono={mono}>
              <input
                value={company.waterIntake?.koordinatorRegu || ""}
                onChange={(e) =>
                  setCompany({
                    ...company,
                    waterIntake: {
                      ...company.waterIntake,
                      koordinatorRegu: e.target.value,
                    },
                  })
                }
                placeholder="Earlan S."
                className={inputCls}
              />
            </Field>
            <Field label="Adm. SDM" mono={mono}>
              <input
                value={company.waterIntake?.admSdm || ""}
                onChange={(e) =>
                  setCompany({
                    ...company,
                    waterIntake: { ...company.waterIntake, admSdm: e.target.value },
                  })
                }
                placeholder="(nama Adm. SDM)"
                className={inputCls}
              />
            </Field>
            <Field label="Mengetahui (TTD kanan)" mono={mono}>
              <input
                value={company.waterIntake?.mengetahui || ""}
                onChange={(e) =>
                  setCompany({
                    ...company,
                    waterIntake: {
                      ...company.waterIntake,
                      mengetahui: e.target.value,
                    },
                  })
                }
                placeholder="Roshie"
                className={inputCls}
              />
            </Field>
          </div>
          <div className="text-[11px] text-zinc-500 leading-relaxed">
            Data ini muncul di kop & blok tanda tangan Daftar Lembur Water Intake.
          </div>
        </FormSection>
      )}

      <button
        onClick={onLogout}
        className="w-full border border-zinc-300 dark:border-zinc-600 text-zinc-600 dark:text-zinc-400 py-3 rounded-lg hover:border-red-400 hover:text-red-500 transition flex items-center justify-center gap-2"
      >
        <LogOut className="w-4 h-4" />
        Keluar
      </button>

      <div className="text-center text-[10px] text-zinc-500 pt-4" style={mono}>
        LEMBUR.ID — DEMO LOKAL
      </div>
    </div>
  );
}

// ---------- Surat Perintah Lembur (printable) ----------
function SuratView({ entry, company, currentUser, onApprove, onClose, mono, serif }) {
  const s = STATUS[entry.status] || STATUS.Draft;

  const handlePrint = () => {
    window.print();
  };

  return (
    <div
      className="fixed inset-0 z-50 bg-black/80 backdrop-blur-sm flex items-center justify-center p-2 md:p-6 surat-modal"
      onClick={onClose}
    >
      <div
        className="bg-white text-zinc-900 w-full max-w-2xl max-h-[95vh] overflow-y-auto rounded-xl surat-sheet"
        onClick={(e) => e.stopPropagation()}
      >
        {/* Header bar (hidden in print) */}
        <div className="surat-toolbar flex items-center justify-between px-4 py-3 border-b border-zinc-200 sticky top-0 bg-white z-10">
          <div className="flex items-center gap-2">
            <span
              className={`text-[10px] px-2 py-1 rounded uppercase tracking-widest ${s.cls}`}
              style={mono}
            >
              {s.label}
            </span>
            <span className="text-[10px] uppercase tracking-widest text-zinc-500" style={mono}>
              No. {String(entry.id).slice(-6)}/SPL/{new Date(entry.tanggalMulai || entry.tanggal).getFullYear()}
            </span>
          </div>
          <div className="flex items-center gap-2">
            {onApprove && (
              <button
                onClick={onApprove}
                className="flex items-center gap-1.5 bg-emerald-500 text-white text-xs px-3 py-1.5 rounded-lg hover:bg-emerald-400 transition"
              >
                <Check className="w-3.5 h-3.5" />
                Setujui
              </button>
            )}
            <button
              onClick={handlePrint}
              className="flex items-center gap-1.5 bg-amber-400 text-zinc-950 text-xs font-medium px-3 py-1.5 rounded-lg hover:bg-amber-300 transition"
            >
              <Printer className="w-3.5 h-3.5" />
              Cetak
            </button>
            <button
              onClick={onClose}
              className="w-8 h-8 rounded-full border border-zinc-300 flex items-center justify-center text-zinc-500 hover:text-red-500 hover:border-red-400 transition"
            >
              <X className="w-4 h-4" />
            </button>
          </div>
        </div>

        {/* The Surat itself (printable area) */}
        <div className="surat-printable p-6 md:p-10" style={{ fontFamily: "'Geist', system-ui, sans-serif" }}>
          {/* Kop */}
          <div className="flex items-center gap-4 mb-6">
            <div className="w-16 h-16 rounded-full border-2 border-zinc-900 flex items-center justify-center text-xs font-bold tracking-wider shrink-0">
              {company.inisial}
            </div>
            <div className="flex-1 text-center">
              <div className="text-xl md:text-2xl font-bold tracking-wide">
                SURAT PERINTAH LEMBUR
              </div>
              <div className="text-xs italic text-zinc-600 mt-0.5">
                ( Penugasan Kerja Lebih )
              </div>
              <div className="text-xs font-semibold mt-1">
                STAFF {company.nama.toUpperCase()}
              </div>
            </div>
            <div className="w-16 shrink-0" />
          </div>

          {/* Identitas */}
          <SuratTable>
            <SuratRow label="NAMA" value={entry.userNama} />
            <SuratRow label="BAGIAN" value={entry.bagian} />
          </SuratTable>

          <div className="h-3" />

          <SuratTable>
            <SuratRow label="PEKERJAAN" value={entry.pekerjaan || entry.deskripsi} multiline />
          </SuratTable>

          <div className="h-3" />

          <SuratTable>
            <SuratRow label="TUJUAN" value={entry.tujuan} multiline />
          </SuratTable>

          <div className="h-3" />

          {/* Waktu Lembur */}
          <table className="w-full border-collapse text-sm">
            <tbody>
              <tr>
                <td colSpan={4} className="border border-zinc-900 px-3 py-2 text-center font-semibold">
                  WAKTU LEMBUR
                </td>
              </tr>
              <tr>
                <td className="border border-zinc-900 px-3 py-2 w-[28%] font-medium">
                  TANGGAL MULAI <span className="float-right">:</span>
                </td>
                <td className="border border-zinc-900 px-3 py-2 w-[22%]">
                  {formatTanggalPanjang(entry.tanggalMulai || entry.tanggal)}
                </td>
                <td className="border border-zinc-900 px-3 py-2 w-[28%] font-medium">
                  JAM MULAI <span className="float-right">:</span>
                </td>
                <td className="border border-zinc-900 px-3 py-2 w-[22%]">
                  {entry.jamMulai || "-"}
                </td>
              </tr>
              <tr>
                <td className="border border-zinc-900 px-3 py-2 font-medium">
                  TANGGAL SELESAI <span className="float-right">:</span>
                </td>
                <td className="border border-zinc-900 px-3 py-2">
                  {formatTanggalPanjang(entry.tanggalSelesai || entry.tanggal)}
                </td>
                <td className="border border-zinc-900 px-3 py-2 font-medium">
                  JAM SELESAI <span className="float-right">:</span>
                </td>
                <td className="border border-zinc-900 px-3 py-2">
                  {entry.jamSelesai || "-"}
                </td>
              </tr>
            </tbody>
          </table>

          <div className="h-3" />

          <SuratTable>
            <SuratRow label="HASIL" value={entry.hasil} multiline />
          </SuratTable>

          <div className="h-3" />

          <SuratTable>
            <SuratRow label="KETERANGAN" value={entry.keterangan} multiline />
          </SuratTable>

          {/* Ringkasan durasi */}
          <div className="mt-4 grid grid-cols-2 gap-3 text-xs surat-summary">
            <div className="border border-zinc-300 rounded p-2 text-center">
              <div className="text-zinc-500">Total Durasi</div>
              <div className="font-semibold text-base mt-0.5">
                {Number(entry.durasiJam || 0).toFixed(2)} jam
              </div>
            </div>
            <div className="border border-zinc-300 rounded p-2 text-center">
              <div className="text-zinc-500">Jenis Hari</div>
              <div className="font-semibold text-base mt-0.5">{entry.jenis}</div>
            </div>
          </div>

          {/* Tanda tangan */}
          <div className="mt-8 text-right text-sm">
            {company.kota}, {formatTanggalPanjang(new Date().toISOString())}
          </div>

          <div className="mt-6 grid grid-cols-3 gap-4 text-center text-sm">
            <SignatureSlot
              label="Yang Membuat"
              nama={entry.userNama}
              ttd={entry.ttdPekerja}
            />
            <SignatureSlot
              label="Penanggung Jawab"
              nama={entry.penanggungJawab}
              ttd={entry.ttdPenanggungJawab}
            />
            <SignatureSlot
              label="Mengetahui"
              nama={
                entry.atasanNama ||
                (entry.status === "Disetujui" && currentUser?.peran === "Atasan"
                  ? currentUser.nama
                  : null)
              }
              ttd={entry.ttdAtasan}
            />
          </div>

          {entry.catatanAtasan && (
            <div className="mt-4 bg-emerald-50 border border-emerald-200 rounded p-2 text-[11px] text-emerald-900">
              <span className="font-semibold uppercase tracking-widest">
                Catatan Atasan:
              </span>{" "}
              {entry.catatanAtasan}
            </div>
          )}

          {/* QR + Evidence footer */}
          <div className="mt-6 pt-4 border-t border-dashed border-zinc-300 flex items-center gap-4">
            <div className="bg-white border border-zinc-300 p-1.5 rounded shrink-0">
              <QRCodeView data={buildQRPayload(entry, company)} size={84} />
            </div>
            <div className="text-[10px] text-zinc-500 leading-relaxed flex-1">
              <div className="uppercase tracking-widest font-semibold text-zinc-700">
                Kode Verifikasi
              </div>
              <div className="mt-0.5">
                Pindai QR untuk verifikasi keaslian Surat Perintah Lembur ini.
              </div>
              {entry.evidence && (
                <div className="mt-1.5">
                  Lampiran bukti:{" "}
                  <span className="font-mono">{entry.evidence.name}</span> (
                  {(entry.evidence.size / 1024).toFixed(0)} KB)
                </div>
              )}
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

// ---------- Daftar Lembur Water Intake (multi-row) ----------
function DaftarLemburView({
  anchorEntry,
  allEntries,
  company,
  currentUser,
  onApproveBatch,
  onClose,
  mono,
  serif,
}) {
  // Find all entries of SAME kategori (Water Intake / Pengemudi) on the same date as anchor
  const tanggal = anchorEntry.tanggalMulai || anchorEntry.tanggal;
  const kategoriId = anchorEntry.kategori;
  const rows = useMemo(() => {
    return allEntries
      .filter(
        (e) =>
          e.kategori === kategoriId &&
          (e.tanggalMulai || e.tanggal) === tanggal
      )
      .sort((a, b) => (a.jamMulai || "").localeCompare(b.jamMulai || ""));
  }, [allEntries, tanggal, kategoriId]);

  // Template config: per-kategori (waterIntake) or fallback for Pengemudi
  const wi = kategoriId === "Pengemudi"
    ? (company.pengemudi || {
        judul: "DAFTAR LEMBUR PENGEMUDI",
        subjudul: company.waterIntake?.subjudul || "",
        kota: company.kota,
        instansi: company.waterIntake?.instansi || "",
        koordinatorRegu: "",
        admSdm: "",
        mengetahui: "",
      })
    : company.waterIntake || {};
  const pendingIds = rows
    .filter((e) => e.status === "Terkirim")
    .map((e) => e.id);
  const allDisetujui =
    rows.length > 0 && rows.every((e) => e.status === "Disetujui");
  const anySigned = rows.find((e) => e.ttdAtasan);

  const handlePrint = () => window.print();

  return (
    <div
      className="fixed inset-0 z-50 bg-black/80 backdrop-blur-sm flex items-center justify-center p-2 md:p-6 surat-modal"
      onClick={onClose}
    >
      <div
        className="bg-white text-zinc-900 w-full max-w-3xl max-h-[95vh] overflow-y-auto rounded-xl surat-sheet"
        onClick={(e) => e.stopPropagation()}
      >
        <div className="surat-toolbar flex items-center justify-between px-4 py-3 border-b border-zinc-200 sticky top-0 bg-white z-10">
          <div className="flex items-center gap-2 flex-wrap">
            <span
              className="text-[10px] uppercase tracking-widest text-zinc-500 flex items-center gap-1"
              style={mono}
            >
              {kategoriId === "Pengemudi" ? (
                <Wallet className="w-3 h-3" />
              ) : (
                <Droplets className="w-3 h-3" />
              )}
              {KATEGORI[kategoriId]?.short || "Daftar"} ·{" "}
              {formatTanggal(tanggal)}
            </span>
            <span
              className="text-[10px] uppercase tracking-widest text-zinc-500"
              style={mono}
            >
              · {rows.length} pekerja
            </span>
          </div>
          <div className="flex items-center gap-2">
            {onApproveBatch && pendingIds.length > 0 && (
              <button
                onClick={() => onApproveBatch(pendingIds)}
                className="flex items-center gap-1.5 bg-emerald-500 text-white text-xs px-3 py-1.5 rounded-lg hover:bg-emerald-400 transition"
              >
                <Check className="w-3.5 h-3.5" />
                Setujui {pendingIds.length}
              </button>
            )}
            <button
              onClick={handlePrint}
              className="flex items-center gap-1.5 bg-amber-400 text-zinc-950 text-xs font-medium px-3 py-1.5 rounded-lg hover:bg-amber-300 transition"
            >
              <Printer className="w-3.5 h-3.5" />
              Cetak
            </button>
            <button
              onClick={onClose}
              className="w-8 h-8 rounded-full border border-zinc-300 flex items-center justify-center text-zinc-500 hover:text-red-500 hover:border-red-400 transition"
            >
              <X className="w-4 h-4" />
            </button>
          </div>
        </div>

        {/* Printable document */}
        <div
          className="surat-printable p-6 md:p-10"
          style={{ fontFamily: "'Geist', system-ui, sans-serif" }}
        >
          <div className="text-center mb-4">
            <div className="text-lg md:text-xl font-bold tracking-wide">
              {wi.judul ||
                (kategoriId === "Pengemudi"
                  ? "DAFTAR LEMBUR PENGEMUDI"
                  : "DAFTAR LEMBUR PETUGAS KEBERSIHAN")}
            </div>
            <div className="text-sm font-semibold mt-0.5">
              {wi.subjudul || "WATER INTAKE PLTGU BLOK 1 MUARA KARANG"}
            </div>
          </div>

          {/* Hari & Tanggal */}
          <table className="w-full border-collapse text-sm mb-2">
            <tbody>
              <tr>
                <td className="px-2 py-1 w-[60%]">
                  <span className="font-medium">Hari</span> :{" "}
                  <span className="underline">{namaHari(tanggal)}</span>
                </td>
                <td className="px-2 py-1 text-right">
                  <span className="font-medium">Tanggal</span> :{" "}
                  <span className="underline">{formatTanggal(tanggal)}</span>
                </td>
              </tr>
            </tbody>
          </table>

          {/* Main table */}
          <table className="w-full border-collapse text-sm">
            <thead>
              <tr>
                <th className="border border-zinc-900 px-2 py-1 w-[5%]" rowSpan={2}>NO</th>
                <th className="border border-zinc-900 px-2 py-1 w-[20%]" rowSpan={2}>NAMA</th>
                <th className="border border-zinc-900 px-2 py-1 w-[28%]" rowSpan={2}>PEKERJAAN</th>
                <th className="border border-zinc-900 px-2 py-1" colSpan={2}>JAM</th>
                <th className="border border-zinc-900 px-2 py-1 w-[12%]" rowSpan={2}>LAMA LEMBUR</th>
                <th className="border border-zinc-900 px-2 py-1 w-[15%]" rowSpan={2}>KETERANGAN</th>
              </tr>
              <tr>
                <th className="border border-zinc-900 px-2 py-1 w-[10%]">MSK</th>
                <th className="border border-zinc-900 px-2 py-1 w-[10%]">PLG</th>
              </tr>
            </thead>
            <tbody>
              {rows.map((e, i) => (
                <tr key={e.id}>
                  <td className="border border-zinc-900 px-2 py-2 text-center">
                    {i + 1}
                  </td>
                  <td className="border border-zinc-900 px-2 py-2 align-top">
                    {e.userNama}
                    {e.ttdPekerja && (
                      <img
                        src={e.ttdPekerja}
                        alt=""
                        className="max-h-10 mt-1"
                      />
                    )}
                  </td>
                  <td className="border border-zinc-900 px-2 py-2 align-top">
                    {e.pekerjaan || "-"}
                  </td>
                  <td className="border border-zinc-900 px-2 py-2 text-center font-mono">
                    {e.jamMulai || "-"}
                  </td>
                  <td className="border border-zinc-900 px-2 py-2 text-center font-mono">
                    {e.jamSelesai || "-"}
                  </td>
                  <td className="border border-zinc-900 px-2 py-2 text-center">
                    {Number(e.durasiJam || 0).toFixed(0)} jam
                  </td>
                  <td className="border border-zinc-900 px-2 py-2 align-top text-xs">
                    <div>{e.keterangan || ""}</div>
                    <div
                      className={`text-[9px] uppercase tracking-widest mt-0.5 ${
                        e.status === "Disetujui"
                          ? "text-emerald-700"
                          : e.status === "Terkirim"
                          ? "text-amber-700"
                          : "text-zinc-500"
                      }`}
                    >
                      {e.status}
                    </div>
                  </td>
                </tr>
              ))}
              {/* Padding rows to fill the page like the physical form */}
              {Array.from({ length: Math.max(0, 4 - rows.length) }).map(
                (_, i) => (
                  <tr key={`pad-${i}`}>
                    <td className="border border-zinc-900 px-2 py-3">&nbsp;</td>
                    <td className="border border-zinc-900 px-2 py-3" />
                    <td className="border border-zinc-900 px-2 py-3" />
                    <td className="border border-zinc-900 px-2 py-3" />
                    <td className="border border-zinc-900 px-2 py-3" />
                    <td className="border border-zinc-900 px-2 py-3" />
                    <td className="border border-zinc-900 px-2 py-3" />
                  </tr>
                )
              )}
            </tbody>
          </table>

          {/* Footer / signatures */}
          <div className="mt-4 text-right text-sm">
            {wi.kota || company.kota}, {formatTanggal(tanggal)}
            <br />
            {wi.instansi || ""}
          </div>

          <div className="mt-6 grid grid-cols-3 gap-4 text-center text-sm">
            <SignatureSlot
              label=""
              nama={wi.koordinatorRegu}
              ttd={anchorEntry.ttdPekerja}
              underLabel="Koordinator Regu"
            />
            <div>
              <div className="mb-2">Mengetahui,</div>
              <div className="h-20 flex items-center justify-center">
                {anySigned && anchorEntry.ttdAtasan && (
                  <img
                    src={anchorEntry.ttdAtasan}
                    alt=""
                    className="max-h-20 max-w-full object-contain"
                  />
                )}
              </div>
              <div className="border-t border-zinc-900 pt-1">
                ( {wi.admSdm || "................................"} )
              </div>
              <div className="text-xs mt-0.5">Adm. SDM</div>
            </div>
            <SignatureSlot
              label=""
              nama={
                allDisetujui && currentUser?.peran === "Atasan"
                  ? wi.mengetahui || currentUser.nama
                  : wi.mengetahui
              }
              ttd={anySigned?.ttdAtasan}
              underLabel={wi.instansi || "Atasan"}
            />
          </div>
        </div>
      </div>
    </div>
  );
}

function SignatureSlot({ label, nama, ttd, underLabel }) {
  return (
    <div>
      {label ? <div>{label},</div> : <div>&nbsp;</div>}
      <div className="h-20 flex items-center justify-center">
        {ttd ? (
          <img
            src={ttd}
            alt={`Tanda tangan ${label || underLabel || ""}`}
            className="max-h-20 max-w-full object-contain"
          />
        ) : null}
      </div>
      <div className="border-t border-zinc-900 pt-1">
        ( {nama || "................................"} )
      </div>
      {underLabel && <div className="text-xs mt-0.5">{underLabel}</div>}
    </div>
  );
}

function SuratTable({ children }) {
  return (
    <table className="w-full border-collapse text-sm">
      <tbody>{children}</tbody>
    </table>
  );
}

function SuratRow({ label, value, multiline }) {
  return (
    <tr>
      <td className="border border-zinc-900 px-3 py-2 w-[28%] align-top font-medium">
        {label} <span className="float-right">:</span>
      </td>
      <td
        className={`border border-zinc-900 px-3 py-2 ${
          multiline ? "min-h-[3rem] align-top" : ""
        }`}
      >
        {value || <span className="text-zinc-300">—</span>}
      </td>
    </tr>
  );
}

// ---------- Bottom Nav ----------
function BottomNav({ view, setView, isAdmin, isPayroll, mono }) {
  const items = isPayroll
    ? [
        { k: "dashboard", l: "Beranda", icon: LayoutDashboard },
        { k: "rekap", l: "Rekap Payroll", icon: FileSpreadsheet },
        { k: "profil", l: "Profil", icon: User },
      ]
    : isAdmin
    ? [
        { k: "dashboard", l: "Beranda", icon: LayoutDashboard },
        { k: "rekap", l: "Rekap", icon: Users },
        { k: "profil", l: "Profil", icon: User },
      ]
    : [
        { k: "dashboard", l: "Beranda", icon: LayoutDashboard },
        { k: "input", l: "Tambah", icon: Plus },
        { k: "riwayat", l: "Riwayat", icon: History },
        { k: "profil", l: "Profil", icon: User },
      ];

  return (
    <nav className="fixed bottom-0 inset-x-0 z-20 animate-slide-up">
      <div className="mx-auto max-w-md px-3 pb-3">
        <div
          className={`relative bg-white/85 dark:bg-zinc-700/60 backdrop-blur-xl border border-zinc-200/70 dark:border-zinc-600/70 rounded-2xl shadow-2xl shadow-black/10 dark:shadow-black/50 grid ${
            isAdmin || isPayroll ? "grid-cols-3" : "grid-cols-4"
          }`}
        >
          {items.map((it) => {
            const active = view === it.k;
            const Icon = it.icon;
            const isAdd = it.k === "input";
            return (
              <button
                key={it.k}
                onClick={() => setView(it.k)}
                className={`relative flex flex-col items-center justify-center py-2.5 gap-0.5 transition-all duration-200 ${
                  active && !isAdd
                    ? "text-amber-500 dark:text-amber-400"
                    : isAdd
                    ? "text-zinc-950"
                    : "text-zinc-400 dark:text-zinc-500"
                }`}
              >
                {active && !isAdd && (
                  <div className="absolute top-0 left-1/2 -translate-x-1/2 w-8 h-0.5 rounded-full bg-amber-400 animate-slide-down" />
                )}
                <div
                  className={`flex items-center justify-center transition-transform ${
                    isAdd
                      ? "w-10 h-10 rounded-full bg-gradient-to-br from-amber-300 to-amber-500 text-zinc-950 -mt-1.5 shadow-lg shadow-amber-500/40 active:scale-95"
                      : `w-6 h-6 ${active ? "scale-110" : ""}`
                  }`}
                >
                  <Icon className="w-5 h-5" strokeWidth={isAdd ? 2.6 : 2} />
                </div>
                <span
                  className={`text-[9px] uppercase tracking-widest transition-opacity ${
                    isAdd ? "opacity-80" : ""
                  }`}
                  style={mono}
                >
                  {it.l}
                </span>
              </button>
            );
          })}
        </div>
      </div>
    </nav>
  );
}

// ---------- Signature Pad ----------
function SignaturePad({ value, onChange, height = 140, mono }) {
  const containerRef = useRef(null);
  const canvasRef = useRef(null);
  const onChangeRef = useRef(onChange);
  const [hasInk, setHasInk] = useState(!!value);

  useEffect(() => {
    onChangeRef.current = onChange;
  });

  useEffect(() => {
    const canvas = canvasRef.current;
    const container = containerRef.current;
    if (!canvas || !container) return;

    const cssW = container.clientWidth || 320;
    const cssH = height;
    const ratio = window.devicePixelRatio || 1;
    canvas.style.width = cssW + "px";
    canvas.style.height = cssH + "px";
    canvas.width = Math.floor(cssW * ratio);
    canvas.height = Math.floor(cssH * ratio);

    const ctx = canvas.getContext("2d");
    ctx.scale(ratio, ratio);
    ctx.lineCap = "round";
    ctx.lineJoin = "round";
    ctx.lineWidth = 2.2;
    ctx.strokeStyle = "#0a0a0a";
    ctx.fillStyle = "#0a0a0a";

    if (value) {
      const img = new Image();
      img.onload = () => ctx.drawImage(img, 0, 0, cssW, cssH);
      img.src = value;
    }

    let drawing = false;
    let lastX = 0;
    let lastY = 0;

    const xy = (clientX, clientY) => {
      const r = canvas.getBoundingClientRect();
      return [clientX - r.left, clientY - r.top];
    };

    const beginStroke = (cx, cy) => {
      drawing = true;
      [lastX, lastY] = xy(cx, cy);
      ctx.beginPath();
      ctx.arc(lastX, lastY, 1.4, 0, Math.PI * 2);
      ctx.fill();
      setHasInk(true);
    };

    const extendStroke = (cx, cy) => {
      const [x, y] = xy(cx, cy);
      ctx.beginPath();
      ctx.moveTo(lastX, lastY);
      ctx.lineTo(x, y);
      ctx.stroke();
      lastX = x;
      lastY = y;
    };

    const endStroke = () => {
      if (!drawing) return;
      drawing = false;
      const url = canvas.toDataURL("image/png");
      onChangeRef.current?.(url);
    };

    // Mouse
    const onMouseDown = (e) => {
      if (e.button !== 0) return;
      e.preventDefault();
      beginStroke(e.clientX, e.clientY);
    };
    const onMouseMove = (e) => {
      if (!drawing) return;
      extendStroke(e.clientX, e.clientY);
    };
    const onMouseUp = () => endStroke();

    // Touch
    const onTouchStart = (e) => {
      if (e.touches.length === 0) return;
      e.preventDefault();
      const t = e.touches[0];
      beginStroke(t.clientX, t.clientY);
    };
    const onTouchMove = (e) => {
      if (!drawing || e.touches.length === 0) return;
      e.preventDefault();
      const t = e.touches[0];
      extendStroke(t.clientX, t.clientY);
    };
    const onTouchEnd = () => endStroke();

    canvas.addEventListener("mousedown", onMouseDown);
    window.addEventListener("mousemove", onMouseMove);
    window.addEventListener("mouseup", onMouseUp);
    canvas.addEventListener("touchstart", onTouchStart, { passive: false });
    canvas.addEventListener("touchmove", onTouchMove, { passive: false });
    canvas.addEventListener("touchend", onTouchEnd);
    canvas.addEventListener("touchcancel", onTouchEnd);

    return () => {
      canvas.removeEventListener("mousedown", onMouseDown);
      window.removeEventListener("mousemove", onMouseMove);
      window.removeEventListener("mouseup", onMouseUp);
      canvas.removeEventListener("touchstart", onTouchStart);
      canvas.removeEventListener("touchmove", onTouchMove);
      canvas.removeEventListener("touchend", onTouchEnd);
      canvas.removeEventListener("touchcancel", onTouchEnd);
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const clear = () => {
    const canvas = canvasRef.current;
    if (!canvas) return;
    const ctx = canvas.getContext("2d");
    ctx.save();
    ctx.setTransform(1, 0, 0, 1, 0, 0);
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    ctx.restore();
    setHasInk(false);
    onChange(null);
  };

  return (
    <div>
      <div
        ref={containerRef}
        style={{
          position: "relative",
          background: "white",
          border: "2px dashed #d4d4d8",
          borderRadius: 10,
          overflow: "hidden",
          userSelect: "none",
        }}
      >
        <canvas
          ref={canvasRef}
          style={{
            display: "block",
            touchAction: "none",
            cursor: "crosshair",
            background: "white",
          }}
        />
        {!hasInk && (
          <div
            style={{
              position: "absolute",
              inset: 0,
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              pointerEvents: "none",
              color: "#a1a1aa",
              fontSize: 10,
              letterSpacing: "0.15em",
              textTransform: "uppercase",
              ...mono,
            }}
          >
            tanda tangan di sini
          </div>
        )}
      </div>
      <div className="flex items-center justify-between mt-1.5">
        <span
          className="text-[10px] uppercase tracking-widest text-zinc-500"
          style={mono}
        >
          {hasInk ? "✓ tertandatangani" : "kosong"}
        </span>
        <button
          type="button"
          onClick={clear}
          disabled={!hasInk}
          className="text-xs text-zinc-500 hover:text-red-500 disabled:opacity-40 transition flex items-center gap-1"
        >
          <Eraser className="w-3 h-3" />
          Hapus
        </button>
      </div>
    </div>
  );
}

// ---------- QR Code ----------
function QRCodeView({ data, size = 110 }) {
  const [src, setSrc] = useState(null);
  useEffect(() => {
    let alive = true;
    QRCode.toDataURL(data, {
      width: size * 2,
      margin: 1,
      errorCorrectionLevel: "M",
    })
      .then((url) => {
        if (alive) setSrc(url);
      })
      .catch(() => alive && setSrc(null));
    return () => {
      alive = false;
    };
  }, [data, size]);
  if (!src)
    return (
      <div
        style={{ width: size, height: size }}
        className="bg-zinc-100 dark:bg-zinc-800 rounded"
      />
    );
  return <img src={src} width={size} height={size} alt="QR" className="block" />;
}

const buildQRPayload = (entry, company) =>
  JSON.stringify({
    spl: String(entry.id).slice(-6),
    co: company?.inisial || "WKI",
    nm: entry.userNama,
    bg: entry.bagian,
    tm: entry.tanggalMulai,
    jm: entry.jamMulai,
    ts: entry.tanggalSelesai,
    js: entry.jamSelesai,
    dur: entry.durasiJam,
    upah: entry.estimasiGaji,
    st: entry.status,
  });

// ---------- Approval Dialog ----------
function ApprovalDialog({ target, currentUser, onConfirm, onCancel, serif, mono }) {
  const [sig, setSig] = useState(null);
  const [catatan, setCatatan] = useState("");

  if (!target) return null;
  const count = target.kind === "batch" ? target.ids.length : 1;
  const title =
    target.kind === "batch"
      ? `Setujui ${count} entri sekaligus`
      : `Setujui Surat Perintah`;

  return (
    <div
      className="fixed inset-0 z-[60] bg-black/80 backdrop-blur-sm flex items-center justify-center p-4"
      onClick={onCancel}
    >
      <div
        className="bg-stone-50 dark:bg-zinc-800 text-zinc-900 dark:text-zinc-100 border border-zinc-200 dark:border-zinc-600 w-full max-w-md rounded-2xl overflow-hidden"
        onClick={(e) => e.stopPropagation()}
        style={{ fontFamily: "'Geist', system-ui, sans-serif" }}
      >
        <div className="px-5 py-4 border-b border-zinc-200 dark:border-zinc-600 flex items-center justify-between">
          <div>
            <h3 className="text-lg" style={serif}>
              {title}
            </h3>
            <div
              className="text-[10px] uppercase tracking-widest text-zinc-500 mt-0.5"
              style={mono}
            >
              Mengetahui · {currentUser.nama}
            </div>
          </div>
          <button
            onClick={onCancel}
            className="w-8 h-8 rounded-full border border-zinc-300 dark:border-zinc-600 flex items-center justify-center text-zinc-500 hover:text-red-500 hover:border-red-400 transition"
          >
            <X className="w-4 h-4" />
          </button>
        </div>

        <div className="p-5 space-y-4">
          {target.kind === "single" && (
            <div className="bg-white dark:bg-zinc-800 border border-zinc-200 dark:border-zinc-600 rounded-lg p-3 text-sm">
              <div className="font-medium">{target.entry.userNama}</div>
              <div className="text-xs text-zinc-500 mt-0.5">
                {target.entry.pekerjaan} · {Number(target.entry.durasiJam).toFixed(2)} jam ·{" "}
                {target.entry.jenis}
              </div>
            </div>
          )}

          <div>
            <div
              className="block text-[10px] uppercase tracking-widest text-zinc-500 mb-1.5"
              style={mono}
            >
              Tanda Tangan Atasan
            </div>
            <SignaturePad value={sig} onChange={setSig} mono={mono} />
          </div>

          <Field label="Catatan (opsional)" mono={mono}>
            <textarea
              value={catatan}
              onChange={(e) => setCatatan(e.target.value)}
              rows={2}
              placeholder="cth. Disetujui untuk pembayaran payroll Mei"
              className={inputCls + " resize-none"}
            />
          </Field>
        </div>

        <div className="px-5 py-4 border-t border-zinc-200 dark:border-zinc-600 flex gap-2">
          <button
            onClick={onCancel}
            className="flex-1 py-2.5 rounded-lg border border-zinc-300 dark:border-zinc-600 text-sm hover:border-zinc-500 transition"
          >
            Batal
          </button>
          <button
            disabled={!sig}
            onClick={() => onConfirm({ ttdAtasan: sig, catatanAtasan: catatan })}
            className="flex-1 py-2.5 rounded-lg bg-emerald-500 text-white text-sm font-medium hover:bg-emerald-400 disabled:opacity-40 disabled:cursor-not-allowed transition flex items-center justify-center gap-2"
          >
            <Check className="w-4 h-4" />
            Setujui
          </button>
        </div>
      </div>
    </div>
  );
}

// ---------- Small UI ----------
function Field({ label, children, mono, required }) {
  return (
    <label className="block">
      <span className="block text-[12px] font-medium tracking-wide text-zinc-400 dark:text-zinc-500 mb-1.5">
        {label}
        {required && <span className="text-amber-500 ml-0.5">*</span>}
      </span>
      {children}
    </label>
  );
}
