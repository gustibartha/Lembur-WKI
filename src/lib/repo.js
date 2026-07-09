// ============================================================
// Repository: satu API untuk data lembur, otomatis pilih backend.
// - Supabase aktif (env terisi)  → tabel public.lembur
// - Supabase nonaktif            → localStorage (perilaku lama)
// Semua fungsi async supaya pemanggil seragam.
// ============================================================
import { supabase, isSupabaseEnabled } from "./supabase";

const LS_ENTRIES = "lembur:entries:v6";

// ---------- Mapping camelCase (app) <-> snake_case (db) ----------
const toRow = (e) => ({
  // id dibiarkan Supabase generate untuk entri baru; sertakan jika update
  karyawan_id: e.karyawanId ?? null,
  user_nama: e.userNama,
  kategori: e.kategori ?? "SPL",
  tipe: e.tipe ?? "Lembur",
  bagian: e.bagian ?? null,
  pekerjaan: e.pekerjaan ?? null,
  tujuan: e.tujuan ?? null,
  tanggal_mulai: e.tanggalMulai || null,
  jam_mulai: e.jamMulai ?? null,
  tanggal_selesai: e.tanggalSelesai || null,
  jam_selesai: e.jamSelesai ?? null,
  jenis: e.jenis ?? "Hari Kerja",
  durasi_jam: Number(e.durasiJam || 0),
  estimasi_gaji: Number(e.estimasiGaji || 0),
  hasil: e.hasil ?? null,
  keterangan: e.keterangan ?? null,
  penanggung_jawab: e.penanggungJawab ?? null,
  status: e.status ?? "Draft",
  evidence: e.evidence ?? null,
  ttd_pekerja: e.ttdPekerja ?? null,
  ttd_atasan: e.ttdAtasan ?? null,
  atasan_nama: e.atasanNama ?? null,
  catatan_atasan: e.catatanAtasan ?? null,
  sent_at: e.sentAt ?? null,
  approved_at: e.approvedAt ?? null,
});

const fromRow = (r) => ({
  id: r.id,
  karyawanId: r.karyawan_id,
  userNama: r.user_nama,
  kategori: r.kategori,
  tipe: r.tipe,
  bagian: r.bagian,
  pekerjaan: r.pekerjaan,
  tujuan: r.tujuan,
  tanggalMulai: r.tanggal_mulai,
  jamMulai: r.jam_mulai,
  tanggalSelesai: r.tanggal_selesai,
  jamSelesai: r.jam_selesai,
  jenis: r.jenis,
  durasiJam: Number(r.durasi_jam || 0),
  estimasiGaji: Number(r.estimasi_gaji || 0),
  hasil: r.hasil,
  keterangan: r.keterangan,
  penanggungJawab: r.penanggung_jawab,
  status: r.status,
  evidence: r.evidence,
  ttdPekerja: r.ttd_pekerja,
  ttdAtasan: r.ttd_atasan,
  atasanNama: r.atasan_nama,
  catatanAtasan: r.catatan_atasan,
  sentAt: r.sent_at,
  approvedAt: r.approved_at,
});

// ---------- localStorage helpers ----------
const lsLoad = () => {
  try {
    const v = localStorage.getItem(LS_ENTRIES);
    return v ? JSON.parse(v) : null;
  } catch {
    return null;
  }
};
const lsSave = (arr) => {
  try {
    localStorage.setItem(LS_ENTRIES, JSON.stringify(arr));
  } catch (err) {
    console.warn("localStorage penuh?", err);
  }
};

// ============================================================
// Public API
// ============================================================
export { isSupabaseEnabled };

export async function listEntries(seedFallback = []) {
  if (isSupabaseEnabled) {
    const { data, error } = await supabase
      .from("lembur")
      .select("*")
      .order("tanggal_mulai", { ascending: false });
    if (error) throw error;
    return data.map(fromRow);
  }
  const local = lsLoad();
  return local ?? seedFallback;
}

export async function createEntry(entry) {
  if (isSupabaseEnabled) {
    const { data, error } = await supabase
      .from("lembur")
      .insert(toRow(entry))
      .select()
      .single();
    if (error) throw error;
    return fromRow(data);
  }
  // localStorage mode: caller mengelola array; kembalikan entri apa adanya
  return entry;
}

export async function updateEntry(id, patch) {
  if (isSupabaseEnabled) {
    const { data, error } = await supabase
      .from("lembur")
      .update(toRow(patch))
      .eq("id", id)
      .select()
      .single();
    if (error) throw error;
    return fromRow(data);
  }
  return { id, ...patch };
}

export async function deleteEntry(id) {
  if (isSupabaseEnabled) {
    const { error } = await supabase.from("lembur").delete().eq("id", id);
    if (error) throw error;
  }
}

// Simpan seluruh array (dipakai mode localStorage agar kompatibel dgn kode lama)
export function persistLocal(arr) {
  if (!isSupabaseEnabled) lsSave(arr);
}

// ---------- Karyawan roster ----------
export async function listKaryawan(seedFallback = []) {
  if (isSupabaseEnabled) {
    const { data, error } = await supabase
      .from("karyawan")
      .select("*")
      .order("no", { ascending: true });
    if (error) throw error;
    return data.map((k) => ({
      no: k.no,
      nama: k.nama,
      nomorInduk: k.nomor_induk,
      ttl: k.ttl,
      tahunMasuk: k.tahun_masuk,
      jabatan: k.jabatan,
      kategori: k.kategori,
      bolehLembur: k.boleh_lembur,
      gajiPokok: Number(k.gaji_pokok || 0),
    }));
  }
  return seedFallback;
}
