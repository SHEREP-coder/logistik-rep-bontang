import { createClient } from '@supabase/supabase-js';

const url = import.meta.env.VITE_SUPABASE_URL;
const key = import.meta.env.VITE_SUPABASE_ANON_KEY;
if (!url || !key) {
  // eslint-disable-next-line no-console
  console.error('VITE_SUPABASE_URL / VITE_SUPABASE_ANON_KEY belum diisi di file .env');
}
export const supabase = createClient(url || 'http://localhost', key || 'missing', {
  auth: { persistSession: true, autoRefreshToken: true }
});

// Klien kedua tanpa penyimpanan sesi: dipakai Admin untuk membuat akun baru
// tanpa mengeluarkan Admin yang sedang login.
export const supabaseSignup = createClient(url || 'http://localhost', key || 'missing', {
  auth: { persistSession: false, autoRefreshToken: false, storageKey: 'signup-tmp' }
});

/* ---------- AMBIL SEMUA BARIS (API Supabase membatasi 1000 baris per permintaan) ---------- */
export const fetchAll = async (build) => {
  const out = [];
  const size = 1000;
  for (let from = 0; ; from += size) {
    const { data, error } = await build().range(from, from + size - 1);
    if (error) throw error;
    out.push(...data);
    if (data.length < size) break;
  }
  return out;
};

/* ---------- PENERJEMAH: baris database <-> bentuk data lama (agar tampilan tidak perlu diubah) ---------- */
const nz = (v) => (v === undefined || v === null ? '' : v);
const num = (v) => (v === undefined || v === null || v === '' ? 0 : Number(v));

export const materialFromRow = (r) => ({
  id: r.kode, Kode: r.kode, 'Nama Barang': nz(r.nama_barang), 'Nomor Barang': nz(r.nomor_barang),
  'Part Unit': nz(r.part_unit), Merk: nz(r.merk), 'Base Unit of Measure': nz(r.base_unit),
  Satuan: nz(r.satuan), Lokasi: nz(r.lokasi), MinQty: num(r.min_qty), Qty: num(r.qty)
});

export const materialToRow = (m) => ({
  kode: m.Kode, nama_barang: m['Nama Barang'] ?? '', nomor_barang: m['Nomor Barang'] ?? null,
  part_unit: m['Part Unit'] ?? null, merk: m.Merk ?? null, base_unit: m['Base Unit of Measure'] ?? null,
  satuan: m.Satuan ?? null, lokasi: m.Lokasi || 'Rak A1', min_qty: num(m.MinQty ?? 5), qty: num(m.Qty)
});

export const masukFromRow = (r) => ({
  id: String(r.id), Tanggal: nz(r.tanggal), Periode: r.tanggal ? String(r.tanggal).slice(0, 7) : 'Tanpa-Tanggal',
  Unit: nz(r.unit), Kode: r.kode, 'Material Number': nz(r.material_number), Description: nz(r.description),
  Merk: nz(r.merk), 'No. PO': nz(r.no_po), 'No. Surat Jalan': nz(r.no_surat_jalan), Qty: num(r.qty),
  Satuan: nz(r.satuan), 'Kekurangan Supply': num(r.kekurangan_supply), Keterangan: nz(r.keterangan),
  poId: r.po_id ? String(r.po_id) : undefined
});

export const masukToRow = (f) => ({
  tanggal: f.Tanggal || null, unit: f.Unit ?? null, kode: f.Kode, material_number: f['Material Number'] ?? null,
  description: f.Description ?? null, merk: f.Merk ?? null, no_po: f['No. PO'] ?? null,
  no_surat_jalan: f['No. Surat Jalan'] ?? null, qty: num(f.Qty), satuan: f.Satuan ?? null,
  kekurangan_supply: num(f['Kekurangan Supply']), keterangan: f.Keterangan ?? null,
  ...(f.poId ? { po_id: Number(f.poId) } : {})
});

export const keluarFromRow = (r) => ({
  id: String(r.id), 'Doc.Date': nz(r.doc_date), Periode: r.doc_date ? String(r.doc_date).slice(0, 7) : 'Tanpa-Tanggal',
  Unit: nz(r.unit), Kode: r.kode, 'Nomor Barang': nz(r.nomor_barang), 'Nama Barang': nz(r.nama_barang),
  Mekanik: nz(r.mekanik), 'No. WO': nz(r.no_wo), HM_KM: nz(r.hm_km), Keterangan: nz(r.keterangan), Qty: num(r.qty)
});

export const keluarToRow = (f) => ({
  doc_date: f['Doc.Date'] || null, unit: f.Unit ?? null, kode: f.Kode, nomor_barang: f['Nomor Barang'] ?? null,
  nama_barang: f['Nama Barang'] ?? null, mekanik: f.Mekanik ?? null, no_wo: f['No. WO'] ?? null,
  hm_km: f.HM_KM ?? null, keterangan: f.Keterangan ?? null, qty: num(f.Qty)
});

export const poFromRow = (r) => ({
  id: String(r.id), noPo: nz(r.no_po), equipment: nz(r.equipment), kodeAlat: nz(r.kode_alat), lokasi: nz(r.lokasi),
  tanggal: nz(r.tanggal), diketahuiOleh: nz(r.diketahui_oleh),
  ttd1Nama: nz(r.ttd1_nama), ttd1Jabatan: nz(r.ttd1_jabatan), ttd2Nama: nz(r.ttd2_nama), ttd2Jabatan: nz(r.ttd2_jabatan),
  ttd3Nama: nz(r.ttd3_nama), ttd3Jabatan: nz(r.ttd3_jabatan), ditutup: !!r.ditutup, dibuatOleh: nz(r.dibuat_oleh),
  createdAt: r.created_at,
  items: (r.po_items || []).slice().sort((a, b) => a.urutan - b.urutan).map((i) => ({
    no: nz(i.no), uraian: nz(i.uraian), partNumber: nz(i.part_number), qty: nz(i.qty), unitPrice: nz(i.unit_price),
    keterangan: nz(i.keterangan), ...(num(i.diterima) ? { diterima: num(i.diterima) } : {})
  }))
});

export const poHeaderToRow = (h) => ({
  no_po: h.noPo, equipment: h.equipment || null, kode_alat: h.kodeAlat || null, lokasi: h.lokasi || null,
  tanggal: h.tanggal || null, diketahui_oleh: h.diketahuiOleh || null,
  ttd1_nama: h.ttd1Nama || null, ttd1_jabatan: h.ttd1Jabatan || null,
  ttd2_nama: h.ttd2Nama || null, ttd2_jabatan: h.ttd2Jabatan || null,
  ttd3_nama: h.ttd3Nama || null, ttd3_jabatan: h.ttd3Jabatan || null,
  ditutup: !!h.ditutup, dibuat_oleh: h.dibuatOleh || null
});

export const poItemsToRows = (poId, items) => items.map((i, idx) => ({
  po_id: Number(poId), urutan: idx, no: i.no ?? null, uraian: i.uraian ?? null, part_number: i.partNumber ?? null,
  qty: i.qty ?? null, unit_price: i.unitPrice ?? null, keterangan: i.keterangan ?? null, diterima: num(i.diterima)
}));

export const opnameFromRow = (r) => ({
  id: String(r.id), tanggal: nz(r.tanggal), createdAt: r.created_at, createdBy: nz(r.created_by),
  items: (r.stok_opname_items || []).map((i) => ({ kode: i.kode, nama: nz(i.nama), sistem: num(i.sistem), fisik: num(i.fisik), selisih: num(i.selisih) }))
});

export const userFromRow = (r) => ({
  id: r.id, nama: nz(r.nama), username: r.username, role: r.role, site: nz(r.site), status: r.status,
  email: `${r.username}@logistik-rep.local`
});

/* ---------- PEMUAT DATA ---------- */
export const loadMaterial = async () =>
  (await fetchAll(() => supabase.from('materials').select('*').order('kode'))).map(materialFromRow);

export const loadMasuk = async (fromDate) =>
  (await fetchAll(() => {
    let q = supabase.from('barang_masuk').select('*').order('id');
    if (fromDate) q = q.gte('tanggal', fromDate);
    return q;
  })).map(masukFromRow);

export const loadKeluar = async (fromDate) =>
  (await fetchAll(() => {
    let q = supabase.from('barang_keluar').select('*').order('id');
    if (fromDate) q = q.gte('doc_date', fromDate);
    return q;
  })).map(keluarFromRow);

export const loadPo = async () =>
  (await fetchAll(() => supabase.from('riwayat_po').select('*, po_items(*)').order('id'))).map(poFromRow);

export const loadOpname = async () =>
  (await fetchAll(() => supabase.from('stok_opname').select('*, stok_opname_items(*)').order('id'))).map(opnameFromRow);

export const loadUsers = async () =>
  (await fetchAll(() => supabase.from('profiles').select('*').order('username'))).map(userFromRow);

/* ---------- BANTU: pecah array untuk penulisan massal ---------- */
export const chunkArr = (arr, size) => {
  const out = [];
  for (let i = 0; i < arr.length; i += size) out.push(arr.slice(i, i + size));
  return out;
};

// Ubah error Postgres menjadi pesan yang mudah dibaca
export const pgMessage = (err) => {
  const m = String(err?.message || err || '');
  const mm = m.match(/STOK_MINUS: stok (.+?) tidak cukup/);
  if (mm) return `Stok ${mm[1]} tidak cukup (akan menjadi minus). Perubahan dibatalkan.`;
  if (m.includes('duplicate key') && m.includes('no_po')) return 'Nomor PO sudah dipakai di arsip.';
  if (m.includes('duplicate key')) return 'Data sudah ada (duplikat).';
  if (m.includes('row-level security')) return 'Akses ditolak.';
  return m;
};
