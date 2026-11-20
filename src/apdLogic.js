/* ====================================================================
   Aturan dan hitungan jadwal APD (sepatu dan seragam) - tanpa React/Supabase
   agar mudah diuji. Ubah aturan di bagian ATURAN di bawah bila kebijakan berubah.
   ==================================================================== */

/* ---------- ATURAN ---------- */
// Jabatan yang sepatunya dibagikan tiap 6 bulan (2 kali setahun). Jabatan lain: tiap 12 bulan.
export const JABATAN_SEPATU_6_BULAN = [
  'Mechanic DT', 'Mechanic A2B', 'Tyreman', 'Welder', 'Helper Mechanic', 'Chief Mechanic', 'GL Mechanic',
];
export const INTERVAL_SEPATU_DEFAULT = 12;  // bulan
export const INTERVAL_SEPATU_KHUSUS = 6;    // bulan
export const INTERVAL_SERAGAM = 12;         // bulan, semua jabatan
export const QTY_SEPATU = 1;                // pasang per orang tiap jadwal
export const QTY_SERAGAM = 2;               // potong per orang tiap jadwal

const norm = (s) => String(s || '').trim().replace(/\s+/g, ' ').toLowerCase();
const SET_6 = new Set(JABATAN_SEPATU_6_BULAN.map(norm));
export const intervalSepatu = (jabatan) => (SET_6.has(norm(jabatan)) ? INTERVAL_SEPATU_KHUSUS : INTERVAL_SEPATU_DEFAULT);

/* ---------- TANGGAL (string 'YYYY-MM-DD', tanpa zona waktu) ---------- */
export const pad2 = (n) => String(n).padStart(2, '0');
export const todayISO = () => { const d = new Date(); return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`; };
export const isoParts = (iso) => { const [y, m, d] = String(iso).slice(0, 10).split('-').map(Number); return { y, m, d }; };
export const daysInMonth = (y, m) => new Date(y, m, 0).getDate();
// Tambah bulan seperti EDATE di Excel: bila tanggal tidak ada di bulan tujuan, dipakai tanggal terakhir bulan itu
export const addMonths = (iso, n) => {
  const { y, m, d } = isoParts(iso);
  const t = y * 12 + (m - 1) + n;
  const ny = Math.floor(t / 12); const nm = (t % 12) + 1;
  return `${ny}-${pad2(nm)}-${pad2(Math.min(d, daysInMonth(ny, nm)))}`;
};
export const monthKey = (iso) => String(iso).slice(0, 7);

const maxDate = (arr) => arr.reduce((mx, r) => (!mx || r.tanggal > mx ? r.tanggal : mx), null);

/* ---------- JADWAL ----------
   Jadwal berikutnya = tanggal terima TERAKHIR + interval.
   Untuk sepatu, penggantian sepatu rusak ikut dihitung sebagai tanggal terima, sehingga hitungan berikutnya
   mulai dari tanggal penggantian. Bila belum pernah ada catatan, jadwal pertama = tanggal masuk kerja. */
export const hitungJadwal = (k, riwayat) => {
  const lastSepatu = maxDate(riwayat.filter((r) => r.jenis === 'sepatu'));
  const lastSeragam = maxDate(riwayat.filter((r) => r.jenis === 'seragam'));
  const ivSepatu = intervalSepatu(k.jabatan);
  return {
    intervalSepatu: ivSepatu,
    intervalSeragam: INTERVAL_SERAGAM,
    lastSepatu, lastSeragam,
    nextSepatu: lastSepatu ? addMonths(lastSepatu, ivSepatu) : (k.tanggal_masuk || null),
    nextSeragam: lastSeragam ? addMonths(lastSeragam, INTERVAL_SERAGAM) : (k.tanggal_masuk || null),
    belumPernahSepatu: !lastSepatu,
    belumPernahSeragam: !lastSeragam,
  };
};

// Status satu jadwal terhadap bulan acuan (ym = 'YYYY-MM'): 'terlambat' | 'bulan' | 'aman' | null
export const statusJadwal = (next, ym) => {
  if (!next) return null;
  const nk = monthKey(next);
  return nk < ym ? 'terlambat' : nk === ym ? 'bulan' : 'aman';
};
// Untuk daftar orderan sebuah bulan: ikut bila jatuh di bulan itu, atau (bila diminta) terlambat dari bulan sebelumnya
export const masukOrderan = (next, ym, sertakanTerlambat) => {
  const s = statusJadwal(next, ym);
  return s === 'bulan' || (sertakanTerlambat && s === 'terlambat') ? s : null;
};

/* ---------- UKURAN ---------- */
const URUT_SERAGAM = ['XS', 'S', 'M', 'L', 'XL', 'XXL', '2XL', '3XL', 'XXXL', '4XL'];
export const urutUkuran = (jenis) => (a, b) => {
  if (jenis === 'sepatu') {
    const na = parseFloat(a); const nb = parseFloat(b);
    if (!isNaN(na) && !isNaN(nb)) return na - nb;
  } else {
    const ia = URUT_SERAGAM.indexOf(a); const ib = URUT_SERAGAM.indexOf(b);
    if (ia > -1 && ib > -1) return ia - ib;
    if (ia > -1) return -1;
    if (ib > -1) return 1;
  }
  return String(a).localeCompare(String(b));
};
// Rekap jumlah per ukuran: list = [{ k }], field = 'ukuran_sepatu' | 'ukuran_seragam', qty = jumlah per orang
export const rekapUkuran = (list, field, jenis, qty) => {
  const m = new Map();
  list.forEach(({ k }) => {
    const key = String(k[field] || '').trim().toUpperCase() || '(belum diisi)';
    m.set(key, (m.get(key) || 0) + qty);
  });
  return Array.from(m.entries()).sort((a, b) => {
    if (a[0] === '(belum diisi)') return 1;
    if (b[0] === '(belum diisi)') return -1;
    return urutUkuran(jenis)(a[0], b[0]);
  });
};
