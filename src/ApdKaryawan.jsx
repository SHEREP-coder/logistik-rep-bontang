import React, { useState, useEffect, useMemo, useCallback } from 'react';
import * as XLSX from 'xlsx';
import { Search, Plus, Pencil, Download, Printer, History, X, Check, Trash2, RefreshCw, Upload } from 'lucide-react';
import { supabase, pgMessage } from './supabase';
import { parseRekap } from './apdImport';
import { useTabel, Th, BtnReset, TabelRekapUkuran, rankUkuran } from './ApdTabel';
import {
  JABATAN_SEPATU_6_BULAN, QTY_SEPATU, QTY_SERAGAM, INTERVAL_SERAGAM,
  pad2, todayISO, addMonths, monthKey, hitungJadwal, statusJadwal, masukOrderan, rekapUkuran,
} from './apdLogic';

/* ============================== KONSTANTA ============================== */
const NAMA_BULAN = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
const BLN3 = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des'];
const UKURAN_SERAGAM = ['S', 'M', 'L', 'XL', 'XXL', '3XL'];
const BARIS = 20;
// Tanda tangan pada cetakan orderan (sesuai lembar PO APD di Excel). Ubah di sini bila berganti pejabat.
const TTD = [
  { judul: 'Diketahui oleh,', nama: 'Catur. A', jabatan: 'Site Manager' },
  { judul: 'Check oleh,', nama: 'Tri Susanto', jabatan: 'Safety Officer' },
  { judul: 'Dibuat oleh,', nama: 'Aliuddin', jabatan: 'Logistic' },
];

// Tanda tangan pada lembar Tanda Terima
const TTD_TERIMA = [
  { judul: 'Diserahkan oleh,', nama: 'Aliuddin', jabatan: 'Logistic' },
  { judul: 'Diketahui oleh,', nama: 'Tri Susanto', jabatan: 'Safety Officer' },
];

const fmt = (iso) => {
  if (!iso) return '-';
  const [y, m, d] = String(iso).slice(0, 10).split('-').map(Number);
  return `${d} ${BLN3[m - 1]} ${y}`;
};
const tglTampil = (v) => (v ? fmt(v) : '(kosong)');
const fmtPanjang = (iso) => {
  const [y, m, d] = String(iso).slice(0, 10).split('-').map(Number);
  return `${d} ${NAMA_BULAN[m - 1]} ${y}`;
};

// Baca semua baris tabel (Supabase membatasi 1000 baris per permintaan)
const loadAll = async (table) => {
  const out = []; const size = 1000;
  for (let from = 0; ; from += size) {
    const { data, error } = await supabase.from(table).select('*').order('id', { ascending: true }).range(from, from + size - 1);
    if (error) throw error;
    out.push(...data);
    if (data.length < size) break;
  }
  return out;
};

const emptyForm = () => ({
  id: null, nrp: '', nrp_amodi: '', nama: '', jabatan: '', tanggal_masuk: '', ukuran_sepatu: '', ukuran_seragam: '',
  keterangan: '', terima_sepatu: '', terima_seragam: '',
  // hanya terisi saat mengubah: tanggal dan id catatan terakhir sebelum diubah, agar koreksi bisa dikenali
  awal_sepatu: '', awal_seragam: '', rec_sepatu: null, rec_seragam: null, kat_sepatu: '', kat_seragam: '',
});

// Form ubah data: isi dari data karyawan + catatan terima terakhir (riwayat sudah terurut dari yang terbaru)
const formDari = (k, riwayat) => {
  const base = { ...emptyForm(), ...Object.fromEntries(Object.entries(k).map(([a, b]) => [a, b ?? ''])) };
  const sp = riwayat.find((r) => r.jenis === 'sepatu');
  const sg = riwayat.find((r) => r.jenis === 'seragam');
  return {
    ...base,
    terima_sepatu: sp ? sp.tanggal : '', awal_sepatu: sp ? sp.tanggal : '', rec_sepatu: sp ? sp.id : null, kat_sepatu: sp ? sp.kategori : '',
    terima_seragam: sg ? sg.tanggal : '', awal_seragam: sg ? sg.tanggal : '', rec_seragam: sg ? sg.id : null, kat_seragam: sg ? sg.kategori : '',
  };
};

/* ============================== KOMPONEN KECIL ============================== */
function Modal({ title, onClose, dark, children, wide }) {
  return (
    <div className="fixed inset-0 bg-black bg-opacity-60 flex items-center justify-center z-50 p-4 no-print">
      <div className={`w-full ${wide ? 'max-w-3xl' : 'max-w-lg'} max-h-[92vh] overflow-y-auto p-6 rounded-xl shadow-2xl border ${dark ? 'bg-slate-800 border-slate-700 text-white' : 'bg-white border-slate-200 text-slate-800'}`}>
        <div className="flex justify-between items-center mb-4 border-b pb-2">
          <h3 className="font-bold text-base">{title}</h3>
          <button onClick={onClose} aria-label="Tutup"><X size={20} /></button>
        </div>
        {children}
      </div>
    </div>
  );
}

const Label = ({ text, children, className = '' }) => (
  <div className={className}><label className="block font-semibold mb-1 text-xs">{text}</label>{children}</div>
);

/* ============================== HALAMAN ============================== */
export default function ApdKaryawan({ canEdit, darkMode, notify, profile }) {
  const [tab, setTab] = useState('daftar');
  const [karyawan, setKaryawan] = useState([]);
  const [pembagian, setPembagian] = useState([]);
  const [loading, setLoading] = useState(true);
  const [loadError, setLoadError] = useState('');
  const [saving, setSaving] = useState(false);

  // daftar karyawan
  const [cari, setCari] = useState('');
  const [filterAktif, setFilterAktif] = useState('aktif');
  const [page, setPage] = useState(1);

  // orderan bulanan
  const now = new Date();
  const [bln, setBln] = useState(now.getMonth() + 1);
  const [thn, setThn] = useState(now.getFullYear());
  const [sertakanTerlambat, setSertakanTerlambat] = useState(true);

  // rekap
  const [thnRekap, setThnRekap] = useState(now.getFullYear());

  // tanda terima
  const [selTerima, setSelTerima] = useState(new Set());
  const [tglTerima, setTglTerima] = useState(todayISO());

  // resign
  const [resign, setResign] = useState(null); // { k, tanggal }

  // impor Excel
  const [imp, setImp] = useState(null);
  const [impBusy, setImpBusy] = useState(false);
  const [impUpdate, setImpUpdate] = useState(false);

  // modal
  const [form, setForm] = useState(null);          // tambah / ubah karyawan
  const [riwayatId, setRiwayatId] = useState(null); // riwayat per karyawan
  const [rusak, setRusak] = useState(null);        // { k, tanggal, keterangan }
  const [formRiwayat, setFormRiwayat] = useState({ jenis: 'sepatu', kategori: 'rutin', tanggal: todayISO(), keterangan: '' });

  const oleh = profile?.nama || profile?.username || '';
  const today = todayISO();
  const ymNow = monthKey(today);
  const ym = `${thn}-${pad2(bln)}`;

  /* ---------- muat data ---------- */
  const muat = useCallback(async () => {
    setLoading(true); setLoadError('');
    try {
      const [k, p] = await Promise.all([loadAll('apd_karyawan'), loadAll('apd_pembagian')]);
      setKaryawan(k); setPembagian(p);
    } catch (err) {
      console.error(err);
      setLoadError(pgMessage(err));
    }
    setLoading(false);
  }, []);
  useEffect(() => { muat(); }, [muat]);

  /* ---------- hitung jadwal setiap karyawan ---------- */
  const data = useMemo(() => {
    const by = new Map();
    pembagian.forEach((p) => { const a = by.get(p.karyawan_id) || []; a.push(p); by.set(p.karyawan_id, a); });
    return karyawan.map((k) => {
      const riwayat = (by.get(k.id) || []).slice().sort((a, b) => (a.tanggal < b.tanggal ? 1 : a.tanggal > b.tanggal ? -1 : b.id - a.id));
      return { k, riwayat, j: hitungJadwal(k, riwayat) };
    });
  }, [karyawan, pembagian]);

  const jabatanList = useMemo(() => {
    const s = new Set(JABATAN_SEPATU_6_BULAN);
    karyawan.forEach((k) => { if (k.jabatan) s.add(k.jabatan); });
    return Array.from(s).sort();
  }, [karyawan]);

  /* ---------- daftar karyawan: saring, urut, ringkasan ---------- */
  const statusDaftar = ({ k, j }) => {
    if (!k.aktif) return 'Resign';
    const a = statusJadwal(j.nextSepatu, ymNow); const g = statusJadwal(j.nextSeragam, ymNow);
    return a === 'terlambat' || g === 'terlambat' ? 'Terlambat' : a === 'bulan' || g === 'bulan' ? 'Bulan ini' : 'Aman';
  };
  const ukS = (r) => String(r.k.ukuran_sepatu || '').trim();
  const ukG = (r) => String(r.k.ukuran_seragam || '').trim().toUpperCase();
  const kolomDaftar = [
    { key: 'amodi', label: 'NRP AMODI', ambil: (r) => r.k.nrp_amodi || r.k.nrp },
    { key: 'nama', label: 'NAMA', kiri: true, ambil: (r) => r.k.nama },
    { key: 'jabatan', label: 'JABATAN', kiri: true, ambil: (r) => r.k.jabatan || '' },
    { key: 'uks', label: 'SEPATU', ambil: ukS, urut: (r) => rankUkuran('sepatu', ukS(r)) },
    { key: 'ukg', label: 'SERAGAM', ambil: ukG, urut: (r) => rankUkuran('seragam', ukG(r)) },
    { key: 'ls', label: 'TERIMA SEPATU\nTERAKHIR', ambil: (r) => r.j.lastSepatu || '', tampil: tglTampil },
    { key: 'ns', label: 'SEPATU\nBERIKUTNYA', ambil: (r) => r.j.nextSepatu || '', tampil: tglTampil },
    { key: 'lg', label: 'TERIMA SERAGAM\nTERAKHIR', ambil: (r) => r.j.lastSeragam || '', tampil: tglTampil },
    { key: 'ng', label: 'SERAGAM\nBERIKUTNYA', ambil: (r) => r.j.nextSeragam || '', tampil: tglTampil },
    { key: 'status', label: 'STATUS', ambil: statusDaftar },
  ];
  const tabDaftar = useTabel(kolomDaftar);
  // dasar: filter aktif/resign dan kotak cari, urut nama; lalu filter dan urut dari header kolom
  const daftarDasar = useMemo(() => {
    const q = cari.trim().toLowerCase();
    let rows = data.filter(({ k }) => (filterAktif === 'semua' ? true : filterAktif === 'aktif' ? k.aktif : !k.aktif));
    if (q) rows = rows.filter(({ k }) => [k.nrp, k.nrp_amodi, k.nama, k.jabatan].some((v) => String(v || '').toLowerCase().includes(q)));
    return rows.slice().sort((a, b) => a.k.nama.localeCompare(b.k.nama));
  }, [data, cari, filterAktif]);
  const daftar = useMemo(() => tabDaftar.terapkan(daftarDasar), [daftarDasar, tabDaftar.sort, tabDaftar.filter]); // eslint-disable-line
  useEffect(() => { setPage(1); }, [tabDaftar.sort, tabDaftar.filter]);

  const ringkas = useMemo(() => {
    const aktif = data.filter(({ k }) => k.aktif);
    const st = (j) => [statusJadwal(j.nextSepatu, ymNow), statusJadwal(j.nextSeragam, ymNow)];
    return {
      aktif: aktif.length,
      bulanIni: aktif.filter(({ j }) => st(j).includes('bulan')).length,
      terlambat: aktif.filter(({ j }) => st(j).includes('terlambat')).length,
    };
  }, [data, ymNow]);

  const totalPage = Math.max(1, Math.ceil(daftar.length / BARIS));
  const pageNow = Math.min(page, totalPage);
  const daftarHal = daftar.slice((pageNow - 1) * BARIS, pageNow * BARIS);

  /* ---------- orderan bulanan ---------- */
  const orderan = useMemo(() => {
    const sepatu = []; const seragam = [];
    data.forEach(({ k, j }) => {
      if (!k.aktif) return;
      const s = masukOrderan(j.nextSepatu, ym, sertakanTerlambat);
      if (s) sepatu.push({ k, j, status: s, tgl: j.nextSepatu });
      const g = masukOrderan(j.nextSeragam, ym, sertakanTerlambat);
      if (g) seragam.push({ k, j, status: g, tgl: j.nextSeragam });
    });
    const byTgl = (a, b) => (a.tgl < b.tgl ? -1 : a.tgl > b.tgl ? 1 : a.k.nama.localeCompare(b.k.nama));
    return { sepatu: sepatu.sort(byTgl), seragam: seragam.sort(byTgl) };
  }, [data, ym, sertakanTerlambat]);

  const rekapSepatu = useMemo(() => rekapUkuran(orderan.sepatu, 'ukuran_sepatu', 'sepatu', QTY_SEPATU), [orderan]);
  const rekapSeragam = useMemo(() => rekapUkuran(orderan.seragam, 'ukuran_seragam', 'seragam', QTY_SERAGAM), [orderan]);
  const totSepatu = rekapSepatu.reduce((a, [, n]) => a + n, 0);
  const totSeragam = rekapSeragam.reduce((a, [, n]) => a + n, 0);

  // gabungan satu baris per orang (format lembar PO APD) untuk cetak dan Excel
  const gabungan = useMemo(() => {
    const m = new Map();
    orderan.sepatu.forEach((r) => m.set(r.k.id, { k: r.k, j: r.j, sepatu: r, seragam: null }));
    orderan.seragam.forEach((r) => { const e = m.get(r.k.id) || { k: r.k, j: r.j, sepatu: null, seragam: null }; e.seragam = r; m.set(r.k.id, e); });
    return Array.from(m.values()).sort((a, b) => a.k.nama.localeCompare(b.k.nama));
  }, [orderan]);

  /* ---------- rekap ---------- */
  const rekap = useMemo(() => {
    const aktif = data.filter(({ k }) => k.aktif);
    const yStr = String(thnRekap);
    // pembagian per bulan pada tahun terpilih (semua catatan riwayat)
    const bulan = Array.from({ length: 12 }, (_, i) => ({ bln: i + 1, sepatuRutin: 0, sepatuRusak: 0, seragam: 0 }));
    pembagian.forEach((p) => {
      if (String(p.tanggal).slice(0, 4) !== yStr) return;
      const b = bulan[Number(String(p.tanggal).slice(5, 7)) - 1];
      if (p.jenis === 'seragam') b.seragam += 1; else if (p.kategori === 'rusak') b.sepatuRusak += 1; else b.sepatuRutin += 1;
    });
    // per jabatan (karyawan aktif)
    const jab = new Map();
    aktif.forEach(({ k, j, riwayat }) => {
      const nm = k.jabatan || '(tanpa jabatan)';
      const e = jab.get(nm) || { jabatan: nm, orang: 0, iv: j.intervalSepatu, sepatu: 0, rusak: 0, seragam: 0, telatSepatu: 0, telatSeragam: 0 };
      e.orang += 1;
      riwayat.forEach((r) => {
        if (String(r.tanggal).slice(0, 4) !== yStr) return;
        if (r.jenis === 'seragam') e.seragam += 1; else if (r.kategori === 'rusak') e.rusak += 1; else e.sepatu += 1;
      });
      if (statusJadwal(j.nextSepatu, ymNow) === 'terlambat') e.telatSepatu += 1;
      if (statusJadwal(j.nextSeragam, ymNow) === 'terlambat') e.telatSeragam += 1;
      jab.set(nm, e);
    });
    // proyeksi kebutuhan 12 bulan ke depan dari jadwal saat ini (siklus berikutnya = jadwal + interval)
    const months = Array.from({ length: 12 }, (_, i) => monthKey(addMonths(`${ymNow}-01`, i)));
    const akhir = months[11];
    const proy = { terlambat: { sepatu: 0, seragam: 0 }, bulan: new Map(months.map((mo) => [mo, { sepatu: 0, seragam: 0 }])) };
    aktif.forEach(({ j }) => {
      [['sepatu', j.nextSepatu, j.intervalSepatu], ['seragam', j.nextSeragam, j.intervalSeragam]].forEach(([jn, next, iv]) => {
        if (!next) return;
        let telat = false;
        for (let n = 0; n < 60; n++) {
          const mk = monthKey(addMonths(next, iv * n));
          if (mk > akhir) break;
          if (mk < ymNow) { if (!telat) { proy.terlambat[jn] += 1; telat = true; } continue; }
          proy.bulan.get(mk)[jn] += 1;
        }
      });
    });
    return {
      aktif: aktif.length, bulan, months, proy,
      jabatan: Array.from(jab.values()).sort((a, b) => b.orang - a.orang || a.jabatan.localeCompare(b.jabatan)),
      ukSepatu: rekapUkuran(aktif, 'ukuran_sepatu', 'sepatu', 1),
      ukSeragam: rekapUkuran(aktif, 'ukuran_seragam', 'seragam', 1),
    };
  }, [data, pembagian, thnRekap, ymNow]);
  const sumBulan = (f) => rekap.bulan.reduce((a, b) => a + f(b), 0);
  const labelBulan = (mo) => `${BLN3[Number(mo.slice(5, 7)) - 1]} ${mo.slice(0, 4)}`;

  /* ---------- urut, filter, cari pada header kolom: orderan, tanda terima, rekap ---------- */
  const kolomOrderan = (jenis) => {
    const uk = (r) => String((jenis === 'sepatu' ? r.k.ukuran_sepatu : r.k.ukuran_seragam) || '').trim().toUpperCase();
    return [
      { key: 'amodi', label: 'NRP AMODI', ambil: (r) => r.k.nrp_amodi || r.k.nrp },
      { key: 'nama', label: 'NAMA', kiri: true, ambil: (r) => r.k.nama },
      { key: 'jabatan', label: 'JABATAN', kiri: true, ambil: (r) => r.k.jabatan || '' },
      { key: 'uk', label: 'UKURAN', ambil: uk, urut: (r) => rankUkuran(jenis, uk(r)) },
      { key: 'tgl', label: 'JADWAL', ambil: (r) => r.tgl || '', tampil: tglTampil },
      { key: 'last', label: 'TERIMA TERAKHIR', ambil: (r) => (jenis === 'sepatu' ? r.j.lastSepatu : r.j.lastSeragam) || '', tampil: tglTampil },
      { key: 'status', label: 'STATUS', ambil: (r) => (r.status === 'terlambat' ? 'Terlambat' : 'Bulan ini') },
    ];
  };
  const tabSepatu = useTabel(kolomOrderan('sepatu'));
  const tabSeragam = useTabel(kolomOrderan('seragam'));
  const sepatuTampil = useMemo(() => tabSepatu.terapkan(orderan.sepatu), [orderan, tabSepatu.sort, tabSepatu.filter]); // eslint-disable-line
  const seragamTampil = useMemo(() => tabSeragam.terapkan(orderan.seragam), [orderan, tabSeragam.sort, tabSeragam.filter]); // eslint-disable-line

  const ukTerima = (g, jenis) => String((jenis === 'sepatu' ? g.sepatu && (g.k.ukuran_sepatu || '?') : g.seragam && (g.k.ukuran_seragam || '?')) || '').trim().toUpperCase();
  const kolomTerima = [
    { key: 'amodi', label: 'NRP AMODI', ambil: (g) => g.k.nrp_amodi || g.k.nrp },
    { key: 'nama', label: 'NAMA', kiri: true, ambil: (g) => g.k.nama },
    { key: 'jabatan', label: 'JABATAN', kiri: true, ambil: (g) => g.k.jabatan || '' },
    { key: 'sepatu', label: 'SEPATU', ambil: (g) => ukTerima(g, 'sepatu'), tampil: (v) => v || '(tidak ada)', urut: (g) => (g.sepatu ? rankUkuran('sepatu', ukTerima(g, 'sepatu')) : '') },
    { key: 'seragam', label: 'SERAGAM', ambil: (g) => ukTerima(g, 'seragam'), tampil: (v) => v || '(tidak ada)', urut: (g) => (g.seragam ? rankUkuran('seragam', ukTerima(g, 'seragam')) : '') },
    { key: 'jadwal', label: 'JADWAL', ambil: (g) => [g.sepatu && g.sepatu.tgl, g.seragam && g.seragam.tgl].filter(Boolean).sort()[0] || '', tampil: tglTampil },
    { key: 'status', label: 'STATUS', ambil: (g) => (g.sepatu?.status === 'terlambat' || g.seragam?.status === 'terlambat' ? 'Terlambat' : 'Bulan ini') },
  ];
  const tabTerima = useTabel(kolomTerima);
  const gabunganTampil = useMemo(() => tabTerima.terapkan(gabungan), [gabungan, tabTerima.sort, tabTerima.filter]); // eslint-disable-line

  const kolomBulan = [
    { key: 'bln', label: 'BULAN', kiri: true, ambil: (b) => b.bln, tampil: (v) => NAMA_BULAN[v - 1] },
    { key: 'rutin', label: 'SEPATU RUTIN\n(psg)', ambil: (b) => b.sepatuRutin },
    { key: 'rusak', label: 'SEPATU GANTI RUSAK\n(psg)', ambil: (b) => b.sepatuRusak },
    { key: 'tot', label: 'TOTAL SEPATU\n(psg)', ambil: (b) => b.sepatuRutin + b.sepatuRusak },
    { key: 'sg', label: 'SERAGAM\n(orang)', ambil: (b) => b.seragam },
    { key: 'pcs', label: 'SERAGAM\n(pcs)', ambil: (b) => b.seragam * QTY_SERAGAM },
  ];
  const tabBulan = useTabel(kolomBulan);
  const bulanTampil = useMemo(() => tabBulan.terapkan(rekap.bulan), [rekap, tabBulan.sort, tabBulan.filter]); // eslint-disable-line

  const proyBaris = useMemo(() => [
    { i: -1, label: 'Terlambat', telat: true, sepatu: rekap.proy.terlambat.sepatu, seragam: rekap.proy.terlambat.seragam },
    ...rekap.months.map((mo, i) => ({ i, label: labelBulan(mo), telat: false, sepatu: rekap.proy.bulan.get(mo).sepatu, seragam: rekap.proy.bulan.get(mo).seragam })),
  ], [rekap]); // eslint-disable-line
  const kolomProy = [
    { key: 'per', label: 'PERIODE', kiri: true, ambil: (r) => r.label, urut: (r) => r.i },
    { key: 'sp', label: 'SEPATU\n(orang)', ambil: (r) => r.sepatu },
    { key: 'sg', label: 'SERAGAM\n(orang)', ambil: (r) => r.seragam },
    { key: 'pcs', label: 'SERAGAM\n(pcs)', ambil: (r) => r.seragam * QTY_SERAGAM },
  ];
  const tabProy = useTabel(kolomProy);
  const proyTampil = useMemo(() => tabProy.terapkan(proyBaris), [proyBaris, tabProy.sort, tabProy.filter]); // eslint-disable-line

  const kolomJab = [
    { key: 'jab', label: 'JABATAN', kiri: true, ambil: (e) => e.jabatan },
    { key: 'orang', label: 'AKTIF', ambil: (e) => e.orang },
    { key: 'iv', label: 'SEPATU\nTIAP', ambil: (e) => e.iv, tampil: (v) => `${v} bln` },
    { key: 'sp', label: 'SEPATU RUTIN\nDITERIMA', ambil: (e) => e.sepatu },
    { key: 'rusak', label: 'GANTI\nRUSAK', ambil: (e) => e.rusak },
    { key: 'sg', label: 'SERAGAM\nDITERIMA', ambil: (e) => e.seragam },
    { key: 'ts', label: 'SEPATU\nTERLAMBAT', ambil: (e) => e.telatSepatu },
    { key: 'tg', label: 'SERAGAM\nTERLAMBAT', ambil: (e) => e.telatSeragam },
  ];
  const tabJab = useTabel(kolomJab);
  const jabTampil = useMemo(() => tabJab.terapkan(rekap.jabatan), [rekap, tabJab.sort, tabJab.filter]); // eslint-disable-line

  useEffect(() => {
    setSelTerima((s) => { const ids = new Set(gabungan.map((g) => g.k.id)); const n = new Set([...s].filter((id) => ids.has(id))); return n.size === s.size ? s : n; });
  }, [gabungan]);

  /* ---------- simpan karyawan ---------- */
  const simpanKaryawan = async () => {
    const f = form;
    const nama = f.nama.trim(); const amodi = f.nrp_amodi.trim();
    if (!nama || !amodi) return notify('NRP Amodi dan Nama wajib diisi.', 'error');
    const row = {
      nrp: f.nrp.trim() || amodi, nrp_amodi: amodi, nama, jabatan: f.jabatan.trim() || null,
      tanggal_masuk: f.tanggal_masuk || null, ukuran_sepatu: f.ukuran_sepatu.trim() || null,
      ukuran_seragam: f.ukuran_seragam.trim().toUpperCase() || null, keterangan: f.keterangan.trim() || null,
    };
    // koreksi tanda terima terakhir (hanya saat mengubah data): ubah tanggal, hapus, atau tambah catatan
    const koreksi = [];
    if (f.id) {
      for (const jenis of ['sepatu', 'seragam']) {
        const baru = f[`terima_${jenis}`] || ''; const lama = f[`awal_${jenis}`] || ''; const rec = f[`rec_${jenis}`];
        if (baru === lama) continue;
        if (baru && baru > today) return notify(`Tanggal terima ${jenis} melewati hari ini.`, 'error');
        if (!baru && rec) {
          if (!window.confirm(`Kosongkan tanda terima ${jenis} terakhir (${fmt(lama)})? Catatan itu akan dihapus dan jadwal dihitung dari catatan sebelumnya.`)) return;
          koreksi.push({ op: 'hapus', id: rec });
        } else if (baru && rec) {
          const kat = f[`kat_${jenis}`];
          if (pembagian.some((p) => p.karyawan_id === f.id && p.jenis === jenis && p.kategori === kat && p.tanggal === baru && p.id !== rec)) {
            return notify(`Tanda terima ${jenis} tanggal ${fmt(baru)} sudah tercatat di riwayat.`, 'error');
          }
          koreksi.push({ op: 'ubah', id: rec, tanggal: baru });
        } else if (baru && !rec) {
          koreksi.push({ op: 'tambah', jenis, tanggal: baru });
        }
      }
    }
    setSaving(true);
    try {
      if (f.id) {
        const { error } = await supabase.from('apd_karyawan').update({ ...row, updated_at: new Date().toISOString(), updated_by: oleh }).eq('id', f.id);
        if (error) throw error;
        for (const kx of koreksi) {
          let res;
          if (kx.op === 'hapus') res = await supabase.from('apd_pembagian').delete().eq('id', kx.id);
          else if (kx.op === 'ubah') res = await supabase.from('apd_pembagian').update({ tanggal: kx.tanggal }).eq('id', kx.id);
          else res = await supabase.from('apd_pembagian').insert({ karyawan_id: f.id, jenis: kx.jenis, kategori: 'rutin', tanggal: kx.tanggal, created_by: oleh });
          if (res.error) throw res.error;
        }
      } else {
        const { data: ins, error } = await supabase.from('apd_karyawan').insert({ ...row, created_by: oleh }).select().single();
        if (error) throw error;
        const awal = [];
        if (f.terima_sepatu) awal.push({ karyawan_id: ins.id, jenis: 'sepatu', kategori: 'rutin', tanggal: f.terima_sepatu, created_by: oleh });
        if (f.terima_seragam) awal.push({ karyawan_id: ins.id, jenis: 'seragam', kategori: 'rutin', tanggal: f.terima_seragam, created_by: oleh });
        if (awal.length) {
          const { error: e2 } = await supabase.from('apd_pembagian').insert(awal);
          if (e2) throw e2;
        }
      }
      notify(f.id ? (koreksi.length ? 'Data karyawan dan tanda terima diperbarui. Jadwal berikutnya dihitung ulang.' : 'Data karyawan diperbarui.') : 'Karyawan ditambahkan.');
      setForm(null);
      await muat();
    } catch (err) {
      notify('Gagal menyimpan: ' + pgMessage(err), 'error');
    }
    setSaving(false);
  };

  // Tandai resign: karyawan tidak muncul lagi di orderan, tanda terima, rekap, dan proyeksi. Riwayat tetap tersimpan.
  const simpanResign = async () => {
    const { k, tanggal } = resign;
    if (!tanggal) return notify('Tanggal resign wajib diisi.', 'error');
    setSaving(true);
    try {
      const { error } = await supabase.from('apd_karyawan').update({ aktif: false, tanggal_resign: tanggal, updated_at: new Date().toISOString(), updated_by: oleh }).eq('id', k.id);
      if (error) throw error;
      notify(`${k.nama} ditandai resign (${fmt(tanggal)}) dan tidak lagi muncul di pemantauan.`);
      setResign(null);
      await muat();
    } catch (err) {
      const pesan = pgMessage(err);
      notify(/tanggal_resign/i.test(pesan) ? 'Kolom tanggal_resign belum ada. Jalankan schema_v2_patch5.sql di Supabase dulu.' : 'Gagal menyimpan: ' + pesan, 'error');
    }
    setSaving(false);
  };

  const aktifkanKembali = async (k) => {
    if (!window.confirm(`Aktifkan kembali ${k.nama}? Ia akan muncul lagi di pemantauan sepatu dan seragam.`)) return;
    setSaving(true);
    try {
      const { error } = await supabase.from('apd_karyawan').update({ aktif: true, tanggal_resign: null, updated_at: new Date().toISOString(), updated_by: oleh }).eq('id', k.id);
      if (error) throw error;
      notify(`${k.nama} diaktifkan kembali.`);
      await muat();
    } catch (err) { notify('Gagal mengaktifkan: ' + pgMessage(err), 'error'); }
    setSaving(false);
  };

  /* ---------- catat pembagian ---------- */
  // items: [{ karyawan_id, jenis }]; mengembalikan true bila tersimpan
  const catat = async (items, tanggal, kategori = 'rutin', keterangan = null) => {
    if (!tanggal && !items.every((it) => it.tanggal)) { notify('Tanggal terima wajib diisi.', 'error'); return false; }
    const ada = new Set(pembagian.map((p) => `${p.karyawan_id}|${p.jenis}|${p.kategori}|${p.tanggal}`));
    const baru = items
      .map((it) => ({ ...it, tanggal: it.tanggal || tanggal }))
      .filter((it) => !ada.has(`${it.karyawan_id}|${it.jenis}|${kategori}|${it.tanggal}`))
      .map((it) => ({ ...it, kategori, keterangan: keterangan || null, created_by: oleh }));
    if (!baru.length) { notify('Pembagian itu sudah tercatat pada tanggal tersebut.', 'error'); return false; }
    setSaving(true);
    let ok = false;
    try {
      const { error } = await supabase.from('apd_pembagian').insert(baru);
      if (error) throw error;
      ok = true;
      await muat();
    } catch (err) { notify('Gagal mencatat: ' + pgMessage(err), 'error'); }
    setSaving(false);
    return ok;
  };

  const simpanRusak = async () => {
    const ok = await catat([{ karyawan_id: rusak.k.id, jenis: 'sepatu' }], rusak.tanggal, 'rusak', rusak.keterangan.trim());
    if (!ok) return;
    const iv = hitungJadwal(rusak.k, []).intervalSepatu;
    notify(`Penggantian sepatu rusak ${rusak.k.nama} tercatat. Jadwal sepatu berikutnya ${fmt(addMonths(rusak.tanggal, iv))}.`);
    setRusak(null);
  };

  const hapusRiwayat = async (r) => {
    if (!window.confirm(`Hapus catatan ${r.jenis} ${r.kategori === 'rusak' ? '(ganti rusak) ' : ''}tanggal ${fmt(r.tanggal)}? Jadwal berikutnya akan dihitung ulang.`)) return;
    setSaving(true);
    try {
      const { error } = await supabase.from('apd_pembagian').delete().eq('id', r.id);
      if (error) throw error;
      notify('Catatan dihapus.');
      await muat();
    } catch (err) { notify('Gagal menghapus: ' + pgMessage(err), 'error'); }
    setSaving(false);
  };

  /* ---------- unduh Excel ---------- */
  const unduhExcel = () => {
    const judul = `${NAMA_BULAN[bln - 1]} ${thn}`;
    const aoa = [
      ['PT. RYAN EKA PRATAMA'], ['BONTANG COAL MINE PROJECT'], ['RINCIAN ORDERAN APD KARYAWAN'], [`Periode: ${judul}`], [],
      ['No', 'NRP Amodi', 'Nama', 'Jabatan', 'Tanggal Masuk', 'Sepatu Size', 'Sepatu Qty', 'Jadwal Sepatu', 'Baju Size', 'Baju Qty', 'Jadwal Seragam', 'Terima Terakhir Sepatu', 'Terima Terakhir Seragam', 'Keterangan'],
    ];
    gabungan.forEach((g, i) => {
      const tl = [g.sepatu?.status === 'terlambat' ? 'Sepatu terlambat' : '', g.seragam?.status === 'terlambat' ? 'Seragam terlambat' : ''].filter(Boolean).join('; ');
      aoa.push([
        i + 1, g.k.nrp_amodi || g.k.nrp, g.k.nama, g.k.jabatan || '', g.k.tanggal_masuk ? fmt(g.k.tanggal_masuk) : '',
        g.sepatu ? (g.k.ukuran_sepatu || '') : '', g.sepatu ? QTY_SEPATU : '', g.sepatu ? fmt(g.sepatu.tgl) : '',
        g.seragam ? (g.k.ukuran_seragam || '') : '', g.seragam ? QTY_SERAGAM : '', g.seragam ? fmt(g.seragam.tgl) : '',
        g.j.lastSepatu ? fmt(g.j.lastSepatu) : '', g.j.lastSeragam ? fmt(g.j.lastSeragam) : '', tl,
      ]);
    });
    aoa.push(['', '', 'TOTAL', '', '', '', totSepatu, '', '', totSeragam, '', '', '', '']);
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws['!cols'] = [5, 11, 28, 18, 14, 11, 10, 14, 10, 10, 14, 20, 20, 22].map((wch) => ({ wch }));

    const rk = [['Rekap Ukuran untuk PO', judul], [], ['Jenis', 'Ukuran', 'Qty']];
    rekapSepatu.forEach(([u, n]) => rk.push(['Sepatu (pasang)', u, n]));
    rk.push(['Sepatu', 'TOTAL', totSepatu], []);
    rekapSeragam.forEach(([u, n]) => rk.push([`Seragam (${QTY_SERAGAM} per orang)`, u, n]));
    rk.push(['Seragam', 'TOTAL', totSeragam]);
    const ws2 = XLSX.utils.aoa_to_sheet(rk);
    ws2['!cols'] = [{ wch: 26 }, { wch: 14 }, { wch: 8 }];

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Orderan');
    XLSX.utils.book_append_sheet(wb, ws2, 'Rekap Ukuran');
    XLSX.writeFile(wb, `ORDERAN_APD_${NAMA_BULAN[bln - 1].toUpperCase()}_${thn}.xlsx`);
  };

  const unduhRekap = () => {
    const wb = XLSX.utils.book_new();
    const add = (nama, aoa, cols) => { const ws = XLSX.utils.aoa_to_sheet(aoa); ws['!cols'] = cols.map((wch) => ({ wch })); XLSX.utils.book_append_sheet(wb, ws, nama); };
    const pb = [[`Rekap pembagian APD tahun ${thnRekap}`], [], ['Bulan', 'Sepatu rutin (psg)', 'Sepatu ganti rusak (psg)', 'Total sepatu (psg)', 'Seragam (orang)', `Seragam (pcs, ${QTY_SERAGAM} per orang)`]];
    rekap.bulan.forEach((b) => pb.push([NAMA_BULAN[b.bln - 1], b.sepatuRutin, b.sepatuRusak, b.sepatuRutin + b.sepatuRusak, b.seragam, b.seragam * QTY_SERAGAM]));
    pb.push(['TOTAL', sumBulan((b) => b.sepatuRutin), sumBulan((b) => b.sepatuRusak), sumBulan((b) => b.sepatuRutin + b.sepatuRusak), sumBulan((b) => b.seragam), sumBulan((b) => b.seragam) * QTY_SERAGAM]);
    add('Pembagian per Bulan', pb, [14, 18, 24, 18, 16, 26]);
    const pj = [[`Rekap per jabatan (karyawan aktif), pembagian tahun ${thnRekap}`], [], ['Jabatan', 'Karyawan aktif', 'Sepatu tiap (bulan)', 'Sepatu rutin diterima', 'Sepatu ganti rusak', 'Seragam diterima', 'Sepatu terlambat', 'Seragam terlambat']];
    rekap.jabatan.forEach((e) => pj.push([e.jabatan, e.orang, e.iv, e.sepatu, e.rusak, e.seragam, e.telatSepatu, e.telatSeragam]));
    add('Per Jabatan', pj, [24, 14, 18, 20, 18, 16, 16, 16]);
    const py = [[`Proyeksi kebutuhan 12 bulan ke depan (dari jadwal per ${fmt(today)})`], [], ['Periode', 'Sepatu (orang)', 'Seragam (orang)', `Seragam (pcs, ${QTY_SERAGAM} per orang)`]];
    py.push(['Terlambat', rekap.proy.terlambat.sepatu, rekap.proy.terlambat.seragam, rekap.proy.terlambat.seragam * QTY_SERAGAM]);
    rekap.months.forEach((mo) => { const v = rekap.proy.bulan.get(mo); py.push([labelBulan(mo), v.sepatu, v.seragam, v.seragam * QTY_SERAGAM]); });
    add('Proyeksi 12 Bulan', py, [14, 16, 18, 28]);
    const uk = [['Ukuran karyawan aktif'], [], ['Jenis', 'Ukuran', 'Jumlah orang']];
    rekap.ukSepatu.forEach(([u, n]) => uk.push(['Sepatu', u, n]));
    rekap.ukSeragam.forEach(([u, n]) => uk.push(['Seragam', u, n]));
    add('Ukuran', uk, [12, 16, 14]);
    XLSX.writeFile(wb, `REKAP_APD_${thnRekap}.xlsx`);
  };

  /* ---------- tanda terima ---------- */
  const terimaRows = gabungan.filter((g) => selTerima.has(g.k.id));
  const totTerimaSepatu = terimaRows.filter((g) => g.sepatu).length * QTY_SEPATU;
  const totTerimaSeragam = terimaRows.filter((g) => g.seragam).length * QTY_SERAGAM;

  // Dialog "Catat sudah diterima": sepatu dan seragam punya tanggal terima masing-masing, dan bisa dicatat terpisah
  // karena barang sering tidak datang bersamaan. Item yang tidak dicentang tetap menunggu (belum tercatat).
  const [terimaModal, setTerimaModal] = useState(null); // { semua: { sepatu, seragam }, item: { 'id|jenis': { on, tgl } } }
  const bukaTerima = (rows = terimaRows) => {
    if (!rows.length) return;
    const item = {};
    rows.forEach((g) => {
      if (g.sepatu) item[`${g.k.id}|sepatu`] = { on: true, tgl: tglTerima };
      if (g.seragam) item[`${g.k.id}|seragam`] = { on: true, tgl: tglTerima };
    });
    setTerimaModal({ rows, semua: { sepatu: tglTerima, seragam: tglTerima }, item });
  };
  const ubahItem = (key, patch) => setTerimaModal((m) => ({ ...m, item: { ...m.item, [key]: { ...m.item[key], ...patch } } }));
  const ubahSemua = (jenis, patch) => setTerimaModal((m) => {
    const item = { ...m.item };
    Object.keys(item).forEach((k) => { if (k.endsWith(`|${jenis}`)) item[k] = { ...item[k], ...patch }; });
    return { ...m, item };
  });
  const simpanTerima = async () => {
    const items = [];
    terimaModal.rows.forEach((g) => ['sepatu', 'seragam'].forEach((jenis) => {
      const v = terimaModal.item[`${g.k.id}|${jenis}`];
      if (v && v.on) items.push({ karyawan_id: g.k.id, jenis, tanggal: v.tgl });
    }));
    if (!items.length) return notify('Centang minimal satu item yang sudah diterima.', 'error');
    const salah = items.find((it) => !it.tanggal || it.tanggal > today);
    if (salah) {
      const g = terimaModal.rows.find((x) => x.k.id === salah.karyawan_id);
      return notify(`Tanggal terima ${salah.jenis} ${g ? g.k.nama : ''} kosong atau melewati hari ini.`, 'error');
    }
    const ok = await catat(items);
    if (ok) {
      const nSepatu = items.filter((it) => it.jenis === 'sepatu').length;
      const diproses = terimaModal.rows.map((g) => g.k.id);
      setTerimaModal(null);
      setSelTerima((sel) => { const n = new Set(sel); diproses.forEach((id) => n.delete(id)); return n; });
      notify(`Tercatat ${nSepatu} sepatu dan ${items.length - nSepatu} seragam sesuai tanggal terima masing-masing. Item yang tidak dicentang tetap menunggu.`);
    }
  };

  /* ---------- impor Excel ---------- */
  const bacaFile = async (e) => {
    const file = e.target.files && e.target.files[0];
    try { e.target.value = ''; } catch (_) { /* abaikan */ }
    if (!file) return;
    setImpBusy(true); setImp(null);
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: 'array' });
      const nm = wb.SheetNames.find((n) => n.trim().toLowerCase() === 'resume') || wb.SheetNames[0];
      const aoa = XLSX.utils.sheet_to_json(wb.Sheets[nm], { header: 1, raw: true, defval: null });
      setImp({ nama: file.name, sheet: nm, ...parseRekap(aoa) });
    } catch (err) {
      notify('Gagal membaca file: ' + (err.message || err), 'error');
    }
    setImpBusy(false);
  };

  // bandingkan isi file dengan data di database (acuan: NRP Amodi)
  const impRingkas = useMemo(() => {
    if (!imp) return null;
    const byAmodi = new Map(karyawan.map((k) => [k.nrp_amodi, k]));
    const idToAmodi = new Map(karyawan.map((k) => [k.id, k.nrp_amodi]));
    const ada = new Set(pembagian.map((p) => `${idToAmodi.get(p.karyawan_id)}|${p.jenis}|${p.kategori}|${p.tanggal}`));
    const baru = imp.karyawan.filter((k) => !byAmodi.has(k.nrp_amodi));
    const sudah = imp.karyawan.filter((k) => byAmodi.has(k.nrp_amodi));
    const ubah = [];
    sudah.forEach((k) => {
      const e = byAmodi.get(k.nrp_amodi); const patch = {};
      ['jabatan', 'ukuran_sepatu', 'ukuran_seragam', 'tanggal_masuk'].forEach((f) => { if (k[f] && String(k[f]) !== String(e[f] ?? '')) patch[f] = k[f]; });
      if (Object.keys(patch).length) ubah.push({ k, patch });
    });
    const riwayatBaru = imp.riwayat.filter((r) => !ada.has(`${r.nrp_amodi}|${r.jenis}|${r.kategori}|${r.tanggal}`));
    return { baru, sudah, ubah, riwayatBaru, riwayatAda: imp.riwayat.length - riwayatBaru.length };
  }, [imp, karyawan, pembagian]);

  const jalankanImpor = async () => {
    if (!imp || !impRingkas) return;
    const { baru, ubah, riwayatBaru } = impRingkas;
    const ubahPakai = impUpdate ? ubah : [];
    if (!baru.length && !riwayatBaru.length && !ubahPakai.length) return notify('Tidak ada data baru untuk diimpor.', 'error');
    if (!window.confirm(`Impor ${baru.length} karyawan baru, ${riwayatBaru.length} catatan pembagian${ubahPakai.length ? ` dan perbarui ${ubahPakai.length} karyawan` : ''}?`)) return;
    setImpBusy(true);
    const label = `impor Excel${oleh ? ` (${oleh})` : ''}`;
    try {
      const idBy = new Map(karyawan.map((k) => [k.nrp_amodi, k.id]));
      for (let i = 0; i < baru.length; i += 200) {
        const chunk = baru.slice(i, i + 200).map((k) => ({
          nrp: k.nrp, nrp_amodi: k.nrp_amodi, nama: k.nama, jabatan: k.jabatan, tanggal_masuk: k.tanggal_masuk,
          ukuran_sepatu: k.ukuran_sepatu, ukuran_seragam: k.ukuran_seragam, keterangan: k.keterangan, created_by: label,
        }));
        const { data: ins, error } = await supabase.from('apd_karyawan').insert(chunk).select('id, nrp_amodi');
        if (error) throw error;
        ins.forEach((r) => idBy.set(r.nrp_amodi, r.id));
      }
      for (const { k, patch } of ubahPakai) {
        const { error } = await supabase.from('apd_karyawan').update({ ...patch, updated_at: new Date().toISOString(), updated_by: label }).eq('id', idBy.get(k.nrp_amodi));
        if (error) throw error;
      }
      const rows = riwayatBaru
        .map((r) => ({ karyawan_id: idBy.get(r.nrp_amodi), jenis: r.jenis, kategori: r.kategori, tanggal: r.tanggal, created_by: label }))
        .filter((r) => r.karyawan_id);
      for (let i = 0; i < rows.length; i += 500) {
        const { error } = await supabase.from('apd_pembagian').insert(rows.slice(i, i + 500));
        if (error) throw error;
      }
      notify(`Impor selesai: ${baru.length} karyawan baru, ${rows.length} catatan pembagian${ubahPakai.length ? `, ${ubahPakai.length} karyawan diperbarui` : ''}.`);
      setImp(null);
      await muat();
    } catch (err) {
      notify('Impor berhenti: ' + pgMessage(err) + '. Data yang sudah masuk aman, impor ulang file yang sama untuk melanjutkan.', 'error');
      await muat();
    }
    setImpBusy(false);
  };

  /* ---------- gaya ---------- */
  const card = darkMode ? 'bg-slate-800 border-slate-700' : 'bg-white border-slate-200';
  const thead = darkMode ? 'bg-slate-900 text-slate-400 border-slate-700' : 'bg-slate-50 text-slate-600 border-slate-200';
  const divide = darkMode ? 'divide-slate-700' : 'divide-slate-100';
  const kelasTbl = { card, thead, divide, dark: darkMode };
  const hoverRow = darkMode ? 'hover:bg-slate-700' : 'hover:bg-slate-50';
  const inputClass = 'w-full p-2 border border-slate-300 rounded-lg bg-white text-slate-900 placeholder-slate-400 focus:outline-none focus:border-blue-500 font-medium';
  const selectClass = 'px-3 py-2 text-sm border border-slate-300 rounded-lg bg-white text-slate-900 font-medium';
  const btnTab = (on) => `px-4 py-2 rounded-lg text-sm font-semibold border ${on ? 'bg-yellow-500 text-slate-950 border-transparent shadow' : darkMode ? 'border-slate-600 text-slate-300 hover:bg-slate-700' : 'border-slate-300 text-slate-600 hover:bg-slate-100'}`;

  const chip = (status) => {
    if (status === 'terlambat') return <span className="bg-rose-100 text-rose-800 px-2 py-0.5 rounded-full font-bold text-[10px] whitespace-nowrap">Terlambat</span>;
    if (status === 'bulan') return <span className="bg-amber-100 text-amber-800 px-2 py-0.5 rounded-full font-bold text-[10px] whitespace-nowrap">Bulan ini</span>;
    return <span className="bg-emerald-100 text-emerald-800 px-2 py-0.5 rounded-full font-bold text-[10px] whitespace-nowrap">Aman</span>;
  };
  const chipJadwal = (next) => chip(statusJadwal(next, ymNow));

  if (!canEdit) return null;

  const riwayatK = riwayatId ? data.find((d) => d.k.id === riwayatId) : null;
  const tahunOpsi = Array.from({ length: 7 }, (_, i) => now.getFullYear() - 2 + i);

  /* ---------- tabel orderan (satu jenis) ---------- */
  const TabelOrderan = ({ jenis, rows, sumber, tabel }) => {
    const nama = jenis === 'sepatu' ? 'Sepatu' : 'Seragam';
    const kol = kolomOrderan(jenis);
    const th = (c) => <Th key={c.key} c={c} tabel={tabel} rows={sumber} dark={darkMode} />;
    return (
      <div className={`rounded-xl shadow-sm border overflow-x-auto ${card}`}>
        <div className="p-3 border-b flex flex-wrap items-center justify-between gap-2">
          <h3 className="font-bold text-sm">{jenis === 'sepatu' ? '👟' : '👕'} Orderan {nama} <span className="text-slate-400 font-normal">({rows.length}{tabel.aktif ? ` dari ${sumber.length}` : ''} orang)</span></h3>
          <BtnReset tabel={tabel} />
        </div>
        <table className="w-full min-w-[760px] text-left text-sm">
          <thead className={`border-b font-semibold leading-tight ${thead}`}>
            <tr>
              {th(kol[0])}{th(kol[1])}{th(kol[2])}{th(kol[3])}<th className="p-2.5">JUMLAH</th>{th(kol[4])}{th(kol[5])}{th(kol[6])}
            </tr>
          </thead>
          <tbody className={`divide-y ${divide}`}>
            {rows.length === 0 ? (
              <tr><td colSpan="8" className="p-6 text-center text-slate-400">{sumber.length > 0 ? 'Tidak ada data yang cocok dengan filter.' : `Tidak ada jadwal ${nama.toLowerCase()} pada ${NAMA_BULAN[bln - 1]} ${thn}.`}</td></tr>
            ) : rows.map((r) => (
              <tr key={r.k.id} className={hoverRow}>
                <td className="p-2.5 text-center font-mono text-xs">{r.k.nrp_amodi || r.k.nrp}</td>
                <td className="p-2.5 font-semibold">{r.k.nama}</td>
                <td className="p-2.5">{r.k.jabatan || '-'}</td>
                <td className="p-2.5 text-center font-bold">{(jenis === 'sepatu' ? r.k.ukuran_sepatu : r.k.ukuran_seragam) || <span className="text-rose-500 text-xs">kosong</span>}</td>
                <td className="p-2.5 text-center">{jenis === 'sepatu' ? QTY_SEPATU : QTY_SERAGAM}</td>
                <td className="p-2.5 text-center whitespace-nowrap">{fmt(r.tgl)}{(jenis === 'sepatu' ? r.j.belumPernahSepatu : r.j.belumPernahSeragam) && <div className="text-[10px] text-slate-400">pembagian pertama</div>}</td>
                <td className="p-2.5 text-center whitespace-nowrap">{fmt(jenis === 'sepatu' ? r.j.lastSepatu : r.j.lastSeragam)}</td>
                <td className="p-2.5 text-center">{chip(r.status)}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    );
  };

  // satu sel (sepatu atau seragam) pada dialog tanda terima
  const celTerima = (g, jenis) => {
    const v = terimaModal.item[`${g.k.id}|${jenis}`];
    if (!v) return <td className="p-2 text-center text-slate-300">-</td>;
    const sah = !!v.tgl && v.tgl <= today;
    const last = jenis === 'sepatu' ? g.j.lastSepatu : g.j.lastSeragam;
    const iv = jenis === 'sepatu' ? g.j.intervalSepatu : INTERVAL_SERAGAM;
    const ukr = jenis === 'sepatu' ? g.k.ukuran_sepatu : g.k.ukuran_seragam;
    const key = `${g.k.id}|${jenis}`;
    return (
      <td className={`p-2 text-center align-top ${v.on ? '' : 'opacity-50'}`}>
        <label className="flex items-center justify-center gap-1.5 text-xs font-semibold mb-1">
          <input type="checkbox" checked={v.on} onChange={(e) => ubahItem(key, { on: e.target.checked })} />
          {jenis === 'sepatu' ? 'Sepatu' : 'Seragam'} {ukr || '?'}{jenis === 'seragam' ? ` x${QTY_SERAGAM}` : ''}
        </label>
        <input type="date" max={today} disabled={!v.on} value={v.tgl} onChange={(e) => ubahItem(key, { tgl: e.target.value })}
          className={`px-2 py-1 text-sm border rounded-lg bg-white text-slate-900 font-medium ${v.on && v.tgl && !sah ? 'border-rose-500' : 'border-slate-300'}`} />
        {v.on && v.tgl && !sah && <div className="text-[10px] text-rose-600">melewati hari ini</div>}
        {v.on && sah && last && v.tgl <= last && <div className="text-[10px] text-amber-600">tidak lebih baru dari terima terakhir</div>}
        <div className="text-[10px] text-slate-400">{v.on && sah ? `Berikutnya ${fmt(addMonths(v.tgl, iv))}` : 'belum dicatat'}</div>
      </td>
    );
  };

  /* ============================== RENDER ============================== */
  return (
    <>
      <style>{`
        #apd-print { display: none; }
        @media print {
          #apd-print { display: block !important; position: static !important; padding: 12mm; box-sizing: border-box; background: #fff !important; color: #000 !important; font-family: Arial, Helvetica, sans-serif; font-size: 10.5px; }
          #apd-print, #apd-print * { visibility: visible !important; }
          #apd-print table { width: 100%; border-collapse: collapse; }
          #apd-print th, #apd-print td { border: 0.3mm solid #000; padding: 1mm 1.5mm; }
          #apd-print th { background: #e5e7eb !important; text-align: center; }
          #apd-print thead { display: table-header-group; }
          #apd-print tr { page-break-inside: avoid; }
          #apd-print .apd-ttd { page-break-inside: avoid; }
          #apd-print .apd-tall td { height: 11mm; }
        }
      `}</style>

      <div className="space-y-4 no-print">
        {/* bilah atas */}
        <div className={`p-3 rounded-xl border shadow-sm flex flex-wrap gap-2 items-center justify-between ${card}`}>
          <div className="flex gap-2">
            <button className={btnTab(tab === 'daftar')} onClick={() => setTab('daftar')}>Daftar Karyawan</button>
            <button className={btnTab(tab === 'orderan')} onClick={() => setTab('orderan')}>Orderan Bulanan</button>
            <button className={btnTab(tab === 'terima')} onClick={() => setTab('terima')}>Tanda Terima</button>
            <button className={btnTab(tab === 'rekap')} onClick={() => setTab('rekap')}>Rekap</button>
            <button className={btnTab(tab === 'impor')} onClick={() => setTab('impor')}>Impor Excel</button>
          </div>
          <div className="flex gap-2 items-center">
            {loading && <span className="text-xs text-yellow-500 animate-pulse">Memuat...</span>}
            <button onClick={muat} disabled={loading} className="px-3 py-2 rounded-lg border border-slate-300 text-xs font-semibold flex items-center gap-1 disabled:opacity-40"><RefreshCw size={14} /> Muat ulang</button>
          </div>
        </div>

        {loadError && (
          <div className="p-4 rounded-xl border border-rose-300 bg-rose-50 text-rose-900 text-xs font-medium">
            Data APD gagal dimuat: {loadError}.{' '}
            {/(relation|does not exist|schema cache|42P01)/i.test(loadError) && <>Jalankan dulu <b>schema_v2_patch4.sql</b> di Supabase &gt; SQL Editor, lalu <b>apd_data_awal.sql</b>.</>}
          </div>
        )}

        {/* ===================== DAFTAR KARYAWAN ===================== */}
        {tab === 'daftar' && (
          <>
            <div className="grid grid-cols-3 gap-3">
              <div className={`p-4 rounded-xl border shadow-sm ${card}`}><p className="text-xs text-slate-400 font-semibold">KARYAWAN AKTIF</p><p className="text-2xl font-bold">{ringkas.aktif}</p></div>
              <div className={`p-4 rounded-xl border shadow-sm ${card}`}><p className="text-xs text-slate-400 font-semibold">JADWAL BULAN INI</p><p className="text-2xl font-bold text-amber-500">{ringkas.bulanIni}</p></div>
              <div className={`p-4 rounded-xl border shadow-sm ${card}`}><p className="text-xs text-slate-400 font-semibold">TERLAMBAT (BELUM DICATAT)</p><p className="text-2xl font-bold text-rose-500">{ringkas.terlambat}</p></div>
            </div>

            <div className={`p-4 rounded-xl shadow-sm border flex flex-wrap gap-3 items-center ${card}`}>
              <div className="relative flex-1 min-w-[200px]">
                <Search className="absolute left-3 top-2.5 text-slate-400" size={16} />
                <input type="text" placeholder="Cari nama, NRP Amodi, atau jabatan..." value={cari} onChange={(e) => { setCari(e.target.value); setPage(1); }}
                  className="w-full pl-9 pr-4 py-2 text-sm border border-slate-300 rounded-lg bg-white text-slate-900 focus:outline-none focus:border-blue-500 font-medium" />
              </div>
              <select value={filterAktif} onChange={(e) => { setFilterAktif(e.target.value); setPage(1); }} className={selectClass}>
                <option value="aktif">Aktif</option><option value="nonaktif">Resign</option><option value="semua">Semua</option>
              </select>
              <BtnReset tabel={tabDaftar} />
              <button onClick={() => setForm(emptyForm())} className="bg-blue-600 hover:bg-blue-700 text-white px-3 py-2 rounded-lg text-xs font-semibold flex items-center gap-1"><Plus size={14} /> Tambah Karyawan</button>
            </div>

            <div className={`rounded-xl shadow-sm border overflow-x-auto ${card}`}>
              <table className="w-full min-w-[1250px] text-left text-sm">
                <thead className={`border-b font-semibold leading-tight ${thead}`}>
                  <tr>
                    {kolomDaftar.map((c) => <Th key={c.key} c={c} tabel={tabDaftar} rows={daftarDasar} dark={darkMode} />)}
                    <th className="p-2.5">AKSI</th>
                  </tr>
                </thead>
                <tbody className={`divide-y ${divide}`}>
                  {daftarHal.length === 0 ? (
                    <tr><td colSpan="11" className="p-8 text-center text-slate-400">{loading ? 'Memuat...' : 'Tidak ada data ditemukan.'}</td></tr>
                  ) : daftarHal.map(({ k, j, riwayat }) => {
                    const sa = statusJadwal(j.nextSepatu, ymNow); const sg = statusJadwal(j.nextSeragam, ymNow);
                    const worst = !k.aktif ? 'nonaktif' : (sa === 'terlambat' || sg === 'terlambat') ? 'terlambat' : (sa === 'bulan' || sg === 'bulan') ? 'bulan' : 'aman';
                    return (
                      <tr key={k.id} className={`${hoverRow} ${k.aktif ? '' : 'opacity-60'}`}>
                        <td className="p-2.5 text-center font-mono text-xs">{k.nrp_amodi || k.nrp}{k.nrp_amodi && k.nrp && k.nrp !== k.nrp_amodi && <div className="text-[10px] text-slate-400">{k.nrp}</div>}</td>
                        <td className="p-2.5 font-semibold">{k.nama}</td>
                        <td className="p-2.5">{k.jabatan || '-'}</td>
                        <td className="p-2.5 text-center font-bold">{k.ukuran_sepatu || '-'}</td>
                        <td className="p-2.5 text-center font-bold">{k.ukuran_seragam || '-'}</td>
                        <td className="p-2.5 text-center whitespace-nowrap">{fmt(j.lastSepatu)}</td>
                        <td className="p-2.5 text-center whitespace-nowrap">{fmt(j.nextSepatu)}<div className="text-[10px] text-slate-400">tiap {j.intervalSepatu} bln</div></td>
                        <td className="p-2.5 text-center whitespace-nowrap">{fmt(j.lastSeragam)}</td>
                        <td className="p-2.5 text-center whitespace-nowrap">{fmt(j.nextSeragam)}<div className="text-[10px] text-slate-400">tiap {j.intervalSeragam} bln</div></td>
                        <td className="p-2.5 text-center">
                          {worst === 'nonaktif' ? <><span className="bg-slate-200 text-slate-600 px-2 py-0.5 rounded-full font-bold text-[10px]">Resign</span>{k.tanggal_resign && <div className="text-[10px] text-slate-400 mt-0.5">{fmt(k.tanggal_resign)}</div>}</> : chip(worst)}
                        </td>
                        <td className="p-2.5 text-center whitespace-nowrap">
                          <button title="Ubah data" onClick={() => setForm(formDari(k, riwayat))} className="p-1.5 text-blue-500 hover:bg-blue-50 rounded"><Pencil size={14} /></button>
                          <button title="Riwayat pembagian" onClick={() => { setRiwayatId(k.id); setFormRiwayat({ jenis: 'sepatu', kategori: 'rutin', tanggal: todayISO(), keterangan: '' }); }} className="p-1.5 text-slate-500 hover:bg-slate-100 rounded"><History size={14} /></button>
                          {k.aktif && <button title="Catat penggantian sepatu rusak" onClick={() => setRusak({ k, tanggal: todayISO(), keterangan: '' })} className="px-1.5 py-1 text-[10px] font-bold text-amber-700 bg-amber-100 hover:bg-amber-200 rounded mx-0.5">Ganti rusak</button>}
                          {k.aktif
                            ? <button title="Tandai resign: tidak muncul lagi di pemantauan" onClick={() => setResign({ k, tanggal: todayISO() })} className="px-1.5 py-1 text-[10px] font-bold rounded text-rose-700 bg-rose-100 hover:bg-rose-200">Resign</button>
                            : <button title="Aktifkan kembali" onClick={() => aktifkanKembali(k)} className="px-1.5 py-1 text-[10px] font-bold rounded text-emerald-700 bg-emerald-100 hover:bg-emerald-200">Aktifkan</button>}
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
              <div className={`p-4 border-t flex justify-between items-center text-xs ${darkMode ? 'border-slate-700 text-slate-300' : 'border-slate-200 text-slate-700'}`}>
                <span className="font-semibold">Halaman <span className="text-yellow-500 font-bold">{pageNow}</span> dari {totalPage} (Total: {daftar.length} Data)</span>
                <div className="flex gap-2">
                  <button onClick={() => setPage(Math.max(pageNow - 1, 1))} disabled={pageNow === 1} className="px-3 py-1.5 rounded-lg bg-emerald-600 text-white font-bold disabled:opacity-40 shadow">◀ Sebelumnya</button>
                  <button onClick={() => setPage(Math.min(pageNow + 1, totalPage))} disabled={pageNow >= totalPage} className="px-3 py-1.5 rounded-lg bg-emerald-600 text-white font-bold disabled:opacity-40 shadow">Berikutnya ▶</button>
                </div>
              </div>
            </div>
            <p className="text-[11px] text-slate-400">
              Aturan: sepatu tiap 6 bulan untuk {JABATAN_SEPATU_6_BULAN.join(', ')}; jabatan lain tiap 12 bulan. Seragam tiap {INTERVAL_SERAGAM} bulan, {QTY_SERAGAM} potong per orang.
              Jadwal berikutnya dihitung dari tanggal terima terakhir; sepatu rusak yang diganti menggeser hitungan sepatu mulai dari tanggal penggantian.
            </p>
          </>
        )}

        {/* ===================== TANDA TERIMA ===================== */}
        {tab === 'terima' && (
          <>
            <div className={`p-4 rounded-xl shadow-sm border flex flex-wrap gap-3 items-center ${card}`}>
              <select value={bln} onChange={(e) => setBln(Number(e.target.value))} className={selectClass}>
                {NAMA_BULAN.map((n, i) => <option key={i} value={i + 1}>{n}</option>)}
              </select>
              <select value={thn} onChange={(e) => setThn(Number(e.target.value))} className={selectClass}>
                {tahunOpsi.map((y) => <option key={y} value={y}>{y}</option>)}
              </select>
              <label className="text-xs font-semibold flex items-center gap-1.5">
                <input type="checkbox" checked={sertakanTerlambat} onChange={(e) => setSertakanTerlambat(e.target.checked)} />
                Sertakan yang terlambat
              </label>
              <label className="text-xs font-semibold ml-2">Tanggal pembagian:</label>
              <input type="date" value={tglTerima} onChange={(e) => setTglTerima(e.target.value)} className="px-3 py-1.5 text-sm border border-slate-300 rounded-lg bg-white text-slate-900 font-medium" />
              <BtnReset tabel={tabTerima} />
              <div className="flex-1" />
              <button onClick={() => window.print()} disabled={!terimaRows.length} className="bg-slate-700 disabled:opacity-40 text-white px-3 py-2 rounded-lg text-xs font-semibold flex items-center gap-1"><Printer size={14} /> Cetak Tanda Terima ({terimaRows.length})</button>
              <button onClick={() => bukaTerima()} disabled={!terimaRows.length || saving} className="bg-emerald-600 disabled:opacity-40 text-white px-3 py-2 rounded-lg text-xs font-semibold flex items-center gap-1"><Check size={14} /> Catat sudah diterima ({terimaRows.length})</button>
            </div>
            <p className="text-[11px] text-slate-400">
              Alur: centang karyawan, cetak lembar tanda terima untuk ditandatangani saat pembagian, lalu setelah semua tanda tangan terkumpul klik "Catat sudah diterima".
              Tanggal pembagian di atas menjadi dasar jadwal berikutnya (sepatu 6 atau 12 bulan sesuai jabatan, seragam {INTERVAL_SERAGAM} bulan).
            </p>

            <div className={`rounded-xl shadow-sm border overflow-x-auto ${card}`}>
              <table className="w-full min-w-[820px] text-left text-sm">
                <thead className={`border-b font-semibold leading-tight ${thead}`}>
                  <tr>
                    <th className="p-2.5 w-8"><input type="checkbox" title="Centang semua baris yang tampil"
                      checked={gabunganTampil.length > 0 && gabunganTampil.every((g) => selTerima.has(g.k.id))}
                      onChange={() => { const semuaTerpilih = gabunganTampil.every((g) => selTerima.has(g.k.id)); const n = new Set(selTerima); gabunganTampil.forEach((g) => { if (semuaTerpilih) n.delete(g.k.id); else n.add(g.k.id); }); setSelTerima(n); }} /></th>
                    {kolomTerima.map((c) => <Th key={c.key} c={c} tabel={tabTerima} rows={gabungan} dark={darkMode} />)}
                    <th className="p-2.5">AKSI</th>
                  </tr>
                </thead>
                <tbody className={`divide-y ${divide}`}>
                  {gabungan.length === 0 ? (
                    <tr><td colSpan="9" className="p-8 text-center text-slate-400">{loading ? 'Memuat...' : gabungan.length > 0 ? 'Tidak ada data yang cocok dengan filter.' : `Tidak ada jadwal pada ${NAMA_BULAN[bln - 1]} ${thn}.`}</td></tr>
                  ) : gabunganTampil.map((g) => {
                    const telat = g.sepatu?.status === 'terlambat' || g.seragam?.status === 'terlambat';
                    return (
                      <tr key={g.k.id} className={hoverRow}>
                        <td className="p-2.5 text-center"><input type="checkbox" checked={selTerima.has(g.k.id)} onChange={() => { const n = new Set(selTerima); if (n.has(g.k.id)) n.delete(g.k.id); else n.add(g.k.id); setSelTerima(n); }} /></td>
                        <td className="p-2.5 text-center font-mono text-xs">{g.k.nrp_amodi || g.k.nrp}</td>
                        <td className="p-2.5 font-semibold">{g.k.nama}</td>
                        <td className="p-2.5">{g.k.jabatan || '-'}</td>
                        <td className="p-2.5 text-center">{g.sepatu ? <><b>{g.k.ukuran_sepatu || '?'}</b> <span className="text-slate-400 text-xs">x{QTY_SEPATU}</span></> : '-'}</td>
                        <td className="p-2.5 text-center">{g.seragam ? <><b>{g.k.ukuran_seragam || '?'}</b> <span className="text-slate-400 text-xs">x{QTY_SERAGAM}</span></> : '-'}</td>
                        <td className="p-2.5 text-center text-xs whitespace-nowrap">{g.sepatu && <div>Sepatu {fmt(g.sepatu.tgl)}</div>}{g.seragam && <div>Seragam {fmt(g.seragam.tgl)}</div>}</td>
                        <td className="p-2.5 text-center">{chip(telat ? 'terlambat' : 'bulan')}</td>
                        <td className="p-2.5 text-center">
                          <button onClick={() => bukaTerima([g])} disabled={saving} title={`Catat ${g.k.nama} sudah menerima (tanggal bisa diatur)`}
                            className="bg-emerald-600 hover:bg-emerald-700 disabled:opacity-40 text-white px-3 py-1.5 rounded-lg text-xs font-semibold inline-flex items-center gap-1"><Check size={14} /> Terima</button>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
            {terimaRows.length > 0 && <p className="text-xs font-semibold">Terpilih: {terimaRows.length} orang, sepatu {totTerimaSepatu} psg, seragam {totTerimaSeragam} pcs.</p>}
          </>
        )}

        {/* ===================== IMPOR EXCEL ===================== */}
        {tab === 'impor' && (
          <>
            <div className={`p-4 rounded-xl shadow-sm border space-y-3 ${card}`}>
              <h3 className="font-bold text-sm">Impor data penerimaan sepatu dan seragam dari Excel</h3>
              <p className="text-xs text-slate-400">
                Gunakan file berformat "MONITORING SEPATU DAN SERAGAM" (sheet RESUME): kolom NRP Amodi, NRP, Nama, Jabatan, Tanggal Masuk, Ukuran Sepatu, Ukuran Seragam,
                Tanda Terima Terakhir, Jadwal Pembagian beserta Progress, dan Penggantian Sepatu Rusak. Jadwal yang berstatus SUDAH dicatat sebagai sudah diterima pada tanggal jadwal itu.
                Acuan data adalah NRP Amodi. Impor tidak menghapus apa pun, tidak mengubah status resign, dan aman diulang karena data yang sudah ada dilewati.
              </p>
              <label className="inline-flex items-center gap-2 bg-blue-600 hover:bg-blue-700 text-white px-4 py-2 rounded-lg text-xs font-semibold cursor-pointer">
                <Upload size={14} /> {impBusy ? 'Memproses...' : 'Pilih file Excel (.xlsx)'}
                <input type="file" accept=".xlsx,.xls" className="hidden" disabled={impBusy} onChange={bacaFile} />
              </label>
              {imp && <span className="text-xs ml-2 text-slate-400">{imp.nama} (sheet {imp.sheet})</span>}
            </div>

            {imp && impRingkas && (
              <>
                <div className="grid grid-cols-2 lg:grid-cols-5 gap-3">
                  <div className={`p-4 rounded-xl border shadow-sm ${card}`}><p className="text-xs text-slate-400 font-semibold">KARYAWAN DI FILE</p><p className="text-2xl font-bold">{imp.karyawan.length}</p></div>
                  <div className={`p-4 rounded-xl border shadow-sm ${card}`}><p className="text-xs text-slate-400 font-semibold">KARYAWAN BARU</p><p className="text-2xl font-bold text-emerald-600">{impRingkas.baru.length}</p></div>
                  <div className={`p-4 rounded-xl border shadow-sm ${card}`}><p className="text-xs text-slate-400 font-semibold">SUDAH ADA</p><p className="text-2xl font-bold">{impRingkas.sudah.length}</p></div>
                  <div className={`p-4 rounded-xl border shadow-sm ${card}`}><p className="text-xs text-slate-400 font-semibold">CATATAN BARU</p><p className="text-2xl font-bold text-emerald-600">{impRingkas.riwayatBaru.length}</p></div>
                  <div className={`p-4 rounded-xl border shadow-sm ${card}`}><p className="text-xs text-slate-400 font-semibold">CATATAN DILEWATI</p><p className="text-2xl font-bold text-slate-400">{impRingkas.riwayatAda}</p></div>
                </div>

                {impRingkas.ubah.length > 0 && (
                  <label className={`p-3 rounded-xl border shadow-sm text-xs font-semibold flex items-start gap-2 ${card}`}>
                    <input type="checkbox" className="mt-0.5" checked={impUpdate} onChange={(e) => setImpUpdate(e.target.checked)} />
                    <span>Perbarui juga {impRingkas.ubah.length} karyawan yang sudah ada karena jabatan, ukuran, atau tanggal masuknya berbeda dengan file.<span className="block font-normal text-slate-400">Contoh: {impRingkas.ubah.slice(0, 3).map(({ k, patch }) => `${k.nama} (${Object.keys(patch).join(', ')})`).join('; ')}</span></span>
                  </label>
                )}

                {imp.masalah.length > 0 && (
                  <div className="p-3 rounded-xl border border-amber-300 bg-amber-50 text-amber-900 text-xs space-y-0.5">
                    <p className="font-bold">{imp.masalah.length} catatan dari pembacaan file (baris bermasalah dilewati):</p>
                    {imp.masalah.slice(0, 15).map((m, i) => <p key={i}>Baris {m.baris}: {m.pesan}</p>)}
                    {imp.masalah.length > 15 && <p>... dan {imp.masalah.length - 15} lainnya.</p>}
                  </div>
                )}

                {impRingkas.baru.length > 0 && (
                  <div className={`rounded-xl shadow-sm border overflow-x-auto ${card}`}>
                    <div className="p-3 border-b font-bold text-sm">Karyawan baru yang akan ditambahkan {impRingkas.baru.length > 8 && <span className="text-slate-400 font-normal">(8 pertama)</span>}</div>
                    <table className="w-full min-w-[640px] text-sm">
                      <thead className={`border-b font-semibold ${thead}`}><tr><th className="p-2 text-left">NRP AMODI</th><th className="p-2 text-left">NAMA</th><th className="p-2 text-left">JABATAN</th><th className="p-2">SEPATU</th><th className="p-2">SERAGAM</th></tr></thead>
                      <tbody className={`divide-y ${divide}`}>
                        {impRingkas.baru.slice(0, 8).map((k) => (
                          <tr key={k.nrp_amodi}><td className="p-2 font-mono text-xs">{k.nrp_amodi}</td><td className="p-2 font-semibold">{k.nama}</td><td className="p-2">{k.jabatan || '-'}</td><td className="p-2 text-center">{k.ukuran_sepatu || '-'}</td><td className="p-2 text-center">{k.ukuran_seragam || '-'}</td></tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                )}

                <div className="flex justify-end gap-2">
                  <button onClick={() => setImp(null)} disabled={impBusy} className="px-4 py-2 rounded-lg border border-slate-300 text-xs font-semibold">Batal</button>
                  <button onClick={jalankanImpor} disabled={impBusy || (!impRingkas.baru.length && !impRingkas.riwayatBaru.length && !(impUpdate && impRingkas.ubah.length))}
                    className="px-4 py-2 rounded-lg bg-emerald-600 text-white text-xs font-semibold disabled:opacity-40">{impBusy ? 'Mengimpor...' : 'Impor Sekarang'}</button>
                </div>
              </>
            )}
          </>
        )}

        {/* ===================== REKAP ===================== */}
        {tab === 'rekap' && (
          <>
            <div className={`p-4 rounded-xl shadow-sm border flex flex-wrap gap-3 items-center ${card}`}>
              <label className="text-xs font-semibold">Tahun pembagian:</label>
              <select value={thnRekap} onChange={(e) => setThnRekap(Number(e.target.value))} className={selectClass}>
                {Array.from({ length: 5 }, (_, i) => now.getFullYear() - 3 + i).map((y) => <option key={y} value={y}>{y}</option>)}
              </select>
              <div className="flex-1" />
              <button onClick={unduhRekap} className="bg-emerald-600 text-white px-3 py-2 rounded-lg text-xs font-semibold flex items-center gap-1"><Download size={14} /> Unduh Excel</button>
            </div>

            <div className="grid grid-cols-2 lg:grid-cols-4 gap-3">
              <div className={`p-4 rounded-xl border shadow-sm ${card}`}><p className="text-xs text-slate-400 font-semibold">SEPATU DITERIMA {thnRekap}</p><p className="text-2xl font-bold">{sumBulan((b) => b.sepatuRutin + b.sepatuRusak)} <span className="text-xs font-normal text-slate-400">psg</span></p></div>
              <div className={`p-4 rounded-xl border shadow-sm ${card}`}><p className="text-xs text-slate-400 font-semibold">SEPATU GANTI RUSAK</p><p className="text-2xl font-bold text-amber-500">{sumBulan((b) => b.sepatuRusak)} <span className="text-xs font-normal text-slate-400">psg</span></p></div>
              <div className={`p-4 rounded-xl border shadow-sm ${card}`}><p className="text-xs text-slate-400 font-semibold">SERAGAM DITERIMA {thnRekap}</p><p className="text-2xl font-bold">{sumBulan((b) => b.seragam)} <span className="text-xs font-normal text-slate-400">orang ({sumBulan((b) => b.seragam) * QTY_SERAGAM} pcs)</span></p></div>
              <div className={`p-4 rounded-xl border shadow-sm ${card}`}><p className="text-xs text-slate-400 font-semibold">TERLAMBAT SAAT INI</p><p className="text-2xl font-bold text-rose-500">{ringkas.terlambat} <span className="text-xs font-normal text-slate-400">orang</span></p></div>
            </div>

            <div className={`rounded-xl shadow-sm border overflow-x-auto ${card}`}>
              <div className="p-3 border-b font-bold text-sm flex flex-wrap items-center justify-between gap-2">
                <span>Pembagian per bulan, tahun {thnRekap} <span className="text-slate-400 font-normal">(dari catatan tanggal terima)</span></span>
                <BtnReset tabel={tabBulan} />
              </div>
              <table className="w-full min-w-[640px] text-sm">
                <thead className={`border-b font-semibold leading-tight ${thead}`}><tr>{kolomBulan.map((c) => <Th key={c.key} c={c} tabel={tabBulan} rows={rekap.bulan} dark={darkMode} />)}</tr></thead>
                <tbody className={`divide-y ${divide}`}>
                  {bulanTampil.length === 0 ? <tr><td colSpan="6" className="p-4 text-center text-slate-400">Tidak ada data yang cocok dengan filter.</td></tr> : bulanTampil.map((b) => (
                    <tr key={b.bln} className={hoverRow}>
                      <td className="p-2.5 font-semibold">{NAMA_BULAN[b.bln - 1]}</td>
                      <td className="p-2.5 text-center">{b.sepatuRutin || '-'}</td><td className="p-2.5 text-center">{b.sepatuRusak || '-'}</td>
                      <td className="p-2.5 text-center font-bold">{b.sepatuRutin + b.sepatuRusak || '-'}</td>
                      <td className="p-2.5 text-center">{b.seragam || '-'}</td><td className="p-2.5 text-center font-bold">{b.seragam * QTY_SERAGAM || '-'}</td>
                    </tr>
                  ))}
                  <tr className="font-bold">
                    <td className="p-2.5">TOTAL{tabBulan.aktif ? ' (terfilter)' : ''}</td>
                    <td className="p-2.5 text-center">{bulanTampil.reduce((a, b) => a + b.sepatuRutin, 0)}</td><td className="p-2.5 text-center">{bulanTampil.reduce((a, b) => a + b.sepatuRusak, 0)}</td>
                    <td className="p-2.5 text-center">{bulanTampil.reduce((a, b) => a + b.sepatuRutin + b.sepatuRusak, 0)}</td><td className="p-2.5 text-center">{bulanTampil.reduce((a, b) => a + b.seragam, 0)}</td>
                    <td className="p-2.5 text-center">{bulanTampil.reduce((a, b) => a + b.seragam, 0) * QTY_SERAGAM}</td>
                  </tr>
                </tbody>
              </table>
            </div>

            <div className="grid lg:grid-cols-2 gap-4 items-start">
              <div className={`rounded-xl shadow-sm border overflow-x-auto ${card}`}>
                <div className="p-3 border-b font-bold text-sm flex flex-wrap items-center justify-between gap-2"><span>Proyeksi kebutuhan 12 bulan ke depan</span><BtnReset tabel={tabProy} /></div>
                <table className="w-full min-w-[420px] text-sm">
                  <thead className={`border-b font-semibold leading-tight ${thead}`}><tr>{kolomProy.map((c) => <Th key={c.key} c={c} tabel={tabProy} rows={proyBaris} dark={darkMode} />)}</tr></thead>
                  <tbody className={`divide-y ${divide}`}>
                    {proyTampil.length === 0 ? <tr><td colSpan="4" className="p-4 text-center text-slate-400">Tidak ada data yang cocok dengan filter.</td></tr> : proyTampil.map((r) => (
                      <tr key={r.label} className={r.telat ? 'bg-rose-50 text-rose-800 font-semibold' : hoverRow}>
                        <td className={`p-2.5 ${r.telat ? '' : 'font-semibold'}`}>{r.label}</td>
                        <td className="p-2.5 text-center">{r.sepatu || '-'}</td><td className="p-2.5 text-center">{r.seragam || '-'}</td><td className="p-2.5 text-center">{r.seragam * QTY_SERAGAM || '-'}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
                <p className="p-3 text-[11px] text-slate-400 border-t">Dihitung dari jadwal saat ini; siklus berikutnya = jadwal + interval. Bergeser bila tanggal terima sebenarnya berbeda.</p>
              </div>
              <div className="space-y-4">
                <TabelRekapUkuran judul="Ukuran sepatu karyawan aktif" rows={rekap.ukSepatu} satuan="orang" jenis="sepatu" kelas={kelasTbl} />
                <TabelRekapUkuran judul="Ukuran seragam karyawan aktif" rows={rekap.ukSeragam} satuan="orang" jenis="seragam" kelas={kelasTbl} />
              </div>
            </div>

            <div className={`rounded-xl shadow-sm border overflow-x-auto ${card}`}>
              <div className="p-3 border-b font-bold text-sm flex flex-wrap items-center justify-between gap-2"><span>Rekap per jabatan (karyawan aktif), pembagian tahun {thnRekap}</span><BtnReset tabel={tabJab} /></div>
              <table className="w-full min-w-[820px] text-sm">
                <thead className={`border-b font-semibold leading-tight ${thead}`}><tr>{kolomJab.map((c) => <Th key={c.key} c={c} tabel={tabJab} rows={rekap.jabatan} dark={darkMode} />)}</tr></thead>
                <tbody className={`divide-y ${divide}`}>
                  {jabTampil.length === 0 ? <tr><td colSpan="8" className="p-4 text-center text-slate-400">Tidak ada data yang cocok dengan filter.</td></tr> : jabTampil.map((e) => (
                    <tr key={e.jabatan} className={hoverRow}>
                      <td className="p-2.5 font-semibold">{e.jabatan}</td><td className="p-2.5 text-center font-bold">{e.orang}</td><td className="p-2.5 text-center">{e.iv} bln</td>
                      <td className="p-2.5 text-center">{e.sepatu || '-'}</td><td className="p-2.5 text-center">{e.rusak || '-'}</td><td className="p-2.5 text-center">{e.seragam || '-'}</td>
                      <td className={`p-2.5 text-center ${e.telatSepatu ? 'text-rose-600 font-bold' : ''}`}>{e.telatSepatu || '-'}</td><td className={`p-2.5 text-center ${e.telatSeragam ? 'text-rose-600 font-bold' : ''}`}>{e.telatSeragam || '-'}</td>
                    </tr>
                  ))}
                  <tr className="font-bold">
                    <td className="p-2.5">TOTAL{tabJab.aktif ? ' (terfilter)' : ''}</td><td className="p-2.5 text-center">{jabTampil.reduce((a, e) => a + e.orang, 0)}</td><td className="p-2.5"></td>
                    <td className="p-2.5 text-center">{jabTampil.reduce((a, e) => a + e.sepatu, 0)}</td><td className="p-2.5 text-center">{jabTampil.reduce((a, e) => a + e.rusak, 0)}</td><td className="p-2.5 text-center">{jabTampil.reduce((a, e) => a + e.seragam, 0)}</td>
                    <td className="p-2.5 text-center">{jabTampil.reduce((a, e) => a + e.telatSepatu, 0)}</td><td className="p-2.5 text-center">{jabTampil.reduce((a, e) => a + e.telatSeragam, 0)}</td>
                  </tr>
                </tbody>
              </table>
            </div>
          </>
        )}

        {/* ===================== ORDERAN BULANAN ===================== */}
        {tab === 'orderan' && (
          <>
            <div className={`p-4 rounded-xl shadow-sm border flex flex-wrap gap-3 items-center ${card}`}>
              <select value={bln} onChange={(e) => setBln(Number(e.target.value))} className={selectClass}>
                {NAMA_BULAN.map((n, i) => <option key={i} value={i + 1}>{n}</option>)}
              </select>
              <select value={thn} onChange={(e) => setThn(Number(e.target.value))} className={selectClass}>
                {tahunOpsi.map((y) => <option key={y} value={y}>{y}</option>)}
              </select>
              <label className="text-xs font-semibold flex items-center gap-1.5">
                <input type="checkbox" checked={sertakanTerlambat} onChange={(e) => setSertakanTerlambat(e.target.checked)} />
                Sertakan yang terlambat dari bulan sebelumnya
              </label>
              <div className="flex-1" />
              <button onClick={unduhExcel} disabled={!gabungan.length} className="bg-emerald-600 disabled:opacity-40 text-white px-3 py-2 rounded-lg text-xs font-semibold flex items-center gap-1"><Download size={14} /> Unduh Excel</button>
              <button onClick={() => window.print()} disabled={!gabungan.length} className="bg-slate-700 disabled:opacity-40 text-white px-3 py-2 rounded-lg text-xs font-semibold flex items-center gap-1"><Printer size={14} /> Cetak</button>
            </div>

            <div className={`p-3 rounded-xl border shadow-sm flex flex-wrap gap-3 items-center ${card}`}>
              <span className="text-xs text-slate-400">Halaman ini hanya untuk melihat orderan, rekap ukuran, unduh Excel, dan cetak. Pencatatan penerimaan sepatu dan seragam dilakukan di tab Tanda Terima.</span>
              <button onClick={() => setTab('terima')} className="px-3 py-1.5 rounded-lg bg-emerald-600 text-white text-xs font-semibold">Buka Tanda Terima</button>
            </div>

            <div className="grid lg:grid-cols-4 gap-4 items-start">
              <div className="lg:col-span-3 space-y-4">
                {TabelOrderan({ jenis: 'sepatu', rows: sepatuTampil, sumber: orderan.sepatu, tabel: tabSepatu })}
                {TabelOrderan({ jenis: 'seragam', rows: seragamTampil, sumber: orderan.seragam, tabel: tabSeragam })}
              </div>
              <div className="space-y-4">
                <TabelRekapUkuran judul="Rekap Sepatu untuk PO" rows={rekapSepatu} satuan="psg" jenis="sepatu" kelas={kelasTbl} />
                <TabelRekapUkuran judul={`Rekap Seragam untuk PO (${QTY_SERAGAM} per orang)`} rows={rekapSeragam} satuan="pcs" jenis="seragam" kelas={kelasTbl} />
              </div>
            </div>
            <p className="text-[11px] text-slate-400">Daftar dihitung dari data terkini: bulan yang sudah lewat hanya menampilkan pembagian yang belum dicatat.</p>
          </>
        )}
      </div>

      {/* ===================== AREA CETAK ===================== */}
      {tab === 'terima' && (
        <div id="apd-print">
          <div style={{ textAlign: 'center', marginBottom: '4mm' }}>
            <div style={{ fontWeight: 700, fontSize: '14px' }}>PT. RYAN EKA PRATAMA</div>
            <div style={{ fontWeight: 700 }}>BONTANG COAL MINE PROJECT</div>
            <div style={{ fontWeight: 700, marginTop: '2mm', textDecoration: 'underline' }}>TANDA TERIMA SEPATU DAN SERAGAM</div>
            <div>Tanggal pembagian: {fmtPanjang(tglTerima)}</div>
          </div>
          <table>
            <thead>
              <tr><th rowSpan="2">No</th><th rowSpan="2">NRP Amodi</th><th rowSpan="2">Nama</th><th rowSpan="2">Jabatan</th><th colSpan="2">Sepatu</th><th colSpan="2">Baju Seragam</th><th rowSpan="2" style={{ width: '32mm' }}>Tanda Tangan<br />Penerima</th></tr>
              <tr><th>Size</th><th>Qty</th><th>Size</th><th>Qty</th></tr>
            </thead>
            <tbody>
              {terimaRows.map((g, i) => (
                <tr key={g.k.id} className="apd-tall">
                  <td style={{ textAlign: 'center' }}>{i + 1}</td><td>{g.k.nrp_amodi || g.k.nrp}</td><td>{g.k.nama}</td><td>{g.k.jabatan || ''}</td>
                  <td style={{ textAlign: 'center' }}>{g.sepatu ? g.k.ukuran_sepatu || '' : ''}</td><td style={{ textAlign: 'center' }}>{g.sepatu ? QTY_SEPATU : ''}</td>
                  <td style={{ textAlign: 'center' }}>{g.seragam ? g.k.ukuran_seragam || '' : ''}</td><td style={{ textAlign: 'center' }}>{g.seragam ? QTY_SERAGAM : ''}</td>
                  <td></td>
                </tr>
              ))}
              <tr style={{ fontWeight: 700 }}>
                <td colSpan="4" style={{ textAlign: 'center' }}>T O T A L</td>
                <td></td><td style={{ textAlign: 'center' }}>{totTerimaSepatu}</td><td></td><td style={{ textAlign: 'center' }}>{totTerimaSeragam}</td><td></td>
              </tr>
            </tbody>
          </table>
          <div className="apd-ttd" style={{ marginTop: '5mm' }}>
            <div style={{ fontSize: '10px' }}>Dengan tanda tangan di atas, penerima menyatakan telah menerima APD tersebut dalam keadaan baik.</div>
            <div style={{ textAlign: 'right', marginTop: '5mm' }}>Bontang, {fmtPanjang(tglTerima)}</div>
            <div style={{ display: 'flex', justifyContent: 'space-around', marginTop: '2mm', textAlign: 'center' }}>
              {TTD_TERIMA.map((t) => (
                <div key={t.nama} style={{ width: '40%' }}>
                  <div>{t.judul}</div><div style={{ height: '18mm' }} />
                  <div style={{ fontWeight: 700, textDecoration: 'underline' }}>{t.nama}</div><div>{t.jabatan}</div>
                </div>
              ))}
            </div>
          </div>
        </div>
      )}
      {tab === 'orderan' && (
        <div id="apd-print">
          <div style={{ textAlign: 'center', marginBottom: '4mm' }}>
            <div style={{ fontWeight: 700, fontSize: '14px' }}>PT. RYAN EKA PRATAMA</div>
            <div style={{ fontWeight: 700 }}>BONTANG COAL MINE PROJECT</div>
            <div style={{ fontWeight: 700, marginTop: '2mm', textDecoration: 'underline' }}>RINCIAN ORDERAN APD KARYAWAN</div>
            <div>Periode {NAMA_BULAN[bln - 1]} {thn}</div>
          </div>
          <table>
            <thead>
              <tr><th rowSpan="2">No</th><th rowSpan="2">NRP Amodi</th><th rowSpan="2">Nama</th><th rowSpan="2">Jabatan</th><th rowSpan="2">Tanggal Masuk</th><th colSpan="2">Sepatu</th><th colSpan="2">Baju Seragam</th><th rowSpan="2">Terima Terakhir<br />Sepatu</th></tr>
              <tr><th>Size</th><th>Qty</th><th>Size</th><th>Qty</th></tr>
            </thead>
            <tbody>
              {gabungan.map((g, i) => (
                <tr key={g.k.id}>
                  <td style={{ textAlign: 'center' }}>{i + 1}</td><td>{g.k.nrp_amodi || g.k.nrp}</td><td>{g.k.nama}</td><td>{g.k.jabatan || ''}</td>
                  <td style={{ textAlign: 'center' }}>{g.k.tanggal_masuk ? fmt(g.k.tanggal_masuk) : ''}</td>
                  <td style={{ textAlign: 'center' }}>{g.sepatu ? g.k.ukuran_sepatu || '' : ''}</td><td style={{ textAlign: 'center' }}>{g.sepatu ? QTY_SEPATU : ''}</td>
                  <td style={{ textAlign: 'center' }}>{g.seragam ? g.k.ukuran_seragam || '' : ''}</td><td style={{ textAlign: 'center' }}>{g.seragam ? QTY_SERAGAM : ''}</td>
                  <td style={{ textAlign: 'center' }}>{g.j.lastSepatu ? fmt(g.j.lastSepatu) : ''}</td>
                </tr>
              ))}
              <tr style={{ fontWeight: 700 }}>
                <td colSpan="5" style={{ textAlign: 'center' }}>T O T A L</td>
                <td></td><td style={{ textAlign: 'center' }}>{totSepatu}</td><td></td><td style={{ textAlign: 'center' }}>{totSeragam}</td><td></td>
              </tr>
            </tbody>
          </table>

          <div className="apd-ttd" style={{ marginTop: '5mm' }}>
            <div style={{ display: 'flex', gap: '8mm', alignItems: 'flex-start' }}>
              <div style={{ width: '45%' }}>
                <div style={{ fontWeight: 700, marginBottom: '1mm' }}>RESUME</div>
                <table>
                  <thead><tr><th colSpan="2">Sepatu</th><th colSpan="2">Baju seragam</th></tr><tr><th>Size</th><th>Qty</th><th>Size</th><th>Qty</th></tr></thead>
                  <tbody>
                    {Array.from({ length: Math.max(rekapSepatu.length, rekapSeragam.length, 1) }, (_, i) => (
                      <tr key={i}>
                        <td style={{ textAlign: 'center' }}>{rekapSepatu[i]?.[0] ?? ''}</td><td style={{ textAlign: 'center' }}>{rekapSepatu[i]?.[1] ?? ''}</td>
                        <td style={{ textAlign: 'center' }}>{rekapSeragam[i]?.[0] ?? ''}</td><td style={{ textAlign: 'center' }}>{rekapSeragam[i]?.[1] ?? ''}</td>
                      </tr>
                    ))}
                    <tr style={{ fontWeight: 700 }}><td style={{ textAlign: 'center' }}>TOTAL</td><td style={{ textAlign: 'center' }}>{totSepatu}</td><td style={{ textAlign: 'center' }}>TOTAL</td><td style={{ textAlign: 'center' }}>{totSeragam}</td></tr>
                  </tbody>
                </table>
              </div>
              <div style={{ flex: 1, fontSize: '10px' }}>
                <div style={{ fontWeight: 700 }}>NOTE :</div>
                <div>Penggantian sepatu: {JABATAN_SEPATU_6_BULAN.join(', ')} per 6 bulan; jabatan lain per 1 tahun. Seragam per 1 tahun, {QTY_SERAGAM} potong per orang.</div>
              </div>
            </div>
            <div style={{ textAlign: 'right', marginTop: '6mm' }}>Bontang, {fmtPanjang(today)}</div>
            <div style={{ display: 'flex', justifyContent: 'space-between', marginTop: '2mm', textAlign: 'center' }}>
              {TTD.map((t) => (
                <div key={t.nama} style={{ width: '30%' }}>
                  <div>{t.judul}</div>
                  <div style={{ height: '18mm' }} />
                  <div style={{ fontWeight: 700, textDecoration: 'underline' }}>{t.nama}</div>
                  <div>{t.jabatan}</div>
                </div>
              ))}
            </div>
          </div>
        </div>
      )}

      {/* ===================== MODAL: TAMBAH / UBAH KARYAWAN ===================== */}
      {form && (
        <Modal title={form.id ? 'Ubah Data Karyawan' : 'Tambah Karyawan'} onClose={() => setForm(null)} dark={darkMode} wide>
          <datalist id="apd-jabatan">{jabatanList.map((j) => <option key={j} value={j} />)}</datalist>
          <datalist id="apd-seragam">{UKURAN_SERAGAM.map((u) => <option key={u} value={u} />)}</datalist>
          <form onSubmit={(e) => { e.preventDefault(); simpanKaryawan(); }} className="grid grid-cols-2 gap-3 text-sm">
            <Label text="NRP Amodi * (acuan data)"><input className={inputClass} value={form.nrp_amodi} onChange={(e) => setForm({ ...form, nrp_amodi: e.target.value })} placeholder="RYEP0099" /></Label>
            <Label text="NRP (REP)"><input className={inputClass} value={form.nrp} onChange={(e) => setForm({ ...form, nrp: e.target.value })} placeholder="REP0099" /></Label>
            <Label text="Nama *" className="col-span-2"><input className={inputClass} value={form.nama} onChange={(e) => setForm({ ...form, nama: e.target.value })} /></Label>
            <Label text="Jabatan"><input className={inputClass} list="apd-jabatan" value={form.jabatan} onChange={(e) => setForm({ ...form, jabatan: e.target.value })} /></Label>
            <Label text="Tanggal masuk kerja"><input type="date" className={inputClass} value={form.tanggal_masuk} onChange={(e) => setForm({ ...form, tanggal_masuk: e.target.value })} /></Label>
            <Label text="Ukuran sepatu"><input className={inputClass} inputMode="numeric" value={form.ukuran_sepatu} onChange={(e) => setForm({ ...form, ukuran_sepatu: e.target.value })} placeholder="6" /></Label>
            <Label text="Ukuran seragam"><input className={inputClass} list="apd-seragam" value={form.ukuran_seragam} onChange={(e) => setForm({ ...form, ukuran_seragam: e.target.value })} placeholder="L" /></Label>
            <Label text={`Tanda terima sepatu terakhir${form.kat_sepatu === 'rusak' ? ' (ganti rusak)' : ''}`}><input type="date" max={today} className={inputClass} value={form.terima_sepatu} onChange={(e) => setForm({ ...form, terima_sepatu: e.target.value })} /></Label>
            <Label text={`Tanda terima seragam terakhir`}><input type="date" max={today} className={inputClass} value={form.terima_seragam} onChange={(e) => setForm({ ...form, terima_seragam: e.target.value })} /></Label>
            <p className="col-span-2 text-[11px] text-slate-400">
              {form.id
                ? 'Ubah tanggal bila salah memilih saat mencatat tanda terima. Kosongkan untuk membatalkan catatan terakhir itu. Jadwal berikutnya dihitung ulang otomatis; riwayat lengkap ada di ikon jam pada daftar.'
                : 'Bila dikosongkan, jadwal pertama dihitung dari tanggal masuk kerja.'}
            </p>
            <Label text="Keterangan" className="col-span-2"><input className={inputClass} value={form.keterangan} onChange={(e) => setForm({ ...form, keterangan: e.target.value })} /></Label>
            <div className="col-span-2 flex justify-end gap-2 pt-2">
              <button type="button" onClick={() => setForm(null)} className="px-4 py-2 rounded-lg border border-slate-300 text-xs font-semibold">Batal</button>
              <button type="submit" disabled={saving} className="px-4 py-2 rounded-lg bg-blue-600 text-white text-xs font-semibold disabled:opacity-50">{saving ? 'Menyimpan...' : 'Simpan'}</button>
            </div>
          </form>
        </Modal>
      )}

      {/* ===================== MODAL: CATAT SUDAH DITERIMA (tanggal aktual, sepatu dan seragam terpisah) ===================== */}
      {terimaModal && (
        <Modal title={terimaModal.rows.length === 1 ? `Catat Diterima - ${terimaModal.rows[0].k.nama}` : 'Catat Sudah Diterima'} onClose={() => setTerimaModal(null)} dark={darkMode} wide>
          {terimaModal.rows.length > 1 && ['sepatu', 'seragam'].map((jenis) => (
            <div key={jenis} className="flex flex-wrap items-center gap-2 mb-2 text-sm">
              <label className="text-xs font-semibold w-52">Tanggal terima {jenis} untuk semua:</label>
              <input type="date" max={today} value={terimaModal.semua[jenis]} onChange={(e) => setTerimaModal((m) => ({ ...m, semua: { ...m.semua, [jenis]: e.target.value } }))}
                className="px-3 py-1.5 text-sm border border-slate-300 rounded-lg bg-white text-slate-900 font-medium" />
              <button type="button" disabled={!terimaModal.semua[jenis] || terimaModal.semua[jenis] > today} onClick={() => ubahSemua(jenis, { tgl: terimaModal.semua[jenis] })}
                className="px-3 py-1.5 rounded-lg bg-slate-700 text-white text-xs font-semibold disabled:opacity-40">Terapkan ke semua {jenis}</button>
              <button type="button" onClick={() => ubahSemua(jenis, { on: true })} className="px-2 py-1.5 rounded-lg border border-slate-300 text-xs font-semibold">Centang semua</button>
              <button type="button" onClick={() => ubahSemua(jenis, { on: false })} className="px-2 py-1.5 rounded-lg border border-slate-300 text-xs font-semibold">Kosongkan</button>
            </div>
          ))}
          <p className="text-[11px] text-slate-400 mb-2">
            Sepatu dan seragam dicatat dengan tanggal terimanya sendiri. Hilangkan centang pada barang yang belum datang: barang itu tetap menunggu dan bisa dicatat nanti.
            Jadwal berikutnya dihitung dari tanggal terima aktual.
          </p>
          <div className={`rounded-lg border overflow-x-auto mb-4 ${card}`}>
            <table className="w-full min-w-[640px] text-sm">
              <thead className={`border-b font-semibold leading-tight ${thead}`}><tr><th className="p-2 text-left">NAMA</th><th className="p-2">SEPATU</th><th className="p-2">SERAGAM</th></tr></thead>
              <tbody className={`divide-y ${divide}`}>
                {terimaModal.rows.map((g) => (
                  <tr key={g.k.id}>
                    <td className="p-2 align-top"><div className="font-semibold">{g.k.nama}</div><div className="text-[10px] text-slate-400 font-mono">{g.k.nrp_amodi || g.k.nrp}</div></td>
                    {celTerima(g, 'sepatu')}
                    {celTerima(g, 'seragam')}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <div className="flex justify-end gap-2">
            <button type="button" onClick={() => setTerimaModal(null)} className="px-4 py-2 rounded-lg border border-slate-300 text-xs font-semibold">Batal</button>
            <button type="button" onClick={simpanTerima} disabled={saving || !Object.values(terimaModal.item).some((v) => v.on)} className="px-4 py-2 rounded-lg bg-emerald-600 text-white text-xs font-semibold disabled:opacity-50">
              {saving ? 'Menyimpan...' : `Simpan (${Object.values(terimaModal.item).filter((v) => v.on).length} item)`}
            </button>
          </div>
        </Modal>
      )}

      {/* ===================== MODAL: RESIGN ===================== */}
      {resign && (
        <Modal title="Tandai Karyawan Resign" onClose={() => setResign(null)} dark={darkMode}>
          <form onSubmit={(e) => { e.preventDefault(); simpanResign(); }} className="space-y-3 text-sm">
            <p><b>{resign.k.nama}</b> <span className="text-slate-400">({resign.k.nrp_amodi || resign.k.nrp} - {resign.k.jabatan || '-'})</span></p>
            <Label text="Tanggal resign"><input type="date" className={inputClass} value={resign.tanggal} onChange={(e) => setResign({ ...resign, tanggal: e.target.value })} /></Label>
            <p className="text-[11px] text-slate-400">Karyawan ini tidak akan muncul lagi di orderan, tanda terima, rekap, dan proyeksi sepatu serta seragam. Riwayat penerimaannya tetap tersimpan, dan bisa diaktifkan kembali dari filter "Resign".</p>
            <div className="flex justify-end gap-2 pt-1">
              <button type="button" onClick={() => setResign(null)} className="px-4 py-2 rounded-lg border border-slate-300 text-xs font-semibold">Batal</button>
              <button type="submit" disabled={saving} className="px-4 py-2 rounded-lg bg-rose-600 text-white text-xs font-semibold disabled:opacity-50">{saving ? 'Menyimpan...' : 'Tandai Resign'}</button>
            </div>
          </form>
        </Modal>
      )}

      {/* ===================== MODAL: SEPATU RUSAK DIGANTI ===================== */}
      {rusak && (
        <Modal title="Penggantian Sepatu Rusak" onClose={() => setRusak(null)} dark={darkMode}>
          <form onSubmit={(e) => { e.preventDefault(); simpanRusak(); }} className="space-y-3 text-sm">
            <p><b>{rusak.k.nama}</b> <span className="text-slate-400">({rusak.k.jabatan || '-'})</span></p>
            <Label text="Tanggal penggantian"><input type="date" className={inputClass} value={rusak.tanggal} onChange={(e) => setRusak({ ...rusak, tanggal: e.target.value })} /></Label>
            <Label text="Keterangan (mis. sol lepas)"><input className={inputClass} value={rusak.keterangan} onChange={(e) => setRusak({ ...rusak, keterangan: e.target.value })} /></Label>
            <p className="text-[11px] text-slate-400">
              Jadwal sepatu berikutnya dihitung mulai tanggal ini: {rusak.tanggal ? fmt(addMonths(rusak.tanggal, hitungJadwal(rusak.k, []).intervalSepatu)) : '-'}.
            </p>
            <div className="flex justify-end gap-2 pt-1">
              <button type="button" onClick={() => setRusak(null)} className="px-4 py-2 rounded-lg border border-slate-300 text-xs font-semibold">Batal</button>
              <button type="submit" disabled={saving} className="px-4 py-2 rounded-lg bg-amber-600 text-white text-xs font-semibold disabled:opacity-50">{saving ? 'Menyimpan...' : 'Catat Penggantian'}</button>
            </div>
          </form>
        </Modal>
      )}

      {/* ===================== MODAL: RIWAYAT PER KARYAWAN ===================== */}
      {riwayatK && (
        <Modal title={`Riwayat APD - ${riwayatK.k.nama}`} onClose={() => setRiwayatId(null)} dark={darkMode} wide>
          <p className="text-xs text-slate-400 mb-3">{riwayatK.k.nrp_amodi || riwayatK.k.nrp} - {riwayatK.k.jabatan || '-'}. Sepatu tiap {riwayatK.j.intervalSepatu} bulan, seragam tiap {riwayatK.j.intervalSeragam} bulan.</p>
          <div className={`rounded-lg border overflow-hidden mb-4 ${card}`}>
            <table className="w-full text-sm">
              <thead className={`border-b font-semibold ${thead}`}><tr><th className="p-2 text-left">TANGGAL TERIMA</th><th className="p-2 text-left">JENIS</th><th className="p-2 text-left">KATEGORI</th><th className="p-2 text-left">KETERANGAN</th><th className="p-2"></th></tr></thead>
              <tbody className={`divide-y ${divide}`}>
                {riwayatK.riwayat.length === 0 ? <tr><td colSpan="5" className="p-4 text-center text-slate-400">Belum ada catatan.</td></tr> : riwayatK.riwayat.map((r) => (
                  <tr key={r.id}>
                    <td className="p-2 whitespace-nowrap">{fmt(r.tanggal)}</td>
                    <td className="p-2">{r.jenis === 'sepatu' ? 'Sepatu' : 'Seragam'}</td>
                    <td className="p-2">{r.kategori === 'rusak' ? <span className="bg-amber-100 text-amber-800 px-2 py-0.5 rounded-full font-bold text-[10px]">Ganti rusak</span> : 'Rutin'}</td>
                    <td className="p-2 text-xs">{r.keterangan || ''}</td>
                    <td className="p-2 text-right"><button title="Hapus catatan" onClick={() => hapusRiwayat(r)} className="p-1 text-rose-500 hover:bg-rose-50 rounded"><Trash2 size={14} /></button></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <div className="text-xs font-bold mb-2">Tambah catatan manual (mis. melengkapi data lama)</div>
          <form onSubmit={async (e) => {
            e.preventDefault();
            const ok = await catat([{ karyawan_id: riwayatK.k.id, jenis: formRiwayat.jenis }], formRiwayat.tanggal, formRiwayat.jenis === 'seragam' ? 'rutin' : formRiwayat.kategori, formRiwayat.keterangan.trim());
            if (ok) { notify('Catatan ditambahkan.'); setFormRiwayat({ ...formRiwayat, keterangan: '' }); }
          }} className="grid grid-cols-2 sm:grid-cols-4 gap-2 items-end text-sm">
            <Label text="Jenis"><select className={inputClass} value={formRiwayat.jenis} onChange={(e) => setFormRiwayat({ ...formRiwayat, jenis: e.target.value })}><option value="sepatu">Sepatu</option><option value="seragam">Seragam</option></select></Label>
            <Label text="Kategori"><select className={inputClass} disabled={formRiwayat.jenis === 'seragam'} value={formRiwayat.jenis === 'seragam' ? 'rutin' : formRiwayat.kategori} onChange={(e) => setFormRiwayat({ ...formRiwayat, kategori: e.target.value })}><option value="rutin">Rutin</option><option value="rusak">Ganti rusak</option></select></Label>
            <Label text="Tanggal terima"><input type="date" className={inputClass} value={formRiwayat.tanggal} onChange={(e) => setFormRiwayat({ ...formRiwayat, tanggal: e.target.value })} /></Label>
            <button type="submit" disabled={saving} className="px-3 py-2 rounded-lg bg-blue-600 text-white text-xs font-semibold disabled:opacity-50 h-[38px]">Tambah</button>
            <Label text="Keterangan" className="col-span-2 sm:col-span-4"><input className={inputClass} value={formRiwayat.keterangan} onChange={(e) => setFormRiwayat({ ...formRiwayat, keterangan: e.target.value })} /></Label>
          </form>
        </Modal>
      )}
    </>
  );
}
