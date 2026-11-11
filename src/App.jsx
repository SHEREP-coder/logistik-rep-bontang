import React, { useState, useEffect, useMemo } from 'react';
import * as XLSX from 'xlsx';
import {
  Search, Plus, LayoutDashboard, Box, ArrowDownToLine, ArrowUpFromLine,
  FileText, ClipboardList, BookOpen, BarChart3, Users, Settings, Moon, TrendingUp,
  TrendingDown, Calendar, X, Download, Trash2, AlertTriangle, ShoppingCart, Barcode,
  Printer, History, LogOut, Pencil, Upload, Lock, Save, FilePlus2
} from 'lucide-react';

import { initializeApp } from 'firebase/app';
import {
  initializeFirestore, persistentLocalCache, persistentMultipleTabManager,
  collection, doc, getDoc, setDoc, addDoc, updateDoc, deleteDoc, onSnapshot,
  writeBatch, runTransaction, serverTimestamp
} from 'firebase/firestore';
import {
  getAuth, onAuthStateChanged, signInWithEmailAndPassword, signOut,
  createUserWithEmailAndPassword, updatePassword, reauthenticateWithCredential,
  EmailAuthProvider
} from 'firebase/auth';

/* ============================== FIREBASE ============================== */
const firebaseConfig = {
  apiKey: "AIzaSyAuetnYQDspE1IFbOcFttgSJMq3FOvmft0",
  authDomain: "logistik-rep-bontang.firebaseapp.com",
  projectId: "logistik-rep-bontang",
  storageBucket: "logistik-rep-bontang.firebasestorage.app",
  messagingSenderId: "26293450683",
  appId: "1:26293450683:web:8d6458695a511b75ea369e",
  measurementId: "G-E72GBTQ59N"
};

const app = initializeApp(firebaseConfig);
// Cache lokal: data tidak dibaca ulang penuh setiap buka halaman (hemat kuota baca Firestore)
const db = initializeFirestore(app, {
  localCache: persistentLocalCache({ tabManager: persistentMultipleTabManager() })
});
const auth = getAuth(app);

// App kedua khusus untuk membuat akun pengguna baru tanpa mengeluarkan Admin yang sedang login
const secondaryApp = initializeApp(firebaseConfig, 'secondary');
const secondaryAuth = getAuth(secondaryApp);

/* ============================== HELPER ============================== */
const ROWS = 20;
const PO_KANTOR = 'BALIKPAPAN'; // tulisan di bawah nama PT pada blangko PO
const MONTHS_ID = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des'];
const INK = '#1e3a5f';

const usernameToEmail = (u) => {
  const s = String(u || '').trim().toLowerCase();
  return s.includes('@') ? s : `${s}@logistik-rep.local`;
};

const safeId = (kode) => {
  let id = String(kode).trim().replace(/\//g, '_');
  if (id === '.' || id === '..' || /^__.*__$/.test(id)) id = 'K_' + id;
  return id.slice(0, 500);
};
const clean = (obj) => Object.fromEntries(Object.entries(obj).filter(([, v]) => v !== undefined));
const pad2 = (n) => String(n).padStart(2, '0');
const todayISO = () => { const d = new Date(); return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`; };

const toISODate = (v) => {
  if (v === undefined || v === null || v === '') return '';
  if (typeof v === 'number' || (typeof v === 'string' && /^\d+(\.\d+)?$/.test(v.trim()))) {
    const d = new Date(Math.round((Number(v) - 25569) * 86400 * 1000));
    return isNaN(d.getTime()) ? '' : d.toISOString().slice(0, 10);
  }
  const s = String(v).trim();
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
  const m = s.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/);
  if (m) return `${m[3]}-${pad2(m[2])}-${pad2(m[1])}`;
  const d = new Date(s);
  return isNaN(d.getTime()) ? '' : `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
};
const dateParts = (v) => {
  const iso = toISODate(v);
  if (!iso) return null;
  const [y, m, d] = iso.split('-').map(Number);
  return { y, m, d };
};
const fmtDate = (v) => {
  const p = dateParts(v);
  return p ? `${p.d} ${MONTHS_ID[p.m - 1]} ${p.y}` : '-';
};

const poNumberValue = (s) => { const m = String(s || '').match(/(\d+)\s*$/); return m ? Number(m[1]) : 0; };
const formatPoNumber = (n) => `A ${String(n).padStart(6, '0')}`;

const chunk = (arr, size) => { const out = []; for (let i = 0; i < arr.length; i += size) out.push(arr.slice(i, i + size)); return out; };

const stripMeta = (row) => {
  const { id, createdAt, ...rest } = row;
  return rest;
};

/* ============================== KOMPONEN KECIL ============================== */
function Pager({ page, setPage, total, dark }) {
  const totalPage = Math.ceil(total / ROWS) || 1;
  const p = Math.min(page, totalPage);
  return (
    <div className={`p-4 border-t flex justify-between items-center text-xs ${dark ? 'border-slate-700 text-slate-300' : 'border-slate-200 text-slate-700'}`}>
      <span className="font-semibold">Halaman <span className="text-yellow-500 font-bold">{p}</span> dari {totalPage} (Total: {total} Data)</span>
      <div className="flex gap-2">
        <button onClick={() => setPage(Math.max(p - 1, 1))} disabled={p === 1} className="px-3 py-1.5 rounded-lg bg-emerald-600 text-white font-bold disabled:opacity-40 shadow">◀ Sebelumnya</button>
        <button onClick={() => setPage(Math.min(p + 1, totalPage))} disabled={p >= totalPage} className="px-3 py-1.5 rounded-lg bg-emerald-600 text-white font-bold disabled:opacity-40 shadow">Berikutnya ▶</button>
      </div>
    </div>
  );
}

function Modal({ title, icon, onClose, dark, children, wide }) {
  return (
    <div className="fixed inset-0 bg-black bg-opacity-60 flex items-center justify-center z-50 p-4 no-print">
      <div className={`w-full ${wide ? 'max-w-xl' : 'max-w-lg'} max-h-[92vh] overflow-y-auto p-6 rounded-xl shadow-2xl border ${dark ? 'bg-slate-800 border-slate-700 text-white' : 'bg-white border-slate-200 text-slate-800'}`}>
        <div className="flex justify-between items-center mb-4 border-b pb-2">
          <h3 className="font-bold text-base flex items-center gap-1.5">{icon}{title}</h3>
          <button onClick={onClose}><X size={20} /></button>
        </div>
        {children}
      </div>
    </div>
  );
}

/* ============================== HALAMAN LOGIN ============================== */
function LoginScreen({ onLogin, error, busy }) {
  const [username, setUsername] = useState('');
  const [password, setPassword] = useState('');
  return (
    <div className="min-h-screen flex items-center justify-center bg-slate-900 p-4">
      <form
        onSubmit={(e) => { e.preventDefault(); onLogin(username, password); }}
        className="w-full max-w-sm bg-white rounded-2xl shadow-2xl p-8 space-y-4"
      >
        <div className="flex flex-col items-center gap-2 pb-2">
          <img src="/logo_PT_REP_transparan.png" alt="Logo" className="w-16 h-16 object-contain" />
          <h1 className="text-sm font-extrabold uppercase tracking-wider text-slate-900">PT. Ryan Eka Pratama</h1>
          <p className="text-[11px] font-semibold text-yellow-600">Logistic Management System - Site Bontang</p>
        </div>
        <div>
          <label className="block text-xs font-semibold text-slate-700 mb-1">Username</label>
          <input autoFocus required value={username} onChange={(e) => setUsername(e.target.value)} autoComplete="username"
            className="w-full p-2.5 border border-slate-300 rounded-lg text-sm text-slate-900 focus:outline-none focus:border-yellow-500" placeholder="contoh: admin_rep" />
        </div>
        <div>
          <label className="block text-xs font-semibold text-slate-700 mb-1">Password</label>
          <input required type="password" value={password} onChange={(e) => setPassword(e.target.value)} autoComplete="current-password"
            className="w-full p-2.5 border border-slate-300 rounded-lg text-sm text-slate-900 focus:outline-none focus:border-yellow-500" placeholder="••••••••" />
        </div>
        {error && <div className="text-xs bg-rose-50 border border-rose-200 text-rose-700 rounded-lg p-2.5 font-medium">{error}</div>}
        <button disabled={busy} className="w-full bg-yellow-500 hover:bg-yellow-400 disabled:opacity-60 text-slate-950 font-bold py-2.5 rounded-lg text-sm flex items-center justify-center gap-2">
          <Lock size={16} /> {busy ? 'Memproses...' : 'Masuk'}
        </button>
        <p className="text-[10px] text-center text-slate-400">Akun dibuat oleh Admin. Hubungi Admin jika lupa password.</p>
      </form>
    </div>
  );
}

/* ============================== APLIKASI UTAMA ============================== */
export default function App() {
  /* ---------- auth ---------- */
  const [authLoading, setAuthLoading] = useState(true);
  const [authBusy, setAuthBusy] = useState(false);
  const [authError, setAuthError] = useState('');
  const [profile, setProfile] = useState(null);

  /* ---------- ui ---------- */
  const [activeTab, setActiveTab] = useState('Dashboard');
  const [pencarian, setPencarian] = useState('');
  const [page, setPage] = useState(1);
  const [selectedMonth, setSelectedMonth] = useState('');
  const [selectedYear, setSelectedYear] = useState('');
  const [darkMode, setDarkMode] = useState(false);
  const [fontSize, setFontSize] = useState(() => { try { return Number(localStorage.getItem('rep_font')) || 18; } catch (e) { return 18; } });
  useEffect(() => {
    document.documentElement.style.fontSize = fontSize + 'px';
    try { localStorage.setItem('rep_font', String(fontSize)); } catch (e) { /* abaikan */ }
  }, [fontSize]);
  const [loadingCloud, setLoadingCloud] = useState(false);
  const [saving, setSaving] = useState(false);
  const [toast, setToast] = useState(null);

  /* ---------- data firestore ---------- */
  const [dataMaterial, setDataMaterial] = useState([]);
  const [dataMasuk, setDataMasuk] = useState([]);
  const [dataKeluar, setDataKeluar] = useState([]);
  const [riwayatPo, setRiwayatPo] = useState([]);
  const [riwayatOpname, setRiwayatOpname] = useState([]);
  const [daftarPengguna, setDaftarPengguna] = useState([]);
  const [riwayatLoaded, setRiwayatLoaded] = useState(false);

  /* ---------- modal & form ---------- */
  const [showModalMaterial, setShowModalMaterial] = useState(false);
  const [editingMaterialId, setEditingMaterialId] = useState(null);
  const [showModalMasuk, setShowModalMasuk] = useState(false);
  const [showModalKeluar, setShowModalKeluar] = useState(false);
  const [showModalUser, setShowModalUser] = useState(false);

  const emptyMaterial = { kode: '', nama: '', nomor: '', unit: 'Ea', lokasi: 'Rak A1', minQty: 5, qty: 0 };
  const emptyMasuk = { tanggal: todayISO(), unit: '', kode: '', description: '', merk: '', qty: '', satuan: 'Pcs', noPo: '' };
  const emptyKeluar = { tanggal: todayISO(), unit: '', kode: '', namaBarang: '', mekanik: '', qty: '', keterangan: '' };
  const [formMaterial, setFormMaterial] = useState(emptyMaterial);
  const [formMasuk, setFormMasuk] = useState(emptyMasuk);
  const [formKeluar, setFormKeluar] = useState(emptyKeluar);
  const [formUser, setFormUser] = useState({ nama: '', username: '', password: '', role: 'Editor' });

  /* ---------- opname / kartu stok ---------- */
  const [physicalStocks, setPhysicalStocks] = useState({});
  const [selectedItemCode, setSelectedItemCode] = useState('');
  const [kartuSearch, setKartuSearch] = useState('');

  /* ---------- PO ---------- */
  const emptyPoItem = () => ({ id: Date.now() + Math.random(), uraian: '', partNumber: '', qty: '', unitPrice: '', keterangan: '' });
  const [poHeader, setPoHeader] = useState({ noPo: '', equipment: '', kodeAlat: '', lokasi: 'Bontang', tanggal: todayISO(), diketahuiOleh: '' });
  const [poItems, setPoItems] = useState([emptyPoItem()]);
  const [editingPoId, setEditingPoId] = useState(null);

  /* ---------- pengaturan ---------- */
  const [pwForm, setPwForm] = useState({ lama: '', baru: '', ulang: '' });
  const [importMsg, setImportMsg] = useState('');
  const [importing, setImporting] = useState(false);

  const role = profile?.role;
  const isAdmin = role === 'Admin';
  const canEdit = role === 'Admin' || role === 'Editor';

  const notify = (msg, type = 'ok') => setToast({ msg, type, t: Date.now() });
  useEffect(() => {
    if (!toast) return;
    const id = setTimeout(() => setToast(null), 4000);
    return () => clearTimeout(id);
  }, [toast]);

  const goTab = (name) => { setActiveTab(name); setPencarian(''); setPage(1); };

  /* ---------- AUTH ---------- */
  useEffect(() => {
    const unsub = onAuthStateChanged(auth, async (u) => {
      if (!u) { setProfile(null); setAuthLoading(false); return; }
      try {
        const snap = await getDoc(doc(db, 'users', u.uid));
        if (snap.exists() && snap.data().status !== 'Nonaktif') {
          setProfile({ uid: u.uid, email: u.email, ...snap.data() });
          setAuthError('');
        } else {
          setAuthError(snap.exists() ? 'Akun Anda dinonaktifkan. Hubungi Admin.' : 'Akun belum terdaftar di sistem. Hubungi Admin.');
          await signOut(auth);
          setProfile(null);
        }
      } catch (err) {
        console.error(err);
        setAuthError('Gagal membaca profil pengguna. Periksa Firestore Rules.');
        await signOut(auth);
        setProfile(null);
      }
      setAuthLoading(false);
    });
    return () => unsub();
  }, []);

  const handleLogin = async (username, password) => {
    setAuthBusy(true); setAuthError('');
    try {
      await signInWithEmailAndPassword(auth, usernameToEmail(username), password);
    } catch (err) {
      const code = err.code || '';
      setAuthError(
        code.includes('invalid-credential') || code.includes('wrong-password') || code.includes('user-not-found')
          ? 'Username atau password salah.'
          : code.includes('too-many-requests') ? 'Terlalu banyak percobaan. Coba lagi beberapa menit lagi.'
          : code.includes('network') ? 'Tidak ada koneksi internet.'
          : 'Gagal login: ' + (err.message || code)
      );
    }
    setAuthBusy(false);
  };

  const handleLogout = async () => {
    await signOut(auth);
    setActiveTab('Dashboard');
    setDataMaterial([]); setDataMasuk([]); setDataKeluar([]); setRiwayatPo([]); setRiwayatOpname([]); setDaftarPengguna([]);
    setRiwayatLoaded(false);
  };

  /* ---------- LANGGANAN DATA FIRESTORE (realtime) ---------- */
  useEffect(() => {
    if (!profile) return;
    setLoadingCloud(true);
    let pending = 5;
    const finish = () => { pending -= 1; if (pending <= 0) setLoadingCloud(false); };
    const unsubs = [];
    const sub = (name, setter, extra) => {
      let first = true;
      const un = onSnapshot(collection(db, name), (snap) => {
        setter(snap.docs.map((d) => ({ id: d.id, ...d.data() })));
        if (extra) extra();
        if (first) { first = false; finish(); }
      }, (err) => {
        console.error(name, err);
        notify(`Gagal memuat "${name}": ${err.code || err.message}`, 'error');
        if (first) { first = false; finish(); }
      });
      unsubs.push(un);
    };
    sub('material', setDataMaterial);
    sub('barangMasuk', setDataMasuk);
    sub('barangKeluar', setDataKeluar);
    sub('riwayatPo', setRiwayatPo, () => setRiwayatLoaded(true));
    sub('stokOpname', setRiwayatOpname);
    if (profile.role === 'Admin') {
      sub('users', setDaftarPengguna);
    }
    return () => unsubs.forEach((u) => u());
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [profile?.uid]);

  /* ---------- TURUNAN DATA ---------- */
  const materialByKode = useMemo(() => {
    const m = new Map();
    dataMaterial.forEach((x) => { if (x.Kode) m.set(String(x.Kode).toLowerCase(), x); });
    return m;
  }, [dataMaterial]);

  const sortedMasuk = useMemo(() => [...dataMasuk].sort((a, b) => toISODate(b.Tanggal).localeCompare(toISODate(a.Tanggal))), [dataMasuk]);
  const sortedKeluar = useMemo(() => [...dataKeluar].sort((a, b) => toISODate(b['Doc.Date']).localeCompare(toISODate(a['Doc.Date']))), [dataKeluar]);

  const matchesPeriod = (v) => {
    if (!selectedMonth && !selectedYear) return true;
    const p = dateParts(v);
    if (!p) return false;
    if (selectedMonth && p.m !== Number(selectedMonth)) return false;
    if (selectedYear && p.y !== Number(selectedYear)) return false;
    return true;
  };

  const q = pencarian.toLowerCase();
  const has = (v) => String(v ?? '').toLowerCase().includes(q);

  const filteredMaterial = useMemo(() => dataMaterial.filter((i) => !q || has(i['Nama Barang']) || has(i.Kode) || has(i['Nomor Barang']) || has(i.Lokasi)), [dataMaterial, q]); // eslint-disable-line
  const filteredMasuk = sortedMasuk.filter((i) => (!q || has(i.Description) || has(i.Kode) || has(i['No. PO']) || has(i.Unit)) && matchesPeriod(i.Tanggal));
  const filteredKeluar = sortedKeluar.filter((i) => (!q || has(i['Nama Barang']) || has(i.Kode) || has(i.Unit) || has(i.Mekanik)) && matchesPeriod(i['Doc.Date']));

  const poNeededItems = useMemo(() => dataMaterial.filter((i) => Number(i.Qty || 0) <= Number(i.MinQty ?? 5)), [dataMaterial]);

  const totalsByKode = useMemo(() => {
    const t = new Map();
    dataMasuk.forEach((r) => { const k = String(r.Kode || '').toLowerCase(); const c = t.get(k) || { in: 0, out: 0 }; c.in += Number(r.Qty || 0); t.set(k, c); });
    dataKeluar.forEach((r) => { const k = String(r.Kode || '').toLowerCase(); const c = t.get(k) || { in: 0, out: 0 }; c.out += Number(r.Qty || 0); t.set(k, c); });
    return t;
  }, [dataMasuk, dataKeluar]);

  const paged = (arr) => {
    const tp = Math.ceil(arr.length / ROWS) || 1;
    const p = Math.min(page, tp);
    return arr.slice((p - 1) * ROWS, p * ROWS);
  };

  // Nomor PO otomatis (lanjutan dari PO terakhir)
  const nextPoNo = useMemo(() => {
    const max = riwayatPo.reduce((m, p) => Math.max(m, poNumberValue(p.noPo)), 85523);
    return formatPoNumber(max + 1);
  }, [riwayatPo]);
  useEffect(() => {
    if (riwayatLoaded && !poHeader.noPo && !editingPoId) setPoHeader((h) => ({ ...h, noPo: nextPoNo }));
  }, [riwayatLoaded, nextPoNo]); // eslint-disable-line

  useEffect(() => {
    if (profile && !poHeader.diketahuiOleh) { /* dibiarkan kosong untuk tanda tangan */ }
  }, [profile]); // eslint-disable-line

  /* ============================== AKSI: MASTER BARANG ============================== */
  const openNewMaterial = () => { setEditingMaterialId(null); setFormMaterial(emptyMaterial); setShowModalMaterial(true); };
  const openEditMaterial = (item) => {
    setEditingMaterialId(item.id);
    setFormMaterial({
      kode: item.Kode || '', nama: item['Nama Barang'] || '', nomor: item['Nomor Barang'] ?? '',
      unit: item['Part Unit'] || 'Ea', lokasi: item.Lokasi || 'Rak A1', minQty: item.MinQty ?? 5, qty: item.Qty ?? 0
    });
    setShowModalMaterial(true);
  };

  const handleSaveMaterial = async (e) => {
    e.preventDefault();
    if (!canEdit) return notify('Akses ditolak!', 'error');
    const kode = formMaterial.kode.trim();
    if (!kode) return notify('Kode wajib diisi', 'error');
    const base = {
      'Nama Barang': formMaterial.nama.trim(),
      'Nomor Barang': String(formMaterial.nomor).trim(),
      'Part Unit': formMaterial.unit.trim() || 'Ea',
      Lokasi: formMaterial.lokasi.trim() || 'Rak A1',
      MinQty: Number(formMaterial.minQty) || 0
    };
    setSaving(true);
    try {
      if (editingMaterialId) {
        await updateDoc(doc(db, 'material', editingMaterialId), { ...base, updatedAt: serverTimestamp(), updatedBy: profile.username });
        notify('Data barang diperbarui.');
      } else {
        const id = safeId(kode);
        if (dataMaterial.some((m) => m.id === id || String(m.Kode).toLowerCase() === kode.toLowerCase())) {
          setSaving(false);
          return notify('Kode barang sudah ada di Master Barang.', 'error');
        }
        await setDoc(doc(db, 'material', id), {
          Kode: kode, ...base, 'Base Unit of Measure': 'Ea',
          Qty: Number(formMaterial.qty) || 0, createdAt: serverTimestamp(), createdBy: profile.username
        });
        notify('Produk baru berhasil disimpan.');
      }
      setShowModalMaterial(false);
    } catch (err) { console.error(err); notify('Gagal menyimpan: ' + (err.code || err.message), 'error'); }
    setSaving(false);
  };

  const handleDeleteMaterial = async (item) => {
    if (!isAdmin) return notify('Hanya Admin yang dapat menghapus.', 'error');
    if (!window.confirm(`Hapus barang ${item.Kode} - ${item['Nama Barang']}? Riwayat transaksi tidak ikut terhapus.`)) return;
    try { await deleteDoc(doc(db, 'material', item.id)); notify('Barang dihapus.'); }
    catch (err) { notify('Gagal menghapus: ' + (err.code || err.message), 'error'); }
  };

  /* ============================== AKSI: STOK MASUK / KELUAR ============================== */
  const lookupMasuk = (kode) => {
    const found = materialByKode.get(String(kode).trim().toLowerCase());
    setFormMasuk((p) => ({ ...p, kode, ...(found ? { description: found['Nama Barang'] || '', merk: found['Nomor Barang'] ?? '' } : {}) }));
  };
  const lookupKeluar = (kode) => {
    const found = materialByKode.get(String(kode).trim().toLowerCase());
    setFormKeluar((p) => ({ ...p, kode, ...(found ? { namaBarang: found['Nama Barang'] || '' } : {}) }));
  };
  const preventEnterSubmit = (e) => { if (e.key === 'Enter') e.preventDefault(); };

  const handleAddMasuk = async (e) => {
    e.preventDefault();
    if (!canEdit) return notify('Akses ditolak!', 'error');
    const mat = materialByKode.get(formMasuk.kode.trim().toLowerCase());
    if (!mat) return notify('Kode tidak ditemukan di Master Barang. Tambahkan dulu lewat "Produk Baru".', 'error');
    const qty = Number(formMasuk.qty);
    if (!(qty > 0)) return notify('Qty harus lebih dari 0.', 'error');
    setSaving(true);
    try {
      const matRef = doc(db, 'material', mat.id);
      const newRef = doc(collection(db, 'barangMasuk'));
      await runTransaction(db, async (t) => {
        const s = await t.get(matRef);
        if (!s.exists()) throw new Error('Master barang tidak ditemukan');
        t.update(matRef, { Qty: Number(s.data().Qty || 0) + qty });
        t.set(newRef, {
          Tanggal: formMasuk.tanggal, Unit: formMasuk.unit.trim(), Kode: mat.Kode,
          Description: formMasuk.description.trim(), Merk: String(formMasuk.merk).trim(),
          Qty: qty, Satuan: formMasuk.satuan.trim(), 'No. PO': formMasuk.noPo.trim(),
          createdAt: serverTimestamp(), createdBy: profile.username
        });
      });
      setShowModalMasuk(false); setFormMasuk(emptyMasuk);
      notify('Stok masuk tercatat & stok bertambah.');
    } catch (err) { console.error(err); notify('Gagal: ' + (err.message || err.code), 'error'); }
    setSaving(false);
  };

  const handleAddKeluar = async (e) => {
    e.preventDefault();
    if (!canEdit) return notify('Akses ditolak!', 'error');
    const mat = materialByKode.get(formKeluar.kode.trim().toLowerCase());
    if (!mat) return notify('Kode tidak ditemukan di Master Barang.', 'error');
    const qty = Number(formKeluar.qty);
    if (!(qty > 0)) return notify('Qty harus lebih dari 0.', 'error');
    setSaving(true);
    try {
      const matRef = doc(db, 'material', mat.id);
      const newRef = doc(collection(db, 'barangKeluar'));
      await runTransaction(db, async (t) => {
        const s = await t.get(matRef);
        if (!s.exists()) throw new Error('Master barang tidak ditemukan');
        const cur = Number(s.data().Qty || 0);
        if (cur < qty) throw new Error(`Stok tidak cukup. Stok saat ini: ${cur}`);
        t.update(matRef, { Qty: cur - qty });
        t.set(newRef, {
          'Doc.Date': formKeluar.tanggal, Unit: formKeluar.unit.trim(), Kode: mat.Kode,
          'Nama Barang': formKeluar.namaBarang.trim(), Mekanik: formKeluar.mekanik.trim(), Qty: qty,
          Keterangan: formKeluar.keterangan.trim(), createdAt: serverTimestamp(), createdBy: profile.username
        });
      });
      setShowModalKeluar(false); setFormKeluar(emptyKeluar);
      notify('Stok keluar tercatat & stok berkurang.');
    } catch (err) { console.error(err); notify('Gagal: ' + (err.message || err.code), 'error'); }
    setSaving(false);
  };

  // Hapus transaksi sekaligus mengembalikan stok
  const handleDeleteTrx = async (type, item) => {
    if (!isAdmin) return notify('Hanya Admin yang dapat menghapus transaksi.', 'error');
    if (!window.confirm('Hapus transaksi ini? Stok barang akan dikoreksi otomatis.')) return;
    const colName = type === 'masuk' ? 'barangMasuk' : 'barangKeluar';
    const mat = materialByKode.get(String(item.Kode || '').toLowerCase());
    const qty = Number(item.Qty || 0);
    try {
      await runTransaction(db, async (t) => {
        const trxRef = doc(db, colName, item.id);
        if (mat) {
          const matRef = doc(db, 'material', mat.id);
          const s = await t.get(matRef);
          if (s.exists()) {
            const cur = Number(s.data().Qty || 0);
            if (type === 'masuk' && cur < qty) throw new Error(`Stok saat ini (${cur}) lebih kecil dari qty transaksi (${qty}); tidak bisa dibatalkan.`);
            t.update(matRef, { Qty: type === 'masuk' ? cur - qty : cur + qty });
          }
        }
        t.delete(trxRef);
      });
      notify('Transaksi dihapus & stok dikoreksi.');
    } catch (err) { notify('Gagal: ' + (err.message || err.code), 'error'); }
  };

  /* ============================== AKSI: STOK OPNAME ============================== */
  const opnameChanges = useMemo(() => dataMaterial.filter((m) => {
    const v = physicalStocks[m.id];
    return v !== undefined && v !== '' && Number(v) !== Number(m.Qty || 0);
  }), [dataMaterial, physicalStocks]);

  const saveOpname = async () => {
    if (!canEdit) return notify('Akses ditolak!', 'error');
    if (opnameChanges.length === 0) return notify('Tidak ada selisih yang perlu disimpan.', 'error');
    if (!window.confirm(`Simpan opname? Stok ${opnameChanges.length} barang akan disesuaikan dengan hitungan fisik.`)) return;
    setSaving(true);
    try {
      const tanggal = todayISO();
      const items = opnameChanges.map((m) => {
        const fisik = Number(physicalStocks[m.id]);
        const sistem = Number(m.Qty || 0);
        return { kode: m.Kode, nama: m['Nama Barang'] || '', sistem, fisik, selisih: fisik - sistem };
      });
      for (const part of chunk(opnameChanges, 400)) {
        const batch = writeBatch(db);
        part.forEach((m) => batch.update(doc(db, 'material', m.id), { Qty: Number(physicalStocks[m.id]) }));
        await batch.commit();
      }
      await addDoc(collection(db, 'stokOpname'), { tanggal, items, createdAt: serverTimestamp(), createdBy: profile.username });
      setPhysicalStocks({});
      notify(`Opname tersimpan. ${items.length} barang disesuaikan.`);
    } catch (err) { console.error(err); notify('Gagal menyimpan opname: ' + (err.code || err.message), 'error'); }
    setSaving(false);
  };

  /* ============================== KARTU STOK ============================== */
  const kartu = useMemo(() => {
    if (!selectedItemCode) return null;
    const key = selectedItemCode.toLowerCase();
    const mat = materialByKode.get(key);
    if (!mat) return null;
    const rows = [];
    dataMasuk.filter((r) => String(r.Kode || '').toLowerCase() === key).forEach((r) =>
      rows.push({ date: toISODate(r.Tanggal), ket: `Masuk${r['No. PO'] ? ' - PO ' + r['No. PO'] : ''}${r.Unit ? ' / ' + r.Unit : ''}`, masuk: Number(r.Qty || 0), keluar: 0 }));
    dataKeluar.filter((r) => String(r.Kode || '').toLowerCase() === key).forEach((r) =>
      rows.push({ date: toISODate(r['Doc.Date']), ket: `Keluar${r.Unit ? ' - ' + r.Unit : ''}${r.Mekanik ? ' (' + r.Mekanik + ')' : ''}${r.Keterangan ? ' - ' + r.Keterangan : ''}`, masuk: 0, keluar: Number(r.Qty || 0) }));
    riwayatOpname.forEach((o) => (o.items || []).filter((i) => String(i.kode || '').toLowerCase() === key).forEach((i) =>
      rows.push({ date: o.tanggal, ket: 'Penyesuaian Stok Opname', masuk: i.selisih > 0 ? i.selisih : 0, keluar: i.selisih < 0 ? -i.selisih : 0 })));
    rows.sort((a, b) => a.date.localeCompare(b.date));
    const sumIn = rows.reduce((s, r) => s + r.masuk, 0);
    const sumOut = rows.reduce((s, r) => s + r.keluar, 0);
    const saldoAwal = Number(mat.Qty || 0) - sumIn + sumOut;
    let saldo = saldoAwal;
    const withSaldo = rows.map((r) => { saldo += r.masuk - r.keluar; return { ...r, saldo }; });
    return { mat, saldoAwal, rows: withSaldo, sumIn, sumOut };
  }, [selectedItemCode, materialByKode, dataMasuk, dataKeluar, riwayatOpname]);

  const kartuOptions = useMemo(() => {
    const s = kartuSearch.toLowerCase();
    return dataMaterial.filter((m) => !s || String(m.Kode).toLowerCase().includes(s) || String(m['Nama Barang'] || '').toLowerCase().includes(s) || String(m['Nomor Barang'] ?? '').toLowerCase().includes(s)).slice(0, 100);
  }, [dataMaterial, kartuSearch]);

  /* ============================== AKSI: PO ============================== */
  const updatePoItem = (id, field, value) => setPoItems((items) => items.map((it) => (it.id === id ? { ...it, [field]: value } : it)));
  const addPoRow = () => setPoItems((items) => [...items, emptyPoItem()]);
  const removePoRow = (id) => setPoItems((items) => (items.length > 1 ? items.filter((it) => it.id !== id) : [emptyPoItem()]));

  const fillPoFromNeeded = () => {
    if (poNeededItems.length === 0) return notify('Tidak ada barang yang wajib order.', 'error');
    const hasContent = poItems.some((i) => i.uraian || i.partNumber);
    if (hasContent && !window.confirm('Isi tabel PO akan diganti dengan daftar barang wajib order. Lanjutkan?')) return;
    setPoItems(poNeededItems.map((m) => {
      const cur = Number(m.Qty || 0); const min = Number(m.MinQty ?? 5);
      return { id: Date.now() + Math.random(), uraian: m['Nama Barang'] || '', partNumber: String(m['Nomor Barang'] || m.Kode || ''), qty: String(Math.max(1, min * 2 - cur)), unitPrice: '', keterangan: `Stok ${cur} / Min ${min}` };
    }));
  };

  const resetPoForm = () => {
    setEditingPoId(null);
    setPoHeader({ noPo: nextPoNo, equipment: '', kodeAlat: '', lokasi: 'Bontang', tanggal: todayISO(), diketahuiOleh: '' });
    setPoItems([emptyPoItem()]);
  };

  const savePo = async () => {
    if (!canEdit) return notify('Akses ditolak!', 'error');
    const items = poItems.filter((i) => i.uraian.trim() || i.partNumber.trim());
    if (items.length === 0) return notify('Isi minimal satu baris barang.', 'error');
    if (!poHeader.noPo.trim()) return notify('Nomor PO wajib diisi.', 'error');
    if (riwayatPo.some((p) => p.noPo === poHeader.noPo.trim() && p.id !== editingPoId)) return notify('Nomor PO sudah dipakai di arsip.', 'error');
    setSaving(true);
    try {
      const payload = {
        ...poHeader, noPo: poHeader.noPo.trim(), dibuatOleh: profile.nama,
        items: items.map(({ id, ...r }) => r)
      };
      if (editingPoId) {
        await updateDoc(doc(db, 'riwayatPo', editingPoId), { ...payload, updatedAt: serverTimestamp() });
        notify('PO di arsip diperbarui.');
      } else {
        await addDoc(collection(db, 'riwayatPo'), { ...payload, createdAt: serverTimestamp(), createdBy: profile.username });
        notify('PO disimpan ke arsip.');
        setEditingPoId(null);
        setPoHeader((h) => ({ ...h, noPo: formatPoNumber(poNumberValue(h.noPo) + 1) }));
        setPoItems([emptyPoItem()]);
      }
    } catch (err) { console.error(err); notify('Gagal menyimpan PO: ' + (err.code || err.message), 'error'); }
    setSaving(false);
  };

  const openPoFromArchive = (po) => {
    setEditingPoId(po.id);
    setPoHeader({ noPo: po.noPo || '', equipment: po.equipment || '', kodeAlat: po.kodeAlat || '', lokasi: po.lokasi || '', tanggal: po.tanggal || todayISO(), diketahuiOleh: po.diketahuiOleh || '' });
    setPoItems((po.items || []).map((i) => ({ id: Date.now() + Math.random(), uraian: i.uraian || '', partNumber: i.partNumber || '', qty: String(i.qty ?? ''), unitPrice: String(i.unitPrice ?? ''), keterangan: i.keterangan || '' })));
    goTab('Form & Cetak PO');
  };

  const deletePo = async (po) => {
    if (!isAdmin) return notify('Hanya Admin yang dapat menghapus arsip PO.', 'error');
    if (!window.confirm(`Hapus arsip PO ${po.noPo}?`)) return;
    try { await deleteDoc(doc(db, 'riwayatPo', po.id)); notify('Arsip PO dihapus.'); }
    catch (err) { notify('Gagal: ' + (err.code || err.message), 'error'); }
  };

  /* ============================== AKSI: PENGGUNA ============================== */
  const handleAddUser = async (e) => {
    e.preventDefault();
    if (!isAdmin) return notify('Hanya Admin!', 'error');
    const username = formUser.username.trim().toLowerCase();
    if (!/^[a-z0-9._-]{3,}$/.test(username)) return notify('Username minimal 3 karakter (huruf kecil, angka, titik, strip, garis bawah).', 'error');
    if (formUser.password.length < 6) return notify('Password minimal 6 karakter.', 'error');
    setSaving(true);
    try {
      const email = usernameToEmail(username);
      const cred = await createUserWithEmailAndPassword(secondaryAuth, email, formUser.password);
      await setDoc(doc(db, 'users', cred.user.uid), {
        nama: formUser.nama.trim(), username, email, role: formUser.role, site: 'Bontang', status: 'Aktif',
        createdAt: serverTimestamp(), createdBy: profile.username
      });
      await signOut(secondaryAuth);
      setShowModalUser(false); setFormUser({ nama: '', username: '', password: '', role: 'Editor' });
      notify('Pengguna berhasil ditambahkan.');
    } catch (err) {
      console.error(err);
      notify(err.code === 'auth/email-already-in-use' ? 'Username sudah dipakai.' : 'Gagal menambah pengguna: ' + (err.code || err.message), 'error');
    }
    setSaving(false);
  };

  const updateUserField = async (u, patch) => {
    if (!isAdmin) return;
    if (u.id === profile.uid) return notify('Anda tidak dapat mengubah akun Anda sendiri.', 'error');
    try { await updateDoc(doc(db, 'users', u.id), patch); notify('Pengguna diperbarui.'); }
    catch (err) { notify('Gagal: ' + (err.code || err.message), 'error'); }
  };

  const handleDeleteUser = async (u) => {
    if (!isAdmin) return;
    if (u.id === profile.uid) return notify('Anda tidak dapat menghapus akun Anda sendiri.', 'error');
    if (!window.confirm(`Cabut akses pengguna ${u.username}?`)) return;
    try { await deleteDoc(doc(db, 'users', u.id)); notify('Akses pengguna dicabut.'); }
    catch (err) { notify('Gagal: ' + (err.code || err.message), 'error'); }
  };

  const changePassword = async (e) => {
    e.preventDefault();
    if (pwForm.baru.length < 6) return notify('Password baru minimal 6 karakter.', 'error');
    if (pwForm.baru !== pwForm.ulang) return notify('Konfirmasi password tidak sama.', 'error');
    try {
      const cred = EmailAuthProvider.credential(auth.currentUser.email, pwForm.lama);
      await reauthenticateWithCredential(auth.currentUser, cred);
      await updatePassword(auth.currentUser, pwForm.baru);
      setPwForm({ lama: '', baru: '', ulang: '' });
      notify('Password berhasil diganti.');
    } catch (err) {
      notify(err.code === 'auth/invalid-credential' || err.code === 'auth/wrong-password' ? 'Password lama salah.' : 'Gagal: ' + (err.code || err.message), 'error');
    }
  };

  /* ============================== IMPOR / EKSPOR EXCEL ============================== */
  const sheetRows = (wb, name, headerRowSkip = 0, mustHave = 'Kode') => {
    const ws = wb.Sheets[name];
    if (!ws) return [];
    let rows = XLSX.utils.sheet_to_json(ws, headerRowSkip ? { range: headerRowSkip } : {});
    if (headerRowSkip && rows.length && !(mustHave in rows[0])) rows = XLSX.utils.sheet_to_json(ws);
    return rows;
  };

  const importWorkbook = async (wb) => {
    if (!isAdmin) return notify('Hanya Admin yang dapat mengimpor.', 'error');
    setImporting(true); setImportMsg('Membaca file...');
    let stage = 'membaca file';
    try {
      const matRows = sheetRows(wb, 'MATERIAL');
      const masukRows = sheetRows(wb, 'Barang Masuk');
      const keluarRows = sheetRows(wb, 'Barang Keluar', 1);

      // 1) Material: hanya tambah yang belum ada (stok yang sudah ada TIDAK ditimpa)
      const existingIds = new Set(dataMaterial.map((m) => m.id));
      const str = (v) => (v === undefined || v === null ? '' : String(v).trim());
      const num = (v, d = 0) => { const n = Number(v); return Number.isFinite(n) ? n : d; };
      const newMat = matRows.filter((r) => r.Kode !== undefined && String(r.Kode).trim() !== '' && !existingIds.has(safeId(r.Kode)))
        .map((r) => {
          const kode = String(r.Kode).trim();
          return { id: safeId(kode), data: {
            Kode: kode, 'Nama Barang': str(r['Nama Barang']), 'Nomor Barang': str(r['Nomor Barang']),
            'Part Unit': str(r['Part Unit']) || 'Ea', 'Base Unit of Measure': str(r['Base Unit of Measure']) || 'Ea',
            Qty: num(r.Qty), Lokasi: str(r.Lokasi) || 'Rak A1', MinQty: r.MinQty !== undefined ? num(r.MinQty, 5) : 5,
            createdAt: serverTimestamp(), createdBy: profile.username
          } };
        });
      stage = 'master barang';
      let done = 0;
      for (const part of chunk(newMat, 400)) {
        const batch = writeBatch(db);
        part.forEach((m) => batch.set(doc(db, 'material', m.id), m.data));
        await batch.commit();
        done += part.length;
        setImportMsg(`Mengimpor master barang... ${done}/${newMat.length}`);
      }

      // 2) Barang masuk/keluar: hanya jika koleksi masih kosong (hindari data ganda)
      let masukCount = 0; let keluarCount = 0;
      if (dataMasuk.length === 0) {
        stage = 'barang masuk';
        const rows = masukRows.filter((r) => r.Kode !== undefined).map((r) => clean({
          Tanggal: toISODate(r.Tanggal), Unit: r.Unit !== undefined ? String(r.Unit) : '', Kode: String(r.Kode).trim(),
          Description: r.Description !== undefined ? String(r.Description) : '', Merk: r.Merk !== undefined ? String(r.Merk) : '',
          Qty: Number(r.Qty) || 0, Satuan: r['Qty.1'] || r.Satuan || '', 'No. PO': r['No. PO'] !== undefined ? String(r['No. PO']) : '',
          createdAt: serverTimestamp(), createdBy: profile.username
        }));
        for (const part of chunk(rows, 400)) {
          const batch = writeBatch(db);
          part.forEach((r) => batch.set(doc(collection(db, 'barangMasuk')), r));
          await batch.commit(); masukCount += part.length;
          setImportMsg(`Mengimpor barang masuk... ${masukCount}/${rows.length}`);
        }
      }
      if (dataKeluar.length === 0) {
        stage = 'barang keluar';
        const rows = keluarRows.filter((r) => r.Kode !== undefined).map((r) => clean({
          'Doc.Date': toISODate(r['Doc.Date']), Unit: r.Unit !== undefined ? String(r.Unit) : '', Kode: String(r.Kode).trim(),
          'Nama Barang': r['Nama Barang'] !== undefined ? String(r['Nama Barang']) : '', Mekanik: r.Mekanik !== undefined ? String(r.Mekanik) : '',
          Qty: Number(r.Qty) || 0, Keterangan: r.Keterangan !== undefined ? String(r.Keterangan) : '',
          createdAt: serverTimestamp(), createdBy: profile.username
        }));
        for (const part of chunk(rows, 400)) {
          const batch = writeBatch(db);
          part.forEach((r) => batch.set(doc(collection(db, 'barangKeluar')), r));
          await batch.commit(); keluarCount += part.length;
          setImportMsg(`Mengimpor barang keluar... ${keluarCount}/${rows.length}`);
        }
      }
      setImportMsg(`Selesai. Master barang baru: ${newMat.length}, barang masuk: ${masukCount}, barang keluar: ${keluarCount}.`);
      notify('Impor selesai.');
    } catch (err) { console.error(err); const m = `Gagal impor tahap "${stage}": ${err.code || ''} ${err.message || ''}`; setImportMsg(m); notify(m, 'error'); }
    setImporting(false);
  };

  const importFromFile = async (file) => {
    if (!file) return;
    const buf = await file.arrayBuffer();
    importWorkbook(XLSX.read(buf, { type: 'array' }));
  };
  const importFromPublic = async () => {
    try {
      const res = await fetch('/BACKLOG_LOGISTIK_2026_update (1).xlsx');
      if (!res.ok) throw new Error('File tidak ditemukan di folder public');
      importWorkbook(XLSX.read(await res.arrayBuffer(), { type: 'array' }));
    } catch (err) { notify('Gagal: ' + err.message, 'error'); }
  };

  const recalcStock = async () => {
    if (!isAdmin) return notify('Hanya Admin.', 'error');
    const changes = dataMaterial.map((m) => {
      const t = totalsByKode.get(String(m.Kode || '').toLowerCase()) || { in: 0, out: 0 };
      return { m, qty: t.in - t.out };
    }).filter((x) => x.qty !== Number(x.m.Qty || 0));
    if (changes.length === 0) return notify('Stok sudah sesuai dengan transaksi.');
    const neg = changes.filter((x) => x.qty < 0).length;
    if (!window.confirm(`Hitung ulang stok ${changes.length} barang menjadi (total masuk - total keluar)?${neg ? `\n${neg} barang akan bernilai minus karena barang keluar lebih besar dari barang masuk.` : ''}\nStok yang ada saat ini akan diganti. Disarankan backup dulu (menu Laporan).`)) return;
    setImporting(true); setImportMsg('Menghitung ulang stok...');
    try {
      let done = 0;
      for (const part of chunk(changes, 400)) {
        const batch = writeBatch(db);
        part.forEach((x) => batch.update(doc(db, 'material', x.m.id), { Qty: x.qty }));
        await batch.commit();
        done += part.length;
        setImportMsg(`Menghitung ulang stok... ${done}/${changes.length}`);
      }
      setImportMsg(`Selesai. ${changes.length} barang diperbarui.`);
      notify('Stok berhasil dihitung ulang.');
    } catch (err) {
      console.error(err);
      const m = err.code === 'resource-exhausted' ? 'Kuota tulis harian Firebase habis. Coba lagi setelah kuota reset (sekitar pukul 15.00 WIB).' : 'Gagal: ' + (err.code || err.message);
      setImportMsg(m); notify(m, 'error');
    }
    setImporting(false);
  };

  const downloadSheets = (filename, sheets) => {
    const wb = XLSX.utils.book_new();
    Object.entries(sheets).forEach(([name, rows]) => XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows.map(stripMeta)), name));
    XLSX.writeFile(wb, filename);
  };
  const exportAll = () => downloadSheets('BACKLOG_LOGISTIK_REP_UPDATE.xlsx', { MATERIAL: dataMaterial, 'Barang Masuk': sortedMasuk, 'Barang Keluar': sortedKeluar });

  /* ============================== STYLE ============================== */
  const card = darkMode ? 'bg-slate-800 border-slate-700' : 'bg-white border-slate-200';
  const thead = darkMode ? 'bg-slate-900 text-slate-400 border-slate-700' : 'bg-slate-50 text-slate-600 border-slate-200';
  const divide = darkMode ? 'divide-slate-700' : 'divide-slate-100';
  const hoverRow = darkMode ? 'hover:bg-slate-700' : 'hover:bg-slate-50';
  const inputClass = 'w-full p-2 border border-slate-300 rounded-lg bg-white text-slate-900 placeholder-slate-400 focus:outline-none focus:border-blue-500 font-medium';
  const selectClass = 'px-3 py-2 text-sm border border-slate-300 rounded-lg bg-white text-slate-900 font-medium';
  const searchClass = 'w-full pl-9 pr-4 py-2 text-sm border border-slate-300 rounded-lg bg-white text-slate-900 focus:outline-none focus:border-blue-500 font-medium';

  const PeriodFilter = () => (
    <div className={`p-4 rounded-xl shadow-sm border flex flex-wrap gap-3 items-center ${card}`}>
      <div className="text-xs font-semibold flex items-center gap-2"><Calendar size={16} className="text-yellow-500" /> Filter Periode:</div>
      <select value={selectedMonth} onChange={(e) => { setSelectedMonth(e.target.value); setPage(1); }} className={selectClass}>
        <option value="">Semua Bulan</option>
        {['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'].map((n, i) => <option key={i} value={i + 1}>{n}</option>)}
      </select>
      <select value={selectedYear} onChange={(e) => { setSelectedYear(e.target.value); setPage(1); }} className={selectClass}>
        <option value="">Semua Tahun</option>
        {Array.from({ length: 6 }, (_, i) => 2024 + i).map((y) => <option key={y} value={y}>{y}</option>)}
      </select>
    </div>
  );

  const SearchBox = ({ placeholder }) => (
    <div className={`p-4 rounded-xl shadow-sm border flex gap-3 items-center ${card}`}>
      <div className="relative flex-1">
        <Search className="absolute left-3 top-2.5 text-slate-400" size={16} />
        <input type="text" placeholder={placeholder} className={searchClass} value={pencarian} onChange={(e) => { setPencarian(e.target.value); setPage(1); }} />
      </div>
    </div>
  );

  /* ============================== RENDER AWAL (AUTH) ============================== */
  if (authLoading) {
    return <div className="min-h-screen flex items-center justify-center bg-slate-900 text-slate-300 text-sm animate-pulse">Memuat...</div>;
  }
  if (!profile) return <LoginScreen onLogin={handleLogin} error={authError} busy={authBusy} />;

  const menus = [
    { name: 'Dashboard', icon: LayoutDashboard },
    { name: 'Master Barang', icon: Box },
    { name: 'Stok Masuk', icon: ArrowDownToLine },
    { name: 'Stok Keluar', icon: ArrowUpFromLine },
    { name: 'Form & Cetak PO', icon: ShoppingCart },
    { name: 'Arsip Database PO', icon: History },
    { name: 'Stok Opname', icon: ClipboardList },
    { name: 'Kartu Stok', icon: BookOpen },
    { name: 'Rekap Stok', icon: BarChart3 },
    { name: 'Laporan', icon: FileText },
    ...(isAdmin ? [{ name: 'Pengguna', icon: Users }] : []),
    { name: 'Pengaturan', icon: Settings },
  ];

  const emptyHint = dataMaterial.length === 0 && !loadingCloud;

  /* ============================== RENDER ============================== */
  return (
    <div className={`app-root flex h-screen font-sans ${darkMode ? 'bg-slate-900 text-white' : 'bg-slate-50 text-slate-800'}`}>
      <style>{`
        .po-input { background: transparent; color: ${INK}; outline: none; width: 100%; }
        .po-input:focus { background: #fffbe6; }
        @page { size: A4 portrait; margin: 0; }
        .app-main, .app-main > div, .overflow-x-auto { scrollbar-width: auto; scrollbar-color: #94a3b8 rgba(148,163,184,0.2); }
        .app-main::-webkit-scrollbar, .overflow-x-auto::-webkit-scrollbar { width: 14px; height: 14px; }
        .app-main::-webkit-scrollbar-thumb, .overflow-x-auto::-webkit-scrollbar-thumb { background: #94a3b8; border-radius: 8px; border: 3px solid transparent; background-clip: content-box; }
        .app-main::-webkit-scrollbar-track, .overflow-x-auto::-webkit-scrollbar-track { background: rgba(148,163,184,0.2); }
        @media print {
          html { font-size: 16px !important; }
          html, body, #root, .app-root, .app-main, .app-content { overflow: visible !important; height: auto !important; display: block !important; background: #fff !important; }
          body * { visibility: hidden !important; }
          #po-sheet, #po-sheet * { visibility: visible !important; }
          #po-sheet { position: absolute !important; left: 0; top: 0; width: 210mm !important; height: 148.5mm !important; max-width: none !important; margin: 0 !important; border: none !important; box-shadow: none !important; padding: 8mm 10mm !important; box-sizing: border-box !important; overflow: hidden !important; background: #fff !important; }
          .no-print { display: none !important; }
          .po-input:focus { background: transparent; }
          * { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
        }
      `}</style>

      {/* TOAST */}
      {toast && (
        <div className={`fixed top-4 right-4 z-[60] px-4 py-3 rounded-lg shadow-xl text-xs font-semibold max-w-sm no-print ${toast.type === 'error' ? 'bg-rose-600 text-white' : 'bg-emerald-600 text-white'}`}>
          {toast.msg}
        </div>
      )}

      {/* SIDEBAR */}
      <div className={`no-print w-72 shrink-0 flex flex-col border-r ${darkMode ? 'bg-slate-950 border-slate-800 text-slate-300' : 'bg-slate-900 border-slate-800 text-slate-300'}`}>
        <div className="p-4 border-b border-slate-800 flex items-center gap-3 bg-slate-950">
          <div className="w-10 h-10 flex items-center justify-center">
            <img src="/logo_PT_REP_transparan.png" alt="Logo" className="w-full h-full object-contain" />
          </div>
          <div className="overflow-hidden">
            <h2 className="text-xs font-bold text-white uppercase tracking-wider">PT. Ryan Eka Pratama</h2>
            <p className="text-[10px] text-yellow-400 font-semibold">Site Bontang - Logistic</p>
          </div>
        </div>

        <nav className="flex-1 p-3 space-y-1 overflow-y-auto text-sm">
          {menus.map((menu) => {
            const Icon = menu.icon;
            const isActive = activeTab === menu.name;
            const showBadge = menu.name === 'Form & Cetak PO' && poNeededItems.length > 0;
            return (
              <button key={menu.name} onClick={() => goTab(menu.name)}
                className={`w-full flex items-center justify-between px-3.5 py-2.5 rounded-lg font-medium transition-all ${isActive ? 'bg-yellow-500 text-slate-950 font-bold shadow' : 'hover:bg-slate-800 text-slate-300'}`}>
                <div className="flex items-center gap-3"><Icon size={18} /> {menu.name}</div>
                {showBadge && <span className="bg-rose-600 text-white text-[10px] font-bold px-2 py-0.5 rounded-full">{poNeededItems.length}</span>}
              </button>
            );
          })}
        </nav>

        <div className="p-3 border-t border-slate-800 space-y-2 text-sm">
          <button onClick={() => setDarkMode(!darkMode)} className="w-full flex items-center gap-3 px-3 py-2 rounded-lg hover:bg-slate-800 text-slate-400 hover:text-white">
            <Moon size={16} /> {darkMode ? 'Mode Terang' : 'Mode Gelap'}
          </button>
          <div className="flex items-center justify-between px-3 py-1 text-slate-400 text-xs">
            <span>Ukuran Huruf</span>
            <div className="flex gap-1">
              {[[16, 'text-xs'], [18, 'text-sm'], [20, 'text-lg']].map(([sz, cls]) => (
                <button key={sz} onClick={() => setFontSize(sz)} title={`Huruf ${sz}px`}
                  className={`w-8 h-8 rounded-lg font-bold flex items-center justify-center ${cls} ${fontSize === sz ? 'bg-yellow-500 text-slate-950' : 'bg-slate-800 text-slate-300 hover:bg-slate-700'}`}>A</button>
              ))}
            </div>
          </div>
          <div className="pt-2 border-t border-slate-800 flex items-center gap-3 px-3 py-2">
            <div className="w-8 h-8 rounded-full bg-emerald-600 flex items-center font-bold text-white justify-center text-xs">{(profile.nama || '?').charAt(0)}</div>
            <div className="overflow-hidden flex-1">
              <p className="text-xs font-semibold text-white truncate">{profile.nama}</p>
              <p className="text-[10px] text-yellow-400 truncate">Role: {profile.role}</p>
            </div>
            <button onClick={handleLogout} title="Keluar" className="p-1.5 rounded-lg hover:bg-slate-800 text-slate-400 hover:text-rose-400"><LogOut size={16} /></button>
          </div>
        </div>
      </div>

      {/* KONTEN UTAMA */}
      <div className="app-content flex-1 min-w-0 flex flex-col overflow-hidden">
        <header className={`no-print p-4 border-b flex flex-wrap gap-2 justify-between items-center shadow-sm ${darkMode ? 'bg-slate-900 border-slate-800' : 'bg-white border-slate-200'}`}>
          <div>
            <h1 className="text-xl font-bold flex items-center gap-2">
              {activeTab}
              {loadingCloud && <span className="text-xs text-yellow-500 font-normal animate-pulse">☁️ Sinkronisasi Cloud...</span>}
            </h1>
            <p className="text-xs text-slate-400">Logistic Management System - PT. Ryan Eka Pratama Site Bontang</p>
          </div>
          <div className="flex items-center gap-2">
            {canEdit ? (
              <>
                <button onClick={openNewMaterial} className="bg-emerald-600 hover:bg-emerald-700 text-white px-3 py-1.5 text-xs font-semibold rounded-lg flex items-center gap-1"><Plus size={14} /> Produk Baru</button>
                <button onClick={() => { setFormMasuk({ ...emptyMasuk, tanggal: todayISO() }); setShowModalMasuk(true); }} className="bg-blue-600 hover:bg-blue-700 text-white px-3 py-1.5 text-xs font-semibold rounded-lg flex items-center gap-1"><Barcode size={14} /> Stok Masuk</button>
                <button onClick={() => { setFormKeluar({ ...emptyKeluar, tanggal: todayISO() }); setShowModalKeluar(true); }} className="bg-amber-600 hover:bg-amber-700 text-white px-3.5 py-1.5 text-xs font-semibold rounded-lg flex items-center gap-1"><Barcode size={14} /> Stok Keluar</button>
              </>
            ) : (<span className="bg-slate-200 text-slate-500 px-3 py-1.5 rounded-lg text-xs font-medium border">👁️ Mode Viewer</span>)}
          </div>
        </header>

        <main className="app-main flex-1 p-6 overflow-auto">

          {emptyHint && activeTab !== 'Pengaturan' && activeTab !== 'Pengguna' && (
            <div className="no-print mb-4 p-4 rounded-xl border border-yellow-300 bg-yellow-50 text-yellow-900 text-xs font-medium">
              Database masih kosong. {isAdmin ? <>Buka menu <button className="underline font-bold" onClick={() => goTab('Pengaturan')}>Pengaturan</button> lalu impor data Excel ke Firestore.</> : 'Minta Admin untuk mengimpor data awal.'}
            </div>
          )}

          {/* ===================== DASHBOARD ===================== */}
          {activeTab === 'Dashboard' && (
            <div className="space-y-6">
              {PeriodFilter()}
              <div className="grid grid-cols-1 md:grid-cols-4 gap-4">
                <div onClick={() => goTab('Master Barang')} className={`p-5 rounded-xl border shadow-sm flex items-center justify-between cursor-pointer hover:border-blue-400 ${card}`}>
                  <div><p className="text-xs font-semibold text-slate-400 uppercase">Total Master Barang</p><h3 className="text-2xl font-bold mt-1 text-blue-600">{dataMaterial.length} Item</h3></div>
                  <div className="p-3 bg-blue-50 text-blue-600 rounded-xl"><Box size={24} /></div>
                </div>
                <div onClick={() => goTab('Stok Masuk')} className={`p-5 rounded-xl border shadow-sm flex items-center justify-between cursor-pointer hover:border-emerald-400 ${card}`}>
                  <div><p className="text-xs font-semibold text-slate-400 uppercase">Barang Masuk</p><h3 className="text-2xl font-bold mt-1 text-emerald-600">{filteredMasuk.length} Trx</h3></div>
                  <div className="p-3 bg-emerald-50 text-emerald-600 rounded-xl"><TrendingUp size={24} /></div>
                </div>
                <div onClick={() => goTab('Stok Keluar')} className={`p-5 rounded-xl border shadow-sm flex items-center justify-between cursor-pointer hover:border-rose-400 ${card}`}>
                  <div><p className="text-xs font-semibold text-slate-400 uppercase">Barang Keluar</p><h3 className="text-2xl font-bold mt-1 text-rose-600">{filteredKeluar.length} Trx</h3></div>
                  <div className="p-3 bg-rose-50 text-rose-600 rounded-xl"><TrendingDown size={24} /></div>
                </div>
                <div onClick={() => goTab('Form & Cetak PO')} className={`p-5 rounded-xl border shadow-sm flex items-center justify-between cursor-pointer hover:border-yellow-500 ${card}`}>
                  <div><p className="text-xs font-semibold text-slate-400 uppercase">Wajib Order (PO)</p><h3 className="text-2xl font-bold mt-1 text-amber-500">{poNeededItems.length} Item</h3></div>
                  <div className="p-3 bg-amber-50 text-amber-600 rounded-xl"><ShoppingCart size={24} /></div>
                </div>
              </div>

              <div className={`rounded-xl shadow-sm border overflow-x-auto ${card}`}>
                <div className="p-4 border-b flex items-center justify-between bg-amber-50">
                  <h3 className="font-bold text-xs uppercase tracking-wider text-amber-700 flex items-center gap-2"><AlertTriangle size={16} /> Perhatian: Material Mencapai Stok Minimum & Wajib Segera Di-Order (PO)</h3>
                  <button onClick={() => goTab('Form & Cetak PO')} className="text-xs bg-yellow-500 text-slate-950 font-bold px-3 py-1 rounded-lg">Buka Form PO ➔</button>
                </div>
                <table className="w-full min-w-[860px] text-left text-sm">
                  <thead className={`border-b font-semibold ${thead}`}>
                    <tr><th className="p-3.5">KODE</th><th className="p-3.5">NAMA BARANG</th><th className="p-3.5">LOKASI RAK</th><th className="p-3.5 text-center">STOK SAAT INI</th><th className="p-3.5 text-center">BATAS MINIMUM</th><th className="p-3.5 text-center">STATUS</th></tr>
                  </thead>
                  <tbody className={`divide-y ${divide}`}>
                    {poNeededItems.length === 0 ? (
                      <tr><td colSpan="6" className="p-6 text-center text-slate-400">Semua stok barang di atas batas minimum.</td></tr>
                    ) : poNeededItems.slice(0, 10).map((item) => {
                      const qn = Number(item.Qty || 0);
                      return (
                        <tr key={item.id} className={hoverRow}>
                          <td className="p-3.5 font-mono font-bold">{item.Kode || '-'}</td>
                          <td className="p-3.5 font-semibold">{item['Nama Barang'] || '-'}</td>
                          <td className="p-3.5"><span className="bg-yellow-100 text-yellow-800 px-2 py-0.5 rounded font-bold text-[10px]">{item.Lokasi || 'Rak A1'}</span></td>
                          <td className="p-3.5 text-center font-bold text-rose-600 text-sm">{qn}</td>
                          <td className="p-3.5 text-center font-bold text-blue-600">{item.MinQty ?? 5}</td>
                          <td className="p-3.5 text-center">{qn === 0 ? <span className="bg-rose-100 text-rose-800 px-2 py-1 rounded font-bold text-[10px]">Habis</span> : <span className="bg-amber-100 text-amber-800 px-2 py-1 rounded font-bold text-[10px]">Di Bawah Min</span>}</td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
                {poNeededItems.length > 10 && <div className="p-3 text-center text-[11px] text-slate-400">+ {poNeededItems.length - 10} barang lainnya di Form PO</div>}
              </div>

              <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
                {[
                  { title: 'Barang Masuk Terbaru', rows: filteredMasuk.slice(0, 5), color: 'text-emerald-600', sign: '+', tab: 'Stok Masuk', date: (r) => r.Tanggal, name: (r) => r.Description },
                  { title: 'Barang Keluar Terbaru', rows: filteredKeluar.slice(0, 5), color: 'text-rose-600', sign: '-', tab: 'Stok Keluar', date: (r) => r['Doc.Date'], name: (r) => r['Nama Barang'] },
                ].map((b) => (
                  <div key={b.title} className={`rounded-xl shadow-sm border overflow-x-auto ${card}`}>
                    <div className="p-3.5 border-b flex justify-between items-center"><h3 className="font-bold text-xs uppercase tracking-wider">{b.title}</h3><button onClick={() => goTab(b.tab)} className="text-[11px] text-blue-500 font-semibold">Lihat semua ➔</button></div>
                    <table className="w-full text-xs"><tbody className={`divide-y ${divide}`}>
                      {b.rows.length === 0 ? <tr><td className="p-4 text-center text-slate-400">Belum ada transaksi.</td></tr> : b.rows.map((r) => (
                        <tr key={r.id}><td className="p-3 whitespace-nowrap">{fmtDate(b.date(r))}</td><td className="p-3 font-mono">{r.Kode}</td><td className="p-3 truncate max-w-[180px]">{b.name(r)}</td><td className={`p-3 text-right font-bold ${b.color}`}>{b.sign}{r.Qty}</td></tr>
                      ))}
                    </tbody></table>
                  </div>
                ))}
              </div>
            </div>
          )}

          {/* ===================== MASTER BARANG ===================== */}
          {activeTab === 'Master Barang' && (
            <div className="space-y-4">
              {SearchBox({ placeholder: "Cari nama barang, kode barcode, part number, atau lokasi..." })}
              <div className={`rounded-xl shadow-sm border overflow-x-auto ${card}`}>
                <table className="w-full min-w-[860px] text-left text-sm">
                  <thead className={`border-b font-semibold ${thead}`}>
                    <tr><th className="p-3.5">KODE / BARCODE</th><th className="p-3.5">NAMA BARANG</th><th className="p-3.5">PART NO.</th><th className="p-3.5">LOKASI</th><th className="p-3.5 text-center">MIN</th><th className="p-3.5 text-center">STOK</th><th className="p-3.5 text-center">STATUS PO</th>{canEdit && <th className="p-3.5 text-center">AKSI</th>}</tr>
                  </thead>
                  <tbody className={`divide-y ${divide}`}>
                    {filteredMaterial.length === 0 ? (
                      <tr><td colSpan="8" className="p-8 text-center text-slate-400">Tidak ada data ditemukan.</td></tr>
                    ) : paged(filteredMaterial).map((item) => {
                      const qn = Number(item.Qty || 0); const minQ = Number(item.MinQty ?? 5);
                      return (
                        <tr key={item.id} className={hoverRow}>
                          <td className="p-3.5 font-mono font-bold text-blue-600">{item.Kode || '-'}</td>
                          <td className="p-3.5 font-semibold">{item['Nama Barang'] || '-'}</td>
                          <td className="p-3.5">{item['Nomor Barang'] ?? '-'}</td>
                          <td className="p-3.5"><span className="bg-slate-100 text-slate-800 px-2 py-0.5 rounded font-bold text-[10px]">📍 {item.Lokasi || 'Rak A1'}</span></td>
                          <td className="p-3.5 text-center font-bold">{minQ}</td>
                          <td className="p-3.5 text-center font-bold text-emerald-600 text-sm">{qn}</td>
                          <td className="p-3.5 text-center">{qn <= minQ ? <span className="bg-rose-100 text-rose-800 px-2.5 py-1 rounded-full font-bold text-[10px]">⚠️ Perlu Order PO</span> : <span className="bg-emerald-100 text-emerald-800 px-2.5 py-1 rounded-full font-bold text-[10px]">Aman</span>}</td>
                          {canEdit && (
                            <td className="p-3.5 text-center whitespace-nowrap">
                              <button onClick={() => openEditMaterial(item)} title="Edit" className="p-1.5 text-blue-500 hover:bg-blue-50 rounded"><Pencil size={14} /></button>
                              {isAdmin && <button onClick={() => handleDeleteMaterial(item)} title="Hapus" className="p-1.5 text-rose-500 hover:bg-rose-50 rounded"><Trash2 size={14} /></button>}
                            </td>
                          )}
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
                <Pager page={page} setPage={setPage} total={filteredMaterial.length} dark={darkMode} />
              </div>
            </div>
          )}

          {/* ===================== STOK MASUK ===================== */}
          {activeTab === 'Stok Masuk' && (
            <div className="space-y-4">
              {PeriodFilter()}
              {SearchBox({ placeholder: "Cari kode, deskripsi, unit, atau No. PO..." })}
              <div className={`rounded-xl shadow-sm border overflow-x-auto ${card}`}>
                <table className="w-full min-w-[860px] text-left text-sm">
                  <thead className={`border-b font-semibold ${thead}`}>
                    <tr><th className="p-3.5">TANGGAL</th><th className="p-3.5">UNIT</th><th className="p-3.5">KODE</th><th className="p-3.5">DESKRIPSI</th><th className="p-3.5">NO. PO</th><th className="p-3.5">QTY</th>{isAdmin && <th className="p-3.5 text-center">AKSI</th>}</tr>
                  </thead>
                  <tbody className={`divide-y ${divide}`}>
                    {filteredMasuk.length === 0 ? <tr><td colSpan="7" className="p-8 text-center text-slate-400">Tidak ada transaksi.</td></tr> : paged(filteredMasuk).map((item) => (
                      <tr key={item.id} className={hoverRow}>
                        <td className="p-3.5 whitespace-nowrap">{fmtDate(item.Tanggal)}</td><td className="p-3.5">{item.Unit}</td>
                        <td className="p-3.5 font-mono">{item.Kode}</td><td className="p-3.5">{item.Description}</td><td className="p-3.5">{item['No. PO']}</td>
                        <td className="p-3.5 font-bold text-emerald-600">+{item.Qty} <span className="text-slate-400 font-normal">{item.Satuan || item['Qty.1'] || ''}</span></td>
                        {isAdmin && <td className="p-3.5 text-center"><button onClick={() => handleDeleteTrx('masuk', item)} className="p-1.5 text-rose-500 hover:bg-rose-50 rounded"><Trash2 size={14} /></button></td>}
                      </tr>
                    ))}
                  </tbody>
                </table>
                <Pager page={page} setPage={setPage} total={filteredMasuk.length} dark={darkMode} />
              </div>
            </div>
          )}

          {/* ===================== STOK KELUAR ===================== */}
          {activeTab === 'Stok Keluar' && (
            <div className="space-y-4">
              {PeriodFilter()}
              {SearchBox({ placeholder: "Cari kode, nama barang, unit, atau mekanik..." })}
              <div className={`rounded-xl shadow-sm border overflow-x-auto ${card}`}>
                <table className="w-full min-w-[860px] text-left text-sm">
                  <thead className={`border-b font-semibold ${thead}`}>
                    <tr><th className="p-3.5">TANGGAL</th><th className="p-3.5">UNIT</th><th className="p-3.5">KODE</th><th className="p-3.5">NAMA BARANG</th><th className="p-3.5">MEKANIK</th><th className="p-3.5">KETERANGAN</th><th className="p-3.5">QTY</th>{isAdmin && <th className="p-3.5 text-center">AKSI</th>}</tr>
                  </thead>
                  <tbody className={`divide-y ${divide}`}>
                    {filteredKeluar.length === 0 ? <tr><td colSpan="8" className="p-8 text-center text-slate-400">Tidak ada transaksi.</td></tr> : paged(filteredKeluar).map((item) => (
                      <tr key={item.id} className={hoverRow}>
                        <td className="p-3.5 whitespace-nowrap">{fmtDate(item['Doc.Date'])}</td><td className="p-3.5">{item.Unit}</td>
                        <td className="p-3.5 font-mono">{item.Kode}</td><td className="p-3.5">{item['Nama Barang']}</td><td className="p-3.5">{item.Mekanik}</td><td className="p-3.5">{item.Keterangan}</td>
                        <td className="p-3.5 font-bold text-rose-600">-{item.Qty}</td>
                        {isAdmin && <td className="p-3.5 text-center"><button onClick={() => handleDeleteTrx('keluar', item)} className="p-1.5 text-rose-500 hover:bg-rose-50 rounded"><Trash2 size={14} /></button></td>}
                      </tr>
                    ))}
                  </tbody>
                </table>
                <Pager page={page} setPage={setPage} total={filteredKeluar.length} dark={darkMode} />
              </div>
            </div>
          )}

          {/* ===================== FORM & CETAK PO (putih) ===================== */}
          {activeTab === 'Form & Cetak PO' && (
            <div className="space-y-4">
              <div className="no-print p-4 rounded-xl border shadow-sm bg-yellow-50 border-yellow-300 flex flex-wrap gap-2 items-center justify-between">
                <div>
                  <h3 className="font-bold text-sm text-slate-900 flex items-center gap-2"><ShoppingCart size={18} className="text-yellow-600" /> Form Purchase Order {editingPoId && <span className="text-[10px] bg-blue-600 text-white px-2 py-0.5 rounded-full">Mengedit arsip</span>}</h3>
                  <p className="text-xs text-slate-600">Isi langsung pada lembar di bawah, simpan ke arsip, lalu cetak. Hasil cetak berlatar putih, ukuran setengah A4 portrait (210 × 148,5 mm), maksimal ±8 baris per lembar.</p>
                </div>
                <div className="flex flex-wrap gap-2">
                  <button onClick={fillPoFromNeeded} className="bg-amber-500 hover:bg-amber-600 text-slate-950 px-3 py-2 rounded-lg text-xs font-bold flex items-center gap-1.5"><AlertTriangle size={14} /> Isi dari Barang Wajib Order ({poNeededItems.length})</button>
                  <button onClick={addPoRow} className="bg-slate-700 hover:bg-slate-800 text-white px-3 py-2 rounded-lg text-xs font-semibold flex items-center gap-1.5"><Plus size={14} /> Baris</button>
                  <button onClick={() => setPoItems((items) => (items.length > 1 ? items.slice(0, -1) : items))} className="bg-rose-600 hover:bg-rose-700 text-white px-3 py-2 rounded-lg text-xs font-semibold flex items-center gap-1.5"><Trash2 size={14} /> Hapus Baris Terakhir</button>
                  {canEdit && <button onClick={savePo} disabled={saving} className="bg-emerald-600 hover:bg-emerald-700 disabled:opacity-60 text-white px-3 py-2 rounded-lg text-xs font-semibold flex items-center gap-1.5"><Save size={14} /> {editingPoId ? 'Perbarui Arsip' : 'Simpan ke Arsip'}</button>}
                  <button onClick={() => window.print()} className="bg-blue-600 hover:bg-blue-700 text-white px-3 py-2 rounded-lg text-xs font-semibold flex items-center gap-1.5"><Printer size={14} /> Cetak</button>
                  <button onClick={resetPoForm} className="bg-white border border-slate-300 text-slate-700 px-3 py-2 rounded-lg text-xs font-semibold flex items-center gap-1.5"><FilePlus2 size={14} /> Form Baru</button>
                </div>
              </div>

              <div id="po-sheet" className="bg-white border border-slate-300 shadow-md mx-auto" style={{ width: '210mm', minHeight: '148.5mm', padding: '8mm 10mm', boxSizing: 'border-box', color: INK, fontFamily: 'Arial, Helvetica, sans-serif' }}>
                <div className="flex items-center gap-4">
                  <img src="/logo_PT_REP_transparan.png" alt="" className="w-12 h-12 object-contain" />
                  <div>
                    <div className="text-lg font-extrabold tracking-wide" style={{ fontFamily: 'Georgia, "Times New Roman", serif' }}>PT. RYAN EKA PRATAMA</div>
                    <div className="text-xs font-bold" style={{ letterSpacing: '0.4em' }}>{PO_KANTOR}</div>
                  </div>
                </div>

                <div className="relative flex items-center justify-center mt-2 mb-2">
                  <div className="text-lg font-extrabold tracking-wide">PURCHASE ORDER</div>
                  <div className="absolute right-0 flex items-center gap-2 text-base" style={{ fontFamily: 'Georgia, serif' }}>
                    <span>No.</span>
                    <input className="po-input font-bold text-red-600 w-36 border-b border-dotted" style={{ color: '#c0392b' }} value={poHeader.noPo} onChange={(e) => setPoHeader({ ...poHeader, noPo: e.target.value })} />
                  </div>
                </div>

                <div className="space-y-0.5 text-xs font-semibold mb-2 max-w-md">
                  {[['EQUIPMENT', 'equipment'], ['KODE ALAT', 'kodeAlat'], ['LOKASI', 'lokasi']].map(([label, key]) => (
                    <div key={key} className="flex items-center gap-2"><span className="w-24 shrink-0 whitespace-nowrap">{label}</span><span>:</span><input className="po-input border-b border-dotted" value={poHeader[key]} onChange={(e) => setPoHeader({ ...poHeader, [key]: e.target.value })} /></div>
                  ))}
                  <div className="flex items-center gap-2"><span className="w-24 shrink-0 whitespace-nowrap">TANGGAL</span><span>:</span><input type="date" className="po-input border-b border-dotted" value={poHeader.tanggal} onChange={(e) => setPoHeader({ ...poHeader, tanggal: e.target.value })} /></div>
                </div>

                <table className="w-full text-xs" style={{ borderCollapse: 'collapse', border: `2px solid ${INK}` }}>
                  <thead>
                    <tr style={{ borderBottom: `2px solid ${INK}` }}>
                      {[['NO', '6%'], ['URAIAN', '32%'], ['PART NUMBER', '18%'], ['QTY', '8%'], ['UNIT PRICE', '14%'], ['KETERANGAN', '22%']].map(([h, w]) => (
                        <th key={h} className="py-1 font-bold" style={{ width: w, borderRight: `1.5px solid ${INK}` }}>{h}</th>
                      ))}
                      <th className="no-print w-12" style={{ border: 'none' }}></th>
                    </tr>
                  </thead>
                  <tbody>
                    {[...poItems, ...Array.from({ length: Math.max(0, 8 - poItems.length) }, (_, i) => ({ filler: true, id: 'f' + i }))].map((it, idx) => (
                      <tr key={it.id} style={{ height: 22 }}>
                        <td className="text-center font-semibold" style={{ borderRight: `1.5px solid ${INK}` }}>{it.filler ? '' : idx + 1}</td>
                        {['uraian', 'partNumber', 'qty', 'unitPrice', 'keterangan'].map((f) => (
                          <td key={f} style={{ borderRight: `1.5px solid ${INK}` }} className="px-1">
                            {it.filler ? null : <input className={`po-input px-1 py-0.5 ${f === 'qty' || f === 'unitPrice' ? 'text-center' : ''}`} value={it[f]} onChange={(e) => updatePoItem(it.id, f, e.target.value)} />}
                          </td>
                        ))}
                        <td className="no-print text-center" style={{ border: 'none' }}>{!it.filler && <button onClick={() => removePoRow(it.id)} title="Hapus baris ini" className="p-1.5 rounded-md text-rose-600 bg-rose-50 hover:bg-rose-600 hover:text-white"><Trash2 size={15} /></button>}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>

                <div className="flex justify-between mt-2 text-xs font-semibold px-8">
                  <div className="text-center w-56">
                    <div>DIKETAHUI OLEH,</div>
                    <div style={{ height: 56 }} />
                  </div>
                  <div className="text-center w-56">
                    <div>DIBUAT OLEH,</div>
                    <div style={{ height: 56 }} />
                  </div>
                </div>
              </div>
            </div>
          )}

          {/* ===================== ARSIP PO ===================== */}
          {activeTab === 'Arsip Database PO' && (() => {
            const list = [...riwayatPo]
              .filter((p) => !q || has(p.noPo) || has(p.equipment) || has(p.kodeAlat) || (p.items || []).some((i) => has(i.uraian) || has(i.partNumber)))
              .sort((a, b) => poNumberValue(b.noPo) - poNumberValue(a.noPo));
            return (
              <div className="space-y-4">
                {SearchBox({ placeholder: "Cari no. PO, equipment, kode alat, atau nama barang..." })}
                <div className={`rounded-xl shadow-sm border overflow-x-auto ${card}`}>
                  <table className="w-full min-w-[860px] text-left text-sm">
                    <thead className={`border-b font-semibold ${thead}`}>
                      <tr><th className="p-3.5">NO. PO</th><th className="p-3.5">TANGGAL</th><th className="p-3.5">EQUIPMENT</th><th className="p-3.5">KODE ALAT</th><th className="p-3.5 text-center">JML ITEM</th><th className="p-3.5">DIBUAT OLEH</th><th className="p-3.5 text-center">AKSI</th></tr>
                    </thead>
                    <tbody className={`divide-y ${divide}`}>
                      {list.length === 0 ? <tr><td colSpan="7" className="p-8 text-center text-slate-400">Belum ada PO tersimpan.</td></tr> : paged(list).map((po) => (
                        <tr key={po.id} className={hoverRow}>
                          <td className="p-3.5 font-bold text-red-600">{po.noPo}</td><td className="p-3.5">{fmtDate(po.tanggal)}</td>
                          <td className="p-3.5">{po.equipment}</td><td className="p-3.5">{po.kodeAlat}</td>
                          <td className="p-3.5 text-center font-bold">{(po.items || []).length}</td><td className="p-3.5">{po.dibuatOleh}</td>
                          <td className="p-3.5 text-center whitespace-nowrap">
                            <button onClick={() => openPoFromArchive(po)} className="px-2.5 py-1 bg-blue-600 text-white rounded-lg font-semibold mr-1 inline-flex items-center gap-1"><Printer size={12} /> Buka / Cetak</button>
                            {isAdmin && <button onClick={() => deletePo(po)} className="p-1.5 text-rose-500 hover:bg-rose-50 rounded"><Trash2 size={14} /></button>}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                  <Pager page={page} setPage={setPage} total={list.length} dark={darkMode} />
                </div>
              </div>
            );
          })()}

          {/* ===================== STOK OPNAME ===================== */}
          {activeTab === 'Stok Opname' && (
            <div className="space-y-4">
              <div className={`p-4 rounded-xl border shadow-sm flex justify-between items-center ${card}`}>
                <div><h3 className="font-bold text-sm">Stock Opname</h3><p className="text-xs text-slate-400">Isi hitungan fisik. Hanya barang yang selisih yang disesuaikan ({opnameChanges.length} barang berubah).</p></div>
                {canEdit && <button onClick={saveOpname} disabled={saving || opnameChanges.length === 0} className="bg-emerald-600 disabled:opacity-50 text-white px-4 py-2 rounded-lg text-xs font-semibold shadow flex items-center gap-1.5"><Save size={14} /> Simpan Opname</button>}
              </div>
              {SearchBox({ placeholder: "Cari kode, nama barang, atau lokasi..." })}
              <div className={`rounded-xl shadow-sm border overflow-x-auto ${card}`}>
                <table className="w-full min-w-[860px] text-left text-sm">
                  <thead className={`border-b font-semibold ${thead}`}>
                    <tr><th className="p-3.5">KODE</th><th className="p-3.5">NAMA BARANG</th><th className="p-3.5">LOKASI</th><th className="p-3.5 text-center">SISTEM</th><th className="p-3.5 text-center">FISIK</th><th className="p-3.5 text-center">SELISIH</th></tr>
                  </thead>
                  <tbody className={`divide-y ${divide}`}>
                    {paged(filteredMaterial).map((item) => {
                      const sys = Number(item.Qty) || 0;
                      const raw = physicalStocks[item.id];
                      const phys = raw === undefined || raw === '' ? sys : Number(raw);
                      const diff = phys - sys;
                      return (
                        <tr key={item.id}>
                          <td className="p-3.5 font-mono font-bold">{item.Kode}</td><td className="p-3.5">{item['Nama Barang']}</td>
                          <td className="p-3.5 font-semibold text-blue-600">{item.Lokasi || 'Rak A1'}</td>
                          <td className="p-3.5 text-center font-bold text-blue-600">{sys}</td>
                          <td className="p-3.5 text-center"><input type="number" min="0" disabled={!canEdit} className="w-20 p-1 text-center border rounded font-bold bg-white text-slate-900" value={raw === undefined ? sys : raw} onChange={(e) => setPhysicalStocks({ ...physicalStocks, [item.id]: e.target.value })} /></td>
                          <td className="p-3.5 text-center">{diff === 0 ? <span className="bg-emerald-100 text-emerald-800 px-2 py-0.5 rounded font-bold text-[10px]">Sesuai</span> : <span className="bg-rose-100 text-rose-800 px-2 py-0.5 rounded font-bold text-[10px]">{diff > 0 ? '+' : ''}{diff}</span>}</td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
                <Pager page={page} setPage={setPage} total={filteredMaterial.length} dark={darkMode} />
              </div>
            </div>
          )}

          {/* ===================== KARTU STOK ===================== */}
          {activeTab === 'Kartu Stok' && (
            <div className="space-y-4">
              <div className={`p-4 rounded-xl border shadow-sm flex flex-wrap items-center gap-3 ${card}`}>
                <span className="text-xs font-semibold">Pilih Barang:</span>
                <input type="text" placeholder="Cari kode / nama / part no..." value={kartuSearch} onChange={(e) => setKartuSearch(e.target.value)} className={`${selectClass} w-64`} />
                <select value={selectedItemCode} onChange={(e) => setSelectedItemCode(e.target.value)} className={`${selectClass} flex-1 min-w-[260px]`}>
                  <option value="">-- Pilih Material ({kartuOptions.length} ditampilkan) --</option>
                  {kartuOptions.map((m) => <option key={m.id} value={m.Kode}>{m.Kode} - {m['Nama Barang']} (Lokasi: {m.Lokasi || 'Rak A1'})</option>)}
                </select>
              </div>
              {kartu ? (
                <div className={`rounded-xl shadow-sm border overflow-x-auto ${card}`}>
                  <div className="p-4 border-b flex flex-wrap gap-6 text-xs">
                    <div><p className="text-slate-400 uppercase font-semibold">Kode</p><p className="font-mono font-bold text-sm">{kartu.mat.Kode}</p></div>
                    <div><p className="text-slate-400 uppercase font-semibold">Nama Barang</p><p className="font-bold text-sm">{kartu.mat['Nama Barang']}</p></div>
                    <div><p className="text-slate-400 uppercase font-semibold">Lokasi</p><p className="font-bold text-sm">{kartu.mat.Lokasi || 'Rak A1'}</p></div>
                    <div><p className="text-slate-400 uppercase font-semibold">Stok Saat Ini</p><p className="font-bold text-sm text-emerald-600">{kartu.mat.Qty}</p></div>
                  </div>
                  <table className="w-full min-w-[860px] text-left text-sm">
                    <thead className={`border-b font-semibold ${thead}`}><tr><th className="p-3.5">TANGGAL</th><th className="p-3.5">KETERANGAN</th><th className="p-3.5 text-center">MASUK</th><th className="p-3.5 text-center">KELUAR</th><th className="p-3.5 text-center">SALDO</th></tr></thead>
                    <tbody className={`divide-y ${divide}`}>
                      <tr className="bg-slate-50 text-slate-700"><td className="p-3.5">-</td><td className="p-3.5 font-semibold">Saldo Awal</td><td /><td /><td className="p-3.5 text-center font-bold">{kartu.saldoAwal}</td></tr>
                      {kartu.rows.map((r, i) => (
                        <tr key={i}><td className="p-3.5 whitespace-nowrap">{fmtDate(r.date)}</td><td className="p-3.5">{r.ket}</td>
                          <td className="p-3.5 text-center font-bold text-emerald-600">{r.masuk || ''}</td><td className="p-3.5 text-center font-bold text-rose-600">{r.keluar || ''}</td><td className="p-3.5 text-center font-bold">{r.saldo}</td></tr>
                      ))}
                      <tr className="font-bold"><td className="p-3.5" colSpan="2">TOTAL</td><td className="p-3.5 text-center text-emerald-600">{kartu.sumIn}</td><td className="p-3.5 text-center text-rose-600">{kartu.sumOut}</td><td className="p-3.5 text-center">{kartu.mat.Qty}</td></tr>
                    </tbody>
                  </table>
                </div>
              ) : <p className="text-xs text-slate-400 text-center py-10">Pilih material untuk melihat kartu stok.</p>}
            </div>
          )}

          {/* ===================== REKAP STOK ===================== */}
          {activeTab === 'Rekap Stok' && (
            <div className="space-y-4">
              {SearchBox({ placeholder: "Cari kode, nama barang, part number, atau lokasi..." })}
              <div className={`rounded-xl shadow-sm border overflow-x-auto ${card}`}>
                <div className="p-4 border-b flex justify-between items-center"><h3 className="font-bold text-sm">Rekap Persediaan</h3><button onClick={exportAll} className="bg-emerald-600 text-white px-3 py-1.5 rounded-lg text-xs font-semibold flex items-center gap-1.5"><Download size={14} /> Excel</button></div>
                <table className="w-full min-w-[860px] text-left text-sm">
                  <thead className={`border-b font-semibold ${thead}`}>
                    <tr><th className="p-3.5">KODE</th><th className="p-3.5">NAMA</th><th className="p-3.5">PART NO.</th><th className="p-3.5">LOKASI</th><th className="p-3.5 text-center">TOTAL MASUK</th><th className="p-3.5 text-center">TOTAL KELUAR</th><th className="p-3.5 text-center">STOK</th></tr>
                  </thead>
                  <tbody className={`divide-y ${divide}`}>
                    {paged(filteredMaterial).map((item) => {
                      const t = totalsByKode.get(String(item.Kode || '').toLowerCase()) || { in: 0, out: 0 };
                      return (
                        <tr key={item.id} className={hoverRow}>
                          <td className="p-3.5 font-mono">{item.Kode}</td><td className="p-3.5">{item['Nama Barang']}</td><td className="p-3.5">{item['Nomor Barang'] ?? '-'}</td>
                          <td className="p-3.5 font-bold text-blue-600">{item.Lokasi || 'Rak A1'}</td>
                          <td className="p-3.5 text-center text-emerald-600 font-semibold">{t.in}</td><td className="p-3.5 text-center text-rose-600 font-semibold">{t.out}</td>
                          <td className="p-3.5 text-center font-bold text-emerald-600">{item.Qty}</td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
                <Pager page={page} setPage={setPage} total={filteredMaterial.length} dark={darkMode} />
              </div>
            </div>
          )}

          {/* ===================== LAPORAN ===================== */}
          {activeTab === 'Laporan' && (
            <div className="space-y-4">
              {PeriodFilter()}
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                {[
                  { t: 'Laporan Stok Seluruh Barang', d: `${dataMaterial.length} item`, f: () => downloadSheets('Laporan_Stok.xlsx', { MATERIAL: dataMaterial }) },
                  { t: 'Laporan Barang Masuk (sesuai periode)', d: `${sortedMasuk.filter((r) => matchesPeriod(r.Tanggal)).length} transaksi`, f: () => downloadSheets('Laporan_Barang_Masuk.xlsx', { 'Barang Masuk': sortedMasuk.filter((r) => matchesPeriod(r.Tanggal)) }) },
                  { t: 'Laporan Barang Keluar (sesuai periode)', d: `${sortedKeluar.filter((r) => matchesPeriod(r['Doc.Date'])).length} transaksi`, f: () => downloadSheets('Laporan_Barang_Keluar.xlsx', { 'Barang Keluar': sortedKeluar.filter((r) => matchesPeriod(r['Doc.Date'])) }) },
                  { t: 'Daftar Barang Wajib Order (PO)', d: `${poNeededItems.length} item`, f: () => downloadSheets('Daftar_Wajib_Order.xlsx', { 'Wajib Order': poNeededItems }) },
                  { t: 'Cadangan Seluruh Data (Backup)', d: 'Material, masuk, keluar', f: exportAll },
                ].map((r) => (
                  <div key={r.t} className={`p-4 rounded-xl border shadow-sm flex justify-between items-center ${card}`}>
                    <div><h3 className="font-bold text-sm">{r.t}</h3><p className="text-xs text-slate-400">{r.d}</p></div>
                    <button onClick={r.f} className="bg-emerald-600 text-white px-3 py-2 rounded-lg text-xs font-semibold flex items-center gap-1.5"><Download size={14} /> Unduh</button>
                  </div>
                ))}
              </div>
            </div>
          )}

          {/* ===================== PENGGUNA ===================== */}
          {activeTab === 'Pengguna' && isAdmin && (
            <div className={`rounded-xl shadow-sm border overflow-x-auto ${card}`}>
              <div className="p-4 border-b flex justify-between items-center"><h3 className="font-bold text-sm">Manajemen Pengguna</h3><button onClick={() => setShowModalUser(true)} className="bg-blue-600 text-white px-3 py-1.5 rounded-lg text-xs font-semibold flex items-center gap-1"><Plus size={14} /> Tambah Pengguna</button></div>
              <table className="w-full min-w-[860px] text-left text-sm">
                <thead className={`border-b font-semibold ${thead}`}><tr><th className="p-3.5">NAMA</th><th className="p-3.5">USERNAME</th><th className="p-3.5">ROLE</th><th className="p-3.5">STATUS</th><th className="p-3.5 text-center">AKSI</th></tr></thead>
                <tbody className={`divide-y ${divide}`}>
                  {daftarPengguna.map((u) => (
                    <tr key={u.id}>
                      <td className="p-3.5 font-semibold">{u.nama}{u.id === profile.uid && <span className="ml-2 text-[10px] bg-yellow-100 text-yellow-800 px-1.5 py-0.5 rounded">Anda</span>}</td>
                      <td className="p-3.5 font-mono">{u.username}</td>
                      <td className="p-3.5">
                        <select disabled={u.id === profile.uid} value={u.role} onChange={(e) => updateUserField(u, { role: e.target.value })} className="p-1 border rounded bg-white text-slate-900 text-xs">
                          <option>Admin</option><option>Editor</option><option>Viewer</option>
                        </select>
                      </td>
                      <td className="p-3.5">
                        <button disabled={u.id === profile.uid} onClick={() => updateUserField(u, { status: u.status === 'Aktif' ? 'Nonaktif' : 'Aktif' })}
                          className={`px-2.5 py-1 rounded-full font-bold text-[10px] ${u.status === 'Aktif' ? 'bg-emerald-100 text-emerald-800' : 'bg-slate-200 text-slate-600'}`}>{u.status}</button>
                      </td>
                      <td className="p-3.5 text-center">{u.id !== profile.uid && <button onClick={() => handleDeleteUser(u)} title="Cabut akses" className="p-1.5 text-rose-500 hover:bg-rose-50 rounded"><Trash2 size={14} /></button>}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
              <p className="p-3 text-[11px] text-slate-400 border-t">Catatan: "Cabut akses" menghapus profil pengguna sehingga tidak bisa masuk lagi. Akun loginnya dapat dihapus total dari Firebase Console → Authentication.</p>
            </div>
          )}

          {/* ===================== PENGATURAN ===================== */}
          {activeTab === 'Pengaturan' && (
            <div className="space-y-4 max-w-2xl">
              <div className={`p-6 rounded-xl border shadow-sm space-y-3 ${card}`}>
                <h3 className="font-bold text-sm border-b pb-2">Informasi Akun & Sistem</h3>
                <div className="text-xs grid grid-cols-[120px_1fr] gap-y-1.5">
                  <span className="text-slate-400">Perusahaan</span><span className="font-semibold">PT. Ryan Eka Pratama - Site Bontang</span>
                  <span className="text-slate-400">Nama</span><span className="font-semibold">{profile.nama}</span>
                  <span className="text-slate-400">Username</span><span className="font-mono font-semibold">{profile.username}</span>
                  <span className="text-slate-400">Role</span><span className="font-semibold">{profile.role}</span>
                </div>
              </div>

              <form onSubmit={changePassword} className={`p-6 rounded-xl border shadow-sm space-y-3 ${card}`}>
                <h3 className="font-bold text-sm border-b pb-2">Ganti Password</h3>
                <input type="password" required placeholder="Password lama" className={inputClass} value={pwForm.lama} onChange={(e) => setPwForm({ ...pwForm, lama: e.target.value })} autoComplete="current-password" />
                <input type="password" required placeholder="Password baru (min. 6 karakter)" className={inputClass} value={pwForm.baru} onChange={(e) => setPwForm({ ...pwForm, baru: e.target.value })} autoComplete="new-password" />
                <input type="password" required placeholder="Ulangi password baru" className={inputClass} value={pwForm.ulang} onChange={(e) => setPwForm({ ...pwForm, ulang: e.target.value })} autoComplete="new-password" />
                <button className="bg-blue-600 text-white px-4 py-2 rounded-lg text-xs font-semibold">Simpan Password</button>
              </form>

              {isAdmin && (
                <div className={`p-6 rounded-xl border shadow-sm space-y-3 ${card}`}>
                  <h3 className="font-bold text-sm border-b pb-2 flex items-center gap-2"><Upload size={16} /> Impor Data Excel ke Firestore</h3>
                  <p className="text-xs text-slate-400">Membaca sheet <b>MATERIAL</b>, <b>Barang Masuk</b>, dan <b>Barang Keluar</b>. Master barang yang sudah ada tidak ditimpa. Barang masuk/keluar hanya diimpor bila datanya masih kosong (mencegah data ganda).</p>
                  <div className="flex flex-wrap gap-2 items-center">
                    <label className={`px-3 py-2 rounded-lg text-xs font-semibold cursor-pointer bg-emerald-600 text-white ${importing ? 'opacity-50 pointer-events-none' : ''}`}>
                      Pilih File .xlsx
                      <input type="file" accept=".xlsx,.xls" className="hidden" onChange={(e) => { importFromFile(e.target.files?.[0]); e.target.value = ''; }} />
                    </label>
                    <button disabled={importing} onClick={importFromPublic} className="px-3 py-2 rounded-lg text-xs font-semibold border border-slate-300 bg-white text-slate-700 disabled:opacity-50">Impor dari file bawaan (public)</button>
                  </div>
                  {importMsg && <p className="text-xs font-semibold text-blue-600">{importMsg}</p>}
                </div>
              )}

              {isAdmin && (
                <div className={`p-6 rounded-xl border shadow-sm space-y-3 ${card}`}>
                  <h3 className="font-bold text-sm border-b pb-2 flex items-center gap-2"><ClipboardList size={16} /> Hitung Ulang Stok dari Transaksi</h3>
                  <p className="text-xs text-slate-400">Mengisi stok setiap barang = total barang masuk − total barang keluar yang tercatat. Gunakan bila stok hasil impor Excel tampil 0. Stok yang sudah ada akan diganti.</p>
                  <button disabled={importing} onClick={recalcStock} className="px-3 py-2 rounded-lg text-xs font-semibold bg-amber-500 text-slate-950 disabled:opacity-50">Hitung Ulang Stok</button>
                  {importMsg && <p className="text-xs font-semibold text-blue-600">{importMsg}</p>}
                </div>
              )}
            </div>
          )}

        </main>
      </div>

      {/* ===================== MODAL: PRODUK ===================== */}
      {showModalMaterial && (
        <Modal dark={darkMode} title={editingMaterialId ? 'Edit Data Barang' : 'Tambah Produk Baru & Stok Minimum'} onClose={() => setShowModalMaterial(false)}>
          <form onSubmit={handleSaveMaterial} className="space-y-3 text-xs">
            <div><label className="block font-semibold mb-1">Kode / Barcode Barang</label><input type="text" required readOnly={!!editingMaterialId} className={`${inputClass} ${editingMaterialId ? 'bg-slate-100' : ''}`} value={formMaterial.kode} onChange={(e) => setFormMaterial({ ...formMaterial, kode: e.target.value })} onKeyDown={preventEnterSubmit} autoFocus={!editingMaterialId} placeholder="Scan atau ketik kode barcode..." /></div>
            <div><label className="block font-semibold mb-1">Nama Barang</label><input type="text" required className={inputClass} value={formMaterial.nama} onChange={(e) => setFormMaterial({ ...formMaterial, nama: e.target.value })} placeholder="FILTER OLI CAT" /></div>
            <div><label className="block font-semibold mb-1">Nomor Part</label><input type="text" className={inputClass} value={formMaterial.nomor} onChange={(e) => setFormMaterial({ ...formMaterial, nomor: e.target.value })} placeholder="1R-1808" /></div>
            <div className="grid grid-cols-2 gap-2">
              <div><label className="block font-semibold mb-1">Lokasi Rak / Yard</label><input type="text" required className={inputClass} value={formMaterial.lokasi} onChange={(e) => setFormMaterial({ ...formMaterial, lokasi: e.target.value })} /></div>
              <div><label className="block font-semibold mb-1">Batas Stok Minimum</label><input type="number" min="0" required className={inputClass} value={formMaterial.minQty} onChange={(e) => setFormMaterial({ ...formMaterial, minQty: e.target.value })} /></div>
            </div>
            <div className="grid grid-cols-2 gap-2">
              <div><label className="block font-semibold mb-1">Part Unit</label><input type="text" className={inputClass} value={formMaterial.unit} onChange={(e) => setFormMaterial({ ...formMaterial, unit: e.target.value })} /></div>
              <div><label className="block font-semibold mb-1">{editingMaterialId ? 'Stok Saat Ini' : 'Qty Awal'}</label><input type="number" min="0" required readOnly={!!editingMaterialId} className={`${inputClass} ${editingMaterialId ? 'bg-slate-100' : ''}`} value={formMaterial.qty} onChange={(e) => setFormMaterial({ ...formMaterial, qty: e.target.value })} /></div>
            </div>
            {editingMaterialId && <p className="text-[11px] text-slate-400">Stok diubah lewat Stok Masuk, Stok Keluar, atau Stok Opname agar riwayat tetap akurat.</p>}
            <div className="flex justify-end gap-2 mt-4 pt-2 border-t"><button type="button" onClick={() => setShowModalMaterial(false)} className="px-4 py-2 border rounded-lg">Batal</button><button disabled={saving} className="px-4 py-2 bg-emerald-600 disabled:opacity-60 text-white font-semibold rounded-lg">{saving ? 'Menyimpan...' : 'Simpan'}</button></div>
          </form>
        </Modal>
      )}

      {/* ===================== MODAL: STOK MASUK ===================== */}
      {showModalMasuk && (
        <Modal dark={darkMode} title="Scan / Input Stok Masuk" icon={<Barcode size={18} className="text-blue-500" />} onClose={() => setShowModalMasuk(false)}>
          <form onSubmit={handleAddMasuk} className="space-y-3 text-xs">
            <div>
              <label className="block font-semibold mb-1 text-blue-600">Scan Barcode / Ketik Kode Barang:</label>
              <input type="text" autoFocus required value={formMasuk.kode} onChange={(e) => lookupMasuk(e.target.value)} onKeyDown={preventEnterSubmit}
                className="w-full p-2.5 border-2 border-blue-500 rounded-lg bg-blue-50 text-slate-900 font-bold placeholder-blue-300 focus:outline-none" placeholder="Arahkan scanner ke barcode atau ketik kode..." />
              {formMasuk.kode && (materialByKode.get(formMasuk.kode.trim().toLowerCase())
                ? <p className="mt-1 text-emerald-600 font-semibold">✔ Ditemukan. Stok saat ini: {materialByKode.get(formMasuk.kode.trim().toLowerCase()).Qty}</p>
                : <p className="mt-1 text-rose-500 font-semibold">Kode belum ada di Master Barang.</p>)}
            </div>
            <div className="grid grid-cols-2 gap-2">
              <div><label className="block font-semibold mb-1">Tanggal</label><input type="date" required className={inputClass} value={formMasuk.tanggal} onChange={(e) => setFormMasuk({ ...formMasuk, tanggal: e.target.value })} /></div>
              <div><label className="block font-semibold mb-1">Unit</label><input type="text" className={inputClass} value={formMasuk.unit} onChange={(e) => setFormMasuk({ ...formMasuk, unit: e.target.value })} placeholder="REP 110" /></div>
            </div>
            <div><label className="block font-semibold mb-1">Deskripsi</label><input type="text" required className={inputClass} value={formMasuk.description} onChange={(e) => setFormMasuk({ ...formMasuk, description: e.target.value })} /></div>
            <div><label className="block font-semibold mb-1">Merk / Part No.</label><input type="text" className={inputClass} value={formMasuk.merk} onChange={(e) => setFormMasuk({ ...formMasuk, merk: e.target.value })} /></div>
            <div className="grid grid-cols-3 gap-2">
              <div><label className="block font-semibold mb-1">Qty</label><input type="number" min="1" required className={inputClass} value={formMasuk.qty} onChange={(e) => setFormMasuk({ ...formMasuk, qty: e.target.value })} /></div>
              <div><label className="block font-semibold mb-1">Satuan</label><input type="text" required className={inputClass} value={formMasuk.satuan} onChange={(e) => setFormMasuk({ ...formMasuk, satuan: e.target.value })} /></div>
              <div><label className="block font-semibold mb-1">No. PO</label><input type="text" className={inputClass} value={formMasuk.noPo} onChange={(e) => setFormMasuk({ ...formMasuk, noPo: e.target.value })} /></div>
            </div>
            <div className="flex justify-end gap-2 mt-4 pt-2 border-t"><button type="button" onClick={() => setShowModalMasuk(false)} className="px-4 py-2 border rounded-lg">Batal</button><button disabled={saving} className="px-4 py-2 bg-blue-600 disabled:opacity-60 text-white font-semibold rounded-lg">{saving ? 'Menyimpan...' : 'Simpan Stok Masuk'}</button></div>
          </form>
        </Modal>
      )}

      {/* ===================== MODAL: STOK KELUAR ===================== */}
      {showModalKeluar && (
        <Modal dark={darkMode} title="Scan / Input Stok Keluar" icon={<Barcode size={18} className="text-amber-500" />} onClose={() => setShowModalKeluar(false)}>
          <form onSubmit={handleAddKeluar} className="space-y-3 text-xs">
            <div>
              <label className="block font-semibold mb-1 text-amber-600">Scan Barcode / Ketik Kode Barang:</label>
              <input type="text" autoFocus required value={formKeluar.kode} onChange={(e) => lookupKeluar(e.target.value)} onKeyDown={preventEnterSubmit}
                className="w-full p-2.5 border-2 border-amber-500 rounded-lg bg-amber-50 text-slate-900 font-bold placeholder-amber-300 focus:outline-none" placeholder="Arahkan scanner ke barcode atau ketik kode..." />
              {formKeluar.kode && (materialByKode.get(formKeluar.kode.trim().toLowerCase())
                ? <p className="mt-1 text-emerald-600 font-semibold">✔ Ditemukan. Stok tersedia: {materialByKode.get(formKeluar.kode.trim().toLowerCase()).Qty}</p>
                : <p className="mt-1 text-rose-500 font-semibold">Kode belum ada di Master Barang.</p>)}
            </div>
            <div className="grid grid-cols-2 gap-2">
              <div><label className="block font-semibold mb-1">Tanggal</label><input type="date" required className={inputClass} value={formKeluar.tanggal} onChange={(e) => setFormKeluar({ ...formKeluar, tanggal: e.target.value })} /></div>
              <div><label className="block font-semibold mb-1">Unit</label><input type="text" required className={inputClass} value={formKeluar.unit} onChange={(e) => setFormKeluar({ ...formKeluar, unit: e.target.value })} placeholder="REP 110" /></div>
            </div>
            <div><label className="block font-semibold mb-1">Nama Barang</label><input type="text" required className={inputClass} value={formKeluar.namaBarang} onChange={(e) => setFormKeluar({ ...formKeluar, namaBarang: e.target.value })} /></div>
            <div><label className="block font-semibold mb-1">Mekanik</label><input type="text" className={inputClass} value={formKeluar.mekanik} onChange={(e) => setFormKeluar({ ...formKeluar, mekanik: e.target.value })} placeholder="Endang" /></div>
            <div className="grid grid-cols-2 gap-2">
              <div><label className="block font-semibold mb-1">Qty Keluar</label><input type="number" min="1" required className={inputClass} value={formKeluar.qty} onChange={(e) => setFormKeluar({ ...formKeluar, qty: e.target.value })} /></div>
              <div><label className="block font-semibold mb-1">Keterangan</label><input type="text" className={inputClass} value={formKeluar.keterangan} onChange={(e) => setFormKeluar({ ...formKeluar, keterangan: e.target.value })} placeholder="Ganti / Rusak" /></div>
            </div>
            <div className="flex justify-end gap-2 mt-4 pt-2 border-t"><button type="button" onClick={() => setShowModalKeluar(false)} className="px-4 py-2 border rounded-lg">Batal</button><button disabled={saving} className="px-4 py-2 bg-amber-600 disabled:opacity-60 text-white font-semibold rounded-lg">{saving ? 'Menyimpan...' : 'Simpan Stok Keluar'}</button></div>
          </form>
        </Modal>
      )}

      {/* ===================== MODAL: PENGGUNA ===================== */}
      {showModalUser && (
        <Modal dark={darkMode} title="Tambah Pengguna" onClose={() => setShowModalUser(false)}>
          <form onSubmit={handleAddUser} className="space-y-3 text-xs">
            <div><label className="block font-semibold mb-1">Nama Lengkap</label><input type="text" required className={inputClass} value={formUser.nama} onChange={(e) => setFormUser({ ...formUser, nama: e.target.value })} /></div>
            <div><label className="block font-semibold mb-1">Username (untuk login)</label><input type="text" required className={inputClass} value={formUser.username} onChange={(e) => setFormUser({ ...formUser, username: e.target.value })} placeholder="endang_mech" autoComplete="off" /></div>
            <div><label className="block font-semibold mb-1">Password Awal</label><input type="password" required minLength={6} className={inputClass} value={formUser.password} onChange={(e) => setFormUser({ ...formUser, password: e.target.value })} placeholder="Minimal 6 karakter" autoComplete="new-password" /></div>
            <div><label className="block font-semibold mb-1">Role</label><select className={selectClass} value={formUser.role} onChange={(e) => setFormUser({ ...formUser, role: e.target.value })}><option value="Admin">Admin (akses penuh)</option><option value="Editor">Editor (input data)</option><option value="Viewer">Viewer (hanya lihat)</option></select></div>
            <div className="flex justify-end gap-2 mt-4 pt-2 border-t"><button type="button" onClick={() => setShowModalUser(false)} className="px-4 py-2 border rounded-lg">Batal</button><button disabled={saving} className="px-4 py-2 bg-blue-600 disabled:opacity-60 text-white font-semibold rounded-lg">{saving ? 'Membuat...' : 'Simpan'}</button></div>
          </form>
        </Modal>
      )}
    </div>
  );
}