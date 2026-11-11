import React, { useState, useEffect } from 'react';
import * as XLSX from 'xlsx';
import { 
  Search, Plus, LayoutDashboard, Box, ArrowDownToLine, ArrowUpFromLine, 
  FileText, ClipboardList, BookOpen, BarChart3, Users, Settings, 
  Moon, TrendingUp, TrendingDown, Calendar, X, Download, UserPlus, Sliders, Trash2
} from 'lucide-react';

// Import Firebase SDK
import { initializeApp } from "firebase/app";
import { getAnalytics } from "firebase/analytics";
import { getFirestore, collection, getDocs, addDoc } from "firebase/firestore";

// Konfigurasi Firebase Anda
const firebaseConfig = {
  apiKey: "AIzaSyAuetnYQDspE1IFbOcFttgSJMq3FOvmft0",
  authDomain: "logistik-rep-bontang.firebaseapp.com",
  projectId: "logistik-rep-bontang",
  storageBucket: "logistik-rep-bontang.firebasestorage.app",
  messagingSenderId: "26293450683",
  appId: "1:26293450683:web:8d6458695a511b75ea369e",
  measurementId: "G-E72GBTQ59N"
};

// Inisialisasi Firebase
const app = initializeApp(firebaseConfig);
const analytics = getAnalytics(app);
const db = getFirestore(app);

export default function App() {
  const [activeTab, setActiveTab] = useState('Dashboard');
  const [pencarian, setPencarian] = useState("");
  const [selectedMonth, setSelectedMonth] = useState("");
  const [selectedYear, setSelectedYear] = useState("");
  const [darkMode, setDarkMode] = useState(false);
  const [loadingCloud, setLoadingCloud] = useState(false);

  // State Simulasi Login Pengguna
  const [currentUser, setCurrentUser] = useState({
    nama: "Administrator PT. REP",
    username: "admin_rep",
    role: "Admin"
  });

  // State Data (Cloud Firestore / Excel)
  const [dataMaterial, setDataMaterial] = useState([]);
  const [dataMasuk, setDataMasuk] = useState([]);
  const [dataKeluar, setDataKeluar] = useState([]);

  // State Pilihan Kartu Stok & Stock Opname
  const [selectedItemCode, setSelectedItemCode] = useState("");
  const [physicalStocks, setPhysicalStocks] = useState({});

  // State Modul Laporan
  const [jenisLaporan, setJenisLaporan] = useState("stok");

  // State Modul Pengguna
  const [daftarPengguna, setDaftarPengguna] = useState([
    { id: '1', nama: "Administrator PT. REP", username: "admin_rep", role: "Admin", site: "Bontang", status: "Aktif" },
    { id: '2', nama: "Endang (Mekanik Lead)", username: "endang_mech", role: "Editor", site: "Bontang", status: "Aktif" },
    { id: '3', nama: "Petugas Logistik Site", username: "logistik_btg", role: "Viewer", site: "Bontang", status: "Aktif" }
  ]);
  const [showModalUser, setShowModalUser] = useState(false);
  const [formUser, setFormUser] = useState({ nama: '', username: '', role: 'Editor' });

  // State Modal Input Utama
  const [showModalMaterial, setShowModalMaterial] = useState(false);
  const [showModalMasuk, setShowModalMasuk] = useState(false);
  const [showModalKeluar, setShowModalKeluar] = useState(false);

  const [formMaterial, setFormMaterial] = useState({ kode: '', nama: '', nomor: '', unit: 'Ea', qty: 0 });
  const [formMasuk, setFormMasuk] = useState({ tanggal: '', unit: '', kode: '', description: '', merk: '', qty: 0, satuan: 'Pcs', noPo: '' });
  const [formKeluar, setFormKeluar] = useState({ tanggal: '', unit: '', kode: '', namaBarang: '', mekanik: '', qty: 0, keterangan: '' });

  const [pageMaterial, setPageMaterial] = useState(1);
  const [pageMasuk, setPageMasuk] = useState(1);
  const [pageKeluar, setPageKeluar] = useState(1);
  const rowsPerPage = 20;

  // Ambil Data dari Firebase Firestore saat aplikasi dibuka
  useEffect(() => {
    const fetchDataFromCloud = async () => {
      setLoadingCloud(true);
      try {
        const querySnapshotMat = await getDocs(collection(db, "material"));
        const matList = querySnapshotMat.docs.map(doc => ({ id: doc.id, ...doc.data() }));

        const querySnapshotMasuk = await getDocs(collection(db, "barangMasuk"));
        const masukList = querySnapshotMasuk.docs.map(doc => ({ id: doc.id, ...doc.data() }));

        const querySnapshotKeluar = await getDocs(collection(db, "barangKeluar"));
        const keluarList = querySnapshotKeluar.docs.map(doc => ({ id: doc.id, ...doc.data() }));

        if (matList.length === 0) {
          const response = await fetch('/BACKLOG_LOGISTIK_2026_update (1).xlsx');
          const arrayBuffer = await response.arrayBuffer();
          const workbook = XLSX.read(arrayBuffer, { type: 'array' });
          
          let mat = workbook.Sheets['MATERIAL'] ? XLSX.utils.sheet_to_json(workbook.Sheets['MATERIAL']) : [];
          let masuk = workbook.Sheets['Barang Masuk'] ? XLSX.utils.sheet_to_json(workbook.Sheets['Barang Masuk']) : [];
          let keluar = workbook.Sheets['Barang Keluar'] ? XLSX.utils.sheet_to_json(workbook.Sheets['Barang Keluar'], { range: 1 }) : [];

          setDataMaterial(mat);
          setDataMasuk(masuk);
          setDataKeluar(keluar);
        } else {
          setDataMaterial(matList);
          setDataMasuk(masukList);
          setDataKeluar(keluarList);
        }
      } catch (error) {
        console.error("Gagal memuat data dari Firebase:", error);
      }
      setLoadingCloud(false);
    };

    fetchDataFromCloud();
  }, []);

  // Fungsi Tambah Master Barang ke Firebase
  const handleAddMaterial = async (e) => {
    e.preventDefault();
    if (currentUser.role === 'Viewer') {
      alert("Akses ditolak! Akun Viewer hanya dapat melihat data.");
      return;
    }
    const newItem = { 
      Kode: formMaterial.kode, 
      'Nama Barang': formMaterial.nama, 
      'Nomor Barang': formMaterial.nomor, 
      'Part Unit': formMaterial.unit, 
      Qty: Number(formMaterial.qty), 
      'Base Unit of Measure': 'Ea' 
    };

    try {
      await addDoc(collection(db, "material"), newItem);
      setDataMaterial([newItem, ...dataMaterial]);
      setShowModalMaterial(false);
      setFormMaterial({ kode: '', nama: '', nomor: '', unit: 'Ea', qty: 0 });
      alert("Master barang berhasil ditambahkan ke database online!");
    } catch (error) {
      console.error("Error adding material: ", error);
      alert("Gagal menyimpan ke database online.");
    }
  };

  // Fungsi Tambah Stok Masuk ke Firebase
  const handleAddMasuk = async (e) => {
    e.preventDefault();
    if (currentUser.role === 'Viewer') {
      alert("Akses ditolak! Akun Viewer hanya dapat melihat data.");
      return;
    }
    const newItem = { 
      Tanggal: formMasuk.tanggal, 
      Unit: formMasuk.unit, 
      Kode: formMasuk.kode, 
      Description: formMasuk.description, 
      Merk: formMasuk.merk, 
      Qty: Number(formMasuk.qty), 
      'Qty.1': formMasuk.satuan, 
      'No. PO': formMasuk.noPo 
    };

    try {
      await addDoc(collection(db, "barangMasuk"), newItem);
      setDataMasuk([newItem, ...dataMasuk]);
      setShowModalMasuk(false);
      setFormMasuk({ tanggal: '', unit: '', kode: '', description: '', merk: '', qty: 0, satuan: 'Pcs', noPo: '' });
      alert("Stok masuk berhasil dicatat ke database online!");
    } catch (error) {
      console.error("Error adding masuk: ", error);
      alert("Gagal menyimpan ke database online.");
    }
  };

  // Fungsi Tambah Stok Keluar ke Firebase
  const handleAddKeluar = async (e) => {
    e.preventDefault();
    if (currentUser.role === 'Viewer') {
      alert("Akses ditolak! Akun Viewer hanya dapat melihat data.");
      return;
    }
    const newItem = { 
      'Doc.Date': formKeluar.tanggal, 
      Unit: formKeluar.unit, 
      Kode: formKeluar.kode, 
      'Nama Barang': formKeluar.namaBarang, 
      Mekanik: formKeluar.mekanik, 
      Qty: Number(formKeluar.qty), 
      Keterangan: formKeluar.keterangan 
    };

    try {
      await addDoc(collection(db, "barangKeluar"), newItem);
      setDataKeluar([newItem, ...dataKeluar]);
      setShowModalKeluar(false);
      setFormKeluar({ tanggal: '', unit: '', kode: '', namaBarang: '', mekanik: '', qty: 0, keterangan: '' });
      alert("Stok keluar berhasil dicatat ke database online!");
    } catch (error) {
      console.error("Error adding keluar: ", error);
      alert("Gagal menyimpan ke database online.");
    }
  };

  // Fungsi Tambah Pengguna (Hanya Admin)
  const handleAddUser = (e) => {
    e.preventDefault();
    if (currentUser.role !== 'Admin') {
      alert("Hanya Admin yang dapat menambahkan pengguna!");
      return;
    }
    const newUser = {
      id: String(Date.now()),
      nama: formUser.nama,
      username: formUser.username,
      role: formUser.role,
      site: "Bontang",
      status: "Aktif"
    };
    const updatedUsers = [...daftarPengguna, newUser];
    setDaftarPengguna(updatedUsers);
    setShowModalUser(false);
    setFormUser({ nama: '', username: '', role: 'Editor' });
    alert("Pengguna baru berhasil ditambahkan!");
  };

  // Fungsi Hapus Pengguna (Hanya Admin)
  const handleDeleteUser = (id, roleTarget) => {
    if (currentUser.role !== 'Admin') {
      alert("Hanya Admin yang dapat menghapus pengguna!");
      return;
    }
    if (roleTarget === 'Admin' && daftarPengguna.filter(u => u.role === 'Admin').length <= 1) {
      alert("Tidak dapat menghapus satu-satunya Admin yang tersisa!");
      return;
    }
    if (confirm("Apakah Anda yakin ingin menghapus pengguna ini?")) {
      const updatedUsers = daftarPengguna.filter(u => u.id !== id);
      setDaftarPengguna(updatedUsers);
      alert("Pengguna berhasil dihapus.");
    }
  };

  const exportToExcel = () => {
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(dataMaterial), "MATERIAL");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(dataMasuk), "Barang Masuk");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(dataKeluar), "Barang Keluar");
    XLSX.writeFile(wb, "BACKLOG_LOGISTIK_REP_UPDATE.xlsx");
  };

  const parseExcelDate = (serial) => {
    if (!serial) return null;
    let dateInfo;
    if (typeof serial === 'string' && serial.includes('-')) {
      dateInfo = new Date(serial);
    } else {
      const parsed = Number(serial);
      if (isNaN(parsed)) return null;
      const utcDays = Math.floor(parsed - 25569);
      dateInfo = new Date(utcDays * 86400 * 1000);
    }
    if (isNaN(dateInfo.getTime())) return null;
    return dateInfo;
  };

  const formatTanggalExcel = (serial) => {
    const dateInfo = parseExcelDate(serial);
    if (!dateInfo) return '-';
    return dateInfo.toLocaleDateString('id-ID', { day: 'numeric', month: 'short', year: 'numeric' });
  };

  const matchesMonthYear = (serial) => {
    if (!selectedMonth && !selectedYear) return true;
    const dateObj = parseExcelDate(serial);
    if (!dateObj) return false;
    const m = dateObj.getMonth() + 1; 
    const y = dateObj.getFullYear();   
    if (selectedMonth && m !== Number(selectedMonth)) return false;
    if (selectedYear && y !== Number(selectedYear)) return false;
    return true;
  };

  const filteredMaterial = dataMaterial.filter(item => 
    item['Nama Barang']?.toLowerCase().includes(pencarian.toLowerCase()) || 
    item.Kode?.toLowerCase().includes(pencarian.toLowerCase()) ||
    item['Nomor Barang']?.toString().toLowerCase().includes(pencarian.toLowerCase())
  );

  const filteredMasuk = dataMasuk.filter(item => {
    const matchText = item.Description?.toLowerCase().includes(pencarian.toLowerCase()) || item.Kode?.toLowerCase().includes(pencarian.toLowerCase());
    return matchText && matchesMonthYear(item.Tanggal);
  });

  const filteredKeluar = dataKeluar.filter(item => {
    const matchText = item['Nama Barang']?.toLowerCase().includes(pencarian.toLowerCase()) || item.Kode?.toLowerCase().includes(pencarian.toLowerCase());
    return matchText && matchesMonthYear(item['Doc.Date']);
  });

  const paginatedMaterial = filteredMaterial.slice((pageMaterial - 1) * rowsPerPage, pageMaterial * rowsPerPage);
  const paginatedMasuk = filteredMasuk.slice((pageMasuk - 1) * rowsPerPage, pageMasuk * rowsPerPage);
  const paginatedKeluar = filteredKeluar.slice((pageKeluar - 1) * rowsPerPage, pageKeluar * rowsPerPage);

  const totalPageMaterial = Math.ceil(filteredMaterial.length / rowsPerPage);
  const totalPageMasuk = Math.ceil(filteredMasuk.length / rowsPerPage);
  const totalPageKeluar = Math.ceil(filteredKeluar.length / rowsPerPage);

  return (
    <div className={`flex h-screen font-sans ${darkMode ? 'bg-slate-900 text-white' : 'bg-slate-50 text-slate-800'}`}>
      
      {/* SIDEBAR */}
      <div className={`w-72 flex flex-col border-r ${darkMode ? 'bg-slate-950 border-slate-800 text-slate-300' : 'bg-slate-900 border-slate-800 text-slate-300'}`}>
        <div className="p-4 border-b border-slate-800 flex items-center gap-3 bg-slate-950">
          <div className="w-10 h-10 flex items-center justify-center">
            <img src="/logo_PT_REP_transparan.png" alt="Logo PT. Ryan Eka Pratama" className="w-full h-full object-contain filter drop-shadow-[0_2px_4px_rgba(0,0,0,0.5)]" />
          </div>
          <div className="overflow-hidden">
            <h2 className="text-xs font-bold text-white uppercase tracking-wider">PT. Ryan Eka Pratama</h2>
            <p className="text-[10px] text-yellow-400 font-semibold">Site Bontang - Logistic</p>
          </div>
        </div>

        <nav className="flex-1 p-3 space-y-1 overflow-y-auto text-sm">
          {[
            { name: 'Dashboard', icon: LayoutDashboard },
            { name: 'Master Barang', icon: Box },
            { name: 'Stok Masuk', icon: ArrowDownToLine },
            { name: 'Stok Keluar', icon: ArrowUpFromLine },
            { name: 'Stok Opname', icon: ClipboardList },
            { name: 'Kartu Stok', icon: BookOpen },
            { name: 'Rekap Stok', icon: BarChart3 },
            { name: 'Laporan', icon: FileText },
            { name: 'Pengguna', icon: Users },
            { name: 'Pengaturan', icon: Settings },
          ].map((menu) => {
            const Icon = menu.icon;
            const isActive = activeTab === menu.name;
            return (
              <button
                key={menu.name}
                onClick={() => setActiveTab(menu.name)}
                className={`w-full flex items-center gap-3 px-3.5 py-2.5 rounded-lg font-medium transition-all ${
                  isActive ? 'bg-yellow-500 text-slate-950 font-bold shadow' : 'hover:bg-slate-800 text-slate-300'
                }`}
              >
                <Icon size={18} /> {menu.name}
              </button>
            );
          })}
        </nav>

        <div className="p-3 border-t border-slate-800 space-y-2 text-sm">
          <button onClick={() => setDarkMode(!darkMode)} className="w-full flex items-center gap-3 px-3 py-2 rounded-lg hover:bg-slate-800 text-slate-400 hover:text-white">
            <Moon size={16} /> {darkMode ? 'Mode Terang' : 'Mode Gelap'}
          </button>
          
          <div className="px-3 py-2 bg-slate-950 rounded-lg border border-slate-800 text-xs space-y-1">
            <p className="text-[10px] text-slate-400">Simulasi Mode Akun:</p>
            <select 
              className="w-full bg-slate-900 text-white p-1.5 rounded border border-slate-700 text-xs"
              value={currentUser.role}
              onChange={(e) => {
                const r = e.target.value;
                setCurrentUser({ ...currentUser, role: r, nama: r === 'Admin' ? 'Administrator PT. REP' : r === 'Editor' ? 'Staff Editor' : 'User Viewer' });
              }}
            >
              <option value="Admin">Admin (Full Akses)</option>
              <option value="Editor">Editor (Input Data)</option>
              <option value="Viewer">Viewer (Hanya Lihat)</option>
            </select>
          </div>

          <div className="pt-2 border-t border-slate-800 flex items-center gap-3 px-3 py-2">
            <div className="w-8 h-8 rounded-full bg-emerald-600 flex items-center font-bold text-white justify-center text-xs">
              {currentUser.nama.charAt(0)}
            </div>
            <div className="overflow-hidden">
              <p className="text-xs font-semibold text-white truncate">{currentUser.nama}</p>
              <p className="text-[10px] text-yellow-400 truncate">Role: {currentUser.role}</p>
            </div>
          </div>
        </div>
      </div>

      {/* KONTEN UTAMA */}
      <div className="flex-1 flex flex-col overflow-hidden">
        
        <header className={`p-4 border-b flex justify-between items-center shadow-sm ${darkMode ? 'bg-slate-900 border-slate-800' : 'bg-white border-slate-200'}`}>
          <div>
            <h1 className="text-xl font-bold flex items-center gap-2">
              {activeTab} 
              {loadingCloud && <span className="text-xs text-yellow-500 font-normal animate-pulse">☁️ Sinkronisasi Cloud...</span>}
            </h1>
            <p className="text-xs text-slate-400">Logistic Management System - PT. Ryan Eka Pratama Site Bontang</p>
          </div>

          <div className="flex items-center gap-2">
            {currentUser.role !== 'Viewer' ? (
              <>
                <button onClick={() => setShowModalMaterial(true)} className="bg-emerald-600 hover:bg-emerald-700 text-white px-3 py-1.5 text-xs font-semibold rounded-lg flex items-center gap-1 shadow">
                  <Plus size={14} /> Produk Baru
                </button>
                <button onClick={() => setShowModalMasuk(true)} className="bg-blue-600 hover:bg-blue-700 text-white px-3 py-1.5 text-xs font-semibold rounded-lg flex items-center gap-1 shadow">
                  <Plus size={14} /> Stok Masuk
                </button>
                <button onClick={() => setShowModalKeluar(true)} className="bg-amber-600 hover:bg-amber-700 text-white px-3.5 py-1.5 text-xs font-semibold rounded-lg flex items-center gap-1 shadow">
                  <Plus size={14} /> Stok Keluar
                </button>
              </>
            ) : (
              <span className="bg-slate-200 dark:bg-slate-800 text-slate-500 px-3 py-1.5 rounded-lg text-xs font-medium border">
                👁️ Mode Viewer (Hanya Lihat)
              </span>
            )}
          </div>
        </header>

        <main className="flex-1 p-6 overflow-auto">
          
          {/* DASHBOARD */}
          {activeTab === 'Dashboard' && (
            <div className="space-y-6">
              <div className={`p-4 rounded-xl shadow-sm border flex flex-wrap gap-3 items-center ${darkMode ? 'bg-slate-800 border-slate-700' : 'bg-white border-slate-200'}`}>
                <div className="text-xs font-semibold flex items-center gap-2">
                  <Calendar size={16} className="text-yellow-500" /> Filter Periode Dashboard:
                </div>
                <select value={selectedMonth} onChange={(e) => setSelectedMonth(e.target.value)} className={`px-3 py-2 text-sm border rounded-lg ${darkMode ? 'bg-slate-900 border-slate-700 text-white' : 'bg-slate-50 border-slate-200'}`}>
                  <option value="">Semua Bulan</option>
                  <option value="1">Januari</option><option value="2">Februari</option><option value="3">Maret</option>
                  <option value="4">April</option><option value="5">Mei</option><option value="6">Juni</option>
                  <option value="7">Juli</option><option value="8">Agustus</option><option value="9">September</option>
                  <option value="10">Oktober</option><option value="11">November</option><option value="12">Desember</option>
                </select>
                <select value={selectedYear} onChange={(e) => setSelectedYear(e.target.value)} className={`px-3 py-2 text-sm border rounded-lg ${darkMode ? 'bg-slate-900 border-slate-700 text-white' : 'bg-slate-50 border-slate-200'}`}>
                  <option value="">Semua Tahun</option>
                  {Array.from({ length: (new Date().getFullYear() + 2) - 2024 + 1 }, (_, i) => 2024 + i).map(y => (
                    <option key={y} value={y}>{y}</option>
                  ))}
                </select>
              </div>

              <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                <div className={`p-5 rounded-xl border shadow-sm flex items-center justify-between ${darkMode ? 'bg-slate-800 border-slate-700' : 'bg-white border-slate-200'}`}>
                  <div>
                    <p className="text-xs font-semibold text-slate-400 uppercase">Total Master Barang</p>
                    <h3 className="text-2xl font-bold mt-1 text-blue-600">{dataMaterial.length} Item</h3>
                  </div>
                  <div className="p-3 bg-blue-50 text-blue-600 rounded-xl"><Box size={24} /></div>
                </div>

                <div className={`p-5 rounded-xl border shadow-sm flex items-center justify-between ${darkMode ? 'bg-slate-800 border-slate-700' : 'bg-white border-slate-200'}`}>
                  <div>
                    <p className="text-xs font-semibold text-slate-400 uppercase">Barang Masuk (Periode Ini)</p>
                    <h3 className="text-2xl font-bold mt-1 text-emerald-600">{filteredMasuk.length} Transaksi</h3>
                  </div>
                  <div className="p-3 bg-emerald-50 text-emerald-600 rounded-xl"><TrendingUp size={24} /></div>
                </div>

                <div className={`p-5 rounded-xl border shadow-sm flex items-center justify-between ${darkMode ? 'bg-slate-800 border-slate-700' : 'bg-white border-slate-200'}`}>
                  <div>
                    <p className="text-xs font-semibold text-slate-400 uppercase">Barang Keluar (Periode Ini)</p>
                    <h3 className="text-2xl font-bold mt-1 text-rose-600">{filteredKeluar.length} Transaksi</h3>
                  </div>
                  <div className="p-3 bg-rose-50 text-rose-600 rounded-xl"><TrendingDown size={24} /></div>
                </div>
              </div>
            </div>
          )}

          {/* MASTER BARANG */}
          {activeTab === 'Master Barang' && (
            <div>
              <div className={`p-4 rounded-xl shadow-sm border mb-4 flex gap-3 items-center ${darkMode ? 'bg-slate-800 border-slate-700' : 'bg-white border-slate-200'}`}>
                <div className="relative flex-1">
                  <Search className="absolute left-3 top-2.5 text-slate-400" size={16} />
                  <input type="text" placeholder="Cari nama barang, kode, atau nomor part..." className={`w-full pl-9 pr-4 py-2 text-sm border rounded-lg focus:outline-none ${darkMode ? 'bg-slate-900 border-slate-700 text-white' : 'bg-slate-50 border-slate-200'}`} value={pencarian} onChange={(e) => { setPencarian(e.target.value); setPageMaterial(1); }} />
                </div>
              </div>

              <div className={`rounded-xl shadow-sm border overflow-hidden ${darkMode ? 'bg-slate-800 border-slate-700' : 'bg-white border-slate-200'}`}>
                <table className="w-full text-left text-xs">
                  <thead className={`border-b font-semibold ${darkMode ? 'bg-slate-900 text-slate-400 border-slate-700' : 'bg-slate-50 text-slate-600 border-slate-200'}`}>
                    <tr>
                      <th className="p-3.5 w-10 text-center"><input type="checkbox" className="rounded" /></th>
                      <th className="p-3.5">NAMA BARANG</th>
                      <th className="p-3.5">KODE / NOMOR</th>
                      <th className="p-3.5">UNIT</th>
                      <th className="p-3.5">LOKASI</th>
                      <th className="p-3.5">STOK</th>
                    </tr>
                  </thead>
                  <tbody className={`divide-y ${darkMode ? 'divide-slate-700' : 'divide-slate-100'}`}>
                    {paginatedMaterial.map((item, index) => (
                      <tr key={index} className={`hover:bg-opacity-50 ${darkMode ? 'hover:bg-slate-700' : 'hover:bg-slate-50'}`}>
                        <td className="p-3.5 text-center"><input type="checkbox" className="rounded" /></td>
                        <td className="p-3.5 font-semibold">{item['Nama Barang'] || '-'}</td>
                        <td className="p-3.5 text-slate-500 font-mono">{item.Kode} / {item['Nomor Barang'] || '-'}</td>
                        <td className="p-3.5"><span className="bg-yellow-100 text-yellow-800 px-2 py-0.5 rounded font-semibold text-[10px]">{item['Part Unit'] || '-'}</span></td>
                        <td className="p-3.5">Site Bontang</td>
                        <td className="p-3.5 font-bold text-emerald-600">{item.Qty || 0} {item['Base Unit of Measure'] || 'Ea'}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>

                <div className={`p-4 border-t flex justify-between items-center text-xs ${darkMode ? 'border-slate-700 text-slate-400' : 'border-slate-200 text-slate-600'}`}>
                  <span>Menampilkan halaman {pageMaterial} dari {totalPageMaterial || 1} (Total {filteredMaterial.length} data)</span>
                  <div className="flex gap-2">
                    <button onClick={() => setPageMaterial(p => Math.max(p - 1, 1))} disabled={pageMaterial === 1} className="px-3 py-1.5 border rounded-lg disabled:opacity-50">Sebelumnya</button>
                    <button onClick={() => setPageMaterial(p => Math.min(p + 1, totalPageMaterial))} disabled={pageMaterial >= totalPageMaterial} className="px-3 py-1.5 border rounded-lg disabled:opacity-50">Berikutnya</button>
                  </div>
                </div>
              </div>
            </div>
          )}

          {/* STOK MASUK */}
          {activeTab === 'Stok Masuk' && (
            <div>
              <div className={`p-4 rounded-xl shadow-sm border mb-4 flex flex-wrap gap-3 items-center ${darkMode ? 'bg-slate-800 border-slate-700' : 'bg-white border-slate-200'}`}>
                <div className="relative flex-1 min-w-[200px]">
                  <Search className="absolute left-3 top-2.5 text-slate-400" size={16} />
                  <input type="text" placeholder="Cari deskripsi atau kode..." className={`w-full pl-9 pr-4 py-2 text-sm border rounded-lg focus:outline-none ${darkMode ? 'bg-slate-900 border-slate-700 text-white' : 'bg-slate-50 border-slate-200'}`} value={pencarian} onChange={(e) => { setPencarian(e.target.value); setPageMasuk(1); }} />
                </div>
                <select value={selectedMonth} onChange={(e) => { setSelectedMonth(e.target.value); setPageMasuk(1); }} className={`px-3 py-2 text-sm border rounded-lg ${darkMode ? 'bg-slate-900 border-slate-700 text-white' : 'bg-slate-50 border-slate-200'}`}>
                  <option value="">Semua Bulan</option>
                  <option value="1">Januari</option><option value="2">Februari</option><option value="3">Maret</option>
                  <option value="4">April</option><option value="5">Mei</option><option value="6">Juni</option>
                  <option value="7">Juli</option><option value="8">Agustus</option><option value="9">September</option>
                  <option value="10">Oktober</option><option value="11">November</option><option value="12">Desember</option>
                </select>
                <select value={selectedYear} onChange={(e) => { setSelectedYear(e.target.value); setPageMasuk(1); }} className={`px-3 py-2 text-sm border rounded-lg ${darkMode ? 'bg-slate-900 border-slate-700 text-white' : 'bg-slate-50 border-slate-200'}`}>
                  <option value="">Semua Tahun</option>
                  {Array.from({ length: (new Date().getFullYear() + 2) - 2024 + 1 }, (_, i) => 2024 + i).map(y => (
                    <option key={y} value={y}>{y}</option>
                  ))}
                </select>
              </div>

              <div className={`rounded-xl shadow-sm border overflow-hidden ${darkMode ? 'bg-slate-800 border-slate-700' : 'bg-white border-slate-200'}`}>
                <table className="w-full text-left text-xs">
                  <thead className={`border-b font-semibold ${darkMode ? 'bg-slate-900 text-slate-400 border-slate-700' : 'bg-slate-50 text-slate-600 border-slate-200'}`}>
                    <tr>
                      <th className="p-3.5">TANGGAL</th>
                      <th className="p-3.5">UNIT</th>
                      <th className="p-3.5">KODE</th>
                      <th className="p-3.5">DESKRIPSI BARANG</th>
                      <th className="p-3.5">MERK</th>
                      <th className="p-3.5">QTY MASUK</th>
                      <th className="p-3.5">NO PO / SURAT JALAN</th>
                    </tr>
                  </thead>
                  <tbody className={`divide-y ${darkMode ? 'divide-slate-700' : 'divide-slate-100'}`}>
                    {paginatedMasuk.map((item, index) => (
                      <tr key={index} className={`hover:bg-opacity-50 ${darkMode ? 'hover:bg-slate-700' : 'hover:bg-slate-50'}`}>
                        <td className="p-3.5">{formatTanggalExcel(item.Tanggal)}</td>
                        <td className="p-3.5 font-semibold">{item.Unit || '-'}</td>
                        <td className="p-3.5 font-mono">{item.Kode || '-'}</td>
                        <td className="p-3.5 font-medium">{item.Description || '-'}</td>
                        <td className="p-3.5"><span className="bg-slate-200 dark:bg-slate-700 px-2 py-0.5 rounded text-[10px]">{item.Merk || '-'}</span></td>
                        <td className="p-3.5 font-bold text-emerald-600">{item.Qty || 0} {item['Qty.1'] || ''}</td>
                        <td className="p-3.5 text-slate-500">{item['No. PO'] || item['NO Surat Jalan'] || '-'}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>

                <div className={`p-4 border-t flex justify-between items-center text-xs ${darkMode ? 'border-slate-700 text-slate-400' : 'border-slate-200 text-slate-600'}`}>
                  <span>Menampilkan halaman {pageMasuk} dari {totalPageMasuk || 1} (Total {filteredMasuk.length} data)</span>
                  <div className="flex gap-2">
                    <button onClick={() => setPageMasuk(p => Math.max(p - 1, 1))} disabled={pageMasuk === 1} className="px-3 py-1.5 border rounded-lg disabled:opacity-50">Sebelumnya</button>
                    <button onClick={() => setPageMasuk(p => Math.min(p + 1, totalPageMasuk))} disabled={pageMasuk >= totalPageMasuk} className="px-3 py-1.5 border rounded-lg disabled:opacity-50">Berikutnya</button>
                  </div>
                </div>
              </div>
            </div>
          )}

          {/* STOK KELUAR */}
          {activeTab === 'Stok Keluar' && (
            <div>
              <div className={`p-4 rounded-xl shadow-sm border mb-4 flex flex-wrap gap-3 items-center ${darkMode ? 'bg-slate-800 border-slate-700' : 'bg-white border-slate-200'}`}>
                <div className="relative flex-1 min-w-[200px]">
                  <Search className="absolute left-3 top-2.5 text-slate-400" size={16} />
                  <input type="text" placeholder="Cari nama barang atau kode..." className={`w-full pl-9 pr-4 py-2 text-sm border rounded-lg focus:outline-none ${darkMode ? 'bg-slate-900 border-slate-700 text-white' : 'bg-slate-50 border-slate-200'}`} value={pencarian} onChange={(e) => { setPencarian(e.target.value); setPageKeluar(1); }} />
                </div>
                <select value={selectedMonth} onChange={(e) => { setSelectedMonth(e.target.value); setPageKeluar(1); }} className={`px-3 py-2 text-sm border rounded-lg ${darkMode ? 'bg-slate-900 border-slate-700 text-white' : 'bg-slate-50 border-slate-200'}`}>
                  <option value="">Semua Bulan</option>
                  <option value="1">Januari</option><option value="2">Februari</option><option value="3">Maret</option>
                  <option value="4">April</option><option value="5">Mei</option><option value="6">Juni</option>
                  <option value="7">Juli</option><option value="8">Agustus</option><option value="9">September</option>
                  <option value="10">Oktober</option><option value="11">November</option><option value="12">Desember</option>
                </select>
                <select value={selectedYear} onChange={(e) => { setSelectedYear(e.target.value); setPageKeluar(1); }} className={`px-3 py-2 text-sm border rounded-lg ${darkMode ? 'bg-slate-900 border-slate-700 text-white' : 'bg-slate-50 border-slate-200'}`}>
                  <option value="">Semua Tahun</option>
                  {Array.from({ length: (new Date().getFullYear() + 2) - 2024 + 1 }, (_, i) => 2024 + i).map(y => (
                    <option key={y} value={y}>{y}</option>
                  ))}
                </select>
              </div>

              <div className={`rounded-xl shadow-sm border overflow-hidden ${darkMode ? 'bg-slate-800 border-slate-700' : 'bg-white border-slate-200'}`}>
                <table className="w-full text-left text-xs">
                  <thead className={`border-b font-semibold ${darkMode ? 'bg-slate-900 text-slate-400 border-slate-700' : 'bg-slate-50 text-slate-600 border-slate-200'}`}>
                    <tr>
                      <th className="p-3.5">TANGGAL</th>
                      <th className="p-3.5">UNIT</th>
                      <th className="p-3.5">KODE</th>
                      <th className="p-3.5">NAMA BARANG</th>
                      <th className="p-3.5">MEKANIK</th>
                      <th className="p-3.5">QTY KELUAR</th>
                      <th className="p-3.5">KETERANGAN</th>
                    </tr>
                  </thead>
                  <tbody className={`divide-y ${darkMode ? 'divide-slate-700' : 'divide-slate-100'}`}>
                    {paginatedKeluar.map((item, index) => (
                      <tr key={index} className={`hover:bg-opacity-50 ${darkMode ? 'hover:bg-slate-700' : 'hover:bg-slate-50'}`}>
                        <td className="p-3.5">{formatTanggalExcel(item['Doc.Date'])}</td>
                        <td className="p-3.5 font-semibold">{item.Unit || '-'}</td>
                        <td className="p-3.5 font-mono">{item.Kode || '-'}</td>
                        <td className="p-3.5 font-medium">{item['Nama Barang'] || '-'}</td>
                        <td className="p-3.5">{item.Mekanik || '-'}</td>
                        <td className="p-3.5 font-bold text-rose-600">{item.Qty || 0}</td>
                        <td className="p-3.5 text-slate-400">{item.Keterangan || '-'}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>

                <div className={`p-4 border-t flex justify-between items-center text-xs ${darkMode ? 'border-slate-700 text-slate-400' : 'border-slate-200 text-slate-600'}`}>
                  <span>Menampilkan halaman {pageKeluar} dari {totalPageKeluar || 1} (Total {filteredKeluar.length} data)</span>
                  <div className="flex gap-2">
                    <button onClick={() => setPageKeluar(p => Math.max(p - 1, 1))} disabled={pageKeluar === 1} className="px-3 py-1.5 border rounded-lg disabled:opacity-50">Sebelumnya</button>
                    <button onClick={() => setPageKeluar(p => Math.min(p + 1, totalPageKeluar))} disabled={pageKeluar >= totalPageKeluar} className="px-3 py-1.5 border rounded-lg disabled:opacity-50">Berikutnya</button>
                  </div>
                </div>
              </div>
            </div>
          )}

          {/* STOK OPNAME */}
          {activeTab === 'Stok Opname' && (
            <div className="space-y-4">
              <div className={`p-4 rounded-xl border shadow-sm flex justify-between items-center ${darkMode ? 'bg-slate-800 border-slate-700' : 'bg-white border-slate-200'}`}>
                <div>
                  <h3 className="font-bold text-sm">Lembar Kerja Stock Opname Gudang Site Bontang</h3>
                  <p className="text-xs text-slate-400">Masukkan jumlah fisik barang yang dihitung di rak untuk mendeteksi selisih.</p>
                </div>
                {currentUser.role !== 'Viewer' ? (
                  <button onClick={() => alert("Hasil Opname berhasil disimpan!")} className="bg-emerald-600 hover:bg-emerald-700 text-white px-4 py-2 rounded-lg text-xs font-semibold shadow">
                    Simpan Sesi Opname
                  </button>
                ) : (
                  <span className="text-xs text-rose-500 font-semibold">Mode Viewer: Tidak dapat mengubah data</span>
                )}
              </div>

              <div className={`rounded-xl shadow-sm border overflow-hidden ${darkMode ? 'bg-slate-800 border-slate-700' : 'bg-white border-slate-200'}`}>
                <table className="w-full text-left text-xs">
                  <thead className={`border-b font-semibold ${darkMode ? 'bg-slate-900 text-slate-400 border-slate-700' : 'bg-slate-50 text-slate-600 border-slate-200'}`}>
                    <tr>
                      <th className="p-3.5">KODE</th>
                      <th className="p-3.5">NAMA BARANG</th>
                      <th className="p-3.5 text-center">STOK SISTEM</th>
                      <th className="p-3.5 text-center">STOK FISIK (HITUNG)</th>
                      <th className="p-3.5 text-center">SELISIH (VARIANCE)</th>
                    </tr>
                  </thead>
                  <tbody className={`divide-y ${darkMode ? 'divide-slate-700' : 'divide-slate-100'}`}>
                    {dataMaterial.slice(0, 30).map((item, index) => {
                      const kode = item.Kode || index;
                      const stokSistem = Number(item.Qty) || 0;
                      const stokFisik = physicalStocks[kode] !== undefined ? Number(physicalStocks[kode]) : stokSistem;
                      const selisih = stokFisik - stokSistem;

                      return (
                        <tr key={index} className={`hover:bg-opacity-50 ${darkMode ? 'hover:bg-slate-700' : 'hover:bg-slate-50'}`}>
                          <td className="p-3.5 font-mono font-bold">{item.Kode || '-'}</td>
                          <td className="p-3.5 font-medium">{item['Nama Barang'] || '-'}</td>
                          <td className="p-3.5 text-center font-bold text-blue-600">{stokSistem}</td>
                          <td className="p-3.5 text-center">
                            <input 
                              type="number" 
                              disabled={currentUser.role === 'Viewer'}
                              className={`w-24 p-1 text-center border rounded font-bold ${darkMode ? 'bg-slate-900 border-slate-700 text-white' : 'bg-slate-50 border-slate-300 text-slate-900'} ${currentUser.role === 'Viewer' ? 'opacity-50 cursor-not-allowed' : ''}`}
                              value={physicalStocks[kode] !== undefined ? physicalStocks[kode] : stokSistem}
                              onChange={(e) => setPhysicalStocks({...physicalStocks, [kode]: e.target.value})}
                            />
                          </td>
                          <td className="p-3.5 text-center">
                            {selisih === 0 ? (
                              <span className="bg-emerald-100 text-emerald-800 px-2 py-1 rounded font-bold text-[10px]">Sesuai (0)</span>
                            ) : selisih > 0 ? (
                              <span className="bg-blue-100 text-blue-800 px-2 py-1 rounded font-bold text-[10px]">Lebih (+{selisih})</span>
                            ) : (
                              <span className="bg-rose-100 text-rose-800 px-2 py-1 rounded font-bold text-[10px]">Kurang ({selisih})</span>
                            )}
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {/* KARTU STOK */}
          {activeTab === 'Kartu Stok' && (
            <div className="space-y-4">
              <div className={`p-4 rounded-xl border shadow-sm flex items-center gap-4 ${darkMode ? 'bg-slate-800 border-slate-700' : 'bg-white border-slate-200'}`}>
                <span className="text-xs font-semibold">Pilih atau Cari Barang:</span>
                <select 
                  value={selectedItemCode} 
                  onChange={(e) => setSelectedItemCode(e.target.value)}
                  className={`flex-1 p-2 text-xs border rounded-lg ${darkMode ? 'bg-slate-900 border-slate-700 text-white' : 'bg-slate-50 border-slate-300'}`}
                >
                  <option value="">-- Pilih Material Master Barang --</option>
                  {dataMaterial.map((m, idx) => (
                    <option key={idx} value={m.Kode}>{m.Kode} - {m['Nama Barang']}</option>
                  ))}
                </select>
              </div>

              {selectedItemCode ? (
                <div className={`p-5 rounded-xl border shadow-sm space-y-4 ${darkMode ? 'bg-slate-800 border-slate-700' : 'bg-white border-slate-200'}`}>
                  <div className="border-b pb-3">
                    <h3 className="font-bold text-sm">Riwayat Kartu Stok untuk Kode: <span className="text-yellow-500">{selectedItemCode}</span></h3>
                  </div>

                  <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                    <div className="border rounded-lg p-3">
                      <h4 className="font-bold text-xs text-emerald-600 mb-2 flex items-center gap-1"><TrendingUp size={14}/> Riwayat Barang Masuk</h4>
                      <div className="space-y-2 max-h-60 overflow-y-auto text-xs">
                        {dataMasuk.filter(m => m.Kode === selectedItemCode).length === 0 ? (
                          <p className="text-slate-400 italic">Tidak ada catatan masuk.</p>
                        ) : (
                          dataMasuk.filter(m => m.Kode === selectedItemCode).map((m, i) => (
                            <div key={i} className="p-2 border rounded bg-slate-50 dark:bg-slate-900 flex justify-between items-center">
                              <div>
                                <p className="font-semibold">{m.Description}</p>
                                <p className="text-[10px] text-slate-400">Tgl: {formatTanggalExcel(m.Tanggal)} | Unit: {m.Unit}</p>
                              </div>
                              <span className="font-bold text-emerald-600">+{m.Qty} {m['Qty.1']}</span>
                            </div>
                          ))
                        )}
                      </div>
                    </div>

                    <div className="border rounded-lg p-3">
                      <h4 className="font-bold text-xs text-rose-600 mb-2 flex items-center gap-1"><TrendingDown size={14}/> Riwayat Barang Keluar</h4>
                      <div className="space-y-2 max-h-60 overflow-y-auto text-xs">
                        {dataKeluar.filter(k => k.Kode === selectedItemCode).length === 0 ? (
                          <p className="text-slate-400 italic">Tidak ada catatan keluar.</p>
                        ) : (
                          dataKeluar.filter(k => k.Kode === selectedItemCode).map((k, i) => (
                            <div key={i} className="p-2 border rounded bg-slate-50 dark:bg-slate-900 flex justify-between items-center">
                              <div>
                                <p className="font-semibold">{k['Nama Barang']}</p>
                                <p className="text-[10px] text-slate-400">Tgl: {formatTanggalExcel(k['Doc.Date'])} | Mekanik: {k.Mekanik}</p>
                              </div>
                              <span className="font-bold text-rose-600">-{k.Qty}</span>
                            </div>
                          ))
                        )}
                      </div>
                    </div>
                  </div>
                </div>
              ) : (
                <div className={`p-12 text-center rounded-xl border ${darkMode ? 'bg-slate-800 border-slate-700' : 'bg-white border-slate-200'}`}>
                  <p className="text-sm text-slate-400">Silakan pilih salah satu kode barang di atas untuk melihat kartu stok dan riwayat keluar-masuknya secara mendetail.</p>
                </div>
              )}
            </div>
          )}

          {/* REKAP STOK */}
          {activeTab === 'Rekap Stok' && (
            <div className="space-y-4">
              <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                <div className={`p-4 rounded-xl border shadow-sm ${darkMode ? 'bg-slate-800 border-slate-700' : 'bg-white border-slate-200'}`}>
                  <p className="text-xs text-slate-400 font-semibold uppercase">Total Jenis Material</p>
                  <h3 className="text-2xl font-bold mt-1 text-blue-600">{dataMaterial.length} Item</h3>
                </div>
                <div className={`p-4 rounded-xl border shadow-sm ${darkMode ? 'bg-slate-800 border-slate-700' : 'bg-white border-slate-200'}`}>
                  <p className="text-xs text-slate-400 font-semibold uppercase">Total Transaksi Masuk</p>
                  <h3 className="text-2xl font-bold mt-1 text-emerald-600">{dataMasuk.length} Catatan</h3>
                </div>
                <div className={`p-4 rounded-xl border shadow-sm ${darkMode ? 'bg-slate-800 border-slate-700' : 'bg-white border-slate-200'}`}>
                  <p className="text-xs text-slate-400 font-semibold uppercase">Total Transaksi Keluar</p>
                  <h3 className="text-2xl font-bold mt-1 text-rose-600">{dataKeluar.length} Catatan</h3>
                </div>
              </div>

              <div className={`rounded-xl shadow-sm border overflow-hidden ${darkMode ? 'bg-slate-800 border-slate-700' : 'bg-white border-slate-200'}`}>
                <div className="p-4 border-b flex justify-between items-center">
                  <h3 className="font-bold text-sm">Rekapitulasi Status Persediaan Gudang</h3>
                  <button onClick={exportToExcel} className="bg-emerald-600 hover:bg-emerald-700 text-white px-3 py-1.5 rounded-lg text-xs font-semibold flex items-center gap-1.5">
                    <Download size={14}/> Unduh Rekap Excel
                  </button>
                </div>
                <table className="w-full text-left text-xs">
                  <thead className={`border-b font-semibold ${darkMode ? 'bg-slate-900 text-slate-400 border-slate-700' : 'bg-slate-50 text-slate-600 border-slate-200'}`}>
                    <tr>
                      <th className="p-3.5">KODE</th>
                      <th className="p-3.5">NAMA MATERIAL</th>
                      <th className="p-3.5 text-center">STOK TERSEDIA</th>
                      <th className="p-3.5 text-center">STATUS PERSEDIAAN</th>
                    </tr>
                  </thead>
                  <tbody className={`divide-y ${darkMode ? 'divide-slate-700' : 'divide-slate-100'}`}>
                    {dataMaterial.slice(0, 40).map((item, index) => {
                      const qty = Number(item.Qty) || 0;
                      return (
                        <tr key={index} className={`hover:bg-opacity-50 ${darkMode ? 'hover:bg-slate-700' : 'hover:bg-slate-50'}`}>
                          <td className="p-3.5 font-mono font-semibold">{item.Kode || '-'}</td>
                          <td className="p-3.5 font-medium">{item['Nama Barang'] || '-'}</td>
                          <td className="p-3.5 text-center font-bold text-emerald-600">{qty} {item['Base Unit of Measure'] || 'Ea'}</td>
                          <td className="p-3.5 text-center">
                            {qty > 5 ? (
                              <span className="bg-emerald-100 text-emerald-800 px-2 py-1 rounded font-bold text-[10px]">Stok Aman</span>
                            ) : qty > 0 ? (
                              <span className="bg-yellow-100 text-yellow-800 px-2 py-1 rounded font-bold text-[10px]">Stok Menipis</span>
                            ) : (
                              <span className="bg-rose-100 text-rose-800 px-2 py-1 rounded font-bold text-[10px]">Habis / Kosong</span>
                            )}
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {/* LAPORAN */}
          {activeTab === 'Laporan' && (
            <div className="space-y-6">
              <div className={`p-5 rounded-xl border shadow-sm flex flex-wrap justify-between items-center gap-4 ${darkMode ? 'bg-slate-800 border-slate-700' : 'bg-white border-slate-200'}`}>
                <div>
                  <h3 className="font-bold text-sm">Pusat Laporan Logistik PT. Ryan Eka Pratama</h3>
                  <p className="text-xs text-slate-400">Pilih jenis laporan yang ingin ditampilkan atau diekspor ke Excel.</p>
                </div>
                <div className="flex items-center gap-2">
                  <select 
                    value={jenisLaporan} 
                    onChange={(e) => setJenisLaporan(e.target.value)}
                    className={`p-2 text-xs border rounded-lg font-medium ${darkMode ? 'bg-slate-900 border-slate-700 text-white' : 'bg-slate-50 border-slate-300'}`}
                  >
                    <option value="stok">Laporan Stok Master Barang</option>
                    <option value="masuk">Laporan Transaksi Masuk</option>
                    <option value="keluar">Laporan Transaksi Keluar</option>
                  </select>
                  <button onClick={exportToExcel} className="bg-emerald-600 hover:bg-emerald-700 text-white px-3.5 py-2 rounded-lg text-xs font-semibold flex items-center gap-1.5 shadow">
                    <Download size={14}/> Cetak / Unduh Laporan (.xlsx)
                  </button>
                </div>
              </div>

              <div className={`rounded-xl shadow-sm border overflow-hidden ${darkMode ? 'bg-slate-800 border-slate-700' : 'bg-white border-slate-200'}`}>
                <div className="p-4 border-b bg-slate-50 dark:bg-slate-900">
                  <h4 className="font-bold text-xs uppercase tracking-wider text-slate-500">Pratinjau: {jenisLaporan === 'stok' ? 'Master Stok Material' : jenisLaporan === 'masuk' ? 'Barang Masuk Gudang' : 'Barang Keluar Pemeliharaan'}</h4>
                </div>
                
                {jenisLaporan === 'stok' && (
                  <table className="w-full text-left text-xs">
                    <thead className={`border-b font-semibold ${darkMode ? 'bg-slate-900 text-slate-400 border-slate-700' : 'bg-slate-50 text-slate-600 border-slate-200'}`}>
                      <tr>
                        <th className="p-3.5">KODE</th>
                        <th className="p-3.5">NAMA BARANG</th>
                        <th className="p-3.5">NOMOR PART</th>
                        <th className="p-3.5">PART UNIT</th>
                        <th className="p-3.5 text-center">JUMLAH STOK</th>
                      </tr>
                    </thead>
                    <tbody className={`divide-y ${darkMode ? 'divide-slate-700' : 'divide-slate-100'}`}>
                      {dataMaterial.slice(0, 25).map((item, idx) => (
                        <tr key={idx}>
                          <td className="p-3.5 font-mono">{item.Kode}</td>
                          <td className="p-3.5 font-semibold">{item['Nama Barang']}</td>
                          <td className="p-3.5 text-slate-500">{item['Nomor Barang'] || '-'}</td>
                          <td className="p-3.5">{item['Part Unit'] || '-'}</td>
                          <td className="p-3.5 text-center font-bold text-emerald-600">{item.Qty || 0}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                )}

                {jenisLaporan === 'masuk' && (
                  <table className="w-full text-left text-xs">
                    <thead className={`border-b font-semibold ${darkMode ? 'bg-slate-900 text-slate-400 border-slate-700' : 'bg-slate-50 text-slate-600 border-slate-200'}`}>
                      <tr>
                        <th className="p-3.5">TANGGAL</th>
                        <th className="p-3.5">UNIT</th>
                        <th className="p-3.5">KODE</th>
                        <th className="p-3.5">DESKRIPSI</th>
                        <th className="p-3.5">QTY MASUK</th>
                      </tr>
                    </thead>
                    <tbody className={`divide-y ${darkMode ? 'divide-slate-700' : 'divide-slate-100'}`}>
                      {dataMasuk.slice(0, 25).map((item, idx) => (
                        <tr key={idx}>
                          <td className="p-3.5">{formatTanggalExcel(item.Tanggal)}</td>
                          <td className="p-3.5 font-semibold">{item.Unit}</td>
                          <td className="p-3.5 font-mono">{item.Kode}</td>
                          <td className="p-3.5">{item.Description}</td>
                          <td className="p-3.5 font-bold text-emerald-600">+{item.Qty} {item['Qty.1']}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                )}

                {jenisLaporan === 'keluar' && (
                  <table className="w-full text-left text-xs">
                    <thead className={`border-b font-semibold ${darkMode ? 'bg-slate-900 text-slate-400 border-slate-700' : 'bg-slate-50 text-slate-600 border-slate-200'}`}>
                      <tr>
                        <th className="p-3.5">TANGGAL</th>
                        <th className="p-3.5">UNIT</th>
                        <th className="p-3.5">KODE</th>
                        <th className="p-3.5">NAMA BARANG</th>
                        <th className="p-3.5">MEKANIK</th>
                        <th className="p-3.5">QTY KELUAR</th>
                      </tr>
                    </thead>
                    <tbody className={`divide-y ${darkMode ? 'divide-slate-700' : 'divide-slate-100'}`}>
                      {dataKeluar.slice(0, 25).map((item, idx) => (
                        <tr key={idx}>
                          <td className="p-3.5">{formatTanggalExcel(item['Doc.Date'])}</td>
                          <td className="p-3.5 font-semibold">{item.Unit}</td>
                          <td className="p-3.5 font-mono">{item.Kode}</td>
                          <td className="p-3.5">{item['Nama Barang']}</td>
                          <td className="p-3.5">{item.Mekanik}</td>
                          <td className="p-3.5 font-bold text-rose-600">-{item.Qty}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                )}
              </div>
            </div>
          )}

          {/* PENGGUNA */}
          {activeTab === 'Pengguna' && (
            <div className="space-y-4">
              <div className={`p-4 rounded-xl border shadow-sm flex justify-between items-center ${darkMode ? 'bg-slate-800 border-slate-700' : 'bg-white border-slate-200'}`}>
                <div>
                  <h3 className="font-bold text-sm">Manajemen Pengguna & Hak Akses Sistem</h3>
                  <p className="text-xs text-slate-400">Daftar staf atau mekanik yang memiliki otoritas mengakses logistik Site Bontang.</p>
                </div>
                {currentUser.role === 'Admin' ? (
                  <button onClick={() => setShowModalUser(true)} className="bg-blue-600 hover:bg-blue-700 text-white px-3.5 py-2 rounded-lg text-xs font-semibold flex items-center gap-1.5 shadow">
                    <UserPlus size={14}/> Tambah Pengguna Baru
                  </button>
                ) : (
                  <span className="text-xs text-amber-500 font-semibold">ℹ️ Hanya Admin yang dapat menambah/menghapus pengguna</span>
                )}
              </div>

              <div className={`rounded-xl shadow-sm border overflow-hidden ${darkMode ? 'bg-slate-800 border-slate-700' : 'bg-white border-slate-200'}`}>
                <table className="w-full text-left text-xs">
                  <thead className={`border-b font-semibold ${darkMode ? 'bg-slate-900 text-slate-400 border-slate-700' : 'bg-slate-50 text-slate-600 border-slate-200'}`}>
                    <tr>
                      <th className="p-3.5">NAMA LENGKAP</th>
                      <th className="p-3.5">USERNAME</th>
                      <th className="p-3.5">PERAN (ROLE)</th>
                      <th className="p-3.5">SITE LOKASI</th>
                      <th className="p-3.5 text-center">STATUS</th>
                      {currentUser.role === 'Admin' && <th className="p-3.5 text-center">AKSI</th>}
                    </tr>
                  </thead>
                  <tbody className={`divide-y ${darkMode ? 'divide-slate-700' : 'divide-slate-100'}`}>
                    {daftarPengguna.map((u) => (
                      <tr key={u.id} className={`hover:bg-opacity-50 ${darkMode ? 'hover:bg-slate-700' : 'hover:bg-slate-50'}`}>
                        <td className="p-3.5 font-semibold flex items-center gap-2">
                          <div className="w-6 h-6 rounded-full bg-yellow-500 text-slate-950 flex items-center justify-center font-bold text-[10px]">{u.nama.charAt(0)}</div>
                          {u.nama}
                        </td>
                        <td className="p-3.5 font-mono text-slate-500">{u.username}</td>
                        <td className="p-3.5">
                          <span className={`px-2 py-0.5 rounded font-semibold text-[10px] ${
                            u.role === 'Admin' ? 'bg-purple-100 text-purple-800' :
                            u.role === 'Editor' ? 'bg-blue-100 text-blue-800' : 'bg-slate-200 text-slate-800'
                          }`}>
                            {u.role}
                          </span>
                        </td>
                        <td className="p-3.5">{u.site}</td>
                        <td className="p-3.5 text-center"><span className="bg-emerald-100 text-emerald-800 px-2 py-0.5 rounded font-bold text-[10px]">{u.status}</span></td>
                        {currentUser.role === 'Admin' && (
                          <td className="p-3.5 text-center">
                            <button 
                              onClick={() => handleDeleteUser(u.id, u.role)} 
                              className="p-1.5 bg-rose-100 hover:bg-rose-200 text-rose-700 rounded-lg transition-colors"
                              title="Hapus Pengguna"
                            >
                              <Trash2 size={14} />
                            </button>
                          </td>
                        )}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {/* PENGATURAN */}
          {activeTab === 'Pengaturan' && (
            <div className="space-y-6 max-w-3xl">
              <div className={`p-6 rounded-xl border shadow-sm space-y-4 ${darkMode ? 'bg-slate-800 border-slate-700' : 'bg-white border-slate-200'}`}>
                <h3 className="font-bold text-sm border-b pb-3 flex items-center gap-2 text-yellow-500">
                  <Sliders size={18}/> Konfigurasi Sistem Logistik Site Bontang
                </h3>

                <div className="space-y-3 text-xs">
                  <div>
                    <label className="block font-semibold mb-1">Nama Perusahaan / Instansi</label>
                    <input type="text" disabled className="w-full p-2 border rounded-lg bg-slate-100 dark:bg-slate-900 dark:border-slate-700 text-slate-500" value="PT. Ryan Eka Pratama"/>
                  </div>
                  <div>
                    <label className="block font-semibold mb-1">Lokasi Site Operasional</label>
                    <input type="text" disabled className="w-full p-2 border rounded-lg bg-slate-100 dark:bg-slate-900 dark:border-slate-700 text-slate-500" value="Site Bontang, Kalimantan Timur"/>
                  </div>
                  <div>
                    <label className="block font-semibold mb-1">Status Login Saat Ini</label>
                    <input type="text" disabled className="w-full p-2 border rounded-lg bg-slate-100 dark:bg-slate-900 dark:border-slate-700 font-bold text-yellow-500" value={`${currentUser.nama} (${currentUser.role})`}/>
                  </div>
                </div>

                <div className="pt-4 border-t space-y-3">
                  <h4 className="font-bold text-xs">Pemeliharaan Data & Pencadangan (Backup)</h4>
                  <div className="flex flex-wrap gap-3">
                    <button onClick={exportToExcel} className="px-4 py-2 bg-emerald-600 hover:bg-emerald-700 text-white rounded-lg text-xs font-semibold shadow flex items-center gap-1.5">
                      <Download size={14} /> Unduh / Ekspor Database Excel (.xlsx)
                    </button>
                  </div>
                </div>
              </div>
            </div>
          )}

          {/* MODAL INPUT: TAMBAH PENGGUNA */}
          {showModalUser && (
            <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50 p-4">
              <div className={`w-full max-w-md p-6 rounded-xl shadow-lg border ${darkMode ? 'bg-slate-800 border-slate-700 text-white' : 'bg-white border-slate-200 text-slate-800'}`}>
                <div className="flex justify-between items-center mb-4 border-b pb-2">
                  <h3 className="font-bold text-base">Tambah Pengguna Baru</h3>
                  <button onClick={() => setShowModalUser(false)} className="text-slate-400 hover:text-red-500"><X size={20}/></button>
                </div>
                <form onSubmit={handleAddUser} className="space-y-3 text-xs">
                  <div>
                    <label className="block font-semibold mb-1">Nama Lengkap</label>
                    <input type="text" required className={`w-full p-2 border rounded-lg ${darkMode ? 'bg-slate-900 border-slate-700 text-white placeholder-slate-500' : 'bg-slate-50 border-slate-300 text-slate-900'}`} value={formUser.nama} onChange={e => setFormUser({...formUser, nama: e.target.value})} placeholder="Misal: Budi Santoso"/>
                  </div>
                  <div>
                    <label className="block font-semibold mb-1">Username Login</label>
                    <input type="text" required className={`w-full p-2 border rounded-lg ${darkMode ? 'bg-slate-900 border-slate-700 text-white placeholder-slate-500' : 'bg-slate-50 border-slate-300 text-slate-900'}`} value={formUser.username} onChange={e => setFormUser({...formUser, username: e.target.value})} placeholder="budi_logistik"/>
                  </div>
                  <div>
                    <label className="block font-semibold mb-1">Peran / Hak Akses (Role)</label>
                    <select className={`w-full p-2 border rounded-lg font-semibold ${darkMode ? 'bg-slate-900 border-slate-700 text-white' : 'bg-slate-50 border-slate-300 text-slate-900'}`} value={formUser.role} onChange={e => setFormUser({...formUser, role: e.target.value})}>
                      <option value="Admin">Admin (Full Akses & Kelola User)</option>
                      <option value="Editor">Editor (Input & Kelola Stok)</option>
                      <option value="Viewer">Viewer (Hanya Bisa Melihat Data)</option>
                    </select>
                  </div>
                  <div className="flex justify-end gap-2 mt-4 pt-2 border-t">
                    <button type="button" onClick={() => setShowModalUser(false)} className="px-4 py-2 border rounded-lg">Batal</button>
                    <button type="submit" className="px-4 py-2 bg-blue-600 hover:bg-blue-700 text-white font-semibold rounded-lg">Simpan Pengguna</button>
                  </div>
                </form>
              </div>
            </div>
          )}

          {/* MODAL INPUT: PRODUK BARU */}
          {showModalMaterial && (
            <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50 p-4">
              <div className={`w-full max-w-lg p-6 rounded-xl shadow-lg border ${darkMode ? 'bg-slate-800 border-slate-700 text-white' : 'bg-white border-slate-200 text-slate-800'}`}>
                <div className="flex justify-between items-center mb-4 border-b pb-2">
                  <h3 className="font-bold text-base">Tambah Produk / Master Barang Baru</h3>
                  <button onClick={() => setShowModalMaterial(false)} className="text-slate-400 hover:text-red-500"><X size={20}/></button>
                </div>
                <form onSubmit={handleAddMaterial} className="space-y-3 text-xs">
                  <div><label className="block font-semibold mb-1">Kode Barang</label><input type="text" required className={`w-full p-2 border rounded-lg ${darkMode ? 'bg-slate-900 border-slate-700 text-white placeholder-slate-500' : 'bg-slate-50 border-slate-300 text-slate-900'}`} value={formMaterial.kode} onChange={e => setFormMaterial({...formMaterial, kode: e.target.value})} placeholder="U198"/></div>
                  <div><label className="block font-semibold mb-1">Nama Barang (Description)</label><input type="text" required className={`w-full p-2 border rounded-lg ${darkMode ? 'bg-slate-900 border-slate-700 text-white placeholder-slate-500' : 'bg-slate-50 border-slate-300 text-slate-900'}`} value={formMaterial.nama} onChange={e => setFormMaterial({...formMaterial, nama: e.target.value})} placeholder="FILTER OLI CAT"/></div>
                  <div><label className="block font-semibold mb-1">Nomor Part / Barang</label><input type="text" className={`w-full p-2 border rounded-lg ${darkMode ? 'bg-slate-900 border-slate-700 text-white placeholder-slate-500' : 'bg-slate-50 border-slate-300 text-slate-900'}`} value={formMaterial.nomor} onChange={e => setFormMaterial({...formMaterial, nomor: e.target.value})} placeholder="1R-1808"/></div>
                  <div className="grid grid-cols-2 gap-2">
                    <div><label className="block font-semibold mb-1">Part Unit</label><input type="text" className={`w-full p-2 border rounded-lg ${darkMode ? 'bg-slate-900 border-slate-700 text-white placeholder-slate-500' : 'bg-slate-50 border-slate-300 text-slate-900'}`} value={formMaterial.unit} onChange={e => setFormMaterial({...formMaterial, unit: e.target.value})} placeholder="REP 110"/></div>
                    <div><label className="block font-semibold mb-1">Qty Awal</label><input type="number" required className={`w-full p-2 border rounded-lg ${darkMode ? 'bg-slate-900 border-slate-700 text-white placeholder-slate-500' : 'bg-slate-50 border-slate-300 text-slate-900'}`} value={formMaterial.qty} onChange={e => setFormMaterial({...formMaterial, qty: e.target.value})} placeholder="0"/></div>
                  </div>
                  <div className="flex justify-end gap-2 mt-4 pt-2 border-t"><button type="button" onClick={() => setShowModalMaterial(false)} className="px-4 py-2 border rounded-lg">Batal</button><button type="submit" className="px-4 py-2 bg-emerald-600 hover:bg-emerald-700 text-white font-semibold rounded-lg">Simpan</button></div>
                </form>
              </div>
            </div>
          )}

          {/* MODAL INPUT: STOK MASUK */}
          {showModalMasuk && (
            <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50 p-4">
              <div className={`w-full max-w-lg p-6 rounded-xl shadow-lg border ${darkMode ? 'bg-slate-800 border-slate-700 text-white' : 'bg-white border-slate-200 text-slate-800'}`}>
                <div className="flex justify-between items-center mb-4 border-b pb-2">
                  <h3 className="font-bold text-base">Tambah Transaksi Stok Masuk</h3>
                  <button onClick={() => setShowModalMasuk(false)} className="text-slate-400 hover:text-red-500"><X size={20}/></button>
                </div>
                <form onSubmit={handleAddMasuk} className="space-y-3 text-xs">
                  <div className="grid grid-cols-2 gap-2">
                    <div><label className="block font-semibold mb-1">Tanggal</label><input type="date" required className={`w-full p-2 border rounded-lg ${darkMode ? 'bg-slate-900 border-slate-700 text-white' : 'bg-slate-50 border-slate-300 text-slate-900'}`} value={formMasuk.tanggal} onChange={e => setFormMasuk({...formMasuk, tanggal: e.target.value})}/></div>
                    <div><label className="block font-semibold mb-1">Unit</label><input type="text" required className={`w-full p-2 border rounded-lg ${darkMode ? 'bg-slate-900 border-slate-700 text-white placeholder-slate-500' : 'bg-slate-50 border-slate-300 text-slate-900'}`} value={formMasuk.unit} onChange={e => setFormMasuk({...formMasuk, unit: e.target.value})} placeholder="REP 110"/></div>
                  </div>
                  <div className="grid grid-cols-2 gap-2">
                    <div><label className="block font-semibold mb-1">Kode Barang</label><input type="text" required className={`w-full p-2 border rounded-lg ${darkMode ? 'bg-slate-900 border-slate-700 text-white placeholder-slate-500' : 'bg-slate-50 border-slate-300 text-slate-900'}`} value={formMasuk.kode} onChange={e => setFormMasuk({...formMasuk, kode: e.target.value})} placeholder="U045"/></div>
                    <div><label className="block font-semibold mb-1">Merk</label><input type="text" className={`w-full p-2 border rounded-lg ${darkMode ? 'bg-slate-900 border-slate-700 text-white placeholder-slate-500' : 'bg-slate-50 border-slate-300 text-slate-900'}`} value={formMasuk.merk} onChange={e => setFormMasuk({...formMasuk, merk: e.target.value})} placeholder="Genuine"/></div>
                  </div>
                  <div><label className="block font-semibold mb-1">Deskripsi Barang</label><input type="text" required className={`w-full p-2 border rounded-lg ${darkMode ? 'bg-slate-900 border-slate-700 text-white placeholder-slate-500' : 'bg-slate-50 border-slate-300 text-slate-900'}`} value={formMasuk.description} onChange={e => setFormMasuk({...formMasuk, description: e.target.value})} placeholder="OIL ENGINE"/></div>
                  <div className="grid grid-cols-3 gap-2">
                    <div><label className="block font-semibold mb-1">Qty</label><input type="number" required className={`w-full p-2 border rounded-lg ${darkMode ? 'bg-slate-900 border-slate-700 text-white placeholder-slate-500' : 'bg-slate-50 border-slate-300 text-slate-900'}`} value={formMasuk.qty} onChange={e => setFormMasuk({...formMasuk, qty: e.target.value})} placeholder="10"/></div>
                    <div><label className="block font-semibold mb-1">Satuan</label><input type="text" required className={`w-full p-2 border rounded-lg ${darkMode ? 'bg-slate-900 border-slate-700 text-white placeholder-slate-500' : 'bg-slate-50 border-slate-300 text-slate-900'}`} value={formMasuk.satuan} onChange={e => setFormMasuk({...formMasuk, satuan: e.target.value})} placeholder="Liter"/></div>
                    <div><label className="block font-semibold mb-1">No. PO</label><input type="text" className={`w-full p-2 border rounded-lg ${darkMode ? 'bg-slate-900 border-slate-700 text-white placeholder-slate-500' : 'bg-slate-50 border-slate-300 text-slate-900'}`} value={formMasuk.noPo} onChange={e => setFormMasuk({...formMasuk, noPo: e.target.value})} placeholder="PO-001"/></div>
                  </div>
                  <div className="flex justify-end gap-2 mt-4 pt-2 border-t"><button type="button" onClick={() => setShowModalMasuk(false)} className="px-4 py-2 border rounded-lg">Batal</button><button type="submit" className="px-4 py-2 bg-blue-600 hover:bg-blue-700 text-white font-semibold rounded-lg">Simpan</button></div>
                </form>
              </div>
            </div>
          )}

          {/* MODAL INPUT: STOK KELUAR */}
          {showModalKeluar && (
            <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50 p-4">
              <div className={`w-full max-w-lg p-6 rounded-xl shadow-lg border ${darkMode ? 'bg-slate-800 border-slate-700 text-white' : 'bg-white border-slate-200 text-slate-800'}`}>
                <div className="flex justify-between items-center mb-4 border-b pb-2">
                  <h3 className="font-bold text-base">Tambah Transaksi Stok Keluar</h3>
                  <button onClick={() => setShowModalKeluar(false)} className="text-slate-400 hover:text-red-500"><X size={20}/></button>
                </div>
                <form onSubmit={handleAddKeluar} className="space-y-3 text-xs">
                  <div className="grid grid-cols-2 gap-2">
                    <div><label className="block font-semibold mb-1">Tanggal</label><input type="date" required className={`w-full p-2 border rounded-lg ${darkMode ? 'bg-slate-900 border-slate-700 text-white' : 'bg-slate-50 border-slate-300 text-slate-900'}`} value={formKeluar.tanggal} onChange={e => setFormKeluar({...formKeluar, tanggal: e.target.value})}/></div>
                    <div><label className="block font-semibold mb-1">Unit</label><input type="text" required className={`w-full p-2 border rounded-lg ${darkMode ? 'bg-slate-900 border-slate-700 text-white placeholder-slate-500' : 'bg-slate-50 border-slate-300 text-slate-900'}`} value={formKeluar.unit} onChange={e => setFormKeluar({...formKeluar, unit: e.target.value})} placeholder="REP 110"/></div>
                  </div>
                  <div className="grid grid-cols-2 gap-2">
                    <div><label className="block font-semibold mb-1">Kode Barang</label><input type="text" placeholder="SC081" required className={`w-full p-2 border rounded-lg ${darkMode ? 'bg-slate-900 border-slate-700 text-white placeholder-slate-500' : 'bg-slate-50 border-slate-300 text-slate-900'}`} value={formKeluar.kode} onChange={e => setFormKeluar({...formKeluar, kode: e.target.value})}/></div>
                    <div><label className="block font-semibold mb-1">Mekanik</label><input type="text" placeholder="Endang" className={`w-full p-2 border rounded-lg ${darkMode ? 'bg-slate-900 border-slate-700 text-white placeholder-slate-500' : 'bg-slate-50 border-slate-300 text-slate-900'}`} value={formKeluar.mekanik} onChange={e => setFormKeluar({...formKeluar, mekanik: e.target.value})}/></div>
                  </div>
                  <div><label className="block font-semibold mb-1">Nama Barang</label><input type="text" placeholder="SPRING BELAKANG" required className={`w-full p-2 border rounded-lg ${darkMode ? 'bg-slate-900 border-slate-700 text-white placeholder-slate-500' : 'bg-slate-50 border-slate-300 text-slate-900'}`} value={formKeluar.namaBarang} onChange={e => setFormKeluar({...formKeluar, namaBarang: e.target.value})}/></div>
                  <div className="grid grid-cols-2 gap-2">
                    <div><label className="block font-semibold mb-1">Qty Keluar</label><input type="number" placeholder="1" required className={`w-full p-2 border rounded-lg ${darkMode ? 'bg-slate-900 border-slate-700 text-white placeholder-slate-500' : 'bg-slate-50 border-slate-300 text-slate-900'}`} value={formKeluar.qty} onChange={e => setFormKeluar({...formKeluar, qty: e.target.value})}/></div>
                    <div><label className="block font-semibold mb-1">Keterangan</label><input type="text" placeholder="Rusak / Ganti" className={`w-full p-2 border rounded-lg ${darkMode ? 'bg-slate-900 border-slate-700 text-white placeholder-slate-500' : 'bg-slate-50 border-slate-300 text-slate-900'}`} value={formKeluar.keterangan} onChange={e => setFormKeluar({...formKeluar, keterangan: e.target.value})}/></div>
                  </div>
                  <div className="flex justify-end gap-2 mt-4 pt-2 border-t"><button type="button" onClick={() => setShowModalKeluar(false)} className="px-4 py-2 border rounded-lg">Batal</button><button type="submit" className="px-4 py-2 bg-amber-600 hover:bg-amber-700 text-white font-semibold rounded-lg">Simpan</button></div>
                </form>
              </div>
            </div>
          )}

        </main>
      </div>
    </div>
  );
}