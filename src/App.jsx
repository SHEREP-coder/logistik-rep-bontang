import React, { useState, useEffect } from 'react';
import * as XLSX from 'xlsx';
import { 
  Search, Plus, LayoutDashboard, Box, ArrowDownToLine, ArrowUpFromLine, 
  FileText, ClipboardList, BookOpen, BarChart3, Users, Settings, 
  Moon, TrendingUp, TrendingDown, Calendar, X, Download, UserPlus, Sliders, Trash2
} from 'lucide-react';

import { initializeApp } from "firebase/app";
import { getFirestore, collection, getDocs, addDoc } from "firebase/firestore";

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
const db = getFirestore(app);

export default function App() {
  const [activeTab, setActiveTab] = useState('Dashboard');
  const [pencarian, setPencarian] = useState("");
  const [selectedMonth, setSelectedMonth] = useState("");
  const [selectedYear, setSelectedYear] = useState("");
  const [darkMode, setDarkMode] = useState(false);
  const [loadingCloud, setLoadingCloud] = useState(false);

  const [currentUser, setCurrentUser] = useState({
    nama: "Administrator PT. REP",
    username: "admin_rep",
    role: "Admin"
  });

  const [dataMaterial, setDataMaterial] = useState([]);
  const [dataMasuk, setDataMasuk] = useState([]);
  const [dataKeluar, setDataKeluar] = useState([]);
  const [selectedItemCode, setSelectedItemCode] = useState("");
  const [physicalStocks, setPhysicalStocks] = useState({});
  const [jenisLaporan, setJenisLaporan] = useState("stok");

  const [daftarPengguna, setDaftarPengguna] = useState([
    { id: '1', nama: "Administrator PT. REP", username: "admin_rep", role: "Admin", site: "Bontang", status: "Aktif" },
    { id: '2', nama: "Endang (Mekanik Lead)", username: "endang_mech", role: "Editor", site: "Bontang", status: "Aktif" },
    { id: '3', nama: "Petugas Logistik Site", username: "logistik_btg", role: "Viewer", site: "Bontang", status: "Aktif" }
  ]);
  const [showModalUser, setShowModalUser] = useState(false);
  const [formUser, setFormUser] = useState({ nama: '', username: '', role: 'Editor' });

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
        console.error("Gagal memuat data:", error);
      }
      setLoadingCloud(false);
    };
    fetchDataFromCloud();
  }, []);

  const handleAddMaterial = async (e) => {
    e.preventDefault();
    if (currentUser.role === 'Viewer') return alert("Akses ditolak!");
    const newItem = { Kode: formMaterial.kode, 'Nama Barang': formMaterial.nama, 'Nomor Barang': formMaterial.nomor, 'Part Unit': formMaterial.unit, Qty: Number(formMaterial.qty), 'Base Unit of Measure': 'Ea' };
    try {
      await addDoc(collection(db, "material"), newItem);
      setDataMaterial([newItem, ...dataMaterial]);
      setShowModalMaterial(false);
      setFormMaterial({ kode: '', nama: '', nomor: '', unit: 'Ea', qty: 0 });
      alert("Berhasil disimpan!");
    } catch (error) { alert("Gagal menyimpan."); }
  };

  const handleAddMasuk = async (e) => {
    e.preventDefault();
    if (currentUser.role === 'Viewer') return alert("Akses ditolak!");
    const newItem = { Tanggal: formMasuk.tanggal, Unit: formMasuk.unit, Kode: formMasuk.kode, Description: formMasuk.description, Merk: formMasuk.merk, Qty: Number(formMasuk.qty), 'Qty.1': formMasuk.satuan, 'No. PO': formMasuk.noPo };
    try {
      await addDoc(collection(db, "barangMasuk"), newItem);
      setDataMasuk([newItem, ...dataMasuk]);
      setShowModalMasuk(false);
      setFormMasuk({ tanggal: '', unit: '', kode: '', description: '', merk: '', qty: 0, satuan: 'Pcs', noPo: '' });
      alert("Berhasil dicatat!");
    } catch (error) { alert("Gagal."); }
  };

  const handleAddKeluar = async (e) => {
    e.preventDefault();
    if (currentUser.role === 'Viewer') return alert("Akses ditolak!");
    const newItem = { 'Doc.Date': formKeluar.tanggal, Unit: formKeluar.unit, Kode: formKeluar.kode, 'Nama Barang': formKeluar.namaBarang, Mekanik: formKeluar.mekanik, Qty: Number(formKeluar.qty), Keterangan: formKeluar.keterangan };
    try {
      await addDoc(collection(db, "barangKeluar"), newItem);
      setDataKeluar([newItem, ...dataKeluar]);
      setShowModalKeluar(false);
      setFormKeluar({ tanggal: '', unit: '', kode: '', namaBarang: '', mekanik: '', qty: 0, keterangan: '' });
      alert("Berhasil dicatat!");
    } catch (error) { alert("Gagal."); }
  };

  const handleAddUser = (e) => {
    e.preventDefault();
    if (currentUser.role !== 'Admin') return alert("Hanya Admin!");
    const newUser = { id: String(Date.now()), nama: formUser.nama, username: formUser.username, role: formUser.role, site: "Bontang", status: "Aktif" };
    setDaftarPengguna([...daftarPengguna, newUser]);
    setShowModalUser(false);
    setFormUser({ nama: '', username: '', role: 'Editor' });
    alert("Pengguna ditambahkan!");
  };

  const handleDeleteUser = (id, roleTarget) => {
    if (currentUser.role !== 'Admin') return alert("Hanya Admin!");
    if (confirm("Hapus pengguna ini?")) setDaftarPengguna(daftarPengguna.filter(u => u.id !== id));
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
    let dateInfo = typeof serial === 'string' && serial.includes('-') ? new Date(serial) : new Date((Math.floor(Number(serial) - 25569)) * 86400 * 1000);
    return isNaN(dateInfo.getTime()) ? null : dateInfo;
  };

  const formatTanggalExcel = (serial) => {
    const dateInfo = parseExcelDate(serial);
    return !dateInfo ? '-' : dateInfo.toLocaleDateString('id-ID', { day: 'numeric', month: 'short', year: 'numeric' });
  };

  const matchesMonthYear = (serial) => {
    if (!selectedMonth && !selectedYear) return true;
    const dateObj = parseExcelDate(serial);
    if (!dateObj) return false;
    if (selectedMonth && (dateObj.getMonth() + 1) !== Number(selectedMonth)) return false;
    if (selectedYear && dateObj.getFullYear() !== Number(selectedYear)) return false;
    return true;
  };

  const filteredMaterial = dataMaterial.filter(item => 
    item['Nama Barang']?.toLowerCase().includes(pencarian.toLowerCase()) || 
    item.Kode?.toLowerCase().includes(pencarian.toLowerCase()) ||
    item['Nomor Barang']?.toString().toLowerCase().includes(pencarian.toLowerCase())
  );

  const filteredMasuk = dataMasuk.filter(item => (item.Description?.toLowerCase().includes(pencarian.toLowerCase()) || item.Kode?.toLowerCase().includes(pencarian.toLowerCase())) && matchesMonthYear(item.Tanggal));
  const filteredKeluar = dataKeluar.filter(item => (item['Nama Barang']?.toLowerCase().includes(pencarian.toLowerCase()) || item.Kode?.toLowerCase().includes(pencarian.toLowerCase())) && matchesMonthYear(item['Doc.Date']));

  const paginatedMaterial = filteredMaterial.slice((pageMaterial - 1) * rowsPerPage, pageMaterial * rowsPerPage);
  const paginatedMasuk = filteredMasuk.slice((pageMasuk - 1) * rowsPerPage, pageMasuk * rowsPerPage);
  const paginatedKeluar = filteredKeluar.slice((pageKeluar - 1) * rowsPerPage, pageKeluar * rowsPerPage);

  const totalPageMaterial = Math.ceil(filteredMaterial.length / rowsPerPage);
  const totalPageMasuk = Math.ceil(filteredMasuk.length / rowsPerPage);
  const totalPageKeluar = Math.ceil(filteredKeluar.length / rowsPerPage);

  const inputClass = "w-full p-2 border border-slate-300 rounded-lg bg-white text-slate-900 placeholder-slate-400 focus:outline-none focus:border-blue-500 font-medium";
  const selectClass = "px-3 py-2 text-sm border border-slate-300 rounded-lg bg-white text-slate-900 font-medium";
  const searchClass = "w-full pl-9 pr-4 py-2 text-sm border border-slate-300 rounded-lg bg-white text-slate-900 focus:outline-none focus:border-blue-500 font-medium";

  return (
    <div className={`flex h-screen font-sans ${darkMode ? 'bg-slate-900 text-white' : 'bg-slate-50 text-slate-800'}`}>
      
      {/* SIDEBAR */}
      <div className={`w-72 flex flex-col border-r ${darkMode ? 'bg-slate-950 border-slate-800 text-slate-300' : 'bg-slate-900 border-slate-800 text-slate-300'}`}>
        <div className="p-4 border-b border-slate-800 flex items-center gap-3 bg-slate-950">
          <div className="w-10 h-10 flex items-center justify-center">
            <img src="/logo_PT_REP_transparan.png" alt="Logo" className="w-full h-full object-contain filter drop-shadow-[0_2px_4px_rgba(0,0,0,0.5)]" />
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
              className="w-full p-1.5 rounded border text-xs border-slate-700 bg-white text-slate-900 font-medium"
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
            <div className="w-8 h-8 rounded-full bg-emerald-600 flex items-center font-bold text-white justify-center text-xs">{currentUser.nama.charAt(0)}</div>
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
                <button onClick={() => setShowModalMaterial(true)} className="bg-emerald-600 hover:bg-emerald-700 text-white px-3 py-1.5 text-xs font-semibold rounded-lg flex items-center gap-1"><Plus size={14} /> Produk Baru</button>
                <button onClick={() => setShowModalMasuk(true)} className="bg-blue-600 hover:bg-blue-700 text-white px-3 py-1.5 text-xs font-semibold rounded-lg flex items-center gap-1"><Plus size={14} /> Stok Masuk</button>
                <button onClick={() => setShowModalKeluar(true)} className="bg-amber-600 hover:bg-amber-700 text-white px-3.5 py-1.5 text-xs font-semibold rounded-lg flex items-center gap-1"><Plus size={14} /> Stok Keluar</button>
              </>
            ) : (<span className="bg-slate-200 dark:bg-slate-800 text-slate-500 px-3 py-1.5 rounded-lg text-xs font-medium border">👁️ Mode Viewer</span>)}
          </div>
        </header>

        <main className="flex-1 p-6 overflow-auto">
          
          {activeTab === 'Dashboard' && (
            <div className="space-y-6">
              <div className={`p-4 rounded-xl shadow-sm border flex flex-wrap gap-3 items-center ${darkMode ? 'bg-slate-800 border-slate-700' : 'bg-white border-slate-200'}`}>
                <div className="text-xs font-semibold flex items-center gap-2"><Calendar size={16} className="text-yellow-500" /> Filter Periode:</div>
                <select value={selectedMonth} onChange={(e) => setSelectedMonth(e.target.value)} className={selectClass}>
                  <option value="">Semua Bulan</option>
                  <option value="1">Januari</option><option value="2">Februari</option><option value="3">Maret</option>
                  <option value="4">April</option><option value="5">Mei</option><option value="6">Juni</option>
                  <option value="7">Juli</option><option value="8">Agustus</option><option value="9">September</option>
                  <option value="10">Oktober</option><option value="11">November</option><option value="12">Desember</option>
                </select>
                <select value={selectedYear} onChange={(e) => setSelectedYear(e.target.value)} className={selectClass}>
                  <option value="">Semua Tahun</option>
                  {Array.from({ length: 4 }, (_, i) => 2024 + i).map(y => (<option key={y} value={y}>{y}</option>))}
                </select>
              </div>
              <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                <div className={`p-5 rounded-xl border shadow-sm flex items-center justify-between ${darkMode ? 'bg-slate-800 border-slate-700' : 'bg-white border-slate-200'}`}>
                  <div><p className="text-xs font-semibold text-slate-400 uppercase">Total Master Barang</p><h3 className="text-2xl font-bold mt-1 text-blue-600">{dataMaterial.length} Item</h3></div>
                  <div className="p-3 bg-blue-50 text-blue-600 rounded-xl"><Box size={24} /></div>
                </div>
                <div className={`p-5 rounded-xl border shadow-sm flex items-center justify-between ${darkMode ? 'bg-slate-800 border-slate-700' : 'bg-white border-slate-200'}`}>
                  <div><p className="text-xs font-semibold text-slate-400 uppercase">Barang Masuk</p><h3 className="text-2xl font-bold mt-1 text-emerald-600">{filteredMasuk.length} Trx</h3></div>
                  <div className="p-3 bg-emerald-50 text-emerald-600 rounded-xl"><TrendingUp size={24} /></div>
                </div>
                <div className={`p-5 rounded-xl border shadow-sm flex items-center justify-between ${darkMode ? 'bg-slate-800 border-slate-700' : 'bg-white border-slate-200'}`}>
                  <div><p className="text-xs font-semibold text-slate-400 uppercase">Barang Keluar</p><h3 className="text-2xl font-bold mt-1 text-rose-600">{filteredKeluar.length} Trx</h3></div>
                  <div className="p-3 bg-rose-50 text-rose-600 rounded-xl"><TrendingDown size={24} /></div>
                </div>
              </div>
            </div>
          )}

          {activeTab === 'Master Barang' && (
            <div>
              <div className={`p-4 rounded-xl shadow-sm border mb-4 flex gap-3 items-center ${darkMode ? 'bg-slate-800 border-slate-700' : 'bg-white border-slate-200'}`}>
                <div className="relative flex-1">
                  <Search className="absolute left-3 top-2.5 text-slate-400" size={16} />
                  <input type="text" placeholder="Cari nama barang..." className={searchClass} value={pencarian} onChange={(e) => { setPencarian(e.target.value); setPageMaterial(1); }} />
                </div>
              </div>
              <div className={`rounded-xl shadow-sm border overflow-hidden ${darkMode ? 'bg-slate-800 border-slate-700' : 'bg-white border-slate-200'}`}>
                <table className="w-full text-left text-xs">
                  <thead className={`border-b font-semibold ${darkMode ? 'bg-slate-900 text-slate-400 border-slate-700' : 'bg-slate-50 text-slate-600 border-slate-200'}`}>
                    <tr>
                      <th className="p-3.5">NAMA BARANG</th>
                      <th className="p-3.5">KODE / NOMOR</th>
                      <th className="p-3.5">UNIT</th>
                      <th className="p-3.5">STOK</th>
                    </tr>
                  </thead>
                  <tbody className={`divide-y ${darkMode ? 'divide-slate-700' : 'divide-slate-100'}`}>
                    {paginatedMaterial.map((item, index) => (
                      <tr key={index} className={`hover:bg-opacity-50 ${darkMode ? 'hover:bg-slate-700' : 'hover:bg-slate-50'}`}>
                        <td className="p-3.5 font-semibold">{item['Nama Barang'] || '-'}</td>
                        <td className="p-3.5 text-slate-500 font-mono">{item.Kode} / {item['Nomor Barang'] || '-'}</td>
                        <td className="p-3.5">{item['Part Unit'] || '-'}</td>
                        <td className="p-3.5 font-bold text-emerald-600">{item.Qty || 0}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {activeTab === 'Stok Masuk' && (
            <div>
              <div className={`p-4 rounded-xl shadow-sm border mb-4 flex gap-3 items-center ${darkMode ? 'bg-slate-800 border-slate-700' : 'bg-white border-slate-200'}`}>
                <div className="relative flex-1">
                  <Search className="absolute left-3 top-2.5 text-slate-400" size={16} />
                  <input type="text" placeholder="Cari nama barang..." className={searchClass} value={pencarian} onChange={(e) => { setPencarian(e.target.value); setPageMasuk(1); }} />
                </div>
                <select value={selectedMonth} onChange={(e) => setSelectedMonth(e.target.value)} className={selectClass}>
                  <option value="">Semua Bulan</option>
                  <option value="1">Januari</option><option value="2">Februari</option><option value="3">Maret</option>
                </select>
              </div>
              <div className={`rounded-xl shadow-sm border overflow-hidden ${darkMode ? 'bg-slate-800 border-slate-700' : 'bg-white border-slate-200'}`}>
                <table className="w-full text-left text-xs">
                  <thead className={`border-b font-semibold ${darkMode ? 'bg-slate-900 text-slate-400 border-slate-700' : 'bg-slate-50 text-slate-600 border-slate-200'}`}>
                    <tr>
                      <th className="p-3.5">TANGGAL</th>
                      <th className="p-3.5">KODE</th>
                      <th className="p-3.5">DESKRIPSI</th>
                      <th className="p-3.5">QTY</th>
                    </tr>
                  </thead>
                  <tbody className={`divide-y ${darkMode ? 'divide-slate-700' : 'divide-slate-100'}`}>
                    {paginatedMasuk.map((item, index) => (
                      <tr key={index} className={`hover:bg-opacity-50 ${darkMode ? 'hover:bg-slate-700' : 'hover:bg-slate-50'}`}>
                        <td className="p-3.5">{formatTanggalExcel(item.Tanggal)}</td>
                        <td className="p-3.5 font-mono">{item.Kode}</td>
                        <td className="p-3.5">{item.Description}</td>
                        <td className="p-3.5 font-bold text-emerald-600">+{item.Qty}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {activeTab === 'Stok Keluar' && (
            <div>
              <div className={`p-4 rounded-xl shadow-sm border mb-4 flex gap-3 items-center ${darkMode ? 'bg-slate-800 border-slate-700' : 'bg-white border-slate-200'}`}>
                <div className="relative flex-1">
                  <Search className="absolute left-3 top-2.5 text-slate-400" size={16} />
                  <input type="text" placeholder="Cari nama barang..." className={searchClass} value={pencarian} onChange={(e) => { setPencarian(e.target.value); setPageKeluar(1); }} />
                </div>
              </div>
              <div className={`rounded-xl shadow-sm border overflow-hidden ${darkMode ? 'bg-slate-800 border-slate-700' : 'bg-white border-slate-200'}`}>
                <table className="w-full text-left text-xs">
                  <thead className={`border-b font-semibold ${darkMode ? 'bg-slate-900 text-slate-400 border-slate-700' : 'bg-slate-50 text-slate-600 border-slate-200'}`}>
                    <tr>
                      <th className="p-3.5">TANGGAL</th>
                      <th className="p-3.5">KODE</th>
                      <th className="p-3.5">NAMA BARANG</th>
                      <th className="p-3.5">QTY</th>
                    </tr>
                  </thead>
                  <tbody className={`divide-y ${darkMode ? 'divide-slate-700' : 'divide-slate-100'}`}>
                    {paginatedKeluar.map((item, index) => (
                      <tr key={index} className={`hover:bg-opacity-50 ${darkMode ? 'hover:bg-slate-700' : 'hover:bg-slate-50'}`}>
                        <td className="p-3.5">{formatTanggalExcel(item['Doc.Date'])}</td>
                        <td className="p-3.5 font-mono">{item.Kode}</td>
                        <td className="p-3.5">{item['Nama Barang']}</td>
                        <td className="p-3.5 font-bold text-rose-600">-{item.Qty}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {/* STOK OPNAME */}
          {activeTab === 'Stok Opname' && (
            <div className="space-y-4">
              <div className={`p-4 rounded-xl border shadow-sm flex justify-between items-center ${darkMode ? 'bg-slate-800 border-slate-700' : 'bg-white border-slate-200'}`}>
                <div>
                  <h3 className="font-bold text-sm">Lembar Kerja Stock Opname Gudang Site Bontang</h3>
                  <p className="text-xs text-slate-400">Masukkan jumlah fisik barang untuk memeriksa selisih stok.</p>
                </div>
                <button onClick={() => alert("Hasil Opname berhasil disimpan!")} className="bg-emerald-600 hover:bg-emerald-700 text-white px-4 py-2 rounded-lg text-xs font-semibold shadow">
                  Simpan Sesi Opname
                </button>
              </div>

              <div className={`rounded-xl shadow-sm border overflow-hidden ${darkMode ? 'bg-slate-800 border-slate-700' : 'bg-white border-slate-200'}`}>
                <table className="w-full text-left text-xs">
                  <thead className={`border-b font-semibold ${darkMode ? 'bg-slate-900 text-slate-400 border-slate-700' : 'bg-slate-50 text-slate-600 border-slate-200'}`}>
                    <tr>
                      <th className="p-3.5">KODE</th>
                      <th className="p-3.5">NAMA BARANG</th>
                      <th className="p-3.5 text-center">STOK SISTEM</th>
                      <th className="p-3.5 text-center">STOK FISIK (HITUNG)</th>
                      <th className="p-3.5 text-center">SELISIH</th>
                    </tr>
                  </thead>
                  <tbody className={`divide-y ${darkMode ? 'divide-slate-700' : 'divide-slate-100'}`}>
                    {dataMaterial.slice(0, 30).map((item, index) => {
                      const kode = item.Kode || index;
                      const stokSistem = Number(item.Qty) || 0;
                      const stokFisik = physicalStocks[kode] !== undefined ? Number(physicalStocks[kode]) : stokSistem;
                      const selisih = stokFisik - stokSistem;

                      return (
                        <tr key={index}>
                          <td className="p-3.5 font-mono font-bold">{item.Kode || '-'}</td>
                          <td className="p-3.5 font-medium">{item['Nama Barang'] || '-'}</td>
                          <td className="p-3.5 text-center font-bold text-blue-600">{stokSistem}</td>
                          <td className="p-3.5 text-center">
                            <input 
                              type="number" 
                              className="w-24 p-1 text-center border border-slate-300 rounded font-bold bg-white text-slate-900"
                              value={physicalStocks[kode] !== undefined ? physicalStocks[kode] : stokSistem}
                              onChange={(e) => setPhysicalStocks({...physicalStocks, [kode]: e.target.value})}
                            />
                          </td>
                          <td className="p-3.5 text-center">
                            {selisih === 0 ? <span className="bg-emerald-100 text-emerald-800 px-2 py-1 rounded font-bold text-[10px]">Sesuai (0)</span> : selisih > 0 ? <span className="bg-blue-100 text-blue-800 px-2 py-1 rounded font-bold text-[10px]">Lebih (+{selisih})</span> : <span className="bg-rose-100 text-rose-800 px-2 py-1 rounded font-bold text-[10px]">Kurang ({selisih})</span>}
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
                <select value={selectedItemCode} onChange={(e) => setSelectedItemCode(e.target.value)} className={selectClass}>
                  <option value="">-- Pilih Material Master Barang --</option>
                  {dataMaterial.map((m, idx) => (<option key={idx} value={m.Kode}>{m.Kode} - {m['Nama Barang']}</option>))}
                </select>
              </div>

              {selectedItemCode ? (
                <div className={`p-5 rounded-xl border shadow-sm space-y-4 ${darkMode ? 'bg-slate-800 border-slate-700' : 'bg-white border-slate-200'}`}>
                  <h3 className="font-bold text-sm border-b pb-2">Riwayat Kartu Stok: <span className="text-yellow-500">{selectedItemCode}</span></h3>
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                    <div className="border rounded-lg p-3">
                      <h4 className="font-bold text-xs text-emerald-600 mb-2">Riwayat Masuk</h4>
                      <div className="space-y-2 max-h-60 overflow-y-auto text-xs">
                        {dataMasuk.filter(m => m.Kode === selectedItemCode).map((m, i) => (
                          <div key={i} className="p-2 border rounded bg-white text-slate-900 flex justify-between">
                            <span>{m.Description} (Tgl: {formatTanggalExcel(m.Tanggal)})</span>
                            <span className="font-bold text-emerald-600">+{m.Qty}</span>
                          </div>
                        ))}
                      </div>
                    </div>
                    <div className="border rounded-lg p-3">
                      <h4 className="font-bold text-xs text-rose-600 mb-2">Riwayat Keluar</h4>
                      <div className="space-y-2 max-h-60 overflow-y-auto text-xs">
                        {dataKeluar.filter(k => k.Kode === selectedItemCode).map((k, i) => (
                          <div key={i} className="p-2 border rounded bg-white text-slate-900 flex justify-between">
                            <span>{k['Nama Barang']} (Mekanik: {k.Mekanik})</span>
                            <span className="font-bold text-rose-600">-{k.Qty}</span>
                          </div>
                        ))}
                      </div>
                    </div>
                  </div>
                </div>
              ) : (
                <div className={`p-12 text-center rounded-xl border ${darkMode ? 'bg-slate-800 border-slate-700' : 'bg-white border-slate-200'}`}>
                  <p className="text-sm text-slate-400">Silakan pilih salah satu kode barang di atas.</p>
                </div>
              )}
            </div>
          )}

          {/* REKAP STOK */}
          {activeTab === 'Rekap Stok' && (
            <div className={`rounded-xl shadow-sm border overflow-hidden ${darkMode ? 'bg-slate-800 border-slate-700' : 'bg-white border-slate-200'}`}>
              <div className="p-4 border-b flex justify-between items-center">
                <h3 className="font-bold text-sm">Rekapitulasi Status Persediaan Gudang</h3>
                <button onClick={exportToExcel} className="bg-emerald-600 hover:bg-emerald-700 text-white px-3 py-1.5 rounded-lg text-xs font-semibold flex items-center gap-1.5"><Download size={14}/> Unduh Excel</button>
              </div>
              <table className="w-full text-left text-xs">
                <thead className={`border-b font-semibold ${darkMode ? 'bg-slate-900 text-slate-400 border-slate-700' : 'bg-slate-50 text-slate-600 border-slate-200'}`}>
                  <tr><th className="p-3.5">KODE</th><th className="p-3.5">NAMA MATERIAL</th><th className="p-3.5 text-center">STOK TERSEDIA</th><th className="p-3.5 text-center">STATUS</th></tr>
                </thead>
                <tbody className={`divide-y ${darkMode ? 'divide-slate-700' : 'divide-slate-100'}`}>
                  {dataMaterial.slice(0, 40).map((item, index) => {
                    const qty = Number(item.Qty) || 0;
                    return (
                      <tr key={index}>
                        <td className="p-3.5 font-mono font-semibold">{item.Kode || '-'}</td>
                        <td className="p-3.5 font-medium">{item['Nama Barang'] || '-'}</td>
                        <td className="p-3.5 text-center font-bold text-emerald-600">{qty}</td>
                        <td className="p-3.5 text-center">
                          {qty > 5 ? <span className="bg-emerald-100 text-emerald-800 px-2 py-1 rounded font-bold text-[10px]">Aman</span> : <span className="bg-rose-100 text-rose-800 px-2 py-1 rounded font-bold text-[10px]">Menipis</span>}
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          )}

          {/* LAPORAN */}
          {activeTab === 'Laporan' && (
            <div className="space-y-4">
              <div className={`p-4 rounded-xl border shadow-sm flex justify-between items-center ${darkMode ? 'bg-slate-800 border-slate-700' : 'bg-white border-slate-200'}`}>
                <h3 className="font-bold text-sm">Pusat Laporan Logistik</h3>
                <button onClick={exportToExcel} className="bg-emerald-600 hover:bg-emerald-700 text-white px-3 py-2 rounded-lg text-xs font-semibold flex items-center gap-1.5"><Download size={14}/> Unduh Laporan Excel</button>
              </div>
            </div>
          )}

          {/* PENGGUNA */}
          {activeTab === 'Pengguna' && (
            <div className="space-y-4">
              <div className={`p-4 rounded-xl border shadow-sm flex justify-between items-center ${darkMode ? 'bg-slate-800 border-slate-700' : 'bg-white border-slate-200'}`}>
                <div>
                  <h3 className="font-bold text-sm">Manajemen Pengguna</h3>
                  <p className="text-xs text-slate-400">Daftar staf yang memiliki akses sistem.</p>
                </div>
                {currentUser.role === 'Admin' && (
                  <button onClick={() => setShowModalUser(true)} className="bg-blue-600 hover:bg-blue-700 text-white px-3.5 py-2 rounded-lg text-xs font-semibold flex items-center gap-1.5"><UserPlus size={14}/> Tambah Pengguna</button>
                )}
              </div>
              <div className={`rounded-xl shadow-sm border overflow-hidden ${darkMode ? 'bg-slate-800 border-slate-700' : 'bg-white border-slate-200'}`}>
                <table className="w-full text-left text-xs">
                  <thead className={`border-b font-semibold ${darkMode ? 'bg-slate-900 text-slate-400 border-slate-700' : 'bg-slate-50 text-slate-600 border-slate-200'}`}>
                    <tr><th className="p-3.5">NAMA</th><th className="p-3.5">USERNAME</th><th className="p-3.5">ROLE</th><th className="p-3.5 text-center">AKSI</th></tr>
                  </thead>
                  <tbody className={`divide-y ${darkMode ? 'divide-slate-700' : 'divide-slate-100'}`}>
                    {daftarPengguna.map((u) => (
                      <tr key={u.id}>
                        <td className="p-3.5 font-semibold">{u.nama}</td>
                        <td className="p-3.5 font-mono">{u.username}</td>
                        <td className="p-3.5"><span className="px-2 py-0.5 rounded font-semibold text-[10px] bg-purple-100 text-purple-800">{u.role}</span></td>
                        <td className="p-3.5 text-center">
                          {currentUser.role === 'Admin' && <button onClick={() => handleDeleteUser(u.id, u.role)} className="p-1 bg-rose-100 text-rose-700 rounded"><Trash2 size={14} /></button>}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {/* PENGATURAN */}
          {activeTab === 'Pengaturan' && (
            <div className={`p-6 rounded-xl border shadow-sm max-w-2xl space-y-4 ${darkMode ? 'bg-slate-800 border-slate-700' : 'bg-white border-slate-200'}`}>
              <h3 className="font-bold text-sm border-b pb-2 flex items-center gap-2 text-yellow-500"><Sliders size={18}/> Pengaturan Sistem</h3>
              <div className="space-y-2 text-xs">
                <label className="block font-semibold">Nama Perusahaan</label>
                <input type="text" disabled className={inputClass} value="PT. Ryan Eka Pratama" />
                <label className="block font-semibold">Lokasi Site</label>
                <input type="text" disabled className={inputClass} value="Site Bontang, Kalimantan Timur" />
              </div>
            </div>
          )}

          {/* MODAL INPUT: PRODUK BARU */}
          {showModalMaterial && (
            <div className="fixed inset-0 bg-black bg-opacity-60 flex items-center justify-center z-50 p-4">
              <div className={`w-full max-w-lg p-6 rounded-xl shadow-2xl border ${darkMode ? 'bg-slate-800 border-slate-700 text-white' : 'bg-white border-slate-200 text-slate-800'}`}>
                <div className="flex justify-between items-center mb-4 border-b pb-2"><h3 className="font-bold text-base">Tambah Produk Baru</h3><button onClick={() => setShowModalMaterial(false)}><X size={20}/></button></div>
                <form onSubmit={handleAddMaterial} className="space-y-3 text-xs">
                  <div><label className="block font-semibold mb-1">Kode Barang</label><input type="text" required className={inputClass} value={formMaterial.kode} onChange={e => setFormMaterial({...formMaterial, kode: e.target.value})}/></div>
                  <div><label className="block font-semibold mb-1">Nama Barang</label><input type="text" required className={inputClass} value={formMaterial.nama} onChange={e => setFormMaterial({...formMaterial, nama: e.target.value})}/></div>
                  <div><label className="block font-semibold mb-1">Nomor Part</label><input type="text" className={inputClass} value={formMaterial.nomor} onChange={e => setFormMaterial({...formMaterial, nomor: e.target.value})}/></div>
                  <div className="grid grid-cols-2 gap-2">
                    <div><label className="block font-semibold mb-1">Unit</label><input type="text" className={inputClass} value={formMaterial.unit} onChange={e => setFormMaterial({...formMaterial, unit: e.target.value})}/></div>
                    <div><label className="block font-semibold mb-1">Qty Awal</label><input type="number" required className={inputClass} value={formMaterial.qty} onChange={e => setFormMaterial({...formMaterial, qty: e.target.value})}/></div>
                  </div>
                  <div className="flex justify-end gap-2 mt-4 pt-2 border-t"><button type="button" onClick={() => setShowModalMaterial(false)} className="px-4 py-2 border rounded-lg">Batal</button><button type="submit" className="px-4 py-2 bg-emerald-600 text-white font-semibold rounded-lg">Simpan</button></div>
                </form>
              </div>
            </div>
          )}

          {/* MODAL INPUT: STOK MASUK */}
          {showModalMasuk && (
            <div className="fixed inset-0 bg-black bg-opacity-60 flex items-center justify-center z-50 p-4">
              <div className={`w-full max-w-lg p-6 rounded-xl shadow-2xl border ${darkMode ? 'bg-slate-800 border-slate-700 text-white' : 'bg-white border-slate-200 text-slate-800'}`}>
                <div className="flex justify-between items-center mb-4 border-b pb-2"><h3 className="font-bold text-base">Tambah Stok Masuk</h3><button onClick={() => setShowModalMasuk(false)}><X size={20}/></button></div>
                <form onSubmit={handleAddMasuk} className="space-y-3 text-xs">
                  <div className="grid grid-cols-2 gap-2">
                    <div><label className="block font-semibold mb-1">Tanggal</label><input type="date" required className={inputClass} value={formMasuk.tanggal} onChange={e => setFormMasuk({...formMasuk, tanggal: e.target.value})}/></div>
                    <div><label className="block font-semibold mb-1">Unit</label><input type="text" required className={inputClass} value={formMasuk.unit} onChange={e => setFormMasuk({...formMasuk, unit: e.target.value})}/></div>
                  </div>
                  <div className="grid grid-cols-2 gap-2">
                    <div><label className="block font-semibold mb-1">Kode Barang</label><input type="text" required className={inputClass} value={formMasuk.kode} onChange={e => setFormMasuk({...formMasuk, kode: e.target.value})}/></div>
                    <div><label className="block font-semibold mb-1">Merk</label><input type="text" className={inputClass} value={formMasuk.merk} onChange={e => setFormMasuk({...formMasuk, merk: e.target.value})}/></div>
                  </div>
                  <div><label className="block font-semibold mb-1">Deskripsi</label><input type="text" required className={inputClass} value={formMasuk.description} onChange={e => setFormMasuk({...formMasuk, description: e.target.value})}/></div>
                  <div className="grid grid-cols-3 gap-2">
                    <div><label className="block font-semibold mb-1">Qty</label><input type="number" required className={inputClass} value={formMasuk.qty} onChange={e => setFormMasuk({...formMasuk, qty: e.target.value})}/></div>
                    <div><label className="block font-semibold mb-1">Satuan</label><input type="text" required className={inputClass} value={formMasuk.satuan} onChange={e => setFormMasuk({...formMasuk, satuan: e.target.value})}/></div>
                    <div><label className="block font-semibold mb-1">No. PO</label><input type="text" className={inputClass} value={formMasuk.noPo} onChange={e => setFormMasuk({...formMasuk, noPo: e.target.value})}/></div>
                  </div>
                  <div className="flex justify-end gap-2 mt-4 pt-2 border-t"><button type="button" onClick={() => setShowModalMasuk(false)} className="px-4 py-2 border rounded-lg">Batal</button><button type="submit" className="px-4 py-2 bg-blue-600 text-white font-semibold rounded-lg">Simpan</button></div>
                </form>
              </div>
            </div>
          )}

          {/* MODAL INPUT: STOK KELUAR */}
          {showModalKeluar && (
            <div className="fixed inset-0 bg-black bg-opacity-60 flex items-center justify-center z-50 p-4">
              <div className={`w-full max-w-lg p-6 rounded-xl shadow-2xl border ${darkMode ? 'bg-slate-800 border-slate-700 text-white' : 'bg-white border-slate-200 text-slate-800'}`}>
                <div className="flex justify-between items-center mb-4 border-b pb-2"><h3 className="font-bold text-base">Tambah Stok Keluar</h3><button onClick={() => setShowModalKeluar(false)}><X size={20}/></button></div>
                <form onSubmit={handleAddKeluar} className="space-y-3 text-xs">
                  <div className="grid grid-cols-2 gap-2">
                    <div><label className="block font-semibold mb-1">Tanggal</label><input type="date" required className={inputClass} value={formKeluar.tanggal} onChange={e => setFormKeluar({...formKeluar, tanggal: e.target.value})}/></div>
                    <div><label className="block font-semibold mb-1">Unit</label><input type="text" required className={inputClass} value={formKeluar.unit} onChange={e => setFormKeluar({...formKeluar, unit: e.target.value})}/></div>
                  </div>
                  <div className="grid grid-cols-2 gap-2">
                    <div><label className="block font-semibold mb-1">Kode Barang</label><input type="text" required className={inputClass} value={formKeluar.kode} onChange={e => setFormKeluar({...formKeluar, kode: e.target.value})}/></div>
                    <div><label className="block font-semibold mb-1">Mekanik</label><input type="text" className={inputClass} value={formKeluar.mekanik} onChange={e => setFormKeluar({...formKeluar, mekanik: e.target.value})}/></div>
                  </div>
                  <div><label className="block font-semibold mb-1">Nama Barang</label><input type="text" required className={inputClass} value={formKeluar.namaBarang} onChange={e => setFormKeluar({...formKeluar, namaBarang: e.target.value})}/></div>
                  <div className="grid grid-cols-2 gap-2">
                    <div><label className="block font-semibold mb-1">Qty Keluar</label><input type="number" required className={inputClass} value={formKeluar.qty} onChange={e => setFormKeluar({...formKeluar, qty: e.target.value})}/></div>
                    <div><label className="block font-semibold mb-1">Keterangan</label><input type="text" className={inputClass} value={formKeluar.keterangan} onChange={e => setFormKeluar({...formKeluar, keterangan: e.target.value})}/></div>
                  </div>
                  <div className="flex justify-end gap-2 mt-4 pt-2 border-t"><button type="button" onClick={() => setShowModalKeluar(false)} className="px-4 py-2 border rounded-lg">Batal</button><button type="submit" className="px-4 py-2 bg-amber-600 text-white font-semibold rounded-lg">Simpan</button></div>
                </form>
              </div>
            </div>
          )}

          {/* MODAL PENGGUNA BARU */}
          {showModalUser && (
            <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50 p-4">
              <div className={`w-full max-w-md p-6 rounded-xl shadow-lg border ${darkMode ? 'bg-slate-800 border-slate-700 text-white' : 'bg-white border-slate-200 text-slate-800'}`}>
                <div className="flex justify-between items-center mb-4 border-b pb-2"><h3 className="font-bold text-base">Tambah Pengguna Baru</h3><button onClick={() => setShowModalUser(false)}><X size={20}/></button></div>
                <form onSubmit={handleAddUser} className="space-y-3 text-xs">
                  <div><label className="block font-semibold mb-1">Nama Lengkap</label><input type="text" required className={inputClass} value={formUser.nama} onChange={e => setFormUser({...formUser, nama: e.target.value})}/></div>
                  <div><label className="block font-semibold mb-1">Username</label><input type="text" required className={inputClass} value={formUser.username} onChange={e => setFormUser({...formUser, username: e.target.value})}/></div>
                  <div><label className="block font-semibold mb-1">Role</label><select className={selectClass} value={formUser.role} onChange={e => setFormUser({...formUser, role: e.target.value})}><option value="Admin">Admin</option><option value="Editor">Editor</option><option value="Viewer">Viewer</option></select></div>
                  <div className="flex justify-end gap-2 mt-4 pt-2 border-t"><button type="button" onClick={() => setShowModalUser(false)} className="px-4 py-2 border rounded-lg">Batal</button><button type="submit" className="px-4 py-2 bg-blue-600 text-white font-semibold rounded-lg">Simpan</button></div>
                </form>
              </div>
            </div>
          )}

        </main>
      </div>
    </div>
  );
}