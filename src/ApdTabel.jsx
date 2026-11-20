import React, { useState, useRef, useEffect } from 'react';
import { Filter, ArrowUp, ArrowDown, ArrowUpDown } from 'lucide-react';

/* ====================================================================
   Urut, filter, dan cari pada header kolom tabel (gaya filter Excel).
   Setiap kolom: { key, label, ambil(row), tampil?(nilai), urut?(row), kiri?, kanan? }
   - Klik judul kolom: urut naik, turun, lalu batal.
   - Ikon corong: panel berisi urutan, kotak cari, dan daftar nilai yang bisa dicentang.
   ==================================================================== */

const kosong = (v) => v === null || v === undefined || v === '';
const tampilDasar = (c, v) => (c.tampil ? c.tampil(v) : kosong(v) ? '(kosong)' : String(v));
const nilaiUrut = (c, r) => (c.urut ? c.urut(r) : c.ambil(r));
const banding = (x, y) => {
  if (typeof x === 'number' && typeof y === 'number') return x - y;
  return String(x).localeCompare(String(y), 'id', { numeric: true });
};

const URUTAN_SERAGAM = ['XS', 'S', 'M', 'L', 'XL', 'XXL', '2XL', '3XL', 'XXXL', '4XL'];
// Peringkat ukuran agar urutannya wajar: sepatu menurut angka, seragam S, M, L, XL, dst.
export const rankUkuran = (jenis, u) => {
  const t = String(u || '').trim().toUpperCase();
  if (!t || t === '(BELUM DIISI)') return 1e6;
  if (jenis === 'sepatu') { const n = parseFloat(t); return Number.isNaN(n) ? 1e5 : n; }
  const i = URUTAN_SERAGAM.indexOf(t);
  return i < 0 ? 1e5 : i;
};

/* ---------- state urut dan filter satu tabel ---------- */
export const useTabel = (kolom) => {
  const [sort, setSort] = useState(null);   // { key, dir: 'asc' | 'desc' } atau null
  const [filter, setFilter] = useState({}); // key -> { teks, pilih: null | string[] }
  const def = {};
  kolom.forEach((c) => { def[c.key] = c; });

  const geserSort = (key) => setSort((s) => (!s || s.key !== key ? { key, dir: 'asc' } : s.dir === 'asc' ? { key, dir: 'desc' } : null));
  const aturSort = (key, dir) => setSort(dir ? { key, dir } : null);
  const ubahFilter = (key, patch) => setFilter((f) => {
    const n = { teks: '', pilih: null, ...(f[key] || {}), ...patch };
    const baru = { ...f };
    if (!(n.teks || '').trim() && !n.pilih) delete baru[key]; else baru[key] = n;
    return baru;
  });
  const hapusKolom = (key) => setFilter((f) => { const baru = { ...f }; delete baru[key]; return baru; });
  const reset = () => { setSort(null); setFilter({}); };
  const aktif = !!sort || Object.keys(filter).length > 0;

  // terapkan filter lalu urutan ke daftar baris (urutan dasar dipertahankan bila tidak ada urutan kolom)
  const terapkan = (rows) => {
    let out = rows;
    Object.entries(filter).forEach(([key, f]) => {
      const c = def[key];
      if (!c) return;
      const teks = (f.teks || '').trim().toLowerCase();
      const pilih = f.pilih ? new Set(f.pilih) : null;
      out = out.filter((r) => {
        const t = tampilDasar(c, c.ambil(r));
        return (!teks || t.toLowerCase().includes(teks)) && (!pilih || pilih.has(t));
      });
    });
    if (sort && def[sort.key]) {
      const c = def[sort.key]; const m = sort.dir === 'desc' ? -1 : 1;
      out = out.slice().sort((a, b) => {
        const x = nilaiUrut(c, a); const y = nilaiUrut(c, b);
        const kx = kosong(x); const ky = kosong(y);
        if (kx || ky) return kx && ky ? 0 : kx ? 1 : -1; // yang kosong selalu di bawah
        return m * banding(x, y);
      });
    }
    return out;
  };
  return { sort, filter, geserSort, aturSort, ubahFilter, hapusKolom, reset, aktif, terapkan };
};

/* ---------- tombol reset ---------- */
export const BtnReset = ({ tabel }) => (tabel.aktif
  ? <button type="button" onClick={tabel.reset} className="px-2.5 py-1.5 rounded-lg bg-yellow-400 text-slate-900 text-xs font-bold whitespace-nowrap">Reset filter dan urutan</button>
  : null);

/* ---------- panel filter (melayang, tidak terpotong tabel yang bisa digeser) ---------- */
function Panel({ c, tabel, rows, pos, tutup, tombol, dark }) {
  const ref = useRef(null);
  useEffect(() => {
    const klikLuar = (e) => {
      if (ref.current && ref.current.contains(e.target)) return;
      if (tombol.current && tombol.current.contains(e.target)) return;
      tutup();
    };
    const gulir = (e) => { if (ref.current && ref.current.contains(e.target)) return; tutup(); };
    const esc = (e) => { if (e.key === 'Escape') tutup(); };
    document.addEventListener('mousedown', klikLuar);
    window.addEventListener('scroll', gulir, true);
    window.addEventListener('resize', tutup);
    document.addEventListener('keydown', esc);
    return () => {
      document.removeEventListener('mousedown', klikLuar);
      window.removeEventListener('scroll', gulir, true);
      window.removeEventListener('resize', tutup);
      document.removeEventListener('keydown', esc);
    };
  }, []); // eslint-disable-line

  const f = tabel.filter[c.key] || { teks: '', pilih: null };
  const peta = new Map();
  rows.forEach((r) => { const t = tampilDasar(c, c.ambil(r)); if (!peta.has(t)) peta.set(t, nilaiUrut(c, r)); });
  const semua = Array.from(peta.entries()).sort((a, b) => {
    const kx = kosong(a[1]); const ky = kosong(b[1]);
    if (kx || ky) return kx && ky ? 0 : kx ? 1 : -1;
    return banding(a[1], b[1]);
  }).map((e) => e[0]);
  const teks = (f.teks || '').trim().toLowerCase();
  const terlihat = semua.filter((t) => !teks || t.toLowerCase().includes(teks));
  const terpilih = f.pilih ? new Set(f.pilih) : null;
  const toggleNilai = (t) => {
    const cur = terpilih ? Array.from(terpilih) : semua;
    const next = cur.includes(t) ? cur.filter((x) => x !== t) : [...cur, t];
    tabel.ubahFilter(c.key, { pilih: next.length === semua.length ? null : next });
  };
  const urutan = tabel.sort && tabel.sort.key === c.key ? tabel.sort.dir : null;
  const btn = 'px-2 py-1 rounded border text-[11px] font-semibold';
  const btnAktif = 'bg-yellow-400 text-slate-900 border-yellow-500';
  const btnBiasa = dark ? 'border-slate-500 hover:bg-slate-700' : 'border-slate-300 hover:bg-slate-100';

  return (
    <div ref={ref} onClick={(e) => e.stopPropagation()}
      style={{ position: 'fixed', left: pos.left, top: pos.top, width: 256, zIndex: 60, maxHeight: '75vh', overflowY: 'auto', textAlign: 'left', fontWeight: 400, textTransform: 'none', letterSpacing: 'normal', whiteSpace: 'normal' }}
      className={`rounded-lg shadow-2xl border p-3 text-xs ${dark ? 'bg-slate-800 border-slate-600 text-slate-100' : 'bg-white border-slate-300 text-slate-800'}`}>
      <div className="font-bold mb-2">{c.label.replace('\n', ' ')}</div>
      <div className="flex gap-1 mb-2">
        <button type="button" className={`${btn} ${urutan === 'asc' ? btnAktif : btnBiasa}`} onClick={() => tabel.aturSort(c.key, urutan === 'asc' ? null : 'asc')}>Urut naik</button>
        <button type="button" className={`${btn} ${urutan === 'desc' ? btnAktif : btnBiasa}`} onClick={() => tabel.aturSort(c.key, urutan === 'desc' ? null : 'desc')}>Urut turun</button>
      </div>
      <input autoFocus type="text" placeholder="Cari di kolom ini..." value={f.teks || ''} onChange={(e) => tabel.ubahFilter(c.key, { teks: e.target.value })}
        className="w-full px-2 py-1.5 mb-1 border border-slate-300 rounded bg-white text-slate-900 placeholder-slate-400 focus:outline-none focus:border-blue-500" />
      <div className="flex justify-between my-1">
        <button type="button" className="text-blue-500 font-semibold" onClick={() => tabel.ubahFilter(c.key, { pilih: null })}>Pilih semua</button>
        <button type="button" className="text-blue-500 font-semibold" onClick={() => tabel.ubahFilter(c.key, { pilih: [] })}>Kosongkan</button>
      </div>
      <div className={`border rounded overflow-y-auto ${dark ? 'border-slate-600' : 'border-slate-200'}`} style={{ maxHeight: 190 }}>
        {terlihat.length === 0 ? <div className="p-2 text-slate-400">Tidak ada nilai yang cocok</div> : terlihat.slice(0, 300).map((t) => (
          <label key={t} className={`flex items-center gap-2 px-2 py-1 cursor-pointer ${dark ? 'hover:bg-slate-700' : 'hover:bg-slate-100'}`}>
            <input type="checkbox" checked={!terpilih || terpilih.has(t)} onChange={() => toggleNilai(t)} />
            <span className="truncate">{t}</span>
          </label>
        ))}
      </div>
      {terlihat.length > 300 && <div className="mt-1 text-slate-400">Menampilkan 300 nilai pertama, persempit lewat kotak cari.</div>}
      <div className="flex justify-between mt-2">
        <button type="button" className="text-rose-500 font-semibold" onClick={() => tabel.hapusKolom(c.key)}>Hapus filter kolom</button>
        <button type="button" className="font-semibold" onClick={tutup}>Tutup</button>
      </div>
    </div>
  );
}

/* ---------- sel judul kolom: urut + filter/cari ---------- */
export function Th({ c, tabel, rows, dark, kiri, kanan }) {
  const [buka, setBuka] = useState(false);
  const [pos, setPos] = useState({ left: 0, top: 0 });
  const tombol = useRef(null);
  const left = kiri || c.kiri; const right = kanan || c.kanan;
  const filterAktif = !!tabel.filter[c.key];
  const urutan = tabel.sort && tabel.sort.key === c.key ? tabel.sort.dir : null;
  const toggle = () => {
    if (!buka && tombol.current) {
      const r = tombol.current.getBoundingClientRect();
      setPos({ left: Math.max(8, Math.min(r.left, window.innerWidth - 268)), top: r.bottom + 4 });
    }
    setBuka(!buka);
  };
  return (
    <th className={`p-2.5 ${left ? 'text-left' : right ? 'text-right' : ''}`}>
      <div className={`flex items-center gap-1 ${left ? 'justify-start' : right ? 'justify-end' : 'justify-center'}`}>
        <button type="button" onClick={() => tabel.geserSort(c.key)} title="Urutkan (klik berulang: naik, turun, batal)"
          className="inline-flex items-center gap-1 font-semibold leading-tight" style={{ whiteSpace: 'pre-line', textAlign: left ? 'left' : right ? 'right' : 'center' }}>
          <span>{c.label}</span>
          {urutan === 'asc' ? <ArrowUp size={12} className="text-yellow-500 shrink-0" /> : urutan === 'desc' ? <ArrowDown size={12} className="text-yellow-500 shrink-0" /> : <ArrowUpDown size={12} className="opacity-40 shrink-0" />}
        </button>
        <button ref={tombol} type="button" onClick={toggle} title="Filter dan cari di kolom ini"
          className={`p-0.5 rounded shrink-0 ${filterAktif ? 'bg-yellow-400 text-slate-900' : 'opacity-60 hover:opacity-100'}`}><Filter size={12} /></button>
      </div>
      {buka && <Panel c={c} tabel={tabel} rows={rows} pos={pos} tutup={() => setBuka(false)} tombol={tombol} dark={dark} />}
    </th>
  );
}

/* ---------- tabel rekap per ukuran (dipakai di Orderan dan Rekap) ---------- */
export function TabelRekapUkuran({ judul, rows, satuan, jenis, kelas }) {
  const kolom = [
    { key: 'u', label: 'UKURAN', kiri: true, ambil: (r) => r.u, urut: (r) => rankUkuran(jenis, r.u) },
    { key: 'n', label: `JUMLAH\n(${satuan})`, kanan: true, ambil: (r) => r.n },
  ];
  const tabel = useTabel(kolom);
  const dasar = rows.map(([u, n]) => ({ u, n }));
  const tampil = tabel.terapkan(dasar);
  const total = tampil.reduce((a, r) => a + r.n, 0);
  return (
    <div className={`rounded-xl shadow-sm border overflow-hidden ${kelas.card}`}>
      <div className="p-3 border-b font-bold text-sm flex flex-wrap items-center justify-between gap-2"><span>{judul}</span><BtnReset tabel={tabel} /></div>
      <table className="w-full text-sm">
        <thead className={`border-b font-semibold ${kelas.thead}`}><tr>{kolom.map((c) => <Th key={c.key} c={c} tabel={tabel} rows={dasar} dark={kelas.dark} />)}</tr></thead>
        <tbody className={`divide-y ${kelas.divide}`}>
          {tampil.length === 0
            ? <tr><td colSpan="2" className="p-4 text-center text-slate-400">-</td></tr>
            : tampil.map((r) => <tr key={r.u}><td className="p-2 font-bold">{r.u}</td><td className="p-2 text-right font-semibold">{r.n}</td></tr>)}
          <tr className="font-bold"><td className="p-2">TOTAL{tabel.aktif ? ' (terfilter)' : ''}</td><td className="p-2 text-right">{total}</td></tr>
        </tbody>
      </table>
    </div>
  );
}
