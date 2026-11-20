/* ====================================================================
   Pembaca file Excel format "TANDA TERIMA SEPATU DAN SERAGAM" (sheet RESUME):
   blok per tahun berisi Tanda Terima Sepatu Terakhir, Jadwal Pembagian Sepatu I/II + Progress +
   Penggantian Sepatu Rusak, Tanda Terima Seragam Terakhir, Jadwal Pembagian Seragam + Progress.
   Fungsi di sini murni (tanpa React/Supabase) supaya mudah diuji.
   ==================================================================== */

const pad2 = (n) => String(n).padStart(2, '0');
const norm = (v) => String(v ?? '').trim().replace(/\s+/g, ' ').toLowerCase();
const teks = (v) => (v === null || v === undefined ? '' : String(v).trim());

// Nomor seri tanggal Excel -> 'YYYY-MM-DD'
export const serialToISO = (n) => {
  const d = new Date(Math.round((n - 25569) * 86400 * 1000));
  return `${d.getUTCFullYear()}-${pad2(d.getUTCMonth() + 1)}-${pad2(d.getUTCDate())}`;
};

// Nilai sel -> 'YYYY-MM-DD', null bila kosong, undefined bila ada isi tetapi bukan tanggal yang dikenali
export const toISO = (v) => {
  if (v === null || v === undefined || v === '') return null;
  if (v instanceof Date && !isNaN(v)) return `${v.getFullYear()}-${pad2(v.getMonth() + 1)}-${pad2(v.getDate())}`;
  if (typeof v === 'number') return v > 20000 && v < 80000 ? serialToISO(v) : undefined;
  const s = String(v).trim();
  if (!s) return null;
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return `${m[1]}-${pad2(m[2])}-${pad2(m[3])}`;
  m = s.match(/^(\d{1,2})[/.-](\d{1,2})[/.-](\d{4})$/);
  if (m) return `${m[3]}-${pad2(m[2])}-${pad2(m[1])}`;
  return undefined;
};

const cari = (header, label, from, to) => {
  for (let c = from; c < to; c++) if (norm(header[c]) === label) return c;
  return -1;
};

/* Membaca array-of-arrays (hasil sheet_to_json header:1). Mengembalikan:
   { karyawan: [...], riwayat: [...], masalah: [{ baris, pesan }], blok } */
export const parseRekap = (aoa) => {
  const masalah = [];
  // cari baris judul kolom
  let hr = -1;
  for (let r = 0; r < Math.min(aoa.length, 40); r++) {
    const row = aoa[r] || [];
    if (row.some((c) => norm(c) === 'nrp amodi') && row.some((c) => norm(c) === 'nama')) { hr = r; break; }
  }
  if (hr < 0) return { karyawan: [], riwayat: [], masalah: [{ baris: 0, pesan: 'Judul kolom "NRP Amodi" dan "Nama" tidak ditemukan. Pastikan memakai sheet RESUME dari file MONITORING SEPATU DAN SERAGAM.' }], blok: 0 };
  const header = aoa[hr];
  const col = {
    amodi: cari(header, 'nrp amodi', 0, header.length),
    nrp: cari(header, 'nrp', 0, header.length),
    nama: cari(header, 'nama', 0, header.length),
    jabatan: cari(header, 'jabatan', 0, header.length),
    masuk: cari(header, 'tanggal masuk', 0, header.length),
    ukSepatu: cari(header, 'ukuran sepatu', 0, header.length),
    ukSeragam: cari(header, 'ukuran seragam', 0, header.length),
  };
  // setiap blok tahun dimulai dari kolom "Tanda Terima Sepatu Terakhir"
  const starts = [];
  header.forEach((c, i) => { if (norm(c) === 'tanda terima sepatu terakhir') starts.push(i); });
  if (!starts.length) return { karyawan: [], riwayat: [], masalah: [{ baris: hr + 1, pesan: 'Kolom "Tanda Terima Sepatu Terakhir" tidak ditemukan.' }], blok: 0 };
  const bloks = starts.map((st, i) => {
    const en = i + 1 < starts.length ? starts[i + 1] : header.length + 1;
    const jad1 = cari(header, 'jadwal pembagian sepatu i', st, en);
    const jad2 = cari(header, 'jadwal pembagian sepatu ii', st, en);
    const jadSer = cari(header, 'jadwal pembagian seragam', st, en);
    return {
      lastSepatu: st,
      lastSeragam: cari(header, 'tanda terima seragam terakhir', st, en),
      sepatu: [jad1, jad2].filter((c) => c >= 0).map((c) => ({ jad: c, prog: c + 1, rusak: c + 2 })),
      jadSeragam: jadSer, progSeragam: jadSer >= 0 ? jadSer + 1 : -1,
      ket: cari(header, 'keterangan', st, en),
    };
  });

  const karyawan = []; const riwayat = []; const seen = new Set(); const rk = new Set();
  const tambah = (amodi, jenis, kategori, v, baris, label) => {
    const t = toISO(v);
    if (t === null) return;
    if (t === undefined) { masalah.push({ baris, pesan: `${label}: "${teks(v)}" bukan tanggal yang dikenali (dilewati).` }); return; }
    const key = `${amodi}|${jenis}|${kategori}|${t}`;
    if (rk.has(key)) return;
    rk.add(key);
    riwayat.push({ nrp_amodi: amodi, jenis, kategori, tanggal: t });
  };

  for (let r = hr + 1; r < aoa.length; r++) {
    const row = aoa[r] || [];
    const amodi = teks(row[col.amodi]); const nama = teks(row[col.nama]);
    if (!amodi && !nama) continue;
    const baris = r + 1;
    if (!amodi) { masalah.push({ baris, pesan: `${nama}: NRP Amodi kosong (dilewati).` }); continue; }
    if (!nama) { masalah.push({ baris, pesan: `${amodi}: Nama kosong (dilewati).` }); continue; }
    if (seen.has(amodi)) { masalah.push({ baris, pesan: `${amodi} (${nama}): NRP Amodi dobel dalam file, baris ini dilewati.` }); continue; }
    seen.add(amodi);
    const masuk = toISO(row[col.masuk]);
    if (masuk === undefined) masalah.push({ baris, pesan: `${nama}: Tanggal Masuk tidak dikenali (diabaikan).` });
    const ukS = teks(row[col.ukSepatu]);
    karyawan.push({
      nrp_amodi: amodi, nrp: teks(row[col.nrp]) || amodi, nama, jabatan: teks(row[col.jabatan]) || null,
      tanggal_masuk: masuk || null, ukuran_sepatu: ukS || null, ukuran_seragam: teks(row[col.ukSeragam]).toUpperCase() || null,
      keterangan: (bloks[0].ket >= 0 ? teks(row[bloks[0].ket]) : '') || null, baris,
    });
    bloks.forEach((b, bi) => {
      // Kolom "Tanda Terima ... Terakhir" hanya dipercaya pada blok tahun pertama. Pada blok tahun berikutnya
      // isinya proyeksi rumus (mis. jadwal yang belum berstatus SUDAH), bukan penerimaan yang sudah terjadi.
      if (bi === 0) tambah(amodi, 'sepatu', 'rutin', row[b.lastSepatu], baris, `${nama}, Tanda Terima Sepatu Terakhir`);
      b.sepatu.forEach((s, i) => {
        // jadwal yang berstatus SUDAH berarti pembagiannya sudah dilakukan pada tanggal jadwal itu
        if (teks(row[s.prog]).toUpperCase() === 'SUDAH') tambah(amodi, 'sepatu', 'rutin', row[s.jad], baris, `${nama}, Jadwal Sepatu ${i === 0 ? 'I' : 'II'}`);
        tambah(amodi, 'sepatu', 'rusak', row[s.rusak], baris, `${nama}, Penggantian Sepatu Rusak`);
      });
      if (bi === 0 && b.lastSeragam >= 0) tambah(amodi, 'seragam', 'rutin', row[b.lastSeragam], baris, `${nama}, Tanda Terima Seragam Terakhir`);
      if (b.jadSeragam >= 0 && teks(row[b.progSeragam]).toUpperCase() === 'SUDAH') tambah(amodi, 'seragam', 'rutin', row[b.jadSeragam], baris, `${nama}, Jadwal Seragam`);
    });
  }
  return { karyawan, riwayat, masalah, blok: bloks.length };
};
