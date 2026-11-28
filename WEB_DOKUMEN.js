/* ======================================================
   WEB_DOKUMEN.js   (dipakai web SL & web SANFELO, isi file sama)
   ------------------------------------------------------
   Halaman daftar + detail nota (seperti web Laurence):
   - webDokumen_      : isi 1 nota (baris-baris sheet sumbernya) + riwayat revisi
   - REVISI_NOTA      : salinan isi nota setiap kali disimpan dari web
                        (versi 1 = "Asli"). Hanya MENAMBAH baris di sheet
                        REVISI_NOTA; sheet transaksi tidak disentuh.
   - webRingkasPO_    : ringkasan PO masuk & PO keluar untuk dashboard
   ====================================================== */

// sumber tiap jenis nota: sheet + kolom nomor (nama header)
const WEB_DOK_ = {
  PO: [{ sheet: 'HISTORY_PO_MASUK', kunci: ['NO PO'], saring: function (r, h) { return String(r[h['JENIS']] || '').trim().toUpperCase() === 'MASUK'; } }],
  POK: [{ sheet: 'HISTORY_PO_KELUAR', kunci: ['NO PO'] }],
  // SJ & SPR juga bisa dibuka dari nomor invoice (dari Pembayaran Piutang / Laporan); nomor resmi = NO SJ
  SJ: [{ sheet: 'TRANSAKSI', kunci: ['NO SJ', 'NO INV'] }, { sheet: 'DETAIL_TRANSAKSI', kunci: ['NO SJ'], dari: 'NO SJ' }, { sheet: 'PENJUALAN', kunci: ['NO INV'], dari: 'NO INV' }],
  SPR: [{ sheet: 'SPAREPART', kunci: ['NO SJ', 'NO SPR', 'NO INV'] }],
  TT: [{ sheet: 'PEMBELIAN', kunci: ['NO PB'] }]
};
const WEB_REVISI_SHEET_ = 'REVISI_NOTA';
const WEB_REVISI_KOLOM_ = ['WAKTU', 'JENIS', 'NOMOR', 'VERSI', 'ADMIN', 'ALASAN', 'TOTAL', 'DATA'];

function webNilaiTampil_(v) {
  if (v instanceof Date) return webTgl_(v);
  return (v && typeof v === 'object') ? '' : v;
}

// baris-baris 1 sheet yang kolom kuncinya = nomor
function webBarisNota_(sumber, nomor) {
  const sh = webSS_().getSheetByName(sumber.sheet);
  if (!sh || sh.getLastRow() < 1) return { header: [], rows: [] };
  const lebar = Math.max(1, sh.getLastColumn());
  const semua = sh.getRange(1, 1, sh.getLastRow(), lebar).getValues();
  const header = semua[0].map(function (x, i) { return String(x || '').trim() || ('KOLOM ' + (i + 1)); });
  const h = {};
  header.forEach(function (x, i) { if (h[x] === undefined) h[x] = i; });
  const ks = sumber.kunci.filter(function (n) { return h[n] !== undefined; }).map(function (n) { return h[n]; });
  if (!ks.length) return { header: header, rows: [] };
  const target = String(nomor).trim().toUpperCase();
  const rows = [];
  for (let i = 1; i < semua.length; i++) {
    const r = semua[i];
    if (!ks.some(function (k) { return String(r[k] || '').trim().toUpperCase() === target; })) continue;
    if (sumber.saring && !sumber.saring(r, h)) continue;
    rows.push(r.map(webNilaiTampil_));
  }
  // buang kolom kosong di ujung kanan header
  let akhir = header.length;
  while (akhir > 1 && /^KOLOM \d+$/.test(header[akhir - 1]) && rows.every(function (r) { return r[akhir - 1] === '' || r[akhir - 1] === null; })) akhir--;
  return { header: header.slice(0, akhir), rows: rows.map(function (r) { return r.slice(0, akhir); }) };
}

// isi lengkap 1 nota: { jenis, nomor, data: { SHEET: { header, rows } } }
function webIsiNota_(jenis, nomor) {
  const cfg = WEB_DOK_[jenis];
  if (!cfg) throw new Error('Jenis nota tidak dikenal.');
  const data = {};
  let nomorLain = {};
  let resmi = String(nomor).trim();
  cfg.forEach(function (s, urutan) {
    let no = nomor;
    if (s.dari) {
      no = nomorLain[s.dari];
      if (!no) { data[s.sheet] = { header: [], rows: [] }; return; }
    }
    const b = webBarisNota_(s, no);
    data[s.sheet] = b;
    // nomor resmi nota = kolom kunci pertama sheet utama (mis. dibuka dari NO INV -> NO SJ)
    if (urutan === 0 && b.rows.length) { const ik = b.header.indexOf(s.kunci.filter(function (n) { return b.header.indexOf(n) !== -1; })[0]); if (ik !== -1 && b.rows[0][ik]) resmi = String(b.rows[0][ik]).trim(); }
    if (b.rows.length) b.header.forEach(function (x, i) { if (nomorLain[x] === undefined && b.rows[0][i]) nomorLain[x] = String(b.rows[0][i]).trim(); });
  });
  return { jenis: jenis, nomor: resmi, data: data };
}

function webTotalNota_(isi) {
  const ambil = function (sheet, kolom) {
    const b = isi.data[sheet]; if (!b) return 0;
    const i = b.header.indexOf(kolom); if (i === -1) return 0;
    return b.rows.reduce(function (s, r) { return s + (Number(r[i]) || 0); }, 0);
  };
  switch (isi.jenis) {
    case 'PO': return ambil('HISTORY_PO_MASUK', 'TOTAL');
    case 'POK': return ambil('HISTORY_PO_KELUAR', 'JUMLAH');
    case 'SJ': return ambil('TRANSAKSI', 'TOTAL');
    case 'SPR': return ambil('SPAREPART', 'HARGA');
    case 'TT': return ambil('PEMBELIAN', 'TOTAL');
  }
  return 0;
}

function webAdaIsiNota_(isi) {
  return Object.keys(isi.data).some(function (k) { return isi.data[k].rows.length > 0; });
}

/* ---------------- REVISI_NOTA ---------------- */

function webRevisiSheet_(buat) {
  const ss = webSS_();
  let sh = ss.getSheetByName(WEB_REVISI_SHEET_);
  if (!sh && buat) {
    sh = ss.insertSheet(WEB_REVISI_SHEET_);
    sh.getRange(1, 1, 1, WEB_REVISI_KOLOM_.length).setValues([WEB_REVISI_KOLOM_]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

// semua revisi 1 nota, urut versi. denganData=true -> ikut isi nota versi itu
function webDaftarRevisi_(jenis, nomor, denganData) {
  const sh = webRevisiSheet_(false);
  if (!sh || sh.getLastRow() < 2) return [];
  const target = String(nomor).trim().toUpperCase();
  const list = [];
  sh.getRange(2, 1, sh.getLastRow() - 1, WEB_REVISI_KOLOM_.length).getValues().forEach(function (r) {
    if (String(r[1]).trim() !== jenis || String(r[2]).trim().toUpperCase() !== target) return;
    const x = { waktu: r[0] instanceof Date ? Utilities.formatDate(r[0], WEB_TZ, 'dd/MM/yyyy HH:mm') : String(r[0]), versi: Number(r[3]) || 0, admin: String(r[4] || ''), alasan: String(r[5] || ''), total: Number(r[6]) || 0 };
    if (denganData) { try { x.isi = JSON.parse(String(r[7] || '')); } catch (err) { x.isi = null; } }
    list.push(x);
  });
  list.sort(function (a, b) { return a.versi - b.versi; });
  return list;
}

function webCatatRevisi_(isi, admin, alasan) {
  const sh = webRevisiSheet_(true);
  const lama = webDaftarRevisi_(isi.jenis, isi.nomor, false);
  const versi = lama.length ? lama[lama.length - 1].versi + 1 : 1;
  let json = JSON.stringify(isi);
  if (json.length > 49000) json = JSON.stringify({ jenis: isi.jenis, nomor: isi.nomor, terlaluBesar: true });
  sh.appendRow([new Date(), isi.jenis, isi.nomor, versi, admin || '', alasan || '', webTotalNota_(isi), json]);
  return versi;
}

// aksi tulis web yang menyimpan revisi: jenis, nomor sebelum simpan (ubah), nomor sesudah simpan
const WEB_REVISI_AKSI_ = {
  simpanPO: { jenis: 'PO', lama: function () { return ''; }, baru: function (a, h) { return h && h.poNo; } },
  ubahPO: { jenis: 'PO', lama: function (a) { return a.poNo; }, baru: function (a, h) { return h && h.poNo; } },
  simpanSJ: { jenis: 'SJ', lama: function (a) { return a.edit && a.edit.noSJ; }, baru: function (a, h) { return h && h.noSJ; } },
  simpanSparepart: { jenis: 'SPR', lama: function (a) { return a.edit && a.edit.noSJ; }, baru: function (a, h) { return h && h.noSJ; } },
  simpanTT: { jenis: 'TT', lama: function (a) { return a.editPB && a.editPB.noPB; }, baru: function (a, h) { return h && h.noPB; } },
  poKeluarSimpan: { jenis: 'POK', lama: function (a) { return a.poNo; }, baru: function (a, h) { return h && h.poNo; } }
};

// dipanggil webJalankanAksi_ SEBELUM aksi tulis: nota lama yang belum punya riwayat dicatat dulu sebagai "Asli"
function webRevisiSebelum_(aksi, args, admin) {
  const c = WEB_REVISI_AKSI_[aksi];
  if (!c) return;
  try {
    const no = String(c.lama(args) || '').trim();
    if (!no || webDaftarRevisi_(c.jenis, no, false).length) return;
    const isi = webIsiNota_(c.jenis, no);
    if (webAdaIsiNota_(isi)) webCatatRevisi_(isi, '', '(isi sebelum revisi pertama dari web)');
  } catch (err) {
    Logger.log('Revisi (sebelum) gagal: ' + err);
  }
}

// dipanggil webJalankanAksi_ SESUDAH aksi tulis berhasil
function webRevisiSesudah_(aksi, args, hasil, admin) {
  const c = WEB_REVISI_AKSI_[aksi];
  if (!c) return;
  try {
    const no = String(c.baru(args, hasil) || '').trim();
    if (!no) return;
    const isi = webIsiNota_(c.jenis, no);
    if (webAdaIsiNota_(isi)) webCatatRevisi_(isi, admin, String(args.alasanRevisi || '').trim());
  } catch (err) {
    Logger.log('Revisi (sesudah) gagal: ' + err);
  }
}

/** aksi "dokumen": a = { jenis, nomor, versi? } */
function webDokumen_(a) {
  const jenis = String(a.jenis || '').trim().toUpperCase();
  const nomor = String(a.nomor || '').trim();
  if (!WEB_DOK_[jenis]) throw new Error('Jenis nota tidak dikenal.');
  if (!nomor) throw new Error('Nomor nota kosong.');
  let revisi = webDaftarRevisi_(jenis, nomor, !!a.versi);
  let isi;
  if (a.versi) {
    const v = revisi.filter(function (x) { return x.versi === Number(a.versi); })[0];
    if (!v || !v.isi || v.isi.terlaluBesar) throw new Error('Versi ' + a.versi + ' tidak bisa dibuka.');
    isi = v.isi;
  } else {
    isi = webIsiNota_(jenis, nomor);
    if (!webAdaIsiNota_(isi)) throw new Error(nomor + ' tidak ditemukan.');
    if (isi.nomor !== nomor) revisi = webDaftarRevisi_(jenis, isi.nomor, false); // dibuka dari nomor invoice
  }
  return {
    isi: isi, versi: a.versi ? Number(a.versi) : 0,
    revisi: revisi.map(function (x) { return { waktu: x.waktu, versi: x.versi, admin: x.admin, alasan: x.alasan, total: x.total }; })
  };
}

/* ---------------- ringkasan PO untuk dashboard ---------------- */

function webRingkasPO_() {
  const bulanIni = Utilities.formatDate(new Date(), WEB_TZ, 'yyyy-MM');
  const masuk = { terbuka: 0, nilaiSisa: 0, bulanIni: 0, barangSisa: 0 };
  const poM = {};
  webBaca_('HISTORY_PO_MASUK', 14).forEach(function (r) {
    if (String(r[0] || '').trim().toUpperCase() !== 'MASUK') return;
    const no = String(r[1] || '').trim(); if (!no) return;
    const p = poM[no] = poM[no] || { sisa: 0, nilai: 0, batal: true, bulan: webBulanDari_(r[2]) };
    const status = String(r[11] || '').trim().toUpperCase();
    if (status === 'CANCELLED') return;
    p.batal = false;
    const qty = Number(r[10]) || 0;
    let kirim = 0;
    if (status.indexOf('DONE SJ') === 0) { const n = Number(status.replace('DONE SJ', '').trim()); kirim = n > 0 ? n : qty; }
    const sisa = Math.max(0, qty - kirim);
    p.sisa += sisa;
    if (qty > 0) p.nilai += (Number(r[13]) || 0) * sisa / qty;
  });
  Object.keys(poM).forEach(function (no) {
    const p = poM[no];
    if (p.batal) return;
    if (p.bulan === bulanIni) masuk.bulanIni++;
    if (p.sisa > 0) { masuk.terbuka++; masuk.nilaiSisa += p.nilai; masuk.barangSisa += p.sisa; }
  });
  const keluar = { terbuka: 0, barangBelum: 0, bulanIni: 0 };
  const poK = {};
  webBaca_('HISTORY_PO_KELUAR', 8).forEach(function (r) {
    const no = String(r[0] || '').trim(); if (!no) return;
    const p = poK[no] = poK[no] || { belum: 0, bulan: webBulanDari_(r[1]) };
    p.belum += Math.max(0, (Number(r[4]) || 0) - (Number(r[7]) || 0));
  });
  Object.keys(poK).forEach(function (no) {
    const p = poK[no];
    if (p.bulan === bulanIni) keluar.bulanIni++;
    if (p.belum > 0) { keluar.terbuka++; keluar.barangBelum += p.belum; }
  });
  masuk.nilaiSisa = Math.round(masuk.nilaiSisa);
  return { masuk: masuk, keluar: keluar };
}

/** dipanggil dari layar login (sebelum masuk): nama-nama akun admin untuk dipilih */
function webNamaAdmin() {
  const urut = ['NUR', 'SILVA'];
  return webDaftarAdmin_().sort(function (a, b) {
    const ia = a === 'ADMIN' ? 99 : (urut.indexOf(a) === -1 ? 50 : urut.indexOf(a));
    const ib = b === 'ADMIN' ? 99 : (urut.indexOf(b) === -1 ? 50 : urut.indexOf(b));
    return ia - ib || (a < b ? -1 : 1);
  });
}

/**
 * PDF dari HTML sederhana buatan web (cetak pengajuan hutang, rekap invoice / SJ), langsung diunduh
 * seperti PDF nota lain. Tidak membaca / menulis sheet dan tidak disimpan ke Drive.
 */
function webPdfHtml_(a) {
  const html = String(a.html || '');
  if (!html || html.length > 500000) throw new Error('Isi PDF kosong atau terlalu besar.');
  const nama = String(a.nama || 'dokumen').replace(/[\\\/:*?"<>|]/g, '').trim().slice(0, 120) || 'dokumen';
  const pdf = Utilities.newBlob(html, 'text/html', 'dokumen.html').getAs('application/pdf');
  return { nama: (WEB_LATIHAN_ ? 'LATIHAN - ' : '') + nama + '.pdf', pdf: Utilities.base64Encode(pdf.getBytes()) };
}
