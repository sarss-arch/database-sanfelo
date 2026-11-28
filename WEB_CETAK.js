/* ======================================================
   WEB_CETAK.js   (PROJECT: database_SANFELO)
   ------------------------------------------------------
   Cetak nota PERSIS seperti di inventory_SANFELO: Surat Jalan, Invoice,
   SJ Sparepart, Invoice Sparepart, Tanda Terima Barang, Invoice Palet.

   Caranya: web memakai "FILE CETAK" = salinan utuh spreadsheet
   inventory_SANFELO (template, logo, font, border, rumus sama persis).
   Form di file itu diisi oleh FUNGSI INVENTORY YANG SAMA (cetak ulang SJ,
   buka draft sparepart, isi form tanda terima, isi invoice palet), lalu
   diekspor jadi PDF dengan pengaturan export inventory (A4, potret, tanpa
   garis). PDF TIDAK disimpan ke Drive: dikirim ke browser dan diunduh ke
   komputer admin. inventory_SANFELO asli tidak disentuh.

   SETUP (pemilik, sekali -- dan ulangi kalau template nota di inventory diubah):
   jalankan webSiapkanFileCetak dari editor.
   ====================================================== */

const WEB_PROP_CETAK = 'WEB_FILE_CETAK_ID';
const WEB_SHEET_CETAK_ = ['PO_KELUAR', 'SURAT JALAN', 'INVOICE', 'SJ_SPR', 'INV_SPR', 'TANDA TERIMA BRG', 'INVOICE PALET', 'SETTING', 'SYSTEM'];
var WEB_CETAK_SS_ = null; // diisi selama proses cetak (dipakai WEB_INV_SALINAN.js)

function webSiapkanFileCetak() {
  const props = PropertiesService.getScriptProperties();
  const lama = props.getProperty(WEB_PROP_CETAK);
  const salinan = DriveApp.getFileById(INVENTORY_SS_ID_).makeCopy('SANFELO CETAK WEB - JANGAN DIUBAH');
  props.setProperty(WEB_PROP_CETAK, salinan.getId());
  if (lama) {
    try { DriveApp.getFileById(lama).setTrashed(true); } catch (err) { /* sudah tidak ada */ }
  }
  Logger.log('File cetak siap: ' + salinan.getUrl());
}

function webFileCetak_() {
  const id = PropertiesService.getScriptProperties().getProperty(WEB_PROP_CETAK);
  if (!id) throw new Error('File cetak belum disiapkan. Pemilik perlu menjalankan webSiapkanFileCetak sekali dari editor Apps Script.');
  return SpreadsheetApp.openById(id);
}

// DATA_BERSIH & DISKON di file cetak disamakan dulu dengan inventory (dipakai rumus nama barang)
function webSegarkanMasterCetak_(ss) {
  WEB_SHEET_MASTER_.forEach(function (nama) {
    const src = webMaster_(nama);
    const dst = ss.getSheetByName(nama);
    if (!dst) return;
    const lr = Math.max(src.getLastRow(), 1);
    const lc = Math.max(src.getLastColumn(), 1);
    if (dst.getMaxRows() < lr) dst.insertRowsAfter(dst.getMaxRows(), lr - dst.getMaxRows());
    if (dst.getMaxColumns() < lc) dst.insertColumnsAfter(dst.getMaxColumns(), lc - dst.getMaxColumns());
    dst.getRange(1, 1, Math.max(dst.getLastRow(), 1), Math.max(dst.getLastColumn(), 1)).clearContent();
    dst.getRange(1, 1, lr, lc).setValues(src.getRange(1, 1, lr, lc).getValues());
  });
}

// export persis seperti exportSJPDF / exportInvoicePDF di inventory
function webEksporPdf_(ss, sh, barisAkhir, kolomAkhir, namaFile, jenis, margin) {
  const m = margin || 0.2;
  const area = sh.getRange(1, 1, barisAkhir, kolomAkhir);
  area.setBackground('#ffffff'); // inventory memutihkan latar sebelum export
  SpreadsheetApp.flush();
  const url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() +
    '/export?format=pdf&gid=' + sh.getSheetId() + '&size=a4&portrait=true&fitw=true' +
    '&gridlines=false&printtitle=false&sheetnames=false' +
    '&top_margin=' + m + '&bottom_margin=' + m + '&left_margin=' + m + '&right_margin=' + m +
    '&horizontal_alignment=CENTER&r1=0&c1=0&r2=' + barisAkhir + '&c2=' + kolomAkhir;
  const resp = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
  if (resp.getResponseCode() !== 200) throw new Error('Export PDF gagal (HTTP ' + resp.getResponseCode() + ').');
  const nama = (WEB_LATIHAN_ ? 'LATIHAN - ' : '') + namaFile.replace(/[\\\/:*?"<>|]/g, '').trim() + '.pdf';
  // tidak disimpan ke Drive: isi PDF dikirim ke browser lalu diunduh ke komputer
  return { nama: nama, pdf: Utilities.base64Encode(resp.getBlob().getBytes()) };
}

// Date -> Date yang tampil dengan tanggal WIB yang sama di spreadsheet ss (zona waktunya bisa beda)
function webTglUntukSheet_(ss, v) {
  if (typeof v === 'string' && /^\d{4}-\d{2}-\d{2}T/.test(v)) v = new Date(v);
  if (!(v instanceof Date) || isNaN(v.getTime())) return null;
  let tz = '';
  try { tz = String(ss.getSpreadsheetTimeZone() || ''); } catch (err) { tz = ''; }
  if (!tz || tz === 'null' || tz === 'undefined') tz = WEB_TZ; // zona waktu file tidak terbaca: anggap WIB
  try {
    return Utilities.parseDate(Utilities.formatDate(v, WEB_TZ, 'yyyy-MM-dd'), tz, 'yyyy-MM-dd');
  } catch (err) {
    return null; // jangan sampai cetak gagal karena ini: tanggal dari fungsi inventory dipakai apa adanya
  }
}

// file cetak (salinan pribadi) disamakan ke WIB supaya tanggal yang ditulis kode tidak bergeser 1 hari
function webSamakanZonaCetak_(ss) {
  try {
    if (String(ss.getSpreadsheetTimeZone() || '') !== WEB_TZ) ss.setSpreadsheetTimeZone(WEB_TZ);
  } catch (err) { Logger.log('Zona waktu file cetak tidak bisa diatur: ' + err); }
}

// baris terakhir nota: anchor + tambah, DIPERPANJANG sampai baris terakhir yang masih ada tulisannya
// (mis. "HARAP DITULIS", "Keterangan:", "PEMBAYARAN KAMI ANGGAP LUNAS ...", rekening) -- sama dengan
// yang tercetak dari sheet inventory. Hanya melihat sampai 12 baris di bawah anchor.
function webBarisAkhirNota_(kode, sh, teks, mulai, tambah, cadangan, kolom) {
  const r = kode.cariBarisAnchor_(sh, teks, mulai);
  if (r === -1) return cadangan;
  const sampai = Math.min(sh.getLastRow(), r + 12);
  let akhir = r + tambah;
  if (sampai >= r) {
    const v = sh.getRange(r, 1, sampai - r + 1, kolom).getDisplayValues();
    for (let i = v.length - 1; i >= 0; i--) {
      if (v[i].some(function (x) { return String(x).trim() !== ''; })) { akhir = Math.max(akhir, r + i); break; }
    }
  }
  return akhir;
}

function webBarisAkhirAnchor_(inv, sh, teks, mulai, tambah, cadangan) {
  const r = inv.cariBarisAnchor_(sh, teks, mulai);
  return r === -1 ? cadangan : r + tambah;
}

/**
 * a.jenis: SJ | INVOICE | SJ_SPR | INV_SPR | TT | PALET
 * a.nomor: No SJ / No Invoice / No SJ sparepart / No Invoice sparepart / No PB / No Invoice (referensi palet)
 * a.tandaiCetak (INVOICE / INV_SPR): catat SIAP CETAK di LOG_INVOICE_CETAK, sama dengan menu inventory
 */
function webCetak_(a, admin) {
  const jenis = String(a.jenis || '');
  const nomor = String(a.nomor || '').trim().toUpperCase();
  if (!nomor) throw new Error('Nomor dokumen kosong.');
  const ss = webFileCetak_();
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) throw new Error('Sedang ada nota lain yang dicetak, coba lagi sebentar.');
  const inv = webInv_();
  let hasil;
  WEB_CETAK_SS_ = ss;
  try {
    webSamakanZonaCetak_(ss);
    webSegarkanMasterCetak_(ss);
    if (jenis === 'SJ' || jenis === 'INVOICE') hasil = webCetakSJInvoice_(inv, ss, jenis, nomor, a);
    else if (jenis === 'SJ_SPR' || jenis === 'INV_SPR') hasil = webCetakSparepart_(inv, ss, jenis, nomor, a);
    else if (jenis === 'TT') hasil = webCetakTT_(inv, ss, nomor);
    else if (jenis === 'PALET') hasil = webCetakPalet_(inv, ss, nomor);
    else if (jenis === 'POK') hasil = webCetakPOKeluar_(inv, ss, nomor);
    else throw new Error('Jenis nota tidak dikenal.');
  } finally {
    WEB_CETAK_SS_ = null;
    lock.releaseLock();
  }
  webCatatLog_(admin, 'CETAK ' + jenis, nomor + ' | ' + hasil.nama + (hasil.tercetak ? ' | ' + hasil.tercetak : ''));
  return hasil;
}

function webCetakSJInvoice_(inv, ss, jenis, nomor, a) {
  const trx = webBaca_('TRANSAKSI', 17).filter(function (r) {
    return String(r[jenis === 'SJ' ? 1 : 2] || '').trim().toUpperCase() === nomor;
  }).pop();
  if (!trx) throw new Error(nomor + ' tidak ditemukan di TRANSAKSI.');
  const noSJ = String(trx[1]).trim();
  inv.cetakUlangSJDariNoSJ_(noSJ); // fungsi inventory: isi SURAT JALAN + INVOICE dari TRANSAKSI/DETAIL
  const sj = ss.getSheetByName('SURAT JALAN');
  const iv = ss.getSheetByName('INVOICE');
  sj.getRange('H6').setValue(noSJ);
  sj.getRange('H3').clearContent(); // exportSJPDF juga mengosongkan No PO sebelum export
  iv.getRange('H9').setValue(trx[16] || '');
  // tanggal nota = tanggal TRANSAKSI menurut jam WIB, apa pun zona waktu file cetak
  // (tanpa ini tanggal bisa mundur 1 hari kalau zona waktu file cetak bukan Asia/Jakarta)
  const tglTrx = webTglUntukSheet_(ss, trx[3]);
  if (tglTrx) {
    sj.getRange('H5').setValue(tglTrx).setNumberFormat('dd/mm/yyyy');
    iv.getRange('H7').setValue(tglTrx).setNumberFormat('dd/mm/yyyy');
  }
  const tglTempo = webTglUntukSheet_(ss, trx[6]);
  if (tglTempo) iv.getRange('H8').setValue(tglTempo).setNumberFormat('dd/mm/yyyy');
  SpreadsheetApp.flush();
  if (jenis === 'SJ') {
    const akhir = webBarisAkhirNota_(inv, sj, 'Yang Menyerahkan,', 9, 3, 20, 8);
    const r = webEksporPdf_(ss, sj, akhir, 8, noSJ + ' - ' + trx[4], 'SJ');
    return r;
  }
  let tercetak = '';
  if (a.tandaiCetak && WEB_LATIHAN_) tercetak = 'Mode latihan: invoice tidak ditandai SIAP CETAK.';
  if (a.tandaiCetak && !WEB_LATIHAN_) {
    tercetak = inv.konfirmasiSiapCetak({ sheetName: 'INVOICE', noInv: String(trx[2]).trim(), mode: 'cetak', alasan: String(a.alasan || '') }).pesan;
  }
  const akhir = webBarisAkhirNota_(inv, iv, 'Accounting,', 11, 4, 27, 9);
  const r = webEksporPdf_(ss, iv, akhir, 9, String(trx[2]).trim() + ' - ' + trx[4], 'INVOICE');
  r.tercetak = tercetak;
  return r;
}

function webCetakSparepart_(inv, ss, jenis, nomor, a) {
  const spr = webBaca_('SPAREPART', 2).filter(function (r) {
    return String(r[jenis === 'SJ_SPR' ? 0 : 1] || '').trim().toUpperCase() === nomor;
  })[0];
  if (!spr) throw new Error(nomor + ' tidak ditemukan di SPAREPART.');
  const props = PropertiesService.getDocumentProperties();
  try {
    inv.terapkanPilihanSJSparepartDraft(String(spr[0]).trim()); // fungsi Buka Draft sparepart inventory
  } finally {
    props.deleteProperty('SPAREPART_EDIT_NOSJ');   // penanda mode edit milik form inventory, tidak dipakai di sini
    props.deleteProperty('SPAREPART_EDIT_STATUS');
  }
  SpreadsheetApp.flush();
  const sh = ss.getSheetByName(jenis);
  let tercetak = '';
  if (jenis === 'INV_SPR' && a.tandaiCetak && WEB_LATIHAN_) tercetak = 'Mode latihan: invoice tidak ditandai SIAP CETAK.';
  if (jenis === 'INV_SPR' && a.tandaiCetak && !WEB_LATIHAN_) {
    tercetak = inv.konfirmasiSiapCetak({ sheetName: 'INV_SPR', noInv: String(spr[1]).trim(), mode: 'cetak', alasan: String(a.alasan || '') }).pesan;
  }
  const akhir = jenis === 'SJ_SPR' ? webBarisAkhirNota_(inv, sh, 'Yang Menyerahkan,', 9, 3, 20, 8) : webBarisAkhirNota_(inv, sh, 'Accounting,', 11, 8, Math.max(webBarisIsiTerakhir_(sh, 9, 1), 27), 9);
  const r = webEksporPdf_(ss, sh, akhir, jenis === 'SJ_SPR' ? 8 : 9, String(jenis === 'SJ_SPR' ? spr[0] : spr[1]).trim(), 'LAIN');
  r.tercetak = tercetak;
  return r;
}

function webCetakTT_(inv, ss, noPB) {
  const rows = webBaca_('PEMBELIAN', 15).filter(function (r) { return String(r[0] || '').trim().toUpperCase() === noPB; });
  if (!rows.length) throw new Error(noPB + ' tidak ditemukan di PEMBELIAN.');
  const existing = inv.cariBarisPembelianByNoSJ_(String(rows[0][2]).trim(), rows[0][3]).filter(function (e) {
    return String(e.noPB || '').trim().toUpperCase() === noPB;
  });
  if (!existing.length) throw new Error('Data ' + noPB + ' tidak terbaca.');
  const sh = ss.getSheetByName('TANDA TERIMA BRG');
  inv.isiFormTandaTerimaDariExisting_(sh, existing[0], existing); // fungsi inventory (edit SJ Beli)
  const jumlah = inv.hitungUlangHargaTandaTerima_(sh);
  // harga yang dicetak = harga yang TERSIMPAN di PEMBELIAN
  const harga = existing.map(function (e) { return [Number(e.harga) || 0]; });
  if (jumlah > 0) {
    sh.getRange(7, 5, harga.length, 1).setValues(harga);
    sh.getRange(7 + jumlah, 5).setValue(harga.reduce(function (s, h) { return s + h[0]; }, 0));
  }
  SpreadsheetApp.flush();
  return webEksporPdf_(ss, sh, 7 + jumlah + 2, 8, noPB + ' - ' + existing[0].supplier, 'LAIN'); // sampai baris Total
}

function webCetakPalet_(inv, ss, ref) {
  const rows = inv.bacaEkspedisiDb_().filter(function (r) {
    return String(r[1] || '').trim().toUpperCase() === ref && inv.jenisBarisEkspedisi_(r);
  });
  if (!rows.length) throw new Error('Tidak ada biaya EKSPEDISI/PALET untuk ' + ref + '.');
  try {
    inv.isiSheetInvoicePalet_(ref, rows); // fungsi inventory: susun sheet INVOICE PALET
  } catch (err) {
    if (!ss.getSheetByName('INVOICE PALET')) throw err; // error setelah sheet terisi (pindah tab) diabaikan
  }
  SpreadsheetApp.flush();
  const sh = ss.getSheetByName('INVOICE PALET');
  return webEksporPdf_(ss, sh, webBarisIsiTerakhir_(sh, 7, 1), 7, 'INVOICE PALET ' + ref + ' - ' + rows[0][5], 'LAIN');
}

// PO Keluar: isi form PO_KELUAR dengan isiPOKeluarDariHistory_ lalu export seperti exportPOKeluarPDF
function webCetakPOKeluar_(inv, ss, poNo) {
  const items = inv.ambilItemHistoryPOKeluar_(null, poNo);
  if (!items.length) throw new Error('PO keluar ' + poNo + ' tidak ditemukan di HISTORY_PO_KELUAR.');
  const tglRow = webBaca_('HISTORY_PO_KELUAR', 2).filter(function (r) { return String(r[0] || '').trim() === poNo; })[0];
  const sh = ss.getSheetByName('PO_KELUAR');
  sh.getRange('B4').setValue(poNo);
  if (tglRow && tglRow[1]) sh.getRange('E4').setValue(tglRow[1]);
  inv.isiPOKeluarDariHistory_(sh, items);
  webRapikanKepalaPOK_(sh, poNo, tglRow && tglRow[1]); // cetakan untuk pabrik: T-xxx, kepala rapi
  SpreadsheetApp.flush();
  // sampai baris TOTAL QTY (dulu sampai baris terakhir sheet -> banyak halaman kosong)
  let akhir = 7 + items.length + 1;
  try { akhir = webBarisAkhirPOK_(sh, items.length); } catch (err) { Logger.log('Batas PO keluar: ' + err); }
  return webEksporPdf_(ss, sh, akhir, 6, 'PO KELUAR - ' + webNomorPOCetak_(poNo), 'POK', 0.3);
}


const WEB_AWALAN_PO_CETAK_ = 'T'; // SANFELO: #2026 -> T di cetakan PO keluar

/* ---------------- batas cetak & PO keluar untuk pabrik ---------------- */

// Baris terakhir yang benar-benar berisi teks (dibaca dari atas). Berhenti setelah 15 baris kosong
// berturut-turut, supaya sisa isian jauh di bawah template tidak ikut tercetak (halaman kosong).
function webBarisIsiTerakhir_(sh, kolom, mulai) {
  const m = mulai || 1;
  const n = Math.max(1, sh.getLastRow() - m + 1);
  const v = sh.getRange(m, 1, n, kolom).getDisplayValues();
  let akhir = m, kosong = 0;
  for (let i = 0; i < v.length; i++) {
    if (v[i].some(function (x) { return String(x).trim() !== ''; })) { akhir = m + i; kosong = 0; }
    else if (++kosong >= 15) break;
  }
  return akhir;
}

// PO_KELUAR: nota selesai di baris "TOTAL QTY" (kolom C)
function webBarisAkhirPOK_(sh, jumlahItem) {
  const batas = Math.min(sh.getLastRow(), 7 + (jumlahItem || 0) + 60);
  const v = sh.getRange(7, 3, Math.max(1, batas - 6), 1).getDisplayValues();
  for (let i = 0; i < v.length; i++) if (String(v[i][0]).trim().toUpperCase() === 'TOTAL QTY') return 7 + i;
  return webBarisIsiTerakhir_(sh, 6, 1);
}

// nomor PO di cetakan untuk pabrik: #2026-096 -> T-096 (SANFELO) / N-096 (SL). Data tidak berubah.
function webNomorPOCetak_(poNo) {
  const m = String(poNo || '').trim().match(/^#?\d{4}-(\d+)$/);
  return m ? WEB_AWALAN_PO_CETAK_ + '-' + m[1] : String(poNo || '').trim();
}

// kepala PO keluar (hanya di file cetak web, form inventory tidak disentuh):
// tata letak sama dengan form PO_KELUAR (logo A1, judul & alamat C1:E1 / C2:E2), alamat rata kanan,
// NO. PO tetap di B4 (T-/N-), TGL rata kanan. TANPA merge sel baru (merge bisa gagal "Kesalahan layanan");
// setiap langkah dijaga supaya PDF tetap jadi walaupun perapian gagal.
function webRapikanKepalaPOK_(sh, poNo, tgl) {
  const coba = function (fn) { try { fn(); } catch (err) { Logger.log('Rapikan PO keluar: ' + err); } };
  // pulihkan bentuk template (percobaan versi sebelumnya bisa membuat C1:F1 / E4:F4 tergabung)
  [1, 2].forEach(function (r) {
    coba(function () { sh.getRange(r, 3, 1, 4).breakApart(); });
    coba(function () { sh.getRange(r, 3, 1, 3).merge(); });
  });
  coba(function () { sh.getRange('E4:F4').breakApart(); });
  coba(function () { sh.getRange(1, 3, 1, 3).setHorizontalAlignment('center'); });
  coba(function () { sh.getRange(2, 3, 1, 3).setHorizontalAlignment('right'); });
  coba(function () { sh.getRange('B4').setValue(webNomorPOCetak_(poNo)).setHorizontalAlignment('left'); });
  coba(function () {
    const teksTgl = tgl instanceof Date ? Utilities.formatDate(tgl, WEB_TZ, 'dd/MM/yyyy') : String(sh.getRange('E4').getDisplayValue() || '').trim();
    sh.getRange('D4').clearContent();
    sh.getRange('E4').setNumberFormat('@').setValue('TGL:  ' + teksTgl).setHorizontalAlignment('right');
  });
}
