/* ======================================================
   WEB_INV_SALINAN.js   (PROJECT: database_SANFELO) -- DIBUAT OTOMATIS
   ------------------------------------------------------
   Salinan PERSIS beberapa file inventory_SANFELO (Utils.js, Database.js, TRANSAKSI_DB.js, Pembayaran.js, StatusDone.js, PENJUALAN.js, TempoToko.js, Transaction.js, PO_HISTORY.js, SJ_DRAFT_FETCH.js, Kode.js, InvoiceGroupingPopup.js, PembelianHutangSync.js, QTY_TERKIRIM_ULANG.js, TANDA_TERIMA_BRG.js, BiayaTambahan.js, INVOICE_PALET.js, SPAREPART_SAVE.js, STOK_BARANG.js, PettyCash.js, ReprintSJ.js, INVOICE_CETAK.js, SPLIT_SURAT.js, SyncDone.js, PO_KELUAR.js),
   dibungkus di dalam satu fungsi supaya tidak bentrok dengan kode
   database_SANFELO. Dipakai web supaya hitungan Surat Jalan / Invoice
   (tempo, potongan 3SET/5SET, status PO, PENJUALAN, TEMPO TOKO) SAMA
   PERSIS dengan tombol Simpan Draft di spreadsheet inventory.
   Jalur baca/tulis database tetap lewat gateway yang sama
   (kirimAksiGateway_ di sini memanggil prosesRequestGateway_ langsung,
   dengan isi dikirim & diterima dalam bentuk JSON seperti lewat internet).
   JANGAN diedit manual -- kalau file inventory berubah, buat ulang file ini.
   ====================================================== */

var WEB_INV_CACHE_ = null;

function webInv_() {
  if (WEB_INV_CACHE_) return WEB_INV_CACHE_;

  const kirimAksiGateway_ = function (action, payload) {
    return JSON.parse(JSON.stringify(webGateway_(action, payload)));
  };
  const bacaSheetGateway_ = function (sheetName, startRow, startCol, numCols) {
    const payload = { sheet: sheetName, startRow: startRow, startCol: startCol };
    if (numCols) payload.numCols = numCols;
    return kirimAksiGateway_('bacaSheet', payload).rows || [];
  };
  const pesanErrorGatewayBukanJson_ = function (kode) { return 'Gateway gagal (kode ' + kode + ').'; };
  const SpreadsheetApp = webInvSpreadsheetApp_();
  const LockService = webInvLockService_();

  // ======================= Utils.js (salinan persis) =======================
/* ======================================================
   ⬅️ BARU: konstanta & helper TEMPO DISPLAY, dipakai bareng oleh
   Transaction.js (simpanDraft), TempoToko.js (syncTempoToko_), dan
   PUSH_LAPORAN.js (buildRawLaporanRows_).

   MASALAH YANG DIPERBAIKI:
   Tempo 90 hari untuk toko yang dapat diskon DISPLAY cuma kepasang
   otomatis KALAU status DISPLAY-nya sempat ketulis rapi di
   HISTORY_PO_MASUK (kolom DISKON atau kolom METHOD). Di lapangan,
   admin kadang cuma nulis "DISPLAY" di kolom Keterangan invoice
   (sheet INVOICE, baris anchor "Keterangan:") tanpa nandain status
   itu di HISTORY_PO_MASUK. Akibatnya tempo jatuh ke default toko
   (mis. SIRO = 45 HARI) padahal harusnya 90 HARI.

   FIX: kalau teks Keterangan mengandung kata "DISPLAY", itu dipakai
   sebagai SINYAL TAMBAHAN (bukan pengganti) buat menentukan tempo —
   digabung pakai logika "ambil yang PALING PANJANG" yang sudah ada,
   supaya tetap konsisten sama urutan hitung tempo yang lama.
   ====================================================== */
const TEMPO_HARI_DISPLAY = 90;

function cekKeteranganDisplay_(teks) {
  return String(teks || "").trim().toUpperCase().indexOf("DISPLAY") !== -1;
}

/**
 * Gabungkan kandidat jumlah hari tempo dari PO (bisa null) dengan
 * override dari teks Keterangan invoice (kalau mengandung "DISPLAY").
 * Selalu ambil yang PALING PANJANG (paling aman buat toko), sama
 * seperti logika lama saat ada beberapa PO dengan status berbeda.
 */
function gabungTempoDenganKeterangan_(bestDaysDariPO, keteranganText) {
  let bestDays = bestDaysDariPO;
  if (cekKeteranganDisplay_(keteranganText)) {
    if (bestDays === null || bestDays === undefined || TEMPO_HARI_DISPLAY > bestDays) {
      bestDays = TEMPO_HARI_DISPLAY;
    }
  }
  return bestDays;
}

/**
 * Buang jam dari sebuah tanggal -> cuma tanggalnya (jam 00:00).
 * Menerima Date, atau teks "dd/MM/yyyy" / "dd/MM/yyyy HH:mm:ss". Selain itu dibalikin apa adanya.
 */
function hanyaTanggal_(v) {
  if (v instanceof Date && !isNaN(v.getTime())) {
    return new Date(v.getFullYear(), v.getMonth(), v.getDate());
  }
  if (typeof v === "string") {
    const m = v.trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    if (m) return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
  }
  return v;
}

/** Tulis tanggal ke sel TANPA jam + paksa tampilan dd/MM/yyyy. */
function setTanggalSaja_(range, v) {
  range.setValue(hanyaTanggal_(v)).setNumberFormat("dd/MM/yyyy");
}

function getTimestamp() {

  return Utilities.formatDate(
    new Date(),
    "Asia/Jakarta",
    "dd/MM/yyyy HH:mm:ss"
  );

}

function generateID() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("TRANSAKSI");

  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    return 1;
  }

  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();

  let maxID = 0;

  ids.forEach(r => {
    const id = Number(r[0]);
    if (!isNaN(id) && id > maxID) {
      maxID = id;
    }
  });

  return maxID + 1;

}

function bersihkanBarisKosongTransaksi() {
  // ⬅️ MIGRASI (Oktober 2026): TRANSAKSI & DETAIL_TRANSAKSI lokal sekarang
  // salinan IMPORTRANGE -- menghapus baris di sini bisa merusak salinan.
  SpreadsheetApp.getUi().alert("Tidak dipakai lagi: TRANSAKSI & DETAIL_TRANSAKSI sekarang salinan dari database_SANFELO. Rapikan baris kosong langsung di database.");
  return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const namaSheetList = ["TRANSAKSI", "DETAIL_TRANSAKSI"];
  let laporan = [];

  namaSheetList.forEach(function (namaSheet) {
    const sheet = ss.getSheetByName(namaSheet);
    if (!sheet) return;

    const maxRows = sheet.getMaxRows();
    const dataRange = sheet.getDataRange();
    const data = dataRange.getValues();

    let lastRealRow = 0;
    for (let i = 0; i < data.length; i++) {
      const adaIsi = data[i].some(function (cell) { return String(cell).trim() !== ""; });
      if (adaIsi) lastRealRow = i + 1;
    }

    if (maxRows > lastRealRow) {
      const jumlahDihapus = maxRows - lastRealRow;
      sheet.deleteRows(lastRealRow + 1, jumlahDihapus);
      laporan.push(namaSheet + ": " + jumlahDihapus + " baris kosong dihapus (sisa " + lastRealRow + " baris).");
    } else {
      laporan.push(namaSheet + ": tidak ada baris kosong berlebih.");
    }
  });

  ui.alert("Pembersihan selesai:\n\n" + laporan.join("\n"));
}

  // ======================= Database.js (salinan persis) =======================
function getTimestamp() {
  return new Date();
}

function cariTransaksi(noSJ) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("TRANSAKSI");

  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][1] == noSJ) {
      return i + 1; // nomor baris spreadsheet
    }
  }

  return -1;

}

  // ======================= TRANSAKSI_DB.js (salinan persis) =======================
/* ======================================================
   TRANSAKSI_DB.js   (PROJECT: inventory_SANFELO)
   ------------------------------------------------------
   FILE BARU (migrasi TRANSAKSI + DETAIL_TRANSAKSI + PENJUALAN,
   Oktober 2026). Sejak migrasi ini, ketiga sheet itu di inventory cuma
   SALINAN (IMPORTRANGE) dari database_SANFELO -- tidak bisa ditulis
   langsung. Semua penulisan lewat gateway (lihat GatewayTransaksi.js &
   aksi PENJUALAN di DatabaseGateway.js sisi database).

   Isi file ini:
   - bacaTransaksiDB_ / bacaDetailDB_ / bacaPenjualanDB_ : baca FRESH
     langsung dari database (bukan dari salinan lokal yang bisa telat
     beberapa detik setelah ada simpan). Tanggal dari gateway datang
     sebagai teks ISO -> dikembalikan jadi Date di sini, supaya kode lama
     yang dulu baca sheet lokal tetap menerima tipe yang sama.
   - refreshSalinan_ : paksa salinan IMPORTRANGE lokal ikut ter-update.
   - ubahStatusTransaksi : menu "✏️ Ubah Status Transaksi" (pengganti
     dropdown STATUS ORDER yang dulu diketik langsung di sheet).
   ====================================================== */

const DB_SPREADSHEET_ID_ = '1UPHKIUAornJxCZ_9CbEJAUTeqfhcKJPJ2la_6GMSqfo'; // database_SANFELO

const KOLOM_TANGGAL_TRANSAKSI_ = [3, 6, 8];       // D TANGGAL, G TEMPO, I LAST UPDATE (index 0-based)
const KOLOM_TANGGAL_DETAIL_ = [1, 4, 10, 11];     // B TANGGAL, E TEMPO, K CREATED_AT, L UPDATED_AT
const KOLOM_TANGGAL_PENJUALAN_ = [0];             // A TGL KIRIM

/** Teks ISO dari gateway ("2026-10-01T03:10:58.553Z") -> Date. Nilai lain dibiarkan. */
function keTanggalDB_(v) {
  if (typeof v === 'string' && /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}/.test(v)) {
    const d = new Date(v);
    return isNaN(d.getTime()) ? v : d;
  }
  return v;
}

function konversiTanggalBaris_(rows, idxList) {
  return rows.map(function (r) {
    const out = r.slice();
    idxList.forEach(function (i) { if (i < out.length) out[i] = keTanggalDB_(out[i]); });
    return out;
  });
}

/** TRANSAKSI A..Q (17 kolom), mulai baris 2, fresh dari database. */
function bacaTransaksiDB_() {
  return konversiTanggalBaris_(bacaSheetGateway_('TRANSAKSI', 2, 1, 17), KOLOM_TANGGAL_TRANSAKSI_);
}

/** DETAIL_TRANSAKSI A..U (21 kolom), mulai baris 2, fresh dari database. */
function bacaDetailDB_() {
  return konversiTanggalBaris_(bacaSheetGateway_('DETAIL_TRANSAKSI', 2, 1, 21), KOLOM_TANGGAL_DETAIL_);
}

/** PENJUALAN selebar sheet (A..riwayat cicilan), mulai baris 2, fresh dari database. */
function bacaPenjualanDB_() {
  return konversiTanggalBaris_(bacaSheetGateway_('PENJUALAN', 2, 1), KOLOM_TANGGAL_PENJUALAN_);
}

/** Paksa salinan IMPORTRANGE di sheet-sheet ini ikut ter-update sekarang. */
function refreshSalinan_(namaSheets) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  (namaSheets || []).forEach(function (nama) {
    try {
      const sh = ss.getSheetByName(nama);
      if (!sh) return;
      const cell = sh.getRange('A2');
      const f = cell.getFormula();
      if (f && f.toUpperCase().indexOf('IMPORTRANGE') !== -1) cell.setFormula(f);
    } catch (err) {
      // tidak fatal -- salinan tetap ter-update sendiri beberapa detik kemudian
    }
  });
}

/* ------------------------------------------------------
   MENU: ✏️ Ubah Status Transaksi
   Cara pakai: buka sheet TRANSAKSI atau DETAIL_TRANSAKSI, klik SATU sel
   di baris yang mau diubah, lalu pilih menu ini. Status dikirim ke
   database; salinan di sheet ikut berubah beberapa detik kemudian.
   Aturan lama tetap berlaku: DONE hanya bisa kalau invoice sudah LUNAS.
   ------------------------------------------------------ */
function ubahStatusTransaksi() {
  const ui = SpreadsheetApp.getUi();
  const sh = SpreadsheetApp.getActiveSheet();
  const nama = sh.getName();

  if (nama !== 'TRANSAKSI' && nama !== 'DETAIL_TRANSAKSI') {
    ui.alert('Buka sheet TRANSAKSI atau DETAIL_TRANSAKSI, klik baris yang mau diubah statusnya, lalu pilih menu ini lagi.');
    return;
  }
  const row = sh.getActiveRange().getRow();
  if (row < 2) {
    ui.alert('Klik dulu baris data (bukan baris judul).');
    return;
  }

  let payload, info, pilihan, statusLama;
  if (nama === 'TRANSAKSI') {
    const v = sh.getRange(row, 1, 1, 17).getValues()[0];
    if (!String(v[1] || '').trim()) { ui.alert('Baris ' + row + ' kosong.'); return; }
    payload = { sheet: nama, noSJ: String(v[1]).trim() };
    info = v[1] + ' / ' + v[2] + ' / ' + v[4];
    statusLama = v[13];
    pilihan = 'PROGRESS / REVISI / DONE';
  } else {
    const v = sh.getRange(row, 1, 1, 21).getValues()[0];
    if (!String(v[0] || '').trim()) { ui.alert('Baris ' + row + ' kosong.'); return; }
    payload = { sheet: nama, noSJ: String(v[0]).trim(), baris: v[15], kode: String(v[5] || '').trim() };
    info = v[0] + ' / ' + v[2] + ' / ' + (v[5] || '') + ' ' + (v[6] || '');
    statusLama = v[19];
    pilihan = 'PO / DONE';
  }

  const jawab = ui.prompt('Ubah Status - ' + nama,
    'Baris ' + row + ': ' + info + '\nStatus sekarang: ' + (statusLama || '(kosong)') +
    '\n\nKetik status baru (' + pilihan + '):', ui.ButtonSet.OK_CANCEL);
  if (jawab.getSelectedButton() !== ui.Button.OK) return;
  payload.status = String(jawab.getResponseText() || '').trim().toUpperCase();
  if (!payload.status) return;

  try {
    const hasil = kirimAksiGateway_('ubahStatusTransaksi', payload);
    refreshSalinan_([nama]);
    ui.alert('✅ Status ' + info + ' diubah: ' + (hasil.statusLama || '(kosong)') + ' → ' + hasil.statusBaru +
             '.\nSheet akan menampilkan perubahan dalam beberapa detik.');
  } catch (err) {
    ui.alert('Status tidak diubah:\n' + String(err.message || err).replace(/^Gateway menolak: /, ''));
  }
}

/* ------------------------------------------------------
   Rumus salinan tiap sheet (dipakai MIGRASI_TRANSAKSI_PENJUALAN.js saat
   pasang, dan onEdit di bawah untuk memulihkan kalau rumusnya tertimpa).
   PENJUALAN dibuat lebar (sampai Z) karena riwayat cicilan nambah ke kanan.
   ------------------------------------------------------ */
const SALINAN_DB_ = {
  'TRANSAKSI':        { range: 'TRANSAKSI!A2:Q',        lebar: 17 },
  'DETAIL_TRANSAKSI': { range: 'DETAIL_TRANSAKSI!A2:U', lebar: 21 },
  'PENJUALAN':        { range: 'PENJUALAN!A2:Z',        lebar: 26 }
};

function rumusSalinanDB_(namaSheet) {
  return '=IMPORTRANGE("' + DB_SPREADSHEET_ID_ + '";"' + SALINAN_DB_[namaSheet].range + '")';
}

/**
 * Dipanggil onEdit (PO_HISTORY.js) kalau ada yang mengetik langsung di
 * sheet salinan. Ketikan di area data membuat IMPORTRANGE error (#REF!),
 * jadi langsung dipulihkan: sel yang diketik dikosongkan lagi, atau
 * rumus di A2 dipasang ulang kalau A2 yang tertimpa. Baris judul (1) dan
 * kolom di luar lebar salinan dibiarkan.
 */
function pulihkanSheetSalinan_(e, namaSheet) {
  try {
    const cfg = SALINAN_DB_[namaSheet];
    if (!cfg || !e || !e.range) return;
    const r = e.range;
    if (r.getLastRow() < 2 || r.getColumn() > cfg.lebar) return;

    const sh = r.getSheet();
    const a2 = sh.getRange('A2');
    const kenaA2 = r.getRow() <= 2 && r.getColumn() === 1;
    if (kenaA2) {
      a2.setFormula(rumusSalinanDB_(namaSheet));
    } else {
      const baris = Math.max(r.getRow(), 2);
      const kolomAkhir = Math.min(r.getLastColumn(), cfg.lebar);
      sh.getRange(baris, r.getColumn(), r.getLastRow() - baris + 1, kolomAkhir - r.getColumn() + 1).clearContent();
    }
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Sheet ' + namaSheet + ' adalah salinan dari database, jadi tidak bisa diketik langsung — ketikan tadi dibatalkan otomatis. ' +
      'Untuk ubah status pakai menu 🏢 SANFELO > ✏️ Ubah Status Transaksi; koreksi data lain lewat IT/owner di database.',
      '⚠️ Sheet salinan', 15);
  } catch (err) {
    Logger.log('pulihkanSheetSalinan_ error: ' + err);
  }
}


  // ======================= Pembayaran.js (salinan persis) =======================
/* ======================================================
   Pembayaran.gs
   - showPembayaranDialog()      : buka popup Pembayaran
   - cariDataInvoice()           : cari data invoice dari sheet PENJUALAN
   - prosesPembayaran()          : simpan pembayaran/cicilan baru ke PENJUALAN

   ⬅️ Fungsi Biaya Tambahan (EKSPEDISI/PALET) sudah dipindah ke
   file terpisah: BiayaTambahan.gs (showBiayaTambahDialog,
   simpanBiayaTambahan, hitungUlangSisaBayar_). Konstanta &
   helper umum (PENJUALAN_COL_*, formatRupiahText_, dst) tetap
   di sini karena dipakai bareng sama file Biaya Tambahan.

   ⬅️ BARU: pas status invoice jadi LUNAS, selain nandain DONE di
   TRANSAKSI/DETAIL_TRANSAKSI (tandaiDoneByNoInv_, dari StatusDone.gs),
   sekarang juga nandain DONE + hijau di SPAREPART kalau invoice-nya
   asalnya dari situ (tandaiDoneSparepartByNoInv_, fungsi baru di
   bawah). Realtime, langsung jalan begitu prosesPembayaran() bikin
   status LUNAS.
   ====================================================== */

/**
 * ====== STRUKTUR SHEET PENJUALAN (mulai baris 2) ======
 *   A: TGL KIRIM   B: CUSTOMER   C: NO INV   D: METHOD   E: TOTAL
 *   F: STATUS ("CICIL" / "LUNAS")
 *   G: SISA BAYAR (angka, = E + H + I - total sudah dibayar)
 *   H: EKSPEDISI (biaya tambahan, angka)
 *   I: PALET (biaya tambahan, angka)
 *   J, K, L, ...: satu kolom per cicilan, isi teks gabungan
 *                 "dd/MM/yyyy - RpX.XXX.XXX - METODE", nambah ke
 *                 kanan tiap ada pembayaran baru.
 *   Kalau STATUS = LUNAS -> background 1 baris penuh jadi hijau.
 */

const PENJUALAN_COL_TGLKIRIM = 1;
const PENJUALAN_COL_CUSTOMER = 2;
const PENJUALAN_COL_NOINV = 3;
const PENJUALAN_COL_METHOD = 4;
const PENJUALAN_COL_TOTAL = 5;
const PENJUALAN_COL_STATUS = 6;
const PENJUALAN_COL_SISA = 7;
const PENJUALAN_COL_EKSPEDISI = 8;  // H
const PENJUALAN_COL_PALET = 9;      // I
const PENJUALAN_COL_CICILAN_START = 10; // J
const PENJUALAN_WARNA_LUNAS = "#d9ead3";
const PENJUALAN_WARNA_LEBIH = "#cfe2f3";

/* ------------------------------------------------------
   1) BUKA POPUP PEMBAYARAN
   ------------------------------------------------------ */
function showPembayaranDialog() {
  const html = HtmlService.createHtmlOutputFromFile("PembayaranDialog")
    .setWidth(480)
    .setHeight(620);
  SpreadsheetApp.getUi().showModalDialog(html, "💰 Pembayaran Invoice");
}

/* ------------------------------------------------------
   2) HELPER FORMAT & PARSE RUPIAH
   ------------------------------------------------------ */
function formatRupiahText_(number) {
  return "Rp" + Number(number || 0).toLocaleString("id-ID");
}

function parseRupiahText_(text) {
  const digits = String(text).replace(/[^0-9]/g, "");
  return Number(digits) || 0;
}

// Ubah "yyyy-MM-dd" (hasil <input type="date">) -> "dd/MM/yyyy"
function formatTglBayarInput_(tglBayarStr) {
  const parts = String(tglBayarStr || "").split("-"); // [yyyy, MM, dd]
  if (parts.length !== 3) return null;
  const yyyy = parts[0];
  const mm = parts[1];
  const dd = parts[2];
  if (!yyyy || !mm || !dd) return null;
  return dd + "/" + mm + "/" + yyyy;
}

// ⬅️ BARU — "No Post": nomor unik yang ditempel di depan tiap baris riwayat
// cicilan PENJUALAN, format simpel "POST-NNN" (POST-001, POST-002, dst),
// TERUS NAIK (gak reset tiap bulan). Counter-nya di-cache di sheet SYSTEM
// sel B4, TAPI tiap generate nomor baru selalu dibandingkan dulu ke hasil
// scan ULANG semua riwayat cicilan yang beneran ada di PENJUALAN (sumber
// kebenaran), diambil yang PALING BESAR -- jadi gak akan pernah nabrak/
// dobel walau cache SYSTEM-nya ketinggalan.
const NOPOST_PREFIX_ = "POST-";
const SYSTEM_NOPOST_CELL_COUNTER_ = "B4";

function getNextNoPost_(ss) {
  const system = ss.getSheetByName("SYSTEM");
  const shPenjualan = ss.getSheetByName("PENJUALAN");

  // 1) scan ULANG semua riwayat cicilan yang SUDAH ADA di PENJUALAN,
  // cari No Post TERTINGGI yang pernah dipakai (sumber kebenaran utama).
  let maxDariData = 0;
  if (shPenjualan) {
    const lastRow = shPenjualan.getLastRow();
    const lastCol = shPenjualan.getLastColumn();
    if (lastRow >= 2 && lastCol >= PENJUALAN_COL_CICILAN_START) {
      const data = shPenjualan
        .getRange(2, PENJUALAN_COL_CICILAN_START, lastRow - 1, lastCol - PENJUALAN_COL_CICILAN_START + 1)
        .getValues();

      data.forEach(function (row) {
        row.forEach(function (cell) {
          const str = String(cell || "").trim();
          // ⬅️ PENTING: cuma dianggap "No Post" valid kalau formatnya PERSIS
          // "POST-<angka> - " (gak ada tanda "-" lain nyempil di dalam
          // angkanya). Ini supaya cell format LAMA-nya fitur ini (sempat
          // dites pakai "POST-2026-09-001 - ...") TIDAK ikut kehitung /
          // gak bikin nomor lompat jauh -- cell semacam itu justru
          // ditangani sebagai "belum bernomor" oleh backfillNoPostDataLama().
          const m = str.match(/^POST-(\d+) - /);
          if (!m) return;
          const num = parseInt(m[1], 10);
          if (!isNaN(num) && num > maxDariData) maxDariData = num;
        });
      });
    }
  }

  // 2) cache counter di SYSTEM
  let counterCache = 0;
  if (system) {
    counterCache = Number(system.getRange(SYSTEM_NOPOST_CELL_COUNTER_).getValue()) || 0;
  }

  // 3) ambil yang PALING BESAR antara scan data vs cache, +1
  const nextNum = Math.max(maxDariData, counterCache) + 1;

  // 4) update cache SYSTEM
  if (system) {
    system.getRange(SYSTEM_NOPOST_CELL_COUNTER_).setValue(nextNum);
  }

  return NOPOST_PREFIX_ + Utilities.formatString("%03d", nextNum);
}

// Helper: parse teks tanggal "dd/MM/yyyy" -> Date object (null kalau formatnya gak sesuai)
function parseTanggalDDMMYYYY_(str) {
  const parts = String(str || "").split("/");
  if (parts.length !== 3) return null;
  const dd = Number(parts[0]);
  const mm = Number(parts[1]);
  const yyyy = Number(parts[2]);
  if (!dd || !mm || !yyyy) return null;
  return new Date(yyyy, mm - 1, dd);
}

/* ------------------------------------------------------
   ⬅️ BARU — KETERANGAN ADJUSTMENT
   Kalau metode bayar = ADJUSTMENT, admin WAJIB mengetik keterangannya
   (mis. pembulatan, biaya admin, sparepart) supaya nanti tidak salah
   paham ini adjustment untuk apa. Keterangan disimpan di dalam teks
   riwayat pembayaran (PENJUALAN dan HUTANG) pada bagian METODE:
       "POST-NNN - dd/MM/yyyy - RpX - ADJUSTMENT (pembulatan)"
   Karena tetap berada di bagian METODE, semua parser lama yang memecah
   teks dengan " - " tidak terganggu. Metode TUNAI / TRANSFER tidak
   berubah sama sekali. Data adjustment LAMA (tanpa keterangan) tetap
   terbaca, keterangannya kosong.
   ------------------------------------------------------ */
const ADJUSTMENT_KET_MIN_ = 3;
const ADJUSTMENT_KET_MAX_ = 60;

// Bersihkan teks keterangan: tanpa enter/tab, tanpa " - " (pemisah riwayat)
// dan tanpa tanda kurung (dipakai sebagai pembungkus keterangan).
function bersihkanKeteranganAdjustment_(ket) {
  return String(ket || "")
    .replace(/[\r\n\t]+/g, " ")
    .replace(/\s+-\s+/g, " ")
    .replace(/[()]/g, " ")
    .replace(/\s{2,}/g, " ")
    .trim()
    .substring(0, ADJUSTMENT_KET_MAX_)
    .trim();
}

// metode harus sudah huruf besar (TUNAI / TRANSFER / ADJUSTMENT).
function validasiKeteranganAdjustment_(metode, ketMentah) {
  if (metode !== "ADJUSTMENT") return { ok: true, keterangan: "" };
  const ket = bersihkanKeteranganAdjustment_(ketMentah);
  if (ket.length < ADJUSTMENT_KET_MIN_) {
    return {
      ok: false,
      keterangan: "",
      message: "Metode Adjustment wajib diisi keterangannya (minimal " + ADJUSTMENT_KET_MIN_ +
        " huruf), contoh: pembulatan, biaya admin, sparepart."
    };
  }
  return { ok: true, keterangan: ket };
}

// Teks METODE yang ditulis ke sel riwayat.
function susunTeksMetodeBayar_(metode, keterangan) {
  if (metode === "ADJUSTMENT" && keterangan) return "ADJUSTMENT (" + keterangan + ")";
  return metode;
}

// Kebalikan susunTeksMetodeBayar_: "ADJUSTMENT (pembulatan)" -> {metode, keterangan}
function pecahMetodeBayar_(teks) {
  const str = String(teks || "").trim();
  const m = str.match(/^ADJUSTMENT\s*\(([\s\S]*)\)\s*$/i);
  if (m) return { metode: "ADJUSTMENT", keterangan: m[1].trim() };
  return { metode: str, keterangan: "" };
}

// Parse 1 sel cicilan -> {tanggal, jumlah, metode, keterangan, metodeTeks, noPost}
// Format BARU (sejak ada No Post): "POST-NNN - dd/MM/yyyy - RpX - METODE"
// Format LAMA (data historis sebelum ada No Post): "dd/MM/yyyy - RpX - METODE"
// ⬅️ Dua-duanya tetap kebaca benar -- dikenali dari bagian pertama diawali
// "POST-" atau bukan (termasuk format lama "POST-yyyy-MM-NNN" yang sempat
// dites, tetap kebaca sebagai punya noPost meski nanti direnumber ulang
// oleh backfillNoPostDataLama()).
function parseRiwayatCell_(text) {
  const str = String(text).trim();
  if (!str) return null;
  const parts = str.split(" - ");
  if (parts.length < 2) return null;

  let idx = 0;
  let noPost = "";
  if (parts[0].trim().indexOf(NOPOST_PREFIX_) === 0) {
    noPost = parts[0].trim();
    idx = 1;
  }

  const tanggal = (parts[idx] || "").trim();
  const jumlah = parseRupiahText_(parts[idx + 1]);
  const metodeTeks = parts.length > idx + 2 ? parts.slice(idx + 2).join(" - ").trim() : "";
  // ⬅️ BARU: "ADJUSTMENT (pembulatan)" dipecah -> metode "ADJUSTMENT" + keterangan "pembulatan".
  // metodeTeks = teks asli apa adanya (dipakai backfill supaya keterangan tidak hilang).
  const pecah = pecahMetodeBayar_(metodeTeks);
  return { tanggal: tanggal, jumlah: jumlah, metode: pecah.metode, keterangan: pecah.keterangan, metodeTeks: metodeTeks, noPost: noPost };
}

/* ------------------------------------------------------
   3) CARI DATA INVOICE DARI SHEET PENJUALAN
   ------------------------------------------------------ */
function cariDataInvoice(noInvoice) {
  if (!noInvoice || String(noInvoice).trim() === "") {
    return { found: false, message: "No Invoice tidak boleh kosong." };
  }

  const trimmed = String(noInvoice).trim();
  const formatValid = /^(INV|SPRIN)-\d{4}-\d{2}-\d{3}$/.test(trimmed);
  if (!formatValid) {
    return { found: false, message: "Format No Invoice salah. Gunakan format: INV-2026-07-001" };
  }

  // ⬅️ MIGRASI (Oktober 2026): dibaca FRESH dari PENJUALAN di database
  // (PENJUALAN lokal cuma salinan yang bisa telat beberapa detik setelah
  // ada pembayaran/biaya tambahan baru).
  let data;
  try {
    data = bacaPenjualanDB_();
  } catch (err) {
    return { found: false, message: "Gagal membaca PENJUALAN dari database: " + err.message };
  }
  if (data.length === 0) return { found: false, message: "Data PENJUALAN masih kosong." };
  data = data.map(function (r) {
    const row = r.slice();
    while (row.length < PENJUALAN_COL_CICILAN_START) row.push("");
    return row;
  });

  let idx = -1;
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][PENJUALAN_COL_NOINV - 1] || "").trim() === trimmed) {
      idx = i;
      break;
    }
  }

  if (idx === -1) {
    return { found: false, message: "Invoice " + trimmed + " tidak ditemukan di PENJUALAN." };
  }

  let namaToko = String(data[idx][PENJUALAN_COL_CUSTOMER - 1] || "").trim();
  let scan = idx;
  while (!namaToko && scan > 0) {
    scan--;
    namaToko = String(data[scan][PENJUALAN_COL_CUSTOMER - 1] || "").trim();
  }

  const totalDasar = Number(data[idx][PENJUALAN_COL_TOTAL - 1]) || 0;   // E
  const ekspedisi = Number(data[idx][PENJUALAN_COL_EKSPEDISI - 1]) || 0; // H (cuma buat ditampilin di dialog)
  const palet = Number(data[idx][PENJUALAN_COL_PALET - 1]) || 0;         // I (cuma buat ditampilin di dialog)
  const total = totalDasar; // ⬅️ G = E doang, gak ikut H & I lagi

  const statusRaw = String(data[idx][PENJUALAN_COL_STATUS - 1] || "").trim().toUpperCase();
  const sisaCell = data[idx][PENJUALAN_COL_SISA - 1]; // G

  const sisaTagihan = statusRaw ? Number(sisaCell) || 0 : total;
  const totalSudahDibayar = total - sisaTagihan;
  const status = statusRaw === "LUNAS" ? "Lunas" : "Cicilan";

  const riwayatPembayaran = [];
  for (let c = PENJUALAN_COL_CICILAN_START - 1; c < data[idx].length; c++) {
    const parsed = parseRiwayatCell_(data[idx][c]);
    if (parsed) riwayatPembayaran.push(parsed);
  }

  return {
    found: true,
    rowIndex: idx + 2,
    noInvoice: String(data[idx][PENJUALAN_COL_NOINV - 1]).trim(),
    namaToko: namaToko,
    totalTagihan: total,
    totalSudahDibayar: totalSudahDibayar,
    sisaTagihan: sisaTagihan,
    status: status,
    riwayatPembayaran: riwayatPembayaran,
    ekspedisi: ekspedisi,
    palet: palet,
  };
}

/* ------------------------------------------------------
   4) SIMPAN PEMBAYARAN BARU KE SHEET PENJUALAN
   ⬅️ FIX: sekarang dibungkus LockService (pola yang sama kayak
   simpanBiayaTambahan() di BiayaTambahan.gs). Sebelumnya kalau 2 orang
   posting pembayaran hampir bersamaan (atau admin klik submit 2x saking
   loading-nya lama), dua eksekusi bisa sama-sama baca targetCol/sisa
   SEBELUM salah satu selesai nulis -> keduanya nulis ke kolom cicilan
   yang sama / hitungan sisa bayar jadi salah -> KEJADIAN "1 invoice
   ke-post 2x" yang dilaporkan. Dengan lock, satu proses posting harus
   selesai TOTAL dulu (baca sisa -> generate No Post -> tulis) sebelum
   proses posting berikutnya boleh mulai baca datanya sendiri.
   ------------------------------------------------------ */
function prosesPembayaran(data) {
  // ⬅️ MIGRASI (Oktober 2026): lock lokal tidak diperlukan lagi -- gateway
  // database sudah mengunci SEMUA aksi tulis (LockService di
  // prosesRequestGateway_), jadi dua posting yang hampir bersamaan tetap
  // diproses satu per satu di sana.
  return prosesPembayaran_(data);
}

// ⬅️ MIGRASI (Oktober 2026): penulisan pembayaran ke PENJUALAN (STATUS,
// SISA BAYAR, No Post, riwayat cicilan, warna) dipindah ke aksi gateway
// "prosesPembayaranPenjualan" (sudah ada di DatabaseGateway.js, port persis
// dari logic lama di sini). Validasi input & bentuk hasil ke popup tetap sama.
function prosesPembayaran_(data) {
  const noInvoice = String(data.noInvoice || "").trim();
  const jumlahBayar = Number(data.jumlahBayar);
  const metodeBayar = String(data.metodeBayar || "").trim().toUpperCase();
  const tglBayarText = formatTglBayarInput_(data.tglBayar);

  if (!jumlahBayar || jumlahBayar <= 0) {
    return { success: false, message: "Jumlah bayar tidak valid." };
  }

  if (metodeBayar !== "TUNAI" && metodeBayar !== "TRANSFER" && metodeBayar !== "ADJUSTMENT") {
    return { success: false, message: "Metode pembayaran harus TUNAI atau TRANSFER atau ADJUSTMENT." };
  }

  // ⬅️ BARU: ADJUSTMENT wajib punya keterangan (mis. pembulatan / biaya admin / sparepart)
  const ketAdj = validasiKeteranganAdjustment_(metodeBayar, data.keteranganAdjustment);
  if (!ketAdj.ok) {
    return { success: false, message: ketAdj.message };
  }
  const metodeTeks = susunTeksMetodeBayar_(metodeBayar, ketAdj.keterangan);

  if (!tglBayarText) {
    return { success: false, message: "Tanggal pembayaran tidak valid." };
  }

  let hasil;
  try {
    hasil = kirimAksiGateway_("prosesPembayaranPenjualan", {
      noInvoice: noInvoice,
      jumlahBayar: jumlahBayar,
      metodeBayar: metodeBayar,
      keteranganAdjustment: ketAdj.keterangan, // ⬅️ BARU: wajib utk ADJUSTMENT (divalidasi lagi di gateway)
      tglBayarText: tglBayarText
    });
  } catch (err) {
    return { success: false, message: String(err.message || err).replace(/^Gateway menolak: /, "") };
  }
  refreshSalinan_(["PENJUALAN"]);

  if (hasil.jadiLunas) {
    try {
      tandaiDoneByNoInv_(hasil.noInvoice);
    } catch (err) {
      Logger.log("Auto-tandai DONE (dari prosesPembayaran) error: " + err);
    }

    try {
      tandaiDoneSparepartByNoInv_(hasil.noInvoice);
    } catch (err) {
      Logger.log("Auto-tandai DONE SPAREPART (dari prosesPembayaran) error: " + err);
    }
  }

  // ⬅️ PENGAMAN: kalau DatabaseGateway.js belum versi baru, pembayaran bisa tersimpan TANPA keterangan.
  // Jangan diam-diam: beri tahu admin supaya segera memperbarui gateway.
  if (metodeBayar === "ADJUSTMENT" && String(hasil.keteranganAdjustment || "") !== ketAdj.keterangan) {
    return { success: false, message: "Pembayaran TERSIMPAN, tetapi keterangan adjustment BELUM ikut tersimpan karena DatabaseGateway.js belum diperbarui. Perbarui & deploy gateway versi baru, lalu cek riwayat invoice ini." };
  }

  return {
    success: true,
    noInvoice: hasil.noInvoice,
    namaToko: hasil.namaToko,
    totalTagihan: hasil.totalTagihan,
    totalSudahDibayar: hasil.totalSudahDibayar,
    sisaTagihan: hasil.sisaTagihan,
    status: hasil.status,
    noPost: hasil.noPost,
    riwayatPembayaran: (hasil.riwayatPembayaran || []).map(function (r) {
      const pecah = pecahMetodeBayar_(r && r.metode);
      return Object.assign({}, r, { metode: pecah.metode, keterangan: pecah.keterangan });
    }),
  };
}

/**
 * ⬅️ REPAIR SEKALI JALAN — kasih No Post ke SEMUA riwayat cicilan LAMA di
 * PENJUALAN yang belum punya format baru "POST-NNN - ..." (termasuk cell
 * yang sempat dites pakai format lama "POST-yyyy-MM-NNN - ..." sebelum
 * formatnya disederhanain -- itu juga ikut dirapikan/dinomori ulang di sini,
 * BUKAN dipertahankan formatnya).
 *
 * URUTAN PENOMORAN: berdasarkan TANGGAL BAYAR di tiap cicilan (paling lama
 * duluan = POST-001), BUKAN posisi baris/kolom di sheet -- soalnya posisi
 * baris pernah kena bug urut-ulang/grouping sebelumnya jadi gak bisa
 * dijadiin acuan urutan waktu. Kalau ada beberapa cicilan tanggalnya SAMA
 * PERSIS, urutan di antara mereka ngikutin urutan baris/kolom saat ini
 * (stabil, gak diacak).
 *
 * TANGGAL, NOMINAL, DAN METODE ASLI TIDAK BERUBAH SAMA SEKALI -- dibaca
 * ulang lewat parseRiwayatCell_, ditulis ulang APA ADANYA cuma dengan
 * "POST-XXX - " ditambahkan di depan.
 *
 * Setelah selesai, counter No Post di SYSTEM (B4) otomatis di-set ke nomor
 * terakhir yang dipakai backfill ini, jadi postingan BARU ke depannya lewat
 * popup Pembayaran otomatis lanjut dari situ (gak akan nabrak).
 *
 * SEBELUM menulis apa pun, muncul ringkasan (berapa cell bakal diberi
 * nomor) dan minta konfirmasi Ya/Tidak dulu.
 *
 * Cara pakai: Apps Script Editor -> pilih function ini di dropdown atas ->
 * Run. Sekali aja.
 */
function backfillNoPostDataLama() {
  // ⬅️ MIGRASI (Oktober 2026): PENJUALAN lokal sekarang salinan IMPORTRANGE.
  SpreadsheetApp.getUi().alert("Tidak dipakai lagi: PENJUALAN sekarang salinan dari database_SANFELO. Penomoran No Post dilakukan otomatis di database.");
  return;
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shPenjualan = ss.getSheetByName("PENJUALAN");
  if (!shPenjualan) throw new Error('Sheet "PENJUALAN" tidak ditemukan.');

  const lastRow = shPenjualan.getLastRow();
  const lastCol = shPenjualan.getLastColumn();
  if (lastRow < 2 || lastCol < PENJUALAN_COL_CICILAN_START) {
    ui.alert("Tidak ada riwayat cicilan di PENJUALAN.");
    return;
  }

  const numCicilanCols = lastCol - PENJUALAN_COL_CICILAN_START + 1;
  const data = shPenjualan
    .getRange(2, PENJUALAN_COL_CICILAN_START, lastRow - 1, numCicilanCols)
    .getValues();

  const formatBaruRegex = /^POST-\d+ - /; // udah sesuai format baru, gak usah disentuh/dinomori ulang

  const candidates = [];
  for (let r = 0; r < data.length; r++) {
    for (let c = 0; c < data[r].length; c++) {
      const str = String(data[r][c] || "").trim();
      if (!str) continue;
      if (formatBaruRegex.test(str)) continue; // sudah format baru, skip

      const parsed = parseRiwayatCell_(str);
      if (!parsed) continue; // bukan format riwayat yang dikenali, biarkan apa adanya

      candidates.push({
        r: r,
        c: c,
        tanggal: parsed.tanggal,
        tglObj: parseTanggalDDMMYYYY_(parsed.tanggal),
        jumlah: parsed.jumlah,
        metode: parsed.metodeTeks, // ⬅️ teks asli (termasuk keterangan adjustment)
      });
    }
  }

  if (candidates.length === 0) {
    ui.alert("Tidak ada riwayat cicilan lama yang perlu diberi No Post -- semuanya sudah punya.");
    return;
  }

  // Urutkan berdasarkan tanggal bayar, paling lama duluan. Array.sort() di
  // Apps Script (V8) STABIL, jadi kalau tanggalnya identik, urutan asli
  // (baris lalu kolom) dipertahankan -- gak acak.
  candidates.sort(function (a, b) {
    const ta = a.tglObj ? a.tglObj.getTime() : 0;
    const tb = b.tglObj ? b.tglObj.getTime() : 0;
    return ta - tb;
  });

  const totalBaru = Utilities.formatString("%03d", candidates.length);
  const resp = ui.alert(
    "Konfirmasi Backfill No Post",
    candidates.length + " riwayat cicilan lama akan diberi nomor POST-001 s/d POST-" + totalBaru +
      ", diurutkan dari tanggal bayar PALING LAMA ke paling baru.\n\n" +
      "Tanggal, nominal, dan metode pembayaran TIDAK akan berubah sama sekali -- cuma nomor posting " +
      "yang ditambahkan di depan tiap riwayat.\n\nLanjutkan?",
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) {
    ui.alert("Dibatalkan. Tidak ada yang diubah.");
    return;
  }

  // Tulis balik SEKALIGUS (1x setValues untuk seluruh area cicilan), bukan
  // satu-satu per cell -- biar cepat walau jumlah cell-nya ratusan (lihat
  // catatan performa di repairHapusDuplikatPenjualan, PENJUALAN.gs).
  candidates.forEach(function (item, i) {
    const noPost = NOPOST_PREFIX_ + Utilities.formatString("%03d", i + 1);
    data[item.r][item.c] =
      noPost + " - " + item.tanggal + " - " + formatRupiahText_(item.jumlah) + " - " + item.metode;
  });

  shPenjualan.getRange(2, PENJUALAN_COL_CICILAN_START, lastRow - 1, numCicilanCols).setValues(data);

  const system = ss.getSheetByName("SYSTEM");
  if (system) {
    system.getRange(SYSTEM_NOPOST_CELL_COUNTER_).setValue(candidates.length);
  }

  ui.alert(
    "Selesai. " + candidates.length + " riwayat cicilan lama sudah diberi No Post (POST-001 s/d POST-" +
      totalBaru + ").\nPostingan baru ke depannya lewat popup Pembayaran otomatis lanjut dari POST-" +
      Utilities.formatString("%03d", candidates.length + 1) + "."
  );
}

/* ------------------------------------------------------
   5) TANDAIN DONE + HIJAU DI SPAREPART SAAT LUNAS

   ⬅️ BARU (migrasi database_SANFELO, September 2026): SPAREPART sekarang
   cuma mirror IMPORTRANGE di sisi inventory, jadi ini cuma 1x panggilan
   ke gateway (aksi "tandaiDoneSparepart"), yang jalan terhadap
   SPAREPART asli di database. Sekalian ketahuan & DIPERBAIKI bug lama:
   versi lokal ini dulu nulis STATUS ORDER ke kolom 18 (R) -- padahal
   kolom 18 itu LABA, STATUS ORDER yang benar itu kolom 19 (S). Jadi
   dulu diam-diam nimpa angka LABA dengan teks "DONE" tiap invoice
   sparepart lunas. Di gateway (aksiTandaiDoneSparepart_) sudah ditulis
   ke kolom yang benar.
   ------------------------------------------------------ */
function tandaiDoneSparepartByNoInv_(noInvoice) {
  const noInvTarget = String(noInvoice || "").trim();
  if (!noInvTarget) return;

  kirimAksiGateway_("tandaiDoneSparepart", { noInvoice: noInvTarget });
}

  // ======================= StatusDone.js (salinan persis) =======================
/* ======================================================
   StatusDone.gs
   - Nandain status DONE + warna hijau di TRANSAKSI (kolom N)
     dan DETAIL_TRANSAKSI (kolom T) begitu invoice terkait LUNAS
     di sheet PENJUALAN.
   - Proteksi: kalau user manual ubah dropdown N/T jadi DONE padahal
     PENJUALAN belum LUNAS, otomatis direvert (lihat onEdit gabungan
     di TempoToko.gs).
   ====================================================== */

const TRANSAKSI_COL_NOINV_STATUS = 3;   // C, dipakai buat lookup NO INV
const TRANSAKSI_COL_STATUS = 14;        // N (dropdown: DRAFT / PROGRESS / REVISI / DONE)
const TRANSAKSI_WARNA_DONE = "#d9ead3";

const DETAIL_COL_NOINV_STATUS = 3;      // C, dipakai buat lookup NO INV (⬅️ UPDATE: NO INV geser B->C, dipakai juga di TempoToko.js baris onEditDetailStatus_)
const DETAIL_COL_STATUS = 20;           // T (dropdown: PO / DONE)
const DETAIL_WARNA_DONE = "#d9ead3";

/**
 * Cek status LUNAS/CICIL suatu NO INV langsung dari sheet PENJUALAN.
 * Return "LUNAS" / "CICIL" / "" (gak ketemu).
 */
function getStatusPenjualanByNoInv_(noInvoice) {
  const trimmed = String(noInvoice || "").trim();
  if (!trimmed) return "";

  // ⬅️ MIGRASI (Oktober 2026): dibaca FRESH dari PENJUALAN di database
  const data = bacaSheetGateway_("PENJUALAN", 2, PENJUALAN_COL_NOINV, PENJUALAN_COL_STATUS - PENJUALAN_COL_NOINV + 1); // C..F

  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0] || "").trim() === trimmed) {
      return String(data[i][PENJUALAN_COL_STATUS - PENJUALAN_COL_NOINV] || "").trim().toUpperCase();
    }
  }
  return "";
}

/**
 * ⬅️ BARU: jalankan MANUAL lewat Apps Script editor (pilih function ini di
 * dropdown, klik Run) buat backfill data lama — scan SEMUA NO INV di
 * PENJUALAN yang statusnya udah LUNAS, terus tandain DONE + hijau di
 * TRANSAKSI & DETAIL_TRANSAKSI kalau belum ke-update.
 * Aman dijalankan berkali-kali (idempotent, gak dobel nulis).
 */
function tandaiDoneSemuaLunas() {
  // ⬅️ MIGRASI (Oktober 2026): PENJUALAN dibaca FRESH dari database, lalu
  // semua NO INV yang LUNAS dikirim SEKALIGUS ke aksi gateway
  // "tandaiDoneTransaksi" (TRANSAKSI & DETAIL_TRANSAKSI ada di database).
  const data = bacaSheetGateway_("PENJUALAN", 2, PENJUALAN_COL_NOINV, PENJUALAN_COL_STATUS - PENJUALAN_COL_NOINV + 1); // C..F
  if (data.length === 0) {
    SpreadsheetApp.getUi().alert("Data PENJUALAN masih kosong.");
    return;
  }

  const noInvLunasList = [];
  data.forEach(function (r) {
    const noInv = String(r[0] || "").trim();
    const status = String(r[PENJUALAN_COL_STATUS - PENJUALAN_COL_NOINV] || "").trim().toUpperCase();
    if (noInv && status === "LUNAS") {
      noInvLunasList.push(noInv);
    }
  });

  if (noInvLunasList.length > 0) {
    kirimAksiGateway_("tandaiDoneTransaksi", { noInvs: noInvLunasList });
    refreshSalinan_(["TRANSAKSI", "DETAIL_TRANSAKSI"]);
  }

  SpreadsheetApp.getUi().alert(
    "Selesai. " + noInvLunasList.length + " invoice LUNAS diproses (TRANSAKSI & DETAIL_TRANSAKSI ditandai DONE)."
  );
}

/**
 * Dipanggil dari prosesPembayaran() begitu status invoice jadi LUNAS.
 * Update TRANSAKSI (kolom N) dan SEMUA baris item di DETAIL_TRANSAKSI
 * (kolom T) yang NO INV-nya cocok.
 * ⬅️ MIGRASI (Oktober 2026): sekarang lewat aksi gateway "tandaiDoneTransaksi"
 * (database yang menulis status DONE + warna hijau).
 */
function tandaiDoneByNoInv_(noInvoice) {
  const trimmed = String(noInvoice || "").trim();
  if (!trimmed) return;

  kirimAksiGateway_("tandaiDoneTransaksi", { noInvs: [trimmed] });
  refreshSalinan_(["TRANSAKSI", "DETAIL_TRANSAKSI"]);
}

/**
 * ⬅️ FIX (performa): helper batching dipakai bareng oleh tandaiDoneByNoInv_
 * utk TRANSAKSI & DETAIL_TRANSAKSI. "sortedRows" = daftar rowIdx yang match
 * (udah pasti urut naik karena hasil scan linear dari atas ke bawah), tiap
 * rentang baris yang nyambung ditulis SEKALI lewat setValues()/
 * setBackgrounds(), bukan satu-satu per baris.
 */
function tandaiDoneRentangBaris_(sheet, sortedRows, colStatus, warnaDone) {
  if (!sortedRows || sortedRows.length === 0) return;

  let idx = 0;
  while (idx < sortedRows.length) {
    const startRow = sortedRows[idx];
    let endRow = startRow;
    let next = idx + 1;
    while (next < sortedRows.length && sortedRows[next] === endRow + 1) {
      endRow = sortedRows[next];
      next++;
    }

    const numRows = endRow - startRow + 1;
    const statusValues = [];
    const warnaGrid = [];
    for (let r = 0; r < numRows; r++) {
      statusValues.push(["DONE"]);
      const rowWarna = [];
      for (let c = 0; c < colStatus; c++) rowWarna.push(warnaDone);
      warnaGrid.push(rowWarna);
    }

    sheet.getRange(startRow, colStatus, numRows, 1).setValues(statusValues);
    sheet.getRange(startRow, 1, numRows, colStatus).setBackgrounds(warnaGrid);

    idx = next;
  }
}

/**
 * Bersihin warna 1 baris (dipakai waktu status diganti manual jadi
 * SELAIN DONE, atau waktu revert karena PENJUALAN belum LUNAS).
 */
function bersihkanWarnaBarisTransaksi_(trx, rowIdx) {
  trx.getRange(rowIdx, 1, 1, TRANSAKSI_COL_STATUS).setBackground(null);
}

function bersihkanWarnaBarisDetail_(det, rowIdx) {
  det.getRange(rowIdx, 1, 1, DETAIL_COL_STATUS).setBackground(null);
}

  // ======================= PENJUALAN.js (salinan persis) =======================
/**
 * ====== SYNC PENJUALAN DARI TRANSAKSI (AUTO, VIA onEdit) ======
 *
 * Mapping kolom:
 *   TRANSAKSI!D (TANGGAL)  -> PENJUALAN!A (TGL KIRIM)
 *   TRANSAKSI!E (CUSTOMER) -> PENJUALAN!B (CUSTOMER)
 *   TRANSAKSI!C (NO INV)   -> PENJUALAN!C (NO INV)
 *   TRANSAKSI!J (METHOD)   -> PENJUALAN!D (METHOD)
 *   TRANSAKSI!M (TOTAL)    -> PENJUALAN!E (TOTAL)
 *
 * Cara kerja tiap sync:
 * 1. UPSERT: invoice yang NO INV-nya sudah ada di PENJUALAN -> kolom
 *    A/D/E di-update di baris yang sama, kolom F ke kanan (data
 *    pembayaran) SAMA SEKALI TIDAK DISENTUH. Invoice baru -> ditambah
 *    di baris paling bawah, kolom F ke kanan dikosongkan.
 * 2. GROUPING: semua baris disusun ulang per customer (grup customer
 *    diurutkan dari kemunculan pertama), nama customer di kolom B
 *    cuma ditulis di baris pertama tiap grup.
 * 3. BACKFILL SISA BAYAR: ⬅️ BARU — kolom G (SISA BAYAR) di-default-in
 *    = TOTAL (E) untuk semua baris yang belum pernah disentuh
 *    prosesPembayaran() sama sekali (STATUS kosong DAN G kosong).
 *    Berlaku buat invoice baru yang baru masuk MAUPUN baris lama yang
 *    kelewatan/belum keisi, karena fungsi ini dipanggil ulang tiap
 *    sync jalan. Baris yang udah CICIL/LUNAS tidak disentuh.
 * 4. RE-COLOR: setValues() cuma mindahin ISI SEL, background
 *    warna GAK IKUT PINDAH. Jadi abis grouping, baris yang tadinya
 *    hijau (LUNAS) bisa nyangkut di posisi lama sementara datanya udah
 *    pindah. Makanya SETIAP ABIS ADA PERUBAHAN SUSUNAN BARIS, warna
 *    di-cek ULANG dari kolom STATUS (F) tiap baris, bukan diwariskan
 *    dari posisi baris sebelumnya.
 * 5. Semua METHOD (CASH & CREDIT) tetap masuk, tanpa filter.
 *
 * KOLOM YANG DIPANTAU DI TRANSAKSI biar auto-sync gak jalan tiap sel
 * apapun diedit, cuma kolom relevan: C, D, E, J, M
 */

const TRANSAKSI_WATCH_COLS = [3, 4, 5, 10, 13]; // C, D, E, J, M

/**
 * Dipanggil dari menu "🔄 Sync Penjualan dari Transaksi" -> nunjukin alert.
 */
function syncPenjualanFromTransaksi() {
  syncPenjualanFromTransaksi_(true);
}

/**
 * Versi inti, dipakai bareng oleh menu (showAlert=true) dan
 * auto-trigger dari onEdit (showAlert=false, biar gak ganggu user
 * tiap ngetik).
 */
function syncPenjualanFromTransaksi_(showAlert) {
  // ⬅️ MIGRASI (Oktober 2026): TRANSAKSI dibaca FRESH dari database, dan
  // PENJUALAN ditulis lewat aksi gateway "upsertPenjualan" (sudah ada di
  // DatabaseGateway.js). Isi aksi itu = port persis dari logic lama di file
  // ini: UPSERT (kolom F ke kanan / data pembayaran tidak disentuh, SISA
  // BAYAR ikut menyesuaikan kalau TOTAL berubah), GROUPING per customer,
  // BACKFILL SISA BAYAR, AUTO-LUNAS, RE-COLOR -- semuanya jalan di database.
  const dataT = bacaTransaksiDB_(); // A..Q
  if (dataT.length === 0) {
    if (showAlert) SpreadsheetApp.getUi().alert("Tidak ada data di TRANSAKSI.");
    return;
  }

  const entries = [];
  const sudah = {};
  dataT.forEach(function (r) {
    const noInv = String(r[2] || "").trim(); // C
    if (!noInv) return;                       // lewati baris kosong
    const e = { noInv: noInv, tanggal: r[3], customer: r[4], method: r[9], total: r[12] }; // D, E, J, M
    if (sudah[noInv] !== undefined) {
      entries[sudah[noInv]] = e;              // NO INV dobel: pakai baris paling bawah (sama seperti versi lama)
    } else {
      sudah[noInv] = entries.length;
      entries.push(e);
    }
  });

  const hasil = kirimAksiGateway_("upsertPenjualan", { entries: entries, updateCustomer: false });
  refreshSalinan_(["PENJUALAN"]);

  // TEMPO TOKO ikut auto-sync, karena datanya nurunin dari PENJUALAN
  try {
    if (typeof syncTempoToko_ === "function") {
      syncTempoToko_(false); // silent, gak munculin alert
    }
  } catch (err) {
    Logger.log("Auto-sync TEMPO TOKO error: " + err);
  }

  if (showAlert) {
    SpreadsheetApp.getUi().alert(
      "✅ Sinkron selesai: " + (hasil.insertedCount || 0) + " invoice baru ditambahkan ke PENJUALAN."
    );
  }
}

/**
 * Susun ulang baris PENJUALAN biar customer yang sama nempel jadi satu
 * grup (urutan grup = urutan customer pertama kali muncul). Nama
 * customer di kolom B cuma ditulis di baris pertama tiap grup.
 *
 * PENTING: tiap baris digeser SEBAGAI SATU PAKET UTUH mulai dari
 * kolom A sampai kolom paling kanan yang kepake (dideteksi otomatis
 * pakai getLastColumn(), jadi berapa pun banyaknya kolom riwayat
 * cicilan di kanan F, semua ikut kebawa bareng baris yang benar).
 */
function groupPenjualanByCustomer_(shPenjualan) {
  // ✅ FIX (anti-duplikat): baris "TOTAL ..." virtual (kalau ada yang
  // kelewatan/nyangkut) dibersihkan LEBIH DULU pakai fungsi yang sudah ada
  // (removeVirtualTotalRowsPenjualan_ di FILTER_SIDEBAR.gs, satu project =
  // satu global scope) -- ini betul-betul MENGHAPUS barisnya dari sheet
  // (bukan cuma di-skip di memory), jadi jumlah baris fisik selalu match
  // sama jumlah baris yang mau ditulis ulang di bawah. Aman dipanggil
  // walau gak ada baris virtual sama sekali (langsung no-op).
  if (typeof removeVirtualTotalRowsPenjualan_ === "function") {
    removeVirtualTotalRowsPenjualan_(shPenjualan);
  }

  const lastRow = shPenjualan.getLastRow();
  const lastCol = shPenjualan.getLastColumn();
  if (lastRow < 2 || lastCol < 2) return;

  const data = shPenjualan.getRange(2, 1, lastRow - 1, lastCol).getValues();

  const customerOrder = [];
  const customerRows = {};
  let lastCustomer = "";

  data.forEach(function (row) {
    const rawCustomer = String(row[1] || "").trim();
    const customer = rawCustomer || lastCustomer;
    if (rawCustomer) lastCustomer = rawCustomer;
    // ✅ FIX (anti-duplikat): baris SEBELUMNYA di-skip kalau "customer"
    // kosong -> baris itu gak ikut ditulis balik ke flatRows, tapi baris
    // fisik lamanya di posisi bawah TETAP ADA (gak pernah dibersihkan) ->
    // jadi "nyangkut" sebagai duplikat stale. Sekarang baris tanpa
    // customer sama sekali (harusnya nyaris gak pernah kejadian) tetap
    // DIPERTAHANKAN dalam grup kuncinya "" (string kosong), bukan dibuang,
    // supaya jumlah baris SELALU sama persis sebelum & sesudah -> gak ada
    // baris lama yang bisa nyangkut jadi duplikat.
    if (!customerRows[customer]) {
      customerRows[customer] = [];
      customerOrder.push(customer);
    }
    customerRows[customer].push(row);
  });

  const flatRows = [];
  customerOrder.forEach(function (customer) {
    const rows = customerRows[customer];
    rows.forEach(function (row) {
      const newRow = row.slice(); // copy, jangan ubah array asli
      newRow[1] = customer; // ⬅️ diisi di SETIAP baris, gak cuma baris pertama grup
      flatRows.push(newRow);
    });
  });

  if (flatRows.length === 0) return;

  shPenjualan.getRange(2, 1, flatRows.length, lastCol).setValues(flatRows);
  shPenjualan.getRange(2, 1, flatRows.length, 1).setNumberFormat("dd/mm/yyyy"); // A: tgl kirim
  shPenjualan.getRange(2, 5, flatRows.length, 1).setNumberFormat("#,##0"); // E: total

  // ✅ FIX (anti-duplikat, jaring pengaman terakhir): kalau karena SEBAB
  // APAPUN jumlah baris hasil grouping (flatRows.length) ternyata masih
  // lebih sedikit dari jumlah baris asli (data.length) -- harusnya sudah
  // gak mungkin lagi setelah 2 perbaikan di atas -- sisa baris lama di
  // bawah DIBERSIHKAN ISINYA (bukan dibiarkan nyangkut jadi duplikat
  // senyap). Baris kosong jelas kelihatan aneh & gampang dicek manual;
  // baris duplikat yang keliatan sama persis dengan data asli JAUH lebih
  // berbahaya karena kebaca seolah data valid (dobel piutang/omzet).
  if (flatRows.length < data.length) {
    const sisaBaris = data.length - flatRows.length;
    shPenjualan.getRange(2 + flatRows.length, 1, sisaBaris, lastCol).clearContent();
  }
}

/**
 * ⬅️ BARU — default-kan kolom G (SISA BAYAR) = TOTAL (E) untuk semua
 * baris yang STATUS (F)-nya masih kosong DAN G-nya juga masih kosong.
 * Artinya: invoice yang belum pernah disentuh prosesPembayaran() sama
 * sekali (baik baru masuk lewat sync, maupun baris lama yang kelewatan)
 * akan otomatis nunjukin sisa piutang = total utuh, bukan kosong melompong.
 *
 * Baris yang udah punya STATUS (CICIL/LUNAS) TIDAK disentuh sama sekali,
 * biar progress pembayaran yang udah ada gak ketimpa.
 *
 * Dipanggil tiap habis sync (aman dijalankan berkali-kali / idempotent),
 * jadi berlaku otomatis buat invoice baru ke depannya juga, bukan cuma
 * sekali jalan buat data lama.
 */
function backfillSisaBayarPenjualan_(shPenjualan) {
  const lastRow = shPenjualan.getLastRow();
  if (lastRow < 2) return;

  const colTotal = typeof PENJUALAN_COL_TOTAL !== "undefined" ? PENJUALAN_COL_TOTAL : 5;
  const colStatus = typeof PENJUALAN_COL_STATUS !== "undefined" ? PENJUALAN_COL_STATUS : 6;
  const colSisa = typeof PENJUALAN_COL_SISA !== "undefined" ? PENJUALAN_COL_SISA : 7;
  const colEksp = typeof PENJUALAN_COL_EKSPEDISI !== "undefined" ? PENJUALAN_COL_EKSPEDISI : 8;
  const colPalet = typeof PENJUALAN_COL_PALET !== "undefined" ? PENJUALAN_COL_PALET : 9;

  const lastCol = Math.max(colSisa, colEksp, colPalet, colTotal, colStatus);
  const data = shPenjualan.getRange(2, 1, lastRow - 1, lastCol).getValues();

  const sisaBaru = data.map(function (row) {
    const total = Number(row[colTotal - 1]) || 0;
    const status = String(row[colStatus - 1] || "").trim();
    const sisa = row[colSisa - 1];
    const ekspedisi = Number(row[colEksp - 1]) || 0;
    const palet = Number(row[colPalet - 1]) || 0;

    if (!status && row[colTotal - 1] !== "" && row[colTotal - 1] !== null) {
      return [total]; // G = E doang, gak ikut H & I
    }
    return [sisa];
  });

  shPenjualan.getRange(2, colSisa, sisaBaru.length, 1).setValues(sisaBaru);
  shPenjualan.getRange(2, colSisa, sisaBaru.length, 1).setNumberFormat("#,##0");
}

/**
 * ⬅️ HELPER BARU — cek ulang STATUS (kolom F / PENJUALAN_COL_STATUS) di
 * SETIAP baris, lalu terapkan warna berdasarkan ISI SEL saat ini, bukan
 * ngandelin background yang mungkin nyangkut dari posisi baris
 * sebelumnya. Dipanggil setiap kali habis ada perubahan susunan baris
 * (setelah upsert + grouping), supaya baris LUNAS selalu ke-highlight
 * hijau di posisi yang BENAR, dan baris yang bukan LUNAS gak kebawa
 * warna hijau nyasar.
 *
 * Konstanta PENJUALAN_COL_STATUS & PENJUALAN_WARNA_LUNAS dipakai dari
 * Pembayaran.gs (satu project = satu global scope, gak perlu di-copy
 * ulang di sini). Kalau kedua konstanta itu belum ke-load (Pembayaran.gs
 * belum ada di project), fallback ke nilai default biar gak error.
 */
function reapplyWarnaLunasPenjualan_(shPenjualan) {
  const lastRow = shPenjualan.getLastRow();
  const lastCol = shPenjualan.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return;

  const colStatus = typeof PENJUALAN_COL_STATUS !== "undefined" ? PENJUALAN_COL_STATUS : 6;
  const warnaLunas = typeof PENJUALAN_WARNA_LUNAS !== "undefined" ? PENJUALAN_WARNA_LUNAS : "#d9ead3";

  if (colStatus > lastCol) return; // kolom STATUS belum ada isinya sama sekali

  const statusValues = shPenjualan.getRange(2, colStatus, lastRow - 1, 1).getValues();

  // ⬅️ FIX (performa): dulu setBackground() dipanggil SATU-SATU per baris
  // (sampai lastRow-1 kali, tiap kali lebar lastCol kolom). Fungsi ini
  // jalan tiap kali sync PENJUALAN-dari-TRANSAKSI (termasuk lewat onEdit
  // otomatis), jadi buat sheet yang udah banyak baris ini salah satu
  // penyebab utama lambatnya. Sekarang dikumpulin ke satu grid 2D,
  // ditulis SEKALI lewat setBackgrounds() buat seluruh range sekaligus --
  // hasil akhirnya identik, cuma 1 panggilan API bukan N panggilan.
  const numRows = statusValues.length;
  const warnaGrid = [];
  for (let i = 0; i < numRows; i++) {
    const status = String(statusValues[i][0] || "").trim().toUpperCase();
    const warnaBaris = status === "LUNAS" ? warnaLunas : null;
    const row = [];
    for (let c = 0; c < lastCol; c++) row.push(warnaBaris);
    warnaGrid.push(row);
  }

  if (numRows > 0) {
    shPenjualan.getRange(2, 1, numRows, lastCol).setBackgrounds(warnaGrid);
  }
}

/* ======================================================
   ⬅️ BARU — SYNC PENJUALAN DARI SPAREPART
   ======================================================
 *
 * Sumber: sheet SPAREPART (1 baris = 1 item barang, banyak baris bisa
 * share NO INV yang sama kalau 1 transaksi isi beberapa barang).
 *
 * Mapping kolom:
 *   SPAREPART!C (TANGGAL)    -> PENJUALAN!A (TGL KIRIM)
 *   SPAREPART!F (NAMA TOKO)  -> PENJUALAN!B (CUSTOMER)
 *   SPAREPART!B (NO INV)     -> PENJUALAN!C (NO INV)
 *   SPAREPART!O (METHOD)     -> PENJUALAN!D (METHOD)
 *   SPAREPART!M (HARGA)      -> PENJUALAN!E (TOTAL), DIJUMLAHKAN dulu
 *                                per NO INV yang sama (karena satu
 *                                invoice bisa punya banyak baris barang)
 *
 * Full rescan tiap dipanggil: baca SEMUA baris SPAREPART, agregasi per
 * NO INV, baru upsert ke PENJUALAN (kolom F ke kanan / data pembayaran
 * SAMA SEKALI TIDAK DISENTUH — sama seperti sync dari TRANSAKSI).
 * Logic grouping/backfill/recolor PENJUALAN dipakai bareng, gak diubah.
 *
 * Dipanggil dari finalizeSparepartSave_() di SPAREPART_SAVE.gs, silent
 * (gak munculin alert), tiap kali simpanSparepart() sukses nulis data.
 *
 * ⬅️ BARU (migrasi database_SANFELO, September 2026): SPAREPART sekarang
 * cuma mirror IMPORTRANGE di sisi inventory. Baca di sini pakai
 * bacaSheetGateway_ (FRESH langsung dari database, bukan mirror yang
 * bisa delay) — penting karena fungsi ini dipanggil LANGSUNG SETELAH
 * finalizeSparepartSave_() nulis data baru, jadi gak boleh nunggu
 * IMPORTRANGE catch up.
 */

// ⬅️ FIX: dipanggil dari menu "Sync Penjualan dari Sparepart". Kalau ada
// error, Google cuma nampilin "Skrip syncSparepart mengalami kesalahan"
// tanpa detail -- sekarang ditangkap & pesan aslinya (plus lokasi baris
// kodenya) ditampilin ke admin biar ketahuan gagal di mana.
function syncSparepart(){
  try {
    syncPenjualanFromSparepart_(true);
  } catch (err) {
    Logger.log("syncSparepart error: " + (err && err.stack ? err.stack : err));
    SpreadsheetApp.getUi().alert(
      "❌ Sync Penjualan dari Sparepart gagal:\n\n" +
      (err && err.message ? err.message : String(err)) +
      (err && err.stack ? "\n\nLokasi:\n" + String(err.stack).split("\n").slice(0, 4).join("\n") : "")
    );
  }
}

function syncPenjualanFromSparepart_(showAlert) {
  // ⬅️ MIGRASI (Oktober 2026): PENJUALAN ditulis lewat aksi gateway
  // "upsertPenjualan" dengan updateCustomer = true (NAMA TOKO di SPAREPART
  // bisa diedit ulang, jadi kolom B ikut ditimpa -- sama seperti versi lama).
  // Ambil kolom A..S (19 kolom) sekaligus, fresh dari database
  const dataSPR = bacaSheetGateway_("SPAREPART", 2, 1, 19);
  if (dataSPR.length === 0) {
    if (showAlert) SpreadsheetApp.getUi().alert("Tidak ada data di SPAREPART.");
    return;
  }
  // index kolom (0-based dari A):
  // 0:A(NO SJ) 1:B(NO INV) 2:C(REFRENSI) 3:D(TANGGAL) 4:E(TEMPO)
  // 5:F(NO KEND) 6:G(NAMA TOKO) 7:H(ALAMAT) ... 13:N(HARGA) ... 15:P(METHOD)

  const sparepartMap = {}; // noInv -> {tanggal, customer, method, total}
  const orderList = [];

  dataSPR.forEach(function (r) {
    const noInv = String(r[1] || "").trim(); // B
    if (!noInv) return; // lewati baris kosong

    const harga = Number(r[13]) || 0; // N = HARGA

    if (!sparepartMap[noInv]) {
      sparepartMap[noInv] = {
        tanggal: r[3],   // D = TANGGAL
        customer: r[6],  // G = NAMA TOKO
        method: r[15],   // P = METHOD
        total: 0,
      };
      orderList.push(noInv);
    }
    sparepartMap[noInv].total += harga; // jumlahkan HARGA semua baris item dgn NO INV sama
  });

  const entries = orderList.map(function (noInv) {
    const t = sparepartMap[noInv];
    return { noInv: noInv, tanggal: t.tanggal, customer: t.customer, method: t.method, total: t.total };
  });

  const hasil = kirimAksiGateway_("upsertPenjualan", { entries: entries, updateCustomer: true });
  refreshSalinan_(["PENJUALAN"]);

  // ⬅️ BARU: TEMPO TOKO ikut di-sync (sama kayak syncPenjualanFromTransaksi_)
  // -- invoice SPRIN yang baru masuk PENJUALAN (dan status LUNAS-nya)
  // langsung kebaca di TEMPO TOKO, gak nunggu sync berikutnya.
  try {
    if (typeof syncTempoToko_ === "function") {
      syncTempoToko_(false); // silent, gak munculin alert
    }
  } catch (err) {
    Logger.log("Auto-sync TEMPO TOKO (dari SPAREPART) error: " + err);
  }

  if (showAlert) {
    SpreadsheetApp.getUi().alert(
      "✅ Sinkron PENJUALAN dari SPAREPART selesai: " + (hasil.insertedCount || 0) + " invoice baru ditambahkan."
    );
  }
}

function autoLunaskanSisaBayarNol_(shPenjualan) {
  const lastRow = shPenjualan.getLastRow();
  if (lastRow < 2) return;

  const colStatus = typeof PENJUALAN_COL_STATUS !== "undefined" ? PENJUALAN_COL_STATUS : 6;
  const colSisa = typeof PENJUALAN_COL_SISA !== "undefined" ? PENJUALAN_COL_SISA : 7;

  const lastCol = Math.max(colStatus, colSisa);
  const data = shPenjualan.getRange(2, 1, lastRow - 1, lastCol).getValues();

  const statusBaru = data.map(function (row) {
    const sisaRaw = row[colSisa - 1];
    const sisaKosong = sisaRaw === "" || sisaRaw === null;

    if (!sisaKosong && Number(sisaRaw) === 0) {
      return ["LUNAS"];
    }
    return [row[colStatus - 1]]; // biarkan nilai asli, gak diubah
  });

  shPenjualan.getRange(2, colStatus, statusBaru.length, 1).setValues(statusBaru);
}

function lunaskanSisaBayarNolManual() {
  // ⬅️ MIGRASI (Oktober 2026): PENJUALAN lokal sekarang salinan IMPORTRANGE,
  // perbaikan massal semacam ini harus dijalankan langsung di database.
  SpreadsheetApp.getUi().alert("Tidak dipakai lagi: PENJUALAN sekarang salinan dari database_SANFELO. Perbaikan data PENJUALAN dilakukan langsung di database.");
  return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shPenjualan = ss.getSheetByName("PENJUALAN");
  if (!shPenjualan) throw new Error('Sheet "PENJUALAN" tidak ditemukan.');

  autoLunaskanSisaBayarNol_(shPenjualan);
  reapplyWarnaLunasPenjualan_(shPenjualan);

  SpreadsheetApp.getUi().alert("Selesai: baris dengan sisa bayar 0 sudah di-set LUNAS dan diwarnai hijau.");
}

/**
 * ⬅️ REPAIR SEKALI JALAN — reset kolom G (SISA BAYAR) jadi = E (TOTAL)
 * untuk data LAMA, cuma di baris yang STATUS (F)-nya masih kosong
 * (belum pernah disentuh prosesPembayaran()).
 *
 * Baris yang STATUS-nya sudah CICIL/LUNAS TIDAK disentuh sama sekali.
 *
 * Ini gak nyentuh kolom lain (A/B/C/D/E) sama sekali, cuma G doang.
 * Aman dijalankan berkali-kali (idempotent).
 *
 * Cara pakai: buka Apps Script editor, pilih function ini di dropdown
 * atas, klik Run. Sekali aja.
 */
function repairSisaBayarJadiTotalSaja() {
  // ⬅️ MIGRASI (Oktober 2026): PENJUALAN lokal sekarang salinan IMPORTRANGE,
  // perbaikan massal semacam ini harus dijalankan langsung di database.
  SpreadsheetApp.getUi().alert("Tidak dipakai lagi: PENJUALAN sekarang salinan dari database_SANFELO. Perbaikan data PENJUALAN dilakukan langsung di database.");
  return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shPenjualan = ss.getSheetByName("PENJUALAN");
  if (!shPenjualan) throw new Error('Sheet "PENJUALAN" tidak ditemukan.');

  const lastRow = shPenjualan.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert("Tidak ada data di PENJUALAN.");
    return;
  }

  const colTotal = typeof PENJUALAN_COL_TOTAL !== "undefined" ? PENJUALAN_COL_TOTAL : 5;   // E
  const colStatus = typeof PENJUALAN_COL_STATUS !== "undefined" ? PENJUALAN_COL_STATUS : 6; // F
  const colSisa = typeof PENJUALAN_COL_SISA !== "undefined" ? PENJUALAN_COL_SISA : 7;       // G

  const lastCol = Math.max(colTotal, colStatus, colSisa);
  const data = shPenjualan.getRange(2, 1, lastRow - 1, lastCol).getValues();

  let totalDiperbaiki = 0;

  const sisaBaru = data.map(function (row) {
    const total = row[colTotal - 1];
    const status = String(row[colStatus - 1] || "").trim();

    // cuma sentuh baris yang belum pernah dicicil/lunas sama sekali
    if (!status && total !== "" && total !== null) {
      totalDiperbaiki++;
      return [total]; // G = E
    }
    return [row[colSisa - 1]]; // biarkan apa adanya
  });

  shPenjualan.getRange(2, colSisa, sisaBaru.length, 1).setValues(sisaBaru);
  shPenjualan.getRange(2, colSisa, sisaBaru.length, 1).setNumberFormat("#,##0");

  SpreadsheetApp.getUi().alert(
    "Selesai. " + totalDiperbaiki + " baris (STATUS kosong) sudah di-reset jadi SISA BAYAR = TOTAL."
  );
}

/**
 * ⬅️ REPAIR SEKALI JALAN — hapus baris DUPLIKAT di PENJUALAN (No INV yang
 * sama muncul di lebih dari 1 baris), hasil dari bug lama di
 * groupPenjualanByCustomer_ / sortPenjualanBySeries_ (baris "ketinggalan"
 * di bawah pas grouping/sorting jalan tapi jumlah barisnya mengecil tanpa
 * baris sisanya dibersihkan). Sudah diperbaiki permanen di kedua fungsi
 * itu (lihat komentar ✅ FIX di groupPenjualanByCustomer_ & FILTER_SIDEBAR.gs)
 * -- function ini KHUSUS buat beresin data yang SUDAH kadung dobel.
 *
 * ATURAN PEMILIHAN BARIS YANG DIPERTAHANKAN (supaya TIDAK ADA satupun
 * riwayat pembayaran yang hilang):
 * 1) Untuk tiap No INV yang muncul > 1 kali: kalau HANYA SATU baris di
 *    antara duplikatnya yang punya "jejak pembayaran" (STATUS terisi,
 *    SISA BAYAR beda dari TOTAL, EKSPEDISI/PALET terisi, atau ada riwayat
 *    cicilan di kolom J dst), baris ITU yang dipertahankan, sisanya
 *    (yang kosong/default) dihapus.
 * 2) Kalau LEBIH dari satu baris duplikat sama-sama punya jejak
 *    pembayaran tapi ISINYA BEDA satu sama lain -> KELOMPOK INI DILEWATI
 *    SAMA SEKALI (tidak ada yang dihapus), dilaporkan di ringkasan supaya
 *    dicek manual satu-satu (jangan sampai salah pilih).
 * 3) Kalau SEMUA baris duplikat sama-sama tidak ada jejak pembayaran,
 *    yang dipertahankan adalah baris yang TOTAL (kolom E)-nya cocok
 *    dengan TRANSAKSI!M untuk No INV yang sama SAAT INI; kalau masih
 *    seri, baris PALING BAWAH yang dipertahankan.
 * 4) SEBELUM menghapus apa pun, ditampilkan dulu ringkasan lengkap (No
 *    INV mana saja, berapa baris akan dihapus, No INV mana yang dilewati
 *    karena ambigu) dan MINTA KONFIRMASI Ya/Tidak dulu.
 * 5) Cuma menghapus baris yang stale/kelebihan -- TIDAK PERNAH mengubah
 *    isi sel apa pun di baris yang dipertahankan, dan TIDAK PERNAH
 *    menyentuh sheet TRANSAKSI/DETAIL_TRANSAKSI (data sumber dipakai
 *    cuma buat pembanding/pelaporan).
 *
 * Cara pakai: Apps Script Editor -> pilih function ini di dropdown atas
 * -> Run. Disarankan jalankan repairHapusBarisTotalPenjualan() (di
 * FILTER_SIDEBAR.gs) dulu sebelum ini, kalau-kalau masih ada baris
 * "TOTAL ..." virtual yang nyangkut.
 */
function repairHapusDuplikatPenjualan() {
  // ⬅️ MIGRASI (Oktober 2026): PENJUALAN lokal sekarang salinan IMPORTRANGE,
  // perbaikan massal semacam ini harus dijalankan langsung di database.
  SpreadsheetApp.getUi().alert("Tidak dipakai lagi: PENJUALAN sekarang salinan dari database_SANFELO. Perbaikan data PENJUALAN dilakukan langsung di database.");
  return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shPenjualan = ss.getSheetByName("PENJUALAN");
  const shTransaksi = ss.getSheetByName("TRANSAKSI");
  if (!shPenjualan) throw new Error('Sheet "PENJUALAN" tidak ditemukan.');

  const ui = SpreadsheetApp.getUi();

  const lastRow = shPenjualan.getLastRow();
  const lastCol = shPenjualan.getLastColumn();
  if (lastRow < 2) {
    ui.alert("Tidak ada data di PENJUALAN.");
    return;
  }

  // Total TRANSAKSI saat ini per No INV, cuma dipakai buat pembanding/pelaporan
  const totalTransaksiPerNoInv = {};
  if (shTransaksi) {
    const lastRowT = shTransaksi.getLastRow();
    if (lastRowT >= 2) {
      const dataT = shTransaksi.getRange(2, 3, lastRowT - 1, 11).getValues(); // C..M
      dataT.forEach(function (r) {
        const noInv = String(r[0] || "").trim(); // C = NO INV
        if (noInv) totalTransaksiPerNoInv[noInv] = Number(r[10]) || 0; // M = TOTAL
      });
    }
  }

  const data = shPenjualan.getRange(2, 1, lastRow - 1, lastCol).getValues();

  // Kelompokkan index baris (0-based di `data`) per No INV (kolom C = index 2)
  const groupsByNoInv = {};
  data.forEach(function (row, i) {
    const noInv = String(row[2] || "").trim();
    if (!noInv) return; // baris tanpa No INV, bukan urusan repair ini
    if (!groupsByNoInv[noInv]) groupsByNoInv[noInv] = [];
    groupsByNoInv[noInv].push(i);
  });

  function adaJejakPembayaran(row) {
    const status = String(row[5] || "").trim(); // F
    const total = row[4];                       // E
    const sisa = row[6];                        // G
    const eksp = Number(row[7]) || 0;            // H
    const palet = Number(row[8]) || 0;           // I
    const sisaBeda = sisa !== "" && sisa !== null && Number(sisa) !== Number(total);
    let cicilanAda = false;
    for (let c = 9; c < row.length; c++) { // J dst = riwayat cicilan
      if (row[c] !== "" && row[c] !== null && row[c] !== undefined) { cicilanAda = true; break; }
    }
    return !!status || sisaBeda || eksp !== 0 || palet !== 0 || cicilanAda;
  }

  function kunciRiwayat(row) {
    return JSON.stringify(row.slice(5)); // F dst, dipakai buat bandingin persis sama/beda
  }

  const rowsToDelete = []; // 1-based row index di sheet
  const ringkasanHapus = [];
  const dilewatiAmbigu = [];
  const totalMismatchReport = [];

  Object.keys(groupsByNoInv).forEach(function (noInv) {
    const idxList = groupsByNoInv[noInv];
    if (idxList.length < 2) return; // bukan duplikat

    const withRiwayat = idxList.filter(function (i) { return adaJejakPembayaran(data[i]); });
    let keepIdx;

    if (withRiwayat.length === 1) {
      keepIdx = withRiwayat[0];
    } else if (withRiwayat.length > 1) {
      const uniqueKeys = {};
      withRiwayat.forEach(function (i) { uniqueKeys[kunciRiwayat(data[i])] = true; });
      if (Object.keys(uniqueKeys).length > 1) {
        // riwayat pembayaran BEDA di lebih dari satu baris -> ambigu, jangan sentuh
        dilewatiAmbigu.push(noInv);
        return;
      }
      keepIdx = withRiwayat[withRiwayat.length - 1]; // riwayatnya identik, aman ambil yang paling bawah
    } else {
      // gak ada satupun yang punya jejak pembayaran -> cocokkan ke TOTAL TRANSAKSI saat ini
      const totalTrx = totalTransaksiPerNoInv[noInv];
      const cocok = idxList.filter(function (i) {
        return typeof totalTrx === "number" && Number(data[i][4]) === totalTrx;
      });
      keepIdx = cocok.length > 0 ? cocok[cocok.length - 1] : idxList[idxList.length - 1];
    }

    idxList.forEach(function (i) {
      if (i === keepIdx) return;
      rowsToDelete.push(i + 2); // balik ke nomor baris 1-based di sheet
      ringkasanHapus.push(noInv + " (baris " + (i + 2) + ")");
    });

    const totalTrx = totalTransaksiPerNoInv[noInv];
    if (typeof totalTrx === "number" && Number(data[keepIdx][4]) !== totalTrx) {
      totalMismatchReport.push(noInv + ": PENJUALAN=" + data[keepIdx][4] + " vs TRANSAKSI=" + totalTrx);
    }
  });

  if (rowsToDelete.length === 0 && dilewatiAmbigu.length === 0) {
    ui.alert("Tidak ditemukan baris duplikat No INV di PENJUALAN. Tidak ada yang perlu dihapus.");
    return;
  }

  let msg = "";
  if (rowsToDelete.length > 0) {
    msg += rowsToDelete.length + " baris duplikat (data lama/stale) akan DIHAPUS:\n"
      + ringkasanHapus.slice(0, 30).join("\n")
      + (ringkasanHapus.length > 30 ? "\n... dan " + (ringkasanHapus.length - 30) + " lainnya" : "")
      + "\n\n";
  }
  if (dilewatiAmbigu.length > 0) {
    msg += "⚠️ " + dilewatiAmbigu.length + " No INV DILEWATI (>1 baris dengan riwayat pembayaran yang BEDA isinya, perlu dicek manual satu-satu):\n"
      + dilewatiAmbigu.join(", ") + "\n\n";
  }
  if (totalMismatchReport.length > 0) {
    msg += "ℹ️ " + totalMismatchReport.length + " baris yang DIPERTAHANKAN TOTAL-nya beda dari TRANSAKSI saat ini (dibiarkan apa adanya, sekadar info, cek manual kalau perlu):\n"
      + totalMismatchReport.slice(0, 20).join("\n")
      + (totalMismatchReport.length > 20 ? "\n..." : "") + "\n\n";
  }

  if (rowsToDelete.length === 0) {
    ui.alert(msg + "Tidak ada baris yang dihapus otomatis (semua kelompok duplikat ambigu, cek manual).");
    return;
  }

  msg += "Lanjutkan hapus " + rowsToDelete.length + " baris duplikat di atas? Baris yang dipertahankan TIDAK akan diubah sama sekali.";

  const resp = ui.alert("Konfirmasi Hapus Duplikat PENJUALAN", msg, ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) {
    ui.alert("Dibatalkan. Tidak ada baris yang dihapus.");
    return;
  }

  // ✅ FIX (performa): SEBELUMNYA di sini deleteRow() dipanggil SATU-SATU per
  // baris duplikat -- tiap deleteRow() adalah 1 panggilan terpisah ke server
  // Google Sheets, jadi kalau duplikatnya ratusan/ribuan baris, totalnya bisa
  // gampang lewat batas waktu eksekusi Apps Script (6 menit) dan macet di
  // tengah jalan (baris di bagian bawah kehapus duluan karena urutan hapusnya
  // dari bawah ke atas, tapi belum sempat nyampe ke baris di atas -> keliatan
  // "gak ada yang berubah" padahal sebagian sebenarnya sudah kehapus).
  //
  // Sekarang caranya diganti jadi CUMA 2 operasi total, seberapa pun banyak
  // baris duplikatnya: 1x tulis ulang semua baris yang DIPERTAHANKAN (dalam
  // urutan asli, gak ada yang digeser2 selain baris duplikatnya sendiri),
  // lalu 1x hapus blok baris sisa di paling bawah yang jadi kosong.
  const deleteSet = {};
  rowsToDelete.forEach(function (r) { deleteSet[r] = true; }); // r = nomor baris 1-based di sheet

  const keptRows = [];
  for (let i = 0; i < data.length; i++) {
    const sheetRow = i + 2;
    if (!deleteSet[sheetRow]) keptRows.push(data[i]);
  }

  if (keptRows.length > 0) {
    shPenjualan.getRange(2, 1, keptRows.length, lastCol).setValues(keptRows);
    shPenjualan.getRange(2, 1, keptRows.length, 1).setNumberFormat("dd/mm/yyyy"); // A: tgl kirim
    shPenjualan.getRange(2, 5, keptRows.length, 1).setNumberFormat("#,##0"); // E: total
  }

  const sisaBarisKosong = data.length - keptRows.length; // = rowsToDelete.length
  if (sisaBarisKosong > 0) {
    shPenjualan.deleteRows(2 + keptRows.length, sisaBarisKosong); // 1x hapus blok kontigu di bawah
  }

  if (typeof reapplyWarnaLunasPenjualan_ === "function") {
    reapplyWarnaLunasPenjualan_(shPenjualan);
  }

  ui.alert(
    "Selesai. " + rowsToDelete.length + " baris duplikat sudah dihapus.\n" +
    (dilewatiAmbigu.length > 0 ? dilewatiAmbigu.length + " No INV dilewati (cek manual): " + dilewatiAmbigu.join(", ") : "Tidak ada yang dilewati.")
  );
}

  // ======================= TempoToko.js (salinan persis) =======================
const TEMPO_COL_TGLKIRIM = 1;
const TEMPO_COL_CUSTOMER = 2;
const TEMPO_COL_NOINV = 3;
const TEMPO_COL_TEMPO = 4;
const TEMPO_COL_PEMBELIAN = 5;

const TEMPO_WARNA_UNLOCKED = "#d9ead3";
const TEMPO_WARNA_LOCKED = "#f4cccc";
const TEMPO_WARNA_GRACE = "#fff2cc"; // kuning buat masa tenggang "-N HARI"

const TEMPO_GRACE_DAYS = 15; // batas hari sebelum LOCKED beneran

const RINGKASAN_COL_CUSTOMER = 7; // G
const RINGKASAN_COL_PEMBELIAN = 8; // H

const TEMPO_PASSWORD = "9346"; // password buat verifikasi sebelum perpanjang tempo

function syncTempoToko() {
  syncTempoToko_(true);
}

function syncTempoToko_(showAlert) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shPenjualan = ss.getSheetByName("PENJUALAN");
  const shTransaksi = ss.getSheetByName("TRANSAKSI");
  const shTempo = ss.getSheetByName("TEMPO TOKO");
  const shHistoryPO = ss.getSheetByName("HISTORY_PO_MASUK");
  const shDiskon = ss.getSheetByName("DISKON");

  if (!shPenjualan || !shTransaksi || !shTempo) {
    if (showAlert) SpreadsheetApp.getUi().alert("Sheet PENJUALAN/TRANSAKSI/TEMPO TOKO tidak ditemukan.");
    return;
  }

  // ⬅️ MIGRASI (Oktober 2026): PENJUALAN & TRANSAKSI dibaca FRESH dari
  // database (sheet lokal cuma salinan yang bisa telat beberapa detik --
  // fungsi ini dipanggil LANGSUNG setelah Simpan Draft / sync PENJUALAN).
  const penjualanMap = {};
  const dataPDB = bacaPenjualanDB_().map(function (r) { return r.slice(0, 3); }); // A..C
  if (dataPDB.length > 0) {
    const dataP = dataPDB;
    let lastCustomer = "";
    dataP.forEach(function (r) {
      const noInv = String(r[2] || "").trim();
      if (!noInv) return;
      const customer = String(r[1] || "").trim() || lastCustomer;
      if (customer) lastCustomer = customer;
      penjualanMap[noInv] = { tglKirim: r[0], customer: customer };
    });
  }

    // noInv -> NO PO (kolom Q di TRANSAKSI), dan noInv -> Keterangan (kolom P)
  const transaksiPOMap = {};
  // ⬅️ BARU: dipakai buat cek teks "DISPLAY" di Keterangan sebagai sinyal
  // tambahan tempo 90 hari (lihat gabungTempoDenganKeterangan_ di Utils.js).
  const transaksiKeteranganMap = {};
  const dataTDB = bacaTransaksiDB_().map(function (r) { return r.slice(2, 17); }); // C..Q (15 kolom)
  if (dataTDB.length > 0) {
    const dataT = dataTDB;
    dataT.forEach(function (r) {
      const noInv = String(r[0] || "").trim();  // C
      const noPO = String(r[14] || "").trim();  // Q (offset 17-3=14)
      const keterangan = r[13];                 // P (offset 16-3=13)
      if (noInv) {
        transaksiPOMap[noInv] = noPO;
        transaksiKeteranganMap[noInv] = keterangan;
      }
    });
  }

  // NO PO -> status (kolom E di HISTORY_PO_MASUK, key di kolom B)
  // ⬅️ FIX: toleransi human error — cek kolom R (METHOD) sebagai fallback
  // kalau kolom E bukan salah satu dari SALE/DISPLAY/3SET/5SET.
  const historyPOMap = {};
  const KATA_KUNCI_STATUS_TEMPO = ["SALE", "DISPLAY", "3SET", "5SET"];
  if (shHistoryPO) {
    const lastRowH = shHistoryPO.getLastRow();
    if (lastRowH >= 2) {
      const dataH = shHistoryPO.getRange(2, 2, lastRowH - 1, 17).getValues(); // B..R
      dataH.forEach(function (r) {
        const noPO = String(r[0] || "").trim();                    // B
        const statusDiskon = String(r[3] || "").trim().toUpperCase();  // E
        const statusMethod = String(r[16] || "").trim().toUpperCase(); // R
        if (!noPO) return;

        if (KATA_KUNCI_STATUS_TEMPO.indexOf(statusDiskon) !== -1) {
          historyPOMap[noPO] = statusDiskon;
        } else if (KATA_KUNCI_STATUS_TEMPO.indexOf(statusMethod) !== -1) {
          historyPOMap[noPO] = statusMethod;
        } else {
          historyPOMap[noPO] = statusDiskon;
        }
      });
    }
  }

  // NAMA TOKO -> jumlah hari tempo, dari DISKON kolom E (format "n HARI")
  const diskonTempoMap = {};
  if (shDiskon) {
    const lastRowD = shDiskon.getLastRow();
    if (lastRowD >= 2) {
      const dataD = shDiskon.getRange(2, 1, lastRowD - 1, 5).getValues(); // A..E
      dataD.forEach(function (r) {
        const namaToko = String(r[0] || "").trim(); // A
        const raw = String(r[4] || "").trim();      // E
        const match = raw.match(/\d+/);
        if (namaToko && match) diskonTempoMap[namaToko] = parseInt(match[0], 10);
      });
    }
  }

  // ⬅️ BARU (migrasi): kolom A-E TEMPO TOKO sekarang di-mirror lewat
  // IMPORTRANGE di sisi inventory (read-only), jadi TIDAK BISA lagi ditulis
  // langsung ke shTempo kayak dulu. Data existing diambil FRESH langsung
  // dari database_SANFELO lewat bacaSheetGateway_ (bukan dari mirror lokal
  // yang bisa telat beberapa detik), dipakai buat: (1) tau invoice mana
  // yang sudah ada biar gak dobel, (2) bawa status (kolom E) & TEMPO (D)
  // lama apa adanya kalau memang gak berubah -- status toh bakal dihitung
  // ulang lagi oleh updateStatusPembelianAll_() yang dipanggil di akhir
  // fungsi ini.
  const existingRows = bacaSheetGateway_("TEMPO TOKO", 2, 1, 5); // A..E
  const existingByInv = {};
  const urutanInv = [];
  existingRows.forEach(function (r) {
    const noInv = String(r[2] || "").trim();
    if (!noInv) return;
    existingByInv[noInv] = r;
    urutanInv.push(noInv);
  });

  const fullRows = [];

  urutanInv.forEach(function (noInv) {
    const p = penjualanMap[noInv];
    const old = existingByInv[noInv];

    if (!p) {
      // Invoice lama yang sudah gak ada lagi di PENJUALAN (edge case) --
      // dipertahankan apa adanya biar gak hilang diam-diam.
      fullRows.push(old);
      return;
    }

    // ⬅️ FIX: dulu ada argumen ke-4 "transaksiMethodMap" yang gak pernah
    // didefinisikan di file ini sama sekali (bukan cuma salah posisi —
    // variabelnya memang gak ada), jadi baris ini SELALU error
    // "transaksiMethodMap is not defined" begitu fungsi ini jalan, dan
    // syncTempoToko_ langsung berhenti di tengah jalan (sheet TEMPO TOKO
    // gak pernah ke-refresh dengan benar). Sudah dihapus, diganti dengan
    // transaksiKeteranganMap (buat cek teks "DISPLAY" di Keterangan).
    const tempoDays = getTempoDays_(noInv, p.customer, transaksiPOMap, transaksiKeteranganMap, historyPOMap, diskonTempoMap);
    const tempoDate = hitungTanggalTempo_(p.tglKirim, tempoDays);

    fullRows.push([
      p.tglKirim,
      p.customer,
      noInv,
      tempoDate || old[3] || "",
      old[4] || "UNLOCKED"
    ]);
  });

  let jumlahBaru = 0;
  Object.keys(penjualanMap).forEach(function (noInv) {
    if (existingByInv[noInv]) return;
    const p = penjualanMap[noInv];
    const tempoDays = getTempoDays_(noInv, p.customer, transaksiPOMap, transaksiKeteranganMap, historyPOMap, diskonTempoMap);
    const tempoDate = hitungTanggalTempo_(p.tglKirim, tempoDays);
    fullRows.push([p.tglKirim, p.customer, noInv, tempoDate || "", "UNLOCKED"]);
    jumlahBaru++;
  });

  kirimAksiGateway_("syncTempoTokoFull", { rows: fullRows });

  updateStatusPembelianAll_();
  paksaRefreshMirrorTempoToko_(shTempo);

  if (showAlert) {
    SpreadsheetApp.getUi().alert("Sinkron TEMPO TOKO selesai: " + jumlahBaru + " toko baru ditambahkan.");
  }
}

/**
 * ⬅️ FIX: split dulu, karena kolom Q di TRANSAKSI bisa berisi gabungan
 * beberapa NO PO dipisah koma (multi-PO per transaksi). Sebelumnya
 * string gabungan ini dipakai langsung sebagai key ke historyPOMap,
 * yang cuma di-key per SATU PO, jadi selalu gagal match.
 *
 * ⬅️ FIX (baru): parameter ke-4 sekarang transaksiKeteranganMap (noInv ->
 * teks Keterangan). Kalau teksnya mengandung "DISPLAY", itu jadi kandidat
 * 90 hari juga — jaring pengaman kalau HISTORY_PO_MASUK belum ditandai
 * DISPLAY dengan benar. Lihat gabungTempoDenganKeterangan_ di Utils.js.
 */
function getTempoDays_(noInv, customer, transaksiPOMap, transaksiKeteranganMap, historyPOMap, diskonTempoMap) {
  const noPORaw = transaksiPOMap[noInv];
  let bestDays = null;

  if (noPORaw) {
    const noPOList = String(noPORaw)
      .split(",")
      .map(function (s) { return s.trim(); })
      .filter(function (s) { return s; });

    noPOList.forEach(function (noPO) {
      const status = historyPOMap[noPO];
      let days = null;
      if (status === "SALE") days = 14;
      else if (status === "DISPLAY") days = 90;
      if (days !== null && (bestDays === null || days > bestDays)) {
        bestDays = days;
      }
    });
  }

  const keteranganText = transaksiKeteranganMap ? transaksiKeteranganMap[noInv] : "";
  bestDays = gabungTempoDenganKeterangan_(bestDays, keteranganText);

  if (bestDays !== null) return bestDays;

  if (customer && diskonTempoMap.hasOwnProperty(customer)) {
    return diskonTempoMap[customer];
  }

  return null;
}

// Tanggal tempo = TGL KIRIM + jumlah hari tempo
function hitungTanggalTempo_(tglKirim, tempoDays) {
  if (!tglKirim || tempoDays === null || tempoDays === undefined || isNaN(tempoDays)) return null;
  const d = new Date(tglKirim);
  d.setDate(d.getDate() + tempoDays);
  return d;
}

// Ambil status pelunasan per NO INV dari sheet PENJUALAN kolom F (LUNAS / CICIL / kosong)
// ⬅️ MIGRASI (Oktober 2026): dibaca FRESH dari PENJUALAN di database;
// parameter shPenjualan dipertahankan supaya pemanggil lama tidak perlu diubah.
function buildPenjualanStatusMap_(shPenjualan) {
  const statusMap = {};
  const dataP = bacaSheetGateway_("PENJUALAN", 2, 3, 4); // C..F
  dataP.forEach(function (r) {
    const noInv = String(r[0] || "").trim();
    if (!noInv) return;
    const status = String(r[3] || "").trim().toUpperCase(); // F
    statusMap[noInv] = status;
  });
  return statusMap;
}

/**
 * Hitung status tempo per invoice.
 * - Belum jatuh tempo / sudah LUNAS -> UNLOCKED
 * - Sudah lewat tempo tapi < TEMPO_GRACE_DAYS hari -> "-N HARI" (grace, kuning)
 * - Sudah lewat tempo >= TEMPO_GRACE_DAYS hari -> LOCKED
 */
function hitungStatusTempo_(tempoVal, pelunasan, todayStr) {
  if (pelunasan === "LUNAS") {
    return { status: "UNLOCKED", warna: TEMPO_WARNA_UNLOCKED, daysOver: 0 };
  }

  if (!tempoVal) {
    return { status: "UNLOCKED", warna: TEMPO_WARNA_UNLOCKED, daysOver: 0 };
  }

  // ⬅️ FIX (Oktober 2026 -- "-NaN HARI" muncul lagi di TEMPO TOKO): dulu
  // tempoVal langsung dibungkus new Date(...). Kalau isinya bukan tanggal
  // yang bisa dibaca (mis. teks "30/10/2026" yang tersimpan sebagai TEKS,
  // bukan tanggal asli), hasilnya Invalid Date -> daysOver = NaN. NaN gagal
  // di SEMUA pengecekan di bawah (<= 0, > GRACE), jadi jatuh ke cabang
  // "masa grace" dan menulis status "-NaN HARI" (kuning) ke database.
  // Sekarang tanggal dibaca lewat parseTempoTanggalAman_ (terima Date, ISO,
  // dd/MM/yyyy, angka serial Sheets); kalau tetap gak bisa dibaca, baris
  // diperlakukan seperti "belum ada tempo" (UNLOCKED) + dicatat di log --
  // JANGAN pernah menulis status NaN.
  const tempoParsed = parseTempoTanggalAman_(tempoVal);
  if (!tempoParsed) {
    Logger.log("hitungStatusTempo_: TEMPO tidak bisa dibaca (" + tempoVal + ") -- dianggap belum ada tempo.");
    return { status: "UNLOCKED", warna: TEMPO_WARNA_UNLOCKED, daysOver: 0 };
  }
  const tempoStr = Utilities.formatDate(tempoParsed, "Asia/Jakarta", "yyyy-MM-dd");

  const msPerDay = 24 * 60 * 60 * 1000;
  const todayDate = new Date(todayStr + "T00:00:00");
  const tempoDate = new Date(tempoStr + "T00:00:00");
  const daysOver = Math.round((todayDate - tempoDate) / msPerDay);
  if (isNaN(daysOver)) {
    return { status: "UNLOCKED", warna: TEMPO_WARNA_UNLOCKED, daysOver: 0 };
  }

  // Belum lewat tempo
  if (daysOver <= 0) {
    return {
      status: "UNLOCKED",
      warna: TEMPO_WARNA_UNLOCKED,
      daysOver: 0
    };
  }

  // Lewat masa grace -> LOCKED
  if (daysOver > TEMPO_GRACE_DAYS) {
    return {
      status: "LOCKED",
      warna: TEMPO_WARNA_LOCKED,
      daysOver: daysOver
    };
  }

  // Masih masa grace
  const sisaHari = TEMPO_GRACE_DAYS - daysOver;

  return {
    status: "-" + sisaHari + " HARI",
    warna: TEMPO_WARNA_GRACE,
    daysOver: daysOver
  };
}

/**
 * Baca nilai TEMPO jadi Date yang valid, atau null kalau memang gak bisa.
 * Menerima: Date, string ISO (yyyy-MM-dd... dari gateway), teks dd/MM/yyyy
 * (format Indonesia), dan angka serial tanggal Google Sheets.
 */
function parseTempoTanggalAman_(v) {
  if (v === null || v === undefined || v === "") return null;
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v;

  if (typeof v === "number") {
    if (v > 20000 && v < 80000) { // angka serial Sheets (hari sejak 30/12/1899)
      const d = new Date(Date.UTC(1899, 11, 30) + Math.round(v) * 86400000);
      return isNaN(d.getTime()) ? null : d;
    }
    return null;
  }

  const t = String(v).trim();
  if (!t) return null;

  // dd/MM/yyyy atau dd-MM-yyyy (opsional diikuti jam)
  const m = t.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})(?:\s.*)?$/);
  if (m) {
    const d = new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]), 12, 0, 0);
    return (isNaN(d.getTime()) || d.getMonth() !== Number(m[2]) - 1) ? null : d;
  }

  const d2 = new Date(t); // ISO "2026-10-30" / "2026-10-30T..."
  return isNaN(d2.getTime()) ? null : d2;
}

function updateStatusPembelianAll_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shPenjualan = ss.getSheetByName("PENJUALAN");
  const sh = ss.getSheetByName("TEMPO TOKO");
  if (!sh || !shPenjualan) return;

  // ⬅️ BARU (migrasi): kolom A-E (termasuk E / STATUS PEMBELIAN) sekarang
  // di-mirror lewat IMPORTRANGE, jadi gak bisa ditulis langsung lokal lagi
  // (makanya blok clearDataValidations yang dulu ada di sini juga sudah
  // gak relevan -- gak ada lagi dropdown lokal yang perlu dicabut, karena
  // kolom E emang bukan sel yang ditulis IMPORTRANGE gak punya validasi).
  // Data A:E diambil FRESH dari database (bukan dari mirror lokal yang
  // bisa telat beberapa detik), STATUS (E) dihitung ulang di sini
  // berdasarkan TEMPO (D) yang sudah ada + status LUNAS/belum di
  // PENJUALAN, lalu SEMUA baris (bukan cuma kolom E) dikirim balik lewat
  // aksi "syncTempoTokoFull" yang sama dipakai syncTempoToko_ -- gateway
  // itu cuma tau "replace semua", jadi kolom A-D yang gak berubah pun
  // harus ikut dikirim apa adanya biar gak hilang.
  const dataTempo = bacaSheetGateway_("TEMPO TOKO", 2, 1, 5); // A..E
  if (dataTempo.length === 0) return;

  const statusMap = buildPenjualanStatusMap_(shPenjualan);
  const todayStr = Utilities.formatDate(new Date(), "Asia/Jakarta", "yyyy-MM-dd");

  const customerOrder = [];
  // per customer sekarang nyimpen locked flag + info grace (hari tersisa)
  const customerInfo = {};
  const fullRows = [];
  // ⬅️ FIX (Oktober 2026 -- warna A:E gak cocok sama STATUS, mis. SIRO
  // isinya UNLOCKED tapi barisnya merah/kuning): warna A:E sekarang
  // TIDAK lagi ditulis per baris lewat setBackgrounds(). Isi A:E datang
  // dari IMPORTRANGE yang refresh-nya telat/asinkron, jadi warna statis
  // per nomor baris gampang ketinggalan (baris yang dulu LOCKED lalu
  // LUNAS tetap merah). Ditambah lagi, conditional formatting lama di
  // sheet ini salah kolom ($L2="UNLOCKED", harusnya $E2) -- jadi baris
  // UNLOCKED gak pernah ke-hijau-in dan warna basi itu yang kelihatan.
  // Sekarang warna ditentukan conditional formatting yang membaca
  // langsung kolom E (lihat pasangFormatWarnaTempoToko_), jadi selalu
  // ikut isi terbaru tanpa nunggu script jalan.
  pasangFormatWarnaTempoToko_(sh);

  // ⬅️ FIX (SIRO LOCKED padahal semua invoice-nya LUNAS): invoice yang
  // sama sekali GAK ADA di PENJUALAN (mis. SPRIN-2026-08-001, aksesoris
  // Rp0 / tidak ditagihkan) dulu dianggap "belum lunas" selamanya, jadi
  // begitu lewat tempo + grace, customer-nya ke-LOCKED permanen dan
  // Surat Jalan ke-blokir. Gak ada di PENJUALAN = gak ada piutang yang
  // ditagih, jadi diperlakukan kayak LUNAS. Pengaman: kalau PENJUALAN
  // kebaca kosong (gagal baca), aturan ini gak dipakai.
  const penjualanTerbaca = Object.keys(statusMap).length > 0;

  for (let i = 0; i < dataTempo.length; i++) {
    const r = dataTempo[i];
    const customer = String(r[TEMPO_COL_CUSTOMER - 1] || "").trim();
    const noInv = String(r[TEMPO_COL_NOINV - 1] || "").trim();
    const tempoVal = r[TEMPO_COL_TEMPO - 1];

    const adaDiPenjualan = statusMap.hasOwnProperty(noInv);
    const pelunasan = adaDiPenjualan ? statusMap[noInv] : (penjualanTerbaca ? "LUNAS" : "");
    const hasil = hitungStatusTempo_(tempoVal, pelunasan, todayStr);

    fullRows.push([r[0], r[1], r[2], r[3], hasil.status]);

    if (customer) {
      if (!(customer in customerInfo)) {
        customerOrder.push(customer);
        customerInfo[customer] = { locked: false, hasGrace: false, minSisaHari: TEMPO_GRACE_DAYS };
      }
      if (hasil.status === "LOCKED") {
        customerInfo[customer].locked = true;
      } else if (hasil.daysOver > 0) {
        customerInfo[customer].hasGrace = true;

        const sisaHari = TEMPO_GRACE_DAYS - hasil.daysOver;

        if (sisaHari < customerInfo[customer].minSisaHari) {
          customerInfo[customer].minSisaHari = sisaHari;
        }
      }
    }
  }

  kirimAksiGateway_("syncTempoTokoFull", { rows: fullRows });

  // Ringkasan G:H tetap murni lokal (lihat writeRingkasanToko_) -- tidak
  // pernah dikirim ke database, jadi tidak ada round-trip tambahan di sini.
  writeRingkasanToko_(sh, customerOrder, customerInfo);
}

/**
 * Paksa IMPORTRANGE kolom A-E di TEMPO TOKO (mirror lokal) buat langsung
 * re-fetch dari database_SANFELO, bukan nunggu jadwal recalculation
 * background Google Sheets (yang gak otomatis ke-trigger walau datanya
 * berubah di spreadsheet lain). Dipanggil di titik-titik aksi admin yang
 * eksplisit (sync manual, perpanjang tempo) -- BUKAN di
 * updateStatusPembelianAll_() sendiri, karena fungsi itu jalan sering
 * banget lewat trigger onEdit di sheet lain dan gak perlu maksa refresh
 * tiap kali. Aman dipanggil walau formulanya belum/sudah gak ada di A2
 * (dicek dulu isi formula-nya).
 */
function paksaRefreshMirrorTempoToko_(sh) {
  try {
    const cell = sh.getRange(2, TEMPO_COL_TGLKIRIM); // A2
    const formula = cell.getFormula();
    if (formula && formula.toUpperCase().indexOf("IMPORTRANGE") !== -1) {
      cell.setFormula(formula);
    }
  } catch (err) {
    // non-blocking -- kalau gagal refresh paksa, biarin aja, nanti
    // ke-refresh sendiri lewat jadwal background Google Sheets.
  }
}

/**
 * Pastikan conditional formatting A:E TEMPO TOKO benar: warna baris
 * ngikutin kolom E (STATUS PEMBELIAN) -- UNLOCKED hijau, LOCKED merah,
 * "-N HARI" kuning. Rule lama yang nempel di A:E (termasuk yang salah
 * kolom $L2) dibuang dan diganti; rule di kolom lain dibiarkan.
 * Cuma nulis ulang kalau rule-nya memang belum persis begini, jadi aman
 * dipanggil tiap kali updateStatusPembelianAll_() jalan.
 */
function pasangFormatWarnaTempoToko_(sh) {
  try {
    const target = sh.getRange(2, TEMPO_COL_TGLKIRIM, sh.getMaxRows() - 1, 5); // A2:E
    const rumus = [
      { formula: '=$E2="UNLOCKED"', warna: TEMPO_WARNA_UNLOCKED },
      { formula: '=$E2="LOCKED"', warna: TEMPO_WARNA_LOCKED },
      { formula: '=ISNUMBER(SEARCH("HARI",$E2))', warna: TEMPO_WARNA_GRACE }
    ];

    const semua = sh.getConditionalFormatRules();
    const kenaAE = function (rule) {
      return rule.getRanges().some(function (rg) {
        return rg.getColumn() <= TEMPO_COL_PEMBELIAN && rg.getLastColumn() >= TEMPO_COL_TGLKIRIM;
      });
    };
    const normal = function (f) { return String(f || "").replace(/\s/g, "").toUpperCase(); };

    const lamaAE = semua.filter(kenaAE);
    const sudahBenar = lamaAE.length === rumus.length && lamaAE.every(function (rule, i) {
      const cond = rule.getBooleanCondition();
      const vals = cond ? cond.getCriteriaValues() : [];
      return normal(vals[0]) === normal(rumus[i].formula) &&
        String(cond.getBackground() || "").toLowerCase() === rumus[i].warna &&
        rule.getRanges()[0].getRow() === 2 &&
        rule.getRanges()[0].getLastRow() >= sh.getMaxRows();
    });
    if (sudahBenar) return;

    const baru = rumus.map(function (x) {
      return SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(x.formula)
        .setBackground(x.warna)
        .setRanges([target])
        .build();
    });
    sh.setConditionalFormatRules(semua.filter(function (rule) { return !kenaAE(rule); }).concat(baru));

    // Buang warna statis lama di A:E (sisa setBackgrounds versi dulu) biar
    // gak ada warna basi yang ketahuan kalau STATUS-nya kosong.
    target.setBackground(null);
  } catch (err) {
    // non-blocking -- urusan tampilan doang, gak pengaruh ke data.
    Logger.log("pasangFormatWarnaTempoToko_ error: " + err);
  }
}

// Tentukan status final customer, prioritas LOCKED > grace (paling parah) > UNLOCKED
function statusFinalCustomer_(info) {
  if (info.locked) return "LOCKED";
  if (info.hasGrace) return "-" + info.minSisaHari + " HARI";
  return "UNLOCKED";
}

function warnaUntukStatus_(status) {
  if (status === "LOCKED") return TEMPO_WARNA_LOCKED;
  if (status === "UNLOCKED") return TEMPO_WARNA_UNLOCKED;
  return TEMPO_WARNA_GRACE; // "-N HARI"
}

// Tulis tabel ringkasan per toko ke kolom G:H
function writeRingkasanToko_(sh, customerOrder, customerInfo) {
  const lastRow = sh.getLastRow();

  // bersihkan area G:H lama dulu
  if (lastRow >= 2) {
    sh.getRange(2, RINGKASAN_COL_CUSTOMER, lastRow - 1, 2).clearContent().clearDataValidations().setBackground(null);
  }

  if (customerOrder.length === 0) return;

  const rows = customerOrder.map(function (customer) {
    return [customer, statusFinalCustomer_(customerInfo[customer])];
  });

  sh.getRange(2, RINGKASAN_COL_CUSTOMER, rows.length, 2).setValues(rows);

  // ⬅️ FIX (performa): sama kayak updateStatusPembelianAll_ -- dulu
  // setDataValidation() & setBackground() dipanggil SATU-SATU per baris
  // customer di dalam loop. Sekarang dikumpulin ke array dulu, ditulis
  // SEKALI lewat setDataValidations()/setBackgrounds() (2D array,
  // sebaris-sebaris ngikutin bentuk range-nya) -- hasil akhirnya sama
  // persis, cuma jauh lebih sedikit panggilan API.
  const rules = [];
  const warnaGrid = [];

  for (let i = 0; i < rows.length; i++) {
    const statusAktif = rows[i][1];
    const warna = warnaUntukStatus_(statusAktif);

    // dropdown per baris isinya status aktif customer itu + UNLOCKED + LOCKED
    // (dedup biar gak dobel kalau statusnya emang udah UNLOCKED/LOCKED)
    const opsi = [];
    [statusAktif, "UNLOCKED", "LOCKED"].forEach(function (v) {
      if (opsi.indexOf(v) === -1) opsi.push(v);
    });
    const rule = SpreadsheetApp.newDataValidation().requireValueInList(opsi, true).build();

    rules.push([rule]);
    warnaGrid.push([warna, warna]);
  }

  sh.getRange(2, RINGKASAN_COL_PEMBELIAN, rules.length, 1).setDataValidations(rules);
  sh.getRange(2, RINGKASAN_COL_CUSTOMER, warnaGrid.length, 2).setBackgrounds(warnaGrid);
}

/* ======================================================
   ⬅️ CATATAN MERGE:
   Dulu di sini ada function onEdit(e) yang jadi dispatcher
   (route ke onEditTempoToko_, onEditTransaksiStatus_,
   onEditDetailStatus_). Sekarang SATU-SATUNYA function onEdit
   di seluruh project ada di PO_HISTORY.gs, dan dia yang
   manggil ketiga handler di bawah ini. Jangan bikin function
   onEdit(e) lagi di file manapun selain PO_HISTORY.gs, karena
   Apps Script cuma boleh punya satu simple trigger onEdit
   per project (kalau ada dua, yang satu bakal ketimpa/rebutan).
   ====================================================== */

/**
 * Handler khusus TEMPO TOKO (kolom H / ringkasan).
 * Dipanggil dari onEdit(e) di PO_HISTORY.gs saat sheetName === "TEMPO TOKO".
 */
function onEditTempoToko_(e, sh) {
  if (e.range.getColumn() !== RINGKASAN_COL_PEMBELIAN) return;
  if (e.range.getRow() < 2) return;
  if (e.range.getNumRows() !== 1 || e.range.getNumColumns() !== 1) return;

  const newVal = String(e.range.getValue() || "").trim().toUpperCase();
  if (newVal !== "UNLOCKED") return;

  const customer = String(sh.getRange(e.range.getRow(), RINGKASAN_COL_CUSTOMER).getValue() || "").trim();
  if (!customer) return;

  // simpan nilai SEBELUM diubah (mis. "LOCKED" / "-5 HARI"), dipakai buat
  // restore persis kalau password salah / dibatalkan, gak cuma ngandelin
  // recompute ulang.
  const oldVal = e.oldValue !== undefined ? String(e.oldValue).trim() : "";

  handlePerpanjangTempo_(sh, customer, e.range, oldVal);
}

/**
 * Handler khusus TRANSAKSI (kolom N / status DRAFT-PROGRESS-REVISI-DONE).
 * Kalau user manual pilih DONE tapi PENJUALAN utk NO INV itu belum LUNAS,
 * cell direvert ke nilai sebelumnya + alert.
 * Dipanggil dari onEdit(e) di PO_HISTORY.gs saat sheetName === "TRANSAKSI".
 */
function onEditTransaksiStatus_(e, sh) {
  if (e.range.getColumn() !== TRANSAKSI_COL_STATUS) return;
  if (e.range.getRow() < 2) return;
  if (e.range.getNumRows() !== 1 || e.range.getNumColumns() !== 1) return;

  const rowIdx = e.range.getRow();
  const newVal = String(e.range.getValue() || "").trim().toUpperCase();
  const oldVal = e.oldValue !== undefined ? String(e.oldValue).trim() : "";
  const noInv = String(sh.getRange(rowIdx, TRANSAKSI_COL_NOINV_STATUS).getValue() || "").trim();

  if (newVal === "DONE") {
    const statusPenjualan = getStatusPenjualanByNoInv_(noInv);
    if (statusPenjualan !== "LUNAS") {
      e.range.setValue(oldVal);
      bersihkanWarnaBarisTransaksi_(sh, rowIdx);
      SpreadsheetApp.getUi().alert(
        "Invoice " + noInv + " belum LUNAS di PENJUALAN. Status tidak bisa diubah jadi DONE secara manual."
      );
      return;
    }
    // Penjualan memang sudah LUNAS -> boleh, kasih warna hijau
    sh.getRange(rowIdx, 1, 1, TRANSAKSI_COL_STATUS).setBackground(TRANSAKSI_WARNA_DONE);
  } else {
    // Status selain DONE -> gak usah diwarnain
    bersihkanWarnaBarisTransaksi_(sh, rowIdx);
  }
}

/**
 * Handler khusus DETAIL_TRANSAKSI (kolom T / status PO-DONE).
 * Sama logic-nya kayak TRANSAKSI, tapi NO INV-nya ada di kolom B.
 * Dipanggil dari onEdit(e) di PO_HISTORY.gs saat sheetName === "DETAIL_TRANSAKSI".
 */
function onEditDetailStatus_(e, sh) {
  if (e.range.getColumn() !== DETAIL_COL_STATUS) return;
  if (e.range.getRow() < 2) return;
  if (e.range.getNumRows() !== 1 || e.range.getNumColumns() !== 1) return;

  const rowIdx = e.range.getRow();
  const newVal = String(e.range.getValue() || "").trim().toUpperCase();
  const oldVal = e.oldValue !== undefined ? String(e.oldValue).trim() : "";
  const noInv = String(sh.getRange(rowIdx, DETAIL_COL_NOINV_STATUS).getValue() || "").trim();

  if (newVal === "DONE") {
    const statusPenjualan = getStatusPenjualanByNoInv_(noInv);
    if (statusPenjualan !== "LUNAS") {
      e.range.setValue(oldVal);
      bersihkanWarnaBarisDetail_(sh, rowIdx);
      SpreadsheetApp.getUi().alert(
        "Invoice " + noInv + " belum LUNAS di PENJUALAN. Status tidak bisa diubah jadi DONE secara manual."
      );
      return;
    }
    sh.getRange(rowIdx, 1, 1, DETAIL_COL_STATUS).setBackground(DETAIL_WARNA_DONE);
  } else {
    bersihkanWarnaBarisDetail_(sh, rowIdx);
  }
}

/**
 * Popup verifikasi password dulu, baru minta jumlah hari perpanjangan,
 * lalu update kolom D (TEMPO) di SEMUA baris invoice LOCKED milik customer
 * tsb, jadi tempo lama + input hari. Setelah itu status di-refresh ulang
 * (jadi cell dropdown balik sesuai hasil hitung, bukan mentah-mentah
 * kesimpen "UNLOCKED").
 */
function handlePerpanjangTempo_(sh, customer, range, oldVal) {
  const ui = SpreadsheetApp.getUi();

  // 1) Verifikasi password dulu
  const respPass = ui.prompt(
    "Verifikasi - " + customer,
    "Masukkan password untuk melanjutkan perpanjangan tempo:",
    ui.ButtonSet.OK_CANCEL
  );

  if (respPass.getSelectedButton() !== ui.Button.OK) {
    batalkanEditStatus_(range, oldVal); // batal, balikin cell PERSIS ke status sebelumnya
    return;
  }

  const passInput = respPass.getResponseText().trim();
  if (passInput !== TEMPO_PASSWORD) {
    ui.alert("Password salah. Perpanjangan dibatalkan.");
    batalkanEditStatus_(range, oldVal);
    return;
  }

  // 2) Password benar, lanjut minta jumlah hari perpanjangan
  const resp = ui.prompt(
    "Perpanjang Tempo - " + customer,
    "Masukkan jumlah hari perpanjangan tempo (dihitung dari tanggal tempo lama tiap invoice):",
    ui.ButtonSet.OK_CANCEL
  );

  if (resp.getSelectedButton() !== ui.Button.OK) {
    batalkanEditStatus_(range, oldVal);
    return;
  }

  const jumlahHari = parseInt(resp.getResponseText().trim(), 10);
  if (isNaN(jumlahHari) || jumlahHari <= 0) {
    ui.alert("Input tidak valid. Harus angka lebih dari 0. Perpanjangan dibatalkan.");
    batalkanEditStatus_(range, oldVal);
    return;
  }

  // ⬅️ BARU (migrasi): kolom D (TEMPO) sekarang bagian dari mirror A:E,
  // jadi gak bisa ditulis langsung lokal lagi -- dikirim sebagai aksi baru
  // ke gateway ("perpanjangTempoPerCustomer"), database yang nyari semua
  // baris LOCKED milik customer ini dan nambah TEMPO-nya di sana.
  let hasil;
  try {
    hasil = kirimAksiGateway_("perpanjangTempoPerCustomer", {
      customer: customer,
      jumlahHari: jumlahHari
    });
  } catch (err) {
    ui.alert("Gagal menyimpan perpanjangan tempo ke database_SANFELO: " + err.message + "\nCoba lagi sebentar.");
    batalkanEditStatus_(range, oldVal);
    return;
  }

  updateStatusPembelianAll_();
  paksaRefreshMirrorTempoToko_(sh);

  ui.alert(
    "Tempo diperpanjang " + jumlahHari + " hari untuk " + hasil.updatedCount +
    " invoice LOCKED milik " + customer + "."
  );
}

// paksa cell H balik ke nilai sebelumnya + warna yang sesuai,
// dipanggil di semua jalur batal (password salah/cancel, cancel jumlah hari, dst).
function batalkanEditStatus_(range, oldVal) {
  if (oldVal) {
    range.setValue(oldVal);
    range.setBackground(warnaUntukStatus_(oldVal));
  }
  // tetap recompute penuh biar kolom E, warna baris lain, dan customer lain konsisten
  updateStatusPembelianAll_();
}

function installTempoDailyTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function (t) {
    if (
      t.getHandlerFunction() === "updateStatusPembelianAll_" ||
      t.getHandlerFunction() === "syncTempoToko_"
    ) {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger("updateStatusPembelianAll_")
    .timeBased()
    .atHour(0)
    .everyDays(1)
    .create();
  SpreadsheetApp.getUi().alert("Trigger harian terpasang, jalan tiap jam 00:00.");
}

  // ======================= Transaction.js (salinan persis) =======================
/**
 * ====== KONSTANTA KOLOM SHEET "TRANSAKSI" ======
 * D = Tanggal kirim
 * G = Tempo (tanggal jatuh tempo)
 */
const TRANSAKSI_COL_TGL = 4;   // D
const TRANSAKSI_COL_TEMPO = 7; // G

// ⬅️ MIGRASI (Oktober 2026): parameter pertama sekarang DAFTAR NOMOR (array
// 1 kolom hasil baca FRESH dari database), bukan sheet TRANSAKSI lokal --
// TRANSAKSI lokal cuma salinan yang bisa telat beberapa detik setelah
// Simpan Draft, dan kalau dibaca dari situ dua transaksi beruntun bisa
// dapat nomor SJ/INV yang sama.
function getNextNoTransaksiDariHistori_(listNomor, system, prefixNo, periodeSekarang, cellCounter, cellPeriode) {
  let maxNum = 0;
  const prefixCari = prefixNo + periodeSekarang + "-";

  (listNomor || []).forEach(function(r) {
    const val = String(r[0] || "").trim();
    if (val.indexOf(prefixCari) === 0) {
      const num = parseInt(val.substring(prefixCari.length), 10);
      if (!isNaN(num) && num > maxNum) maxNum = num;
    }
  });

  const nextNum = maxNum + 1;

  // SYSTEM cuma dijadiin cache/referensi, bukan sumber utama lagi
  system.getRange(cellCounter).setValue(nextNum);
  system.getRange(cellPeriode).setValue(periodeSekarang);

  return nextNum;
}

function transaksiBaru() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sj = ss.getSheetByName("SURAT JALAN");
  const inv = ss.getSheetByName("INVOICE");
  const system = ss.getSheetByName("SYSTEM");

  clearJumlahKertasManual_();

  pastikanKapasitasItem_(sj, SJ_ITEM_START_ROW, SJ_ITEM_MAX_ROWS, SJ_ITEM_MAX_ROWS, SJ_FOOTER_COL_START, SJ_FOOTER_COL_END, SJ_FOOTER_TEKS);
  if (inv) pastikanKapasitasItem_(inv, INV_DISPLAY_START_ROW, INV_ITEM_MAX_ROWS, INV_ITEM_MAX_ROWS, INV_FOOTER_COL_START, INV_FOOTER_COL_END, INV_FOOTER_TEKS);

  const today = new Date();
  const tahun = today.getFullYear();
  const bulan = Utilities.formatString("%02d", today.getMonth()+1);
  const periodeSekarang = tahun + "-" + bulan;

  // ⬅️ GANTI: hitung ulang langsung dari TRANSAKSI tiap kali, bukan increment tersimpan
  // ⬅️ MIGRASI: No SJ (B) & No INV (C) dibaca FRESH dari TRANSAKSI di database
  const nomorDB = bacaSheetGateway_("TRANSAKSI", 2, 2, 2); // B..C
  const lastSJ  = getNextNoTransaksiDariHistori_(nomorDB.map(function (r) { return [r[0]]; }), system, "SJ-",  periodeSekarang, "B1", "C1"); // kolom B = No SJ
  const lastINV = getNextNoTransaksiDariHistori_(nomorDB.map(function (r) { return [r[1]]; }), system, "INV-", periodeSekarang, "B2", "C2"); // kolom C = No INV

  const noSJ = `SJ-${tahun}-${bulan}-${Utilities.formatString("%03d",lastSJ)}`;
  const noINV = `INV-${tahun}-${bulan}-${Utilities.formatString("%03d",lastINV)}`;

  sj.getRange("H6").setValue(noSJ);
  inv.getRange("H6").setValue(noINV);

  setTanggalSaja_(sj.getRange("H5"), today);   // cuma tanggal, tanpa jam
  setTanggalSaja_(inv.getRange("H7"), today);

  sj.getRange("A5").clearContent();
  sj.getRange("A6").clearContent();

  inv.getRange("A9").clearContent();

  sj.getRange(SJ_ITEM_START_ROW, 2, SJ_ITEM_MAX_ROWS, 1).clearContent();
  sj.getRange(SJ_ITEM_START_ROW, 5, SJ_ITEM_MAX_ROWS, 1).clearContent();
  sj.getRange(SJ_ITEM_START_ROW, 6, SJ_ITEM_MAX_ROWS, 1).clearContent();
  sj.getRange(SJ_ITEM_START_ROW, 7, SJ_ITEM_MAX_ROWS, 1).clearContent();

  inv.getRange(INV_DISPLAY_START_ROW, 4, INV_ITEM_MAX_ROWS, 1).clearContent();
}

/* ------------------------------------------------------
   HELPER BARU — hitung jumlah baris item aktual
   berdasarkan posisi footer (anchor teks), bukan angka mati.
   ------------------------------------------------------ */
function hitungJumlahBarisItem_(sheet, startRow, footerTeks) {
  const footerRow = cariBarisAnchor_(sheet, footerTeks, startRow);
  if (footerRow === -1) {
    throw new Error('Anchor footer "' + footerTeks + '" tidak ditemukan di sheet "' + sheet.getName() + '".');
  }
  return footerRow - startRow; // jumlah baris item sebelum footer
}


/**
 * ====== KONSTANTA KOLOM SHEET "DETAIL_TRANSAKSI" (BARU) ======
 * Ditambahkan supaya nomor kolom LABA & SOURCE_PO gak "angka mati"
 * yang tersebar di banyak tempat dalam kode.
 *  O = HARGA NET   (15)
 *  R = HPP BARANG  (18)
 *  S = LABA        (19)  -> SEKARANG RUMUS, bukan angka statis
 *  T = STATUS ORDER(20)  -> tidak disentuh script ini (manual/dropdown)
 *  U = SOURCE_PO   (21)  -> BARU, dipindah dari kolom S supaya tidak
 *                           bentrok dengan formula LABA
 */
const DET_COL_HARGA_NET = 15; // O
const DET_COL_HPP_BARANG = 18; // R
const DET_COL_LABA = 19; // S
const DET_COL_SOURCE_PO = 21; // U

/**
 * Bangun rumus LABA untuk satu baris: HARGA NET (O) - HPP BARANG (R).
 * Dibungkus IFERROR supaya tidak nongol #VALUE!/#REF! kalau salah satu
 * sel kosong/berubah jadi teks.
 */
function buatRumusLaba_(row) {
  return "=IFERROR(O" + row + "-R" + row + ";\"\")";
}

function simpanDraft() {
  const ss = SpreadsheetApp.getActive();

  const sj = ss.getSheetByName("SURAT JALAN");
  const inv = ss.getSheetByName("INVOICE");
  // ⬅️ MIGRASI (Oktober 2026): TRANSAKSI & DETAIL_TRANSAKSI sekarang ditulis
  // ke database lewat gateway (aksi "simpanTransaksiDraft"), bukan ke sheet
  // lokal -- sheet lokal cuma salinan IMPORTRANGE.

  const now = getTimestamp();

  const noSJ = sj.getRange("H6").getValue();
  const noINV = inv.getRange("H6").getValue();
  const tgl = hanyaTanggal_(sj.getRange("H5").getValue()); // buang jam kalau ada
  const toko = sj.getRange("A5").getValue();
  const daerah = sj.getRange("A6").getValue();
  const method = inv.getRange("I3").getValue();
  const noPol = sj.getRange("H7").getValue();
  let noPOGabungan = sj.getRange("H3").getValue(); // No PO (bisa multi, dipisah koma)

  // ⬅️ FIX: kalau H3 kosong, dulu noPOGabungan langsung dianggap "" apa
  // adanya. Ini kejadian tiap kali draft dibuka ulang lewat "📂 Buka
  // Draft" (SJ_DRAFT_FETCH.js SENGAJA ngosongin H3 di mode itu, karena
  // provenance PO-nya disimpan per-item di kolom I, bukan di H3). Kalau
  // abis buka draft admin langsung klik "💾 Simpan Draft" lagi (mis. cuma
  // mau benerin qty/harga) TANPA isi ulang H3 manual, noPOGabungan kosong
  // ini dulu KETULIS APA ADANYA ke TRANSAKSI!Q -- ngilangin jejak No PO
  // transaksi itu buat selamanya -- DAN bikin hitung tempo di bawah gak
  // ketemu status SALE/DISPLAY PO-nya, jatuh ke default DISKON toko
  // (itu salah satu penyebab Jatuh Tempo di INVOICE keliatan "acak"
  // antara nilai asli vs default toko tiap kali draft dibuka-tutup ulang).
  // Sekarang: kalau H3 kosong, coba rekonstruksi dari mapping per-item
  // di kolom I (sumber yang sama dipakai "Buka Draft" buat isi ulang
  // kolom I itu sendiri), baru kalau itu juga kosong beneran fallback ke "".
  if (!String(noPOGabungan || "").trim()) {
    const rekonstruksi = ambilNoPOGabunganDariKolomMappingSJ_(sj);
    if (rekonstruksi) noPOGabungan = rekonstruksi;
  }

  // ⬅️ FIX: kalau H5 (Tgl) kosong/rusak, dulu tetap lanjut nulis Date
  // yang invalid ke TRANSAKSI/DETAIL_TRANSAKSI/PENJUALAN — hasilnya sel
  // itu jadi error #VALUE! diam-diam (kejadian nyata di SJ-2026-09-042,
  // transaksi multi-PO buat MEBEL UTOMO). Sekarang divalidasi dulu di
  // sini: kalau tanggalnya nggak valid, STOP dengan pesan jelas, jangan
  // biarkan data rusak ikut tersimpan.
  if (!(tgl instanceof Date) || isNaN(tgl.getTime())) {
    throw new Error('Tanggal (H5) di sheet SURAT JALAN kosong atau tidak valid. Isi ulang tanggalnya dulu sebelum Simpan Draft.');
  }

  // ⬅️ FIX: ketRow/keterangan dipindah ke ATAS, SEBELUM hitung tempo,
  // karena teks Keterangan (mis. "DISPLAY") sekarang jadi salah satu
  // sinyal penentu tempo juga (lihat catatan di bawah). ketRow cuma
  // butuh sheet INVOICE + INV_DISPLAY_START_ROW, jadi aman dipindah
  // lebih awal (tidak butuh footerRowInv).
  const ketRow = cariBarisAnchor_(inv, INV_KETERANGAN_ANCHOR_TEKS, INV_DISPLAY_START_ROW);
  if (ketRow === -1) {
    throw new Error('Anchor "' + INV_KETERANGAN_ANCHOR_TEKS + '" tidak ditemukan di sheet "INVOICE".');
  }
  const keterangan = inv.getRange(ketRow, 3).getValue();

  // ⬅️ BARU: tempo dihitung OTOMATIS (bukan lagi manual dari INVOICE!H8).
  // Urutan penentuan jumlah hari tempo:
  // 1) Semua NO PO di noPOGabungan (dipisah koma) dicek ke HISTORY_PO_MASUK
  //    kolom E: SALE -> 14 hari, DISPLAY -> 90 hari. Kalau ada beberapa PO
  //    dengan status beda, dipakai yang PALING PANJANG (paling aman buat toko).
  // 1b) ⬅️ FIX: kalau teks Keterangan invoice (kolom C di baris anchor
  //     "Keterangan:") mengandung kata "DISPLAY", itu JUGA dipakai sebagai
  //     kandidat 90 hari — supaya kasus admin cuma nulis "DISPLAY" di
  //     Keterangan (tanpa sempat nandain status di HISTORY_PO_MASUK) tetap
  //     kehitung 90 hari, bukan jatuh ke default toko di DISKON.
  // 2) Kalau semua di atas gak ketemu -> fallback ke DISKON kolom E
  //    (format "n HARI") berdasarkan NAMA TOKO di kolom A.
  // Hasil akhir (tanggal) DITULIS ULANG ke INVOICE!H8, nimpa apapun isi
  // manual sebelumnya — H8 sekarang murni output, bukan input lagi.
  const tempoHari = hitungTempoHariUntukTransaksi_(noPOGabungan, toko, keterangan);
  const tempo = hitungTanggalTempoTransaksi_(tgl, tempoHari);

  if (tempoHari !== null && !isNaN(tempoHari)) {
    inv.getRange("H8").setFormula("=H7+" + tempoHari).setNumberFormat("dd/mm/yyyy");
  } else {
    inv.getRange("H8").setValue("");
  }

  if (!noSJ) throw new Error("No SJ kosong.");
  if (!noINV) throw new Error("No Invoice kosong.");
  if (!toko) throw new Error("Nama toko kosong.");

  // ⬅️ MIGRASI: mapping lama dibaca FRESH dari DETAIL_TRANSAKSI di database
  const oldMapping = ambilMappingLamaDariDetailTransaksi_(bacaDetailDB_(), noSJ);

  const footerRowSJ  = cariBarisAnchor_(sj, SJ_FOOTER_TEKS, SJ_ITEM_START_ROW);
  const footerRowInv = cariBarisAnchor_(inv, INV_FOOTER_TEKS, INV_DISPLAY_START_ROW);

  if (footerRowSJ === -1) throw new Error('Anchor footer "' + SJ_FOOTER_TEKS + '" tidak ditemukan di SURAT JALAN.');
  if (footerRowInv === -1) throw new Error('Anchor footer "' + INV_FOOTER_TEKS + '" tidak ditemukan di INVOICE.');

  const sjNumRows  = footerRowSJ - SJ_ITEM_START_ROW;
  const invNumRows = footerRowInv - INV_DISPLAY_START_ROW;

  if (sjNumRows !== invNumRows) {
    throw new Error(
      "Jumlah baris item SJ (" + sjNumRows + ") dan INVOICE (" + invNumRows + ") tidak sinkron. " +
      "Cek kapasitas kedua sheet sebelum simpan draft."
    );
  }

  const numRows = sjNumRows;

  const sales = inv.getRange(footerRowInv + 1, 3).getValue();
  const totalBayar = inv.getRange(footerRowInv + 3, 8).getValue();

  const kode   = sj.getRange(SJ_ITEM_START_ROW, 2, numRows, 1).getValues(); // B
  const barang = sj.getRange(SJ_ITEM_START_ROW, 3, numRows, 1).getValues(); // C
  const qty    = sj.getRange(SJ_ITEM_START_ROW, 5, numRows, 1).getValues(); // E
  const sat    = sj.getRange(SJ_ITEM_START_ROW, 6, numRows, 1).getValues(); // F
  const ket    = sj.getRange(SJ_ITEM_START_ROW, 7, numRows, 1).getValues(); // G
  const srcPO  = sj.getRange(SJ_ITEM_START_ROW, 9, numRows, 1).getValues(); // I

  const harga    = inv.getRange(INV_DISPLAY_START_ROW, 7, numRows, 1).getValues(); // G
  const diskon   = inv.getRange(INV_DISPLAY_START_ROW, 8, numRows, 1).getValues(); // H
  const hargaNet = inv.getRange(INV_DISPLAY_START_ROW, 9, numRows, 1).getValues(); // I

  // ⬅️ BARU: map status semua NO PO (dari HISTORY_PO_MASUK), dibangun sekali
  // di luar loop biar gak scan sheet berkali-kali. Dipakai buat cek potongan
  // tambahan 3SET/5SET per baris item, berdasarkan PO SPESIFIK sumber baris
  // itu sendiri (SOURCE_PO), bukan noPOGabungan gabungan semua PO transaksi.
  const historyPOMap = buildHistoryPOStatusMap_();

  const rows = [];
  // ⬅️ BARU: SOURCE_PO ditampung terpisah (tidak lagi ikut nempel di ujung
  // array "rows"), karena kolom S sekarang dipakai formula LABA. Dipakai
  // untuk ditulis belakangan ke kolom U (lihat di bawah, setelah targetRowDet
  // diketahui).
  const srcPOList = [];

  for (let i = 0; i < kode.length; i++) {
    if (kode[i][0] == "") continue;

    let hppSatuan = Number(getExtraData(kode[i][0])) || 0;
    let qtyBarang = Number(qty[i][0]) || 0;
    let hppTotal = hppSatuan * qtyBarang;

    // ⬅️ BARU: potongan tambahan 3SET (-5%) / 5SET (-7,5%) dari HARGA NET
    // (kolom I INVOICE), berdasarkan status PO spesifik sumber baris ini.
    // Kalau baris ini asalnya dari beberapa PO (multi-PO per baris),
    // status yang dipakai adalah PO PERTAMA di SOURCE_PO.
    const persenSetPO = hitungPersenPotonganSetPO_(srcPO[i][0], historyPOMap);
    const hargaNetAsli = hargaNet[i][0];
    const hargaNetFinal = (persenSetPO > 0 && typeof Number(hargaNetAsli) === "number" && hargaNetAsli !== "")
      ? Number(hargaNetAsli) * (1 - persenSetPO)
      : hargaNetAsli;

    // ⬅️ UPDATE (swap kolom B<->C dan D<->E di DETAIL_TRANSAKSI): urutan
    // fisik sheet sekarang jadi NO SJ, TANGGAL, NO INV, NAMA TOKO, TEMPO,
    // KODE, BARANG, QTY, SATUAN, KETERANGAN, CREATED, UPDATED, HARGA,
    // DISKON, HARGA NET, BARIS SJ, METHOD, HPP TOTAL (kolom A sampai R,
    // 18 kolom). Nilai variabel JS (noSJ, noINV, tgl, tempo, toko) TIDAK
    // berubah maknanya, cuma URUTAN taruhnya di array ini yang disesuaikan
    // biar tetap nulis ke kolom yang benar setelah kolomnya ditukar posisi
    // di sheet.
    // ⬅️ PENTING: array ini SEKARANG BERHENTI DI KOLOM R (HPP TOTAL).
    // Kolom S (LABA) & U (SOURCE_PO) SENGAJA TIDAK dimasukkan ke sini
    // lagi, supaya setValues() di bawah tidak menimpa rumus LABA yang
    // ditulis terpisah setelahnya (lihat blok setFormulas di bawah).
    rows.push([
      noSJ,             // A
      tgl,              // B TANGGAL (ditukar posisi dgn NO INV)
      noINV,             // C NO INV (ditukar posisi dgn TANGGAL)
      toko,             // D NAMA TOKO (ditukar posisi dgn TEMPO)
      tempo,            // E TEMPO (ditukar posisi dgn NAMA TOKO)
      kode[i][0],       // F KODE
      barang[i][0],     // G BARANG
      qty[i][0],        // H QTY
      sat[i][0],        // I SATUAN
      ket[i][0],        // J KETERANGAN
      now,              // K CREATED
      now,              // L UPDATED
      harga[i][0],      // M HARGA
      diskon[i][0],     // N DISKON
      hargaNetFinal,    // O HARGA NET (sudah dipotong 3SET/5SET kalau berlaku)
      SJ_ITEM_START_ROW + i, // P BARIS SJ
      method,           // Q METHOD
      hppTotal          // R HPP TOTAL
    ]);

    srcPOList.push(srcPO[i][0] || "");
  }

  if (rows.length === 0) {
    throw new Error("Tidak ada barang terisi. Draft dibatalkan agar data lama tidak terhapus.");
  }

  // ⬅️ FIX: mapping status HISTORY_PO sekarang dibaca lewat
  // bacaMappingSumberSJDariSheet_(), yang parse JSON array di kolom I
  // (format baru hasil terapkanPilihanSJPopup_, support multi-PO per baris).
  // Sebelumnya baca kolom I/J mentah-mentah (poMapCol/lineMapCol) yang
  // udah gak sesuai lagi sama format JSON, jadi HISTORY_PO gak pernah kesync.
  const mappingBaru = bacaMappingSumberSJDariSheet_();
  updateStatusDoneSJHistoryPODelta_(oldMapping, mappingBaru);

  // ⬅️ MIGRASI (Oktober 2026): penulisan TRANSAKSI + DETAIL_TRANSAKSI
  // dipindah ke database (aksi gateway "simpanTransaksiDraft", lihat
  // GatewayTransaksi.js). Perilakunya sama persis dengan versi lokal lama:
  //  - NO SJ belum ada -> baris TRANSAKSI baru (ID = max ID + 1), H = DRAFT,
  //    N = PROGRESS; sudah ada -> B..N ditimpa, item DETAIL lama NO SJ ini dihapus
  //  - item ditulis ke DETAIL_TRANSAKSI A..R, S = rumus LABA (=O-R),
  //    U = SOURCE_PO
  kirimAksiGateway_("simpanTransaksiDraft", {
    header: {
      noSJ: noSJ, noInv: noINV, tgl: tgl, toko: toko, daerah: daerah,
      tempo: tempo, // G = Tempo (dihitung otomatis, ditulis ulang ke INVOICE!H8)
      now: now, method: method, noPol: noPol, sales: sales, total: totalBayar,
      keterangan: keterangan, // P
      noPO: noPOGabungan      // Q
    },
    items: rows,          // A..R
    sourcePO: srcPOList   // U
  });
  refreshSalinan_(["TRANSAKSI", "DETAIL_TRANSAKSI"]);

  // ⬅️ TAMBAHAN: sync PIUTANG dipanggil manual di sini, karena data TRANSAKSI
  // di atas ditulis lewat setValues() (dari script), BUKAN diketik langsung
  // sama user di sheet. onEdit trigger cuma kepicu kalau user ngedit cell
  // langsung lewat UI, jadi kalau simpanDraft() cuma ngandelin onEdit, sync
  // gak akan pernah jalan otomatis pas tombol "Simpan Draft" diklik.
  // ⬅️ MIGRASI: syncPenjualanFromTransaksi_ sudah memanggil syncTempoToko_
  // di dalamnya. Dulu syncTempoToko_ dipanggil LAGI di sini (dobel) -- sejak
  // TEMPO TOKO & PENJUALAN dibaca dari database, panggilan dobel itu bikin
  // Simpan Draft makin lama, jadi cukup sekali. Kalau sync PENJUALAN gagal,
  // TEMPO TOKO tetap dicoba sendiri.
  try {
    syncPenjualanFromTransaksi_(false);
  } catch (err) {
    Logger.log("Auto-sync PIUTANG (dari simpanDraft) error: " + err);
    try {
      syncTempoToko_(false);
    } catch (err2) {
      Logger.log("Auto-sync TEMPO TOKO (dari simpanDraft) error: " + err2);
    }
  }

  SpreadsheetApp.getUi().alert("Draft berhasil disimpan (" + rows.length + " item).");
}

/**
 * Lookup Nama Toko dari sheet TRANSAKSI berdasarkan No Invoice.
 * TRANSAKSI: kolom C = No INV, kolom E = Toko.
 */
function getNamaTokoByNoInv_(noINV) {
  if (!noINV) return "";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const trx = ss.getSheetByName("TRANSAKSI");
  const lastRow = trx.getLastRow();
  if (lastRow < 2) return "";

  const data = trx.getRange(2, 3, lastRow - 1, 3).getValues(); // C:E -> [noINV, tgl, toko]
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(noINV).trim()) {
      return data[i][2]; // kolom E = Toko
    }
  }
  return "";
}

/**
 * Lookup TANGGAL & TEMPO dari sheet TRANSAKSI berdasarkan No Invoice.
 * TRANSAKSI: kolom C = No INV, kolom D = Tanggal, kolom G = Tempo.
 */
function getTanggalTempoByNoInv_(noINV) {
  if (!noINV) return { tanggal: "", tempo: "" };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const trx = ss.getSheetByName("TRANSAKSI");
  const lastRow = trx.getLastRow();
  if (lastRow < 2) return { tanggal: "", tempo: "" };

  const data = trx.getRange(2, 3, lastRow - 1, 5).getValues(); // C:G -> [noINV, tgl, customer, daerah, tempo]
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(noINV).trim()) {
      return { tanggal: data[i][1], tempo: data[i][4] };
    }
  }
  return { tanggal: "", tempo: "" };
}

/**
 * (one-time migration LAMA — sudah pernah dijalankan, JANGAN dijalankan lagi)
 * Nyisipin kolom "NAMA TOKO" di posisi C pada DETAIL_TRANSAKSI.
 */
function migrasiTambahKolomNamaToko() {
  throw new Error("Fungsi migrasi lama ini sudah tidak berlaku: sejak Oktober 2026 DETAIL_TRANSAKSI ada di database_SANFELO (sheet lokal cuma salinan).");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const det = ss.getSheetByName("DETAIL_TRANSAKSI");

  det.insertColumnBefore(3);
  det.getRange(1, 3).setValue("NAMA TOKO");

  const lastRow = det.getLastRow();
  if (lastRow >= 2) {
    const noINVList = det.getRange(2, 2, lastRow - 1, 1).getValues();
    const namaTokoList = noINVList.map(function(r) {
      return [getNamaTokoByNoInv_(r[0])];
    });
    det.getRange(2, 3, namaTokoList.length, 1).setValues(namaTokoList);
  }

  SpreadsheetApp.getUi().alert(
    "Migrasi selesai: kolom NAMA TOKO ditambahkan di C, data lama sudah diisi otomatis.\n" +
    "JANGAN jalankan fungsi ini lagi."
  );
}

/**
 * (one-time migration BARU — jalankan MANUAL SEKALI SAJA lewat script
 * editor, pilih fungsi ini di dropdown, klik Run)
 * Nyisipin 2 kolom baru "TANGGAL" & "TEMPO" di posisi C & D pada
 * DETAIL_TRANSAKSI (sebelum NAMA TOKO yang sekarang ada di C), lalu
 * ngisi ulang data yang sudah ada sebelumnya pakai lookup dari
 * TRANSAKSI berdasarkan No INV.
 *
 * SETELAH dijalankan sekali dan berhasil, JANGAN dijalankan lagi.
 */
function migrasiTambahKolomTanggalTempo() {
  throw new Error("Fungsi migrasi lama ini sudah tidak berlaku: sejak Oktober 2026 DETAIL_TRANSAKSI ada di database_SANFELO (sheet lokal cuma salinan).");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const det = ss.getSheetByName("DETAIL_TRANSAKSI");

  det.insertColumnsBefore(3, 2); // sisip 2 kolom kosong sebelum C lama (NAMA TOKO)
  det.getRange(1, 3).setValue("TANGGAL");
  det.getRange(1, 4).setValue("TEMPO");

  const lastRow = det.getLastRow();
  if (lastRow >= 2) {
    const noINVList = det.getRange(2, 2, lastRow - 1, 1).getValues(); // kolom B = No INV

    const tglList = [];
    const tempoList = [];
    noINVList.forEach(function(r) {
      const info = getTanggalTempoByNoInv_(r[0]);
      tglList.push([info.tanggal]);
      tempoList.push([info.tempo]);
    });

    det.getRange(2, 3, tglList.length, 1).setValues(tglList);
    det.getRange(2, 4, tempoList.length, 1).setValues(tempoList);
    det.getRange(2, 3, tglList.length, 1).setNumberFormat("dd/mm/yyyy");
    det.getRange(2, 4, tempoList.length, 1).setNumberFormat("dd/mm/yyyy");
  }

  SpreadsheetApp.getUi().alert(
    "Migrasi selesai: kolom TANGGAL & TEMPO ditambahkan di C & D, data lama sudah diisi otomatis.\n" +
    "JANGAN jalankan fungsi ini lagi."
  );
}

/**
 * ⬅️ BARU (one-time migration) — jalankan MANUAL SEKALI SAJA lewat script
 * editor (pilih fungsi ini di dropdown, klik Run), SETELAH kamu update
 * simpanDraft() ke versi yang menulis LABA sebagai rumus.
 *
 * Fungsi ini mengonversi SEMUA baris LABA yang sudah ada di
 * DETAIL_TRANSAKSI (yang saat ini masih berupa ANGKA STATIS hasil
 * hitungan lama) menjadi RUMUS "=O{row}-R{row}", supaya:
 *  - Data lama & baru konsisten sama-sama pakai rumus.
 *  - Kalau suatu saat HARGA NET atau HPP BARANG historis dikoreksi,
 *    LABA baris itu ikut update otomatis, tidak perlu dihitung ulang manual.
 *
 * Baris yang diproses: baris 2 sampai lastRow, HANYA yang kolom A (NO SJ)
 * atau F (KODE) tidak kosong (baris benar-benar berisi data item).
 * Nilai HARGA NET (O) & HPP BARANG (R) yang sudah ada TIDAK diubah sama
 * sekali — cuma kolom LABA (S) yang ditimpa jadi rumus.
 *
 * SETELAH dijalankan sekali dan hasilnya dicek benar, JANGAN dijalankan
 * lagi (aman diulang sih, cuma menimpa rumus yang sama, tapi tidak perlu).
 */
function migrasiFormulaLaba() {
  throw new Error("Fungsi migrasi lama ini sudah tidak berlaku: sejak Oktober 2026 DETAIL_TRANSAKSI ada di database_SANFELO (sheet lokal cuma salinan).");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const det = ss.getSheetByName("DETAIL_TRANSAKSI");
  if (!det) throw new Error('Sheet "DETAIL_TRANSAKSI" tidak ditemukan.');

  const lastRow = det.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert("Tidak ada data di DETAIL_TRANSAKSI.");
    return;
  }

  const numRows = lastRow - 1;
  const dataAF = det.getRange(2, 1, numRows, 6).getValues(); // A..F, buat cek baris kosong

  const formulas = [];
  for (let i = 0; i < numRows; i++) {
    const noSJVal = String(dataAF[i][0] || "").trim();
    const kodeVal = String(dataAF[i][5] || "").trim(); // F = KODE
    const row = 2 + i;

    if (noSJVal === "" && kodeVal === "") {
      formulas.push([""]); // baris kosong, biarkan LABA kosong juga
    } else {
      formulas.push([buatRumusLaba_(row)]);
    }
  }

  det.getRange(2, DET_COL_LABA, formulas.length, 1).setFormulas(formulas);

  SpreadsheetApp.getUi().alert(
    "Migrasi selesai: " + numRows + " baris kolom LABA sudah dikonversi jadi rumus (=HARGA NET - HPP BARANG).\n" +
    "JANGAN jalankan fungsi ini berulang-ulang tanpa perlu."
  );
}

function getExtraData(kodeBarang) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const db = ss.getSheetByName("DATA_BERSIH");
  const data = db.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == kodeBarang) {
      return data[i][10];
    }
  }
  return "";
}

/**
 * Ambil mapping poNo+line LAMA (sebelum diedit) dari baris DETAIL_TRANSAKSI
 * yang masih ada untuk noSJ ini. Dipanggil SEBELUM baris lama dihapus.
 * ⬅️ DIPERBARUI: SOURCE_PO sekarang dibaca dari kolom U (21), bukan S (19)
 * lagi — karena kolom S sudah dipakai untuk rumus LABA (lihat simpanDraft()).
 */
// ⬅️ MIGRASI: parameter pertama sekarang DATA baris DETAIL_TRANSAKSI (A..U,
// hasil bacaDetailDB_()), bukan sheet lokal.
function ambilMappingLamaDariDetailTransaksi_(dataDetail, noSJ) {
  const data = dataDetail || [];
  if (data.length === 0) return [];
  const mapping = [];

  data.forEach(function(row) {
    if (String(row[0] || "").trim() !== noSJ) return; // A = NO SJ
    const raw = row[DET_COL_SOURCE_PO - 1]; // U = SOURCE_PO
    if (!raw) return;
    try {
      const sources = JSON.parse(raw);
      sources.forEach(function(src) {
        mapping.push({ poNo: src.poNo, line: src.line, qty: src.qty });
      });
    } catch (err) {
      Logger.log("Gagal parse SOURCE_PO lama DETAIL_TRANSAKSI: " + err);
    }
  });

  return mapping;
}

/**
 * Hitung jumlah HARI tempo untuk transaksi baru, berdasarkan:
 * 1) Semua NO PO di noPOGabungan (dipisah koma) dicek ke HISTORY_PO_MASUK
 *    kolom B (key) / kolom E (status). SALE -> 14, DISPLAY -> 90.
 *    Kalau ada beberapa PO dengan status beda, ambil yang PALING PANJANG.
 * 1b) ⬅️ FIX: teks Keterangan invoice (keteranganText) dicek juga — kalau
 *     mengandung "DISPLAY", jadi kandidat 90 hari, digabung pakai logika
 *     "paling panjang" yang sama dengan poin 1. Ini jaring pengaman kalau
 *     HISTORY_PO_MASUK belum/tidak sempat ditandai DISPLAY oleh admin.
 * 2) Kalau tidak ada satupun sinyal di atas, fallback ke sheet DISKON:
 *    kolom A = NAMA TOKO, kolom E = "n HARI".
 * 3) Kalau semua gagal -> null.
 */
/**
 * ⬅️ BARU: rekonstruksi noPOGabungan langsung dari kolom I (mapping
 * SOURCE_PO per-item, JSON) di sheet SURAT JALAN yang lagi kebuka --
 * dipakai sebagai fallback di simpanDraft() kalau H3 kosong (kejadian
 * normal begitu draft dibuka ulang lewat "Buka Draft"). Logic parsing-nya
 * sama persis kayak ambilNoPOGabunganDariItems_() di SJ_DRAFT_FETCH.js,
 * cuma sumbernya baca langsung dari sheet (belum lewat DETAIL_TRANSAKSI)
 * karena dipanggil SEBELUM baris di-scan jadi array item.
 */
function ambilNoPOGabunganDariKolomMappingSJ_(sj) {
  const lastRow = sj.getLastRow();
  if (lastRow < SJ_ITEM_START_ROW) return "";

  const numRows = lastRow - SJ_ITEM_START_ROW + 1;
  const mappingData = sj.getRange(SJ_ITEM_START_ROW, 9, numRows, 1).getValues(); // I

  const poSet = {};
  mappingData.forEach(function (r) {
    const raw = r[0];
    if (!raw) return;
    try {
      const sources = JSON.parse(raw);
      if (Array.isArray(sources)) {
        sources.forEach(function (src) {
          const poNo = String((src && src.poNo) || "").trim();
          if (poNo) poSet[poNo] = true;
        });
      }
    } catch (err) {
      // bukan JSON / format lama -> lewati, gak fatal
    }
  });

  return Object.keys(poSet).join(",");
}

function hitungTempoHariUntukTransaksi_(noPOGabungan, toko, keteranganText) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shDiskon = ss.getSheetByName("DISKON");

  const noPOList = String(noPOGabungan || "")
    .split(",")
    .map(function (s) { return s.trim(); })
    .filter(function (s) { return s; });

  let bestDays = null;

  if (noPOList.length > 0) {
    const historyPOMap = buildHistoryPOStatusMap_();

    noPOList.forEach(function (noPO) {
      const status = historyPOMap[noPO];
      let days = null;
      if (status === "SALE") days = 14;
      else if (status === "DISPLAY") days = 90;
      if (days !== null && (bestDays === null || days > bestDays)) {
        bestDays = days;
      }
    });
  }

  // ⬅️ FIX: gabungkan dengan sinyal "DISPLAY" dari teks Keterangan invoice.
  bestDays = gabungTempoDenganKeterangan_(bestDays, keteranganText);

  if (bestDays !== null) return bestDays;

  // Fallback ke DISKON kolom E berdasarkan NAMA TOKO (kolom A)
  if (shDiskon && toko) {
    const lastRowD = shDiskon.getLastRow();
    if (lastRowD >= 2) {
      const dataD = shDiskon.getRange(2, 1, lastRowD - 1, 5).getValues(); // A..E
      const targetToko = String(toko).trim();
      for (let i = 0; i < dataD.length; i++) {
        const namaToko = String(dataD[i][0] || "").trim();
        if (namaToko === targetToko) {
          const raw = String(dataD[i][4] || "").trim();
          const match = raw.match(/\d+/);
          if (match) return parseInt(match[0], 10);
          break;
        }
      }
    }
  }

  return null;
}

/**
 * ⬅️ BARU: bangun map status semua NO PO dari HISTORY_PO_MASUK sekali jalan.
 * { NO_PO: STATUS_UPPERCASE }  (kolom B = No PO, kolom E = Status)
 * Dipakai bareng oleh hitungTempoHariUntukTransaksi_() dan
 * hitungPersenPotonganSetPO_() biar gak scan sheet dua kali.
 *
 * ⬅️ FIX: toleransi human error — kalau kolom E (DISKON) bukan salah satu
 * dari SALE/DISPLAY/3SET/5SET, cek juga kolom R (METHOD) untuk kata
 * kunci yang sama. Ini karena di lapangan kadang admin salah taruh nilai
 * status (mis. "DISPLAY") ke kolom METHOD bukan ke kolom DISKON.
 * Kolom METHOD yang isinya normal (CASH/CREDIT) tidak match kata kunci
 * ini jadi tidak terpengaruh.
 */
function buildHistoryPOStatusMap_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("HISTORY_PO_MASUK");
  const map = {};
  if (!sh) return map;

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return map;

  const KATA_KUNCI_STATUS = ["SALE", "DISPLAY", "3SET", "5SET"];

  const data = sh.getRange(2, 2, lastRow - 1, 17).getValues(); // B..R
  data.forEach(function (r) {
    const noPO = String(r[0] || "").trim();
    if (!noPO) return;

    const statusDiskon = String(r[3] || "").trim().toUpperCase();   // E
    const statusMethod = String(r[16] || "").trim().toUpperCase();  // R

    if (KATA_KUNCI_STATUS.indexOf(statusDiskon) !== -1) {
      map[noPO] = statusDiskon;
    } else if (KATA_KUNCI_STATUS.indexOf(statusMethod) !== -1) {
      map[noPO] = statusMethod; // fallback: nyasar ke kolom METHOD
    } else {
      map[noPO] = statusDiskon;
    }
  });

  return map;
}

/**
 * ⬅️ BARU: hitung persen potongan tambahan (di luar diskon toko biasa)
 * berdasarkan status PO SPESIFIK sumber baris item ini, bukan gabungan
 * semua PO transaksi. Status 3SET -> potong 5%, status 5SET -> potong 7,5%.
 *
 * srcPoRaw = isi kolom SOURCE_PO baris SJ (JSON string, format
 * [{poNo, line, qty}, ...]). Kalau baris ini asalnya dari beberapa PO
 * sekaligus, dipakai status PO PERTAMA di array.
 *
 * Return: 0 (tanpa potongan tambahan), 0.05, atau 0.075.
 */
function hitungPersenPotonganSetPO_(srcPoRaw, historyPOMap) {
  if (!srcPoRaw) return 0;

  let sources;
  try {
    sources = JSON.parse(srcPoRaw);
  } catch (err) {
    return 0;
  }

  if (!Array.isArray(sources) || sources.length === 0) return 0;

  const poNoUtama = String(sources[0].poNo || "").trim();
  if (!poNoUtama) return 0;

  const status = historyPOMap[poNoUtama];
  if (status === "3SET") return 0.05;
  if (status === "5SET") return 0.075;
  return 0;
}

// Tanggal tempo = TGL KIRIM + jumlah hari tempo
function hitungTanggalTempoTransaksi_(tglKirim, tempoHari) {
  if (!tglKirim || tempoHari === null || tempoHari === undefined || isNaN(tempoHari)) return null;
  const d = new Date(tglKirim);
  d.setDate(d.getDate() + tempoHari);
  return d;
}

/**
 * ⬅️ REPAIR SEKALI JALAN (VERSI CEPAT) — recompute ulang kolom G (TEMPO)
 * di TRANSAKSI untuk SEMUA baris, berdasarkan status PO SAAT INI di
 * HISTORY_PO_MASUK.
 *
 * historyPOMap & diskonTempoMap dibangun SEKALI SAJA di awal (bukan per
 * baris) biar gak timeout.
 */
function perbaikiTempoTransaksiLama() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shDiskon = ss.getSheetByName("DISKON");

  // ⬅️ MIGRASI (Oktober 2026): TRANSAKSI dibaca FRESH dari database dan
  // perbaikannya dikirim lewat gateway (aksi "updateTempoTransaksiBatch"),
  // karena TRANSAKSI lokal sekarang cuma salinan.
  const dataTrxDB = bacaTransaksiDB_(); // A..Q
  if (dataTrxDB.length === 0) {
    SpreadsheetApp.getUi().alert("Tidak ada data di TRANSAKSI.");
    return;
  }

  const historyPOMap = buildHistoryPOStatusMap_();

  const diskonTempoMap = {};
  if (shDiskon) {
    const lastRowD = shDiskon.getLastRow();
    if (lastRowD >= 2) {
      const dataD = shDiskon.getRange(2, 1, lastRowD - 1, 5).getValues();
      dataD.forEach(function (r) {
        const namaToko = String(r[0] || "").trim();
        const raw = String(r[4] || "").trim();
        const match = raw.match(/\d+/);
        if (namaToko && match) diskonTempoMap[namaToko] = parseInt(match[0], 10);
      });
    }
  }

  function hitungTempoHariCepat_(noPOGabungan, toko, keteranganText) {
    const noPOList = String(noPOGabungan || "")
      .split(",")
      .map(function (s) { return s.trim(); })
      .filter(function (s) { return s; });

    let bestDays = null;
    noPOList.forEach(function (noPO) {
      const status = historyPOMap[noPO];
      let days = null;
      if (status === "SALE") days = 14;
      else if (status === "DISPLAY") days = 90;
      if (days !== null && (bestDays === null || days > bestDays)) {
        bestDays = days;
      }
    });

    // ⬅️ FIX: sama seperti hitungTempoHariUntukTransaksi_ — teks Keterangan
    // yang mengandung "DISPLAY" juga jadi kandidat 90 hari, supaya invoice
    // lama yang cuma ditandai DISPLAY di Keterangan (bukan di
    // HISTORY_PO_MASUK) ikut kebenerin pas repair massal ini dijalankan.
    bestDays = gabungTempoDenganKeterangan_(bestDays, keteranganText);

    if (bestDays !== null) return bestDays;

    if (toko) {
      const targetToko = String(toko).trim();
      if (diskonTempoMap.hasOwnProperty(targetToko)) {
        return diskonTempoMap[targetToko];
      }
    }

    return null;
  }

  // ⬅️ FIX: rentang diperlebar 1 kolom (D..Q jadi D..Q tetap 14 kolom, P
  // sudah ikut ke-cover karena P(16) ada di antara D(4) dan Q(17)) — jadi
  // TIDAK perlu nambah lebar range, cuma perlu ambil index kolom P yang
  // benar (lihat r[12] di bawah).
  const data = dataTrxDB.map(function (r) { return r.slice(3, 17); }); // D..Q

  let updatedCount = 0;
  const updates = [];

  data.forEach(function (r, i) {
    const tgl = r[0];
    const toko = r[1];
    const tempoLama = r[3];
    const keteranganRow = r[12]; // P (offset 16-4=12)
    const noPOGabungan = r[13];

    if (!tgl) return;

    const tempoHariBaru = hitungTempoHariCepat_(noPOGabungan, toko, keteranganRow);
    const tempoBaru = hitungTanggalTempoTransaksi_(tgl, tempoHariBaru);

    if (!tempoBaru) return;

    const tempoLamaStr = tempoLama ? Utilities.formatDate(new Date(tempoLama), "Asia/Jakarta", "yyyy-MM-dd") : "";
    const tempoBaruStr = Utilities.formatDate(tempoBaru, "Asia/Jakarta", "yyyy-MM-dd");

    if (tempoLamaStr !== tempoBaruStr) {
      updates.push({ noSJ: String(dataTrxDB[i][1] || "").trim(), tempo: tempoBaru }); // B = NO SJ
    }
  });

  if (updates.length > 0) {
    const hasilUpdate = kirimAksiGateway_("updateTempoTransaksiBatch", { updates: updates });
    updatedCount = hasilUpdate.updatedCount || 0;
    refreshSalinan_(["TRANSAKSI"]);
  }

  try {
    syncTempoToko_(false);
  } catch (err) {
    Logger.log("Re-sync TEMPO TOKO setelah repair error: " + err);
  }

  SpreadsheetApp.getUi().alert(
    "Selesai. " + updatedCount + " baris TRANSAKSI diperbaiki temponya, dan TEMPO TOKO sudah di-re-sync."
  );
}

  // ======================= PO_HISTORY.js (salinan persis) =======================
/* ======================================================
   PO_HISTORY.gs
   - poBaru()          : reset form PO_MASUK, auto-increment No PO
   - onEdit(e)         : SATU-SATUNYA simple trigger di seluruh project
                          (hasil merge dari handleOnEdit + onEdit lama
                          yang sebelumnya ada di file TEMPO_TOKO.gs)
   - syncToHistory()   : logic sync PO_MASUK -> HISTORY_PO_MASUK
   - loadSuratJalanFromHistory() : SJ multi-PO + auto-expand baris
   ====================================================== */

const PO_ITEM_START_ROW = 6;     // baris pertama item di PO_MASUK
const PO_ITEM_MAX_ROWS  = 200;   // batas scan baris item

const HISTORY_PO_ACTION_COL = 17; // Kolom Q
const HISTORY_PO_ACTION_EDIT = "EDIT";
const HISTORY_PO_ACTION_CANCEL = "CANCEL";
const HISTORY_PO_ACTION_UNCANCEL = "BATAL CANCEL";
const HISTORY_PO_BACKUP_COL = 19; // Kolom S (backup status sebelum cancel)

// ⬅️ BARU (Oktober 2026 — perbaikan alur nomor PO Masuk, atas permintaan
// user): key PropertiesService buat nandain "nomor PO yang lagi tampil di
// E2 masih CALON/preview, belum beneran diklaim dari counter bersama".
// Dipakai bareng poBaru() dan syncToHistory() di bawah -- lihat komentar
// di masing-masing fungsi untuk alur lengkapnya.
const PO_MASUK_PREVIEW_PROP_ = "PO_MASUK_NOMOR_PREVIEW";

// ⬅️ BARU (Oktober 2026): guard re-entrancy buat syncToHistory() -- lihat
// komentar di fungsi itu.
var SEDANG_SYNC_HISTORY_PO_MASUK_ = false;

// ⬅️ BARU (Oktober 2026, perbaikan bug nomor PO dobel/triple): kunci cache
// buat nandain "edit PO_MASUK paling baru". Dipakai syncToHistory() buat
// debounce -- lihat komentar di fungsi itu.
const PO_MASUK_EDIT_TS_KEY_ = "PO_MASUK_EDIT_TS";

/* ------------------------------------------------------
   1) PO BARU — reset form & auto increment No PO
   ------------------------------------------------------ */
// ⬅️ MIGRASI KE GATEWAY: dulu fungsi ini baca HISTORY_PO_MASUK lokal
// langsung. Sekarang HISTORY_PO_MASUK di inventory_SANFELO cuma MIRROR
// (IMPORTRANGE) dari database_SANFELO yang bisa telat beberapa saat --
// kalau dipakai buat cek ini di sini bisa salah (PO yang baru aja disimpan
// belum kelihatan di mirror, jadi dikira belum ada). Makanya baca FRESH
// langsung dari database lewat aksi "bacaSheet" gateway, bukan dari mirror.
function noPOSudahTersimpanGateway_(poNo) {
  const target = String(poNo).trim();
  const rows = bacaSheetGateway_("HISTORY_PO_MASUK", 2, 1, 2); // kolom A (JENIS) & B (NO PO)

  for (let i = 0; i < rows.length; i++) {
    const jenis = String(rows[i][0] || "").trim().toUpperCase();
    const noPo  = String(rows[i][1] || "").trim();

    if (jenis === "MASUK" && noPo === target) {
      return true;
    }
  }
  return false;
}

function poBaru() {
  // ⬅️ FIX: dulu fungsi ini baca counter SYSTEM!B3, hitung nomor
  // berikutnya, lalu nulis balik — TANPA lock. Kalau 2 device/tab
  // mengklik "PO Baru" hampir bersamaan, keduanya bisa baca counter yang
  // sama & dapat No PO yang SAMA PERSIS (nomor PO dobel), padahal
  // kembarannya di PO_KELUAR.js (poKeluarBaru()) sudah pakai
  // LockService buat kasus yang identik. Sekarang dikunci dengan pola
  // yang sama: otomatis release walau eksekusi timeout/crash, jadi gak
  // nyangkut permanen kayak flag manual di sheet.
  const lock = LockService.getScriptLock();
  const dapatLock = lock.tryLock(5000); // tunggu maks 5 detik

  if (!dapatLock) {
    SpreadsheetApp.getUi().alert("Masih ada proses PO Baru yang berjalan, coba lagi sebentar.");
    return;
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const masuk  = ss.getSheetByName("PO_MASUK");
    const system = ss.getSheetByName("SYSTEM");

    if (!masuk) throw new Error('Sheet "PO_MASUK" tidak ditemukan.');
    if (!system) throw new Error('Sheet "SYSTEM" tidak ditemukan.');

    const currentNo = String(masuk.getRange("E2").getValue() || "").trim();
    const sudahTersimpan = currentNo && noPOSudahTersimpanGateway_(currentNo);

    let nextNo;
    const docProps = PropertiesService.getDocumentProperties();
    const previewSebelumnya = docProps.getProperty(PO_MASUK_PREVIEW_PROP_);

    if (currentNo && !sudahTersimpan) {
        // CABANG "REUSE" — SYSTEM!B3 TIDAK DISENTUH SAMA SEKALI
        nextNo = currentNo;
        // ⬅️ BARU (Oktober 2026): kalau nomor yang lagi tampil ini BUKAN
        // preview yang barusan kita kasih sendiri (mis. admin ngetik
        // manual nomor lama buat dipakai ulang), hapus flag preview lama
        // -- supaya syncToHistory() nanti gak keliru nyangka nomor REUSE
        // manual ini masih "belum diklaim" terus nyoba klaim dari counter.
        // Kalau nomornya SAMA PERSIS dengan preview yang barusan kita
        // kasih (admin cuma klik "PO Masuk Baru" dua kali tanpa sempat
        // nyimpan), flag DIBIARKAN apa adanya -- nomor itu memang masih
        // berstatus "belum diklaim beneran".
        if (previewSebelumnya && previewSebelumnya !== nextNo) {
          docProps.deleteProperty(PO_MASUK_PREVIEW_PROP_);
        }
    } else {
        // ⬅️ FIX (Oktober 2026, ditemukan dari laporan user — data dobel di
        // HISTORY_PO_MASUK): kalau currentNo ADA tapi SUDAH tersimpan, itu
        // biasanya tandanya form ini lagi nampilin PO hasil "EDIT" dari
        // HISTORY_PO_MASUK (handleHistoryEdit), BUKAN draft kosong. Dulu di
        // sini LANGSUNG diam-diam generate No PO BARU dari counter tanpa
        // peringatan apapun — kalau admin gak sengaja klik menu "🆕 PO Masuk
        // Baru" pas lagi di tengah proses EDIT (alih-alih langsung ngetik
        // item di form yang sudah keisi), PO yang di-edit itu DITINGGAL apa
        // adanya di database (nomor lama), sementara form malah lanjut jalan
        // dengan NOMOR BARU — begitu admin ngisi/ngedit item di form, onEdit
        // nyinkron item itu ke HISTORY_PO_MASUK di BAWAH NOMOR BARU itu,
        // jadi nyangkut sebagai baris "orphan" (gak ada SJ-nya, TGL PO bisa
        // kosong karena ikut ke-clear bareng field lain di bawah). Sekarang
        // admin diminta konfirmasi dulu sebelum form di-reset, supaya kalau
        // ini kejadian gak sengaja, PO yang sedang di-edit gak sampai
        // ditinggal/keduplikat.
        if (sudahTersimpan) {
          const respKonfirmasi = SpreadsheetApp.getUi().alert(
            "No PO " + currentNo + " yang lagi tampil di form ini SUDAH TERSIMPAN di HISTORY_PO_MASUK.",
            "Lanjut akan mengosongkan form ini dan membuat PO BARU dengan nomor baru.\n\n" +
            "Kalau maksudnya mau LANJUTIN EDIT PO " + currentNo + " (misalnya abis klik ACTION \"EDIT\" di HISTORY_PO_MASUK), " +
            "klik Batal — form ini TIDAK perlu di-reset, langsung aja edit item yang sudah keisi.",
            SpreadsheetApp.getUi().ButtonSet.OK_CANCEL
          );
          if (respKonfirmasi !== SpreadsheetApp.getUi().Button.OK) {
            return; // dibatalkan — form & No PO yang lagi di-edit dibiarkan apa adanya
          }
        }

        // ⬅️ MIGRASI KE COUNTER BERSAMA (Oktober 2026): dulu nomor PO Masuk
        // berikutnya diambil dari counter LOKAL (SYSTEM!B3/C3) yang cuma
        // tahu nomor yang dipakai di SANFELO sendiri -- sementara nomor PO
        // Masuk di SANFELO berjalan BERSAMAAN dengan inventory_SL (sistem
        // terpisah). Counter lokal ini gak tahu-menahu nomor yang sudah
        // dipakai SL, begitu juga sebaliknya, jadi secara teori bisa
        // tabrakan (kebetulan aja belum pernah separah itu). Sekarang
        // nomor diminta dari SATU counter bersama di database_SANFELO,
        // yang juga dipakai SL, jadi nomornya beneran saling melanjutkan
        // di kedua sistem.
        //
        // ⬅️ BARU (Oktober 2026 — perbaikan alur, atas permintaan user):
        // dulu baris di bawah ini manggil "getNomorPoMasukBersama" (yang
        // BENERAN nge-increment counter) SAAT ITU JUGA begitu form "PO
        // Masuk Baru" dibuka -- jadi nomornya "kebakar"/diklaim walau
        // admin belum tentu jadi nyimpen PO-nya. Kalau SANFELO & SL
        // sama-sama buka form ini hampir bersamaan, dua-duanya langsung
        // dapat nomor BEDA (mis. SL dapat #231 duluan, begitu SANFELO
        // buka form-nya kebagian #232), padahal belum tentu jadi dipakai
        // dua-duanya.
        //
        // Sekarang dipakai aksi "peekNomorPoMasukBersama" -- cuma INTIP
        // nomor berikutnya, gak nulis/nambah apapun ke counter. Jadi kalau
        // dibuka bersamaan, SANFELO & SL nampilin nomor CALON yang SAMA.
        // Nomor ini ditandai sebagai "preview" (lewat PO_MASUK_PREVIEW_PROP_)
        // dan baru BENERAN diklaim (lewat "getNomorPoMasukBersama" yang
        // asli) pas admin mulai nyimpan barang pertama -- lihat
        // syncToHistory() di bawah. Siapa yang nyimpen duluan itu yang
        // dapat nomornya; yang satunya, begitu DIA juga mulai nyimpen,
        // otomatis ke-deteksi nomor previewnya sudah kepake & diganti ke
        // nomor berikutnya + admin dikasih tau lewat alert.
        const hasilNomor = kirimAksiGateway_("peekNomorPoMasukBersama", { asal: "SANFELO" });
        nextNo = hasilNomor.nomor;
        docProps.setProperty(PO_MASUK_PREVIEW_PROP_, nextNo);
    }

    // Header PO_MASUK (Aman dihapus karena tidak ada rumus)
    masuk.getRange("B2").clearContent(); // Customer — B3 & B4 formula XLOOKUP, biarin auto-update
    masuk.getRange("B3").setFormula('=IFERROR(VLOOKUP(B2;DISKON!A:F;6;FALSE);"")');
    masuk.getRange("B4").setFormula('=IFERROR(VLOOKUP(B2;DISKON!A:G;7;FALSE);"")');
    masuk.getRange("E3:E4").clearContent(); // Tgl PO, Tgl Kirim
    masuk.getRange("H4").clearContent();    // Method

    // Set No PO (baru atau nomor lama yang dipakai ulang)
    masuk.getRange("E2").setValue(nextNo);

    // BERSIHKAN HANYA KOLOM B (KODE) DAN D (QTY) - RUMUS DI A, C, E, F, G, H AMAN!
    masuk.getRange(PO_ITEM_START_ROW, 2, PO_ITEM_MAX_ROWS, 1).clearContent(); // Kolom B
    masuk.getRange(PO_ITEM_START_ROW, 4, PO_ITEM_MAX_ROWS, 1).clearContent(); // Kolom D

    SpreadsheetApp.getUi().alert("PO Masuk baru siap diisi. No PO: " + nextNo);
  } finally {
    lock.releaseLock();
  }
}

/* ------------------------------------------------------
   1b) CEK TEMPO TOKO SEBELUM BUAT SURAT JALAN
   ------------------------------------------------------ */
function cekTempoTokoBlokir_(poNoList) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hist = ss.getSheetByName("HISTORY_PO_MASUK");
  const tempo = ss.getSheetByName("TEMPO TOKO");

  if (!hist) throw new Error('Sheet "HISTORY_PO_MASUK" tidak ditemukan.');
  if (!tempo) throw new Error('Sheet "TEMPO TOKO" tidak ditemukan.');

  const histLastRow = hist.getLastRow();
  if (histLastRow < 2) return false;

  const histData = hist.getRange(2, 1, histLastRow - 1, 4).getValues(); // A..D

  // Kumpulin nama toko unik dari semua No PO yang ada di H3 (bisa multi-PO)
  const namaTokoSet = {};

  poNoList.forEach(function(poNo) {
    for (let i = 0; i < histData.length; i++) {
      const jenis = String(histData[i][0] || "").trim().toUpperCase();
      const rowPoNo = String(histData[i][1] || "").trim();

      if (jenis === "MASUK" && rowPoNo === poNo) {
        const namaToko = String(histData[i][3] || "").trim(); // kolom D CUSTOMER
        if (namaToko) namaTokoSet[namaToko] = true;
        break; // cukup ambil baris pertama yang match buat PO ini
      }
    }
  });

  const namaTokoList = Object.keys(namaTokoSet);
  if (namaTokoList.length === 0) return false;

  const tempoLastRow = tempo.getLastRow();
  if (tempoLastRow < 2) {
    SpreadsheetApp.getUi().alert('Sheet "TEMPO TOKO" masih kosong, tidak bisa cek status toko. Surat Jalan tidak bisa dibuat.');
    return true;
  }

  const tempoData = tempo.getRange(2, 7, tempoLastRow - 1, 2).getValues(); // G=nama, H=status

  for (let t = 0; t < namaTokoList.length; t++) {
    const namaToko = namaTokoList[t];
    let status = "";

    for (let i = 0; i < tempoData.length; i++) {
      const namaTempo = String(tempoData[i][0] || "").trim();

      if (namaTempo === namaToko) {
        status = String(tempoData[i][1] || "").trim().toUpperCase();
        break;
      }
    }

    // Hanya blokir kalau memang ketemu dan statusnya LOCKED
    if (status === "LOCKED") {
      SpreadsheetApp.getUi().alert(
        'Toko "' + namaToko + '" punya pembayaran yang jatuh tempo (LOCKED). Surat Jalan tidak bisa dibuat.'
      );
      return true;
    }
  }

  return false;
}

/* ------------------------------------------------------
   2) TRIGGER OTOMATIS — onEdit (SATU-SATUNYA di project)
   ------------------------------------------------------ */
/* ======================================================
   onEdit(e) — HASIL MERGE dari:
   1. handleOnEdit(e) yang sebelumnya di PO_HISTORY.gs
   2. onEdit(e) dispatcher yang sebelumnya di TEMPO_TOKO.gs
      (yang manggil onEditTempoToko_, onEditTransaksiStatus_,
      onEditDetailStatus_)

   Semua logic lama TETAP SAMA, cuma digabung jadi satu titik masuk
   biar gak ada dua function onEdit yang saling timpa di project.
   ====================================================== */
// kolom di HISTORY_PO_MASUK yang datanya dipakai di Laporan PO Masuk
// A=JENIS, B=NO PO, C=TGL PO, D=CUSTOMER, I=KODE_BARANG, J=NAMA_BARANG, K=QUANTITY
const HISTORY_PO_LAPORAN_WATCH_COLS = [1, 2, 3, 4, 9, 10, 11, 14]; // ⬅️ tambah 14 (kolom N)

function handleOnEdit(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();
  const row = e.range.getRow();
  const col = e.range.getColumn();

  if (sheetName === "PO_MASUK") {
    try {
      syncToHistory();
    } catch (err) {
      Logger.log("Sync error: " + err);
    }
    return;
  }

  if (sheetName === "PO_KELUAR") {
    if (row === 4 && col === 5) { // E4 = kolom TGL
      try {
        generateNoPOKeluarDariTanggal_(sheet);
      } catch (err) {
        Logger.log("Generate No PO Keluar error: " + err);
      }
    }
    return;
  }

  if (sheetName === "SURAT JALAN") {
    if (col === 8 && row === 3) { // H3
      try {
        const poNo = String(e.range.getValue() || "").trim();
        const oldValue = String(e.oldValue || "").trim();

        if (poNo) {
          const poNoListCek = poNo.split(",").map(function(s) { return s.trim(); }).filter(Boolean);
          const blokir = cekTempoTokoBlokir_(poNoListCek);
          if (blokir) {
            sheet.getRange("H3").clearContent();
            return;
          }
        }

        if (poNo && !oldValue) {
          const noSJSudahAda = String(sheet.getRange("H6").getValue() || "").trim();
          if (!noSJSudahAda) {
            transaksiBaru();
          }
        }

        if (!poNo) {
          const noSJval = String(sheet.getRange("H6").getValue() || "").trim();
          if (noSJval) {
            cetakUlangSJDariNoSJ_(noSJval);
            return;
          }
        }

        loadSuratJalanFromHistory();
        adjustSJItemBorders();
        adjustInvoiceItemBorders();
      } catch (err) {
        Logger.log("Load SJ error: " + err);
        SpreadsheetApp.getUi().alert("Gagal load Surat Jalan: " + err.message);
      }
      return;
    }

    if (col === 8 && row === 6) { // H6 = No SJ (reprint by No SJ, hanya kalau No PO/H3 kosong)
      try {
        const noPOval = String(sheet.getRange("H3").getValue() || "").trim();
        const noSJval = String(e.range.getValue() || "").trim();
        const oldVal  = String(e.oldValue || "").trim();

        if (noPOval) return;

        if (noSJval && !oldVal) {
          cetakUlangSJDariNoSJ_(noSJval);
        }
      } catch (err) {
        Logger.log("Reprint SJ error: " + err);
        SpreadsheetApp.getUi().alert("Gagal reprint: " + err.message);
      }
      return;
    }

    if (row >= SJ_DISPLAY_START_ROW && col >= 1 && col <= 8) {
      try {
        adjustSJItemBorders();
      } catch (err) {
        Logger.log("Border SJ error: " + err);
      }
    }
    return;
  }

  if (sheetName === "INVOICE") {
    if (row >= INV_DISPLAY_START_ROW && col >= 1 && col <= 9) {
      try {
        adjustInvoiceItemBorders();
      } catch (err) {
        Logger.log("Border Invoice error: " + err);
      }
    }
    return;
  }

  // ⬅️ MIGRASI (Oktober 2026): TRANSAKSI, DETAIL_TRANSAKSI, dan PENJUALAN
  // sekarang SALINAN dari database_SANFELO (IMPORTRANGE), tidak bisa diketik
  // langsung. Handler lama di sini (convert tempo kolom G, auto-sync
  // PENJUALAN saat C/D/E/J/M diedit, guard dropdown STATUS ORDER DONE,
  // warna status DETAIL) tidak berlaku lagi:
  //  - ubah status -> menu 🏢 SANFELO > ✏️ Ubah Status Transaksi
  //    (guard "DONE hanya kalau LUNAS" sekarang dicek di database)
  //  - sync PENJUALAN sudah otomatis setiap Simpan Draft
  // Kalau ada yang mengetik di sheet salinan, ketikannya langsung
  // dibatalkan supaya IMPORTRANGE tidak rusak (lihat TRANSAKSI_DB.js).
  if (sheetName === "TRANSAKSI" || sheetName === "DETAIL_TRANSAKSI" || sheetName === "PENJUALAN") {
    pulihkanSheetSalinan_(e, sheetName);
    return;
  }

  if (sheetName === "TEMPO TOKO") {
    // ⬅️ MERGED dari onEdit dispatcher lama (TEMPO_TOKO.gs):
    // handle klik dropdown kolom H (ringkasan) jadi UNLOCKED -> minta
    // password + jumlah hari perpanjangan tempo.
    try {
      onEditTempoToko_(e, sheet);
    } catch (err) {
      Logger.log("onEditTempoToko_ error: " + err);
    }
    return;
  }

  // ⬅️ BARU: sheet KAS KECIL - tiap ada edit manual di area data (kolom TANGGAL
  // s/d SALDO, mulai baris data), otomatis sort ulang berdasarkan TANGGAL
  // (ascending) dan hitung ulang SALDO biar tetap konsisten sama urutannya.
  // (Sort ini dipicu oleh script sendiri via setValues, jadi TIDAK bakal
  // memicu onEdit lagi secara berulang / infinite loop.)
  if (sheetName === KAS_KECIL_SHEET_NAME) {
    if (row >= KAS_KECIL_FIRST_DATA_ROW && col >= 1 && col <= KAS_KECIL_LAST_COL) {
      try {
        sortKasKecilByTanggal(sheet);
      } catch (err) {
        Logger.log("Sort KAS KECIL error: " + err);
      }
    }
    return;
  }

  if (sheetName === "HISTORY_PO_MASUK") {
    if (row === 1 && (col === 18 || col === 20 || col === 22)) {
      try {
        applyHistoryFilter();
      } catch (err) {
        Logger.log("Filter error: " + err);
      }
      return;
    }

    // ACTION dropdown di kolom Q
    if (row >= 2 && col === HISTORY_PO_ACTION_COL) {
      try {
        const action = String(e.range.getValue() || "").trim().toUpperCase();

        if (!action) return;

        if (action === HISTORY_PO_ACTION_EDIT) {
          handleHistoryEdit(row);
          return;
        }

        if (action === HISTORY_PO_ACTION_CANCEL) {
          handleHistoryCancel(row);
          return;
        }

        if (action === HISTORY_PO_ACTION_UNCANCEL) {
          handleHistoryUncancel(row);
          return;
        }

        e.range.clearContent();
        SpreadsheetApp.getUi().alert("ACTION tidak dikenali: " + action);

      } catch (err) {
        Logger.log("History PO Action error: " + err);
        SpreadsheetApp.getUi().alert("Error ACTION HISTORY_PO_MASUK: " + err.message);
      }
      return;
    }

    // ⬅️ TAMBAHAN: auto-refresh Laporan PO Masuk (file terpisah) kalau
    // ada edit di kolom yang datanya dipakai laporan (JENIS/NO PO/TGL PO/
    // CUSTOMER/KODE_BARANG/NAMA_BARANG/QUANTITY/NILAI N). Ini yang bikin
    // file Laporan (+ LEADERBOARD) auto-update tiap kali HISTORY_PO_MASUK diedit.
    if (row >= 2 && HISTORY_PO_LAPORAN_WATCH_COLS.indexOf(col) !== -1) {
      try {
        refreshLaporanPOMasuk_();
      } catch (err) {
        Logger.log("Refresh Laporan PO Masuk error: " + err);
      }
      return;
    }
  }

  // ⬅️ BARU (migrasi HISTORY_PO_KELUAR): kolom H (QTY TERKIRIM) sengaja
  // TIDAK ikut di-mirror IMPORTRANGE, jadi setiap kali admin ngetik di
  // situ, nilainya diforward ke database_SANFELO lewat gateway (lihat
  // handleHistoryPOKeluarQtyEdit_ di PO_KELUAR.js).
  if (sheetName === "HISTORY_PO_KELUAR") {
    if (row >= 2 && col === 8) {
      try {
        handleHistoryPOKeluarQtyEdit_(row, e.range.getValue());
      } catch (err) {
        Logger.log("Forward QTY TERKIRIM PO Keluar error: " + err);
      }
    }
    return;
  }

  // ⬅️ TAMBAHAN: auto-sort sheet DISKON berdasarkan kolom A (NAMA TOKO) A-Z,
  // kolom B-E ikut sesuai barisnya. Jalan tiap ada edit di sheet DISKON
  // (row > 1 biar header di row 1 gak ikut ke-sort/gak trigger apa-apa).
  if (sheetName === "DISKON") {
    if (row > 1) {
      try {
        autoSortDiskon_(sheet);
      } catch (err) {
        Logger.log("Sort DISKON error: " + err);
      }
    }
    return;
  }

  if (sheetName === "TANDA TERIMA BRG" && e.range.getA1Notation() === "B3") {
    loadTandaTerimaFromHistory();
  }
}

/**
 * Sort sheet DISKON berdasarkan kolom A (NAMA TOKO), A-Z.
 * HANYA kolom A-G yang ikut disort — kolom lain (H dst, kalau ada) TIDAK disentuh.
 */
function autoSortDiskon_(sheet) {
  const lastRow = sheet.getLastRow();

  // minimal harus ada 2 baris data (row 2 ke atas) buat worth di-sort
  if (lastRow < 3) return;

  const range = sheet.getRange(2, 1, lastRow - 1, 7); // A-G = 7 kolom
  range.sort({ column: 1, ascending: true });
}

/**
 * Buat dipanggil manual sekali dari Apps Script editor (pilih function ini
 * di dropdown, klik Run), khusus untuk sort data yang udah ada SEKARANG.
 * Setelah ini, edit-edit selanjutnya bakal auto-sort sendiri lewat handleOnEdit.
 */
function sortDiskonManual() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("DISKON");
  if (!sheet) throw new Error('Sheet "DISKON" tidak ditemukan.');

  autoSortDiskon_(sheet);
  SpreadsheetApp.getUi().alert("Sheet DISKON berhasil di-sort A-Z berdasarkan Nama Toko.");
}

/* ------------------------------------------------------
   3) FILTER IN-PLACE
   ------------------------------------------------------ */
function applyHistoryFilter() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hist = ss.getSheetByName("HISTORY_PO_MASUK");

  if (!hist) throw new Error('Sheet "HISTORY_PO_MASUK" tidak ditemukan.');

  const lastRow = hist.getLastRow();
  if (lastRow < 2) return;

  const jenisFilter    = String(hist.getRange("R1").getValue() || "Semua").trim().toLowerCase();
  const customerFilter = String(hist.getRange("T1").getValue() || "Semua").trim().toLowerCase();
  const supplierFilter = String(hist.getRange("V1").getValue() || "Semua").trim().toLowerCase();

  hist.showRows(2, lastRow - 1);
  const data = hist.getRange(2, 1, lastRow - 1, 5).getValues();

  for (let i = 0; i < data.length; i++) {
    const jenis    = String(data[i][0] || "").trim().toLowerCase();
    const customer = String(data[i][3] || "").trim().toLowerCase();
    const supplier = String(data[i][4] || "").trim().toLowerCase();

    const matchJenis    = (jenisFilter === "semua")    || (jenis === jenisFilter);
    const matchCustomer = (customerFilter === "semua") || (customer === customerFilter);
    const matchSupplier = (supplierFilter === "semua") || (supplier === supplierFilter);

    if (!(matchJenis && matchCustomer && matchSupplier)) {
      hist.hideRows(i + 2);
    }
  }
}

/* ------------------------------------------------------
   4) SYNC LOGIC — PO_MASUK -> HISTORY_PO_MASUK
   ------------------------------------------------------ */
function syncToHistory() {
  // ⬅️ DIREVISI (Oktober 2026, setelah muncul notifikasi "Melebihi jumlah
  // eksekusi maksimum"): versi sebelumnya membuat SETIAP ketikan menunggu
  // giliran sampai 60 detik -- tiap yang menunggu tetap memakai 1 dari ~30
  // slot eksekusi Apps Script, jadi ketikan cepat bisa menghabiskan slot.
  //
  // Aturan sekarang (tanpa antrian panjang):
  //   1) Tiap ketikan menulis "penanda" baru ke cache.
  //   2) Hanya SATU eksekusi yang memegang DocumentLock dan menjalankan sync.
  //      Eksekusi lain menunggu paling lama ~8 detik; kalau lock masih
  //      dipegang, dia langsung KELUAR (tidak memakai slot lama-lama).
  //   3) Setelah selesai, pemegang lock mengecek penanda: kalau berubah
  //      (ada ketikan baru selama dia bekerja), dia mengulang sync SEKALI
  //      lagi dengan isi sheet terbaru. Jadi ketikan terakhir selalu
  //      tersimpan, tanpa ada yang menumpuk.
  //   4) Nomor PO tetap diklaim cuma sekali (klaim -> tulis E2 -> hapus flag
  //      preview), karena semua dikerjakan satu per satu di dalam lock.
  const cache = CacheService.getScriptCache();
  try { cache.put(PO_MASUK_EDIT_TS_KEY_, String(Date.now()) + "_" + Math.random(), 300); } catch (e) {}

  const lock = LockService.getDocumentLock();
  const daftarPesan = [];
  let errTersimpan = null;

  for (let putaran = 0; putaran < 6; putaran++) {
    if (!lock.tryLock(putaran === 0 ? 8000 : 3000)) break; // ada yang sedang sync -- dia yang menyelesaikan

    let penandaMulai = null;
    try {
      try { penandaMulai = cache.get(PO_MASUK_EDIT_TS_KEY_); } catch (e) {}
      SEDANG_SYNC_HISTORY_PO_MASUK_ = true;
      const pesan = syncToHistoryInternal_();
      if (pesan) daftarPesan.push(pesan);
    } catch (err) {
      errTersimpan = err;
      break;
    } finally {
      SEDANG_SYNC_HISTORY_PO_MASUK_ = false;
      // Paksa semua tulisan sheet (mis. E2 hasil renumber) benar-benar
      // tersimpan SEBELUM lock dilepas, supaya eksekusi berikutnya tidak
      // membaca E2 versi lama (nomor PO lama -> baris "hantu").
      try { SpreadsheetApp.flush(); } catch (eFlush) {}
      lock.releaseLock();
    }

    let penandaSekarang = null;
    try { penandaSekarang = cache.get(PO_MASUK_EDIT_TS_KEY_); } catch (e) {}
    if (penandaSekarang === penandaMulai) break; // tidak ada ketikan baru -- selesai
  }

  // Alert ditampilkan SETELAH lock dilepas, supaya dialog yang nunggu
  // admin klik OK tidak menahan proses lain.
  daftarPesan.forEach(function (pesan) {
    try { SpreadsheetApp.getUi().alert(pesan); } catch (e) { Logger.log(pesan); }
  });
  if (errTersimpan) throw errTersimpan;
}

function syncToHistoryInternal_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masuk = ss.getSheetByName("PO_MASUK");
  const hist  = ss.getSheetByName("HISTORY_PO_MASUK");

  if (!masuk) throw new Error('Sheet "PO_MASUK" tidak ditemukan.');
  if (!hist) throw new Error('Sheet "HISTORY_PO_MASUK" tidak ditemukan.');

  const timestamp = Utilities.formatDate(new Date(), "Asia/Jakarta", "dd/MM/yyyy HH:mm:ss");

  let poNo     = masuk.getRange("E2").getValue();
  const poTgl    = masuk.getRange("E3").getValue();
  const kirimTgl = masuk.getRange("E4").getValue();
  const customer = masuk.getRange("B2").getValue();
  const alamatM  = masuk.getRange("B3").getValue();
  const telpM    = masuk.getRange("B4").getValue();
  const method   = masuk.getRange("H4").getValue();
  const diskon   = masuk.getRange("H3").getValue(); // ⬅️ TAMBAHAN: field DISKON

  if (!poNo) return;

  const itemsMasuk = masuk
    .getRange(PO_ITEM_START_ROW, 1, PO_ITEM_MAX_ROWS, 8)
    .getValues();

  const rows = [];

  itemsMasuk.forEach(function(item, index) {
    const lineNo    = item[0] || index + 1;
    const kode      = item[1];
    const nama      = item[2];
    const qty       = item[3];
    const status    = item[4];
    const hargaUnit = item[5];
    const total     = item[7]; // kolom H

    if (!kode) return;

    rows.push([
      "MASUK",       // A JENIS
      poNo,          // B NO PO
      poTgl,         // C TGL PO
      customer,      // D CUSTOMER
      diskon,        // E DISKON (dulu SUPPLIER, selalu "" utk JENIS=MASUK) ⬅️ TAMBAHAN
      alamatM,       // F ALAMAT
      telpM,         // G TELP
      kirimTgl,      // H TGL_KIRIM
      kode,          // I KODE
      nama,          // J NAMA
      qty,           // K QTY
      status,        // L STATUS
      hargaUnit,     // M HARGA_UNIT
      total,         // N TOTAL
      timestamp,     // O TIMESTAMP_SIMPAN
      lineNo,        // P LINE
      "",            // Q ACTION
      method         // R METHOD
    ]);
  });

  // ⬅️ BARU (Oktober 2026 — perbaikan alur nomor PO, atas permintaan user;
  // lihat juga poBaru() dan aksiPeekNomorPoMasukBersama_ di
  // DatabaseGateway.js): nomor PO yang lagi tampil di E2 bisa aja masih
  // berstatus "preview" (dikasih lewat intip/peek waktu form dibuka, BELUM
  // diklaim beneran dari counter bersama). Begitu admin beneran mulai
  // ngisi barang (rows.length > 0 buat PERTAMA KALInya buat PO ini),
  // sekarang baru nomor itu diklaim SUNGGUHAN dari counter bersama. Kalau
  // nomor yang di-preview ternyata udah keburu dipakai sistem lain
  // (SANFELO/SL) di antara waktu form dibuka dan sekarang mulai nyimpan,
  // counter bersama otomatis kasih nomor BERIKUTNYA yang masih kosong --
  // E2 & rows disesuaikan ke nomor itu, admin dikasih tau lewat alert
  // (BUKAN diam-diam kayak sebelumnya), barang yang sudah diisi TIDAK
  // hilang.
  let pesanRenumber = null;
  if (rows.length > 0) {
    const docProps = PropertiesService.getDocumentProperties();
    const previewSaatIni = docProps.getProperty(PO_MASUK_PREVIEW_PROP_);
    if (previewSaatIni && previewSaatIni === String(poNo).trim()) {
      // Urutan penting (anti dobel): 1) klaim, 2) tulis E2 kalau nomornya
      // berubah, 3) BARU hapus tanda preview. Kalau proses berhenti di
      // tengah jalan, tanda preview masih ada -> aman dicoba ulang; tidak
      // pernah ada kondisi "sudah diklaim tapi E2 masih nomor lama tanpa
      // tanda".
      const hasilKlaim = kirimAksiGateway_("getNomorPoMasukBersama", { asal: "SANFELO" });
      const nomorKlaim = String(hasilKlaim.nomor).trim();

      if (nomorKlaim !== String(poNo).trim()) {
        const poNoLama = poNo;
        poNo = nomorKlaim;
        rows.forEach(function (r) { r[1] = nomorKlaim; }); // B NO PO di tiap baris
        masuk.getRange("E2").setValue(nomorKlaim);
        pesanRenumber =
          "No PO berubah dari " + poNoLama + " menjadi " + nomorKlaim + ".\n\n" +
          "Ini kejadian karena nomor " + poNoLama + " baru saja dipakai duluan oleh sistem lain " +
          "(SANFELO/SL) sepersekian detik sebelum kamu mulai menyimpan. Barang yang sudah kamu isi " +
          "TIDAK hilang -- tetap tersimpan, cuma dengan nomor PO yang baru ini.";
      }
      docProps.deleteProperty(PO_MASUK_PREVIEW_PROP_);
    }
  }

  // ⬅️ MIGRASI KE GATEWAY: dulu fungsi ini hapus baris lama
  // (deleteHistoryPORows_), insert baris baru, dan bersihkan duplikat
  // (bersihkanDuplikatUntukPO_) LANGSUNG ke HISTORY_PO_MASUK lokal, tiga
  // langkah terpisah. HISTORY_PO_MASUK di inventory_SANFELO sekarang cuma
  // MIRROR (IMPORTRANGE) dari database_SANFELO -- satu-satunya jalur tulis
  // yang sah adalah lewat gateway. Ketiga langkah itu sekarang dikerjain
  // SEKALIGUS di server (lihat aksiSyncPoMasuk_ di DatabaseGateway.js),
  // jadi cukup satu panggilan.
  const hasil = kirimAksiGateway_("syncPoMasuk", { jenis: "MASUK", poNo: poNo, rows: rows });

  if (rows.length > 0 && hasil.mulai_dari_baris) {
    // Kolom Q (ACTION) sengaja TIDAK ikut di-mirror (biar tetap bisa
    // diklik admin secara lokal) -- baris di database & di mirror lokal
    // SELALU sama posisinya (mirror 1:1 row-by-row), jadi nomor baris
    // yang dikembalikan gateway langsung dipakai buat nyiapin dropdown
    // ACTION di lokal, walau data hasil sync-nya sendiri baru kelihatan
    // di mirror beberapa saat lagi (delay IMPORTRANGE).
    applyHistoryPOActionValidation_(hasil.mulai_dari_baris, rows.length);
  }

  return pesanRenumber; // teks alert (atau null) -- ditampilkan syncToHistory() setelah lock dilepas
}

function applyHistoryPOActionValidation_(startRow, numRows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hist = ss.getSheetByName("HISTORY_PO_MASUK");

  if (!hist || numRows <= 0) return;

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["EDIT", "CANCEL"], true)
    .setAllowInvalid(false)
    .build();

  hist.getRange(startRow, HISTORY_PO_ACTION_COL, numRows, 1).setDataValidation(rule);
}

/**
 * Cek dampak cancel PO terhadap transaksi berjalan
 * Return:
 * {
 *   blocked: true/false,
 *   reason: [],
 *   detailRows: [],
 *   trxRows: []
 * }
 */
function cekDampakCancelPO_(poNo) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const trx = ss.getSheetByName("TRANSAKSI");
  const det = ss.getSheetByName("DETAIL_TRANSAKSI");

  if (!trx || !det) {
    throw new Error("Sheet TRANSAKSI / DETAIL_TRANSAKSI tidak ditemukan.");
  }


  const result = {
    blocked: false,
    reason: [],
    detailRows: [],
    trxRows: []
  };


  const lastDet = det.getLastRow();

  if (lastDet < 2) {
    return result;
  }


  // ⬅️ FIX: SOURCE_PO ada di kolom U (21), BUKAN S (19) -- kolom S sudah
  // jadi rumus LABA sejak lama (lihat DET_COL_SOURCE_PO di Transaction.js).
  // Dulu baca kolom S -> JSON.parse selalu gagal -> PO yang sudah dipakai
  // transaksi tidak pernah terdeteksi saat Cancel PO.
  // ⬅️ MIGRASI: dibaca FRESH dari database (A..U).
  const detailData = bacaDetailDB_();


  const noSJSet = {};


  detailData.forEach(function(row, index) {

    const sourcePO = row[20]; // U = SOURCE_PO

    if (!sourcePO) return;


    try {

      const sources = JSON.parse(sourcePO);


      const match = sources.some(function(src){
        return String(src.poNo).trim() === String(poNo).trim();
      });


      if (match) {

        const realRow = index + 2;

        result.detailRows.push(realRow);

        noSJSet[row[0]] = true; // A = NO SJ
      }


    } catch(err) {
      Logger.log(
        "SOURCE_PO gagal parse row " + (index+2)
      );
    }

  });



  const noSJList = Object.keys(noSJSet);


  if (noSJList.length === 0) {
    return result;
  }



  // ⬅️ MIGRASI: dibaca FRESH dari database (A..Q)
  const trxData = bacaTransaksiDB_();

  if (trxData.length === 0) {
    return result;
  }

  // ⬅️ FIX: status LUNAS diambil dari PENJUALAN (kolom F) per NO INV --
  // dulu dibaca dari TRANSAKSI kolom Q, padahal Q itu NO PO.
  const statusJualPerInv = {};
  bacaSheetGateway_("PENJUALAN", 2, 3, 4).forEach(function (r) { // C..F
    const k = String(r[0] || "").trim();
    if (k) statusJualPerInv[k] = String(r[3] || "").trim().toUpperCase();
  });



  trxData.forEach(function(row,index){

    const noSJ = row[1]; // B

    if (noSJList.indexOf(String(noSJ)) === -1) {
      return;
    }


    // ⬅️ FIX: STATUS ORDER ada di kolom N (index 13), bukan H (index 7,
    // isinya selalu "DRAFT").
    const statusOrder = String(row[13] || "")
      .trim()
      .toUpperCase();


    const statusJual = statusJualPerInv[String(row[2] || "").trim()] || ""; // C = NO INV



    if (
      statusOrder === "DONE" ||
      statusJual === "LUNAS"
    ) {

      result.blocked = true;

      result.reason.push(
        "SJ " + noSJ +
        " sudah DONE/LUNAS"
      );

    }


    result.trxRows.push(index + 2);


  });



  return result;

}

function warnaiCancelPO_(sheetName, rows){

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(sheetName);

  if(!sh || rows.length === 0){
    return;
  }

  const lastCol = sh.getLastColumn();

  // ⬅️ FIX (performa): dulu setBackground() dipanggil SATU-SATU per baris
  // di "rows" (bisa banyak baris kalau PO-nya kena di banyak baris
  // TRANSAKSI/DETAIL_TRANSAKSI). "rows" biasanya berisi baris yang
  // berurutan (blok item dari PO yang sama), jadi sekarang kita kelompokin
  // dulu jadi rentang-rentang baris yang nyambung (contiguous), baru tiap
  // rentang ditulis SEKALI lewat setBackgrounds() -- hasil akhirnya identik
  // (semua baris di "rows" tetap jadi merah), cuma jauh lebih sedikit
  // panggilan API dibanding 1 panggilan per baris.
  const sortedRows = rows.slice().sort(function(a, b) { return a - b; });

  let idx = 0;
  while (idx < sortedRows.length) {
    const startRow = sortedRows[idx];
    let endRow = startRow;
    let next = idx + 1;
    while (next < sortedRows.length && sortedRows[next] === endRow + 1) {
      endRow = sortedRows[next];
      next++;
    }

    const numRows = endRow - startRow + 1;
    const warnaGrid = [];
    for (let r = 0; r < numRows; r++) {
      const rowWarna = [];
      for (let c = 0; c < lastCol; c++) rowWarna.push("#ffcccc");
      warnaGrid.push(rowWarna);
    }
    sh.getRange(startRow, 1, numRows, lastCol).setBackgrounds(warnaGrid);

    idx = next;
  }

}

/* ============================================================
   HANDLER ACTIONS (EDIT / CANCEL) - AMAN DARI RUMUS
   ============================================================ */
function handleHistoryEdit(clickedRow) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hist = ss.getSheetByName("HISTORY_PO_MASUK");
  const masuk = ss.getSheetByName("PO_MASUK");

  if (!hist || !masuk) return;

  // ⬅️ BARU (Oktober 2026): apapun yang terjadi di bawah, form PO_MASUK
  // akan diisi dengan PO yang SUDAH TERSIMPAN (nomor sudah komit, bukan
  // preview) -- hapus flag preview yang mungkin masih nyangkut dari
  // percobaan "PO Masuk Baru" sebelumnya yang gak jadi disimpan, supaya
  // syncToHistory() nanti gak keliru nyangka nomor PO hasil EDIT ini masih
  // berstatus "belum diklaim".
  PropertiesService.getDocumentProperties().deleteProperty(PO_MASUK_PREVIEW_PROP_);

  const jenis = String(hist.getRange(clickedRow, 1).getValue() || "").trim().toUpperCase();
  const poNo = String(hist.getRange(clickedRow, 2).getValue() || "").trim();

  if (!poNo) {
    SpreadsheetApp.getUi().alert("Baris ini belum punya NO PO.");
    hist.getRange(clickedRow, 17).clearContent();
    return;
  }

  if (jenis !== "MASUK") {
    SpreadsheetApp.getUi().alert("EDIT dari HISTORY_PO_MASUK saat ini hanya mendukung JENIS = MASUK.");
    hist.getRange(clickedRow, 17).clearContent();
    return;
  }

  const statusClicked = String(hist.getRange(clickedRow, 12).getValue() || "").trim().toUpperCase();
  if (statusClicked === "CANCELLED") {
    SpreadsheetApp.getUi().alert(
      "PO " + poNo + " sudah di-CANCEL dan tidak bisa di-edit.\n" +
      "Gunakan ACTION \"BATAL CANCEL\" dulu untuk mengaktifkan kembali."
    );
    hist.getRange(clickedRow, 17).clearContent();
    return;
  }

  const lastRow = hist.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert("HISTORY_PO_MASUK masih kosong.");
    hist.getRange(clickedRow, 17).clearContent();
    return;
  }

  const data = hist.getRange(2, 1, lastRow - 1, 18).getValues();

  const items = [];
  let header = null;

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const rowJenis = String(row[0] || "").trim().toUpperCase();
    const rowPoNo = String(row[1] || "").trim();

    if (rowJenis === "MASUK" && rowPoNo === poNo) {
      if (!header) {
        header = {
          poNo: row[1],
          poTgl: row[2],
          customer: row[3],
          diskon: row[4],   // ⬅️ TAMBAHAN: E DISKON
          alamat: row[5],
          telp: row[6],
          tglKirim: row[7],
          method: row[17]
        };
      }

      items.push({
        line: row[15] || items.length + 1,
        kode: row[8],
        qty: row[10]
      });
    }
  }

  if (!header) {
    SpreadsheetApp.getUi().alert("Data PO " + poNo + " tidak ditemukan.");
    hist.getRange(clickedRow, 17).clearContent();
    return;
  }

  items.sort(function(a, b) {
    return (Number(a.line) || 0) - (Number(b.line) || 0);
  });

  // --- 1. SET HEADER ---
  masuk.getRange("B2").setValue(header.customer);
  masuk.getRange("B3").setValue(header.alamat);
  masuk.getRange("B4").setValue(header.telp);
  masuk.getRange("E2").setValue(header.poNo);
  masuk.getRange("E3").setValue(header.poTgl);
  masuk.getRange("E4").setValue(header.tglKirim);
  masuk.getRange("H3").setValue(header.diskon || ""); // ⬅️ TAMBAHAN: restore DISKON
  masuk.getRange("H4").setValue(header.method);

  // --- 2. BERSIHKAN HANYA KOLOM B DAN D ---
  masuk.getRange(PO_ITEM_START_ROW, 2, PO_ITEM_MAX_ROWS, 1).clearContent(); // Kode (B)
  masuk.getRange(PO_ITEM_START_ROW, 4, PO_ITEM_MAX_ROWS, 1).clearContent(); // Qty (D)

  // --- 3. ISI ITEM BARU (KODE & QTY SAJA) ---
  const kodes = items.map(function(it) { return [it.kode]; });
  const qtys = items.map(function(it) { return [it.qty]; });

  if (kodes.length > 0) {
    masuk.getRange(PO_ITEM_START_ROW, 2, kodes.length, 1).setValues(kodes); // Isi Kode (B)
    masuk.getRange(PO_ITEM_START_ROW, 4, qtys.length, 1).setValues(qtys);   // Isi Qty (D)
  }

  hist.getRange(clickedRow, 17).clearContent();

  ss.setActiveSheet(masuk);
  masuk.getRange("B2").activate();
}

function handleHistoryCancel(clickedRow) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hist = ss.getSheetByName("HISTORY_PO_MASUK");
  const ui = SpreadsheetApp.getUi();
  if (!hist) return;

  const poNo = String(hist.getRange(clickedRow, 2).getValue() || "").trim();
  const jenis = String(hist.getRange(clickedRow, 1).getValue() || "").trim().toUpperCase();

  const cek = cekDampakCancelPO_(poNo);


if(cek.blocked){

  ui.alert(
    "Cancel PO dibatalkan.\n\n" +
    cek.reason.join("\n") +
    "\n\nTransaksi DONE/LUNAS tidak boleh dicancel otomatis."
  );

  warnaiCancelPO_(
  "TRANSAKSI",
  cek.trxRows
);


warnaiCancelPO_(
  "DETAIL_TRANSAKSI",
  cek.detailRows
);


  hist.getRange(clickedRow,17).clearContent();

  return;

}

  if (!poNo) {
    ui.alert("Baris ini belum punya NO PO.");
    hist.getRange(clickedRow, 17).clearContent();
    return;
  }

  const resp = ui.alert(
    "Cancel PO?",
    "Semua baris item dengan NO PO \"" + poNo + "\" akan diubah STATUS-nya jadi CANCELLED. Lanjut?",
    ui.ButtonSet.YES_NO
  );

  if (resp !== ui.Button.YES) {
    hist.getRange(clickedRow, 17).clearContent();
    return;
  }

  // ⬅️ MIGRASI KE GATEWAY: dulu fungsi ini baca+tulis STATUS (L), backup
  // status (S), dan warna (A:R) LANGSUNG ke HISTORY_PO_MASUK lokal, scan
  // seluruh sheet. HISTORY_PO_MASUK di inventory_SANFELO sekarang cuma
  // mirror (IMPORTRANGE) -- data yang benar-benar dipegang cuma yang ada
  // di database_SANFELO, jadi status/backup/warna sekarang ditulis DI
  // SANA lewat gateway (lihat aksiCancelPo_ di DatabaseGateway.js, sudah
  // termasuk fix lastCol/kolom-S). Yang TETAP dikerjain lokal cuma kolom
  // Q (ACTION), karena kolom itu sengaja TIDAK ikut di-mirror.
  const hasil = kirimAksiGateway_("cancelPo", { jenis: jenis, poNo: poNo });

  // Kolom Q (ACTION) lokal: cari baris DI MIRROR yang cocok JENIS+NO PO --
  // aman discan lokal karena baris-baris PO ini pasti sudah kelihatan di
  // mirror (admin baru aja klik CANCEL di salah satunya), gak perlu nunggu
  // status barunya kelihatan dulu (JENIS+NO PO gak berubah pas cancel).
  const lastRow = hist.getLastRow();
  if (lastRow >= 2) {
    const dataAB = hist.getRange(2, 1, lastRow - 1, 2).getValues(); // A, B
    const rowsMatch = [];
    for (let i = 0; i < dataAB.length; i++) {
      const rowJenis = String(dataAB[i][0] || "").trim().toUpperCase();
      const rowPoNo = String(dataAB[i][1] || "").trim();
      if (rowJenis === jenis && rowPoNo === poNo) rowsMatch.push(i + 2);
    }

    if (rowsMatch.length > 0) {
      const ruleBatalCancel = SpreadsheetApp.newDataValidation()
        .requireValueInList(["BATAL CANCEL"], true)
        .setAllowInvalid(false)
        .build();
      const minRow = Math.min.apply(null, rowsMatch);
      const maxRow = Math.max.apply(null, rowsMatch);

      if (maxRow - minRow + 1 === rowsMatch.length) {
        // baris PO ini selalu ditulis bareng pas sync -- normalnya nyambung
        hist.getRange(minRow, HISTORY_PO_ACTION_COL, rowsMatch.length, 1).clearContent();
        hist.getRange(minRow, HISTORY_PO_ACTION_COL, rowsMatch.length, 1).setDataValidation(ruleBatalCancel);
      } else {
        rowsMatch.forEach(function (r) {
          hist.getRange(r, HISTORY_PO_ACTION_COL, 1, 1).clearContent();
          hist.getRange(r, HISTORY_PO_ACTION_COL, 1, 1).setDataValidation(ruleBatalCancel);
        });
      }
    }
  }

  ui.alert(hasil.updatedCount + " baris PO " + poNo + " berhasil di-cancel.");
}

function handleHistoryUncancel(clickedRow) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hist = ss.getSheetByName("HISTORY_PO_MASUK");
  const ui = SpreadsheetApp.getUi();
  if (!hist) return;

  const poNo = String(hist.getRange(clickedRow, 2).getValue() || "").trim();
  const jenis = String(hist.getRange(clickedRow, 1).getValue() || "").trim().toUpperCase();

  if (!poNo) {
    ui.alert("Baris ini belum punya NO PO.");
    hist.getRange(clickedRow, 17).clearContent();
    return;
  }

  const statusClicked = String(hist.getRange(clickedRow, 12).getValue() || "").trim().toUpperCase();
  if (statusClicked !== "CANCELLED") {
    ui.alert("PO " + poNo + " tidak dalam status CANCELLED, tidak perlu di-batal-cancel.");
    hist.getRange(clickedRow, 17).clearContent();
    return;
  }

  const resp = ui.alert(
    "Batal Cancel PO?",
    "Status PO \"" + poNo + "\" akan dikembalikan ke status sebelum di-cancel. Lanjut?",
    ui.ButtonSet.YES_NO
  );

  if (resp !== ui.Button.YES) {
    hist.getRange(clickedRow, 17).clearContent();
    return;
  }

  // ⬅️ MIGRASI KE GATEWAY: sama alasannya kayak handleHistoryCancel -- status
  // restore (L), backup dikosongin (S), dan warna (A:R) sekarang ditulis di
  // database_SANFELO lewat gateway (lihat aksiUncancelPo_). Kolom Q (ACTION)
  // tetap dikerjain lokal karena sengaja TIDAK ikut di-mirror. Gateway juga
  // yang tau status hasil restore-nya apa (buat nentuin dropdown Q jadi
  // EDIT/CANCEL atau BATAL CANCEL kalau ternyata backup-nya CANCELLED juga),
  // jadi kita baca balik lewat bacaSheet biar dropdown-nya akurat.
  const hasil = kirimAksiGateway_("uncancelPo", { jenis: jenis, poNo: poNo });

  const lastRow = hist.getLastRow();
  if (lastRow >= 2) {
    const dataAB = hist.getRange(2, 1, lastRow - 1, 2).getValues(); // A, B
    const rowsMatch = [];
    for (let i = 0; i < dataAB.length; i++) {
      const rowJenis = String(dataAB[i][0] || "").trim().toUpperCase();
      const rowPoNo = String(dataAB[i][1] || "").trim();
      if (rowJenis === jenis && rowPoNo === poNo) rowsMatch.push(i + 2);
    }

    if (rowsMatch.length > 0) {
      // Baca status HASIL RESTORE langsung dari database (bukan mirror yang
      // bisa telat) buat nentuin opsi dropdown yang benar per baris.
      const statusFresh = bacaSheetGateway_("HISTORY_PO_MASUK", 2, 12, 1); // kolom L = STATUS
      const minRow = Math.min.apply(null, rowsMatch);
      const maxRow = Math.max.apply(null, rowsMatch);
      const kontigu = (maxRow - minRow + 1 === rowsMatch.length);

      rowsMatch.forEach(function (r) {
        const statusRow = statusFresh[r - 2] ? String(statusFresh[r - 2][0] || "").trim().toUpperCase() : "";
        const opsi = statusRow === "CANCELLED" ? ["BATAL CANCEL"] : ["EDIT", "CANCEL"];
        const rule = SpreadsheetApp.newDataValidation()
          .requireValueInList(opsi, true)
          .setAllowInvalid(false)
          .build();
        hist.getRange(r, HISTORY_PO_ACTION_COL, 1, 1).clearContent();
        hist.getRange(r, HISTORY_PO_ACTION_COL, 1, 1).setDataValidation(rule);
      });
    }
  }

  ui.alert(hasil.updatedCount + " baris PO " + poNo + " berhasil dikembalikan (batal cancel).");
}

/* ------------------------------------------------------
   5) SURAT JALAN — load data dari HISTORY_PO_MASUK (SUPPORT MULTI-PO)
   ------------------------------------------------------ */
const SJ_ITEM_START_ROW = 9;
const SJ_ITEM_MAX_ROWS  = 8;
const SJ_DISPLAY_START_ROW = 9;
const SJ_DISPLAY_END_ROW   = 16; // fallback kalau anchor gak ketemu

const INV_DISPLAY_START_ROW = 11;
const INV_DISPLAY_END_ROW   = 18; // fallback kalau anchor gak ketemu
const INV_ITEM_MAX_ROWS = 8;

const SJ_FOOTER_COL_START  = 1;   // A
const SJ_FOOTER_COL_END    = 4;   // sampai D (jaga-jaga kalau merge cell mulai dari kolom lain)
const SJ_FOOTER_TEKS       = "Yang Menyerahkan,";

const INV_FOOTER_COL_START = 1;   // A
const INV_FOOTER_COL_END   = 3;   // sampai C
const INV_FOOTER_TEKS      = "Accounting,";
const INV_KETERANGAN_ANCHOR_TEKS = "Keterangan:";

// Cari baris pertama yang MENGANDUNG teks target di kolom manapun (aman dari merged cell)
function cariBarisAnchor_(sheet, teksMengandung, startRow) {
  const lastRow = sheet.getLastRow();
  const lastCol = Math.max(sheet.getLastColumn(), 9);

  for (let r = startRow; r <= lastRow + 5; r++) {
    const rowValues = sheet.getRange(r, 1, 1, lastCol).getValues()[0];
    for (let c = 0; c < rowValues.length; c++) {
      const val = String(rowValues[c] || "").toUpperCase();
      if (val.indexOf(teksMengandung.toUpperCase()) !== -1) return r;
    }
  }
  return -1;
}

// Pastikan area item punya "neededRows" baris.
// Auto insert kalau kurang, auto hapus kalau lebih (collapse balik ke ukuran dasar).
function pastikanKapasitasItem_(sheet, itemStartRow, baseRows, neededRows, footerColStart, footerColEnd, footerTeks) {
  let footerRow = cariBarisAnchor_(sheet, footerTeks, itemStartRow);
  if (footerRow === -1) {
    SpreadsheetApp.getUi().alert(
      'Anchor footer "' + footerTeks + '" tidak ditemukan di sheet "' + sheet.getName() + '".'
    );
    throw new Error('Anchor "' + footerTeks + '" tidak ditemukan di "' + sheet.getName() + '".');
  }

  const currentCapacity = footerRow - itemStartRow;
  const target = Math.max(baseRows, neededRows);

  if (target === currentCapacity) return; // gak ada perubahan, skip semua

  const oldLastRow = footerRow - 1; // ✅ dicatat SEBELUM insert/hapus

  if (target > currentCapacity) {
    const tambah = target - currentCapacity;
    const lastItemRow = footerRow - 1;
    const numCols = sheet.getLastColumn();

    sheet.insertRowsBefore(footerRow, tambah);

    const sumber = sheet.getRange(lastItemRow, 1, 1, numCols);
    const tujuan = sheet.getRange(footerRow, 1, tambah, numCols);
    sumber.copyTo(tujuan, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);

  } else {
    const kurang = currentCapacity - target;
    sheet.deleteRows(footerRow - kurang, kurang);
  }

  const newLastRow = itemStartRow + target - 1; // ✅ posisi baris terakhir yang baru
  perbaruiRangeArrayFormula_(sheet, itemStartRow, oldLastRow, newLastRow); // ✅ rewrite formula H11/I11
}

function loadSuratJalanFromHistory() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const sj   = ss.getSheetByName("SURAT JALAN");
  const hist = ss.getSheetByName("HISTORY_PO_MASUK");
  const inv  = ss.getSheetByName("INVOICE");

  if (!sj || !hist) return;

  const rawH3 = String(sj.getRange("H3").getValue() || "").trim();

  const baseRowsSJ  = getEffectiveBaseRows_(SJ_ITEM_MAX_ROWS);
  const baseRowsInv = getEffectiveBaseRows_(INV_ITEM_MAX_ROWS);

  pastikanKapasitasItem_(sj, SJ_ITEM_START_ROW, baseRowsSJ, baseRowsSJ, SJ_FOOTER_COL_START, SJ_FOOTER_COL_END, SJ_FOOTER_TEKS);
  if (inv) pastikanKapasitasItem_(inv, INV_DISPLAY_START_ROW, baseRowsInv, baseRowsInv, INV_FOOTER_COL_START, INV_FOOTER_COL_END, INV_FOOTER_TEKS);

  sj.getRange("A5").clearContent();
  sj.getRange("A6").clearContent();
  sj.getRange(SJ_ITEM_START_ROW, 2, baseRowsSJ, 1).clearContent(); // B kode
  sj.getRange(SJ_ITEM_START_ROW, 5, baseRowsSJ, 1).clearContent(); // E qty
  sj.getRange(SJ_ITEM_START_ROW, 6, baseRowsSJ, 1).clearContent(); // F satuan
  sj.getRange(SJ_ITEM_START_ROW, 7, baseRowsSJ, 1).clearContent(); // G ket
  sj.getRange(SJ_ITEM_START_ROW, 9, baseRowsSJ, 2).clearContent(); // I,J mapping hidden

  if (!rawH3) return;

  const poNoList = rawH3.split(",").map(function(s) { return s.trim(); }).filter(Boolean);

  const lastRow = hist.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert("HISTORY_PO_MASUK masih kosong.");
    return;
  }

  const data = hist.getRange(2, 1, lastRow - 1, 18).getValues();

  let headerSet = false;
  let namaTokoHeader = "";
  let alamatHeader = "";
  let kirimTglHeader = null;
  let methodHeader = "";
  const combinedItems = [];
  const poGagal = [];

  poNoList.forEach(function(poNo) {
    const matches = data.filter(function(r) {
      return String(r[0] || "").trim().toUpperCase() === "MASUK" &&
             String(r[1] || "").trim() === poNo;
    });

    if (matches.length === 0) {
      poGagal.push(poNo + " (tidak ditemukan)");
      return;
    }

    const statusPO = String(matches[0][11] || "").trim().toUpperCase();
    if (statusPO === "CANCELLED") {
      poGagal.push(poNo + " (CANCELLED)");
      return;
    }

    if (!headerSet) {
      namaTokoHeader = matches[0][3];
      alamatHeader   = matches[0][5];
      kirimTglHeader = matches[0][7];
      methodHeader   = matches[0][17];
      headerSet = true;
    }

    matches.forEach(function(r) {
      const qtyAsli = Number(r[10]) || 0;
      const statusRaw = String(r[11] || "").trim().toUpperCase();
      const qtyDoneSoFar = parseDoneQty_(statusRaw, qtyAsli);
      const sisaQty = Math.max(0, qtyAsli - qtyDoneSoFar);

      combinedItems.push({
        poNo: poNo,
        line: r[15],
        kode: r[8],
        nama: r[9],
        qtyAsli: qtyAsli,
        sisaQty: sisaQty,
        qty: sisaQty,
        status: String(r[11] || "").trim().toUpperCase(),
        harga: Number(r[12]) || 0,
        total: Number(r[13]) || 0
      });
    });
  });

  if (poGagal.length > 0) {
    SpreadsheetApp.getUi().alert("PO berikut dilewati:\n" + poGagal.join("\n"));
  }

  if (combinedItems.length === 0) {
    SpreadsheetApp.getUi().alert('Tidak ada barang yang bisa dimuat dari No PO: "' + rawH3 + '"');
    return;
  }

  // simpan header sementara, dipakai lagi saat user klik "Simpan" di popup
  PropertiesService.getDocumentProperties().setProperty("SJ_POPUP_HEADER", JSON.stringify({
    namaToko: namaTokoHeader,
    alamat: alamatHeader,
    tglKirim: kirimTglHeader ? new Date(kirimTglHeader).toISOString() : null,
    method: methodHeader
  }));

  tampilkanPopupPilihBarangSJ_(combinedItems);
}

/* ------------------------------------------------------
   POPUP PILIH BARANG SJ (baru)
   ------------------------------------------------------ */
function tampilkanPopupPilihBarangSJ_(items) {
  const template = HtmlService.createTemplateFromFile("PopupPilihBarangSJ");
  template.itemsJson = JSON.stringify(items);
  const html = template.evaluate().setWidth(480).setHeight(560);
  SpreadsheetApp.getUi().showModalDialog(html, "Pilih Barang untuk Surat Jalan");
}

// Dipanggil dari client (google.script.run) setelah user klik "Simpan" di popup
function terapkanPilihanSJPopup(selectedJson) {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const sj   = ss.getSheetByName("SURAT JALAN");
  const inv  = ss.getSheetByName("INVOICE");
  const hist = ss.getSheetByName("HISTORY_PO_MASUK");
  if (!sj) return;

  const selected = JSON.parse(selectedJson || "[]");

  // ✅ Build lookup fresh dari HISTORY_PO_MASUK (anti data basi)
  const histLastRow = hist ? hist.getLastRow() : 0;
  const histMap = {};
  if (hist && histLastRow >= 2) {
    const histData = hist.getRange(2, 1, histLastRow - 1, 16).getValues();
    histData.forEach(function(row) {
      const jenis = String(row[0] || "").trim().toUpperCase();
      if (jenis !== "MASUK") return;
      const poNo = String(row[1] || "").trim();
      const line = String(row[15] || "").trim();
      histMap[poNo + "||" + line] = {
        qtyAsli:   Number(row[10]) || 0, // K QTY
        hargaUnit: Number(row[12]) || 0, // M HARGA_UNIT
        total:     Number(row[13]) || 0  // N TOTAL
      };
    });
  }

  // ✅ MERGE: gabungkan item dengan KODE sama (walaupun beda PO) jadi satu baris di SJ
  const mergedMap = {};
  const mergedOrder = [];

  selected.forEach(function(it) {
    const kode = String(it.kode || "").trim();
    if (!mergedMap[kode]) {
      mergedMap[kode] = { kode: kode, qty: 0, sources: [] };
      mergedOrder.push(kode);
    }
    mergedMap[kode].qty += Number(it.qty) || 0;
    mergedMap[kode].sources.push({
      poNo: it.poNo,
      line: it.line,
      qty: Number(it.qty) || 0
    });
  });

  const merged = mergedOrder.map(function(k) { return mergedMap[k]; });

  const headerRaw = PropertiesService.getDocumentProperties().getProperty("SJ_POPUP_HEADER");
  const header = headerRaw ? JSON.parse(headerRaw) : {};

  const baseRowsSJ  = getEffectiveBaseRows_(SJ_ITEM_MAX_ROWS);
  const baseRowsInv = getEffectiveBaseRows_(INV_ITEM_MAX_ROWS);
  const neededRows  = Math.max(baseRowsSJ, merged.length);

  pastikanKapasitasItem_(sj, SJ_ITEM_START_ROW, baseRowsSJ, neededRows, SJ_FOOTER_COL_START, SJ_FOOTER_COL_END, SJ_FOOTER_TEKS);
  if (inv) pastikanKapasitasItem_(inv, INV_DISPLAY_START_ROW, baseRowsInv, neededRows, INV_FOOTER_COL_START, INV_FOOTER_COL_END, INV_FOOTER_TEKS);

  sj.getRange("A5").setValue(header.namaToko || "");
  sj.getRange("A6").setValue(header.alamat || "");
  if (header.tglKirim) {
    const tgl = new Date(header.tglKirim);
    setTanggalSaja_(sj.getRange("H5"), tgl);
    if (inv) setTanggalSaja_(inv.getRange("H7"), tgl);
  }
  if (inv && header.method) {
    inv.getRange("I3").setValue(header.method);
  }

  sj.getRange(SJ_ITEM_START_ROW, 2, neededRows, 1).clearContent();
  sj.getRange(SJ_ITEM_START_ROW, 5, neededRows, 1).clearContent();
  sj.getRange(SJ_ITEM_START_ROW, 6, neededRows, 1).clearContent();
  sj.getRange(SJ_ITEM_START_ROW, 7, neededRows, 1).clearContent();
  sj.getRange(SJ_ITEM_START_ROW, 9, neededRows, 2).clearContent();

  // ⬅️ FIX (performa): dulu tiap item hasil merge nulis KODE, QTY, sumber
  // JSON (SURAT JALAN) + HARGA/DISKON/TOTAL (INVOICE) SATU-SATU DI DALAM
  // LOOP -- sampai 7 panggilan API per baris. Sekarang semua hasil
  // dikumpulin dulu ke array per kolom, baru ditulis SEKALI per kolom
  // lewat setValues() -- hasil akhir di kedua sheet identik persis, cuma
  // jauh lebih sedikit panggilan API. (Kolom J udah dikosongin bareng
  // clearContent() di atas, jadi gak perlu di-clear lagi satu-satu di sini.)
  const kodeGrid = [];
  const qtyGrid = [];
  const sourcesGrid = [];
  const invHargaGrid = [];
  const invDiskonGrid = [];
  const invTotalGrid = [];

  merged.forEach(function(item) {
    kodeGrid.push([item.kode]);
    qtyGrid.push([item.qty]);
    // Kolom I = JSON semua sumber PO+line (bisa lebih dari satu kalau digabung)
    sourcesGrid.push([JSON.stringify(item.sources)]);

    if (inv) {
      let sumHarga = 0;
      let sumTotal = 0;

      item.sources.forEach(function(src) {
        const key = String(src.poNo).trim() + "||" + String(src.line).trim();
        const fresh = histMap[key];

        if (fresh && fresh.qtyAsli > 0) {
          const ratio = src.qty / fresh.qtyAsli;
          sumHarga += Math.round(fresh.hargaUnit * ratio);
          sumTotal += Math.round(fresh.total * ratio);
        }
      });

      const diskon = sumHarga - sumTotal;

      invHargaGrid.push([sumHarga]);
      invDiskonGrid.push([diskon]);
      invTotalGrid.push([sumTotal]);
    }
  });

  if (kodeGrid.length > 0) {
    sj.getRange(SJ_ITEM_START_ROW, 2, kodeGrid.length, 1).setValues(kodeGrid);
    sj.getRange(SJ_ITEM_START_ROW, 5, qtyGrid.length, 1).setValues(qtyGrid);
    sj.getRange(SJ_ITEM_START_ROW, 9, sourcesGrid.length, 1).setValues(sourcesGrid);

    if (inv && invHargaGrid.length > 0) {
      inv.getRange(INV_DISPLAY_START_ROW, 7, invHargaGrid.length, 1).setValues(invHargaGrid);
      inv.getRange(INV_DISPLAY_START_ROW, 8, invDiskonGrid.length, 1).setValues(invDiskonGrid);
      inv.getRange(INV_DISPLAY_START_ROW, 9, invTotalGrid.length, 1).setValues(invTotalGrid);
    }
  }

  sembunyikanKolomMappingSJ_(sj);
  adjustSJItemBorders();
  adjustInvoiceItemBorders();

  PropertiesService.getDocumentProperties().deleteProperty("SJ_POPUP_HEADER");
}

// ✅ BARU: baca balik mapping sumber (PO+line+qty) dari kolom I sheet SURAT JALAN
// Panggil ini dari fungsi finalize/cetak SJ kamu, ganti cara lama yang baca
// kolom I/J satu-satu per baris, lalu hasilnya di-pass ke updateStatusDoneSJHistoryPO_()
function bacaMappingSumberSJDariSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sj = ss.getSheetByName("SURAT JALAN");
  if (!sj) return [];

  const footerRow = cariBarisAnchor_(sj, SJ_FOOTER_TEKS, SJ_ITEM_START_ROW);
  const endRow = footerRow === -1 ? SJ_DISPLAY_END_ROW : footerRow - 1;

  const mappingList = [];

  for (let r = SJ_ITEM_START_ROW; r <= endRow; r++) {
    const kode = sj.getRange(r, 2).getValue();
    if (!kode) continue;

    const rawMapping = sj.getRange(r, 9).getValue();
    if (!rawMapping) continue;

    try {
      const sources = JSON.parse(rawMapping);
      sources.forEach(function(src) {
        mappingList.push({ poNo: src.poNo, line: src.line, qty: src.qty });
      });
    } catch (err) {
      Logger.log("Gagal parse mapping SJ baris " + r + ": " + err);
    }
  }

  return mappingList;
}

function sembunyikanKolomMappingSJ_(sj) {
  try { sj.hideColumns(9, 2); } catch (err) { Logger.log("Hide kolom mapping SJ error: " + err); }
}

/* ------------------------------------------------------
   6) EXPORT SURAT JALAN KE PDF
   ------------------------------------------------------ */
function exportSJPDF() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sj = ss.getSheetByName("SURAT JALAN");
  if (!sj) return;

  const noSJ     = sj.getRange("H6").getValue() || "SJ";
  const customer = sj.getRange("A5").getValue() || "";
  const originalH3 = sj.getRange("H3").getValue();
  sj.getRange("H3").clearContent();

  const footerRow = cariBarisAnchor_(sj, SJ_FOOTER_TEKS, SJ_ITEM_START_ROW);
  const lastItemRow = footerRow === -1 ? (SJ_ITEM_START_ROW + SJ_ITEM_MAX_ROWS - 1) : footerRow - 1;

  // ⬅️ FIX: signatureRow sekarang nyampe baris "(.....)", bukan cuma 1 baris kosong.
  // Blok TTD = footerRow ("Yang Menyerahkan,") + 3 baris tambahan (kosong, kosong, "(.....)")
  const signatureRow = footerRow === -1 ? lastItemRow + 4 : footerRow + 3;

  const totalRowsCapacity = lastItemRow - SJ_ITEM_START_ROW + 1;
  // const jumlahHalaman = Math.max(1, Math.ceil(totalRowsCapacity / KERTAS_ROWS_PER_HALAMAN));
  const jumlahHalaman = 1;

  const colorRangeFull = sj.getRange(1, 1, signatureRow, 8);
  const originalColors = colorRangeFull.getBackgrounds();
  colorRangeFull.setBackground("#ffffff");
  SpreadsheetApp.flush();

  const fileLinks = [];
  const pageErrors = [];

  try {
    const sheetId = sj.getSheetId();
    const spreadsheetId = ss.getId();
    const token = ScriptApp.getOAuthToken();
    const folder = DriveApp.getFolderById("1Lp_RjQrCOLkvXMI4uW_Wr3O0MWJSv3yU");

    for (let h = 0; h < jumlahHalaman; h++) {
      try {
        const isFirstPage = (h === 0);
        const isLastPage  = (h === jumlahHalaman - 1);

        const chunkStart = SJ_ITEM_START_ROW + (h * KERTAS_ROWS_PER_HALAMAN);
        const chunkEnd   = Math.min(chunkStart + KERTAS_ROWS_PER_HALAMAN - 1, lastItemRow);

        const rowAwal  = isFirstPage ? 1 : chunkStart;
        const rowAkhir = isLastPage ? signatureRow : chunkEnd;

        const r1 = rowAwal - 1;
        const c1 = 0;
        const r2 = rowAkhir;
        const c2 = 8;

        const url = "https://docs.google.com/spreadsheets/d/" + spreadsheetId +
          "/export?format=pdf&gid=" + sheetId + "&size=a4&portrait=true&fitw=true" +
          "&gridlines=false&printtitle=false&sheetnames=false" +
          "&top_margin=0.2&bottom_margin=0.2&left_margin=0.2&right_margin=0.2" +
          "&horizontal_alignment=CENTER&r1=" + r1 + "&c1=" + c1 + "&r2=" + r2 + "&c2=" + c2;

        const response = UrlFetchApp.fetch(url, {
          headers: { Authorization: "Bearer " + token },
          muteHttpExceptions: true
        });

        const responseCode = response.getResponseCode();
        if (responseCode !== 200) {
          throw new Error("Export gagal (HTTP " + responseCode + "): " + response.getContentText().substring(0, 200));
        }

        let fileName = String(noSJ) + " - " + String(customer);
        if (jumlahHalaman > 1) fileName += " - hal" + (h + 1);
        fileName = fileName.replace(/[\\\/:*?"<>|]/g, "").trim() + ".pdf";

        const blob = response.getBlob().setName(fileName);
        const file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

        fileLinks.push({ name: fileName, url: "https://drive.google.com/uc?export=download&id=" + file.getId() });

        Utilities.sleep(300);

      } catch (pageErr) {
        pageErrors.push("Halaman " + (h + 1) + ": " + pageErr.message);
        Logger.log("Error export SJ halaman " + (h + 1) + ": " + pageErr);
      }
    }

    const linksHtml = fileLinks.map(function (f) {
      return '<p><a href="' + f.url + '" target="_blank">⬇️ Download ' + f.name + '</a></p>';
    }).join("");

    const errorHtml = pageErrors.length > 0
      ? '<div style="color:red;margin-top:10px;"><p><b>⚠️ Ada halaman gagal:</b></p><p>' + pageErrors.join("<br>") + '</p></div>'
      : '';

    const html = HtmlService.createHtmlOutput(
      '<div style="font-family:sans-serif;padding:10px;">' +
      '<p>PDF berhasil dibuat (' + fileLinks.length + ' dari ' + jumlahHalaman + ' halaman):</p>' +
      linksHtml + errorHtml + '</div>'
    ).setWidth(400).setHeight(150 + fileLinks.length * 30 + (pageErrors.length > 0 ? 80 : 0));

    SpreadsheetApp.getUi().showModalDialog(html, "Export Surat Jalan");

  } finally {
    sj.getRange("H3").setValue(originalH3);
    colorRangeFull.setBackgrounds(originalColors);
    SpreadsheetApp.flush();
  }
}

// /* ------------------------------------------------------
//    7) Refresh manual Surat Jalan
//    ------------------------------------------------------ */
// function refreshSJDariPO() {
//   loadSuratJalanFromHistory();
//   adjustSJItemBorders();
//   adjustInvoiceItemBorders();
//   SpreadsheetApp.getUi().alert("Data Surat Jalan sudah di-refresh.");
// }

/* ------------------------------------------------------
   8) Border dinamis SURAT JALAN
   ------------------------------------------------------ */
function adjustSJItemBorders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sj = ss.getSheetByName("SURAT JALAN");
  if (!sj) return;

  // ✅ signature baru: 5 argumen (sheet, colStart, colEnd, teks, startRow)
  const footerRow = cariBarisAnchor_(sj, SJ_FOOTER_TEKS, SJ_ITEM_START_ROW);
  const endRow = footerRow === -1 ? SJ_DISPLAY_END_ROW : footerRow - 1;

  let lastDataRow = -1;
  for (let r = SJ_ITEM_START_ROW; r <= endRow; r++) {
    const kode = sj.getRange(r, 2).getValue();
    const range = sj.getRange(r, 1, 1, 8);
    if (kode) { range.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID); lastDataRow = r; }
    else { range.setBorder(false, false, false, false, false, false, null, null); }
  }
  if (lastDataRow >= SJ_ITEM_START_ROW) sj.getRange(lastDataRow, 1, 1, 8).setBorder(null, null, true, null, null, null, "black", SpreadsheetApp.BorderStyle.SOLID);
  sj.getRange(endRow + 1, 1, 1, 8).setBorder(true, null, null, null, null, null, "black", SpreadsheetApp.BorderStyle.SOLID);
}

/* ------------------------------------------------------
   9) Border dinamis INVOICE
   ------------------------------------------------------ */
function adjustInvoiceItemBorders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inv = ss.getSheetByName("INVOICE");
  if (!inv) return;

  // ✅ signature baru: 5 argumen
  const footerRow = cariBarisAnchor_(inv, INV_FOOTER_TEKS, INV_DISPLAY_START_ROW);
  const endRow = footerRow === -1 ? INV_DISPLAY_END_ROW : footerRow - 1;

  let lastDataRow = -1;
  for (let r = INV_DISPLAY_START_ROW; r <= endRow; r++) {
    const barang = inv.getRange(r, 2).getValue();
    const range = inv.getRange(r, 1, 1, 9);
    if (barang) { range.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID); lastDataRow = r; }
    else { range.setBorder(false, false, false, false, false, false, null, null); }
  }
  if (lastDataRow >= INV_DISPLAY_START_ROW) inv.getRange(lastDataRow, 1, 1, 9).setBorder(null, null, true, null, null, null, "black", SpreadsheetApp.BorderStyle.SOLID);
  inv.getRange(endRow + 1, 1, 1, 9).setBorder(true, null, null, null, null, null, "black", SpreadsheetApp.BorderStyle.SOLID);
}

/* ------------------------------------------------------
   10) EXPORT INVOICE KE PDF
   ------------------------------------------------------ */
function exportInvoicePDF() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inv = ss.getSheetByName("INVOICE");
  if (!inv) return;

  const noInv    = inv.getRange("H6").getValue() || "INV";
  const customer = inv.getRange("A6").getValue() || "";

  const footerRow = cariBarisAnchor_(inv, INV_FOOTER_TEKS, INV_DISPLAY_START_ROW);
  const lastItemRow = footerRow === -1 ? (INV_DISPLAY_START_ROW + INV_ITEM_MAX_ROWS - 1) : footerRow - 1;
  const footerEndRow = footerRow === -1 ? lastItemRow + 5 : footerRow + 4; // sampai "Yang harus Dibayar"

  const totalRowsCapacity = lastItemRow - INV_DISPLAY_START_ROW + 1;
  // const jumlahHalaman = Math.max(1, Math.ceil(totalRowsCapacity / KERTAS_ROWS_PER_HALAMAN));
  const jumlahHalaman = 1;

  const colorRangeFull = inv.getRange(1, 1, footerEndRow, 9);
  const originalColors = colorRangeFull.getBackgrounds();
  colorRangeFull.setBackground("#ffffff");
  SpreadsheetApp.flush();

  const fileLinks = [];
  const pageErrors = []; // ⬅️ TAMBAHAN

  try {
    const sheetId = inv.getSheetId();
    const spreadsheetId = ss.getId();
    const token = ScriptApp.getOAuthToken();
    const folder = DriveApp.getFolderById("1GBfWEBu2awzzbwtsIh9nEiNghASxCXkX");

    for (let h = 0; h < jumlahHalaman; h++) {
      try { // ⬅️ TAMBAHAN
        const isFirstPage = (h === 0);
        const isLastPage  = (h === jumlahHalaman - 1);

        const chunkStart = INV_DISPLAY_START_ROW + (h * KERTAS_ROWS_PER_HALAMAN);
        const chunkEnd   = Math.min(chunkStart + KERTAS_ROWS_PER_HALAMAN - 1, lastItemRow);

        const rowAwal  = isFirstPage ? 1 : chunkStart;
        const rowAkhir = isLastPage ? footerEndRow : chunkEnd;

        const r1 = rowAwal - 1;
        const c1 = 0;
        const r2 = rowAkhir;
        const c2 = 9;

        const url = "https://docs.google.com/spreadsheets/d/" + spreadsheetId +
          "/export?format=pdf&gid=" + sheetId + "&size=a4&portrait=true&fitw=true" +
          "&gridlines=false&printtitle=false&sheetnames=false" +
          "&top_margin=0.2&bottom_margin=0.2&left_margin=0.2&right_margin=0.2" +
          "&horizontal_alignment=CENTER&r1=" + r1 + "&c1=" + c1 + "&r2=" + r2 + "&c2=" + c2;

        const response = UrlFetchApp.fetch(url, {
          headers: { Authorization: "Bearer " + token },
          muteHttpExceptions: true // ⬅️ TAMBAHAN
        });

        const responseCode = response.getResponseCode();
        if (responseCode !== 200) {
          throw new Error("Export gagal (HTTP " + responseCode + "): " + response.getContentText().substring(0, 200));
        }

        let fileName = String(noInv) + " - " + String(customer);
        if (jumlahHalaman > 1) fileName += " - hal" + (h + 1);
        fileName = fileName.replace(/[\\\/:*?"<>|]/g, "").trim() + ".pdf";

        const blob = response.getBlob().setName(fileName);
        const file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

        fileLinks.push({ name: fileName, url: "https://drive.google.com/uc?export=download&id=" + file.getId() });

        Utilities.sleep(300); // ⬅️ TAMBAHAN

      } catch (pageErr) { // ⬅️ TAMBAHAN
        pageErrors.push("Halaman " + (h + 1) + ": " + pageErr.message);
        Logger.log("Error export Invoice halaman " + (h + 1) + ": " + pageErr);
      }
    }

    const linksHtml = fileLinks.map(function (f) {
      return '<p><a href="' + f.url + '" target="_blank">⬇️ Download ' + f.name + '</a></p>';
    }).join("");

    const errorHtml = pageErrors.length > 0 // ⬅️ TAMBAHAN
      ? '<div style="color:red;margin-top:10px;"><p><b>⚠️ Ada halaman gagal:</b></p><p>' + pageErrors.join("<br>") + '</p></div>'
      : '';

    const html = HtmlService.createHtmlOutput(
      '<div style="font-family:sans-serif;padding:10px;">' +
      '<p>PDF berhasil dibuat (' + fileLinks.length + ' dari ' + jumlahHalaman + ' halaman):</p>' +
      linksHtml + errorHtml + '</div>'
    ).setWidth(400).setHeight(150 + fileLinks.length * 30 + (pageErrors.length > 0 ? 80 : 0));

    SpreadsheetApp.getUi().showModalDialog(html, "Export Invoice");
  } finally {
    colorRangeFull.setBackgrounds(originalColors);
    SpreadsheetApp.flush();
  }
}

function setHistoryPOActionOptionsForRow_(rowIdx) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hist = ss.getSheetByName("HISTORY_PO_MASUK");
  if (!hist) return;

  const status = String(hist.getRange(rowIdx, 12).getValue() || "").trim().toUpperCase();

  const options = (status === "CANCELLED")
    ? ["BATAL CANCEL"]
    : ["EDIT", "CANCEL"];

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(options, true)
    .setAllowInvalid(false)
    .build();

  hist.getRange(rowIdx, HISTORY_PO_ACTION_COL, 1, 1).setDataValidation(rule);
}

// ⬅️ BARU (penting buat mirror): kolom Q (ACTION) sengaja TIDAK ikut
// di-mirror lewat IMPORTRANGE (biar tetap bisa diklik admin secara lokal).
// Konsekuensinya: kalau ADA PO LAIN yang barisnya bergeser di database
// (misal PO lain di-edit ulang, baris lamanya kehapus lalu baris barunya
// nambah di akhir), baris-baris yang ikut bergeser di mirror lokal TIDAK
// otomatis membawa validasi dropdown Q yang lama -- Q bisa "nyasar" nempel
// di baris yang sekarang isinya PO lain. Fungsi ini menyamakan lagi:
// setiap baris di mirror dicek STATUS-nya FRESH dari database (bukan dari
// mirror yang bisa telat) dan dropdown Q dipasang ulang sesuai (EDIT/
// CANCEL untuk yang aktif, BATAL CANCEL untuk yang CANCELLED).
function refreshSemuaValidasiHistoryPO_(diamDiam) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hist = ss.getSheetByName("HISTORY_PO_MASUK");
  if (!hist) return;

  const lastRow = hist.getLastRow();
  if (lastRow < 2) {
    if (!diamDiam) SpreadsheetApp.getUi().alert("HISTORY_PO_MASUK masih kosong.");
    return;
  }

  const numRows = lastRow - 1;
  const statusFresh = bacaSheetGateway_("HISTORY_PO_MASUK", 2, 12, 1); // kolom L = STATUS

  const ruleAktif = SpreadsheetApp.newDataValidation()
    .requireValueInList(["EDIT", "CANCEL"], true)
    .setAllowInvalid(false)
    .build();
  const ruleCancelled = SpreadsheetApp.newDataValidation()
    .requireValueInList(["BATAL CANCEL"], true)
    .setAllowInvalid(false)
    .build();

  const rules = [];
  for (let i = 0; i < numRows; i++) {
    const status = statusFresh[i] ? String(statusFresh[i][0] || "").trim().toUpperCase() : "";
    rules.push([status === "CANCELLED" ? ruleCancelled : ruleAktif]);
  }

  hist.getRange(2, HISTORY_PO_ACTION_COL, numRows, 1).setDataValidations(rules);

  if (!diamDiam) {
    SpreadsheetApp.getUi().alert("Validation ACTION sudah di-refresh untuk semua baris (dibaca fresh dari database_SANFELO).");
  }
}

// Dipanggil manual dari menu / editor Apps Script.
function refreshAllHistoryPOValidation() {
  refreshSemuaValidasiHistoryPO_(false);
}

// Dipanggil otomatis lewat trigger waktu -- lihat
// pasangTriggerRefreshValidasiHistoryPO_() di bawah. Tanpa alert() karena
// trigger otomatis gak punya konteks UI (bakal error kalau ui.alert dipanggil).
function refreshValidasiHistoryPOOtomatis_() {
  refreshSemuaValidasiHistoryPO_(true);
}

// Jalankan SEKALI dari editor Apps Script buat masang auto-refresh tiap 10
// menit (jaring pengaman tambahan di luar refresh yang sudah dilakukan
// otomatis pas sync/cancel/batal-cancel).
function pasangTriggerRefreshValidasiHistoryPO_() {
  hapusTriggerRefreshValidasiHistoryPO_();
  ScriptApp.newTrigger("refreshValidasiHistoryPOOtomatis_")
    .timeBased()
    .everyMinutes(10)
    .create();
  SpreadsheetApp.getUi().alert("Auto-refresh dropdown ACTION HISTORY_PO_MASUK dipasang, jalan tiap 10 menit.");
}

function hapusTriggerRefreshValidasiHistoryPO_() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === "refreshValidasiHistoryPOOtomatis_") {
      ScriptApp.deleteTrigger(t);
    }
  });
}

/* ------------------------------------------------------
   11) BERSIHKAN DATA DOUBLE DI HISTORY_PO_MASUK
   ------------------------------------------------------ */
function buatKeyHistoryPO_(jenis, poNo, line) {
  return String(jenis).trim().toUpperCase() + "||" + String(poNo).trim() + "||" + String(line).trim();
}

function parseTimestamp_(value) {
  if (!value) return 0;
  if (value instanceof Date) return value.getTime();

  const str = String(value).trim();
  const m = str.match(/^(\d{2})\/(\d{2})\/(\d{4})\s+(\d{2}):(\d{2}):(\d{2})$/);
  if (!m) return 0;

  const [, dd, mm, yyyy, hh, mi, ss] = m;
  return new Date(yyyy, mm - 1, dd, hh, mi, ss).getTime();
}

function parseDoneQty_(statusStr, qtyAsli) {
  const s = String(statusStr || "").trim().toUpperCase();
  if (s === "DONE SJ") return qtyAsli;           // udah full
  const m = s.match(/^DONE SJ\s+(\d+(\.\d+)?)$/); // "DONE SJ 3" dst
  if (m) return Number(m[1]);
  return 0;                                       // belum pernah DONE SJ
}

// ⬅️ MIGRASI KE GATEWAY: dulu fungsi ini scan HISTORY_PO_MASUK lokal &
// hapus baris double LANGSUNG (hist.deleteRow). HISTORY_PO_MASUK di
// inventory_SANFELO sekarang cuma mirror (IMPORTRANGE) -- gak bisa lagi
// dihapus barisnya secara lokal (bakal error karena itu hasil array
// formula). Dedup sekarang dikerjain di server lewat aksi
// "bersihkanDuplikat" gateway (lihat aksiBersihkanDuplikat_ /
// jalankanDedupUntukPO_ di DatabaseGateway.js), tanpa jenis/poNo berarti
// bersihkan SELURUH sheet.
function bersihkanDuplikatHistoryPO() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert(
    "Bersihkan Data Double?",
    "Ini akan menghapus baris duplikat (JENIS+NO PO+LINE yang sama, kesimpen lebih dari sekali) di HISTORY_PO_MASUK milik database_SANFELO. Lanjut?",
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;

  const hasil = kirimAksiGateway_("bersihkanDuplikat", {});

  if (hasil.baris_dihapus === 0) {
    ui.alert("Tidak ada data double di HISTORY_PO_MASUK.");
    return;
  }

  ui.alert(hasil.baris_dihapus + " baris double berhasil dihapus dari HISTORY_PO_MASUK (database_SANFELO). Mirror lokal akan menyusul beberapa saat lagi.");
}

function perbaruiRangeArrayFormula_(sheet, anchorRow, oldLastRow, newLastRow) {
  if (oldLastRow === newLastRow) return;

  const lastCol = sheet.getLastColumn();
  const formulas = sheet.getRange(anchorRow, 1, 1, lastCol).getFormulas()[0];

  formulas.forEach(function(formula, idx) {
    if (!formula) return;

    const regex = new RegExp("([A-Z]{1,3}\\$?)" + oldLastRow + "(?![0-9])", "g");
    const formulaBaru = formula.replace(regex, "$1" + newLastRow);

    if (formulaBaru !== formula) {
      sheet.getRange(anchorRow, idx + 1).setFormula(formulaBaru);
    }
  });
}

/**
 * Versi delta-aware dari updateStatusDoneSJHistoryPO_.
 * oldMappingList = qty yg SEBELUMNYA sudah "DONE SJ" utk SJ ini (sebelum diedit).
 * newMappingList = qty yg SEKARANG (setelah diedit) mau di-set jadi "DONE SJ".
 * Selisihnya (delta) yang di-apply ke HISTORY_PO_MASUK, jadi:
 * - qty naik -> nambah DONE (seperti biasa)
 * - qty turun/item dihapus -> ngurangin DONE, otomatis balik jadi sisaQty lagi
 */
// ⬅️ MIGRASI KE GATEWAY: dulu fungsi ini baca+tulis STATUS (L) & warna (A:P)
// LANGSUNG ke HISTORY_PO_MASUK lokal. HISTORY_PO_MASUK di inventory_SANFELO
// sekarang cuma mirror -- perhitungan delta & penulisan STATUS/warna
// sekarang dikerjain di server (lihat aksiUpdateStatusDoneSjDelta_ di
// DatabaseGateway.js, logic delta-nya identik persis dengan yang di bawah
// sebelum migrasi). Tidak ada kolom lokal (Q/ACTION) yang kesentuh fungsi
// ini, jadi tidak perlu langkah tambahan di sisi inventory sesudahnya.
function updateStatusDoneSJHistoryPODelta_(oldMappingList, newMappingList) {
  const deltaMap = {};
  function tambahKe_(list, tanda) {
    (list || []).forEach(function(m) {
      const key = String(m.poNo).trim() + "||" + String(m.line).trim();
      deltaMap[key] = (deltaMap[key] || 0) + tanda * (Number(m.qty) || 0);
    });
  }
  tambahKe_(oldMappingList, -1);
  tambahKe_(newMappingList, 1);

  const keys = Object.keys(deltaMap).filter(function(k) { return deltaMap[k] !== 0; });
  if (keys.length === 0) return; // gak ada perubahan qty PO sama sekali

  kirimAksiGateway_("updateStatusDoneSjDelta", {
    oldMappingList: oldMappingList || [],
    newMappingList: newMappingList || []
  });
}

function getCounterDenganResetPeriode_(system, cellCounter, cellPeriode, periodeSekarang) {
  // ⬅️ MIGRASI KE GATEWAY: dulu langkah 1 di bawah baca HISTORY_PO_MASUK
  // LOKAL langsung (sumber kebenaran utama buat nomor PO tertinggi).
  // Sekarang HISTORY_PO_MASUK lokal cuma mirror (IMPORTRANGE) yang bisa
  // telat -- kalau tetap dipakai di sini risikonya nomor PO baru bisa
  // TABRAKAN sama nomor yang baru aja disimpan device/tab lain tapi belum
  // sempat ke-refresh di mirror. Makanya baca FRESH langsung dari database
  // lewat aksi "bacaSheet" gateway.
  let maxDariHistori = 0;
  const prefixCari = "#" + periodeSekarang + "-";
  const dataNoPO = bacaSheetGateway_("HISTORY_PO_MASUK", 2, 2, 1); // kolom B = NO PO

  dataNoPO.forEach(function(r) {
    const val = String(r[0] || "").trim();
    if (val.indexOf(prefixCari) === 0) {
      const num = parseInt(val.substring(prefixCari.length), 10);
      if (!isNaN(num) && num > maxDariHistori) maxDariHistori = num;
    }
  });

  // ---------- 2) Ambil counter cache di SYSTEM, TAPI cuma valid kalau
  //               periode tersimpan = periode sekarang (kalau beda tahun,
  //               berarti harus reset, jadi cache lama diabaikan) ----------
  const periodeTersimpan = String(system.getRange(cellPeriode).getValue() || "").trim();
  let counterCache = 0;

  if (periodeTersimpan === String(periodeSekarang).trim()) {
    counterCache = Number(system.getRange(cellCounter).getValue()) || 0;
  }

  // ---------- 3) Ambil yang PALING BESAR antara histori (utama) vs cache,
  //               biar aman kalau salah satu ketinggalan / gak sinkron ----------
  const maxNum = Math.max(maxDariHistori, counterCache);
  const nextNum = maxNum + 1;

  // ---------- 4) Update cache SYSTEM ----------
  system.getRange(cellCounter).setValue(nextNum);
  system.getRange(cellPeriode).setValue(periodeSekarang);

  return nextNum;
}

  // ======================= SJ_DRAFT_FETCH.js (salinan persis) =======================
/* ======================================================
   SJ_DRAFT_FETCH.gs
   - bukaDraft()                        : buka popup, pilih PENJUALAN / PEMBELIAN / SPAREPART
   - ambilDaftarSJTransaksi_()          : ambil semua No SJ + Customer dari TRANSAKSI
   - terapkanPilihanSJDraft()           : fetch item dari DETAIL_TRANSAKSI -> SURAT JALAN + INVOICE
   - ambilDaftarSJPembelian_()          : ambil semua No PB + Supplier dari PEMBELIAN
   - terapkanPilihanSJBeliDraft()       : fetch item dari PEMBELIAN -> TANDA TERIMA BRG
   - ambilDaftarSJSparepart_()          : ambil semua No SJ + Nama Toko dari SPAREPART
   - terapkanPilihanSJSparepartDraft()  : fetch item dari SPAREPART -> SJ_SPR + INV_SPR
   (dipanggil dari popup via google.script.run)

   CATATAN:
   - Panggil bukaDraft() dari menu UI (custom menu di onOpen).
   - Fungsi ini TIDAK lewat trigger onEdit, langsung fetch server-side.
     Karena setValue() dari Apps Script gak memicu simple trigger onEdit,
     jadi aman, gak bakal nabrak logic handleOnEdit yang sudah ada di
     PO_HISTORY.gs.
   - Reuse konstanta & helper yang sudah ada di PO_HISTORY.gs (khusus flow
     PENJUALAN): SJ_ITEM_START_ROW, SJ_ITEM_MAX_ROWS, SJ_FOOTER_COL_START,
     SJ_FOOTER_COL_END, SJ_FOOTER_TEKS, INV_DISPLAY_START_ROW, INV_ITEM_MAX_ROWS,
     INV_FOOTER_COL_START, INV_FOOTER_COL_END, INV_FOOTER_TEKS,
     INV_KETERANGAN_ANCHOR_TEKS, getEffectiveBaseRows_(), pastikanKapasitasItem_(),
     adjustSJItemBorders(), adjustInvoiceItemBorders(), cariBarisAnchor_(),
     sembunyikanKolomMappingSJ_(). Semua itu global di project yang sama.
   - Flow PEMBELIAN -> TANDA TERIMA BRG TIDAK pakai kapasitas dinamis / extend
     baris / adjust border (sesuai konfirmasi user), jadi item langsung
     ditulis mulai row 7 tanpa proses tambahan.
   - Flow SPAREPART -> SJ_SPR + INV_SPR juga TIDAK pakai kapasitas dinamis,
     karena SJ_SPR/INV_SPR emang form dengan kapasitas tetap (maks
     SPR_MAX_LINES_ = 8 baris barang, dideklarasikan di SPAREPART_SAVE.gs).
   - Reuse SPR_MAX_LINES_ dari SPAREPART_SAVE.gs (global, project yang sama).

   ⬅️ PATCH (fix 2 bug flow PENJUALAN, per diskusi):
   1) BUG: kolom I (mapping poNo+line, dipakai buat update status DONE SJ di
      HISTORY_PO_MASUK) dulu di-clearContent() aja pas fetch draft, gak
      pernah ditulis ulang -> re-save draft SJ yg asalnya dari PO jadi
      gak update HISTORY_PO_MASUK sama sekali (silent no-op).
      FIX: mapping poNo+line sekarang dibawa & disimpan PERMANEN di
      DETAIL_TRANSAKSI kolom S (SOURCE_PO, JSON) oleh simpanDraft() di
      Transaction.gs (lihat patch terpisah). Fungsi terapkanPilihanSJDraft()
      di file ini tinggal BACA balik kolom S itu dan tulis ulang ke kolom I
      SURAT JALAN, supaya re-save berikutnya bisa hitung delta yg benar
      (naik/turun qty) via updateStatusDoneSJHistoryPODelta_() di PO_HISTORY.gs.
   2) BUG: sheet INVOICE gak pernah ke-fetch sama sekali pas pilih draft SJ
      (cuma sheet SURAT JALAN yg diisi). Ini bisa bikin data invoice gak
      sinkron, atau malah bikin simpanDraft() error "jumlah baris item SJ
      dan INVOICE tidak sinkron" kalau row count beda.
      FIX: setelah nulis ke SURAT JALAN, sekarang juga fetch header (No INV,
      Tempo, Method, Sales, Total Bayar, Keterangan) dari TRANSAKSI, dan
      qty/harga/diskon/hargaNet item dari DETAIL_TRANSAKSI, ke sheet INVOICE.
      Kolom kode/barang/qty di INVOICE diasumsikan ARRAYFORMULA yg narik
      dari SURAT JALAN (sesuai pola perbaruiRangeArrayFormula_ di
      PO_HISTORY.gs) -> jadi cuma perlu isi kolom G/H/I (harga/diskon/net).
      KALAU TERNYATA kolom itu bukan formula di project kamu, kasih tau,
      soalnya perlu ditulis manual juga.

   STRUKTUR SHEET PEMBELIAN (1 baris = 1 item, header info diulang tiap baris):
   A NO PB | B NO PO | C NO SJ | D SUPPLIER | E TGL BELI | F NO KEND |
   G KODE | H NAMA BARANG (skip) | I QTY | J TOTAL (skip) | K KETERANGAN |
   L PENERIMA | M PENGIRIM

   MAPPING PEMBELIAN -> TANDA TERIMA BRG:
   A (NO PB)      -> B4
   B (NO PO)      -> B3
   C (NO SJ)      -> E2
   D (SUPPLIER)   -> B2
   E (TGL BELI)   -> E3
   F (NO KEND)    -> E4
   G (KODE)       -> mulai B7
   I (QTY)        -> mulai C7
   K (KETERANGAN) -> mulai F7
   L (PENERIMA)   -> G2
   M (PENGIRIM)   -> G3
   (kolom A "NO" di TANDA TERIMA BRG diisi otomatis 1,2,3,... mengikuti urutan item)

   STRUKTUR SHEET SPAREPART (1 baris = 1 item, header info diulang tiap baris):
   A NO SJ | B NO INV | C (KOLOM BARU, blm dikonfirmasi namanya) | D TANGGAL |
   E TEMPO | F NO KEND | G NAMA TOKO | H ALAMAT | I BARANG | J QTY | K SAT |
   L KET BRG | M KET INV | N HARGA | O HPP BARANG | P METHOD | Q SALES |
   R LABA | S STATUS ORDER

   MAPPING SPAREPART -> SJ_SPR / INV_SPR (per konfirmasi user 30 Jul 2026,
   revisi dari mapping lama yang sudah gak sesuai struktur sheet terbaru):
   A (NO SJ)     -> SJ_SPR!H6
   B (NO INV)    -> INV_SPR!H5
   C (BARU)      -> INV_SPR!H7
   D (TANGGAL)   -> SJ_SPR!H5
   E (TEMPO)     -> INV_SPR!H9
   F (NO KEND)   -> SJ_SPR!H7
   G (NAMA TOKO) -> SJ_SPR!A5
   H (ALAMAT)    -> SJ_SPR!A6
   I (BARANG)    -> SJ_SPR! mulai B9
   J (QTY)       -> SJ_SPR! mulai E9
   K (SAT)       -> SJ_SPR! mulai F9
   L (KET BRG)   -> SJ_SPR! mulai G9
   M (KET INV)   -> INV_SPR!C25
   N (HARGA)     -> INV_SPR! mulai G11
   P (METHOD)    -> SJ_SPR!H3
   Q (SALES)     -> INV_SPR!C20
   (O HPP BARANG, R LABA gak perlu ditulis balik ke form -- dihitung ulang
    otomatis oleh simpanSparepart() pas disimpan lagi. S STATUS ORDER juga
    gak ditulis ke form, cuma dibawa lewat PropertiesService biar
    dipertahankan pas simpanSparepart() jalan lagi di mode EDIT.)
   ====================================================== */

/* NOTE: TANDA_TERIMA_ITEM_START_ROW TIDAK dideklarasi di sini karena sudah ada
   di file .gs lain pada project ini. Kalau ternyata belum ada / bentrok
   nilai, cari deklarasinya (grep "TANDA_TERIMA_ITEM_START_ROW" di semua file)
   dan pastikan nilainya = 7. */

/* ======================================================
   ========== FLOW PENJUALAN (TRANSAKSI -> SURAT JALAN + INVOICE) ==========
   ====================================================== */

/* ------------------------------------------------------
   1) AMBIL DAFTAR NO SJ DARI TRANSAKSI (untuk isi popup)
   ------------------------------------------------------ */
function ambilDaftarSJTransaksi_() {
  // ⬅️ MIGRASI (Oktober 2026): dibaca FRESH dari TRANSAKSI di database
  // (draft yang baru disimpan langsung muncul, tanpa nunggu salinan lokal).
  // B = NO SJ, E = CUSTOMER
  const data = bacaSheetGateway_("TRANSAKSI", 2, 2, 4); // B..E
  if (data.length === 0) return [];

  const list = [];
  data.forEach(function(row) {
    const noSJ = String(row[0] || "").trim();
    const customer = String(row[3] || "").trim();
    if (!noSJ) return;
    list.push({ noSJ: noSJ, customer: customer });
  });

  // baris paling bawah = data paling baru -> tampilkan paling atas
  list.reverse();
  return list;
}

/* ------------------------------------------------------
   2) FETCH DATA NO SJ TERPILIH -> SHEET SURAT JALAN + INVOICE
      Dipanggil dari popup lewat google.script.run
   ------------------------------------------------------ */
function terapkanPilihanSJDraft(noSJ) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const sj  = ss.getSheetByName("SURAT JALAN");
  const inv = ss.getSheetByName("INVOICE"); // ⬅️ dipakai di bagian fetch INVOICE (boleh null, di-guard)
  // ⬅️ MIGRASI (Oktober 2026): TRANSAKSI & DETAIL_TRANSAKSI dibaca FRESH
  // dari database (bacaTransaksiDB_ / bacaDetailDB_ di TRANSAKSI_DB.js).

  if (!sj) throw new Error('Sheet "SURAT JALAN" tidak ditemukan.');

  noSJ = String(noSJ || "").trim();
  if (!noSJ) throw new Error("No SJ kosong.");

  /* ---------- 1) ambil header dari TRANSAKSI ---------- */
  const trxData = bacaTransaksiDB_(); // A..Q
  if (trxData.length === 0) throw new Error("Sheet TRANSAKSI masih kosong.");

  let header = null;
  // ⬅️ TAMBAHAN: field2 tambahan buat INVOICE, diambil dari baris TRANSAKSI yg sama
  let noINV = "", tempo = null, method = "", sales = "", totalBayar = "", keterangan = "";

  for (let i = 0; i < trxData.length; i++) {
    const rowNoSJ = String(trxData[i][1] || "").trim(); // B = NO SJ
    if (rowNoSJ === noSJ) {
      header = {
        customer: trxData[i][4],   // E CUSTOMER
        alamat:   trxData[i][5],   // F ALAMAT
        tanggal:  trxData[i][3],   // D TANGGAL
        noPol:    trxData[i][10]   // K NO POL
      };
      // ⬅️ TAMBAHAN: dipakai buat fetch INVOICE di bawah
      noINV      = trxData[i][2];  // C NO INV
      tempo      = trxData[i][6];  // G TEMPO
      method     = trxData[i][9];  // J METHOD
      sales      = trxData[i][11]; // L SALES
      totalBayar = trxData[i][12]; // M TOTAL BAYAR
      keterangan = trxData[i][15]; // P KETERANGAN
      break;
    }
  }

  if (!header) throw new Error('No SJ "' + noSJ + '" tidak ditemukan di TRANSAKSI.');

  /* ---------- 2) ambil item dari DETAIL_TRANSAKSI ---------- */

  // ⬅️ FIX: dulu cuma baca A..S (19 kolom) dan ambil SOURCE_PO dari kolom S
  // (row[18]). Tapi SOURCE_PO sudah dipindah ke kolom U (lihat
  // DET_COL_SOURCE_PO = 21 di Transaction.js) — kolom S sekarang isinya
  // LABA, bukan SOURCE_PO lagi. Akibatnya draft yang dibuka ulang selalu
  // baca angka LABA sebagai JSON mapping PO, gagal parse, dan diam-diam
  // gagal nandain barang itu "DONE SJ" ke PO aslinya — stok dari PO itu
  // bisa keliatan masih tersedia dan ke-invoice/terkirim dua kali.
  // Sekarang baca sampai kolom U (21 kolom) dan ambil dari situ.
  const dtlData = bacaDetailDB_(); // A..U
  if (dtlData.length === 0) throw new Error("Sheet DETAIL_TRANSAKSI masih kosong.");

  const items = [];
  dtlData.forEach(function(row) {
    const rowNoSJ = String(row[0] || "").trim(); // A NO SJ
    if (rowNoSJ !== noSJ) return;

    const kode = row[5]; // F KODE
    if (!kode) return;

    items.push({
      kode: kode,        // F KODE
      qty:  row[7],       // H QTY
      sat:  row[8],       // I SAT
      ket:  row[9],        // J KETERANGAN
      harga:    row[12],   // M HARGA        ⬅️ TAMBAHAN (buat INVOICE)
      diskon:   row[13],   // N DISKON       ⬅️ TAMBAHAN (buat INVOICE)
      hargaNet: row[14],   // O HARGA NET    ⬅️ TAMBAHAN (buat INVOICE)
      sourcePO: row[20] || "" // U SOURCE_PO ⬅️ FIX (dulu row[18]/kolom S, sekarang kolom U yang benar)
    });
  });

  if (items.length === 0) {
    throw new Error('Tidak ada item untuk No SJ "' + noSJ + '" di DETAIL_TRANSAKSI.');
  }

  /* ---------- 3) tulis ke sheet SURAT JALAN ---------- */
  // Kosongin No PO (H3), isi No SJ (H6) — sesuai mode "fetch by No SJ"
  sj.getRange("H3").clearContent();
  sj.getRange("H6").setValue(noSJ);

  const baseRowsSJ = getEffectiveBaseRows_(SJ_ITEM_MAX_ROWS);
  const neededRows = Math.max(baseRowsSJ, items.length);

  pastikanKapasitasItem_(
    sj, SJ_ITEM_START_ROW, baseRowsSJ, neededRows,
    SJ_FOOTER_COL_START, SJ_FOOTER_COL_END, SJ_FOOTER_TEKS
  );

  sj.getRange("A5").setValue(header.customer || "");
  sj.getRange("A6").setValue(header.alamat || "");
  if (header.tanggal) setTanggalSaja_(sj.getRange("H5"), header.tanggal);
  sj.getRange("H7").setValue(header.noPol || "");

  // bersihin dulu area item
  sj.getRange(SJ_ITEM_START_ROW, 2, neededRows, 1).clearContent(); // B kode
  sj.getRange(SJ_ITEM_START_ROW, 5, neededRows, 1).clearContent(); // E qty
  sj.getRange(SJ_ITEM_START_ROW, 6, neededRows, 1).clearContent(); // F satuan
  sj.getRange(SJ_ITEM_START_ROW, 7, neededRows, 1).clearContent(); // G keterangan
  sj.getRange(SJ_ITEM_START_ROW, 9, neededRows, 2).clearContent(); // I,J mapping (di-restore lagi di bawah, per item)

  items.forEach(function(item, i) {
    const targetRow = SJ_ITEM_START_ROW + i;
    sj.getRange(targetRow, 2).setValue(item.kode);
    sj.getRange(targetRow, 5).setValue(item.qty);
    sj.getRange(targetRow, 6).setValue(item.sat || "");
    sj.getRange(targetRow, 7).setValue(item.ket || "");
    sj.getRange(targetRow, 9).setValue(item.sourcePO); // ⬅️ TAMBAHAN: restore mapping poNo/line (JSON)
  });

  sembunyikanKolomMappingSJ_(sj);
  adjustSJItemBorders();

  /* ---------- 4) fetch data yg sesuai ke sheet INVOICE ⬅️ TAMBAHAN (fix bug #2) ----------
     Asumsi: kolom kode/barang/qty di INVOICE adalah ARRAYFORMULA yg narik dari
     SURAT JALAN (pola yg sama seperti perbaruiRangeArrayFormula_ di PO_HISTORY.gs).
     Jadi di sini cuma perlu isi header + kolom G/H/I (harga/diskon/hargaNet).
     Kalau ternyata bukan formula di project kamu, kasih tau -- perlu ditulis manual juga. */
  if (inv) {
    const baseRowsInv = getEffectiveBaseRows_(INV_ITEM_MAX_ROWS);
    const neededRowsInv = Math.max(baseRowsInv, items.length);

    pastikanKapasitasItem_(
      inv, INV_DISPLAY_START_ROW, baseRowsInv, neededRowsInv,
      INV_FOOTER_COL_START, INV_FOOTER_COL_END, INV_FOOTER_TEKS
    );

    inv.getRange("H6").setValue(noINV || "");
    if (header.tanggal) setTanggalSaja_(inv.getRange("H7"), header.tanggal);

    // ✅ FIX (bug: Jatuh Tempo "nyantol" gak ikut update DISKON): SEBELUMNYA
    // baris ini cuma nge-paste ULANG tanggal tempo LAMA dari TRANSAKSI!G apa
    // adanya (angka MATI, hasil hitungan waktu draft ini PERTAMA KALI dibuat).
    // Jadi kalau sheet DISKON toko itu diubah belakangan (mis. 30 HARI -> 45
    // HARI), draft yang sudah ada tetap kebawa 30 hari terus tiap kali dibuka
    // ulang lewat "Buka Draft", gak pernah kehitung ulang.
    // FIX: hitung ULANG tempoHari-nya persis kayak transaksiBaru() di
    // Transaction.gs (pakai hitungTempoHariUntukTransaksi_, yang sama), lalu
    // tulis sebagai RUMUS "=H7+n" ke H8 -- biar SELALU ikut nilai DISKON /
    // status SALE-DISPLAY yang PALING BARU, bukan angka beku dari draft lama.
    // noPOGabungan direkonstruksi dari SOURCE_PO tiap item (H3 sengaja
    // dikosongkan di flow "fetch by No SJ" ini, lihat baris 226), supaya
    // status SALE/DISPLAY di HISTORY_PO_MASUK tetap ikut diperhitungkan --
    // "DISPLAY tetap 90 hari" TIDAK berubah oleh fix ini.
    const noPOGabunganDraft = ambilNoPOGabunganDariItems_(items);
    const tempoHariDraft = hitungTempoHariUntukTransaksi_(noPOGabunganDraft, header.customer, keterangan);

    if (tempoHariDraft !== null && !isNaN(tempoHariDraft)) {
      inv.getRange("H8").setFormula("=H7+" + tempoHariDraft).setNumberFormat("dd/mm/yyyy");
    } else if (tempo) {
      // fallback TERAKHIR kalau tempoHari beneran gak bisa dihitung sama
      // sekali (mis. toko-nya sudah gak ada lagi di DISKON) -- daripada H8
      // kosong, tetap dipasang tanggal lama seperti sebelumnya.
      inv.getRange("H8").setValue(tempo);
    }

    inv.getRange("I3").setValue(method || "");

    const footerRowInv = cariBarisAnchor_(inv, INV_FOOTER_TEKS, INV_DISPLAY_START_ROW);
    if (footerRowInv !== -1) {
      inv.getRange(footerRowInv + 1, 3).setValue(sales || "");      // C = Nama Sales
      // ✅ FIX: pakai RUMUS, JANGAN angka statis dari totalBayar — biar H22
      // selalu otomatis = Total Penjualan (H{footerRowInv}) - Discount (H{footerRowInv+1}),
      // gak pernah "beku" jadi angka mati tiap kali draft ini dibuka ulang.
      inv.getRange(footerRowInv + 3, 8).setFormula(
        "=H" + footerRowInv + "-H" + (footerRowInv + 1)
      );
    }

    const ketRow = cariBarisAnchor_(inv, INV_KETERANGAN_ANCHOR_TEKS, INV_DISPLAY_START_ROW);
    if (ketRow !== -1) inv.getRange(ketRow, 3).setValue(keterangan || "");

    // bersihin area harga/diskon/net lama, lalu tulis ulang per item
    inv.getRange(INV_DISPLAY_START_ROW, 7, neededRowsInv, 3).clearContent(); // G,H,I

    items.forEach(function(item, i) {
      const targetRowInv = INV_DISPLAY_START_ROW + i;
      inv.getRange(targetRowInv, 7).setValue(item.harga);    // G HARGA
      inv.getRange(targetRowInv, 8).setValue(item.diskon);   // H DISKON
      inv.getRange(targetRowInv, 9).setValue(item.hargaNet); // I HARGA NET
    });

    adjustInvoiceItemBorders();
  }

  return { ok: true, noSJ: noSJ, jumlahItem: items.length };
}

/**
 * ⬅️ BARU: rekonstruksi daftar NO PO (gabungan, dipisah koma) dari kolom
 * SOURCE_PO tiap item (JSON string "[{poNo, line, qty}, ...]", sama seperti
 * yang ditulis simpanDraft() di Transaction.gs). Dipakai supaya
 * terapkanPilihanSJDraft() bisa manggil hitungTempoHariUntukTransaksi_()
 * (dari Transaction.gs) dengan sinyal NO PO yang sama seperti waktu
 * transaksi ini pertama kali dibuat -- termasuk status SALE/DISPLAY di
 * HISTORY_PO_MASUK -- meskipun field No PO di form SURAT JALAN (H3) sendiri
 * sengaja dikosongkan di flow "fetch by No SJ" ini.
 * Item tanpa SOURCE_PO (mis. barang lama / manual tanpa PO) dilewati aja,
 * gak bikin error -- hasilnya sekadar gabungan NO PO yang KETEMU.
 */
function ambilNoPOGabunganDariItems_(items) {
  const poSet = {};
  items.forEach(function (item) {
    if (!item.sourcePO) return;
    try {
      const sources = JSON.parse(item.sourcePO);
      if (Array.isArray(sources)) {
        sources.forEach(function (src) {
          const poNo = String((src && src.poNo) || "").trim();
          if (poNo) poSet[poNo] = true;
        });
      }
    } catch (err) {
      // SOURCE_PO kosong/format lama yang gak dikenal -> lewati, gak fatal
    }
  });
  return Object.keys(poSet).join(",");
}

/* ======================================================
   ========== FLOW PEMBELIAN (PEMBELIAN -> TANDA TERIMA BRG) ==========
   ⬅️ BARU (migrasi PEMBELIAN, Oktober 2026): kedua fungsi di bawah baca
   FRESH dari database lewat bacaSheetGateway_, bukan dari sheet
   PEMBELIAN lokal lagi (itu sekarang cuma mirror IMPORTRANGE) — sama
   pola kayak flow SPAREPART. Logic isi popup / tulis balik ke form
   TANDA TERIMA BRG TIDAK BERUBAH SAMA SEKALI.
   ====================================================== */

/* ------------------------------------------------------
   3) AMBIL DAFTAR NO PB DARI PEMBELIAN (untuk isi popup)
      1 NO PB bisa berulang di banyak baris (1 baris = 1 item),
      jadi di-dedupe supaya 1 NO PB cuma muncul 1x di daftar.
   ------------------------------------------------------ */
function ambilDaftarSJPembelian_() {
  const data = bacaSheetGateway_("PEMBELIAN", 2, 1, 13); // A..M
  if (data.length === 0) return [];

  const seen = {};
  const list = [];
  data.forEach(function(row) {
    const noPB = String(row[0] || "").trim(); // A NO PB
    if (!noPB || seen[noPB]) return;
    seen[noPB] = true;

    const supplier = String(row[3] || "").trim(); // D SUPPLIER
    list.push({ noPB: noPB, supplier: supplier });
  });

  // baris paling bawah = data paling baru -> tampilkan paling atas
  list.reverse();
  return list;
}

/* ------------------------------------------------------
   4) FETCH DATA NO PB TERPILIH -> SHEET TANDA TERIMA BRG
      Dipanggil dari popup lewat google.script.run

      ⬅️ FIX (Nov 2026): disamakan dengan versi SL. Dulu cuma nulis B (kode),
      C (qty), F (keterangan) -- kolom G (NO PO asal) & H (status SET) gak
      ikut direstore, formula harga gak dihitung ulang, dan gak ada popup.
      Sekarang: item ditulis lengkap, HARGA dihitung ulang, lalu muncul
      popup daftar barang (diskon SET). Tombol "Lewati" di popup tidak
      menghapus apa pun -- daftar barang tetap ada di form.
   ------------------------------------------------------ */
function terapkanPilihanSJBeliDraft(noPB) {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const shTerima  = ss.getSheetByName("TANDA TERIMA BRG");

  if (!shTerima) throw new Error('Sheet "TANDA TERIMA BRG" tidak ditemukan.');

  noPB = String(noPB || "").trim();
  if (!noPB) throw new Error("No PB kosong.");

  const data = bacaSheetGateway_("PEMBELIAN", 2, 1, 15); // A..O (O = STATUS SET)
  if (data.length === 0) throw new Error("Sheet PEMBELIAN masih kosong.");

  let header = null;
  const items = [];

  data.forEach(function(row) {
    const rowNoPB = String(row[0] || "").trim(); // A NO PB
    if (rowNoPB !== noPB) return;

    // header cuma diambil sekali, dari baris pertama yang match
    if (!header) {
      header = {
        noPO:     row[1],   // B NO PO
        noSJ:     row[2],   // C NO SJ
        supplier: row[3],   // D SUPPLIER
        tglBeli:  row[4],   // E TGL BELI
        noKend:   row[5],   // F NO KEND
        penerima: row[11],  // L PENERIMA
        pengirim: row[12]   // M PENGIRIM
      };
    }

    const kode = row[6]; // G KODE
    if (!kode) return;

    // B di PEMBELIAN SANFELO = gabungan semua PO di header. PO asal per item
    // cuma diisi kalau PO-nya tunggal (kalau gabungan, dicari ulang saat simpan).
    const poList = parsePoListTT_(row[1]);

    items.push({
      kode:      kode,                                      // G KODE   -> form B
      nama:      row[7],                                    // H NAMA   -> cuma buat tampilan popup
      qty:       row[8],                                    // I QTY    -> form C
      toko:      row[10],                                   // K NAMA TOKO -> form F
      po:        poList.length === 1 ? poList[0] : "",      // -> form G
      setDiskon: String(row[14] || "").trim().toUpperCase() // O STATUS SET -> form H
    });
  });

  if (!header) throw new Error('No PB "' + noPB + '" tidak ditemukan di PEMBELIAN.');
  if (items.length === 0) {
    throw new Error('Tidak ada item untuk No PB "' + noPB + '" di PEMBELIAN.');
  }
  if (items.length > TANDA_TERIMA_ITEM_MAX_ROWS) {
    throw new Error(
      'Item (' + items.length + ') melebihi kapasitas TANDA TERIMA BRG (' +
      TANDA_TERIMA_ITEM_MAX_ROWS + ' baris).'
    );
  }

  // TGL BELI dari gateway datang sebagai teks ISO -> balikin jadi tanggal
  let tglBeli = header.tglBeli || "";
  if (typeof tglBeli === "string" && /^\d{4}-\d{2}-\d{2}T/.test(tglBeli)) tglBeli = new Date(tglBeli);

  /* ---------- header ---------- */
  shTerima.getRange("B2").setValue(header.supplier || "");
  shTerima.getRange("B3").setValue(header.noPO || "");
  shTerima.getRange("B4").setValue(noPB);
  shTerima.getRange("E2").setValue(header.noSJ || "");
  shTerima.getRange("E3").setValue(tglBeli).setNumberFormat("dd/MM/yyyy");
  shTerima.getRange("E4").setValue(header.noKend || "");
  shTerima.getRange("G2").setValue(header.penerima || "");
  shTerima.getRange("G3").setValue(header.pengirim || "");

  /* ---------- bersihin item lama ----------
     A & D = ARRAYFORMULA (JANGAN disentuh). Yang dibersihkan: B:C dan F:H. */
  const n  = TANDA_TERIMA_ITEM_MAX_ROWS;
  const r0 = TANDA_TERIMA_ITEM_START_ROW;
  shTerima.getRange(r0, 2, n, 2).clearContent(); // B:C (kode, qty)
  shTerima.getRange(r0, 6, n, 3).clearContent(); // F:H (toko, PO asal, SET)

  /* ---------- tulis item baru (batch) ---------- */
  shTerima.getRange(r0, 2, items.length, 2).setValues(
    items.map(function(it) { return [it.kode, it.qty]; })
  );
  shTerima.getRange(r0, 6, items.length, 3).setValues(
    items.map(function(it) { return [it.toko || "", it.po || "", it.setDiskon || ""]; })
  );

  // hitung ulang HARGA (sudah termasuk diskon SET lama) + baris Total
  pastikanFormulaTandaTerimaUtuh_(shTerima);
  SpreadsheetApp.flush();

  /* ---------- popup daftar barang + diskon SET ---------- */
  tampilkanPopupDiskonSetDraft_(items);

  return { ok: true, noPB: noPB, jumlahItem: items.length };
}

// Popup yang sama dengan "Pilih Barang", tapi modeDraft = true:
// checkbox & qty terkunci, dropdown SET terisi pilihan lama.
// Tombol "Lewati" cuma menutup popup -- isi form TIDAK diubah.
function tampilkanPopupDiskonSetDraft_(items) {
  const template = HtmlService.createTemplateFromFile("PopupPilihBarangTandaTerima");
  template.modeDraft = true;
  template.itemsJson = JSON.stringify(items.map(function(it) {
    return {
      poNo:       String(it.po || "-"),
      namaToko:   it.toko || "",
      namaBarang: String(it.nama || it.kode),
      qtyAsli:    Number(it.qty) || 0,
      sisaQty:    Number(it.qty) || 0,
      setDiskon:  it.setDiskon || ""
    };
  }));
  const html = template.evaluate().setWidth(480).setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, "Diskon SET per Barang");
}

// Dipanggil dari popup (mode draft).
// selectedJson = array status SET, urut sesuai baris item di form.
function terapkanDiskonSetDraft(selectedJson) {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const shTerima = ss.getSheetByName("TANDA TERIMA BRG");
  const arr      = JSON.parse(selectedJson);
  if (!arr.length) return { ok: true };

  const vals = arr.map(function(s) {
    const v = String(s || "").trim().toUpperCase();
    return [(v === "3SET" || v === "5SET") ? v : ""];
  });

  shTerima.getRange(TANDA_TERIMA_ITEM_START_ROW, TANDA_TERIMA_COL_SET, vals.length, 1).setValues(vals);
  pastikanFormulaTandaTerimaUtuh_(shTerima); // hitung ulang HARGA + baris Total
  return { ok: true };
}

/* ======================================================
   ========== FLOW SPAREPART (SPAREPART -> SJ_SPR + INV_SPR) ==========
   (Tidak ada perubahan logic dari sebelumnya)
   ====================================================== */

/* ------------------------------------------------------
   5) AMBIL DAFTAR NO SJ DARI SPAREPART (untuk isi popup)
      1 NO SJ bisa berulang di banyak baris (1 baris = 1 barang),
      jadi di-dedupe supaya 1 NO SJ cuma muncul 1x di daftar.
      ⬅️ BARU: baca FRESH dari database lewat bacaSheetGateway_, bukan
      dari sheet SPAREPART lokal lagi (itu sekarang cuma mirror).
   ------------------------------------------------------ */
function ambilDaftarSJSparepart_() {
  const data = bacaSheetGateway_("SPAREPART", 2, 1, 7); // A..G (butuh G = NAMA TOKO)
  if (data.length === 0) return [];

  const seen = {};
  const list = [];
  data.forEach(function(row) {
    const noSJ = String(row[0] || "").trim(); // A NO SJ
    if (!noSJ || seen[noSJ]) return;
    seen[noSJ] = true;

    const namaToko = String(row[6] || "").trim(); // G NAMA TOKO
    list.push({ noSJ: noSJ, namaToko: namaToko });
  });

  // baris paling bawah = data paling baru -> tampilkan paling atas
  list.reverse();
  return list;
}

/* ------------------------------------------------------
   6) FETCH DATA NO SJ TERPILIH -> SHEET SJ_SPR + INV_SPR
      Dipanggil dari popup lewat google.script.run

      Setelah ini, form SJ_SPR/INV_SPR ke-isi ulang seperti kondisi
      transaksi lama tsb, DAN ditandai lagi mode "EDIT" (via
      PropertiesService) supaya nanti pas user pencet tombol
      "Simpan Sparepart" (simpanSparepart() di SPAREPART_SAVE.gs),
      baris lama utk No SJ ini DITIMPA -- bukan ditambah sbg baris
      baru. STATUS ORDER lama ikut dibawa biar dipertahankan.
      ⬅️ BARU: baca FRESH dari database lewat bacaSheetGateway_, bukan
      dari sheet SPAREPART lokal lagi (itu sekarang cuma mirror).
   ------------------------------------------------------ */
function terapkanPilihanSJSparepartDraft(noSJ) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const shSJ  = ss.getSheetByName("SJ_SPR");
  const shINV = ss.getSheetByName("INV_SPR");

  if (!shSJ) throw new Error('Sheet "SJ_SPR" tidak ditemukan.');
  if (!shINV) throw new Error('Sheet "INV_SPR" tidak ditemukan.');

  noSJ = String(noSJ || "").trim();
  if (!noSJ) throw new Error("No SJ kosong.");

  const data = bacaSheetGateway_("SPAREPART", 2, 1, 19); // A..S
  if (data.length === 0) throw new Error("Sheet SPAREPART masih kosong.");

  let header = null;
  const items = [];

  data.forEach(function(row) {
    const rowNoSJ = String(row[0] || "").trim(); // A NO SJ
    if (rowNoSJ !== noSJ) return;

    // header cuma diambil sekali, dari baris pertama yang match
    if (!header) {
      header = {
        noInv:    row[1],  // B NO INV
        kolomC:   row[2],  // C (field baru, blm dikonfirmasi namanya) -> INV_SPR H7
        tanggal:  row[3],  // D TANGGAL
        tempo:    row[4],  // E TEMPO
        noKend:   row[5],  // F NO KEND
        namaToko: row[6],  // G NAMA TOKO
        alamat:   row[7],  // H ALAMAT
        ketInv:   row[12], // M KET INV
        method:   row[15], // P METHOD
        sales:    row[16], // Q SALES
        status:   row[18]  // S STATUS ORDER (dibawa biar dipertahankan)
      };
    }

    const barang = row[8]; // I BARANG
    if (!barang) return;

    items.push({
      barang: barang,
      qty:    row[9],   // J QTY
      sat:    row[10],  // K SAT
      ketBrg: row[11],  // L KET BRG
      harga:  row[13]   // N HARGA
    });
  });

  if (!header) throw new Error('No SJ "' + noSJ + '" tidak ditemukan di SPAREPART.');
  if (items.length === 0) {
    throw new Error('Tidak ada item untuk No SJ "' + noSJ + '" di SPAREPART.');
  }
  if (items.length > SPR_MAX_LINES_) {
    throw new Error(
      'No SJ "' + noSJ + '" punya ' + items.length + ' item, ' +
      'melebihi kapasitas form SJ_SPR/INV_SPR (' + SPR_MAX_LINES_ + ' baris).'
    );
  }

  /* ---------- tulis header ke SJ_SPR & INV_SPR ---------- */
  shSJ.getRange('H3').setValue(header.method || "");   // Metode Pembayaran
  if (header.tanggal) setTanggalSaja_(shSJ.getRange('H5'), header.tanggal); // Tgl SJ
  shSJ.getRange('H6').setValue(noSJ);                   // No SJ
  shSJ.getRange('H7').setValue(header.noKend || "");    // No Kendaraan
  shSJ.getRange('A5').setValue(header.namaToko || "");  // Nama Toko
  shSJ.getRange('A6').setValue(header.alamat || "");    // Alamat

  shINV.getRange('H5').setValue(header.noInv || "");    // No Inv
  shINV.getRange('H7').setValue(header.kolomC || "");   // (field baru, kolom C)
  if (header.tempo) shINV.getRange('H9').setValue(header.tempo); // Jatuh tempo
  shINV.getRange('C20').setValue(header.sales || "");   // Nama Sales
  shINV.getRange('C25').setValue(header.ketInv || "");  // Keterangan invoice

  /* ---------- bersihin area item lama di form, lalu tulis item ---------- */
  shSJ.getRange('B9:G16').clearContent();
  shINV.getRange('G11:G18').clearContent();

  items.forEach(function(item, i) {
    const sjRow  = 9 + i;
    const invRow = 11 + i;
    shSJ.getRange(sjRow, 2).setValue(item.barang);        // B Barang
    shSJ.getRange(sjRow, 5).setValue(item.qty);           // E Qty
    shSJ.getRange(sjRow, 6).setValue(item.sat || "");     // F SAT
    shSJ.getRange(sjRow, 7).setValue(item.ketBrg || "");  // G Keterangan barang
    shINV.getRange(invRow, 7).setValue(item.harga);       // G Total (Rp)
  });

  SpreadsheetApp.flush();

  /* ---------- tandai form ini sedang mode EDIT utk No SJ ini ---------- */
  PropertiesService.getDocumentProperties().setProperties({
    'SPAREPART_EDIT_NOSJ': noSJ,
    'SPAREPART_EDIT_STATUS': String(header.status || 'NOT DONE')
  });

  return { ok: true, noSJ: noSJ, jumlahItem: items.length };
}

/* ======================================================
   ========== POPUP (mode PENJUALAN / PEMBELIAN / SPAREPART) ==========
   ====================================================== */

/* ------------------------------------------------------
   7) BUKA POPUP PILIH NO SJ (Penjualan) / NO PB (Pembelian) / NO SJ (Sparepart)
   ------------------------------------------------------ */
function bukaDraft() {
  const listPenjualan = ambilDaftarSJTransaksi_();
  const listPembelian = ambilDaftarSJPembelian_();
  const listSparepart = ambilDaftarSJSparepart_();
  const listPOMasuk   = ambilDaftarPOMasuk_();

  if (listPenjualan.length === 0 && listPembelian.length === 0 &&
      listSparepart.length === 0 && listPOMasuk.length === 0) {
    SpreadsheetApp.getUi().alert("Belum ada data di sheet TRANSAKSI, PEMBELIAN, SPAREPART, maupun HISTORY_PO_MASUK.");
    return;
  }

  const template = HtmlService.createTemplateFromFile("PopupPilihSJDraft");
  template.itemsJsonPenjualan = JSON.stringify(listPenjualan);
  template.itemsJsonPembelian = JSON.stringify(listPembelian);
  template.itemsJsonSparepart = JSON.stringify(listSparepart);
  template.itemsJsonPOMasuk   = JSON.stringify(listPOMasuk);
  const html = template.evaluate().setWidth(420).setHeight(560);
  SpreadsheetApp.getUi().showModalDialog(html, "Pilih No SJ / No PB / No PO");
}

/* ======================================================
   ========== FLOW PO MASUK (HISTORY_PO_MASUK -> SURAT JALAN, via popup pilih barang) ==========
   ====================================================== */

/* ------------------------------------------------------
   AMBIL DAFTAR NO PO DARI HISTORY_PO_MASUK (untuk isi popup)
   1 NO PO bisa berulang di banyak baris (1 baris = 1 item PO),
   jadi di-dedupe supaya 1 NO PO cuma muncul 1x di daftar.
   HANYA JENIS = "MASUK" yang diambil (HISTORY_PO_MASUK bisa juga
   berisi jenis lain kalau ada, sesuai pola cekTempoTokoBlokir_ dll).
   ------------------------------------------------------ */
function ambilDaftarPOMasuk_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hist = ss.getSheetByName("HISTORY_PO_MASUK");
  if (!hist) throw new Error('Sheet "HISTORY_PO_MASUK" tidak ditemukan.');

  const lastRow = hist.getLastRow();
  if (lastRow < 2) return [];

  const data = hist.getRange(2, 1, lastRow - 1, 12).getValues(); // A..L (butuh D=CUSTOMER, L=STATUS)

  const seen = {};
  const list = [];

  data.forEach(function(row) {
    const jenis = String(row[0] || "").trim().toUpperCase(); // A JENIS
    if (jenis !== "MASUK") return;

    const poNo = String(row[1] || "").trim(); // B NO PO
    if (!poNo || seen[poNo]) return;
    seen[poNo] = true;

    const customer = String(row[3] || "").trim();               // D CUSTOMER
    const status   = String(row[11] || "").trim().toUpperCase(); // L STATUS

    list.push({
      poNo: poNo,
      customer: customer,
      cancelled: status === "CANCELLED"
    });
  });

  // baris paling bawah = data paling baru -> tampilkan paling atas
  list.reverse();
  return list;
}

function terapkanPilihanPOMasukDraft(poNo) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const hist  = ss.getSheetByName("HISTORY_PO_MASUK");
  const masuk = ss.getSheetByName("PO_MASUK");

  if (!hist) throw new Error('Sheet "HISTORY_PO_MASUK" tidak ditemukan.');
  if (!masuk) throw new Error('Sheet "PO_MASUK" tidak ditemukan.');

  poNo = String(poNo || "").trim();
  if (!poNo) throw new Error("No PO kosong.");

  /* ---------- 1) ambil header + item dari HISTORY_PO_MASUK utk No PO ini ---------- */
  const lastRow = hist.getLastRow();
  if (lastRow < 2) throw new Error("Sheet HISTORY_PO_MASUK masih kosong.");

  const data = hist.getRange(2, 1, lastRow - 1, 18).getValues(); // A..R

  let header = null;
  let statusPOAwal = "";
  const items = [];

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const jenis = String(row[0] || "").trim().toUpperCase();
    const rowPoNo = String(row[1] || "").trim();

    if (jenis !== "MASUK" || rowPoNo !== poNo) continue;

    if (!header) {
      header = {
        poNo:     row[1],  // B NO PO
        poTgl:    row[2],  // C TGL PO
        customer: row[3],  // D CUSTOMER
        diskon:   row[4],  // E DISKON ⬅️ TAMBAHAN
        alamat:   row[5],  // F ALAMAT
        telp:     row[6],  // G TELP
        tglKirim: row[7],  // H TGL_KIRIM
        method:   row[17]  // R METHOD
      };
      statusPOAwal = String(row[11] || "").trim().toUpperCase(); // L STATUS baris pertama
    }

    items.push({
      line: row[15] || items.length + 1, // P LINE
      kode: row[8],                       // I KODE_BARANG
      qty:  row[10]                       // K QUANTITY
    });
  }

  if (!header) throw new Error('No PO "' + poNo + '" tidak ditemukan di HISTORY_PO_MASUK.');

  if (statusPOAwal === "CANCELLED") {
    throw new Error(
      'PO "' + poNo + '" sudah di-CANCEL dan tidak bisa di-edit.\n' +
      'Gunakan ACTION "BATAL CANCEL" dulu di HISTORY_PO_MASUK untuk mengaktifkan kembali.'
    );
  }

  items.sort(function(a, b) { return (Number(a.line) || 0) - (Number(b.line) || 0); });

  /* ---------- 2) SET HEADER ke PO_MASUK ---------- */
  masuk.getRange("B2").setValue(header.customer);
  masuk.getRange("B3").setValue(header.alamat);
  masuk.getRange("B4").setValue(header.telp);
  masuk.getRange("E2").setValue(header.poNo);
  masuk.getRange("E3").setValue(header.poTgl);
  masuk.getRange("E4").setValue(header.tglKirim);
  masuk.getRange("H3").setValue(header.diskon || ""); // ⬅️ TAMBAHAN: restore DISKON
  masuk.getRange("H4").setValue(header.method);

  /* ---------- 3) BERSIHKAN HANYA KOLOM B (KODE) DAN D (QTY) ---------- */
  masuk.getRange(PO_ITEM_START_ROW, 2, PO_ITEM_MAX_ROWS, 1).clearContent(); // B KODE
  masuk.getRange(PO_ITEM_START_ROW, 4, PO_ITEM_MAX_ROWS, 1).clearContent(); // D QTY

  /* ---------- 4) ISI ITEM BARU (KODE & QTY SAJA) ---------- */
  const kodes = items.map(function(it) { return [it.kode]; });
  const qtys  = items.map(function(it) { return [it.qty]; });

  if (kodes.length > 0) {
    masuk.getRange(PO_ITEM_START_ROW, 2, kodes.length, 1).setValues(kodes); // Kode (B)
    masuk.getRange(PO_ITEM_START_ROW, 4, qtys.length, 1).setValues(qtys);   // Qty (D)
  }

  return { ok: true, poNo: poNo, jumlahItem: items.length };
}

  // ======================= Kode.js (salinan persis) =======================
function HARGATOKO(harga, namaToko, metode, diskon) {

  // Ambil data sheet DISKON sekali saja (kolom A = nama toko, kolom C = format diskon, misal "35+5+3")
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DISKON");
  const data = sheet.getDataRange().getValues();

  // Cari diskon berdasarkan nama toko (cocok PERSIS / case-sensitive)
  function cariDiskon(toko) {
    const tokoStr = toko.toString().trim();
    for (let i = 0; i < data.length; i++) {
      const namaDiSheet = data[i][0];
      if (namaDiSheet !== "" && namaDiSheet !== null && namaDiSheet.toString().trim() === tokoStr) {
        return data[i][2]; // kolom C (index 2)
      }
    }
    return null;
  }

  // Fungsi untuk menghitung satu baris (satu harga, satu toko, satu metode, satu status display)
  function hitung(h, toko, met, disp) {
    if (h === "" || h == null) return "";
    if (toko === "" || toko == null) return "";

    h = Number(h);

    const dispStr = disp ? disp.toString().trim().toUpperCase() : "";

    // VALIDASI: Jika status display/diskon bertuliskan "SALE", kembalikan diskon 0 (tidak ada diskon)
    if (dispStr === "SALE") {
      return 0;
    }

    if (dispStr === "NODISC") {
      return 0;
    }

    const diskonRaw = cariDiskon(toko);
    if (diskonRaw === null || diskonRaw === "") return "";

    // Ubah string diskon jadi array angka, misal "35+5+3" -> ["35","5","3"]
    const diskonArr = diskonRaw.toString().split("+");

    // Kalau metode CASH, tambahkan 1,5% sebagai faktor tambahan
    if (met && met.toString().trim().toUpperCase() === "CASH") {
      diskonArr.push("1,5");
    }

    // Kalau diskon == DISPLAY, tambahkan 2,5% sebagai faktor tambahan
    if (dispStr === "DISPLAY") {
      diskonArr.push("2,5");
    }

    // Kalau diskon == 3SET, tambahkan 5% sebagai faktor tambahan
    if (dispStr === "3SET") {
      diskonArr.push("5");
    }

    // Kalau diskon == 5SET, tambahkan 7,5% sebagai faktor tambahan
    if (dispStr === "5SET") {
      diskonArr.push("7,5");
    }

    // Hitung sisa harga setelah semua diskon berlaku berurutan (bertingkat/compounding)
    let sisaPersen = 1;
    for (let j = 0; j < diskonArr.length; j++) {
      const d = parseFloat(diskonArr[j].toString().trim().replace(",", "."));
      sisaPersen *= (1 - d / 100);
    }

// Total diskon dalam persen, dipotong (truncate) 4 angka di belakang koma
// ⬅️ FIX: tambah epsilon kecil sebelum Math.round() supaya nilai yang
// SEHARUSNYA persis di titik pembulatan (misal 413375) tidak kebulat ke
// bawah gara-gara floating-point error dari perkalian berantai sisaPersen
// (makin banyak diskon dikombinasikan — CASH/DISPLAY/3SET/5SET —
// makin besar potensi akumulasi galatnya).
    const totalDiskonPersen = Math.round((1 - sisaPersen) * 100 * 10000 + 1e-9) / 10000;

    // Besar nilai diskon (Rupiah) = harga * totalDiskonPersen / 100
    return h * (totalDiskonPersen / 100);
  }

  // Jika ARRAYFORMULA (harga berupa range/array)
  if (Array.isArray(harga)) {
    return harga.map((row, i) => {
      const h = row[0];
      const toko = Array.isArray(namaToko) ? namaToko[i][0] : namaToko;
      const met = Array.isArray(metode) ? metode[i][0] : metode;
      const disp = Array.isArray(diskon) ? diskon[i][0] : diskon;
      return [hitung(h, toko, met, disp)];
    });
  }

  // Jika satu sel
  return hitung(harga, namaToko, metode, diskon);
}

  // ======================= InvoiceGroupingPopup.js (salinan persis) =======================
/* ======================================================
   InvoiceGroupingPopup.gs
   Popup untuk grouping banyak NO SJ (HUTANG) ke 1 NO INV.

   Popup ini HANYA menulis ke kolom A (TRANSAKSI BULAN),
   C (NO INV), G (JATUH TEMPO) untuk baris yang dicentang.
   Kolom lain (SUPPLIER, TOTAL, STATUS, SISA BAYAR, dst)
   tidak disentuh.

   ⬅️ BARU (migrasi database_SANFELO, September 2026):
   - Sheet HUTANG sekarang tinggal di database_SANFELO. Di sisi
     inventory ini HUTANG cuma mirror IMPORTRANGE baca-saja.
   - getSJBelumInvoice() sekarang baca FRESH langsung dari HUTANG asli
     di database lewat bacaSheetGateway_ (bukan mirror lokal yang bisa
     delay), supaya daftar NO SJ yang ditampilin ke popup selalu akurat.
   - simpanGroupingInvoice() sekarang kirim ke gateway (aksi
     "groupingInvoiceHutang"), yang jalan di DatabaseGateway.js sisi
     database_SANFELO: validasi dupe NO INV, tulis A/C/G, & refresh
     STATUS (via updateStatusJatuhTempoGatewayHutang_) semuanya jadi
     terjadi di server terhadap HUTANG yang asli.
   - updateStatusJatuhTempo_() dan refreshSemuaStatusJatuhTempo() yang
     dulu di file ini SUDAH DIHAPUS — logic-nya pindah ke
     updateStatusJatuhTempoGatewayHutang_ / refreshStatusJatuhTempoHutangHarian
     di DatabaseGateway.js (yang terakhir dipasang sebagai trigger
     harian LANGSUNG di Apps Script database_SANFELO, bukan di sini —
     lihat komentar di DatabaseGateway.js). Sebelumnya juga tidak ada
     trigger terpasang buat fungsi ini di sisi inventory, jadi aman
     dihapus.
   - Alur klik admin di popup Grouping HUTANG SAMA SEKALI TIDAK BERUBAH.
   ====================================================== */

const HUTANG_COL = {
  TRANSAKSI_BULAN: 1,
  NO_PB: 2,
  NO_INV: 3,
  NO_SJ: 4,
  SUPPLIER: 5,
  TGL_BELI: 6,
  JATUH_TEMPO: 7,
  TOTAL: 8,
  STATUS: 9,
  SISA_BAYAR: 10
};

function bukaPopupGroupingInvoice() {
  const html = HtmlService.createHtmlOutputFromFile("GroupingInvoicePopup")
    .setWidth(520)
    .setHeight(620);
  SpreadsheetApp.getUi().showModalDialog(html, "Grouping Invoice Hutang");
}

// Tambahkan pemanggilan ini di onOpen() yang sudah ada, misal:
// ui.createMenu("Hutang").addItem("Grouping Invoice", "bukaPopupGroupingInvoice").addToUi();

/* Dipanggil dari popup: ambil semua NO SJ yang belum punya NO INV.
   ⬅️ BARU: baca FRESH dari database lewat bacaSheetGateway_, bukan
   dari sheet HUTANG lokal lagi (itu sekarang cuma mirror). rowIndex
   yang dikembalikan adalah baris di HUTANG ASLI di database. */
function getSJBelumInvoice() {
  const data = bacaSheetGateway_("HUTANG", 2, 1, HUTANG_COL.SISA_BAYAR);
  const list = [];

  data.forEach(function(row, i) {
    const noInv = row[HUTANG_COL.NO_INV - 1];
    const noSJ  = row[HUTANG_COL.NO_SJ - 1];
    if (!noSJ) return;   // baris kosong
    if (noInv) return;   // sudah punya NO INV -> skip

    list.push({
      rowIndex: i + 2,
      noPB: row[HUTANG_COL.NO_PB - 1],
      noSJ: noSJ,
      supplier: row[HUTANG_COL.SUPPLIER - 1],
      tglBeli: formatTglHutang_(row[HUTANG_COL.TGL_BELI - 1]),
      total: Number(row[HUTANG_COL.TOTAL - 1]) || 0,
      sisaBayar: Number(row[HUTANG_COL.SISA_BAYAR - 1]) || 0
    });
  });

  return list;
}

function formatTglHutang_(val) {
  if (val instanceof Date) {
    return Utilities.formatDate(val, "Asia/Jakarta", "dd/MM/yyyy");
  }
  return val ? String(val) : "";
}

/* Dipanggil dari popup pas submit.
   ⬅️ BARU: validasi ringan di sini, lalu SEMUA kerja berat (cek dupe
   NO INV, tulis A/C/G, refresh STATUS) jalan di gateway lewat aksi
   "groupingInvoiceHutang", terhadap HUTANG asli di database. */
function simpanGroupingInvoice(payload) {
  const noInv = String(payload.noInv || "").trim();
  const jatuhTempo = payload.jatuhTempo;
  const transaksiBulan = String(payload.transaksiBulan || "").trim();
  const rows = payload.rows || [];

  if (!noInv) throw new Error("NO INV wajib diisi.");
  if (!rows.length) throw new Error("Pilih minimal 1 NO SJ.");

  return kirimAksiGateway_("groupingInvoiceHutang", {
    rows: rows,
    noInv: noInv,
    jatuhTempo: jatuhTempo,
    transaksiBulan: transaksiBulan
  });
}

  // ======================= PembelianHutangSync.js (salinan persis) =======================
/* ======================================================
   PembelianHutangSync.gs

   FIX (Juli 2026):
   - Key unik grouping SEBELUMNYA pakai NO SJ (kolom C di PEMBELIAN).
     Ternyata NO SJ bisa collision antar PB yang beda (2 PB, 2 supplier
     beda, kebetulan NO SJ-nya sama persis) -> PB kedua ke-merge jadi
     bagian dari PB pertama, gak pernah dapet baris HUTANG sendiri.
   - Sekarang key unik grouping (dan matching ke baris HUTANG yang
     sudah ada) diganti jadi NO PB (kolom A di PEMBELIAN / kolom B di
     HUTANG), sesuai aturan: 1 baris HUTANG = 1 PB. NO SJ tetap dicatat
     di kolom D HUTANG, tapi cuma sebagai informasi, bukan key unik lagi.

   FIX (September 2026) — SISA BAYAR ketinggalan pas TOTAL berubah:
   - Setiap kali sync UPDATE baris yang sudah ada, SISA BAYAR dihitung
     ULANG dari nol: TOTAL terbaru dikurangi TOTAL YANG SUDAH DIBAYAR
     (dihitung dari riwayat pembayaran ASLI, bukan dari angka SISA BAYAR
     lama yang udah gak bisa dipercaya).

   ⬅️ BARU (migrasi database_SANFELO, September 2026):
   - Sheet HUTANG sekarang tinggal di database_SANFELO (cuma IT & owner
     yang bisa akses langsung). Di sisi inventory ini, HUTANG cuma
     mirror IMPORTRANGE baca-saja -- jadi fungsi ini TIDAK PERNAH lagi
     nulis langsung ke sheet HUTANG lokal.
   - PEMBELIAN TETAP LOKAL (belum dimigrasi) -- bagian scan & grouping
     PEMBELIAN di bawah TIDAK BERUBAH SAMA SEKALI. Yang berubah cuma
     bagian akhir: hasil grouping dikirim ke gateway (aksi
     "syncHutangDariPembelian"), yang jalan di DatabaseGateway.js sisi
     database_SANFELO dan melakukan insert/update + recompute SISA
     BAYAR + refresh STATUS terhadap sheet HUTANG YANG ASLI.
   - cariLastRowHutangBerisi_() dan recomputeSisaBayarHutang_() yang
     dulu di sini sudah tidak dipakai lagi (logic-nya pindah ke gateway,
     lihat cariLastRowHutangGatewayBerisi_ dan
     recomputeSisaBayarGatewayHutang_ di DatabaseGateway.js) -- dihapus
     dari sini supaya gak ada 2 salinan yang bisa beda hasil.
   - "Tandain LUNAS + hijau di PEMBELIAN" tetap dipicu dari daftar PB
     yang jadi LUNAS gara-gara koreksi TOTAL (dikasih tahu lewat hasil
     balikan gateway, pbJadiLunasKarenaKoreksi, bukan dihitung sendiri
     di sini) -- TAPI, ⬅️ BARU (migrasi PEMBELIAN, Oktober 2026):
     tandaiLunasPembelianByPB_() sendiri sekarang juga 1x panggilan
     gateway (PEMBELIAN sudah ikut dimigrasi ke database_SANFELO, lihat
     definisi barunya di PembayaranHutang.js).
   - ⬅️ BARU (migrasi PEMBELIAN): bagian scan & grouping PEMBELIAN di
     bawah SEKARANG baca FRESH lewat bacaSheetGateway_, bukan dari sheet
     PEMBELIAN lokal lagi (itu sekarang cuma mirror IMPORTRANGE) --
     WAJIB fresh karena fungsi ini dipanggil LANGSUNG SETELAH
     simpanTandaTerimaKePembelian() nulis baris baru lewat gateway
     (lihat TANDA_TERIMA_BRG.js), jadi gak boleh ketinggalan baca mirror
     yang masih delay (kalau sampai ketinggalan, baris yang baru aja
     disimpan gak ikut ke-grouping ke HUTANG). Logic grouping per NO PB
     sendiri TIDAK BERUBAH SAMA SEKALI, cuma sumber datanya yang beda.
   - Alur klik admin (jalan(), auto-sync abis Tanda Terima Barang) SAMA
     SEKALI TIDAK BERUBAH -- tetap dibungkus try/catch di
     TANDA_TERIMA_BRG.js seperti sebelumnya, jadi kalau gateway lagi
     lambat/timeout, posting Tanda Terima Barang tetap jalan normal.
   ====================================================== */

function jalan(){
  try {
    syncHutangFromPembelian_();
    SpreadsheetApp.getUi().alert("selesai jalan(), cek log/HUTANG");
  } catch (err) {
    SpreadsheetApp.getUi().alert("ERROR: " + err.message + "\n\n" + err.stack);
  }
}

function syncHutangFromPembelian_() {
  // ⬅️ BARU: baca FRESH dari database lewat bacaSheetGateway_ (lihat
  // catatan migrasi PEMBELIAN di komentar atas) -- bukan dari sheet
  // PEMBELIAN lokal lagi.
  const data = bacaSheetGateway_("PEMBELIAN", 2, 1, 13);
  Logger.log("jumlah baris PEMBELIAN (fresh dari gateway): " + data.length);
  if (data.length === 0) return;
  Logger.log("baris pertama PEMBELIAN: " + JSON.stringify(data[0]));

  const groups = {};
  const order = [];

  // Key unik grouping = NO PB (kolom A), bukan NO SJ lagi.
  // NO SJ tetap diambil & disimpan per grup (buat ditulis ke kolom D
  // HUTANG), tapi cuma informasi — bukan penentu baris yang sama/beda.
  data.forEach(function(row) {
    const noPB = String(row[0] || "").trim();
    if (!noPB) return; // lewati baris tanpa NO PB

    const noSJ     = row[2];
    const supplier = row[3];
    const tglBeli  = row[4];
    const total    = Number(row[9]) || 0;

    if (!groups[noPB]) {
      groups[noPB] = { noSJ: noSJ, supplier: supplier, tglBeli: tglBeli, total: 0 };
      order.push(noPB);
    }
    groups[noPB].total += total;
  });

  Logger.log("jumlah grup NO PB ketemu: " + order.length);
  if (order.length === 0) return;

  // ⬅️ BARU: kirim hasil grouping ke gateway, bukan tulis lokal lagi.
  // tglBeli diserialize ke ISO string kalau berupa Date (konsisten sama
  // pola yang sudah dipakai TEMPO TOKO) -- parseTanggalDariGateway_ di
  // sisi gateway yang urus balikin jadi Date lagi.
  const payloadGroups = order.map(function(noPB) {
    const g = groups[noPB];
    return {
      noPB: noPB,
      noSJ: g.noSJ,
      supplier: g.supplier,
      tglBeli: (g.tglBeli instanceof Date) ? g.tglBeli.toISOString() : g.tglBeli,
      total: g.total
    };
  });

  const hasil = kirimAksiGateway_("syncHutangDariPembelian", { groups: payloadGroups });
  Logger.log("hasil syncHutangDariPembelian: " + JSON.stringify(hasil));

  // ⬅️ BARU (Oktober 2026): isi kolom BULAN (A) HUTANG yang masih kosong
  // dari TGL BELI -- baris baru dari sync di atas selalu BULAN-nya kosong,
  // padahal popup Bayar Hutang mencari per BULAN. Butuh aksi gateway
  // isiBulanHutangKosong (database-sanfelo/GatewayHutangBulan.js); kalau
  // belum dipasang, cukup dicatat di log, tidak menggagalkan proses.
  try {
    const hasilBulan = kirimAksiGateway_("isiBulanHutangKosong", {});
    Logger.log("isiBulanHutangKosong: " + JSON.stringify(hasilBulan));
  } catch (err) {
    Logger.log("isiBulanHutangKosong error: " + err);
  }

  // ⬅️ kalau ada PB yang jadi LUNAS gara-gara TOTAL dikoreksi turun (bukan
  // gara-gara ada pembayaran baru), tetap tandain LUNAS + hijau di
  // PEMBELIAN.
  // ⬅️ FIX PERFORMA (Oktober 2026): dulu loop 1x panggilan gateway PER PB
  // di sini -- kalau yang kena koreksi sekaligus banyak (pernah sampai
  // ~94 PB dalam satu sync), itu jadi puluhan-ratusan round-trip jaringan
  // BERURUTAN, bikin simpanTandaTerimaKePembelian() nyangkut bermenit-
  // menit dan sempat kena error 404 dari Google gara-gara request
  // terlalu padat dalam waktu singkat. Sekarang SATU panggilan buat
  // semua PB sekaligus (lihat tandaiLunasPembelianBatch_ di
  // PembayaranHutang.js & aksiTandaiLunasPembelianBatch_ di
  // DatabaseGateway.js).
  const pbUnik = hasil.pbJadiLunasKarenaKoreksi || [];
  try {
    tandaiLunasPembelianBatch_(pbUnik);
  } catch (err) {
    Logger.log("Auto-tandai LUNAS PEMBELIAN batch (dari syncHutangFromPembelian_) error: " + err);
  }
}

  // ======================= QTY_TERKIRIM_ULANG.js (salinan persis) =======================
/* ======================================================
   QTY_TERKIRIM_ULANG.js   (PROJECT: inventory_SANFELO)
   ------------------------------------------------------
   Hitung ULANG kolom H (QTY TERKIRIM) di HISTORY_PO_KELUAR dari PEMBELIAN,
   dari nol, tiap kali dipanggil. Menggantikan cara lama yang menambah /
   mengurangi angka per baris (updateQtyTerkirimHistoryPOKeluar_).

   Kenapa diganti (Oktober 2026):
   1. Barang yang dipecah per warna (KS 001 BIRU 12 / MERAH 12 / COKLAT 4)
      punya NO PO + NAMA BARANG yang sama. Cara lama berhenti di baris
      PERTAMA yang cocok, jadi 12+12+4 = 28 numpuk di baris BIRU dan
      ketiganya gak pernah hijau. Sekarang qty DIBAGI berurutan sesuai
      JUMLAH tiap baris (isi baris 1 sampai penuh, sisanya baris 2, dst).
   2. Kolom A..G di sini salinan IMPORTRANGE dari database, tapi H diisi
      lokal per NOMOR BARIS. Begitu urutan baris di database berubah
      (PO disimpan ulang -> barisnya pindah ke bawah), angka H gak ikut
      pindah dan nempel di barang lain. Sekarang data dibaca FRESH dari
      database (urutannya sama persis dengan salinan) lalu H ditulis
      ulang semua, jadi selalu lurus lagi.
   3. NO PO di PEMBELIAN bisa gabungan ("#2026-052, #2026-054") atau
      salah ketik nomor PO-nya. Pencocokan sekarang bertingkat:
        a. NO PO (salah satu dari daftar) + NAMA BARANG + NAMA TOKO
        b. NO PO + NAMA BARANG
        c. NAMA BARANG + NAMA TOKO di PO lain (NO PO di PEMBELIAN salah)
           -> tetap dihitung, tapi dilaporkan supaya PEMBELIAN-nya dibetulkan
      Yang sama sekali gak ketemu juga dilaporkan.

   Dipanggil otomatis setelah Simpan SJ Beli / hapus PEMBELIAN / simpan
   PO KELUAR, dan manual lewat menu 🏢 SANFELO > 🔁 Hitung Ulang QTY Terkirim.
   Catatan: angka H yang diketik manual akan ikut ditimpa hasil hitungan.
   ====================================================== */

function hitungUlangQtyTerkirim() {
  hitungUlangQtyTerkirimPOKeluar_(true);
}

function hitungUlangQtyTerkirimPOKeluar_(tampilkanHasil) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hist = ss.getSheetByName('HISTORY_PO_KELUAR');
  if (!hist) throw new Error('Sheet "HISTORY_PO_KELUAR" tidak ditemukan.');

  // FRESH dari database -> baris ke-i di sini = baris ke-i salinan IMPORTRANGE lokal
  const histDb = bacaSheetGateway_('HISTORY_PO_KELUAR', 2, 1, 7);  // A..G
  const pmbDb = bacaSheetGateway_('PEMBELIAN', 2, 1, 11);          // A..K
  const hasil = alokasiQtyTerkirim_(histDb, pmbDb);

  const n = Math.max(histDb.length, 1);
  const kolomH = hist.getRange(2, 8, n, 1);
  const lama = kolomH.getValues();
  const baru = [];
  let berubah = 0;
  for (let i = 0; i < n; i++) {
    const v = i < hasil.qty.length && hasil.qty[i] > 0 ? hasil.qty[i] : '';
    if ((Number(lama[i][0]) || 0) !== (Number(v) || 0)) berubah++;
    baru.push([v]);
  }
  kolomH.setValues(baru);
  // sisa baris di bawah data (sisa salinan lama yang sudah bergeser) dikosongkan
  const sisa = hist.getMaxRows() - 1 - n;
  if (sisa > 0) hist.getRange(2 + n, 8, sisa, 1).clearContent();

  // ⬅️ BARU: kirim juga ke kolom H HISTORY_PO_KELUAR di DATABASE (di sana H
  // terpisah dari salinan ini). Butuh aksi gateway setQtyTerkirimPoKeluarBatch
  // (database-sanfelo/GatewayQtyTerkirim.js). Kalau gagal, tampilan di
  // inventory tetap benar -- cuma database yang belum ikut.
  let pesanDb = '';
  try {
    const hasilDb = kirimAksiGateway_('setQtyTerkirimPoKeluarBatch', {
      rows: histDb.map(function (r, i) {
        return { po: r[0], barang: r[3], ket: r[5], qty: hasil.qty[i] || 0 };
      })
    });
    pesanDb = '\nDatabase ikut diperbarui (' + (hasilDb.baris_diubah || 0) + ' baris).';
  } catch (err) {
    Logger.log('Kirim QTY TERKIRIM ke database gagal: ' + err);
    pesanDb = '\n⚠️ Database BELUM ikut diperbarui: ' + err.message +
      (String(err.message).indexOf('tidak dikenali') !== -1
        ? '\n(Pasang GatewayQtyTerkirim.js di project database_SANFELO lalu deploy versi baru.)' : '');
  }

  if (tampilkanHasil) {
    let pesan = '✅ QTY TERKIRIM dihitung ulang dari PEMBELIAN: ' + berubah + ' baris berubah.' + pesanDb;
    if (hasil.bedaPO.length) {
      pesan += '\n\n⚠️ NO PO di PEMBELIAN beda dengan PO keluar-nya (sudah dihitung lewat nama toko, ' +
        'tapi sebaiknya NO PO di PEMBELIAN dibetulkan):\n' + hasil.bedaPO.slice(0, 15).join('\n') +
        (hasil.bedaPO.length > 15 ? '\n... dan ' + (hasil.bedaPO.length - 15) + ' lainnya' : '');
    }
    if (hasil.lebih.length) {
      pesan += '\n\n⚠️ Dibeli LEBIH dari jumlah PO:\n' + hasil.lebih.slice(0, 15).join('\n');
    }
    if (hasil.tidakKetemu.length) {
      pesan += '\n\n❓ Ada di PEMBELIAN tapi tidak ada PO keluar-nya sama sekali:\n' + hasil.tidakKetemu.slice(0, 15).join('\n') +
        (hasil.tidakKetemu.length > 15 ? '\n... dan ' + (hasil.tidakKetemu.length - 15) + ' lainnya' : '');
    }
    SpreadsheetApp.getUi().alert(pesan);
  }
  return hasil;
}

/**
 * Inti perhitungan (tanpa baca/tulis sheet).
 * histRows : HISTORY_PO_KELUAR A..G  (A NO PO, C NAMA TOKO, D NAMA BARANG, E JUMLAH)
 * pmbRows  : PEMBELIAN A..K          (A NO PB, B NO PO, H NAMA BARANG, I QTY, K KETERANGAN=toko)
 * Return { qty: [per baris hist], bedaPO: [], lebih: [], tidakKetemu: [] }
 */
function alokasiQtyTerkirim_(histRows, pmbRows) {
  const U = function (s) { return String(s === null || s === undefined ? '' : s).trim().toUpperCase().replace(/\s+/g, ' '); };
  const qty = histRows.map(function () { return 0; });
  const perNama = {};
  histRows.forEach(function (r, i) {
    if (!U(r[0]) || !U(r[3])) return;
    (perNama[U(r[3])] = perNama[U(r[3])] || []).push(i);
  });
  const bedaPO = [], lebih = [], tidakKetemu = [];

  pmbRows.forEach(function (p) {
    const q = Number(p[8]) || 0;
    const nama = U(p[7]);
    if (!q || !nama) return;
    const poList = String(p[1] || '').split(',').map(U).filter(String);
    const toko = U(p[10]);
    const semua = perNama[nama] || [];
    const samaPO = semua.filter(function (i) { return poList.indexOf(U(histRows[i][0])) !== -1; });
    const samaToko = function (list) { return list.filter(function (i) { return toko && U(histRows[i][2]) === toko; }); };

    let pilihan = samaToko(samaPO);
    if (!pilihan.length) pilihan = samaPO;
    if (!pilihan.length) {
      pilihan = samaToko(semua);
      if (pilihan.length) bedaPO.push(p[0] + ' ' + p[7] + ' (' + p[10] + '): dicatat ' + (p[1] || '-') + ', PO keluar-nya ' +
        pilihan.map(function (i) { return histRows[i][0]; }).filter(function (v, k, a) { return a.indexOf(v) === k; }).join('/'));
    }
    if (!pilihan.length) {
      tidakKetemu.push(p[0] + ' ' + (p[1] || '-') + ' ' + p[7] + ' (' + q + ', ' + (p[10] || '-') + ')');
      return;
    }

    // isi berurutan sesuai JUMLAH tiap baris; kelebihan masuk ke baris terakhir
    let sisa = q;
    for (let k = 0; k < pilihan.length && sisa > 0; k++) {
      const i = pilihan[k];
      const ruang = Math.max(0, (Number(histRows[i][4]) || 0) - qty[i]);
      const ambil = Math.min(ruang, sisa);
      qty[i] += ambil;
      sisa -= ambil;
    }
    if (sisa > 0) {
      const akhir = pilihan[pilihan.length - 1];
      qty[akhir] += sisa;
      lebih.push(histRows[akhir][0] + ' ' + histRows[akhir][3] + ' (' + histRows[akhir][2] + '): lebih ' + sisa + ' dari ' + p[0]);
    }
  });

  return { qty: qty, bedaPO: bedaPO, lebih: lebih, tidakKetemu: tidakKetemu };
}


  // ======================= TANDA_TERIMA_BRG.js (salinan persis) =======================
/* ======================================================
   TANDA_TERIMA_BRG.gs
   - simpanTandaTerimaKePembelian() : simpan data dari sheet TANDA TERIMA BRG
     ke sheet PEMBELIAN, dengan cek duplikat NO SJ
   - NO PB auto-generate: liat blok "NO PB" di bawah

   (Riwayat UPDATE 1-5 + migrasi database_SANFELO: lihat versi sebelumnya.)

   UPDATE 6 (MULTI NO PO):
   - Kolom B (NO PO) di PEMBELIAN menyimpan SEMUA no PO yang ditulis user
     di B3 (mis. "PO-1, PO-2") di setiap baris.
   - NO PO per item (poEfektif) tetap dihitung, tapi HANYA dipakai untuk
     update QTY TERKIRIM di HISTORY_PO_KELUAR.
   - hapusBarisPembelian_ memecah NO PO gabungan & mencari PO yang tepat
     per item sebelum decrement QTY TERKIRIM.

   UPDATE 7 (AUTO CLEAR SETELAH SIMPAN):
   - Setelah tersimpan ke PEMBELIAN, form TANDA TERIMA BRG PASTI dikosongkan
     (header, item B:C, F:H, dan kolom E harga), lalu flush() sebelum alert.
   - Update HISTORY_PO_KELUAR & sync HUTANG dibungkus try/catch supaya error
     di situ tidak menghalangi reset form.
   ====================================================== */

const TANDA_TERIMA_ITEM_START_ROW = 7;   // baris pertama item di TANDA TERIMA BRG
const TANDA_TERIMA_ITEM_MAX_ROWS  = 100; // batas scan/clear baris item

// Kolom A7 & D7 isinya ARRAYFORMULA (anchor formula ada di baris 7, nge-spill ke bawah).
// KEDUA kolom ini TIDAK BOLEH di-clearContent atau di-setValue manual per baris.
// Yang boleh disentuh manual: B, C, E (harga, ditulis script), F, G, H.
// F = NAMA TOKO ASAL per item, G = NO PO ASAL per item (hidden),
// H = STATUS DISKON SET per item (hidden).
const TANDA_TERIMA_LAST_ROW = TANDA_TERIMA_ITEM_START_ROW + TANDA_TERIMA_ITEM_MAX_ROWS - 1; // 106

const TANDA_TERIMA_COL_SET = 8; // H

const TANDA_TERIMA_FORMULA_NO   = '=ARRAYFORMULA(IF(B7:B="";"";SCAN(0;B7:B;LAMBDA(a;b;IF(b="";a;a+1)))))';

// PENTING: sengaja TIDAK pakai AND(...) di dalam ARRAYFORMULA.
const TANDA_TERIMA_FORMULA_NAMA =
  '=ARRAYFORMULA(IF(COUNTA($B$7:$B$' + TANDA_TERIMA_LAST_ROW + ')=0;' +
  'IF(B7:B' + TANDA_TERIMA_LAST_ROW + '="";"";IFERROR(VLOOKUP(B7:B' + TANDA_TERIMA_LAST_ROW + ';DATA_BERSIH!A:B;2;FALSE);""));' +
  'IF(ROW(B7:B' + TANDA_TERIMA_LAST_ROW + ')=7+COUNTA($B$7:$B$' + TANDA_TERIMA_LAST_ROW + ');' +
  '"Total";' +
  'IF(B7:B' + TANDA_TERIMA_LAST_ROW + '="";"";IFERROR(VLOOKUP(B7:B' + TANDA_TERIMA_LAST_ROW + ';DATA_BERSIH!A:B;2;FALSE);""))' +
  ')))';

/**
 * Persen potongan HPP berdasarkan status SET per item.
 */
function persenDiskonSetTandaTerima_(status) {
  const s = String(status || "").trim().toUpperCase();
  if (s === "3SET") return 0.05;
  if (s === "5SET") return 0.075;
  return 0;
}

/**
 * Bangun map { KODE: HPP } dari sheet DATA_BERSIH kolom K (index 11).
 */
function buildHppMapDariDataBersih_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const db = ss.getSheetByName("DATA_BERSIH");
  const map = {};
  if (!db) return map;

  const lastRow = db.getLastRow();
  if (lastRow < 1) return map;

  // ⬅️ FIX (Nov 2026): dulu baca blok A..K sekaligus. Kolom I (BARCODE) berisi
  // gambar, jadi blok itu ikut membawa objek gambar. Sekarang cukup baca kolom
  // A (KODE) dan K (HPP) saja, masing-masing terpisah.
  const kodeCol = db.getRange(1, 1, lastRow, 1).getValues();   // A
  const hppCol  = db.getRange(1, 11, lastRow, 1).getValues();  // K
  for (let i = 0; i < kodeCol.length; i++) {
    const kode = String(kodeCol[i][0] || "").trim();
    if (!kode) continue;
    map[kode] = Number(hppCol[i][0]) || 0;
  }

  return map;
}

/**
 * Hitung ULANG & TULIS kolom E (HARGA) di TANDA TERIMA BRG.
 * Return jumlah item.
 */
function hitungUlangHargaTandaTerima_(shTerima) {
  const startRow = TANDA_TERIMA_ITEM_START_ROW;
  const numRows = TANDA_TERIMA_ITEM_MAX_ROWS;

  // Blok B..H: index 0=B(kode), 1=C(qty), 6=H(set)
  const block = shTerima.getRange(startRow, 2, numRows, 7).getValues();

  const hppMap = buildHppMapDariDataBersih_();

  let jumlahItem = 0;
  const hargaItem = [];

  for (let i = 0; i < numRows; i++) {
    const kode = String(block[i][0] || "").trim();
    if (kode === "") {
      hargaItem.push("");
      continue;
    }
    jumlahItem++;
    const qty = Number(block[i][1]) || 0;
    const hppSatuan = hppMap[kode] || 0;
    const persenDiskon = persenDiskonSetTandaTerima_(block[i][6]);
    hargaItem.push(hppSatuan * (1 - persenDiskon) * qty);
  }

  const outValues = [];
  for (let i = 0; i < numRows; i++) {
    if (jumlahItem > 0 && i === jumlahItem) {
      let total = 0;
      for (let j = 0; j < jumlahItem; j++) total += (Number(hargaItem[j]) || 0);
      outValues.push([total]);
    } else if (typeof hargaItem[i] === "number") {
      outValues.push([hargaItem[i]]);
    } else {
      outValues.push([""]);
    }
  }

  const rangeHarga = shTerima.getRange(startRow, 5, numRows, 1); // E
  rangeHarga.setValues(outValues);
  rangeHarga.setNumberFormat("#,##0");

  return jumlahItem;
}

/**
 * Formula anchor A7 & D7 cuma ditulis kalau kosong/hilang (atau force=true).
 */
function pastikanFormulaTandaTerimaUtuh_(shTerima, force) {
  const a7 = shTerima.getRange("A7");
  const d7 = shTerima.getRange("D7");
  if (force || !a7.getFormula()) a7.setFormula(TANDA_TERIMA_FORMULA_NO);
  if (force || !d7.getFormula()) d7.setFormula(TANDA_TERIMA_FORMULA_NAMA);

  // ⬅️ FIX (Nov 2026): kalau hitung harga error, dulu seluruh proses ikut gagal
  // diam-diam (kolom HARGA tetap kosong, tanpa pesan). Sekarang error-nya
  // ditampilkan sebagai pemberitahuan di layar, dan proses lain tetap jalan.
  let jumlahItem = 0;
  try {
    jumlahItem = hitungUlangHargaTandaTerima_(shTerima);
  } catch (err) {
    Logger.log("Hitung harga Tanda Terima error: " + err);
    try {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        "Harga gagal dihitung: " + (err && err.message ? err.message : err) + ". Klik menu REFRESH BOS untuk mencoba lagi.",
        "⚠️ Tanda Terima Barang", 20);
    } catch (e2) { /* abaikan */ }
  }
  terapkanBoldTotalTandaTerima_(shTerima, jumlahItem);
}

function terapkanBoldTotalTandaTerima_(shTerima, jumlahItemOpt) {
  const startRow = TANDA_TERIMA_ITEM_START_ROW;
  const lastItemRow = TANDA_TERIMA_LAST_ROW;
  const maxRowSheet = shTerima.getMaxRows();
  const totalRowMax = Math.min(lastItemRow + 1, maxRowSheet);

  if (totalRowMax < startRow) return;

  let jumlahItem = jumlahItemOpt;
  if (typeof jumlahItem !== "number") {
    const kodeValues = shTerima.getRange(startRow, 2, lastItemRow - startRow + 1, 1).getValues();
    jumlahItem = 0;
    for (let i = 0; i < kodeValues.length; i++) {
      if (String(kodeValues[i][0] || "").trim() !== "") jumlahItem++;
    }
  }

  shTerima.getRange(startRow, 4, (totalRowMax - startRow + 1), 2).setFontWeight("normal"); // D:E

  if (jumlahItem > 0) {
    const totalRow = startRow + jumlahItem;
    if (totalRow <= maxRowSheet) {
      shTerima.getRange(totalRow, 4, 1, 2).setFontWeight("bold");
    }
  }
}

// Menu: refresh formula/harga di sheet yang sudah kadung dipakai.
function refreshFormulaTandaTerima() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shTerima = ss.getSheetByName("TANDA TERIMA BRG");
  const ui = SpreadsheetApp.getUi();
  if (!shTerima) { ui.alert('Sheet "TANDA TERIMA BRG" tidak ditemukan.'); return; }

  pastikanFormulaTandaTerimaUtuh_(shTerima, true);
  ui.alert("✅ Formula NO/NAMA, HARGA (termasuk diskon SET), dan baris Total di TANDA TERIMA BRG sudah di-refresh.");
}

/* ======================================================
   HELPER MULTI NO PO
   ====================================================== */

// "PO-1, PO-2,PO-1" -> ["PO-1","PO-2"] (trim, buang kosong & duplikat)
function parsePoListTT_(s) {
  const seen = {};
  return String(s || "").split(",")
    .map(function(p) { return p.trim(); })
    .filter(function(p) {
      if (!p) return false;
      const k = p.toLowerCase();
      if (seen[k]) return false;
      seen[k] = true;
      return true;
    });
}

// Baca HISTORY_PO_KELUAR (A..H) sekali. Return [] kalau sheet gak ada / kosong.
function bacaHistoryPOKeluarData_(ss) {
  const hist = ss.getSheetByName("HISTORY_PO_KELUAR");
  if (!hist) return [];
  const lastRow = hist.getLastRow();
  if (lastRow < 2) return [];
  return hist.getRange(2, 1, lastRow - 1, 8).getValues();
}

/**
 * Cari PO yang tepat untuk 1 item di antara poList, berdasarkan HISTORY_PO_KELUAR
 * (match NAMA BARANG, dan NAMA TOKO kalau ada).
 * mode "inc": pilih PO pertama yang sisanya (jumlah - terkirim) >= qty, kalau
 *             gak ada pilih kandidat pertama.
 * mode "dec": pilih PO pertama yang qty terkirimnya > 0, kalau gak ada pilih
 *             kandidat pertama.
 * Return { po, ambiguous } -- ambiguous = barang ditemukan di >1 PO.
 */
function cariPoItemDiHistory_(histData, poList, namaBarang, namaToko, mode, qty) {
  const namaUpper = String(namaBarang || "").trim().toUpperCase();
  const tokoUpper = String(namaToko || "").trim().toUpperCase();
  const q = Number(qty) || 0;
  const candidates = [];

  poList.forEach(function(po) {
    for (let i = 0; i < histData.length; i++) {
      const rowNoPO = String(histData[i][0] || "").trim();
      const rowToko = String(histData[i][2] || "").trim().toUpperCase();
      const rowNama = String(histData[i][3] || "").trim().toUpperCase();
      if (rowNoPO !== po || rowNama !== namaUpper) continue;
      if (tokoUpper && rowToko !== tokoUpper) continue;
      candidates.push({
        po: po,
        jumlah: Number(histData[i][4]) || 0,
        terkirim: Number(histData[i][7]) || 0
      });
      break; // 1 kandidat per PO
    }
  });

  if (candidates.length === 0) return { po: poList[0] || "", ambiguous: false };

  let chosen = null;
  for (let c = 0; c < candidates.length; c++) {
    if (mode === "dec") {
      if (candidates[c].terkirim > 0) { chosen = candidates[c]; break; }
    } else {
      if ((candidates[c].jumlah - candidates[c].terkirim) >= q) { chosen = candidates[c]; break; }
    }
  }
  if (!chosen) chosen = candidates[0];

  return { po: chosen.po, ambiguous: candidates.length > 1 };
}

/* ======================================================
   NO PB (Nomor Pembelian) - auto counter bulanan
   Disimpan di sheet SYSTEM: A5="LAST_PB", B5=angka terakhir, C5=periode (YYYY-MM)
   Format akhir: PB-2026-07-001
   ====================================================== */

const SYSTEM_PB_ROW = 5;

function getPeriodeSekarangPB_() {
  const tz = Session.getScriptTimeZone();
  return Utilities.formatDate(new Date(), tz, "yyyy-MM");
}

function formatNoPB_(periode, angka) {
  return "PB-" + periode + "-" + String(angka).padStart(3, "0");
}

function cariNoPBTerakhirDiPembelian_(periode) {
  const data = bacaSheetGateway_("PEMBELIAN", 2, 1, 1); // kolom A = NO PB
  const prefix = "PB-" + periode + "-";
  let maxAngka = 0;

  data.forEach(function(row) {
    const val = String(row[0] || "").trim();
    if (val.indexOf(prefix) === 0) {
      const angka = parseInt(val.substring(prefix.length), 10);
      if (!isNaN(angka) && angka > maxAngka) maxAngka = angka;
    }
  });

  return maxAngka;
}

function generateNoPBBerikutnya_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const system = ss.getSheetByName("SYSTEM");
  if (!system) throw new Error('Sheet "SYSTEM" tidak ditemukan.');

  const periodeSekarang = getPeriodeSekarangPB_();
  const sysRow = system.getRange(SYSTEM_PB_ROW, 1, 1, 3).getValues()[0];
  const lastPeriodeTersimpan = String(sysRow[2] || "").trim(); // C5
  const lastAngkaTersimpan   = Number(sysRow[1]) || 0;          // B5
  const angkaDiPembelian     = cariNoPBTerakhirDiPembelian_(periodeSekarang);

  let angkaTerakhir;
  if (lastPeriodeTersimpan === periodeSekarang) {
    angkaTerakhir = Math.max(lastAngkaTersimpan, angkaDiPembelian);
  } else {
    angkaTerakhir = angkaDiPembelian;
  }

  const angkaBaru = angkaTerakhir + 1;

  system.getRange(SYSTEM_PB_ROW, 1, 1, 3).setValues([["LAST_PB", angkaBaru, periodeSekarang]]);

  return formatNoPB_(periodeSekarang, angkaBaru);
}

function isiNoPBBerikutnyaKeForm_(shTerima) {
  const noPBBaru = generateNoPBBerikutnya_();
  shTerima.getRange("B4").setValue(noPBBaru);
  return noPBBaru;
}

function cekDanIsiNoPBAwal_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shTerima = ss.getSheetByName("TANDA TERIMA BRG");
  if (!shTerima) return;

  const b4 = String(shTerima.getRange("B4").getValue() || "").trim();
  if (b4) return;

  try {
    isiNoPBBerikutnyaKeForm_(shTerima);
  } catch (err) {
    Logger.log("Gagal auto-generate NO PB awal: " + err);
  }
}

/* ======================================================
   SIMPAN KE PEMBELIAN
   ====================================================== */

function simpanTandaTerimaKePembelian() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shTerima = ss.getSheetByName("TANDA TERIMA BRG");
  const pembelian = ss.getSheetByName("PEMBELIAN");
  const ui = SpreadsheetApp.getUi();

  if (!shTerima) throw new Error('Sheet "TANDA TERIMA BRG" tidak ditemukan.');
  if (!pembelian) throw new Error('Sheet "PEMBELIAN" tidak ditemukan.');

  const hdrB = shTerima.getRange("B2:B4").getValues(); // B2 supplier, B3 noPO, B4 noPB
  const hdrE = shTerima.getRange("E2:E4").getValues(); // E2 noSJ, E3 tglBeli, E4 noKend
  const hdrG = shTerima.getRange("G2:G3").getValues(); // G2 penerima, G3 pengirim

  const noPB     = String(hdrB[2][0] || "").trim();
  const supplier = hdrB[0][0];
  const noPO     = hdrB[1][0]; // bisa gabungan beberapa PO dipisah koma
  const noSJ     = hdrE[0][0];
  const tglBeli  = hdrE[1][0];
  const noKend   = hdrE[2][0];
  const penerima = hdrG[0][0];
  const pengirim = hdrG[1][0];

  const noSJTrim = String(noSJ || "").trim();
  if (!noSJTrim) {
    ui.alert("NO SJ (E2) masih kosong. Isi dulu sebelum disimpan.");
    return;
  }

  if (!noPB) {
    ui.alert("NO PB (B4) masih kosong. Ini harusnya keisi otomatis — cek sheet SYSTEM, atau tutup-buka lagi sheet TANDA TERIMA BRG.");
    return;
  }

  // --- SCAN ITEM (1x getValues blok B:H) ---
  // index: 0=B kode, 1=C qty, 2=D nama, 3=E harga, 4=F toko, 5=G noPO asal, 6=H status SET
  const itemBlock = shTerima.getRange(TANDA_TERIMA_ITEM_START_ROW, 2, TANDA_TERIMA_ITEM_MAX_ROWS, 7).getValues();
  const items = [];
  for (let i = 0; i < itemBlock.length; i++) {
    const row = itemBlock[i];
    const kode      = row[0];
    const qty       = row[1];
    const nama      = row[2];
    const harga     = row[3];
    const tokoAsal  = String(row[4] || "").trim();
    const poAsal    = String(row[5] || "").trim();
    const setDiskon = String(row[6] || "").trim().toUpperCase();

    const kodeKosong  = String(kode || "").trim() === "";
    const qtyKosong   = String(qty || "").trim() === "";
    const tokoKosong  = tokoAsal === "";

    if (kodeKosong && qtyKosong && tokoKosong) continue;

    items.push({ kode: kode, qty: qty, nama: nama, harga: harga, tokoAsal: tokoAsal, poAsal: poAsal, setDiskon: setDiskon });
  }

  if (items.length === 0) {
    items.push({ kode: "", qty: "", nama: "", harga: "", tokoAsal: "", poAsal: "", setDiskon: "" });
  }

  // --- DAFTAR PO DARI HEADER (B3) ---
  const headerPOList = parsePoListTT_(noPO);

  // --- CEK DUPLIKAT NO SJ DI PEMBELIAN (NO SJ + SUPPLIER sama) ---
  const existing = cariBarisPembelianByNoSJ_(noSJTrim, supplier);

  if (existing.length > 0) {
    const first = existing[0];
    const daftarItem = existing.map(function(e, idx) {
      return (idx + 1) + ". " + e.kode + " - " + e.nama + " - qty " + e.qty;
    }).join("\n");

    const resp = ui.alert(
      "NO SJ Sudah Ada",
      'NO SJ "' + noSJTrim + '" untuk supplier "' + first.supplier + '" sudah pernah disimpan.\n\n' +
      "NO PB: " + first.noPB + "\n" +
      "NO PO: " + first.noPO + "\n" +
      "Supplier: " + first.supplier + "\n" +
      "Tgl Beli: " + first.tglBeli + "\n\n" +
      "Item tersimpan:\n" + daftarItem + "\n\n" +
      "YES = Timpa (hapus data lama, simpan yang baru)\n" +
      "NO = Edit (ambil data lama ke form ini untuk diedit)\n" +
      "CANCEL = Batal",
      ui.ButtonSet.YES_NO_CANCEL
    );

    if (resp === ui.Button.CANCEL || resp === ui.Button.CLOSE) return;

    if (resp === ui.Button.NO) {
      // EDIT: tarik data lama balik ke form (termasuk NO PB lama), hapus baris lama
      hapusBarisPembelian_(pembelian, existing);
      isiFormTandaTerimaDariExisting_(shTerima, first, existing);

      try {
        syncHutangFromPembelian_();
      } catch (err) {
        Logger.log("Sync HUTANG error (edit flow): " + err);
      }

      SpreadsheetApp.flush();
      ui.alert("Data lama sudah dimuat balik ke form. Silakan edit lalu simpan ulang.");
      return;
    }

    if (resp === ui.Button.YES) {
      // skipGateway=true: delete-by-NO SJ+SUPPLIER sudah dikerjakan
      // oleh "syncPembelian" di bawah, QTY TERKIRIM dihitung ulang sesudahnya.
      hapusBarisPembelian_(pembelian, existing, true);
    }
  }

  // --- TENTUKAN NO PO EFEKTIF PER ITEM (dipakai HANYA utk QTY TERKIRIM) ---
  const fallbackAmbiguous = [];
  const perluCariHistory = headerPOList.length > 1 && items.some(function(it) { return !it.poAsal; });
  const histData = perluCariHistory ? bacaHistoryPOKeluarData_(ss) : [];

  items.forEach(function(it) {
    if (it.poAsal) {
      it.poEfektif = it.poAsal;
    } else if (headerPOList.length <= 1) {
      it.poEfektif = headerPOList[0] || "";
    } else {
      const hasil = cariPoItemDiHistory_(histData, headerPOList, it.nama || it.kode, it.tokoAsal, "inc", it.qty);
      it.poEfektif = hasil.po;
      if (hasil.ambiguous && String(it.nama || it.kode || "").trim() !== "") {
        fallbackAmbiguous.push(String(it.nama || it.kode));
      }
    }
  });

  // --- NO PO YANG DITULIS KE PEMBELIAN: SEMUA PO DI B3 ---
  let noPoGabungan;
  if (headerPOList.length > 0) {
    noPoGabungan = headerPOList.join(", ");
  } else {
    const poAsalList = [];
    items.forEach(function(it) { if (it.poAsal) poAsalList.push(it.poAsal); });
    noPoGabungan = parsePoListTT_(poAsalList.join(",")).join(", ");
  }

  // --- TULIS KE PEMBELIAN (15 kolom A..O) ---
  const rows = items.map(function(it) {
    return [
      noPB,          // A NO PB
      noPoGabungan,  // B NO PO (SEMUA PO dari B3)
      noSJTrim,      // C NO SJ
      supplier,      // D SUPPLIER
      tglBeli,       // E TGL BELI
      noKend,        // F NO KEND
      it.kode,       // G KODE
      it.nama,       // H NAMA BARANG
      it.qty,        // I QTY
      it.harga,      // J HARGA (sudah termasuk diskon SET)
      it.tokoAsal,   // K NAMA TOKO
      penerima,      // L PENERIMA
      pengirim,      // M PENGIRIM
      "",            // N reserved
      it.setDiskon   // O STATUS DISKON SET
    ];
  });

  const rowsPayload = rows.map(function(row) {
    return row.map(function(val) {
      return (val instanceof Date) ? val.toISOString() : val;
    });
  });

  // Kalau baris ini error, script berhenti & form TIDAK dihapus (data aman).
  kirimAksiGateway_("syncPembelian", { noSJ: noSJTrim, supplier: supplier, rows: rowsPayload });

  // ---- Mulai sini PEMBELIAN sudah tersimpan. Error di langkah berikut
  // ---- tidak boleh menghalangi pengosongan form.
  const peringatanTambahan = [];

  // ⬅️ FIX (Oktober 2026): QTY TERKIRIM di HISTORY_PO_KELUAR sekarang
  // DIHITUNG ULANG penuh dari PEMBELIAN (QTY_TERKIRIM_ULANG.js), bukan
  // ditambah per baris lagi -- cara lama numpuk qty barang yang dipecah per
  // warna di baris pertama & nyasar kalau urutan baris di database berubah.
  try {
    hitungUlangQtyTerkirimPOKeluar_(false);
  } catch (err) {
    Logger.log("Update QTY TERKIRIM error: " + err);
    peringatanTambahan.push("QTY TERKIRIM di HISTORY_PO_KELUAR gagal diupdate (" + err + "). Cek manual.");
  }

  try {
    syncHutangFromPembelian_();
  } catch (err) {
    Logger.log("Sync HUTANG error: " + err);
    peringatanTambahan.push("Sync HUTANG gagal (" + err + "). Jalankan sync HUTANG manual.");
  }

  // ---- KOSONGKAN FORM TANDA TERIMA BRG ----
  resetFormTandaTerima_(shTerima);

  // langsung generate NO PB berikutnya abis form direset
  try {
    cekDanIsiNoPBAwal_();
  } catch (err) {
    Logger.log("Gagal generate NO PB berikutnya: " + err);
  }

  // pastikan semua perubahan sheet sudah tampil SEBELUM alert muncul
  SpreadsheetApp.flush();

  let pesan;
  const adaItem = items.some(function(it) { return String(it.kode || "").trim() !== ""; });
  if (adaItem) {
    pesan = rows.length + " item berhasil disimpan ke PEMBELIAN (NO PO: " + (noPoGabungan || "-") + "). Form sudah dikosongkan.";
  } else {
    pesan = "Data header berhasil disimpan ke PEMBELIAN (tanpa item, KODE kosong). Form sudah dikosongkan.";
  }

  if (fallbackAmbiguous.length > 0) {
    pesan += "\n\nPERHATIAN: " + fallbackAmbiguous.length + " item (" + fallbackAmbiguous.join(", ") + ") " +
      "diinput manual dan barangnya ada di lebih dari satu PO di B3. " +
      "QTY TERKIRIM-nya otomatis dicatat ke PO yang sisanya masih cukup — " +
      "tolong cek HISTORY_PO_KELUAR kalau itu bukan PO yang seharusnya.";
  }

  if (peringatanTambahan.length > 0) {
    pesan += "\n\nPERHATIAN:\n- " + peringatanTambahan.join("\n- ");
  }

  ui.alert(pesan);
}

// Cari baris di PEMBELIAN dengan NO SJ (kolom C) DAN SUPPLIER (kolom D) yang sama persis.
function cariBarisPembelianByNoSJ_(noSJ, supplier) {
  const data = bacaSheetGateway_("PEMBELIAN", 2, 1, 15); // A..O
  const supplierTrim = String(supplier || "").trim().toUpperCase();
  const result = [];

  data.forEach(function(row) {
    const rowNoSJ = String(row[2] || "").trim();
    const rowSupplier = String(row[3] || "").trim().toUpperCase();

    if (rowNoSJ === noSJ && rowSupplier === supplierTrim) {
      result.push({
        noPB: row[0],
        noPO: row[1],
        noSJ: row[2],
        supplier: row[3],
        tglBeli: row[4],
        noKend: row[5],
        kode: row[6],
        nama: row[7],
        qty: row[8],
        harga: row[9],
        namaToko: row[10],
        penerima: row[11],
        pengirim: row[12],
        setDiskon: row[14]
      });
    }
  });

  return result;
}

// Hapus baris PEMBELIAN lewat gateway (rows kosong = hapus doang).
// Decrement QTY TERKIRIM di HISTORY_PO_KELUAR tetap lokal. NO PO di baris
// lama bisa gabungan ("PO-1, PO-2") -> dipecah & dicari PO yang tepat per item.
function hapusBarisPembelian_(pembelian, existingRows, skipGateway) {
  // ⬅️ FIX (Oktober 2026): decrement QTY TERKIRIM per baris dihapus --
  // QTY TERKIRIM sekarang dihitung ulang penuh dari PEMBELIAN setelah
  // baris lama benar-benar terhapus di database (lihat QTY_TERKIRIM_ULANG.js).
  // skipGateway=true (mode Timpa): penghapusan & hitung ulang dikerjakan
  // nanti oleh alur simpan, jadi di sini gak perlu apa-apa.
  if (skipGateway) return;

  if (existingRows.length > 0) {
    const noSJ = String(existingRows[0].noSJ || "").trim();
    const supplier = existingRows[0].supplier;
    kirimAksiGateway_("syncPembelian", { noSJ: noSJ, supplier: supplier, rows: [] });
  }

  try {
    hitungUlangQtyTerkirimPOKeluar_(false);
  } catch (err) {
    Logger.log("Hitung ulang QTY TERKIRIM (hapus PEMBELIAN) error: " + err);
  }
}

function isiFormTandaTerimaDariExisting_(shTerima, header, items) {
  // B2=supplier, B3=noPO (bisa gabungan, tetap ditulis apa adanya), B4=noPB LAMA
  shTerima.getRange("B2:B4").setValues([[header.supplier], [header.noPO], [header.noPB]]);
  // ⬅️ FIX: TGL BELI dari gateway datang sebagai teks ISO ("2026-10-27T00:00:00.000Z");
  // dulu ditulis apa adanya -> E3 jadi teks, dan kalau disimpan ulang PEMBELIAN/HUTANG
  // ikut dapat teks itu. Sekarang diubah balik jadi tanggal.
  let tglBeli = header.tglBeli;
  if (typeof tglBeli === "string" && /^\d{4}-\d{2}-\d{2}T/.test(tglBeli)) tglBeli = new Date(tglBeli);
  shTerima.getRange("E2:E4").setValues([[header.noSJ], [tglBeli], [header.noKend]]);
  shTerima.getRange("E3").setNumberFormat("dd/MM/yyyy");
  shTerima.getRange("G2:G3").setValues([[header.penerima], [header.pengirim]]);

  shTerima.getRange(TANDA_TERIMA_ITEM_START_ROW, 2, TANDA_TERIMA_ITEM_MAX_ROWS, 2).clearContent(); // B:C
  shTerima.getRange(TANDA_TERIMA_ITEM_START_ROW, 6, TANDA_TERIMA_ITEM_MAX_ROWS, 3).clearContent(); // F:H

  if (items.length > 0) {
    const bc = [];
    const fgh = [];
    items.forEach(function(it) {
      // PO asal per item: kalau NO PO lama gabungan, kosongkan (dicari ulang saat simpan)
      const poList = parsePoListTT_(it.noPO);
      const poAsal = poList.length === 1 ? poList[0] : "";
      bc.push([it.kode, it.qty]);
      fgh.push([it.namaToko || "", poAsal, it.setDiskon || ""]);
    });
    shTerima.getRange(TANDA_TERIMA_ITEM_START_ROW, 2, bc.length, 2).setValues(bc);
    shTerima.getRange(TANDA_TERIMA_ITEM_START_ROW, 6, fgh.length, 3).setValues(fgh);
  }

  pastikanFormulaTandaTerimaUtuh_(shTerima);
}

// Kosongkan seluruh isi form TANDA TERIMA BRG (header + item).
// D (NAMA) & A (NO) isinya ARRAYFORMULA, jadi tidak disentuh -- otomatis
// kosong sendiri begitu kolom B kosong.
function resetFormTandaTerima_(shTerima) {
  const s = TANDA_TERIMA_ITEM_START_ROW;
  const n = TANDA_TERIMA_ITEM_MAX_ROWS;
  const e = s + n; // +1 baris buat "Total"

  shTerima.getRangeList([
    "B2:B4",                        // supplier, NO PO, NO PB
    "E2:E4",                        // NO SJ, tgl beli, no kendaraan
    "G2:G3",                        // penerima, pengirim
    "B" + s + ":C" + (s + n - 1),   // kode & qty item
    "E" + s + ":E" + e,             // harga item + baris Total
    "F" + s + ":H" + (s + n - 1)    // toko, NO PO asal, status SET
  ]).clearContent();

  // hitung ulang (item kosong -> harga kosong) & bersihkan bold Total
  pastikanFormulaTandaTerimaUtuh_(shTerima);

  SpreadsheetApp.flush();
}

/* ======================================================
   Load barang dari HISTORY_PO_KELUAR ke TANDA TERIMA BRG
   Trigger: saat B3 (NO. PO) diisi/diubah
   ====================================================== */

/**
 * ⬅️ BARU (Nov 2026): menu "🔍 Pilih Barang dari PO (B3)". Membuka popup pilih
 * barang untuk NO PO yang SUDAH tertulis di B3, tanpa perlu mengetik ulang B3.
 * (Mengetik ulang nilai yang sama di B3 tidak dianggap perubahan oleh Google
 * Sheets, jadi pemicu otomatis tidak jalan -- itu sebabnya popup tidak muncul
 * setelah form dimuat lewat "Simpan SJ Beli" > Edit, karena B3 sudah terisi.)
 */
function pilihBarangDariPO() {
  loadTandaTerimaFromHistory(true);
}

function loadTandaTerimaFromHistory(manual) {
  try {
    loadTandaTerimaFromHistoryInternal_(manual === true);
  } catch (err) {
    Logger.log("Load Tanda Terima error: " + err);
    // ⬅️ FIX (Nov 2026): dulu error di sini diam-diam (cuma masuk log), jadi
    // admin cuma melihat "popup tidak muncul". Sekarang ditampilkan.
    SpreadsheetApp.getUi().alert(
      "Popup pilih barang gagal dibuka",
      String(err && err.message ? err.message : err) + "\n\nCoba lagi lewat menu SANFELO > 🔍 Pilih Barang dari PO (B3). Kalau tetap gagal, kabari IT beserta pesan ini.",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

function loadTandaTerimaFromHistoryInternal_(manual) {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const shTerima = ss.getSheetByName("TANDA TERIMA BRG");
  const hist     = ss.getSheetByName("HISTORY_PO_KELUAR");

  if (!shTerima || !hist) return;

  const rawB3 = String(shTerima.getRange("B3").getValue() || "").trim();

  // ⬅️ FIX (Nov 2026): isi form (kode, qty, toko, harga) TIDAK lagi dikosongkan
  // sebelum popup dikonfirmasi. Dulu form langsung kosong begitu B3 diubah, jadi
  // kalau popup ditutup/dibatalkan, barang dan harga yang sudah ada ikut hilang.
  // Penulisan ulang form sekarang hanya terjadi saat admin menekan "Simpan ke
  // TANDA TERIMA BRG" di popup (terapkanPilihanTandaTerimaPopup). Form baru
  // dikosongkan di sini hanya kalau B3 sengaja dikosongkan.
  if (!rawB3) {
    shTerima.getRange(TANDA_TERIMA_ITEM_START_ROW, 2, TANDA_TERIMA_ITEM_MAX_ROWS, 2).clearContent(); // B:C
    shTerima.getRange(TANDA_TERIMA_ITEM_START_ROW, 6, TANDA_TERIMA_ITEM_MAX_ROWS, 3).clearContent(); // F:H
    pastikanFormulaTandaTerimaUtuh_(shTerima);
    return;
  }

  const poNoList = parsePoListTT_(rawB3);

  const lastRow = hist.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert("HISTORY_PO_KELUAR masih kosong.");
    return;
  }

  const data = hist.getRange(2, 1, lastRow - 1, 8).getValues();

  let tglHeader = null;
  const combinedItems = [];
  const poGagal = [];

  poNoList.forEach(function(poNo) {
    const matches = data.filter(function(r) {
      return String(r[0] || "").trim() === poNo;
    });

    if (matches.length === 0) {
      poGagal.push(poNo + " (tidak ditemukan)");
      return;
    }

    if (tglHeader === null) tglHeader = matches[0][1];

    matches.forEach(function(r) {
      const namaToko    = r[2];
      const namaBarang  = r[3];
      const jumlah      = Number(r[4]) || 0;
      const qtyTerkirim = Number(r[7]) || 0;
      const sisaQty     = Math.max(0, jumlah - qtyTerkirim);

      if (!namaBarang) return;

      combinedItems.push({
        poNo: poNo,
        namaToko: namaToko,
        namaBarang: namaBarang,
        qtyAsli: jumlah,
        sisaQty: sisaQty
      });
    });
  });

  if (poGagal.length > 0) {
    SpreadsheetApp.getUi().alert("No PO berikut dilewati:\n" + poGagal.join("\n"));
  }

  if (combinedItems.length === 0) {
    SpreadsheetApp.getUi().alert('Tidak ada barang yang bisa dimuat dari No PO: "' + rawB3 + '"');
    return;
  }

  // ⬅️ FIX: tanggal PO cuma jadi isian AWAL kalau TGL (E3) masih kosong.
  // Dulu selalu menimpa E3 -> waktu PB lama diedit ulang (pilih barang lagi
  // lewat B3), TGL BELI ikut berubah jadi tanggal PO (kasus PB-2026-10-023:
  // 27/10 jadi 27/08, hutangnya pindah ke bulan Agustus).
  if (tglHeader && String(shTerima.getRange("E3").getValue() || "").trim() === "") {
    shTerima.getRange("E3").setValue(tglHeader);
  }

  tampilkanPopupPilihBarangTandaTerima_(combinedItems);
}

function tampilkanPopupPilihBarangTandaTerima_(items) {
  const template = HtmlService.createTemplateFromFile("PopupPilihBarangTandaTerima");
  template.modeDraft = false; // HTML popup memakai <?!= modeDraft ?>; tanpa ini popup error
  template.itemsJson = JSON.stringify(items);
  const html = template.evaluate().setWidth(480).setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, "Pilih Barang untuk TANDA TERIMA BRG");
}

function terapkanPilihanTandaTerimaPopup(selectedJson) {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const shTerima = ss.getSheetByName("TANDA TERIMA BRG");
  const selected = JSON.parse(selectedJson);

  if (!selected.length) throw new Error("Tidak ada barang yang dipilih.");

  const db = ss.getSheetByName("DATA_BERSIH");
  const dbMap = {};
  if (db && db.getLastRow() >= 1) {
    const dbData = db.getRange(1, 1, db.getLastRow(), 2).getValues();
    dbData.forEach(function(r) {
      const nama = String(r[1] || "").trim().toUpperCase();
      if (nama) dbMap[nama] = r[0];
    });
  }

  shTerima.getRange(TANDA_TERIMA_ITEM_START_ROW, 2, TANDA_TERIMA_ITEM_MAX_ROWS, 2).clearContent(); // B:C
  shTerima.getRange(TANDA_TERIMA_ITEM_START_ROW, 6, TANDA_TERIMA_ITEM_MAX_ROWS, 3).clearContent(); // F:H

  const bcRows  = [];
  const fghRows = [];

  selected.forEach(function(it) {
    const namaUpper = String(it.namaBarang || "").trim().toUpperCase();
    const kode = dbMap[namaUpper] || it.namaBarang;
    bcRows.push([kode, it.qty]);
    fghRows.push([it.namaToko || "", it.poNo || "", it.setDiskon || ""]);
  });

  shTerima.getRange(TANDA_TERIMA_ITEM_START_ROW, 2, bcRows.length, 2).setValues(bcRows);
  shTerima.getRange(TANDA_TERIMA_ITEM_START_ROW, 6, fghRows.length, 3).setValues(fghRows);

  pastikanFormulaTandaTerimaUtuh_(shTerima);
}

// itemsDelta: [{ namaBarang, qty, namaToko }]
// Matching ke HISTORY_PO_KELUAR: NO PO (tunggal) + NAMA BARANG, DAN (kalau namaToko ada)
// NAMA TOKO juga harus cocok. Baca sekali, hitung di memori, tulis kolom H sekali.
function updateQtyTerkirimHistoryPOKeluar_(ss, noPO, itemsDelta, delta) {
  const hist = ss.getSheetByName("HISTORY_PO_KELUAR");
  if (!hist) return;

  const noPOTrim = String(noPO || "").trim();
  if (!noPOTrim || !itemsDelta.length) return;

  const lastRow = hist.getLastRow();
  if (lastRow < 2) return;

  const numRows = lastRow - 1;
  const data = hist.getRange(2, 1, numRows, 8).getValues();
  const rangeH = hist.getRange(2, 8, numRows, 1);
  const formulasH = rangeH.getFormulas();
  const touched = {};
  let changed = false;

  itemsDelta.forEach(function(itd) {
    const namaUpper = String(itd.namaBarang || "").trim().toUpperCase();
    const tokoUpper = String(itd.namaToko || "").trim().toUpperCase();
    const qty = Number(itd.qty) || 0;
    if (!namaUpper || !qty) return;

    for (let i = 0; i < data.length; i++) {
      const rowNoPO = String(data[i][0] || "").trim();
      const rowNama = String(data[i][3] || "").trim().toUpperCase();
      const rowToko = String(data[i][2] || "").trim().toUpperCase();

      const cocokDasar = (rowNoPO === noPOTrim && rowNama === namaUpper);
      const cocokToko = tokoUpper ? (rowToko === tokoUpper) : true;

      if (cocokDasar && cocokToko) {
        const existingVal = Number(data[i][7]) || 0;
        let newVal = existingVal + (qty * delta);
        if (newVal < 0) newVal = 0;
        data[i][7] = newVal;
        touched[i] = true;
        changed = true;
        break;
      }
    }
  });

  if (!changed) return;

  const out = [];
  for (let i = 0; i < numRows; i++) {
    if (!touched[i] && formulasH[i][0]) {
      out.push([formulasH[i][0]]);
    } else {
      out.push([data[i][7]]);
    }
  }
  rangeH.setValues(out);
}

  // ======================= BiayaTambahan.js (salinan persis) =======================
/* ======================================================
   BiayaTambahan.gs
   - showBiayaTambahDialog() : buka popup Biaya Tambahan
   - simpanBiayaTambahan()   : simpan 1 baris ke sheet EKSPEDISI,
                               hitung TEMPO otomatis, lalu update
                               PENJUALAN kolom H/I & SISA BAYAR
   - hitungTempo_()          : lookup TEMPO dari sheet DISKON
   - addDaysToTanggal_()     : helper tambah hari ke tanggal dd/MM/yyyy
   - hitungUlangSisaBayar_() : hitung ulang SISA BAYAR (G) & STATUS (F)

   Catatan: konstanta (PENJUALAN_COL_*, PENJUALAN_WARNA_*) dan
   helper (formatRupiahText_, parseRupiahText_, parseRiwayatCell_,
   formatTglBayarInput_) dideklarasikan di Pembayaran.gs — otomatis
   kepakai di sini karena satu project Apps Script (global scope).

   ====== STRUKTUR SHEET EKSPEDISI (mulai baris 2) ======
   A: NO INV        B: REFRENSI     C: NAM EKSPEDISI  D: TANGGAL
   E: TEMPO (otomatis, hasil TANGGAL + n hari dari sheet DISKON)
   F: NAMA TOKO      G: ALAMAT       H: BARANG         I: QTY
   J: SAT            K: KET BRG      L: HARGA          M: METHOD

   ====== STRUKTUR SHEET DISKON (dipakai buat hitung TEMPO) ======
   A: NAMA TOKO   ...   E: TEMPO, formatnya teks "n HARI" (misal "30 HARI")
   Dicari baris yang NAMA TOKO (kolom A)-nya cocok sama Nama Toko yang
   diinput di popup, angka n diambil dari kolom E, lalu:
     TEMPO (di sheet EKSPEDISI) = TANGGAL input + n hari
   Kalau toko gak ketemu di DISKON, TEMPO dikosongin aja (gak error).

   ====== ATURAN NOMINAL KE PENJUALAN ======
   Nominal yang ditulis ke PENJUALAN kolom H (EKSPEDISI) atau I (PALET)
   = HARGA x QTY (dari input popup ini). Nilai lama di kolom itu
   DITIMPA (bukan ditambah), lalu SISA BAYAR (G) & STATUS (F) dihitung
   ulang otomatis oleh hitungUlangSisaBayar_().

   ⬅️ FIX: EKSPEDISI (H) & PALET (I) TIDAK ikut menambah SISA BAYAR.
   Konsisten dengan backfillSisaBayarPenjualan_() di Sync.gs. Kolom
   H/I tetap ditulis ke PENJUALAN seperti biasa (lihat simpanBiayaTambahan()
   di bawah), cuma gak lagi menambah nilai piutang / SISA BAYAR. SISA
   BAYAR murni dari TOTAL (E) dikurangi cicilan yang sudah masuk.

   ⬅️ PENTING: baris PENJUALAN yang dicari/di-update itu dicocokkan
   pakai field REFERENSI (bukan field "No Invoice"). Jadi REFERENSI
   di popup ini harus SAMA PERSIS dengan isi kolom NO INV (C) di
   sheet PENJUALAN. Field "No Invoice" di popup cuma catatan, ditulis
   ke kolom A sheet EKSPEDISI, gak dipakai buat matching sama sekali.

   ⬅️ BARU (Oktober 2026):
   - EKSPEDISI dapat kolom N: JENIS ("EKSPEDISI"/"PALET"), supaya tiap
     baris jelas masuk ke PENJUALAN H atau I.
   - PENJUALAN H/I = JUMLAH semua baris EKSPEDISI dengan REFERENSI &
     JENIS yang sama (dulu: ditimpa harga baris terakhir, jadi barang
     kedua menghapus barang pertama).
   - Setelah simpan, sheet INVOICE PALET langsung terisi untuk NO INV
     (REFERENSI) itu, siap print. Detail di INVOICE_PALET.js.
   ====================================================== */

const EKSPEDISI_COL_NOINV = 1;
const EKSPEDISI_COL_REFERENSI = 2;
const EKSPEDISI_COL_NAMEKSPEDISI = 3;
const EKSPEDISI_COL_TANGGAL = 4;
const EKSPEDISI_COL_TEMPO = 5;
const EKSPEDISI_COL_NAMATOKO = 6;
const EKSPEDISI_COL_ALAMAT = 7;
const EKSPEDISI_COL_BARANG = 8;
const EKSPEDISI_COL_QTY = 9;
const EKSPEDISI_COL_SAT = 10;
const EKSPEDISI_COL_KETBRG = 11;
const EKSPEDISI_COL_HARGA = 12;
const EKSPEDISI_COL_METHOD = 13;
const EKSPEDISI_COL_JENIS = 14;   // N (BARU): "EKSPEDISI" / "PALET"

/* ------------------------------------------------------
   1) BUKA POPUP BIAYA TAMBAHAN
   ------------------------------------------------------ */
function showBiayaTambahDialog() {
  const html = HtmlService.createHtmlOutputFromFile("BiayaTambahDialog")
    .setWidth(440)
    .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, "🚚 Biaya Tambahan (Ekspedisi/Palet)");
}

/* ------------------------------------------------------
   2) SIMPAN BIAYA TAMBAHAN
   Dipanggil dari HTML: google.script.run.simpanBiayaTambahan(data)
   Param `data` = {
     noInvoice, referensi, namaEkspedisi, tanggal (yyyy-MM-dd),
     namaToko, alamat, barang, qty, satuan, ketBrg, harga,
     method ("CASH"/"CREDIT"), jenisBiaya ("EKSPEDISI"/"PALET")
   }

   ⬅️ PAKAI LockService biar aman kalau ada 2 orang/proses submit
   hampir bersamaan. Tanpa lock, dua eksekusi bisa sama-sama baca
   getLastRow() SEBELUM salah satunya selesai nulis -> keduanya
   dapet target row yang sama -> baris pertama ketimpa baris kedua
   (keliatan kayak "data hilang"). Lock bikin satu eksekusi harus
   selesai TOTAL (baca lastRow -> tulis -> return) sebelum eksekusi
   berikutnya boleh mulai baca lastRow-nya sendiri.
   ------------------------------------------------------ */
function simpanBiayaTambahan(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000); // tunggu max 30 detik kalau lagi dipakai proses lain
  } catch (err) {
    return { success: false, message: "Server sedang sibuk, coba lagi sebentar." };
  }

  try {
    const noInvoice = String(data.noInvoice || "").trim();
    const referensi = String(data.referensi || "").trim();
    const namaEkspedisi = String(data.namaEkspedisi || "").trim();
    const tanggalText = formatTglBayarInput_(data.tanggal);
    const namaToko = String(data.namaToko || "").trim();
    const alamat = String(data.alamat || "").trim();
    const barang = String(data.barang || "").trim();
    const qty = Number(data.qty);
    const satuan = String(data.satuan || "").trim();
    const ketBrg = String(data.ketBrg || "").trim();
    const harga = Number(data.harga);
    const method = String(data.method || "").trim().toUpperCase();
    const jenisBiaya = String(data.jenisBiaya || "").trim().toUpperCase();

    if (!referensi) {
      return { success: false, message: "Referensi tidak boleh kosong (dipakai buat mencocokkan ke NO INV di PENJUALAN)." };
    }
    if (!tanggalText) {
      return { success: false, message: "Tanggal tidak valid." };
    }
    if (!namaToko) {
      return { success: false, message: "Nama Toko tidak boleh kosong." };
    }
    if (!qty || qty <= 0) {
      return { success: false, message: "QTY tidak valid." };
    }
    if (!harga || harga <= 0) {
      return { success: false, message: "HARGA tidak valid." };
    }
    if (method !== "CASH" && method !== "CREDIT") {
      return { success: false, message: "Method harus CASH atau CREDIT." };
    }
    if (jenisBiaya !== "EKSPEDISI" && jenisBiaya !== "PALET") {
      return { success: false, message: "Jenis biaya harus EKSPEDISI atau PALET." };
    }

    // ⬅️ BARU: baca EKSPEDISI fresh dari database (masih di dalam lock),
    // dipakai buat jumlah PENJUALAN H/I dan isi sheet INVOICE PALET.
    let rowsEkspedisi;
    try {
      rowsEkspedisi = bacaEkspedisiDb_();
    } catch (err) {
      return { success: false, message: "Gagal membaca EKSPEDISI dari database_SANFELO: " + err.message };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // ---- 1) Hitung TEMPO otomatis dari sheet DISKON ----
    const tempoText = hitungTempo_(namaToko, tanggalText);

    // ---- 2) Simpan baris baru ke sheet EKSPEDISI ----
    // ⬅️ BARU (migrasi): sheet EKSPEDISI sekarang di-mirror IMPORTRANGE
    // read-only di inventory (seluruh A:M, gak ada bagian yang perlu tetap
    // lokal -- beda dari TEMPO TOKO/HISTORY_PO_*, sheet ini emang gak
    // pernah dibaca balik oleh kode manapun, murni log). Baris baru
    // dikirim ke database lewat gateway, bukan ditulis langsung ke sheet
    // lokal (yang sekarang isinya formula IMPORTRANGE, gak bisa ditulis).
    const barisBaru = [
      noInvoice, referensi, namaEkspedisi, tanggalText, tempoText,
      namaToko, alamat, barang, qty, satuan, ketBrg, harga, method, jenisBiaya
    ];
    try {
      kirimAksiGateway_("appendEkspedisi", { rows: [barisBaru] });
    } catch (err) {
      return { success: false, message: "Gagal menyimpan EKSPEDISI ke database_SANFELO: " + err.message };
    }
    rowsEkspedisi.push(barisBaru);

    // ---- 3) Update PENJUALAN kolom H/I ----
    // ⬅️ BARU: nominal = JUMLAH semua baris EKSPEDISI dengan REFERENSI dan
    // JENIS yang sama (termasuk baris baru ini), bukan harga baris ini saja.
    // Baris lama yang kolom JENIS-nya kosong = tidak pakai palet/ekspedisi,
    // tidak ikut dihitung.
    const nominal = totalBiayaPerReferensi_(rowsEkspedisi, referensi, jenisBiaya);

    // ⬅️ MIGRASI (Oktober 2026): update kolom H (EKSPEDISI) / I (PALET) +
    // hitung ulang SISA BAYAR & STATUS sekarang jalan di database lewat aksi
    // gateway "updateBiayaTambahanPenjualan" (sudah ada di DatabaseGateway.js,
    // port persis dari hitungUlangSisaBayar_ di bawah). PENJUALAN lokal cuma
    // salinan IMPORTRANGE.
    let hasil;
    try {
      hasil = kirimAksiGateway_("updateBiayaTambahanPenjualan", {
        referensi: referensi,
        jenisBiaya: jenisBiaya,
        nominal: nominal
      });
    } catch (err) {
      return {
        success: false,
        message: "Data EKSPEDISI tersimpan, tapi gagal update PENJUALAN: " + String(err.message || err).replace(/^Gateway menolak: /, ""),
      };
    }
    refreshSalinan_(["PENJUALAN"]);

    // ---- 4) Tampilkan di sheet INVOICE PALET (siap print) ----
    let invoiceTampil = true;
    try {
      isiSheetInvoicePalet_(referensi, rowsEkspedisi.filter(function (r) {
        return String(r[EKSPEDISI_COL_REFERENSI - 1] || "").trim().toUpperCase() === referensi.toUpperCase();
      }));
    } catch (err) {
      Logger.log("isiSheetInvoicePalet_ gagal: " + err);
      invoiceTampil = false;
    }

    return {
      invoiceTampil: invoiceTampil,
      success: true,
      noInvoice: noInvoice,
      referensi: referensi,
      jenisBiaya: jenisBiaya,
      nominal: nominal,
      tempo: tempoText,
      totalTagihan: hasil.totalTagihan,
      sisaBayar: hasil.sisaBayar,
      status: hasil.status,
    };
  } finally {
    lock.releaseLock();
  }
}

/* ------------------------------------------------------
   3) HITUNG TEMPO OTOMATIS DARI SHEET DISKON
   Cari baris di sheet DISKON yang NAMA TOKO (kolom A)-nya cocok,
   ambil angka n dari kolom E (format teks "n HARI"), lalu return
   TANGGAL input + n hari (format dd/MM/yyyy).
   Kalau sheet DISKON gak ada, toko gak ketemu, atau format kolom E
   gak punya angka -> return "" (gak error, cuma TEMPO dikosongin).
   ------------------------------------------------------ */
function hitungTempo_(namaToko, tanggalText) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const shDiskon = ss.getSheetByName("DISKON");
    if (!shDiskon) return "";

    const lastRow = shDiskon.getLastRow();
    if (lastRow < 2) return "";

    const target = String(namaToko || "").trim().toUpperCase();
    if (!target) return "";

    const data = shDiskon.getRange(2, 1, lastRow - 1, 5).getValues(); // kolom A - E

    for (let i = 0; i < data.length; i++) {
      const toko = String(data[i][0] || "").trim().toUpperCase();
      if (toko !== target) continue;

      const rawTempo = String(data[i][4] || "").trim(); // kolom E, format "n HARI"
      const match = rawTempo.match(/\d+/);
      if (!match) return "";

      const n = Number(match[0]);
      return addDaysToTanggal_(tanggalText, n);
    }

    return ""; // toko gak ketemu di DISKON
  } catch (err) {
    Logger.log("hitungTempo_ error: " + err);
    return "";
  }
}

/* Tambah `days` hari ke tanggal berformat "dd/MM/yyyy", return string
   dengan format yang sama. Ditulis manual dari angka y/m/d (bukan
   parse ISO string) biar gak ada resiko geser hari gara-gara timezone. */
function addDaysToTanggal_(tanggalText, days) {
  const parts = String(tanggalText || "").split("/"); // [dd, MM, yyyy]
  if (parts.length !== 3) return "";

  const dd = Number(parts[0]);
  const mm = Number(parts[1]);
  const yyyy = Number(parts[2]);
  if (!dd || !mm || !yyyy) return "";

  const dateObj = new Date(yyyy, mm - 1, dd);
  dateObj.setDate(dateObj.getDate() + days);

  const newDd = String(dateObj.getDate()).padStart(2, "0");
  const newMm = String(dateObj.getMonth() + 1).padStart(2, "0");
  const newYyyy = dateObj.getFullYear();
  return newDd + "/" + newMm + "/" + newYyyy;
}

/* ------------------------------------------------------
   4) HITUNG ULANG SISA BAYAR (G) & STATUS (F) UNTUK 1 BARIS PENJUALAN

   ⬅️ FIX: EKSPEDISI (H) & PALET (I) TIDAK ikut dijumlahkan ke SISA
   BAYAR di sini, konsisten dengan backfillSisaBayarPenjualan_() di
   Sync.gs. Kolom H/I tetap ditulis ke PENJUALAN seperti biasa (lihat
   simpanBiayaTambahan() di atas), tapi SISA BAYAR (piutang customer)
   murni dihitung dari TOTAL (E) dikurangi cicilan yang sudah masuk.
   ------------------------------------------------------ */
function hitungUlangSisaBayar_(sh, rowIndex) {
  const lastCol = Math.max(sh.getLastColumn(), PENJUALAN_COL_CICILAN_START);
  const rowValues = sh.getRange(rowIndex, 1, 1, lastCol).getValues()[0];

  // totalTagihan = TOTAL (E) doang, EKSPEDISI (H) & PALET (I) gak diikutkan
  const totalTagihan = Number(rowValues[PENJUALAN_COL_TOTAL - 1]) || 0;

  let totalSudahDibayar = 0;
  for (let c = PENJUALAN_COL_CICILAN_START - 1; c < rowValues.length; c++) {
    const parsed = parseRiwayatCell_(rowValues[c]);
    if (parsed) totalSudahDibayar += parsed.jumlah;
  }

  const statusLama = String(rowValues[PENJUALAN_COL_STATUS - 1] || "").trim().toUpperCase();
  const sisaRange = sh.getRange(rowIndex, PENJUALAN_COL_SISA);
  const statusRange = sh.getRange(rowIndex, PENJUALAN_COL_STATUS);

  if (totalSudahDibayar <= 0) {
    statusRange.setValue("");
    sisaRange.setValue(totalTagihan).setNumberFormat("#,##0").setBackground(null);
    if (statusLama === "LUNAS") {
      sh.getRange(rowIndex, 1, 1, lastCol).setBackground(null);
    }
    return { totalTagihan: totalTagihan, sisaBayar: totalTagihan, status: "" };
  }

  const sisaBaru = totalTagihan - totalSudahDibayar;

  if (sisaBaru <= 0) {
    statusRange.setValue("LUNAS");
    sh.getRange(rowIndex, 1, 1, lastCol).setBackground(PENJUALAN_WARNA_LUNAS);
    if (sisaBaru < 0) {
      sisaRange.setValue("LEBIH " + formatRupiahText_(-sisaBaru)).setBackground(PENJUALAN_WARNA_LEBIH);
    } else {
      sisaRange.setValue(0).setNumberFormat("#,##0").setBackground(PENJUALAN_WARNA_LUNAS);
    }
    return { totalTagihan: totalTagihan, sisaBayar: sisaBaru, status: "LUNAS" };
  }

  statusRange.setValue("CICIL");
  sisaRange.setValue(sisaBaru).setNumberFormat("#,##0").setBackground(null);
  if (statusLama === "LUNAS") {
    sh.getRange(rowIndex, 1, 1, lastCol).setBackground(null);
  }
  return { totalTagihan: totalTagihan, sisaBayar: sisaBaru, status: "CICIL" };
}

  // ======================= INVOICE_PALET.js (salinan persis) =======================
/* ======================================================
   INVOICE_PALET.js   (PROJECT: inventory_SANFELO)
   ------------------------------------------------------
   Sheet "INVOICE PALET" = form siap print untuk biaya EKSPEDISI & PALET
   satu NO INV (= REFERENSI di sheet EKSPEDISI = NO INV di PENJUALAN).
   Tidak ada nomor baru: nomornya ya NO INV yang sudah ada.

   Isinya = semua baris EKSPEDISI dengan REFERENSI itu, dikelompokkan
   per JENIS (kolom N EKSPEDISI):
     BIAYA EKSPEDISI = jumlah baris JENIS "EKSPEDISI"  -> sama dengan PENJUALAN H
     BIAYA PALET     = jumlah baris JENIS "PALET"      -> sama dengan PENJUALAN I
   Baris yang JENIS-nya kosong = tidak pakai palet / ekspedisi (sesuai admin),
   jadi tidak ikut dihitung dan tidak tampil di invoice.
   Kalau H/I di PENJUALAN beda dengan EKSPEDISI, jalankan
   syncBiayaTambahanKePenjualan di database (SYNC_BIAYA_TAMBAHAN.js).

   ALUR:
     🚚 Biaya Tambah -> isi -> Simpan  => sheet INVOICE PALET langsung terisi
     📂 Buka Invoice Palet             => pilih NO INV -> sheet terisi lagi
     Cetak: buka sheet INVOICE PALET -> File > Cetak (pilih "Sheet saat ini")

   Sheet INVOICE PALET dibuat & ditata ulang otomatis oleh kode setiap kali
   diisi -- jangan diketik manual, isinya akan ditimpa.
   ====================================================== */

const INVOICE_PALET_SHEET = 'INVOICE PALET';
const INVOICE_PALET_MIN_BARIS = 8;   // baris barang kosong minimal biar form gak kelihatan "pendek"
const INVOICE_PALET_ITEM_ROW = 8;    // baris pertama barang (header tabel di baris 7)
const INVOICE_PALET_COLS = 7;        // A..G

/* ---------------- hitungan ---------------- */

function jenisBarisEkspedisi_(r) {
  const j = String(r[EKSPEDISI_COL_JENIS - 1] || '').trim().toUpperCase();
  return (j === 'EKSPEDISI' || j === 'PALET') ? j : '';
}

// jumlah HARGA semua baris EKSPEDISI dengan REFERENSI & JENIS yang sama -> nilai PENJUALAN H/I
function totalBiayaPerReferensi_(rowsEkspedisi, referensi, jenisBiaya) {
  const ref = String(referensi || '').trim().toUpperCase();
  return rowsEkspedisi.reduce(function (s, r) {
    if (String(r[EKSPEDISI_COL_REFERENSI - 1] || '').trim().toUpperCase() !== ref) return s;
    if (jenisBarisEkspedisi_(r) !== jenisBiaya) return s;
    return s + (Number(r[EKSPEDISI_COL_HARGA - 1]) || 0);
  }, 0);
}

/* ---------------- baca data ---------------- */

// tanggal dari EKSPEDISI bisa berupa Date, teks ISO (hasil gateway) atau teks dd/MM/yyyy
function tglTeksPalet_(v) {
  if (v === '' || v === null || v === undefined) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return isNaN(v.getTime()) ? '' : Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  }
  const s = String(v).trim();
  if (/^\d{4}-\d{2}-\d{2}T/.test(s)) return Utilities.formatDate(new Date(s), Session.getScriptTimeZone(), 'dd/MM/yyyy');
  return s;
}

// semua baris EKSPEDISI fresh dari database (A..N)
function bacaEkspedisiDb_() {
  return bacaSheetGateway_('EKSPEDISI', 2, 1, EKSPEDISI_COL_JENIS).filter(function (r) {
    return String(r[EKSPEDISI_COL_REFERENSI - 1] || '').trim() !== '' || String(r[EKSPEDISI_COL_BARANG - 1] || '').trim() !== '';
  });
}

/* ---------------- menu: buka invoice ---------------- */

function bukaInvoicePalet() {
  const html = HtmlService.createHtmlOutputFromFile('BukaInvoicePaletDialog').setWidth(560).setHeight(520);
  SpreadsheetApp.getUi().showModalDialog(html, '📂 Buka Invoice Palet / Ekspedisi');
}

// dipanggil dari BukaInvoicePaletDialog: daftar per NO INV, terbaru dulu,
// sekalian dibandingkan dengan PENJUALAN H/I
function getDaftarInvoicePalet() {
  const penjualan = {};
  bacaSheetGateway_('PENJUALAN', 2, 1, 9).forEach(function (r) {
    const no = String(r[2] || '').trim().toUpperCase();
    if (no) penjualan[no] = { h: Number(r[7]) || 0, i: Number(r[8]) || 0 };
  });

  const grup = {}, urutan = [];
  bacaEkspedisiDb_().forEach(function (r) {
    const ref = String(r[EKSPEDISI_COL_REFERENSI - 1] || '').trim();
    const jenis = jenisBarisEkspedisi_(r);
    if (!ref || !jenis) return; // JENIS kosong = tidak pakai palet/ekspedisi
    const k = ref.toUpperCase();
    if (!grup[k]) {
      grup[k] = { referensi: ref, toko: String(r[EKSPEDISI_COL_NAMATOKO - 1] || '').trim(),
        tanggal: tglTeksPalet_(r[EKSPEDISI_COL_TANGGAL - 1]), ekspedisi: 0, palet: 0, item: 0 };
      urutan.push(k);
    }
    const g = grup[k], harga = Number(r[EKSPEDISI_COL_HARGA - 1]) || 0;
    if (jenis === 'EKSPEDISI') g.ekspedisi += harga;
    else g.palet += harga;
    g.item++;
  });

  // urutan baris di EKSPEDISI = urutan input, jadi dibalik -> yang terbaru di atas
  return urutan.reverse().map(function (k) {
    const g = grup[k], p = penjualan[k];
    g.adaDiPenjualan = !!p;
    g.penjualanH = p ? p.h : 0;
    g.penjualanI = p ? p.i : 0;
    g.sejalan = !!p && p.h === g.ekspedisi && p.i === g.palet;
    return g;
  });
}

// dipanggil dari BukaInvoicePaletDialog
function tampilkanInvoicePalet(referensi) {
  const ref = String(referensi || '').trim().toUpperCase();
  const rows = bacaEkspedisiDb_().filter(function (r) {
    return String(r[EKSPEDISI_COL_REFERENSI - 1] || '').trim().toUpperCase() === ref && jenisBarisEkspedisi_(r);
  });
  if (!rows.length) return { success: false, message: 'NO INV ' + referensi + ' tidak punya biaya palet / ekspedisi.' };
  isiSheetInvoicePalet_(String(referensi).trim(), rows);
  return { success: true };
}

/* ---------------- tulis sheet ---------------- */

/**
 * Isi ulang sheet INVOICE PALET untuk satu NO INV.
 * rows : baris-baris EKSPEDISI (A..N) dengan REFERENSI = noInv
 */
function isiSheetInvoicePalet_(noInv, rows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(INVOICE_PALET_SHEET);
  if (!sh) sh = ss.insertSheet(INVOICE_PALET_SHEET);

  // JENIS kosong = tidak pakai palet/ekspedisi -> tidak ikut di invoice
  const isi = susunInvoicePalet_(noInv, rows.filter(jenisBarisEkspedisi_));
  const C = INVOICE_PALET_COLS;
  const nItem = Math.max(isi.items.length, INVOICE_PALET_MIN_BARIS);
  const barisSubtotal = INVOICE_PALET_ITEM_ROW + nItem;
  const ringkasan = [['BIAYA EKSPEDISI', isi.totalEkspedisi], ['BIAYA PALET', isi.totalPalet]];
  ringkasan.push(['TOTAL', isi.total]);
  const barisAkhir = barisSubtotal + ringkasan.length - 1;

  // bersihkan total (isi, format, merge) lalu tata ulang
  if (sh.getMaxRows() < barisAkhir) sh.insertRowsAfter(sh.getMaxRows(), barisAkhir - sh.getMaxRows());
  if (sh.getMaxColumns() < C) sh.insertColumnsAfter(sh.getMaxColumns(), C - sh.getMaxColumns());
  const semua = sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns());
  semua.breakApart();
  semua.clear();

  [90, 240, 95, 55, 60, 200, 120].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  sh.getRange(1, 1, barisAkhir, C).setFontFamily('Times New Roman').setFontSize(11).setVerticalAlignment('middle');
  sh.setHiddenGridlines(true);

  // judul
  sh.getRange(1, 1, 1, C).merge().setValue(isi.judul)
    .setFontSize(14).setFontWeight('bold').setHorizontalAlignment('center');

  // kepala invoice
  const kepala = [
    ['NO. INV:', isi.noInv, 'TANGGAL:', isi.tanggal],
    ['KEPADA:', isi.toko, 'JATUH TEMPO:', isi.tempo],
    ['ALAMAT:', isi.alamat, 'METHOD:', isi.method],
    ['EKSPEDISI:', isi.ekspedisi, '', '']
  ];
  kepala.forEach(function (k, i) {
    const row = 2 + i;
    sh.getRange(row, 1).setValue(k[0]).setHorizontalAlignment('right');
    sh.getRange(row, 2, 1, 4).merge().setValue(k[1]).setHorizontalAlignment('left').setWrap(true);
    sh.getRange(row, 6).setValue(k[2]).setHorizontalAlignment('right');
    sh.getRange(row, 7).setNumberFormat('@').setValue(k[3]).setHorizontalAlignment('left');
  });
  sh.getRange(2, 2).setFontWeight('bold');

  // tabel barang
  sh.getRange(INVOICE_PALET_ITEM_ROW - 1, 1, 1, C)
    .setValues([['NO', 'NAMA BARANG', 'JENIS', 'QTY', 'SAT', 'KETERANGAN', 'JUMLAH (Rp)']])
    .setFontWeight('bold').setHorizontalAlignment('center');

  const dataItem = [];
  for (let i = 0; i < nItem; i++) {
    const it = isi.items[i];
    dataItem.push(it ? [i + 1, it.barang, it.jenis || '-', it.qty, it.sat, it.ket, it.jumlah] : ['', '', '', '', '', '', '']);
  }
  sh.getRange(INVOICE_PALET_ITEM_ROW, 1, nItem, C).setValues(dataItem).setWrap(true);
  sh.getRange(INVOICE_PALET_ITEM_ROW, 1, nItem, 1).setHorizontalAlignment('center');
  sh.getRange(INVOICE_PALET_ITEM_ROW, 3, nItem, 3).setHorizontalAlignment('center');
  sh.getRange(INVOICE_PALET_ITEM_ROW, 7, nItem, 1).setNumberFormat('#,##0').setHorizontalAlignment('right');
  sh.getRange(INVOICE_PALET_ITEM_ROW - 1, 1, nItem + 1, C).setBorder(true, true, true, true, true, true);

  // subtotal per jenis + total
  ringkasan.forEach(function (rk, i) {
    const row = barisSubtotal + i;
    const tebal = rk[0] === 'TOTAL' ? 'bold' : 'normal';
    sh.getRange(row, 1, 1, C - 1).merge().setValue(rk[0]).setFontWeight(tebal).setHorizontalAlignment('right');
    sh.getRange(row, C).setValue(rk[1]).setNumberFormat('#,##0').setFontWeight(tebal).setHorizontalAlignment('right');
  });
  sh.getRange(barisSubtotal, 1, ringkasan.length, C).setBorder(true, true, true, true, true, true);

  // sisa baris di bawah form dirapikan biar area print gak kepanjangan
  if (sh.getMaxRows() > barisAkhir + 5) sh.deleteRows(barisAkhir + 6, sh.getMaxRows() - barisAkhir - 5);

  SpreadsheetApp.flush();
  ss.setActiveSheet(sh);
  sh.setActiveSelection('A1');
}

// inti (tanpa sheet): ubah baris EKSPEDISI jadi isi form
function susunInvoicePalet_(noInv, rows) {
  const ambil = function (col) {
    const daftar = [];
    rows.forEach(function (r) {
      const v = String(r[col - 1] || '').trim();
      if (v && daftar.indexOf(v) === -1) daftar.push(v);
    });
    return daftar;
  };
  const items = rows.map(function (r) {
    return {
      barang: String(r[EKSPEDISI_COL_BARANG - 1] || '').trim(),
      jenis: jenisBarisEkspedisi_(r),
      qty: Number(r[EKSPEDISI_COL_QTY - 1]) || '',
      sat: String(r[EKSPEDISI_COL_SAT - 1] || '').trim().toUpperCase(),
      ket: String(r[EKSPEDISI_COL_KETBRG - 1] || '').trim(),
      jumlah: Number(r[EKSPEDISI_COL_HARGA - 1]) || 0   // HARGA di EKSPEDISI = nominal total baris ini
    };
  });
  const jumlahJenis = function (j) {
    return items.reduce(function (s, it) { return s + (it.jenis === j ? it.jumlah : 0); }, 0);
  };
  const adaEks = items.some(function (it) { return it.jenis === 'EKSPEDISI'; });
  const adaPlt = items.some(function (it) { return it.jenis === 'PALET'; });
  const tempo = rows.map(function (r) { return tglTeksPalet_(r[EKSPEDISI_COL_TEMPO - 1]); }).filter(String);
  return {
    judul: adaPlt && !adaEks ? 'INVOICE PALET' : adaEks && !adaPlt ? 'INVOICE EKSPEDISI' : 'INVOICE EKSPEDISI & PALET',
    noInv: noInv,
    tanggal: tglTeksPalet_(rows[0][EKSPEDISI_COL_TANGGAL - 1]),
    tempo: tempo[0] || '-',
    toko: ambil(EKSPEDISI_COL_NAMATOKO)[0] || '',
    alamat: ambil(EKSPEDISI_COL_ALAMAT)[0] || '',
    ekspedisi: ambil(EKSPEDISI_COL_NAMEKSPEDISI).join(', ') || '-',
    method: ambil(EKSPEDISI_COL_METHOD).join(' / '),
    items: items,
    totalEkspedisi: jumlahJenis('EKSPEDISI'),
    totalPalet: jumlahJenis('PALET'),
    total: items.reduce(function (s, it) { return s + it.jumlah; }, 0)
  };
}


  // ======================= SPAREPART_SAVE.js (salinan persis) =======================
/* ======================================================
   SPAREPART_SAVE.gs
   Simpan data dari form SJ_SPR + INV_SPR ke sheet SPAREPART.
   Panggil fungsi simpanSparepart() dari menu UI yang udah ada.

   ⬅️ BARU (migrasi database_SANFELO, September 2026):
   - Sheet SPAREPART sekarang tinggal di database_SANFELO. Di sisi
     inventory ini SPAREPART cuma mirror IMPORTRANGE baca-saja (A:S,
     19 kolom — TIDAK ADA kolom lokal-only di sheet ini, beda dari
     TEMPO TOKO/HISTORY_PO_KELUAR).
   - finalizeSparepartSave_() sekarang kirim baris final (sudah
     termasuk STATUS ORDER yang tepat buat mode edit/baru, dihitung di
     sini seperti sebelumnya) ke gateway (aksi "syncSparepart"), yang
     melakukan delete-by-NO SJ + insert terhadap SPAREPART ASLI di
     database — sama pola "hapus lalu tulis ulang" yang dulu dipakai
     hapusBarisSparepartByNoSJ_() + shSPR.getRange(...).setValues(...)
     di sini, cuma sekarang jalan di server.
   - hapusBarisSparepartByNoSJ_() lama SUDAH DIHAPUS dari sini (logic-nya
     pindah ke aksiSyncSparepart_ di DatabaseGateway.js).
   - getNextSparepartNumber_() sekarang cross-check ke SPAREPART lewat
     bacaSheetGateway_ (baca FRESH langsung dari database, bukan mirror
     lokal yang bisa delay) — penting biar nomor SJ/INV sparepart gak
     tabrakan kalau ada transaksi lain yang baru aja disimpan device/tab
     lain tapi belum ke-refresh di mirror. Cache counter di sheet SYSTEM
     TETAP LOKAL (SYSTEM belum dimigrasi), logic-nya tidak berubah.
   - buildPembelianMap_() TETAP LOKAL tanpa perubahan (PEMBELIAN belum
     dimigrasi).
   - Alur klik admin (isi form SJ_SPR/INV_SPR -> "Simpan SPAREPART")
     SAMA SEKALI TIDAK BERUBAH.
   ====================================================== */

/* ---------- Jumlah maksimal baris barang per form (SJ baris 9-16, INV baris 11-18) ---------- */
const SPR_MAX_LINES_ = 8;

/* ---------- Fungsi utama: simpan form ke SPAREPART ---------- */
function simpanSparepart() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shSJ  = ss.getSheetByName('SJ_SPR');
  const shINV = ss.getSheetByName('INV_SPR');
  const shSPR = ss.getSheetByName('SPAREPART');
  const shPMB = ss.getSheetByName('PEMBELIAN');
  const shSYS = ss.getSheetByName('SYSTEM');

  if (!shSJ || !shINV || !shSPR || !shPMB || !shSYS) {
    SpreadsheetApp.getUi().alert('Ada sheet yang gak ketemu. Pastikan SJ_SPR, INV_SPR, SPAREPART, PEMBELIAN, SYSTEM semua ada.');
    return;
  }

  // 1. Kumpulkan baris barang dari SJ_SPR (B/E/F/G, baris 9-16) + Total dari INV_SPR (G, baris 10-17)
  const barangRows = [];
  for (let i = 0; i < SPR_MAX_LINES_; i++) {
    const sjRow  = 9 + i;
    const invRow = 11 + i;
    const barang = shSJ.getRange(sjRow, 2).getValue(); // B = Barang

    if (barang === '' || barang === null || barang === undefined) continue;

    barangRows.push({
      barang: barang,
      qty:    shSJ.getRange(sjRow, 5).getValue(),   // E = Qty
      sat:    shSJ.getRange(sjRow, 6).getValue(),   // F = SAT
      ketBrg: shSJ.getRange(sjRow, 7).getValue(),   // G = Keterangan (barang)
      harga:  shINV.getRange(invRow, 7).getValue()  // G = Total (Rp) di INV_SPR, baris 11-18
    });
  }

  if (barangRows.length === 0) {
    SpreadsheetApp.getUi().alert('Belum ada barang yang diisi di SJ_SPR (baris 9-16). Isi dulu ya sebelum simpan.');
    return;
  }

  // ⬅️ BARU (Oktober 2026 -- kasus SPRIN-2026-08-001 SIRO ke-simpan Rp0):
  // harga (INV_SPR kolom G) diisi manual. Kalau kosong/0, dulu langsung
  // kesimpan Rp0 tanpa peringatan, lalu di PENJUALAN otomatis dianggap
  // LUNAS (sisa bayar 0) -- padahal bisa jadi cuma lupa diisi. Sekarang
  // admin ditanya dulu. Dicek SEBELUM nomor SJ/INV di-generate, jadi
  // kalau dibatalkan gak ada nomor yang kepakai.
  const barangTanpaHarga = barangRows.filter(function (item) {
    return !(Number(item.harga) > 0);
  });
  if (barangTanpaHarga.length > 0) {
    const ui = SpreadsheetApp.getUi();
    const daftar = barangTanpaHarga.map(function (item) {
      return '- ' + String(item.barang).trim() + ' (qty ' + item.qty + ')';
    }).join('\n');
    const jawab = ui.alert(
      'Harga Kosong / Rp0',
      'Barang berikut harganya kosong atau Rp0 di INV_SPR kolom G:\n\n' + daftar + '\n\n' +
      'Kalau disimpan, barang ini dianggap TIDAK DITAGIHKAN dan invoice-nya ' +
      'bisa langsung tercatat LUNAS di PENJUALAN.\n\n' +
      'Memang tidak ditagihkan?\n' +
      '• YA = tetap simpan dengan Rp0\n' +
      '• TIDAK = batal, isi harganya dulu',
      ui.ButtonSet.YES_NO
    );
    if (jawab !== ui.Button.YES) {
      ui.alert('Simpan dibatalkan. Isi harga di INV_SPR kolom G (baris 11-18), lalu klik Simpan SPAREPART lagi.');
      return;
    }
  }

  // 2. Ambil data header form
  const tglSJ    = shSJ.getRange('H5').getValue();   // Tgl SJ
  const noKend   = shSJ.getRange('H7').getValue();   // No Kendaraan
  const namaToko = shSJ.getRange('A5').getValue();   // Nama Toko
  const alamat   = shSJ.getRange('A6').getValue();   // Alamat
  const method   = shSJ.getRange('H3').getValue();   // Metode Pembayaran
  const tempo    = shINV.getRange('H9').getValue();  // Jatuh tempo
  const refrensi = shINV.getRange('H7').getValue();  // Refrensi
  const ketInv   = shINV.getRange('C25').getValue(); // Keterangan invoice
  const sales    = shINV.getRange('C20').getValue(); // Nama Sales

  // 3. No SJ & No INV: pakai yang udah "dipesan"/kepampang di H6 & H5
  //    (hasil generate dari transaksi sebelumnya, ATAU hasil load draft
  //    edit dari SPAREPART). Kalau kosong (mis. pertama kali dipakai),
  //    generate baru sekarang juga.
  let noSJ  = shSJ.getRange('H6').getValue();
  let noInv = shINV.getRange('H5').getValue();

  // ⬅️ FIX (Oktober 2026 -- form masih nyimpen SPRIN-2026-08-003 di
  // bulan Oktober): nomor "siap pakai" di H5/H6 dibuat waktu transaksi
  // TERAKHIR disimpan. Kalau bulannya sudah ganti, nomor itu basi dan
  // dulu tetap kepakai (invoice Oktober bernomor Agustus). Sekarang kalau
  // BUKAN mode edit draft dan bulan di nomornya bukan bulan berjalan,
  // nomornya dibuang dan di-generate ulang. Mode edit tetap pakai nomor
  // lamanya apa adanya (baris lama di SPAREPART harus ketimpa).
  const editNoSJAktif = PropertiesService.getDocumentProperties().getProperty('SPAREPART_EDIT_NOSJ');
  const modeEditForm = !!(editNoSJAktif && editNoSJAktif === String(noSJ).trim());
  if (!modeEditForm) {
    if (nomorSparepartBulanLama_(noSJ))  noSJ  = '';
    if (nomorSparepartBulanLama_(noInv)) noInv = '';
  }

  if (!noSJ)  noSJ  = getNextSparepartNumber_(shSPR, shSYS, 'SJ');
  if (!noInv) noInv = getNextSparepartNumber_(shSPR, shSYS, 'INV');

  shSJ.getRange('H6').setValue(noSJ);
  shINV.getRange('H5').setValue(noInv);
  SpreadsheetApp.flush(); // pastikan formula INV_SPR!H6 (=SJ_SPR!H6) & H7 (=SJ_SPR!H5) ke-update

  // 4. Lookup HPP dari PEMBELIAN (kolom H = Nama Barang, kolom I = Qty, kolom J = Total)
  const pembelianMap = buildPembelianMap_();

  // 5. Pisahkan barang yang HPP-nya udah pasti (0 atau 1 entry di PEMBELIAN)
  //    vs yang ambigu (>1 entry, perlu user pilih salah satu)
  const hppKnown = {};   // { KEY_BARANG: unitPrice }
  const ambiguous = [];  // [{ key, namaBarang, entries: [{supplier, total, qty, unitPrice}] }]
  const seenKeys = {};

  barangRows.forEach(function (item) {
    const key = String(item.barang).trim().toUpperCase();
    if (seenKeys.hasOwnProperty(key)) return;
    seenKeys[key] = true;

    const entries = pembelianMap[key] || [];
    if (entries.length === 0) {
      hppKnown[key] = ''; // gak ketemu di PEMBELIAN, dikosongin (kayak logic lama)
    } else if (entries.length === 1) {
      hppKnown[key] = entries[0].unitPrice;
    } else {
      ambiguous.push({
        key: key,
        namaBarang: entries[0].namaBarang,
        entries: entries.map(function (e) {
          return { supplier: e.supplier, total: e.total, qty: e.qty, unitPrice: e.unitPrice };
        })
      });
    }
  });

  // 6. Susun "state" transaksi ini biar bisa dilanjutkan setelah user milih di dialog
  const state = {
    barangRows: barangRows,
    header: {
      tglSJ:    tglSJ instanceof Date ? tglSJ.toISOString() : tglSJ,
      noKend:   noKend,
      namaToko: namaToko,
      alamat:   alamat,
      method:   method,
      tempo:    tempo instanceof Date ? tempo.toISOString() : tempo,
      refrensi: refrensi,
      ketInv:   ketInv,
      sales:    sales
    },
    noSJ: noSJ,
    noInv: noInv,
    hppKnown: hppKnown,
    ambiguous: ambiguous
  };

  // 7. Kalau gak ada yang ambigu, langsung lanjut simpan
  if (ambiguous.length === 0) {
    finalizeSparepartSave_(state, hppKnown);
    return;
  }

  // 8. Kalau ada barang yang datanya dobel di PEMBELIAN, minta user pilih dulu
  const token = Utilities.getUuid();
  CacheService.getScriptCache().put(token, JSON.stringify(state), 1800); // 30 menit
  const html = buildHppDialogHtml_(token, ambiguous);
  SpreadsheetApp.getUi().showModalDialog(html, 'Pilih Data Pembelian');
}

/* ---------- Dipanggil dari dialog setelah user milih entry PEMBELIAN utk tiap barang ambigu ---------- */
function resolveSparepartSelections_(token, selections) {
  const cache = CacheService.getScriptCache();
  const raw = cache.get(token);
  if (!raw) {
    throw new Error('Sesi udah kadaluarsa, silakan ulangi simpan dari awal.');
  }
  cache.remove(token);

  const state = JSON.parse(raw);
  const hppMap = state.hppKnown || {};

  state.ambiguous.forEach(function (g) {
    const idx = (selections && selections.hasOwnProperty(g.key)) ? selections[g.key] : 0;
    const chosen = g.entries[idx] || g.entries[0];
    hppMap[g.key] = chosen.unitPrice;
  });

  finalizeSparepartSave_(state, hppMap);
}

/* ---------- Tahap akhir: hitung HPP & LABA per baris, tulis ke SPAREPART, reset form ---------- */
function finalizeSparepartSave_(state, hppMap) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shSJ  = ss.getSheetByName('SJ_SPR');
  const shINV = ss.getSheetByName('INV_SPR');
  const shSPR = ss.getSheetByName('SPAREPART');
  const shSYS = ss.getSheetByName('SYSTEM');

  const h = state.header;
  const tglSJ = h.tglSJ ? new Date(h.tglSJ) : h.tglSJ;
  const tempo = h.tempo ? new Date(h.tempo) : h.tempo;

  // Cek apakah form ini lagi mode EDIT (hasil load draft dari SPAREPART)
  // utk No SJ yang sama persis. Kalau ya -> baris lama ditimpa & STATUS
  // ORDER lama dipertahankan. Kalau enggak (No SJ beda / gak ada
  // penanda) -> berarti transaksi baru, ditambahkan sbg baris baru biasa.
  const docProps = PropertiesService.getDocumentProperties();
  const editNoSJ = docProps.getProperty('SPAREPART_EDIT_NOSJ');
  const isEditMode = !!(editNoSJ && editNoSJ === state.noSJ);
  const statusOrder = isEditMode
    ? (docProps.getProperty('SPAREPART_EDIT_STATUS') || 'NOT DONE')
    : 'NOT DONE';

  // Susun baris-baris untuk SPAREPART
  // HPP (kolom O) = (TOTAL PEMBELIAN / QTY PEMBELIAN) * QTY SPAREPART (kolom J)
  const rowsToWrite = state.barangRows.map(function (item) {
    const key = String(item.barang).trim().toUpperCase();
    const unitPrice = hppMap.hasOwnProperty(key) ? hppMap[key] : '';
    const harga = item.harga;
    const hpp = (typeof unitPrice === 'number') ? unitPrice * item.qty : '';
    const laba = (typeof harga === 'number' && typeof hpp === 'number') ? (harga - hpp) : '';

    return [
      state.noSJ,   // A  NO SJ
      state.noInv,  // B  NO INV
      h.refrensi,   // C  REFRENSI
      tglSJ,        // D  TANGGAL
      tempo,        // E  TEMPO
      h.noKend,     // F  NO KEND
      h.namaToko,   // G  NAMA TOKO
      h.alamat,     // H  ALAMAT
      item.barang,  // I  BARANG
      item.qty,     // J  QTY
      item.sat,     // K  SAT
      item.ketBrg,  // L  KET BRG
      h.ketInv,     // M  KET INV
      harga,        // N  HARGA
      hpp,          // O  HPP BARANG
      h.method,     // P  METHOD
      h.sales,      // Q  SALES
      laba,         // R  LABA
      statusOrder   // S  STATUS ORDER (dipertahankan kalau mode edit)
    ];
  });

  // ⬅️ BARU: tulis (delete-by-NO SJ + insert) sekarang jalan di gateway,
  // terhadap SPAREPART asli di database — bukan lokal lagi. Serialisasi
  // Date jadi ISO string biar aman lewat JSON (server parse ulang pakai
  // parseTanggalDariGateway_, sama pola kayak TEMPO TOKO/HUTANG).
  const rowsPayload = rowsToWrite.map(function (row) {
    return row.map(function (val) {
      return (val instanceof Date) ? val.toISOString() : val;
    });
  });
  kirimAksiGateway_("syncSparepart", { noSJ: state.noSJ, rows: rowsPayload });

  // Sync ke PIUTANG (full rescan semua SPAREPART, upsert per NO INV).
  // Fungsi ada di SYNC_PIUTANG.gs, logic PIUTANG sendiri gak diubah.
  // ⬅️ FIX: dulu kalau gagal cuma masuk log -- admin gak tau invoice-nya
  // gak masuk PENJUALAN (kasus SPRIN-2026-08-001: gak bisa dibayar/posting
  // karena "tidak ditemukan di PENJUALAN"). Sekarang ikut ditampilin di
  // alert akhir, plus arahan pakai menu sync manual.
  let peringatanSync = '';
  try {
    if (typeof syncPenjualanFromSparepart_ === 'function') {
      syncPenjualanFromSparepart_(false); // silent, gak munculin alert
    }
  } catch (err) {
    Logger.log('Auto-sync PENJUALAN dari SPAREPART error: ' + err);
    peringatanSync = '\n\n⚠️ Invoice ini BELUM masuk PENJUALAN (' + err.message + ').\n' +
      'Jalankan menu SANFELO > "🔄 Sync Penjualan dari Sparepart" sebelum input pembayaran.';
  }

  // Bersihin penanda mode edit karena transaksi ini sudah kelar diproses
  docProps.deleteProperty('SPAREPART_EDIT_NOSJ');
  docProps.deleteProperty('SPAREPART_EDIT_STATUS');

  // Reset form (cell oranye) biar siap dipakai transaksi berikutnya
  resetFormSparepart_(shSJ, shINV);

  // Generate & tulis nomor berikutnya sebagai "nomor siap pakai" buat transaksi selanjutnya
  const nextNoSJ  = getNextSparepartNumber_(shSPR, shSYS, 'SJ');
  const nextNoInv = getNextSparepartNumber_(shSPR, shSYS, 'INV');
  shSJ.getRange('H6').setValue(nextNoSJ);
  shINV.getRange('H5').setValue(nextNoInv);
  SpreadsheetApp.flush();

  SpreadsheetApp.getUi().alert(
    (isEditMode ? 'Data lama utk No SJ ini sudah ditimpa.\n' : '') +
    rowsToWrite.length + ' baris barang berhasil disimpan ke SPAREPART.\n' +
    'No SJ  : ' + state.noSJ + '\n' +
    'No Inv : ' + state.noInv + '\n\n' +
    'Nomor siap untuk transaksi berikutnya:\n' +
    'No SJ  : ' + nextNoSJ + '\n' +
    'No Inv : ' + nextNoInv +
    peringatanSync
  );
}

/* ---------- Generate nomor SJ/INV Sparepart, auto increment per bulan ----------
   ⬅️ BARU: cross-check ke SPAREPART sekarang baca FRESH lewat
   bacaSheetGateway_ (langsung ke database, bukan mirror lokal yang bisa
   delay) — biar nomor gak tabrakan kalau ada transaksi lain yang baru
   aja disimpan dari device/tab lain. Parameter shSPR TETAP diterima
   (dipakai pemanggil lain di file ini) tapi TIDAK dipakai lagi buat
   scan nomor -- dibiarkan di signature biar pemanggil gak perlu diubah. */
function getNextSparepartNumber_(shSPR, shSYS, type) {
  const cfg = (type === 'SJ')
    ? { sysRow: 6, key: 'LAST_SPR',   prefix: 'SPR',   scanCol: 1 }  // SPAREPART kolom A = NO SJ
    : { sysRow: 7, key: 'LAST_SPRIN', prefix: 'SPRIN', scanCol: 2 }; // SPAREPART kolom B = NO INV

  const now = new Date();
  const monthKey = Utilities.formatDate(now, 'Asia/Jakarta', 'yyyy-MM');

  // baca counter yang tersimpan di SYSTEM (tetap lokal, tidak berubah)
  const sysVals = shSYS.getRange(cfg.sysRow, 1, 1, 3).getValues()[0]; // [key, value, monthKey]
  // ⬅️ FIX: Sheets otomatis ngubah teks "2026-08" di kolom C jadi TANGGAL
  // (01/08/2026), jadi dulu sysVals[2] === monthKey selalu false. Samain
  // formatnya dulu sebelum dibandingin.
  const sysMonth = (sysVals[2] instanceof Date)
    ? Utilities.formatDate(sysVals[2], 'Asia/Jakarta', 'yyyy-MM')
    : String(sysVals[2] || '').trim();
  const storedCounter = (sysMonth === monthKey && typeof sysVals[1] === 'number') ? sysVals[1] : 0;

  // cross-check ke SPAREPART ASLI di database (bukan mirror lokal):
  // cari angka terbesar untuk bulan berjalan (jaga-jaga kalau ada input
  // manual / counter di SYSTEM ketinggalan)
  const values = bacaSheetGateway_('SPAREPART', 2, cfg.scanCol, 1);
  let maxFound = 0;
  const pattern = new RegExp('^' + cfg.prefix + '-' + monthKey + '-(\\d+)$');
  values.forEach(function (r) {
    const v = r[0];
    if (!v) return;
    const m = String(v).match(pattern);
    if (m) {
      const n = parseInt(m[1], 10);
      if (n > maxFound) maxFound = n;
    }
  });

  const nextCounter = Math.max(storedCounter, maxFound) + 1;
  const nextNumber = cfg.prefix + '-' + monthKey + '-' + String(nextCounter).padStart(3, '0');

  // simpan counter terbaru ke SYSTEM (A=key, B=angka, C=bulan aktif)
  shSYS.getRange(cfg.sysRow, 1, 1, 3).setValues([[cfg.key, nextCounter, monthKey]]);

  return nextNumber;
}

/* ---------- true kalau nomor SPR/SPRIN-nya dari bulan yang SUDAH LEWAT ----------
   Cuma ngecek format SPR-yyyy-MM-NNN / SPRIN-yyyy-MM-NNN; nomor yang
   formatnya lain (atau kosong) dianggap bukan basi, dibiarkan apa adanya. */
function nomorSparepartBulanLama_(nomor) {
  const m = String(nomor || '').trim().match(/^(SPR|SPRIN)-(\d{4}-\d{2})-\d+$/);
  if (!m) return false;
  return m[2] !== Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM');
}

/* ---------- Bangun map semua data PEMBELIAN per nama barang ----------
   { KEY_BARANG_UPPER: [ {namaBarang, supplier, total, qty, unitPrice}, ... ] }
   Satu barang bisa punya beberapa entry (beda supplier/harga/tanggal beli). ---------- */
// ⬅️ BARU (migrasi PEMBELIAN, Oktober 2026): baca FRESH dari database
// lewat bacaSheetGateway_, bukan dari sheet PEMBELIAN lokal lagi (itu
// sekarang cuma mirror IMPORTRANGE) -- biar lookup HPP gak ketinggalan
// PEMBELIAN yang baru aja diposting (mis. Tanda Terima Barang beberapa
// saat sebelum Simpan Sparepart ini). Parameter shPMB dihapus, gak
// diperlukan lagi.
function buildPembelianMap_() {
  const map = {};
  const data = bacaSheetGateway_("PEMBELIAN", 2, 1, 10); // A..J
  if (data.length === 0) return map;

  data.forEach(function (row) {
    const namaBarang = row[7]; // H = NAMA BARANG
    const supplier   = row[3]; // D = SUPPLIER
    const qty        = row[8]; // I = QTY
    const total      = row[9]; // J = TOTAL

    if (!namaBarang) return;
    if (typeof total !== 'number' || typeof qty !== 'number' || qty === 0) return;

    const key = String(namaBarang).trim().toUpperCase();
    const unitPrice = total / qty;

    if (!map[key]) map[key] = [];
    map[key].push({
      namaBarang: String(namaBarang).trim(),
      supplier: supplier,
      total: total,
      qty: qty,
      unitPrice: unitPrice
    });
  });

  return map;
}

/* ---------- Bangun HTML dialog pilihan data PEMBELIAN utk barang yang datanya dobel ---------- */
function buildHppDialogHtml_(token, ambiguousGroups) {
  const dataJson = JSON.stringify(ambiguousGroups);
  const tokenJson = JSON.stringify(token);

  return HtmlService.createHtmlOutput(
    '<style>' +
    'body{font-family:Arial,sans-serif;font-size:13px;padding:10px;}' +
    '.group{margin-bottom:16px;border-bottom:1px solid #ddd;padding-bottom:10px;}' +
    '.group h4{margin:0 0 6px 0;}' +
    'label{display:block;margin:4px 0;cursor:pointer;}' +
    '.opt-row{display:flex;gap:8px;align-items:center;}' +
    'button{margin-top:10px;padding:6px 14px;}' +
    '#err{color:red;margin-top:8px;}' +
    '</style>' +
    '<div id="container"></div>' +
    '<div id="err"></div>' +
    '<button id="btnSubmit">Simpan Pilihan</button>' +
    '<script>' +
    'var token = ' + tokenJson + ';' +
    'var groups = ' + dataJson + ';' +
    'var container = document.getElementById("container");' +
    'groups.forEach(function(g, gi){' +
    '  var div = document.createElement("div");' +
    '  div.className = "group";' +
    '  var h = document.createElement("h4");' +
    '  h.textContent = "Barang: " + g.namaBarang + " - pilih data pembelian yang dipakai:";' +
    '  div.appendChild(h);' +
    '  g.entries.forEach(function(e, ei){' +
    '    var row = document.createElement("label");' +
    '    row.className = "opt-row";' +
    '    var radio = document.createElement("input");' +
    '    radio.type = "radio";' +
    '    radio.name = "group_" + gi;' +
    '    radio.value = ei;' +
    '    if (ei === 0) radio.checked = true;' +
    '    row.appendChild(radio);' +
    '    var txt = document.createElement("span");' +
    '    txt.textContent = "Supplier: " + (e.supplier || "-") + " | Total: " + e.total + " | Qty: " + e.qty + " | Harga satuan: " + e.unitPrice.toFixed(2);' +
    '    row.appendChild(txt);' +
    '    div.appendChild(row);' +
    '  });' +
    '  container.appendChild(div);' +
    '});' +
    'document.getElementById("btnSubmit").addEventListener("click", function(){' +
    '  var selections = {};' +
    '  groups.forEach(function(g, gi){' +
    '    var checked = document.querySelector("input[name=\\"group_" + gi + "\\"]:checked");' +
    '    selections[g.key] = checked ? parseInt(checked.value, 10) : 0;' +
    '  });' +
    '  document.getElementById("btnSubmit").disabled = true;' +
    '  google.script.run' +
    '    .withSuccessHandler(function(){ google.script.host.close(); })' +
    '    .withFailureHandler(function(err){' +
    '      document.getElementById("err").textContent = "Gagal: " + err.message;' +
    '      document.getElementById("btnSubmit").disabled = false;' +
    '    })' +
    '    .resolveSparepartSelections_(token, selections);' +
    '});' +
    '</script>'
  ).setWidth(520).setHeight(450);
}

/* ---------- Kosongin form SJ_SPR & INV_SPR setelah disimpan ----------
   Cuma cell yang warnanya ORANGE (input manual) yang di-clear.
   H6 (No SJ) & H5 di INV_SPR (No Inv) TIDAK di-clear karena itu hasil
   auto-generate yang baru ditulis. Kolom B/E/F di INV_SPR juga tidak
   di-clear karena isinya formula yang narik dari SJ_SPR (otomatis ikut
   kosong begitu SJ_SPR-nya dikosongin). ---------- */
function resetFormSparepart_(shSJ, shINV) {
  // SJ_SPR - cell oranye
  shSJ.getRange('H3').clearContent();     // Metode Pembayaran
  shSJ.getRange('H5').clearContent();     // Tgl
  shSJ.getRange('H7').clearContent();     // No Kend
  shSJ.getRange('A5').clearContent();     // Nama Toko
  shSJ.getRange('A6').clearContent();     // Alamat
  shSJ.getRange('B9:G16').clearContent(); // baris barang (Barang, Qty, SAT, Keterangan)

  // INV_SPR - cell oranye
  shINV.getRange('H9').clearContent();      // Jatuh tempo
  shINV.getRange('H7').clearContent();      // Refrensi
  shINV.getRange('C20').clearContent();     // Nama Sales
  shINV.getRange('C25').clearContent();     // Keterangan
  shINV.getRange('G11:G18').clearContent(); // Total (Rp) per baris
}

function resetbos(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const shSJ  = ss.getSheetByName('SJ_SPR');
  const shINV = ss.getSheetByName('INV_SPR');
  const shSPR = ss.getSheetByName('SPAREPART');
  const shSYS = ss.getSheetByName('SYSTEM');

  if (!shSJ || !shINV || !shSPR || !shSYS) {
    ui.alert('Ada sheet yang gak ketemu. Pastikan SJ_SPR, INV_SPR, SPAREPART, SYSTEM semua ada.');
    return;
  }

  // Jangan reset kalau form lagi dalam mode EDIT (draft dari SPAREPART),
  // soalnya No SJ yang lagi diedit harus tetep sama biar pas disimpan
  // baris lama ke-timpa, bukan malah nambah baris baru.
  const docProps = PropertiesService.getDocumentProperties();
  const editNoSJ = docProps.getProperty('SPAREPART_EDIT_NOSJ');
  if (editNoSJ) {
    ui.alert('Form ini lagi dalam mode EDIT untuk No SJ ' + editNoSJ + '.\n' +
      'Reset dibatalkan biar data lama gak ketiban transaksi baru.\n' +
      'Selesaikan/batalkan edit dulu sebelum reset.');
    return;
  }

  // 1. Kosongin cell oranye (barang, header, tanggal, dll)
  resetFormSparepart_(shSJ, shINV);

  // 2. Generate ulang No SJ & No INV sesuai bulan berjalan, tulis ke form
  const nextNoSJ  = getNextSparepartNumber_(shSPR, shSYS, 'SJ');
  const nextNoInv = getNextSparepartNumber_(shSPR, shSYS, 'INV');
  shSJ.getRange('H6').setValue(nextNoSJ);
  shINV.getRange('H5').setValue(nextNoInv);
  SpreadsheetApp.flush();

  ui.alert(
    'Form direset.\n\n' +
    'No SJ  : ' + nextNoSJ + '\n' +
    'No Inv : ' + nextNoInv
  );
}

  // ======================= STOK_BARANG.js (salinan persis) =======================
/* ======================================================
   STOK_BARANG.js   (dipakai SAMA PERSIS di inventory_SANFELO & inventory_SL)
   ------------------------------------------------------
   Halaman "📦 Stok Barang" (popup besar, file StokBarang.html).

   CARA HITUNG:
     barang MASUK  = sheet PEMBELIAN        (TGL BELI, KODE, NAMA BARANG, QTY)
     barang KELUAR = sheet DETAIL_TRANSAKSI (TANGGAL, KODE, BARANG, QTY)
                   + sheet SPAREPART        (TANGGAL, BARANG, QTY)
     stok pada suatu tanggal = total masuk - total keluar s/d tanggal itu.
   Belum ada stok awal (DATA_BERSIH kolom STOK masih 0 semua), jadi
   stok MINUS artinya ada barang keluar yang pembeliannya belum tercatat
   (atau dicatat dengan KODE lain) -- ditandai "Perlu dicek" di halaman.

   Kolom dicari dari JUDUL di baris 1 (bukan nomor kolom), karena urutan
   kolom DETAIL_TRANSAKSI SANFELO & SL beda. Di SANFELO sheet-sheet ini
   salinan IMPORTRANGE dari database -- cukup dibaca lokal, ini cuma
   tampilan (tidak ada yang ditulis).

   Barang dikelompokkan per KODE. Baris yang KODE-nya bukan kode barang
   (mis. "SPR", "SPAREPART", "KATALOG", kosong) dikelompokkan per NAMA
   BARANG, masuk kategori "Sparepart & lainnya".

   CARA PAKAI: menu 🏢 SANFELO > 📦 Stok Barang
   ====================================================== */

function bukaStokBarang() {
  const html = HtmlService.createHtmlOutputFromFile('StokBarang')
    .setWidth(1200)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, '📦 Stok Barang');
}

/** Dipanggil dari StokBarang.html lewat google.script.run. */
function ambilDataStokBarang() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = Session.getScriptTimeZone();
  const items = {};   // kunci -> { k: kode tampil, n: nama, m: merk, j: 'produk' | 'lain' }
  const gerak = [];   // [tgl yyyy-MM-dd, kunci, masuk, keluar, no dokumen, pihak, sumber]
  const catatan = [];
  const dilewati = { tanggal: 0, contoh: [] };

  // master nama & merk dari DATA_BERSIH
  const master = {};
  const db = stokBacaSheet_(ss, 'DATA_BERSIH', ['KODE', 'NAMA_PRODUK'], catatan);
  if (db) {
    db.rows.forEach(function (r) {
      const kode = String(r[db.idx.KODE] || '').trim().toUpperCase();
      if (!kode) return;
      master[kode] = {
        nama: String(r[db.idx.NAMA_PRODUK] || '').trim(),
        merk: db.idx.MERK !== undefined ? String(r[db.idx.MERK] || '').trim() : ''
      };
    });
  }

  const daftarItem = function (kunci, kode, nama) {
    if (items[kunci]) return;
    const isProduk = kunci.indexOf('LAIN:') !== 0;
    const m = isProduk ? master[kunci] : null;
    items[kunci] = {
      k: isProduk ? kunci : '',
      n: (m && m.nama) || String(nama || '').trim() || kunci,
      m: (m && m.merk) || '',
      j: isProduk ? 'produk' : 'lain'
    };
  };

  const tambah = function (sumber, cfg, arah) {
    const t = stokBacaSheet_(ss, sumber, cfg.wajib, catatan);
    if (!t) return;
    t.rows.forEach(function (r) {
      const qty = Number(r[t.idx.QTY]);
      if (!qty || isNaN(qty)) return;
      const kode = cfg.kode ? r[t.idx[cfg.kode]] : '';
      const nama = r[t.idx[cfg.nama]];
      const kunci = stokKunci_(kode, nama);
      if (!kunci) return;
      const ref = cfg.ref && t.idx[cfg.ref] !== undefined ? String(r[t.idx[cfg.ref]] || '').trim() : '';
      const tgl = stokTanggal_(r[t.idx[cfg.tgl]], tz);
      if (!tgl) {
        dilewati.tanggal++;
        if (dilewati.contoh.length < 5) dilewati.contoh.push(sumber + ' ' + (ref || String(nama || '')));
        return;
      }
      daftarItem(kunci, kode, nama);
      const pihak = cfg.pihak && t.idx[cfg.pihak] !== undefined ? String(r[t.idx[cfg.pihak]] || '').trim() : '';
      gerak.push([tgl, kunci, arah === 'masuk' ? qty : 0, arah === 'keluar' ? qty : 0, ref, pihak, sumber]);
    });
  };

  tambah('PEMBELIAN', { wajib: ['TGL BELI', 'NAMA BARANG', 'QTY'], tgl: 'TGL BELI', kode: 'KODE', nama: 'NAMA BARANG', ref: 'NO PB', pihak: 'SUPPLIER' }, 'masuk');
  tambah('DETAIL_TRANSAKSI', { wajib: ['TANGGAL', 'BARANG', 'QTY'], tgl: 'TANGGAL', kode: 'KODE', nama: 'BARANG', ref: 'NO INV', pihak: 'NAMA TOKO' }, 'keluar');
  tambah('SPAREPART', { wajib: ['TANGGAL', 'BARANG', 'QTY'], tgl: 'TANGGAL', nama: 'BARANG', ref: 'NO INV', pihak: 'NAMA TOKO' }, 'keluar');

  if (dilewati.tanggal) {
    catatan.push(dilewati.tanggal + ' baris dilewati karena tanggalnya tidak terbaca (contoh: ' + dilewati.contoh.join(', ') + ').');
  }
  gerak.sort(function (a, b) { return a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : 0; });

  return {
    nama: ss.getName(),
    dibuat: Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm'),
    hariIni: Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd'),
    items: items,
    gerak: gerak,
    catatan: catatan
  };
}

/* ---------- helper ---------- */

/** Baca sheet + cari index kolom dari judul baris 1. null kalau sheet/kolom wajib tidak ada. */
function stokBacaSheet_(ss, nama, wajib, catatan) {
  const sh = ss.getSheetByName(nama);
  if (!sh || sh.getLastRow() < 2) {
    catatan.push('Sheet ' + nama + ' tidak ada / masih kosong.');
    return null;
  }
  const data = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
  const idx = {};
  data[0].forEach(function (h, i) {
    const k = String(h || '').trim().toUpperCase();
    if (k && idx[k] === undefined) idx[k] = i;
  });
  const kurang = wajib.filter(function (k) { return idx[k] === undefined; });
  if (kurang.length) {
    catatan.push('Sheet ' + nama + ' tidak punya kolom: ' + kurang.join(', ') + ' -- dilewati.');
    return null;
  }
  return { rows: data.slice(1), idx: idx };
}

/**
 * Kunci barang: KODE kalau bentuknya kode barang asli (SA001-160, SN015,
 * DP012, CS004 ...). Selain itu (SPR, SPAREPART, KATALOG, plat nomor,
 * kosong) dikelompokkan per NAMA BARANG.
 */
function stokKunci_(kode, nama) {
  const k = String(kode || '').trim().toUpperCase();
  if (/^[A-Z]{1,4}\d{2,4}(-\d+)?$/.test(k)) return k;
  const n = String(nama || '').trim().toUpperCase().replace(/\s+/g, ' ');
  return n ? 'LAIN:' + n : '';
}

/** Tanggal apa pun (Date, angka serial Sheets, ISO, dd/MM/yyyy) -> 'yyyy-MM-dd', atau null. */
function stokTanggal_(v, tz) {
  let d = null;
  if (v instanceof Date) d = v;
  else if (typeof v === 'number') {
    // angka serial = jam dinding lokal; ambil bagian tanggalnya saja lalu
    // format sebagai UTC biar gak geser hari karena zona waktu
    if (!(v > 20000 && v < 80000)) return null;
    return Utilities.formatDate(new Date(Date.UTC(1899, 11, 30) + Math.floor(v) * 86400000), 'UTC', 'yyyy-MM-dd');
  } else if (v) {
    const s = String(v).trim();
    const m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})/);
    if (m) d = new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]), 12);
    else { const p = new Date(s); if (!isNaN(p.getTime())) d = p; }
  }
  if (!d || isNaN(d.getTime())) return null;
  const y = Number(Utilities.formatDate(d, tz, 'yyyy'));
  if (y < 2020 || y > 2100) return null; // tanggal salah ketik (mis. tahun 20206)
  return Utilities.formatDate(d, tz, 'yyyy-MM-dd');
}


  // ======================= PettyCash.js (salinan persis) =======================
/* ==========================================================================
 * PETTY CASH -> DATABASE
 *
 * Semua transaksi dari popup sekarang ditulis ke sheet "PETTY CASH" di file
 * DATABASE (bukan lagi di file tempat script ini terpasang). File INVENTORY
 * cukup mengambil data dari file DATABASE lewat IMPORTRANGE.
 *
 * ⚠️ WAJIB: isi DATABASE_SPREADSHEET_ID di bawah dengan ID file DATABASE.
 *    ID = bagian di URL antara /d/ dan /edit
 *    contoh: https://docs.google.com/spreadsheets/d/INI_IDNYA/edit
 * ========================================================================== */

var DATABASE_SPREADSHEET_ID = '1UPHKIUAornJxCZ_9CbEJAUTeqfhcKJPJ2la_6GMSqfo';
var KAS_KECIL_SHEET_NAME = 'PETTY CASH';
var KAS_KECIL_FIRST_DATA_ROW = 4;  // baris 1-3 header, data mulai row 4
var KAS_KECIL_LAST_COL = 9;        // A..I (H = SALDO, I = catatan/nama)
var SALDO_COL = 8;                 // kolom H
var DATE_FORMAT = 'dd/mm/yyyy';    // format tampilan tanggal di kolom A

/**
 * Mapping kategori (dari form) -> kolom di sheet
 */
var COLUMN_MAP = {
  'PIUTANG_USAHA': 'C',
  'LAINNYA': 'D',
  'ONGKOS_KIRIM': 'E',
  'OPERASIONAL': 'F',
  'ADMIN': 'G'
};

// Kategori DEBET (menambah saldo) -> kolom C, D
var DEBET_CATEGORIES = ['PIUTANG_USAHA', 'LAINNYA'];
// Kategori KREDIT (mengurangi saldo) -> kolom E, F, G
var KREDIT_CATEGORIES = ['ONGKOS_KIRIM', 'OPERASIONAL', 'ADMIN'];

/* ==========================================================================
 * POPUP
 * ========================================================================== */

/**
 * Menampilkan popup form (dipanggil dari menu atau tombol/gambar di sheet).
 * Kalau mau dipasang ke tombol/gambar: klik kanan gambar > Assign script >
 * isi dengan "showTransactionDialog"
 */
function showTransactionDialog() {
  var html = HtmlService.createHtmlOutputFromFile('PettyCashPopup')
    .setWidth(420)
    .setHeight(520);
  SpreadsheetApp.getUi().showModalDialog(html, 'Tambah Transaksi Kas');
}

/* ==========================================================================
 * AKSES KE FILE DATABASE
 * ========================================================================== */

/**
 * Ambil sheet PETTY CASH di file DATABASE. Selalu dicari eksplisit lewat ID,
 * jadi tidak tergantung tab/file mana yang lagi aktif.
 */
function getPettyCashSheet_() {
  if (!DATABASE_SPREADSHEET_ID || DATABASE_SPREADSHEET_ID.indexOf('ISI_ID') === 0) {
    throw new Error('DATABASE_SPREADSHEET_ID belum diisi di script.');
  }

  var ss;
  try {
    ss = SpreadsheetApp.openById(DATABASE_SPREADSHEET_ID);
  } catch (e) {
    throw new Error('Gagal membuka file DATABASE. Cek ID-nya dan pastikan akun ini punya akses edit. (' + e.message + ')');
  }

  var sheet = ss.getSheetByName(KAS_KECIL_SHEET_NAME);
  if (!sheet) {
    throw new Error('Sheet "' + KAS_KECIL_SHEET_NAME + '" tidak ditemukan di file DATABASE.');
  }
  return sheet;
}

/**
 * JALANKAN SEKALI dari Apps Script editor untuk tes koneksi sekaligus
 * memunculkan permintaan izin akses ke file DATABASE.
 */
function testKoneksiDatabase() {
  var sheet = getPettyCashSheet_();
  var lastRow = findLastDataRow_(sheet);
  var msg = 'Koneksi OK.\nFile: ' + sheet.getParent().getName() +
    '\nSheet: ' + sheet.getName() +
    '\nBaris data terakhir: ' + lastRow;
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { /* dijalankan dari editor tanpa UI */ }
}

/* ==========================================================================
 * INPUT TRANSAKSI
 * ========================================================================== */

/**
 * Dipanggil dari form saat submit banyak barang sekaligus dalam 1 tanggal.
 * payload = { tanggal: 'yyyy-mm-dd', items: [ { kategori, nominal, keterangan }, ... ] }
 *
 * - Semua item divalidasi DULU. Kalau ada 1 yang salah, tidak ada yang ditulis.
 * - Semua baris ditulis sekaligus (1x setValues), jadi cepat dan tidak setengah-setengah.
 * - Tanggal yang sama dengan grup di atasnya otomatis di-merge.
 * - Kalau tanggal yang diinput lebih lama dari data terakhir, data otomatis
 *   di-sort ulang dan SALDO dihitung ulang.
 *
 * Return: { success: true, results: [ { row, saldo }, ... ] } atau { success: false, error }
 */
function processTransactionBatch(payload) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
  } catch (e) {
    return { success: false, error: 'Sistem sedang dipakai orang lain, coba lagi beberapa detik lagi.' };
  }

  try {
    if (!payload || !payload.tanggal) {
      throw new Error('Tanggal wajib diisi');
    }
    if (!payload.items || payload.items.length === 0) {
      throw new Error('Tidak ada barang yang diinput');
    }

    var tanggal = parseTanggalPettyCash_(payload.tanggal);

    // 1. Validasi semua item dulu
    var parsedItems = [];
    for (var i = 0; i < payload.items.length; i++) {
      var item = payload.items[i];
      var columnLetter = COLUMN_MAP[item.kategori];
      if (!columnLetter) {
        throw new Error('Barang #' + (i + 1) + ': kategori tidak dikenali (' + item.kategori + ')');
      }
      var nominal = parseFloat(item.nominal);
      if (isNaN(nominal) || nominal <= 0) {
        throw new Error('Barang #' + (i + 1) + ': nominal tidak valid');
      }
      parsedItems.push({
        kategori: item.kategori,
        colIndex: columnToIndex_(columnLetter),
        nominal: nominal,
        keterangan: item.keterangan || ''
      });
    }

    // 2. Buka sheet di file DATABASE
    var sheet = getPettyCashSheet_();
    var first = KAS_KECIL_FIRST_DATA_ROW;
    var lastRow = findLastDataRow_(sheet);
    var firstNewRow = lastRow + 1;
    var numNew = parsedItems.length;
    var lastNewRow = firstNewRow + numNew - 1;

    // 3. Info baris terakhir yang sudah ada: tanggal grup & saldo terakhir
    var prevSaldo = 0;
    var lastGroup = null; // { date, topRow }
    if (lastRow >= first) {
      var existingA = sheet.getRange(first, 1, lastRow - first + 1, 1).getValues();
      var existingH = sheet.getRange(first, SALDO_COL, lastRow - first + 1, 1).getValues();

      for (var a = existingA.length - 1; a >= 0; a--) {
        if (existingA[a][0] instanceof Date) {
          lastGroup = { date: existingA[a][0], topRow: first + a };
          break;
        }
      }
      for (var h = existingH.length - 1; h >= 0; h--) {
        if (typeof existingH[h][0] === 'number') {
          prevSaldo = existingH[h][0];
          break;
        }
      }
    }

    // 4. Susun baris baru (A..I) + hitung saldo berjalan
    var newValues = [];
    var results = [];
    var saldo = prevSaldo;
    for (var n = 0; n < numNew; n++) {
      var p = parsedItems[n];
      var row = ['', p.keterangan, '', '', '', '', '', '', ''];
      row[p.colIndex - 1] = p.nominal;

      var isDebet = DEBET_CATEGORIES.indexOf(p.kategori) !== -1;
      saldo = isDebet ? (saldo + p.nominal) : (saldo - p.nominal);
      row[SALDO_COL - 1] = saldo;

      newValues.push(row);
      results.push({ row: firstNewRow + n, saldo: saldo });
    }

    // 5. Tentukan: lanjut grup tanggal yang lama (merge) atau mulai grup baru
    var topRow = firstNewRow;
    var sameGroup = lastGroup && isSameDate_(lastGroup.date, tanggal);
    if (sameGroup) {
      topRow = lastGroup.topRow;
    } else {
      newValues[0][0] = tanggal;
    }

    // 6. Tulis semuanya sekaligus
    sheet.getRange(firstNewRow, 1, numNew, KAS_KECIL_LAST_COL).setValues(newValues);
    sheet.getRange(firstNewRow, 3, numNew, 6).setNumberFormat('#,##0'); // C..H

    // Paksa format tanggal, biar tidak tampil sebagai angka serial (mis. 46280)
    sheet.getRange(topRow, 1, lastNewRow - topRow + 1, 1).setNumberFormat(DATE_FORMAT);

    // 7. Merge kolom tanggal untuk grup ini
    if (lastNewRow > topRow) {
      var mergeRange = sheet.getRange(topRow, 1, lastNewRow - topRow + 1, 1);
      mergeRange.breakApart();
      mergeRange.merge().setVerticalAlignment('middle');
    }

    // 8. Kalau urutan tanggal jadi tidak berurutan, rapikan + hitung ulang saldo
    try {
      var sorted = sortKasKecilByTanggal(sheet, false);
      if (sorted) {
        // posisi baris berubah, ambil ulang saldo akhir (baris terakhir data)
        var finalLast = findLastDataRow_(sheet);
        var finalSaldo = sheet.getRange(finalLast, SALDO_COL).getValue();
        results[results.length - 1].saldo = finalSaldo;
      }
    } catch (sortErr) {
      Logger.log('Sort setelah input transaksi gagal: ' + sortErr.message);
    }

    SpreadsheetApp.flush();
    return { success: true, results: results };

  } catch (err) {
    return { success: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Versi 1 barang (dipertahankan biar kompatibel kalau ada yang masih
 * memanggil processTransaction langsung).
 * data = { tanggal, kategori, nominal, keterangan }
 */
function processTransaction(data) {
  var res = processTransactionBatch({
    tanggal: data.tanggal,
    items: [{ kategori: data.kategori, nominal: data.nominal, keterangan: data.keterangan }]
  });
  if (!res.success) return res;
  return { success: true, row: res.results[0].row, saldo: res.results[0].saldo };
}

/* ==========================================================================
 * HELPER
 * ========================================================================== */

/**
 * Ubah 'yyyy-mm-dd' (dari <input type="date">) jadi Date jam 00:00 lokal.
 * Sengaja tidak pakai new Date('yyyy-mm-dd') karena itu dibaca UTC dan
 * bisa bergeser sehari tergantung timezone.
 */
function parseTanggalPettyCash_(value) {
  if (value instanceof Date) {
    return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  }
  var m = /^(\d{4})-(\d{2})-(\d{2})/.exec(String(value));
  if (!m) throw new Error('Format tanggal tidak valid: ' + value);
  var y = Number(m[1]), mo = Number(m[2]), d = Number(m[3]);
  var date = new Date(y, mo - 1, d);
  if (date.getFullYear() !== y || date.getMonth() !== mo - 1 || date.getDate() !== d) {
    throw new Error('Tanggal tidak valid: ' + value);
  }
  return date;
}

/**
 * Bandingkan dua Date hanya berdasarkan tahun/bulan/tanggal (tanpa jam).
 */
function isSameDate_(d1, d2) {
  return d1.getFullYear() === d2.getFullYear() &&
    d1.getMonth() === d2.getMonth() &&
    d1.getDate() === d2.getDate();
}

/**
 * Cari baris data terakhir (kolom A..G + I yang berisi). Kolom H (saldo)
 * sengaja tidak dihitung supaya rumus/sisa isi di H tidak mengacaukan.
 * Return 3 kalau belum ada data.
 */
function findLastDataRow_(sheet) {
  var first = KAS_KECIL_FIRST_DATA_ROW;
  var lastRow = sheet.getLastRow();
  if (lastRow < first) return first - 1;

  var values = sheet.getRange(first, 1, lastRow - first + 1, KAS_KECIL_LAST_COL).getValues();
  for (var i = values.length - 1; i >= 0; i--) {
    for (var c = 0; c < KAS_KECIL_LAST_COL; c++) {
      if (c === SALDO_COL - 1) continue;
      var cell = values[i][c];
      if (cell !== '' && cell !== null) return first + i;
    }
  }
  return first - 1;
}

/**
 * Ubah huruf kolom (C, D, E, ...) jadi index angka (3, 4, 5, ...)
 */
function columnToIndex_(letter) {
  return letter.charCodeAt(0) - 64; // A=1, B=2, C=3, dst
}

/* ==========================================================================
 * SORT BY TANGGAL + HITUNG ULANG SALDO
 * ========================================================================== */

/**
 * JALANKAN MANUAL kalau mau merapikan data yang sudah ada:
 * Apps Script editor -> pilih "sortKasKecilByTanggalManual" -> Run.
 * (Selalu sort & hitung ulang saldo, walaupun urutannya sudah rapi.)
 */
function sortKasKecilByTanggalManual() {
  sortKasKecilByTanggal(null, true);
  var msg = 'Selesai! Data PETTY CASH di file DATABASE sudah diurutkan berdasarkan tanggal dan SALDO sudah dihitung ulang.';
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { /* tanpa UI */ }
}

/**
 * Sort data PETTY CASH berdasarkan TANGGAL (kolom A) ascending, lalu hitung
 * ulang SALDO (kolom H) dari atas ke bawah. Baris dengan tanggal sama
 * di-merge lagi jadi satu grup.
 *
 * @param {Sheet=}   sheetParam - sheet PETTY CASH di DATABASE (kalau kosong dicari otomatis)
 * @param {boolean=} force      - true = selalu sort & hitung ulang.
 *                                false/kosong = hanya kalau urutan tanggal belum rapi.
 * @return {boolean} true kalau sheet benar-benar ditulis ulang.
 */
function sortKasKecilByTanggal(sheetParam, force) {
  var sheet = sheetParam || getPettyCashSheet_();
  var first = KAS_KECIL_FIRST_DATA_ROW;

  var lastRow = findLastDataRow_(sheet);
  if (lastRow < first) return false; // belum ada data

  var numRows = lastRow - first + 1;

  // 1. Baca A..I, isi tanggal yang kosong (baris bawah dalam grup merge)
  //    dengan tanggal dari baris di atasnya.
  var values = sheet.getRange(first, 1, numRows, KAS_KECIL_LAST_COL).getValues();
  var lastDate = null;
  for (var i = 0; i < values.length; i++) {
    if (values[i][0] instanceof Date) {
      lastDate = values[i][0];
    } else if (lastDate) {
      values[i][0] = lastDate;
    }
  }

  // 2. Kalau sudah berurutan dan tidak dipaksa, tidak usah diapa-apakan
  if (!force && isSortedByDate_(values)) return false;

  // 3. Lepas merge kolom tanggal supaya penulisan ulang aman
  sheet.getRange(first, 1, numRows, 1).breakApart();

  // 4. Sort stabil: tanggal sama tetap urut sesuai input semula
  var indexed = values.map(function (row, idx) { return { row: row, idx: idx }; });
  indexed.sort(function (a, b) {
    var da = a.row[0] instanceof Date ? a.row[0].getTime() : 0;
    var db = b.row[0] instanceof Date ? b.row[0].getTime() : 0;
    if (da !== db) return da - db;
    return a.idx - b.idx;
  });
  var sorted = indexed.map(function (x) { return x.row; });

  // 5. Hitung ulang SALDO: (DEBET C+D) - (KREDIT E+F+G), berurutan dari atas
  var running = 0;
  sorted.forEach(function (row) {
    running += num_(row[2]) + num_(row[3]) - num_(row[4]) - num_(row[5]) - num_(row[6]);
    row[SALDO_COL - 1] = running;
  });

  // 6. Tulis balik, lalu merge ulang baris-baris dengan tanggal sama
  sheet.getRange(first, 1, numRows, KAS_KECIL_LAST_COL).setValues(sorted);
  sheet.getRange(first, 1, numRows, 1).setNumberFormat(DATE_FORMAT);
  remergeSameDates_(sheet, first, lastRow);
  return true;
}

function num_(v) {
  return typeof v === 'number' ? v : 0;
}

/**
 * Cek apakah tanggal (kolom A, sudah di-fill) sudah urut naik.
 */
function isSortedByDate_(values) {
  var prev = -Infinity;
  for (var i = 0; i < values.length; i++) {
    var t = values[i][0] instanceof Date ? values[i][0].getTime() : 0;
    if (t < prev) return false;
    prev = t;
  }
  return true;
}

/**
 * Merge ulang baris-baris berurutan yang tanggalnya sama di kolom A.
 */
function remergeSameDates_(sheet, firstRow, lastRow) {
  var numRows = lastRow - firstRow + 1;
  var dateValues = sheet.getRange(firstRow, 1, numRows, 1).getValues();

  var groupStart = 0;
  for (var i = 1; i <= numRows; i++) {
    var sameAsStart = i < numRows &&
      dateValues[i][0] instanceof Date &&
      dateValues[groupStart][0] instanceof Date &&
      isSameDate_(dateValues[i][0], dateValues[groupStart][0]);

    if (!sameAsStart) {
      var groupSize = i - groupStart;
      if (groupSize > 1) {
        sheet.getRange(firstRow + groupStart, 1, groupSize, 1)
          .merge()
          .setVerticalAlignment('middle');
      }
      groupStart = i;
    }
  }
}

  // ======================= ReprintSJ.js (salinan persis) =======================
/* ======================================================
   ReprintSJ.gs
   - Reprint SJ + biaya tambahan pakai No SJ (syarat: No PO / H3 kosong)
   - Data diambil dari TRANSAKSI (header) + DETAIL_TRANSAKSI (item & biaya tambahan)
   - Filter DETAIL_TRANSAKSI PERSIS by No SJ (bukan No INV), supaya kalau 1 No INV
     dipakai > 1 No SJ, reprint cuma nampilin barang dari SJ yang diminta.
   - Harga/Diskon/Total di INVOICE diisi LANGSUNG dari nilai tersimpan di
     DETAIL_TRANSAKSI (snapshot hasil simpanDraft), BUKAN dihitung ulang lewat
     formula HARGATOKO. Formula itu sudah tidak relevan dipakai di Invoice.

   ⬅️ CATATAN: index kolom DETAIL_TRANSAKSI di file ini sudah disesuaikan
   dengan kolom baru "TANGGAL" & "TEMPO" yang disisipkan di posisi C/D
   (lihat REVISI_HANDLER.gs untuk pemetaan lengkap terbaru).
   ====================================================== */

function sjInvoiceFormulaTemplates_() {
  return {
    sjC: function(r) {
      return `=IFERROR(VLOOKUP(B${r};DATA_BERSIH!$A:$B;2;FALSE);"")`;
    },
    invB: function(ir) {
      const sjRow = ir - 2;
      return `=IFERROR(XLOOKUP('SURAT JALAN'!B${sjRow};DATA_BERSIH!A:A;DATA_BERSIH!B:B);"")`;
    }
  };
}

// Cuma restore formula NAMA BARANG (sjC di SJ, invB di INVOICE).
// Harga/Diskon/Total INVOICE (G/H/I) diisi manual via setValue, bukan formula.
function restoreSJInvoiceFormulas_(fromRowSJ, toRowSJ) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sj = ss.getSheetByName("SURAT JALAN");
  const inv = ss.getSheetByName("INVOICE");
  const t = sjInvoiceFormulaTemplates_();

  for (let r = fromRowSJ; r <= toRowSJ; r++) {
    sj.getRange(r, 3).setFormula(t.sjC(r)); // C = Barang (SJ)

    const ir = r + 2; // baris invoice terkait (offset tetap +2, sama kayak logic asli)
    inv.getRange(ir, 2).setFormula(t.invB(ir)); // B = Barang (Invoice)
  }
}

function cetakUlangSJDariNoSJ_(noSJ) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const sj  = ss.getSheetByName("SURAT JALAN");
  const inv = ss.getSheetByName("INVOICE");
  // ⬅️ MIGRASI (Oktober 2026): TRANSAKSI & DETAIL_TRANSAKSI dibaca FRESH dari
  // database. Baris judul ditambahkan di depan supaya indeks loop lama
  // (mulai i = 1) tetap sama persis.
  const ui  = SpreadsheetApp.getUi();

  noSJ = String(noSJ).trim();
  if (!noSJ) return;

  // --- cari header dari TRANSAKSI ---
  const trxData = [[]].concat(bacaTransaksiDB_());
  let header = null;
  for (let i = 1; i < trxData.length; i++) {
    if (String(trxData[i][1]).trim() === noSJ) { // B = NO SJ
      header = {
        noInv: trxData[i][2],      // C
        tgl: trxData[i][3],        // D
        customer: trxData[i][4],   // E
        daerah: trxData[i][5],     // F
        noPol: trxData[i][10],     // K
        sales: trxData[i][11],     // L
        method: trxData[i][9],    // N
        keterangan: trxData[i][15] // P
      };
      break;
    }
  }
  if (!header) {
    ui.alert('No SJ "' + noSJ + '" tidak ditemukan di TRANSAKSI.');
    return;
  }

  // --- ambil item & biaya tambahan dari DETAIL_TRANSAKSI, filter PERSIS by No SJ ini ---
  const detData = [[]].concat(bacaDetailDB_());
  const barangRows = [];
  const biayaRows = [];

  for (let i = 1; i < detData.length; i++) {
    if (String(detData[i][0]).trim() !== noSJ) continue; // A = NO SJ
    const kode      = detData[i][5];  // F = KODE
    const barang    = detData[i][6];  // G = BARANG
    const qty       = detData[i][7];  // H = QTY
    const sat       = detData[i][8];  // I = SATUAN
    const ketItem   = detData[i][9];  // J = KETERANGAN
    const harga     = detData[i][12]; // M = HARGA
    const diskon    = detData[i][13]; // N = DISKON
    const hargaNet  = detData[i][14]; // O = HARGA NET

    if (kode) {
      barangRows.push({
        kode: kode,
        qty: qty,
        sat: sat,
        ket: ketItem,
        harga: Number(harga) || 0,
        diskon: Number(diskon) || 0,
        hargaNet: Number(hargaNet) || 0
      });
    } else {
      biayaRows.push({ keterangan: barang, nominal: hargaNet });
    }
  }

  if (barangRows.length === 0 && biayaRows.length === 0) {
    ui.alert('Tidak ada data DETAIL_TRANSAKSI untuk No SJ "' + noSJ + '".');
    return;
  }

  const totalNeeded = barangRows.length + biayaRows.length;

  // --- COLLAPSE dulu ke kapasitas dasar, biar sisa barang/biaya
  // dari SJ lama (kalau sebelumnya sempat expand) kebersihan total ---
  const baseRows = getEffectiveBaseRows_(SJ_ITEM_MAX_ROWS);

  pastikanKapasitasItem_(sj, SJ_ITEM_START_ROW, baseRows, baseRows, SJ_FOOTER_COL_START, SJ_FOOTER_COL_END, SJ_FOOTER_TEKS);
  pastikanKapasitasItem_(inv, INV_DISPLAY_START_ROW, baseRows, baseRows, INV_FOOTER_COL_START, INV_FOOTER_COL_END, INV_FOOTER_TEKS);

  sj.getRange(SJ_ITEM_START_ROW, 2, baseRows, 1).clearContent(); // B kode
  sj.getRange(SJ_ITEM_START_ROW, 5, baseRows, 1).clearContent(); // E qty
  sj.getRange(SJ_ITEM_START_ROW, 6, baseRows, 1).clearContent(); // F sat
  sj.getRange(SJ_ITEM_START_ROW, 7, baseRows, 1).clearContent(); // G keterangan

  inv.getRange(INV_DISPLAY_START_ROW, 7, baseRows, 3).clearContent(); // G/H/I Harga/Discount/Total

  // pulihkan formula nama barang di kapasitas dasar (hapus override biaya tambahan lama)
  restoreSJInvoiceFormulas_(SJ_ITEM_START_ROW, SJ_ITEM_START_ROW + baseRows - 1);

  // --- EXPAND kalau item + biaya tambahan melebihi kapasitas dasar ---
  pastikanKapasitasItem_(sj, SJ_ITEM_START_ROW, baseRows, totalNeeded, SJ_FOOTER_COL_START, SJ_FOOTER_COL_END, SJ_FOOTER_TEKS);
  pastikanKapasitasItem_(inv, INV_DISPLAY_START_ROW, baseRows, totalNeeded, INV_FOOTER_COL_START, INV_FOOTER_COL_END, INV_FOOTER_TEKS);

  if (totalNeeded > baseRows) {
    // baris baru hasil expand cuma dapet FORMAT (dari pastikanKapasitasItem_),
    // formula nama barang harus di-set ulang manual di sini
    restoreSJInvoiceFormulas_(SJ_ITEM_START_ROW + baseRows, SJ_ITEM_START_ROW + totalNeeded - 1);
  }

  // --- header ---
  sj.getRange("A5").setValue(header.customer);
  sj.getRange("A6").setValue(header.daerah);
  setTanggalSaja_(sj.getRange("H5"), header.tgl);
  sj.getRange("H7").setValue(header.noPol); // TRANSAKSI kolom K

  inv.getRange("H6").setValue(header.noInv);
  setTanggalSaja_(inv.getRange("H7"), header.tgl);
  inv.getRange("I3").setValue(header.method); // TRANSAKSI kolom N

  // --- isi barang normal: kode+qty+sat+ket ke SJ, harga/diskon/total LANGSUNG ke INVOICE ---
  let r = SJ_ITEM_START_ROW;
  barangRows.forEach(function(item) {
    sj.getRange(r, 2).setValue(item.kode);
    sj.getRange(r, 5).setValue(item.qty);
    sj.getRange(r, 6).setValue(item.sat); // F = Satuan
    sj.getRange(r, 7).setValue(item.ket); // G = Keterangan per-item

    const invRow = r + 2;
    inv.getRange(invRow, 7).setValue(item.harga);    // G = Harga
    inv.getRange(invRow, 8).setValue(item.diskon);   // H = Discount
    inv.getRange(invRow, 9).setValue(item.hargaNet); // I = Total

    r++;
  });

  // --- isi biaya tambahan tepat di bawah barang: override manual ---
  biayaRows.forEach(function(item) {
    const invRow = r + 2;
    sj.getRange(r, 3).setValue(item.keterangan);        // SJ Barang
    inv.getRange(invRow, 2).setValue(item.keterangan);  // INV Barang
    inv.getRange(invRow, 7).setValue(item.nominal);     // INV Harga
    inv.getRange(invRow, 8).setValue(0);                // INV Discount dipaksa 0
    inv.getRange(invRow, 9).setValue(item.nominal);     // INV Total
    r++;
  });

  // --- isi Sales & Keterangan di INVOICE, posisinya dinamis ngikutin footer ---
  const footerRowInvFinal = cariBarisAnchor_(inv, INV_FOOTER_TEKS, INV_DISPLAY_START_ROW);
  if (footerRowInvFinal !== -1) {
    inv.getRange(footerRowInvFinal + 1, 3).setValue(header.sales); // C = Sales, dari TRANSAKSI kolom L
  }

  const ketRow = cariBarisAnchor_(inv, INV_KETERANGAN_ANCHOR_TEKS, INV_DISPLAY_START_ROW);
  if (ketRow !== -1) {
    inv.getRange(ketRow, 3).setValue(header.keterangan); // C = Keterangan, dari TRANSAKSI kolom P
  }

  adjustSJItemBorders();
  adjustInvoiceItemBorders();
}

  // ======================= INVOICE_CETAK.js (salinan persis) =======================
/* ======================================================
   INVOICE_CETAK.js   (PROJECT: inventory_SANFELO)
   ------------------------------------------------------
   Menggantikan menu "Export Invoice" (PDF) dengan "SIAP CETAK":
   admin klik SIAP CETAK -> dialog konfirmasi -> invoice dicatat
   sebagai TERCETAK di database (LOG_INVOICE_CETAK) lewat gateway.
   Tab LAPORAN INVOICE membaca log itu.

   Berlaku untuk sheet INVOICE dan INV_SPR (sparepart).
   Print-nya sendiri tetap dilakukan admin seperti biasa.
   ====================================================== */

// Lokasi sel tiap jenis invoice + tempat mengecek "invoice sudah tersimpan".
const LAPINV_SHEET_CFG_ = {
  // cekLokal = true kalau sheet itu BELUM dimigrasi (dicek di sheet inventory
  // sendiri). Sejak migrasi Oktober 2026, TRANSAKSI & SPAREPART sama-sama
  // ada di database -> dibaca fresh lewat gateway.
  'INVOICE': { noInv: 'H6', tgl: 'H7', customer: 'A6', cekSheet: 'TRANSAKSI', cekCol: 3, cekLokal: false },
  'INV_SPR': { noInv: 'H5', tgl: 'H8', customer: 'A5', cekSheet: 'SPAREPART', cekCol: 2, cekLokal: false }
};

function bacaKolomCekInvoice_(cfg) {
  if (!cfg.cekLokal) return bacaSheetGateway_(cfg.cekSheet, 2, cfg.cekCol, 1);
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(cfg.cekSheet);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, cfg.cekCol, sh.getLastRow() - 1, 1).getValues();
}
const LAPINV_LABEL_TOTAL_ = 'yang harus dibayar';   // dicari di kolom F (tanpa peduli huruf besar/kecil)
const LAPINV_KOLOM_NILAI_TOTAL_ = 8;                // H

function siapCetak()        { bukaDialogSiapCetak_('cetak'); }
function batalTandaiCetak() { bukaDialogSiapCetak_('batal'); }

/* ---------- baca form ---------- */

function bacaFormInvoiceCetak_(namaSheet) {
  const cfg = LAPINV_SHEET_CFG_[namaSheet];
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(namaSheet);
  if (!cfg || !sh) throw new Error('Sheet ' + namaSheet + ' tidak ditemukan.');
  SpreadsheetApp.flush();

  const noInv = String(sh.getRange(cfg.noInv).getDisplayValue()).trim();
  if (!noInv) throw new Error('No Invoice di form masih kosong.');

  const customer = String(sh.getRange(cfg.customer).getDisplayValue()).trim();
  if (!customer) throw new Error('Nama customer di form masih kosong.');

  let tgl = sh.getRange(cfg.tgl).getValue();
  if (!(tgl instanceof Date) || isNaN(tgl.getTime())) throw new Error('Tanggal invoice di form tidak valid.');
  const tglInvoice = Utilities.formatDate(tgl, Session.getScriptTimeZone(), 'yyyy-MM-dd');

  // cari baris "Yang harus Dibayar" di kolom F
  const lastRow = sh.getLastRow();
  const kolF = sh.getRange(11, 6, Math.max(lastRow - 10, 1), 1).getValues();
  let total = null;
  for (let i = 0; i < kolF.length; i++) {
    if (String(kolF[i][0]).toLowerCase().indexOf(LAPINV_LABEL_TOTAL_) !== -1) {
      total = Number(sh.getRange(11 + i, LAPINV_KOLOM_NILAI_TOTAL_).getValue());
      break;
    }
  }
  if (total === null || isNaN(total)) throw new Error('Baris "Yang harus Dibayar" tidak ketemu / bukan angka di ' + namaSheet + '.');
  if (!(total > 0)) throw new Error('Total invoice masih 0. Isi barang dulu sebelum SIAP CETAK.');

  return { sheetName: namaSheet, noInv: noInv, customer: customer, tglInvoice: tglInvoice, total: total };
}

/* ---------- dialog ---------- */

function bukaDialogSiapCetak_(mode) {
  const ui = SpreadsheetApp.getUi();
  const namaSheet = SpreadsheetApp.getActiveSheet().getName();
  const cfg = LAPINV_SHEET_CFG_[namaSheet];
  if (!cfg) {
    ui.alert('Buka dulu sheet INVOICE atau INV_SPR yang mau dicetak, lalu klik menu ini lagi.');
    return;
  }

  let form, status, tersimpan;
  try {
    form = bacaFormInvoiceCetak_(namaSheet);
    tersimpan = bacaKolomCekInvoice_(cfg)
      .some(function (r) { return String(r[0]).trim() === form.noInv; });
    status = kirimAksiGateway_('cekStatusInvoiceCetak', { noInv: form.noInv });
  } catch (err) {
    ui.alert('Tidak bisa lanjut:\n' + err.message);
    return;
  }

  if (mode === 'cetak' && !tersimpan) {
    ui.alert('Invoice ' + form.noInv + ' belum tersimpan di ' + cfg.cekSheet + '.\n' +
             'Simpan dulu (Simpan Draft / Simpan SPAREPART), baru klik SIAP CETAK.');
    return;
  }
  if (mode === 'batal' && !(status.ada && status.status === 'TERCETAK')) {
    ui.alert('Invoice ' + form.noInv + ' tidak sedang berstatus TERCETAK, jadi tidak ada yang perlu dibatalkan.');
    return;
  }

  const data = {
    mode: mode,
    sheetName: namaSheet,
    noInv: form.noInv,
    customer: form.customer,
    tglInvoice: form.tglInvoice,
    total: form.total,
    sudahAda: !!status.ada,
    statusLama: status.ada ? status.status : '',
    jumlahLama: status.ada ? status.jumlah : 0,
    totalLama: status.ada ? status.total : 0,
    tandaTerakhir: status.ada ? status.tandaTerakhir : ''
  };

  const tpl = HtmlService.createTemplateFromFile('SiapCetakDialog');
  tpl.dataJson = JSON.stringify(data).replace(/</g, '\\u003c');
  ui.showModalDialog(tpl.evaluate().setWidth(470).setHeight(520),
                     mode === 'batal' ? 'Batalkan Tanda Tercetak' : 'SIAP CETAK');
}

/**
 * Dipanggil dari dialog (google.script.run). Form dibaca ULANG di server
 * supaya yang tercatat persis isi form saat ini, bukan kiriman dari dialog.
 */
function konfirmasiSiapCetak(p) {
  const form = bacaFormInvoiceCetak_(String(p.sheetName));
  if (form.noInv !== String(p.noInv)) {
    throw new Error('Isi form berubah sejak dialog dibuka (' + p.noInv + ' -> ' + form.noInv + '). Tutup dan klik ulang.');
  }
  const alasan = String(p.alasan || '').trim();

  if (p.mode === 'batal') {
    kirimAksiGateway_('batalInvoiceTercetak', { noInv: form.noInv, alasan: alasan });
    return { pesan: '↩️ Tanda tercetak ' + form.noInv + ' dibatalkan.' };
  }

  const hasil = kirimAksiGateway_('catatInvoiceTercetak', {
    noInv: form.noInv,
    tglInvoice: form.tglInvoice,
    customer: form.customer,
    total: form.total,
    alasan: alasan
  });
  const ket = hasil.hasil === 'baru' ? 'Ditandai tercetak.' : 'Dicatat sebagai cetak ulang ke-' + hasil.jumlah + '.';
  return { pesan: '✅ ' + form.noInv + ': ' + ket + ' Silakan print sekarang.' };
}

  // ======================= SPLIT_SURAT.js (salinan persis) =======================
/* ======================================================
   SPLIT_SURAT.gs
   - aturJumlahKertas() : dialog manual buat nentuin berapa halaman (kertas)
                          yang disiapin untuk SURAT JALAN + INVOICE.

   KONSEP:
   1 halaman fisik (1/2 A4 = A5) = 8 baris item (SJ_ITEM_MAX_ROWS / INV_ITEM_MAX_ROWS
   di PO_HISTORY.gs, udah dikalibrasi sebelumnya biar pas 1 lembar nota).

   ⬅️ FIX: setting manual sekarang DISIMPAN (PropertiesService), dan
   dijadikan acuan minimum baris ("effective base rows") di fungsi-fungsi lain
   (transaksiBaru, loadSuratJalanFromHistory) supaya gak ke-collapse balik ke
   8 baris tiap kali H3 diedit atau transaksi baru dibuat.

   Google Sheets TIDAK punya API buat set page break manual programmatically.
   Jadi cara kerja fitur ini: nge-expand kapasitas baris jadi kelipatan 8
   (pakai pastikanKapasitasItem_ yang udah ada di PO_HISTORY.gs), terus pagination
   fisik antar-halamannya diserahkan ke Google Sheets/PDF export secara natural —
   yang penting break-nya SELALU per-baris utuh (gak akan motong 1 baris jadi
   kepotong 2 halaman), asal ukuran kertas export/print-nya udah bener (A5).

   Footer & tanda tangan otomatis nempel tepat setelah baris item terakhir
   (gak digandain per halaman), soalnya pastikanKapasitasItem_ emang gitu cara
   kerjanya dari awal.
   ====================================================== */

const KERTAS_ROWS_PER_HALAMAN = 8; // = SJ_ITEM_MAX_ROWS = INV_ITEM_MAX_ROWS di PO_HISTORY.gs

const MANUAL_PAGES_PROP_KEY = "JUMLAH_KERTAS_MANUAL"; // ⬅️ FIX: key penyimpanan setting manual

/* ------------------------------------------------------
   ⬅️ FIX: Helper simpan/baca/hapus setting jumlah kertas manual
   Pakai Document Properties biar persist antar-eksekusi function
   (gak ilang kayak variabel biasa), tapi tetap per-spreadsheet.
   ------------------------------------------------------ */
function getJumlahKertasManual_() {
  const props = PropertiesService.getDocumentProperties();
  const val = Number(props.getProperty(MANUAL_PAGES_PROP_KEY));
  return (val && val > 0) ? val : 1; // default 1 halaman kalau belum pernah diset
}

function setJumlahKertasManual_(jumlah) {
  const props = PropertiesService.getDocumentProperties();
  props.setProperty(MANUAL_PAGES_PROP_KEY, String(jumlah));
}

function clearJumlahKertasManual_() {
  const props = PropertiesService.getDocumentProperties();
  props.deleteProperty(MANUAL_PAGES_PROP_KEY);
}

/* ⬅️ FIX: hitung baris minimum efektif = yang lebih besar antara
   basis default (8 baris / 1 halaman) VS setting manual yang tersimpan.
   Dipakai menggantikan SJ_ITEM_MAX_ROWS/INV_ITEM_MAX_ROWS sebagai "baseRows"
   di tempat-tempat yang tadinya collapse paksa ke basis. */
function getEffectiveBaseRows_(defaultBaseRows) {
  // const manualPages = getJumlahKertasManual_();
  // const manualRows = manualPages * KERTAS_ROWS_PER_HALAMAN;
  // return Math.max(defaultBaseRows, manualRows);
  return defaultBaseRows;
}

/* ------------------------------------------------------
   ATUR JUMLAH KERTAS — dialog manual, expand SJ & INVOICE bareng
   ------------------------------------------------------ */
function aturJumlahKertas() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sj = ss.getSheetByName("SURAT JALAN");
  const inv = ss.getSheetByName("INVOICE");

  if (!sj) { ui.alert('Sheet "SURAT JALAN" tidak ditemukan.'); return; }

  const resp = ui.prompt(
    "Atur Jumlah Kertas",
    "Mau berapa lembar (halaman) buat SJ & Invoice ini? (1 lembar = " + KERTAS_ROWS_PER_HALAMAN + " baris item)",
    ui.ButtonSet.OK_CANCEL
  );

  if (resp.getSelectedButton() !== ui.Button.OK) return;

  const jumlah = Number(String(resp.getResponseText() || "").trim());
  if (!jumlah || jumlah < 1 || !Number.isInteger(jumlah)) {
    ui.alert("Jumlah kertas harus angka bulat lebih dari 0.");
    return;
  }

  setJumlahKertasManual_(jumlah); // ⬅️ FIX: simpan setting supaya gak ke-reset otomatis

  const neededRowsSJ  = jumlah * KERTAS_ROWS_PER_HALAMAN;
  const neededRowsInv = jumlah * KERTAS_ROWS_PER_HALAMAN;

  // Reuse pastikanKapasitasItem_ & konstanta dari PO_HISTORY.gs (satu project = satu scope)
  pastikanKapasitasItem_(sj, SJ_ITEM_START_ROW, SJ_ITEM_MAX_ROWS, neededRowsSJ, SJ_FOOTER_COL_START, SJ_FOOTER_COL_END, SJ_FOOTER_TEKS);
  if (inv) pastikanKapasitasItem_(inv, INV_DISPLAY_START_ROW, INV_ITEM_MAX_ROWS, neededRowsInv, INV_FOOTER_COL_START, INV_FOOTER_COL_END, INV_FOOTER_TEKS);

  adjustSJItemBorders();
  if (inv) adjustInvoiceItemBorders();

  ui.alert(
    "Kapasitas disesuaikan jadi " + jumlah + " halaman.\n" +
    "SJ: " + neededRowsSJ + " baris | Invoice: " + neededRowsInv + " baris.\n\n" +
    "Catatan: kalau item yang keisi lebih dikit dari kapasitas ini, sisa baris kosong " +
    "TETAP disediakan (gak collapse otomatis) sampai kamu isi manual/proses ulang, " +
    "karena ini kapasitas hasil setting manual, bukan hasil fetch.\n\n" +
    "Setting ini akan otomatis ke-reset ke 1 halaman kalau kamu bikin Transaksi Baru."
  );
}

/* ------------------------------------------------------
   ⬅️ FIX (opsional tapi disarankan): reset manual ke 1 halaman
   tanpa harus lewat transaksiBaru(). Berguna kalau kamu mau
   balikin ke auto-collapse tanpa bikin transaksi baru dulu.
   ------------------------------------------------------ */
function resetJumlahKertasKeAuto() {
  const ui = SpreadsheetApp.getUi();
  clearJumlahKertasManual_();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sj = ss.getSheetByName("SURAT JALAN");
  const inv = ss.getSheetByName("INVOICE");
  if (!sj) return;

  pastikanKapasitasItem_(sj, SJ_ITEM_START_ROW, SJ_ITEM_MAX_ROWS, SJ_ITEM_MAX_ROWS, SJ_FOOTER_COL_START, SJ_FOOTER_COL_END, SJ_FOOTER_TEKS);
  if (inv) pastikanKapasitasItem_(inv, INV_DISPLAY_START_ROW, INV_ITEM_MAX_ROWS, INV_ITEM_MAX_ROWS, INV_FOOTER_COL_START, INV_FOOTER_COL_END, INV_FOOTER_TEKS);

  adjustSJItemBorders();
  if (inv) adjustInvoiceItemBorders();

  ui.alert("Jumlah kertas dikembalikan ke mode otomatis (1 halaman / sesuai jumlah item).");
}

  // ======================= SyncDone.js (salinan persis) =======================
function perbaikiStatusHistoryPOMassal() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const trx  = ss.getSheetByName("TRANSAKSI");
  const det  = ss.getSheetByName("DETAIL_TRANSAKSI");
  const hist = ss.getSheetByName("HISTORY_PO_MASUK");

  if (!trx || !det || !hist) throw new Error("Sheet TRANSAKSI/DETAIL_TRANSAKSI/HISTORY_PO tidak ditemukan.");

  // 1) Ambil daftar No PO per No SJ dari TRANSAKSI (kolom B=No SJ, Q=No PO gabungan)
  const trxLastRow = trx.getLastRow();
  if (trxLastRow < 2) { ui.alert("TRANSAKSI masih kosong."); return; }

  const trxData = trx.getRange(2, 1, trxLastRow - 1, 17).getValues(); // A..Q
  const poPerSJ = {}; // noSJ -> [poNo,...]
  const skippedMultiPO = [];

  trxData.forEach(function(r) {
    const noSJ = String(r[1] || "").trim();       // B
    const poRaw = String(r[16] || "").trim();     // Q
    if (!noSJ || !poRaw) return;

    const poList = poRaw.split(",").map(function(s){ return s.trim(); }).filter(Boolean);
    poPerSJ[noSJ] = poList;
  });

  // 2) Ambil KODE+QTY per No SJ dari DETAIL_TRANSAKSI (A=No SJ, F=KODE, H=QTY)
  const detLastRow = det.getLastRow();
  if (detLastRow < 2) { ui.alert("DETAIL_TRANSAKSI masih kosong."); return; }

  const detData = det.getRange(2, 1, detLastRow - 1, 8).getValues(); // A..H
  const itemPerSJ = {}; // noSJ -> { kode: qtyTotal }

  detData.forEach(function(r) {
    const noSJ = String(r[0] || "").trim(); // A
    const kode = String(r[5] || "").trim(); // F
    const qty  = Number(r[7]) || 0;         // H
    if (!noSJ || !kode) return;

    if (!itemPerSJ[noSJ]) itemPerSJ[noSJ] = {};
    itemPerSJ[noSJ][kode] = (itemPerSJ[noSJ][kode] || 0) + qty;
  });

  // 3) Akumulasi total qty terkirim per (PO+KODE) — HANYA dari SJ yang PO-nya tunggal
  const totalDeliveredByPoKode = {}; // "poNo||kode" -> qty

  Object.keys(poPerSJ).forEach(function(noSJ) {
    const poList = poPerSJ[noSJ];
    const items = itemPerSJ[noSJ];
    if (!items) return;

    if (poList.length > 1) {
      skippedMultiPO.push(noSJ + " (PO: " + poList.join(", ") + ")");
      return; // ambigu, gak diproses otomatis
    }

    const poNo = poList[0];
    Object.keys(items).forEach(function(kode) {
      const key = poNo + "||" + kode;
      totalDeliveredByPoKode[key] = (totalDeliveredByPoKode[key] || 0) + items[kode];
    });
  });

  // 4) Kelompokkan baris HISTORY_PO (MASUK) per (PO+KODE), urut by line -- sekalian
  //    simpan info Customer/Nama Barang/status saat ini buat ditampilkan di preview.
  const histLastRow = hist.getLastRow();
  if (histLastRow < 2) { ui.alert("HISTORY_PO masih kosong."); return; }

  const histData = hist.getRange(2, 1, histLastRow - 1, 16).getValues(); // A..P
  const lineGroups = {}; // "poNo||kode" -> [{rowIdx, qtyAsli, line, customer, nama, statusSaatIni}]

  histData.forEach(function(r, i) {
    const jenis = String(r[0] || "").trim().toUpperCase();
    if (jenis !== "MASUK") return;

    const poNo = String(r[1] || "").trim();
    const customer = String(r[3] || "").trim(); // D
    const kode = String(r[8] || "").trim();     // I
    const nama = String(r[9] || "").trim();     // J
    const qtyAsli = Number(r[10]) || 0;         // K
    const statusSaatIni = String(r[11] || "").trim(); // L
    const line = r[15];                         // P

    const key = poNo + "||" + kode;
    if (!lineGroups[key]) lineGroups[key] = [];
    lineGroups[key].push({
      rowIdx: i + 2, qtyAsli: qtyAsli, line: line,
      customer: customer, nama: nama, statusSaatIni: statusSaatIni
    });
  });

  // 5) DRY RUN dulu: hitung status BARU tiap baris tanpa nulis apapun ke sheet,
  //    supaya bisa di-preview dan hanya baris yang BENERAN berubah yang dilaporkan.
  const rencana = []; // { rowIdx, poNo, kode, customer, nama, statusLama, statusBaru, warnaBaru }

  Object.keys(totalDeliveredByPoKode).forEach(function(key) {
    let sisaTerkirim = totalDeliveredByPoKode[key];
    const rows = lineGroups[key];
    if (!rows) return; // PO/kode gak ketemu di HISTORY_PO (mungkin PO lama/beda)

    const poNo = key.split("||")[0];
    const kode = key.split("||")[1];

    rows.sort(function(a,b){ return (Number(a.line)||0) - (Number(b.line)||0); });

    rows.forEach(function(entry) {
      if (sisaTerkirim <= 0) return;

      const dipakai = Math.min(sisaTerkirim, entry.qtyAsli);
      sisaTerkirim -= dipakai;
      if (dipakai <= 0) return;

      let statusBaru, warnaBaru;
      if (dipakai >= entry.qtyAsli) {
        statusBaru = "DONE SJ";
        warnaBaru = "#d9ead3"; // hijau
      } else {
        statusBaru = "DONE SJ " + dipakai;
        warnaBaru = "#fff2cc"; // kuning
      }

      // cuma masuk rencana kalau statusnya BENERAN beda dari yang sekarang
      if (entry.statusSaatIni !== statusBaru) {
        rencana.push({
          rowIdx: entry.rowIdx, poNo: poNo, kode: kode,
          customer: entry.customer, nama: entry.nama,
          statusLama: entry.statusSaatIni || "(kosong)",
          statusBaru: statusBaru, warnaBaru: warnaBaru
        });
      }
    });
  });

  // 6) Kalau gak ada yang perlu diubah, cukup laporkan itu
  if (rencana.length === 0) {
    let msg = "Semua baris HISTORY_PO_MASUK sudah sesuai dengan SJ yang tercatat di TRANSAKSI/DETAIL_TRANSAKSI. Tidak ada yang perlu diperbaiki.";
    if (skippedMultiPO.length > 0) {
      msg += "\n\n⚠️ " + skippedMultiPO.length + " SJ gabungan multi-PO tidak ikut dicek otomatis (perlu cek manual):\n" + skippedMultiPO.join("\n");
    }
    ui.alert(msg);
    return;
  }

  // 7) Tampilkan preview & minta konfirmasi SEBELUM nulis apapun ke sheet
  const JUMLAH_PREVIEW = 15;
  let preview = rencana.slice(0, JUMLAH_PREVIEW).map(function (p) {
    return "Baris " + p.rowIdx + ": PO " + p.poNo + " / " + p.customer + " / " + p.kode + " " + p.nama +
      "  ->  '" + p.statusLama + "'  menjadi  '" + p.statusBaru + "'";
  }).join("\n");
  if (rencana.length > JUMLAH_PREVIEW) {
    preview += "\n... dan " + (rencana.length - JUMLAH_PREVIEW) + " baris lainnya.";
  }

  let confirmMsg = "Ditemukan " + rencana.length + " baris HISTORY_PO_MASUK yang SJ-nya SUDAH diposting (ada di TRANSAKSI + DETAIL_TRANSAKSI) tapi status di HISTORY_PO_MASUK belum diupdate:\n\n" + preview;
  if (skippedMultiPO.length > 0) {
    confirmMsg += "\n\n⚠️ " + skippedMultiPO.length + " SJ gabungan multi-PO TIDAK ikut dicek otomatis (perlu cek manual terpisah):\n" + skippedMultiPO.slice(0, 10).join("\n");
    if (skippedMultiPO.length > 10) confirmMsg += "\n... dan " + (skippedMultiPO.length - 10) + " lainnya.";
  }
  confirmMsg += "\n\nLanjutkan update status untuk " + rencana.length + " baris ini?";

  const resp = ui.alert("⚠️ Konfirmasi Perbaikan Status HISTORY PO", confirmMsg, ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) {
    ui.alert("Dibatalkan. Tidak ada perubahan yang dilakukan.");
    return;
  }

  // 8) Baru sekarang benar-benar nulis ke sheet
  const lastCol = hist.getLastColumn();

  // ⬅️ FIX (performa): dulu tiap baris di rencana nulis STATUS (setValue)
  // & warna (setBackground) SATU-SATU DI DALAM LOOP. rencana dikelompokin
  // per (PO+KODE) makanya rowIdx-nya gak urut, jadi di-sort dulu by rowIdx
  // baru dikelompokin jadi rentang baris yang nyambung (contiguous) --
  // rencana bisa berisi ratusan baris kalau data lama belum pernah
  // disinkron, jadi tiap rentang ditulis SEKALI lewat setValues()/
  // setBackgrounds() -- hasil akhirnya identik, cuma jauh lebih sedikit
  // panggilan API.
  const rencanaUrut = rencana.slice().sort(function (a, b) { return a.rowIdx - b.rowIdx; });

  let idxP = 0;
  while (idxP < rencanaUrut.length) {
    const startRow = rencanaUrut[idxP].rowIdx;
    let endRow = startRow;
    let next = idxP + 1;
    while (next < rencanaUrut.length && rencanaUrut[next].rowIdx === endRow + 1) {
      endRow = rencanaUrut[next].rowIdx;
      next++;
    }

    const numRows = endRow - startRow + 1;
    const statusValues = [];
    const warnaGrid = [];
    for (let r = 0; r < numRows; r++) {
      const p = rencanaUrut[idxP + r];
      statusValues.push([p.statusBaru]);
      const rowWarna = [];
      for (let c = 0; c < lastCol; c++) rowWarna.push(p.warnaBaru);
      warnaGrid.push(rowWarna);
    }

    hist.getRange(startRow, 12, numRows, 1).setValues(statusValues);
    hist.getRange(startRow, 1, numRows, lastCol).setBackgrounds(warnaGrid);

    idxP = next;
  }

  // 9) Laporan hasil akhir
  let msg = "✅ " + rencana.length + " baris HISTORY_PO berhasil diupdate.";
  if (skippedMultiPO.length > 0) {
    msg += "\n\n⚠️ " + skippedMultiPO.length + " SJ gabungan multi-PO DILEWATI (perlu cek manual):\n" + skippedMultiPO.slice(0, 15).join("\n");
    if (skippedMultiPO.length > 15) msg += "\n... dan " + (skippedMultiPO.length - 15) + " lainnya.";
  }
  ui.alert(msg);
}

/* ======================================================
   perbaikiWarnaStatusHistoryPO()
   -----------------------------------------------------
   Masalah yang diperbaiki: baris HISTORY_PO_MASUK yang warna
   latar belakangnya HIJAU (#d9ead3, dipakai untuk "DONE SJ")
   atau KUNING (#fff2cc, dipakai untuk "DONE SJ <sisa>" / partial),
   TAPI teks di kolom STATUS bukan "DONE SJ..." -- misalnya masih
   "ORDER 1", "ORDER 2", atau kosong.

   Kenapa ini bisa kejadian: setiap tempat di kode yang menandai
   sebuah baris sebagai "DONE SJ" (updateStatusDoneSJHistoryPODelta_
   di PO_HISTORY.gs, dan perbaikiStatusHistoryPOMassal di atas)
   SELALU menyetel teks STATUS dan warna latar bareng-bareng dalam
   satu langkah -- jadi lewat kode, dua-duanya gak akan pernah beda.
   Tapi kalau admin mengetik langsung ke sel STATUS di spreadsheet
   (mis. mengganti "DONE SJ" jadi "ORDER 1" karena barang itu mau
   dipesan ulang), Google Sheets HANYA mengganti teksnya saja --
   warna latar lama yang sudah ada di sel itu TIDAK ikut kehapus.
   Hasilnya baris itu jadi "keliatan" hijau/done padahal teksnya
   sudah bilang lain.

   Perbaikan ini CUMA mengembalikan warna latar baris yang salah
   itu ke putih -- teks STATUS ("ORDER 1", dst) SAMA SEKALI TIDAK
   diubah/dihapus. Selalu tampil daftar & minta konfirmasi dulu
   sebelum baris manapun diubah.
   ====================================================== */
function perbaikiWarnaStatusHistoryPO() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hist = ss.getSheetByName("HISTORY_PO_MASUK");
  const ui = SpreadsheetApp.getUi();
  if (!hist) { ui.alert("Sheet HISTORY_PO_MASUK tidak ditemukan."); return; }

  const lastRow = hist.getLastRow();
  if (lastRow < 2) { ui.alert("HISTORY_PO_MASUK masih kosong."); return; }
  const lastCol = hist.getLastColumn();

  const HIJAU = "#d9ead3";
  const KUNING = "#fff2cc";

  // baca teks + warna latar kolom STATUS (L = kolom 12) untuk semua baris data
  const statusRange = hist.getRange(2, 12, lastRow - 1, 1);
  const statusValues = statusRange.getValues();
  const statusColors = statusRange.getBackgrounds();

  function isTeksDoneSJ(s) {
    return String(s || "").trim().toUpperCase().indexOf("DONE SJ") === 0;
  }
  function normalisasiWarna(c) {
    return String(c || "").toLowerCase();
  }

  const bermasalah = [];
  for (let i = 0; i < statusValues.length; i++) {
    const teks = statusValues[i][0];
    const warna = normalisasiWarna(statusColors[i][0]);
    const isHijauAtauKuning = (warna === HIJAU || warna === KUNING);
    if (isHijauAtauKuning && !isTeksDoneSJ(teks)) {
      bermasalah.push({
        rowIdx: i + 2,
        teks: teks,
        warnaLabel: (warna === HIJAU ? "HIJAU" : "KUNING")
      });
    }
  }

  if (bermasalah.length === 0) {
    ui.alert(
      "✅ Tidak Ada Masalah",
      "Semua baris HISTORY_PO_MASUK sudah konsisten: baris yang berwarna hijau/kuning teksnya selalu 'DONE SJ...', dan sebaliknya. Tidak ada yang perlu diperbaiki.",
      ui.ButtonSet.OK
    );
    return;
  }

  // ambil detail PO/Customer/Kode buat ditampilkan di preview konfirmasi
  const infoRange = hist.getRange(2, 1, lastRow - 1, 16); // A..P
  const infoValues = infoRange.getValues();

  const JUMLAH_PREVIEW = 15;
  let preview = bermasalah.slice(0, JUMLAH_PREVIEW).map(function (b) {
    const info = infoValues[b.rowIdx - 2];
    const poNo = info[1];       // B = NO PO
    const customer = info[3];   // D = CUSTOMER
    const kode = info[8];       // I = KODE_BARANG
    const nama = info[9];       // J = NAMA_BARANG
    return "Baris " + b.rowIdx + ": PO " + poNo + " / " + customer + " / " + kode + " " + nama +
      " -> STATUS='" + b.teks + "' tapi warna latar " + b.warnaLabel;
  }).join("\n");
  if (bermasalah.length > JUMLAH_PREVIEW) {
    preview += "\n... dan " + (bermasalah.length - JUMLAH_PREVIEW) + " baris lainnya.";
  }

  const resp = ui.alert(
    "⚠️ Ditemukan " + bermasalah.length + " Baris Bermasalah",
    "Baris di bawah ini warna latarnya HIJAU/KUNING (seolah-olah sudah 'DONE SJ') tapi teks STATUS-nya BUKAN 'DONE SJ...'. " +
    "Ini kemungkinan besar sisa warna lama yang ketinggalan waktu teks STATUS-nya diedit manual di spreadsheet.\n\n" +
    "Perbaikan ini HANYA mengembalikan warna latar baris-baris ini ke putih. " +
    "Teks STATUS (mis. 'ORDER 1') TIDAK akan diubah atau dihapus sama sekali, dan tidak ada data lain yang disentuh.\n\n" +
    preview +
    "\n\nLanjutkan perbaikan warna untuk " + bermasalah.length + " baris ini?",
    ui.ButtonSet.YES_NO
  );

  if (resp !== ui.Button.YES) {
    ui.alert("Dibatalkan. Tidak ada perubahan yang dilakukan.");
    return;
  }

  // ⬅️ FIX (performa): dulu setBackground() dipanggil SATU-SATU per baris
  // bermasalah DI DALAM LOOP. "bermasalah" hasil scan forward dari atas ke
  // bawah jadi rowIdx-nya udah pasti urut naik, tinggal dikelompokin jadi
  // rentang baris yang nyambung (contiguous), tiap rentang ditulis SEKALI
  // lewat setBackgrounds() -- hasil akhirnya identik, cuma jauh lebih
  // sedikit panggilan API.
  let idxB = 0;
  while (idxB < bermasalah.length) {
    const startRow = bermasalah[idxB].rowIdx;
    let endRow = startRow;
    let next = idxB + 1;
    while (next < bermasalah.length && bermasalah[next].rowIdx === endRow + 1) {
      endRow = bermasalah[next].rowIdx;
      next++;
    }

    const numRows = endRow - startRow + 1;
    const warnaGrid = [];
    for (let r = 0; r < numRows; r++) {
      const rowWarna = [];
      for (let c = 0; c < lastCol; c++) rowWarna.push("#ffffff");
      warnaGrid.push(rowWarna);
    }
    hist.getRange(startRow, 1, numRows, lastCol).setBackgrounds(warnaGrid);

    idxB = next;
  }

  ui.alert(
    "✅ Selesai",
    bermasalah.length + " baris berhasil diperbaiki warnanya (dikembalikan ke putih). Teks STATUS di semua baris tersebut tidak diubah sama sekali.",
    ui.ButtonSet.OK
  );
}

  // ======================= PO_KELUAR.js (salinan persis) =======================
const PO_KELUAR_ITEM_START_ROW = 7;

/* ---------- Helper: key-value SYSTEM (biar gampang tambah key baru) ---------- */
function getSystemRow_(system, key) {
  const lastRow = Math.max(system.getLastRow(), 1);
  const data = system.getRange(1, 1, lastRow, 1).getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim().toUpperCase() === key) return i + 1;
  }
  const newRow = system.getLastRow() + 1;
  system.getRange(newRow, 1).setValue(key);
  return newRow;
}
function getSystemValue_(key) {
  const system = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SYSTEM");
  return system.getRange(getSystemRow_(system, key), 2).getValue();
}
function setSystemValue_(key, value) {
  const system = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SYSTEM");
  system.getRange(getSystemRow_(system, key), 2).setValue(value);
}

/* ---------- Helper: BARU - parse tanggal dari berbagai format ---------- */
// Data yang balik dari gateway lewat JSON TIDAK punya tipe Date (semua jadi
// string ISO). Helper ini menormalkan Date / string ISO / "yyyy-MM-dd" /
// "dd/MM/yyyy" menjadi objek Date, atau null kalau tidak valid.
function parseTanggal_(v) {
  if (v instanceof Date) return isNaN(v) ? null : v;
  if (v === null || v === undefined || v === "") return null;
  const s = String(v).trim();
  if (!s) return null;

  // dd/MM/yyyy
  let m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));

  // yyyy-MM-dd (tanpa jam) -> jam 12 siang biar aman dari pergeseran timezone
  m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), 12, 0, 0);

  // string ISO lengkap, contoh "2026-09-30T00:00:00.000Z" / "+07:00"
  const d = new Date(s);
  return isNaN(d) ? null : d;
}

/* ---------- Helper: fetch nilai dari sheet DATA_BERSIH (kolom B -> kolom G) ---------- */
// Bikin map sekali (namaBarang di-uppercase+trim sebagai key) biar gak berkali-kali baca sheet dalam loop.
function buatMapDataBersih_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DATA_BERSIH");
  if (!sheet) {
    throw new Error('Sheet "DATA_BERSIH" tidak ditemukan. Cek lagi nama sheet-nya persis sama (huruf besar/kecil & spasi/underscore ikut diperhatikan).');
  }
  const lastRow = Math.max(sheet.getLastRow(), 1);
  const data = sheet.getRange(1, 2, lastRow, 6).getValues(); // kolom B..G (B=nama, G=nilai)

  const map = {};
  data.forEach(r => {
    const nama = r[0];   // kolom B
    const nilai = r[5];  // kolom G
    if (nama) map[String(nama).trim().toUpperCase()] = nilai;
  });
  return map;
}

function cariNilaiDataBersih_(map, namaBarang) {
  if (!namaBarang) return "";
  const key = String(namaBarang).trim().toUpperCase();
  return map.hasOwnProperty(key) ? map[key] : "";
}

/* ---------- Helper: cari baris "TOTAL QTY" di PO_KELUAR ---------- */
function cariBarisTotalQty_(sheet) {
  const lastRow = Math.max(sheet.getLastRow(), PO_KELUAR_ITEM_START_ROW);

  // Ambil kolom C saja
  const values = sheet.getRange(
    PO_KELUAR_ITEM_START_ROW,
    3,
    lastRow - PO_KELUAR_ITEM_START_ROW + 1,
    1
  ).getValues();

  const totalRows = [];

  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim().toUpperCase() === "TOTAL QTY") {
      totalRows.push(PO_KELUAR_ITEM_START_ROW + i);
    }
  }

  // Kalau belum ada TOTAL QTY
  if (totalRows.length === 0) {
    let row = PO_KELUAR_ITEM_START_ROW;

    // cari baris kosong pertama setelah data
    while (sheet.getRange(row, 3).getValue() !== "") {
      row++;
    }

    sheet.getRange(row, 3).setValue("TOTAL QTY");
    sheet.getRange(row, 4).setValue(0);

    return row;
  }

  // Kalau ada lebih dari satu TOTAL QTY, sisakan yang pertama
  const totalRow = totalRows[0];

  for (let i = totalRows.length - 1; i >= 1; i--) {
    sheet.deleteRow(totalRows[i]);
  }

  return totalRow;
}

function hitungUlangTotalQty_(sheet) {
  const totalRow = cariBarisTotalQty_(sheet);
  let total = 0;
  if (totalRow > PO_KELUAR_ITEM_START_ROW) {
    const jumlahs = sheet.getRange(PO_KELUAR_ITEM_START_ROW, 4, totalRow - PO_KELUAR_ITEM_START_ROW, 1).getValues();
    jumlahs.forEach(r => { total += Number(r[0]) || 0; });
  }
  sheet.getRange(totalRow, 4).setValue(total);
}

/* ---------- 1) Tambah item dari PO_MASUK ke PO_KELUAR ---------- */
// CATATAN: PO_ITEM_START_ROW & PO_ITEM_MAX_ROWS harus didefinisikan di file PO_MASUK kamu.
function tambahKePOKeluar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masuk  = ss.getSheetByName("PO_MASUK");
  const keluar = ss.getSheetByName("PO_KELUAR");

  const namaToko = masuk.getRange("B2").getValue();
  if (!namaToko) {
    SpreadsheetApp.getUi().alert("PO_MASUK belum ada customer-nya.");
    return;
  }

  const dataMasuk = masuk.getRange(PO_ITEM_START_ROW, 3, PO_ITEM_MAX_ROWS, 2).getValues(); // C: barang, D: qty
  const itemBaru = [];
  dataMasuk.forEach(r => {
    const [namaBarang, jumlah] = r;
    if (namaBarang && jumlah) itemBaru.push({ namaBarang, jumlah });
  });

  if (itemBaru.length === 0) {
    SpreadsheetApp.getUi().alert("Tidak ada barang di PO_MASUK yang bisa diambil.");
    return;
  }

  // Map DATA_BERSIH (namaBarang -> nilai kolom G), dibuat sekali di awal
  const mapDataBersih = buatMapDataBersih_();

  // ---- baca item lama, rekonstruksi grup per toko ----
  const totalRowLama = cariBarisTotalQty_(keluar);
  const groups = []; // [{ namaToko, items: [{namaBarang, jumlah, keterangan, nilaiF}] }]

  if (totalRowLama > PO_KELUAR_ITEM_START_ROW) {
    const raw = isiTokoKosong_(
      keluar.getRange(PO_KELUAR_ITEM_START_ROW, 2, totalRowLama - PO_KELUAR_ITEM_START_ROW, 5).getValues()
    );
    raw.forEach(r => {
      const [toko, namaBarang, jumlah, keterangan, nilaiF] = r;
      if (!namaBarang) return;
      let group = groups.find(g => g.namaToko === toko);
      if (!group) { group = { namaToko: toko, items: [] }; groups.push(group); }
      group.items.push({ namaBarang, jumlah, keterangan: keterangan || "", nilaiF: nilaiF || "" });
    });
  }

  // ---- merge item baru ke grup toko yang sesuai ----
  let group = groups.find(g => g.namaToko === namaToko);
  if (!group) { group = { namaToko: namaToko, items: [] }; groups.push(group); }

  itemBaru.forEach(baru => {
    const nilaiF = cariNilaiDataBersih_(mapDataBersih, baru.namaBarang);

    group.items.push({
      namaBarang: baru.namaBarang,
      jumlah: baru.jumlah,
      keterangan: "",
      nilaiF: nilaiF
    });
  });

  // ---- flatten: toko cuma ditulis di baris pertama tiap grup ----
  const rows = [];
  let no = 1;

  groups.forEach(g => {
    g.items.forEach((it, i) => {
      rows.push([
        i === 0 ? no : "",
        i === 0 ? g.namaToko : "",
        it.namaBarang,
        it.jumlah,
        it.keterangan,
        it.nilaiF
      ]);
    });

    no++;
  });

  if (totalRowLama > PO_KELUAR_ITEM_START_ROW) {
    keluar.getRange(PO_KELUAR_ITEM_START_ROW, 1, totalRowLama - PO_KELUAR_ITEM_START_ROW, 6).clearContent();
  }

  const totalRowBaru = PO_KELUAR_ITEM_START_ROW + rows.length;
  const selisih = totalRowBaru - totalRowLama;
  if (selisih > 0) keluar.insertRowsBefore(totalRowLama, selisih);
  else if (selisih < 0) keluar.deleteRows(totalRowBaru, -selisih);

  keluar.getRange(PO_KELUAR_ITEM_START_ROW, 1, rows.length, 6).setValues(rows);

  hitungUlangTotalQty_(keluar);
  adjustPOKeluarBorders_(keluar);

  SpreadsheetApp.getUi().alert("Berhasil tarik/update " + itemBaru.length + " barang (" + namaToko + ") ke PO_KELUAR.");
}

/* ---------- 2) Simpan snapshot ke HISTORY_PO_KELUAR (lewat gateway) ---------- */
// FIX: tanggal dikirim sebagai string "yyyy-MM-dd" yang pasti, bukan Date
// mentah yang di-JSON.stringify jadi ISO UTC (rawan bergeser hari / jadi teks).
// Di aksiSyncPoKeluar_ (DatabaseGateway.js), string ini HARUS diubah jadi Date
// asli sebelum ditulis ke sheet (lihat catatan di bawah file ini).
function simpanHistoryPOKeluar_(ss, poNo, tgl, items) {
  const t = parseTanggal_(tgl);
  const tglStr = t ? Utilities.formatDate(t, "Asia/Jakarta", "yyyy-MM-dd") : "";
  kirimAksiGateway_("syncPoKeluar", { poNo: poNo, tgl: tglStr, items: items });

  // ⬅️ BARU: simpan PO keluar bisa menggeser urutan baris di database,
  // jadi QTY TERKIRIM (kolom H lokal, nempel per nomor baris) dihitung
  // ulang supaya ikut lurus lagi. Lihat QTY_TERKIRIM_ULANG.js.
  try {
    hitungUlangQtyTerkirimPOKeluar_(false);
  } catch (err) {
    Logger.log("Hitung ulang QTY TERKIRIM (simpan PO keluar) error: " + err);
  }
}

/* ---------- 3) Cari/generate No PO berdasarkan tanggal ---------- */
// Baca FRESH dari database lewat gateway. Hasil gateway berupa JSON, jadi
// kolom TGL datang sebagai string -> WAJIB lewat parseTanggal_.
function getNoPOByTanggal_(tanggal) {
  const tglInput = parseTanggal_(tanggal);
  if (!tglInput) throw new Error("Tanggal PO tidak valid: " + tanggal);

  const target = Utilities.formatDate(tglInput, "Asia/Jakarta", "yyyy-MM-dd");
  const tahunTarget = Utilities.formatDate(tglInput, "Asia/Jakarta", "yyyy");

  const data = bacaSheetGateway_("HISTORY_PO_KELUAR", 2, 1, 2); // kolom A NO PO, B TGL

  if (data.length === 0) {
    return "#" + tahunTarget + "-001";
  }

  let lastNo = 0;

  for (let i = 0; i < data.length; i++) {
    const no = String(data[i][0]).trim();
    if (!no) continue;

    const tgl = parseTanggal_(data[i][1]); // FIX: bukan lagi data[i][1] mentah

    if (tgl) {
      const tglHist = Utilities.formatDate(tgl, "Asia/Jakarta", "yyyy-MM-dd");
      if (tglHist === target) {
        return no; // tanggal sama persis -> reuse nomor yang sama
      }
    }

    // format baru: #YYYY-NNN -> ambil tahun dari nomornya sendiri
    const match = no.match(/^#(\d{4})-(\d{3})$/);
    if (match) {
      if (match[1] === tahunTarget) {
        const angka = Number(match[2]);
        if (angka > lastNo) lastNo = angka;
      }
    } else if (tgl) {
      // fallback data lama format "#NNN" (belum ada tahun) -> cek tahun dari kolom TGL
      const tahunHist = Utilities.formatDate(tgl, "Asia/Jakarta", "yyyy");
      if (tahunHist === tahunTarget) {
        const angkaLama = Number(no.replace("#", ""));
        if (angkaLama > lastNo) lastNo = angkaLama;
      }
    }
  }

  return "#" + tahunTarget + "-" + Utilities.formatString("%03d", lastNo + 1);
}

/* ---------- 4) Generate No PO otomatis waktu E4 (TGL) diisi manual ---------- */
function generateNoPOKeluarDariTanggal_(keluar) {
  const tgl = keluar.getRange("E4").getValue();

  if (!tgl) {
    keluar.getRange("B4").clearContent();
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const nextNo = getNoPOByTanggal_(tgl);
  keluar.getRange("B4").setValue(nextNo);

  // kalau nomor ini sudah ada datanya di history, fetch balik item-itemnya
  const items = ambilItemHistoryPOKeluar_(ss, nextNo);
  if (items.length > 0) {
    isiPOKeluarDariHistory_(keluar, items);
  }
}

/* ---------- 5) PO Keluar Baru: simpan history dulu, lalu clear ---------- */
function poKeluarBaru() {
  const lock = LockService.getScriptLock();
  const dapatLock = lock.tryLock(5000); // tunggu maks 5 detik

  if (!dapatLock) {
    SpreadsheetApp.getUi().alert("Masih ada proses PO Keluar Baru yang berjalan, coba lagi sebentar.");
    return;
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const keluar = ss.getSheetByName("PO_KELUAR");

    // Simpan PO lama ke history (kalau ada isinya)
    const noLama  = keluar.getRange("B4").getValue();
    const tglLama = keluar.getRange("E4").getValue();
    const totalRow = cariBarisTotalQty_(keluar);

    if (noLama && totalRow > PO_KELUAR_ITEM_START_ROW) {
      const raw = keluar
        .getRange(
          PO_KELUAR_ITEM_START_ROW,
          2,
          totalRow - PO_KELUAR_ITEM_START_ROW,
          5
        )
        .getValues();

      const items = isiTokoKosong_(raw);
      simpanHistoryPOKeluar_(ss, noLama, tglLama, items);
    }

    // Bersihkan isi PO (kolom A:F)
    if (totalRow >= PO_KELUAR_ITEM_START_ROW) {
      keluar.getRange(
          PO_KELUAR_ITEM_START_ROW,
          1,
          totalRow - PO_KELUAR_ITEM_START_ROW + 1,
          6
      ).clearContent();
    }

    // Kosongkan No PO & Tanggal
    keluar.getRange("B4").clearContent();
    keluar.getRange("E4").clearContent();

    // Buat ulang TOTAL QTY
    keluar.getRange(PO_KELUAR_ITEM_START_ROW, 3).setValue("TOTAL QTY");
    keluar.getRange(PO_KELUAR_ITEM_START_ROW, 4).setValue(0);

    adjustPOKeluarBorders_(keluar);

    SpreadsheetApp.getUi().alert(
      "PO Keluar sudah di-clear.\nSilakan isi TGL di E4, No PO akan otomatis muncul di B4."
    );

  } finally {
    lock.releaseLock();
  }
}

function adjustPOKeluarBorders_(sheet) {

  const totalRow = cariBarisTotalQty_(sheet);
  const lastRow = Math.max(sheet.getMaxRows(), totalRow);

  // Hapus semua border area tabel (A:F)
  sheet.getRange(PO_KELUAR_ITEM_START_ROW, 1, lastRow - PO_KELUAR_ITEM_START_ROW + 1, 6)
       .setBorder(false, false, false, false, false, false);

  // Border untuk setiap baris yang ada data
  for (let r = PO_KELUAR_ITEM_START_ROW; r < totalRow; r++) {

    const namaToko = sheet.getRange(r, 2).getValue();
    const namaBarang = sheet.getRange(r, 3).getValue();

    if (namaToko || namaBarang) {
      sheet.getRange(r, 1, 1, 6)
           .setBorder(true, true, true, true, true, true);
    }
  }

  // Border TOTAL QTY
  sheet.getRange(totalRow, 1, 1, 6)
       .setBorder(true, true, true, true, true, true);
}

// isi ulang nama toko yang kosong (baris grup ke-2,3,dst) — dipakai pas mau nulis ke history
// rows 5 kolom: [toko, namaBarang, jumlah, keterangan, nilaiF]
function isiTokoKosong_(rows) {
  let tokoBerjalan = "";
  return rows.map(r => {
    const [toko, namaBarang, jumlah, keterangan, nilaiF] = r;
    if (toko) tokoBerjalan = toko;
    return [tokoBerjalan, namaBarang, jumlah, keterangan, nilaiF];
  });
}

function exportPOKeluarPDF() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pok = ss.getSheetByName("PO_KELUAR");

  const lastRow = pok.getLastRow(); // otomatis ngikut sepanjang data + baris TOTAL QTY
  const lastCol = 6;                // A:F

  const sheetId = pok.getSheetId();
  const spreadsheetId = ss.getId();

  const r1 = 0;
  const c1 = 0;
  const r2 = lastRow;
  const c2 = lastCol;

  const url = "https://docs.google.com/spreadsheets/d/" + spreadsheetId +
    "/export?format=pdf" +
    "&gid=" + sheetId +
    "&size=A4" +
    "&portrait=true" +
    "&fitw=true" +
    "&gridlines=false" +
    "&printtitle=false" +
    "&sheetnames=false" +
    "&top_margin=0.3&bottom_margin=0.3&left_margin=0.3&right_margin=0.3" +
    "&horizontal_alignment=CENTER" +
    "&r1=" + r1 + "&c1=" + c1 + "&r2=" + r2 + "&c2=" + c2;

  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(url, {
    headers: { Authorization: "Bearer " + token }
  });

  // Nama file diambil dari cell B4 (berisi No PO, misal "#2026-063")
  const rawPoLabel = String(pok.getRange("B4").getValue() || "PO_KELUAR").trim();
  let fileName = "PO KELUAR - " + rawPoLabel;
  fileName = fileName.replace(/[\\\/:*?"<>|]/g, "").trim() + ".pdf";

  const blob = response.getBlob().setName(fileName);

  const POK_FOLDER_ID = "1b9MU5pJU6QGc0hVeW-p5KOwyy7mROVMq";
  const folder = DriveApp.getFolderById(POK_FOLDER_ID);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  const downloadUrl = "https://drive.google.com/uc?export=download&id=" + file.getId();

  const html = HtmlService.createHtmlOutput(
    '<div style="font-family:sans-serif;padding:10px;">' +
    '<p>PDF berhasil dibuat:</p>' +
    '<p><a href="' + downloadUrl + '" target="_blank">⬇️ Download ' + fileName + '</a></p>' +
    '</div>'
  ).setWidth(350).setHeight(120);

  SpreadsheetApp.getUi().showModalDialog(html, "Export PO Keluar");
}

function syncData(){
  syncNilaiDataBersihKeHistory_();
}

/* ---------- JALANKAN SEKALI: sync kolom I (NILAI DATA BERSIH) untuk data history yang sudah ada ---------- */
function syncNilaiDataBersihKeHistory_() {
  const mapDataBersih = buatMapDataBersih_();

  const mapUntukGateway = {};
  Object.keys(mapDataBersih).forEach(function (key) {
    mapUntukGateway[key] = mapDataBersih[key];
  });

  const hasil = kirimAksiGateway_("syncNilaiDataBersihBatch", { map: mapUntukGateway });

  SpreadsheetApp.getUi().alert("Selesai sync " + hasil.baris_diubah + " baris ke kolom I (NILAI DATA BERSIH) di database_SANFELO. Mirror lokal akan menyusul beberapa saat lagi.");
}

/* ---------- 6) Ambil item dari HISTORY_PO_KELUAR berdasarkan No PO (fresh dari database) ---------- */
// Catatan: kolom NAMA TOKO di database kosong untuk baris non-pertama grup
// (lihat isiTokoKosong_), jadi toko di-carry ke baris-baris berikutnya
// supaya isiPOKeluarDariHistory_ tidak memecah satu toko jadi beberapa grup.
function ambilItemHistoryPOKeluar_(ss, poNo) {
  const data = bacaSheetGateway_("HISTORY_PO_KELUAR", 2, 1, 9); // A..I
  const items = [];
  const target = String(poNo).trim();
  let tokoBerjalan = "";
  let poBerjalan = "";

  data.forEach(r => {
    const [no, , toko, barang, jumlah, keterangan, , , nilaiF] = r;
    const noStr = String(no).trim();

    if (noStr !== poBerjalan) { poBerjalan = noStr; tokoBerjalan = ""; }
    if (toko) tokoBerjalan = toko;

    if (noStr === target && barang) {
      items.push({
        toko: tokoBerjalan,
        barang,
        jumlah,
        keterangan: keterangan || "",
        nilaiF: nilaiF || ""
      });
    }
  });
  return items;
}

/* ---------- 7) Tulis ulang item history ke sheet PO_KELUAR ---------- */
function isiPOKeluarDariHistory_(keluar, items) {
  // Rekonstruksi grup per toko (toko cuma ditulis di baris pertama tiap grup)
  const groups = [];
  items.forEach(it => {
    let group = groups.find(g => g.namaToko === it.toko);
    if (!group) { group = { namaToko: it.toko, items: [] }; groups.push(group); }
    group.items.push({ namaBarang: it.barang, jumlah: it.jumlah, keterangan: it.keterangan, nilaiF: it.nilaiF });
  });

  const rows = [];
  let no = 1;
  groups.forEach(g => {
    g.items.forEach((it, i) => {
      rows.push([
        i === 0 ? no : "",
        i === 0 ? g.namaToko : "",
        it.namaBarang,
        it.jumlah,
        it.keterangan,
        it.nilaiF
      ]);
    });
    no++;
  });

  const totalRowLama = cariBarisTotalQty_(keluar);
  if (totalRowLama > PO_KELUAR_ITEM_START_ROW) {
    keluar.getRange(PO_KELUAR_ITEM_START_ROW, 1, totalRowLama - PO_KELUAR_ITEM_START_ROW, 6).clearContent();
  }

  const totalRowBaru = PO_KELUAR_ITEM_START_ROW + rows.length;
  const selisih = totalRowBaru - totalRowLama;
  if (selisih > 0) keluar.insertRowsBefore(totalRowLama, selisih);
  else if (selisih < 0) keluar.deleteRows(totalRowBaru, -selisih);

  keluar.getRange(PO_KELUAR_ITEM_START_ROW, 1, rows.length, 6).setValues(rows);

  hitungUlangTotalQty_(keluar);
  adjustPOKeluarBorders_(keluar);
}

/* ------------------------------------------------------
   8) Forward edit kolom H (QTY TERKIRIM) ke database_SANFELO
   ------------------------------------------------------ */
function handleHistoryPOKeluarQtyEdit_(row, qtyTerkirimBaru) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hist = ss.getSheetByName("HISTORY_PO_KELUAR");
  if (!hist) return;

  const poNo = String(hist.getRange(row, 1).getValue() || "").trim(); // A
  const namaBarang = String(hist.getRange(row, 4).getValue() || "").trim(); // D

  if (!poNo || !namaBarang) return; // baris kosong / belum ada datanya, abaikan

  // Cari NAMA TOKO: scan ke atas sampai ketemu kolom C yang keisi, jangan lewatin PO lain.
  let namaToko = "";
  for (let r = row; r >= 2; r--) {
    const rowPoNo = String(hist.getRange(r, 1).getValue() || "").trim();
    if (rowPoNo !== poNo) break;
    const toko = String(hist.getRange(r, 3).getValue() || "").trim();
    if (toko) { namaToko = toko; break; }
  }

  try {
    kirimAksiGateway_("updateQtyTerkirimPoKeluar", {
      poNo: poNo,
      namaToko: namaToko,
      namaBarang: namaBarang,
      qtyTerkirim: qtyTerkirimBaru
    });
  } catch (err) {
    Logger.log("Gagal forward QTY TERKIRIM ke database: " + err);
    SpreadsheetApp.getUi().alert("Gagal menyimpan QTY TERKIRIM ke database_SANFELO: " + err.message + "\nCoba edit ulang sel ini.");
  }
}

/* ======================================================
   CATATAN: WAJIB diubah di DatabaseGateway.js (sisi database_SANFELO)
   ------------------------------------------------------
   Di aksiSyncPoKeluar_, `tgl` sekarang datang sebagai "yyyy-MM-dd".
   Ubah jadi Date asli SEBELUM ditulis ke kolom TGL:

     const p = String(tgl).split("-").map(Number);
     const tglDate = new Date(p[0], p[1] - 1, p[2], 12, 0, 0);

   lalu tulis `tglDate` (bukan string `tgl`) ke sheet. Pastikan timezone
   project database = Asia/Jakarta (Project Settings).
   ====================================================== */

  // status ringkasan per toko (kolom G:H TEMPO TOKO di inventory) dihitung persis seperti
  // updateStatusPembelianAll_, TANPA menulis apa pun
  function statusRingkasanToko_() {
    const m = ringkasanTokoMentah_();
    return m.customerOrder.map(function (c) { return { customer: c, status: statusFinalCustomer_(m.customerInfo[c]) }; });
  }

  // bahan writeRingkasanToko_ (urutan customer + info locked/grace) + invoice LOCKED per customer
  function ringkasanTokoMentah_() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dataTempo = bacaSheetGateway_("TEMPO TOKO", 2, 1, 5);
    const statusMap = buildPenjualanStatusMap_(ss.getSheetByName("PENJUALAN"));
    const todayStr = Utilities.formatDate(new Date(), "Asia/Jakarta", "yyyy-MM-dd");
    const penjualanTerbaca = Object.keys(statusMap).length > 0;
    const customerOrder = [];
    const customerInfo = {};
    const lockedPerToko = {};
    for (let i = 0; i < dataTempo.length; i++) {
      const r = dataTempo[i];
      const customer = String(r[TEMPO_COL_CUSTOMER - 1] || "").trim();
      const noInv = String(r[TEMPO_COL_NOINV - 1] || "").trim();
      const adaDiPenjualan = statusMap.hasOwnProperty(noInv);
      const pelunasan = adaDiPenjualan ? statusMap[noInv] : (penjualanTerbaca ? "LUNAS" : "");
      const hasil = hitungStatusTempo_(r[TEMPO_COL_TEMPO - 1], pelunasan, todayStr);
      if (!customer) continue;
      if (!(customer in customerInfo)) {
        customerOrder.push(customer);
        customerInfo[customer] = { locked: false, hasGrace: false, minSisaHari: TEMPO_GRACE_DAYS };
      }
      if (hasil.status === "LOCKED") {
        customerInfo[customer].locked = true;
        (lockedPerToko[customer] = lockedPerToko[customer] || []).push({ noInv: noInv, tempo: r[TEMPO_COL_TEMPO - 1] });
      } else if (hasil.daysOver > 0) {
        customerInfo[customer].hasGrace = true;
        const sisaHari = TEMPO_GRACE_DAYS - hasil.daysOver;
        if (sisaHari < customerInfo[customer].minSisaHari) customerInfo[customer].minSisaHari = sisaHari;
      }
    }
    return { customerOrder: customerOrder, customerInfo: customerInfo, lockedPerToko: lockedPerToko };
  }

  WEB_INV_CACHE_ = {
    statusRingkasanToko_: statusRingkasanToko_,
    ringkasanTokoMentah_: ringkasanTokoMentah_,
    TEMPO_PASSWORD: TEMPO_PASSWORD,
    hitungTempoHariUntukTransaksi_: hitungTempoHariUntukTransaksi_,
    hitungTanggalTempoTransaksi_: hitungTanggalTempoTransaksi_,
    buildHistoryPOStatusMap_: buildHistoryPOStatusMap_,
    hitungPersenPotonganSetPO_: hitungPersenPotonganSetPO_,
    getExtraData: getExtraData,
    getTimestamp: getTimestamp,
    hanyaTanggal_: hanyaTanggal_,
    ambilMappingLamaDariDetailTransaksi_: ambilMappingLamaDariDetailTransaksi_,
    bacaDetailDB_: bacaDetailDB_,
    bacaTransaksiDB_: bacaTransaksiDB_,
    updateStatusDoneSJHistoryPODelta_: updateStatusDoneSJHistoryPODelta_,
    syncPenjualanFromTransaksi_: syncPenjualanFromTransaksi_,
    syncTempoToko_: syncTempoToko_,
    parseDoneQty_: parseDoneQty_,
    persenDiskonSetTandaTerima_: persenDiskonSetTandaTerima_,
    buildHppMapDariDataBersih_: buildHppMapDariDataBersih_,
    parsePoListTT_: parsePoListTT_,
    bacaHistoryPOKeluarData_: bacaHistoryPOKeluarData_,
    cariPoItemDiHistory_: cariPoItemDiHistory_,
    cariBarisPembelianByNoSJ_: cariBarisPembelianByNoSJ_,
    hapusBarisPembelian_: hapusBarisPembelian_,
    hitungUlangQtyTerkirimPOKeluar_: hitungUlangQtyTerkirimPOKeluar_,
    syncHutangFromPembelian_: syncHutangFromPembelian_,
    simpanBiayaTambahan: simpanBiayaTambahan,
    getDaftarInvoicePalet: getDaftarInvoicePalet,
    bacaEkspedisiDb_: bacaEkspedisiDb_,
    jenisBarisEkspedisi_: jenisBarisEkspedisi_,
    hitungTempo_: hitungTempo_,
    buildPembelianMap_: buildPembelianMap_,
    syncPenjualanFromSparepart_: syncPenjualanFromSparepart_,
    ambilDataStokBarang: ambilDataStokBarang,
    processTransactionBatch: processTransactionBatch,
    cetakUlangSJDariNoSJ_: cetakUlangSJDariNoSJ_,
    terapkanPilihanSJSparepartDraft: terapkanPilihanSJSparepartDraft,
    isiFormTandaTerimaDariExisting_: isiFormTandaTerimaDariExisting_,
    hitungUlangHargaTandaTerima_: hitungUlangHargaTandaTerima_,
    isiSheetInvoicePalet_: isiSheetInvoicePalet_,
    cariBarisAnchor_: cariBarisAnchor_,
    konfirmasiSiapCetak: konfirmasiSiapCetak,
    perbaikiStatusHistoryPOMassal: perbaikiStatusHistoryPOMassal,
    perbaikiTempoTransaksiLama: perbaikiTempoTransaksiLama,
    getSJBelumInvoice: getSJBelumInvoice,
    simpanGroupingInvoice: simpanGroupingInvoice,
    getNoPOByTanggal_: getNoPOByTanggal_,
    ambilItemHistoryPOKeluar_: ambilItemHistoryPOKeluar_,
    buatMapDataBersih_: buatMapDataBersih_,
    cariNilaiDataBersih_: cariNilaiDataBersih_,
    isiPOKeluarDariHistory_: isiPOKeluarDariHistory_,
    updateStatusPembelianAll_: updateStatusPembelianAll_,
    writeRingkasanToko_: writeRingkasanToko_,
    paksaRefreshMirrorTempoToko_: paksaRefreshMirrorTempoToko_,
    statusFinalCustomer_: statusFinalCustomer_
  };
  return WEB_INV_CACHE_;
}

// kunci untuk kode salinan = kunci skrip yang sama dengan gateway; selama dipegang,
// webGateway_ menjalankan aksi langsung (tidak antre ke kunci yang sama dua kali)
function webInvLockService_() {
  const bungkus = function () {
    const asli = LockService.getScriptLock();
    return {
      waitLock: function (ms) { asli.waitLock(ms); WEB_LOCK_DIPEGANG_ = true; },
      tryLock: function (ms) { const ok = asli.tryLock(ms); if (ok) WEB_LOCK_DIPEGANG_ = true; return ok; },
      releaseLock: function () { WEB_LOCK_DIPEGANG_ = false; asli.releaseLock(); },
      hasLock: function () { return asli.hasLock(); }
    };
  };
  return { getScriptLock: bungkus, getDocumentLock: bungkus, getUserLock: bungkus };
}

// SpreadsheetApp untuk kode salinan: sama dengan aslinya, kecuali tidak boleh memunculkan popup/alert
// Semua sheet yang disentuh LANGSUNG oleh kode salinan bersifat BACA SAJA: di inventory
// tulisan seperti itu masuk ke sheet LOKAL inventory (salinan/ringkasan), jadi di sini
// diabaikan supaya database tidak pernah ditulis di luar jalur gateway.
function webInvBacaSaja_(obj) {
  if (obj === null || typeof obj !== 'object') return obj;
  return new Proxy(obj, {
    get: function (t, k) {
      const v = t[k];
      if (typeof v !== 'function') return v;
      const nama = String(k);
      if (/^(set|clear|insert|delete|append|sort|hide|show|unhide|protect|merge|breakApart|activate|copyTo|moveTo|autoResize|trim|remove|add|flush|createFilter|createTextFinder|uncheck|check|randomize|shiftRow|shiftColumn|expand|collapse)/.test(nama)) {
        return function () { return webInvBacaSaja_(t); };
      }
      return function () { return webInvBacaSaja_(v.apply(t, arguments)); };
    }
  });
}

// spreadsheet untuk kode salinan: database_SANFELO, kecuali DATA_BERSIH & DISKON dari inventory
function webInvSpreadsheet_() {
  const ss = webSS_(); // mode latihan: file latihan
  const ro = webInvBacaSaja_(ss);
  return new Proxy(ro, {
    get: function (t, k) {
      if (k === 'getSheetByName') {
        return function (nama) {
          // saat mencetak: sheet form (SURAT JALAN, INVOICE, ...) = file CETAK, boleh ditulis
          if (WEB_CETAK_SS_ && WEB_SHEET_CETAK_.indexOf(nama) !== -1) {
            return WEB_CETAK_SS_.getSheetByName(nama) || (nama === 'INVOICE PALET' ? null : WEB_CETAK_SS_.getSheetByName(nama));
          }
          return webInvBacaSaja_(WEB_SHEET_MASTER_.indexOf(nama) !== -1 ? webMaster_(nama) : ss.getSheetByName(nama));
        };
      }
      if (k === 'insertSheet' && WEB_CETAK_SS_) {
        return function () { return WEB_CETAK_SS_.insertSheet.apply(WEB_CETAK_SS_, arguments); };
      }
      return t[k];
    }
  });
}

function webInvSpreadsheetApp_() {
  const asli = SpreadsheetApp;
  return {
    getActiveSpreadsheet: function () { return webInvSpreadsheet_(); },
    getActive: function () { return webInvSpreadsheet_(); },
    flush: function () { asli.flush(); },
    // file DATABASE dibuka langsung oleh inventory (mis. PETTY CASH) -> boleh ditulis, sama seperti di inventory
    openById: function (id) { return id === asli.getActiveSpreadsheet().getId() ? webSS_() : webInvBacaSaja_(id === INVENTORY_SS_ID_ ? webInventorySS_() : asli.openById(id)); }, // file DATABASE (latihan: file latihan)
    // popup spreadsheet tidak ada di web: pesan alert dijadikan error supaya tampil di web
    getUi: function () {
      const gagal = function (m) { throw new Error(String(m && m.message ? m.message : m)); };
      return {
        // saat alat dijalankan dari web (WEB_UI_KUMPUL_ berisi array), pesan alert dikumpulkan & konfirmasi dianggap YA
        alert: function (a, b) {
          const teks = typeof b === 'string' ? a + ': ' + b : a;
          if (WEB_UI_KUMPUL_) { WEB_UI_KUMPUL_.push(String(teks)); return 'YES'; }
          gagal(teks);
        },
        prompt: function () { gagal('Isian popup spreadsheet tidak tersedia di web.'); },
        showModalDialog: function () { gagal('Popup spreadsheet tidak tersedia di web.'); },
        showSidebar: function () { gagal('Sidebar spreadsheet tidak tersedia di web.'); },
        ButtonSet: { OK: 'OK', OK_CANCEL: 'OK_CANCEL', YES_NO: 'YES_NO', YES_NO_CANCEL: 'YES_NO_CANCEL' },
        Button: { OK: 'OK', YES: 'YES', NO: 'NO', CANCEL: 'CANCEL', CLOSE: 'CLOSE' }
      };
    },
    newDataValidation: function () { return asli.newDataValidation(); },
    BorderStyle: asli.BorderStyle, CopyPasteType: asli.CopyPasteType, ProtectionType: asli.ProtectionType
  };
}
