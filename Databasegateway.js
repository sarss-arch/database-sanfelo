/* ======================================================
   DatabaseGateway.js  (v2 — lengkap untuk HISTORY_PO_MASUK)
   ------------------------------------------------------
   FILE INI DI-PASTE / DI-DEPLOY DI SISI database_SANFELO
   (spreadsheet database, yang cuma bisa diakses IT & owner).

   v2 nambahin dukungan buat SEMUA aksi yang nulis ke
   HISTORY_PO_MASUK (bukan cuma "tambah baris" pas posting PO
   baru, tapi juga Cancel PO, Batal Cancel, bersihkan data
   double, dan update status "DONE SJ").

   SETUP: sama seperti sebelumnya (lihat komentar lama) — kalau
   sudah pernah deploy versi lama, tinggal:
   1) Ganti isi file DatabaseGateway.js lama dengan isi file ini.
   2) Deploy -> Manage deployments -> pencet ikon pensil (edit) di
      deployment yang sudah ada -> Version: "New version" -> Deploy.
      (JANGAN bikin "New deployment" lagi, nanti URL-nya beda dan
      GatewayClient.js di sisi inventory harus diganti lagi.)

   STRUKTUR KOLOM HISTORY_PO_MASUK yang dipakai gateway ini:
   A JENIS, B NO PO, C TGL PO, D CUSTOMER, E DISKON, F ALAMAT,
   G TELP, H TGL_KIRIM, I KODE, J NAMA, K QTY, L STATUS,
   M HARGA_UNIT, N TOTAL, O TIMESTAMP_SIMPAN, P LINE, Q ACTION
   (TIDAK dipakai di sisi database — action-trigger cuma relevan
   di sisi inventory tempat admin klik), R METHOD,
   S (kolom 19) BACKUP_STATUS (dipakai pas Cancel/Batal Cancel).
   ====================================================== */

// Ganti dengan kode rahasia sendiri (bebas, asal panjang & unik).
// ⚠️ HARUS SAMA PERSIS dengan GATEWAY_SECRET_TOKEN di GatewayClient.js sisi
// inventory_SANFELO -- kalau beda, semua aksi ditolak "Token tidak valid".
const SECRET_TOKEN = 'rahasia-sanfelo-2026';

// Sheet yang boleh disentuh action generik "appendRows" (dipakai buat
// sheet2 yang BELUM punya aksi khusus, + sheet "TEST" buat uji coba).
const SHEET_YANG_DIIZINKAN = [
  'TEST',
  'SPAREPART',
  'PENJUALAN',
  'EKSPEDISI',
  'TEMPO TOKO',
  'PEMBELIAN',
  'HUTANG',
  'PETTY CASH'
];

const HISTORY_PO_MASUK_SHEET = 'HISTORY_PO_MASUK';
const HISTORY_PO_MAIN_COLS = 18; // A..R
const HISTORY_PO_BACKUP_COL = 19; // S
const HISTORY_PO_STATUS_COL = 12; // L

// ⬅️ BARU (migrasi HISTORY_PO_KELUAR): struktur kolomnya —
// A NO PO, B TGL, C NAMA TOKO, D NAMA BARANG, E JUMLAH, F KETERANGAN,
// G TIMESTAMP, H QTY TERKIRIM (diisi manual oleh admin, TIDAK ikut
// di-mirror IMPORTRANGE di sisi inventory — persis kayak kolom Q di
// HISTORY_PO_MASUK, bedanya ini beneran data, bukan cuma tombol aksi,
// jadi setiap kali admin ngetik di kolom H lokal, langsung dikirim balik
// ke sini lewat aksi "updateQtyTerkirimPoKeluar" biar database tetap
// punya salinannya), I NILAI DATA BERSIH.
const HISTORY_PO_KELUAR_SHEET = 'HISTORY_PO_KELUAR';
const HISTORY_PO_KELUAR_COLS = 9; // A..I
const HISTORY_PO_KELUAR_QTY_COL = 8; // H
const HISTORY_PO_KELUAR_HEADER = [
  'NO PO', 'TGL', 'NAMA TOKO', 'NAMA BARANG', 'JUMLAH',
  'KETERANGAN', 'TIMESTAMP', '', 'NILAI DATA BERSIH'
];

// ⬅️ BARU (migrasi TEMPO TOKO): struktur kolom A..E — A TGL KIRIM,
// B CUSTOMER, C NO INV, D TEMPO, E STATUS PEMBELIAN (UNLOCKED / LOCKED /
// "-N HARI"). Kolom G:H (ringkasan per customer + dropdown lock/unlock)
// SENGAJA TIDAK ada di sini sama sekali -- itu murni lokal di sisi
// inventory (dipakai juga oleh cekTempoTokoBlokir_ yang butuh baca fresh
// tanpa delay buat nge-blokir Surat Jalan toko LOCKED), jadi gak pernah
// disentuh/disimpan di database. Polanya "replace semua" (bukan per-PO
// kayak HISTORY_PO_MASUK/KELUAR) karena logic sync yang lama emang
// ngitung ULANG SELURUH tabel A:E tiap kali jalan, bukan incremental.
const TEMPO_TOKO_SHEET = 'TEMPO TOKO';
const TEMPO_TOKO_COLS = 5; // A..E

// ⬅️ BARU (migrasi EKSPEDISI): sheet log biaya tambahan (ekspedisi/palet),
// PURE APPEND -- gak pernah di-update/dihapus/dibaca balik oleh kode
// manapun, jadi migrasinya paling simpel: SELURUH kolom A-M di-mirror
// IMPORTRANGE di inventory, gak ada kolom yang perlu tetap lokal.
// A NO INV, B REFERENSI, C NAMA EKSPEDISI, D TANGGAL, E TEMPO (teks
// dd/mm/yyyy, bukan Date), F NAMA TOKO, G ALAMAT, H BARANG, I QTY,
// J SAT, K KET BRG, L HARGA, M METHOD.
const EKSPEDISI_SHEET = 'EKSPEDISI';
const EKSPEDISI_COLS = 13; // A..M
const EKSPEDISI_QTY_COL = 9; // I
const EKSPEDISI_HARGA_COL = 12; // L

// ⬅️ BARU (migrasi HUTANG, September 2026): sheet paling kompleks yang
// dimigrasi sejauh ini karena punya 3 jalur tulis independen (semua
// sekarang lewat gateway ini):
//   1) syncHutangDariPembelian -- insert/update per NO PB dari hasil scan
//      PEMBELIAN (PEMBELIAN sendiri TETAP LOKAL, belum dimigrasi --
//      gateway cuma terima hasil groupingnya, bukan baca PEMBELIAN
//      langsung). Dipanggil OTOMATIS oleh TANDA_TERIMA_BRG.js tiap ada
//      Tanda Terima Barang baru/edit, jadi harus tetap dibungkus
//      try/catch di sisi inventory supaya gak pernah bikin proses
//      posting itu gagal.
//   2) bayarHutangMultiPB -- popup "Bayar Hutang", bisa bayar beberapa PB
//      (baris HUTANG) sekaligus dalam 1x submit, tiap PB independen
//      (bukan FIFO). Riwayat cicilan ditulis ke kolom K dst, di kolom
//      kosong pertama MILIK BARIS ITU SENDIRI.
//   3) groupingInvoiceHutang -- popup "Grouping HUTANG", gabungin
//      beberapa NO SJ (baris HUTANG) jadi 1 NO INV + isi TRANSAKSI BULAN
//      & JATUH TEMPO.
// Ketiganya sama-sama berakhir manggil updateStatusJatuhTempoGatewayHutang_
// di bawah buat refresh STATUS (kolom I) + warna baris -- pola yang sama
// persis kayak versi lokal lama (semua manggil updateStatusJatuhTempo_
// yang sama di InvoiceGroupingPopup.gs), cuma sekarang jalan di sini
// terhadap sheet HUTANG yang ASLI, bukan mirror.
//
// STRUKTUR KOLOM HUTANG:
// A TRANSAKSI BULAN, B NO PB, C NO INV, D NO SJ, E SUPPLIER, F TGL BELI,
// G JATUH TEMPO, H TOTAL, I STATUS, J SISA BAYAR, K dst RIWAYAT
// PEMBAYARAN (teks "dd/MM/yyyy - RpX - METODE", nambah ke kanan tiap ada
// pembayaran baru -- makanya IMPORTRANGE di sisi inventory harus
// di-mirror cukup lebar, disarankan A:Z, biar kebagian ruang buat
// belasan kali cicilan per PB).
const HUTANG_SHEET = 'HUTANG';
const HUTANG_COL_BULAN = 1;
const HUTANG_COL_PB = 2;
const HUTANG_COL_NOINV = 3;
const HUTANG_COL_NOSJ = 4;
const HUTANG_COL_SUPPLIER = 5;
const HUTANG_COL_TGLBELI = 6;
const HUTANG_COL_JATUHTEMPO = 7;
const HUTANG_COL_TOTAL = 8;
const HUTANG_COL_STATUS = 9;
const HUTANG_COL_SISABAYAR = 10;
const HUTANG_RIWAYAT_START_COL = 11; // K

const HUTANG_STATUS_WARNA = {
  MENDEKATI: { bg: '#FFF200', font: '#000000' },
  LEWAT:     { bg: '#FF0000', font: '#FFFFFF' },
  LUNAS:     { bg: '#34A853', font: '#FFFFFF' }
};

// ⬅️ BARU (migrasi SPAREPART, September 2026): sheet transaksi sparepart
// (1 baris = 1 item barang, banyak baris bisa share NO SJ/NO INV yang
// sama kalau 1 transaksi isi beberapa barang). TIDAK ADA kolom
// lokal-only sama sekali di sini (beda dari TEMPO TOKO G:H atau
// HISTORY_PO_KELUAR H) — seluruh A:S (19 kolom) murni data & aman
// di-mirror IMPORTRANGE penuh.
//
// STRUKTUR KOLOM SPAREPART:
// A NO SJ, B NO INV, C REFRENSI, D TANGGAL, E TEMPO, F NO KEND,
// G NAMA TOKO, H ALAMAT, I BARANG, J QTY, K SAT, L KET BRG, M KET INV,
// N HARGA, O HPP BARANG, P METHOD, Q SALES, R LABA, S STATUS ORDER.
//
// Pola sync: **delete-by-NO SJ lalu insert** (sama pola kayak
// HISTORY_PO_KELUAR/HISTORY_PO_MASUK) — dipakai baik buat transaksi
// baru (gak ada yang ke-delete, NO SJ-nya emang baru) maupun mode EDIT
// (baris lama utk NO SJ itu ditimpa versi baru). Keputusan "ini mode
// edit atau bukan" tetap di sisi inventory (dari PropertiesService
// lokal) — gateway cuma terima NO SJ + baris final yang mau ditulis.
const SPAREPART_SHEET = 'SPAREPART';
const SPAREPART_COLS = 19; // A..S
const SPAREPART_COL_NOSJ = 1;   // A
const SPAREPART_COL_NOINV = 2;  // B
const SPAREPART_COL_TANGGAL = 4; // D
const SPAREPART_COL_TEMPO = 5;   // E
const SPAREPART_COL_STATUS = 19; // S — STATUS ORDER (bukan kolom 18/R, itu LABA!)

// ⬅️ BARU (migrasi PEMBELIAN, Oktober 2026): sheet item Tanda Terima
// Barang (1 baris = 1 item barang, banyak baris bisa share NO PB/NO SJ
// yang sama kalau 1 Tanda Terima isi beberapa barang/PO sekaligus).
// TIDAK ADA kolom lokal-only — seluruh A:O (15 kolom) murni data, aman
// di-mirror IMPORTRANGE penuh (A2:O). Kolom N (STATUS LUNAS) memang
// dikirim kosong pas simpan awal, tapi diisi BELAKANGAN oleh aksi
// terpisah (tandaiLunasPembelian) — bukan kolom lokal, tetap bagian
// dari data asli, cuma ditulis di waktu yang beda.
//
// STRUKTUR KOLOM PEMBELIAN:
// A NO PB, B NO PO, C NO SJ, D SUPPLIER, E TGL BELI, F NO KEND, G KODE,
// H NAMA BARANG, I QTY, J HARGA, K NAMA TOKO, L PENERIMA, M PENGIRIM,
// N STATUS LUNAS (diisi belakangan), O STATUS DISKON SET.
//
// Pola sync tulis: **delete-by-(NO SJ + SUPPLIER) lalu insert** — mirip
// SPAREPART, tapi key gandanya NO SJ+SUPPLIER (bukan cuma NO SJ), karena
// NO SJ bisa collision antar PB beda supplier (lihat catatan di
// PembelianHutangSync.js). Satu aksi ini dipakai buat 3 skenario dari
// simpanTandaTerimaKePembelian() di sisi inventory: (a) item baru (gak
// ada yg match, rows dikirim, gak ada yg kehapus), (b) "Timpa" (match
// ketemu, baris lama kehapus, baris baru ditulis), (c) "Edit" (match
// ketemu, baris lama kehapus, rows dikirim kosong — caller cuma mau
// hapus, user edit manual lalu simpan ulang lewat panggilan kedua).
const PEMBELIAN_SHEET = 'PEMBELIAN';
const PEMBELIAN_COLS = 15; // A..O
const PEMBELIAN_COL_PB = 1;             // A
const PEMBELIAN_COL_NOSJ = 3;            // C
const PEMBELIAN_COL_SUPPLIER = 4;        // D
const PEMBELIAN_COL_TGLBELI = 5;         // E
const PEMBELIAN_COL_STATUS_LUNAS = 14;   // N

// ⬅️ BARU (migrasi PENJUALAN, Oktober 2026): sheet paling kompleks sejauh
// ini karena punya 4 jalur tulis independen (semua sekarang lewat gateway
// ini), ditambah 1 masalah arsitektur yang baru ketahuan pas riset migrasi
// ini: fitur "Urutkan berdasarkan Customer" + baris TOTAL kuning virtual
// di sidebar Filter (FILTER_SIDEBAR.gs sisi inventory) itu FISIK NULIS
// ULANG / NYISIP baris ke PENJUALAN -- begitu PENJUALAN jadi mirror
// IMPORTRANGE, dua hal itu PASTI ERROR (gak bisa nulis ke range yang
// isinya formula array). Keputusan (dari diskusi sama user): SORT
// dipindah jadi aksi gateway ini juga (aksiSortPenjualanByCustomer_,
// jalan di database asli, mirror lokal otomatis ikut urut lewat
// IMPORTRANGE), baris TOTAL virtual DIHILANGKAN dari tampilan sheet --
// totalnya tetap ada, cuma dipindah ke panel sidebar (breakdown per
// customer dihitung CLIENT-SIDE dari data lokal yang sedang tampil,
// gak butuh aksi gateway baru -- lihat applyFilterSidebar di
// FILTER_SIDEBAR.gs sisi inventory).
//
// 4 jalur tulis yang dipindah ke sini:
//   1) upsertPenjualan -- dipakai BERSAMA oleh syncPenjualanFromTransaksi_
//      (sumber: TRANSAKSI, tetap lokal) & syncPenjualanFromSparepart_
//      (sumber: SPAREPART, sudah mirror). Upsert by NO INV (kolom C) +
//      regroup per customer + backfill SISA BAYAR + auto-LUNAS + re-color,
//      SEMUA jalan di sini sekarang (bukan lokal lagi).
//   2) prosesPembayaranPenjualan -- versi database dari prosesPembayaran_
//      (popup Pembayaran), termasuk generate No Post (POST-NNN) server-side
//      lewat scan ulang PENJUALAN yang ASLI (bukan mirror -- No Post HARUS
//      gak pernah nabrak walau ada 2 admin posting hampir bersamaan, sama
//      kayak alasan lock di prosesRequestGateway_ di atas).
//   3) updateBiayaTambahanPenjualan -- versi database dari bagian PENJUALAN
//      di simpanBiayaTambahan_() (BiayaTambahan.gs): update kolom H
//      (EKSPEDISI) / I (PALET) + recompute SISA BAYAR & STATUS 1 baris.
//   4) sortPenjualanByCustomer -- lihat catatan arsitektur di atas.
//
// STRUKTUR KOLOM PENJUALAN (sama persis dengan PENJUALAN_COL_* di
// Pembayaran.gs sisi inventory):
// A TGL KIRIM, B CUSTOMER, C NO INV, D METHOD, E TOTAL,
// F STATUS ("CICIL"/"LUNAS"), G SISA BAYAR, H EKSPEDISI, I PALET,
// J dst RIWAYAT CICILAN (teks "POST-NNN - dd/MM/yyyy - RpX - METODE",
// nambah ke kanan tiap ada pembayaran baru -- makanya IMPORTRANGE di
// sisi inventory harus di-mirror cukup lebar, sama kayak saran HUTANG,
// disarankan A:Z atau lebih).
//
// Catatan No Post: counter cache di sheet "SYSTEM" sel B4 dicari DI
// DATABASE INI (bukan di inventory) -- kalau database_SANFELO belum
// punya sheet SYSTEM, fungsi tetap jalan benar (fallback ke scan ulang
// data doang, sedikit lebih lambat tapi tetap akurat), gak wajib dibikin.
const PENJUALAN_SHEET = 'PENJUALAN';
const PENJUALAN_COL_TGLKIRIM = 1;   // A
const PENJUALAN_COL_CUSTOMER = 2;   // B
const PENJUALAN_COL_NOINV = 3;      // C
const PENJUALAN_COL_METHOD = 4;     // D
const PENJUALAN_COL_TOTAL = 5;      // E
const PENJUALAN_COL_STATUS = 6;     // F
const PENJUALAN_COL_SISA = 7;       // G
const PENJUALAN_COL_EKSPEDISI = 8;  // H
const PENJUALAN_COL_PALET = 9;      // I
const PENJUALAN_COL_CICILAN_START = 10; // J
const PENJUALAN_WARNA_LUNAS = '#d9ead3';
const PENJUALAN_WARNA_LEBIH = '#cfe2f3';
const PENJUALAN_NOPOST_PREFIX_ = 'POST-';
const PENJUALAN_SYSTEM_NOPOST_CELL_ = 'B4';

// ⬅️ BARU (Oktober 2026) — NOMOR PO MASUK BERSAMA (SANFELO + SL)
// ------------------------------------------------------------------
// Ditemukan dari laporan user: format nomor "No PO Masuk" (#YYYY-NNN,
// lihat poBaru() di PO_HISTORY.js) TERNYATA dipakai BERSAMAAN oleh 2
// sistem terpisah — inventory_SANFELO (project ini) DAN inventory_SL
// (project Apps Script lain, spreadsheet lain, bukan bagian dari
// migrasi database_SANFELO ini) — nomor yang sudah kepake di satu sisi
// TIDAK BOLEH kepake lagi di sisi satunya, nomornya harus saling
// nyambung/melanjutkan satu urutan yang sama.
//
// Sebelum ini, masing-masing sisi punya counter LOKAL sendiri-sendiri
// (SYSTEM!B3/C3 di tiap spreadsheet, lihat getCounterDenganResetPeriode_
// di PO_HISTORY.js) yang sama sekali gak saling tau — ini pernah
// kejadian nyata bikin data dobel di HISTORY_PO_MASUK SANFELO (PO
// #2026-223/#2026-224 ketinggalan orphan gara-gara nomor baru digenerate
// padahal lagi di tengah proses EDIT PO lama — lihat log proyek).
//
// Fix: counter-nya sekarang dipindah ke SATU tempat (cell SYSTEM!B6/C6
// di database_SANFELO INI — sengaja beda cell dari
// PENJUALAN_SYSTEM_NOPOST_CELL_/B4 yang sudah dipakai lebih dulu untuk
// hal lain), dan KEDUA sisi (SANFELO & SL) manggil aksi
// "getNomorPoMasukBersama" di bawah buat dapet nomor PO MASUK
// berikutnya — request dari kedua sisi otomatis ke-serialize lewat
// LockService yang sama yang sudah ngunci SEMUA aksi gateway (lihat
// prosesRequestGateway_), jadi gak mungkin ada 2 device/sistem dapet
// nomor yang sama. Sisi SL (yang gak pernah ikut migrasi gateway
// SANFELO) dikasih file client minimal baru khusus buat manggil aksi
// ini doang, lewat Web App yang SAMA (SECRET_TOKEN yang SAMA) yang
// sudah dipakai inventory_SANFELO — gak perlu deploy baru.
//
// SETUP WAJIB (SEKALI SAJA, manual, SEBELUM dipakai): pastikan ada
// sheet bernama PERSIS "SYSTEM" di database_SANFELO ini (buat baru
// kalau belum ada). Isi cell B6 dengan ANGKA PO MASUK TERTINGGI yang
// sudah pernah kepake tahun berjalan — cek kedua sisi (SANFELO DAN
// SL), ambil yang LEBIH BESAR — dan cell C6 dengan tahun berjalan
// (contoh: "2026"). Setelah seeding awal ini, counter jalan sendiri,
// gak perlu disentuh manual lagi kecuali mau koreksi.
const SHARED_PO_MASUK_COUNTER_CELL_ = 'B6';
const SHARED_PO_MASUK_PERIODE_CELL_ = 'C6';
const SHARED_PO_MASUK_LOG_SHEET_ = 'LOG_PO_BERSAMA'; // opsional, dilewati diam-diam kalau sheet ini gak ada

// Sheet yang boleh DIBACA fresh langsung dari database (aksi "bacaSheet"),
// dipakai cuma buat pengecekan yang gak boleh kena delay IMPORTRANGE —
// contoh: cek toko LOCKED sebelum bikin Surat Jalan, atau cek nomor PO
// Keluar berikutnya berdasarkan tanggal, atau baca ulang TEMPO TOKO A:E
// sebelum recompute status/perpanjang tempo, atau cek status LUNAS
// PENJUALAN sebelum kasih tau TRANSAKSI/DETAIL_TRANSAKSI boleh DONE.
// Sengaja daftar terpisah dari SHEET_YANG_DIIZINKAN (yang itu buat nulis).
const SHEET_BACA_DIIZINKAN = [
  'HISTORY_PO_MASUK',
  'HISTORY_PO_KELUAR',
  'TEMPO TOKO',
  'HUTANG',
  'SPAREPART',
  'PEMBELIAN',
  'PENJUALAN',
  'LOG_INVOICE_CETAK',  // ⬅️ BARU: dibaca menu "Catat Susulan Invoice Tercetak" di inventory
  'TRANSAKSI',          // ⬅️ BARU (migrasi TRANSAKSI, Oktober 2026) -- lihat GatewayTransaksi.js
  'DETAIL_TRANSAKSI'    // ⬅️ BARU (migrasi DETAIL_TRANSAKSI, Oktober 2026)
];

const ACTION_HANDLERS = {
  appendRows: aksiAppendRows_,
  syncPoMasuk: aksiSyncPoMasuk_,
  cancelPo: aksiCancelPo_,
  uncancelPo: aksiUncancelPo_,
  bersihkanDuplikat: aksiBersihkanDuplikat_,
  updateStatusDoneSjDelta: aksiUpdateStatusDoneSjDelta_,
  bacaSheet: aksiBacaSheet_,
  syncPoKeluar: aksiSyncPoKeluar_,
  updateQtyTerkirimPoKeluar: aksiUpdateQtyTerkirimPoKeluar_,
  setQtyTerkirimPoKeluarBatch: function (ss, body) { return aksiSetQtyTerkirimPoKeluarBatch_(ss, body); },
  isiBulanHutangKosong: function (ss, body) { return aksiIsiBulanHutangKosong_(ss, body); },
  syncNilaiDataBersihBatch: aksiSyncNilaiDataBersihBatch_,
  syncTempoTokoFull: aksiSyncTempoTokoFull_,
  perpanjangTempoPerCustomer: aksiPerpanjangTempoPerCustomer_,
  appendEkspedisi: aksiAppendEkspedisi_,
  syncHutangDariPembelian: aksiSyncHutangDariPembelian_,
  bayarHutangMultiPB: aksiBayarHutangMultiPB_,
  groupingInvoiceHutang: aksiGroupingInvoiceHutang_,
  syncSparepart: aksiSyncSparepart_,
  tandaiDoneSparepart: aksiTandaiDoneSparepart_,
  syncPembelian: aksiSyncPembelian_,
  tandaiLunasPembelian: aksiTandaiLunasPembelian_,
  tandaiLunasPembelianBatch: aksiTandaiLunasPembelianBatch_,
  upsertPenjualan: aksiUpsertPenjualan_,
  sortPenjualanByCustomer: aksiSortPenjualanByCustomer_,
  prosesPembayaranPenjualan: aksiProsesPembayaranPenjualan_,
  updateBiayaTambahanPenjualan: aksiUpdateBiayaTambahanPenjualan_,
  getNomorPoMasukBersama: aksiGetNomorPoMasukBersama_,
  perbaikiStatusHistoryPOMassalBatch: aksiPerbaikiStatusHistoryPOMassalBatch_,
  perbaikiWarnaStatusHistoryPOBatch: aksiPerbaikiWarnaStatusHistoryPOBatch_,
  peekNomorPoMasukBersama: aksiPeekNomorPoMasukBersama_,
  // ⬅️ dibungkus function: isinya ada di GatewayInvoiceCetak.js, jadi aman
  // apa pun urutan file di project
  cekStatusInvoiceCetak: function (ss, body) { return aksiCekStatusInvoiceCetak_(ss, body); },
  catatInvoiceTercetak: function (ss, body) { return aksiCatatInvoiceTercetak_(ss, body); },
  batalInvoiceTercetak: function (ss, body) { return aksiBatalInvoiceTercetak_(ss, body); },
  // ⬅️ BARU (migrasi TRANSAKSI + DETAIL_TRANSAKSI, Oktober 2026) -- isi
  // fungsinya ada di file GatewayTransaksi.js
  // (dibungkus function supaya tidak bergantung urutan file di project)
  simpanTransaksiDraft: function (ss, body) { return aksiSimpanTransaksiDraft_(ss, body); },
  tandaiDoneTransaksi: function (ss, body) { return aksiTandaiDoneTransaksi_(ss, body); },
  ubahStatusTransaksi: function (ss, body) { return aksiUbahStatusTransaksi_(ss, body); },
  updateTempoTransaksiBatch: function (ss, body) { return aksiUpdateTempoTransaksiBatch_(ss, body); }
};

function doPost(e) {
  let hasil;
  try {
    hasil = prosesRequestGateway_(e);
  } catch (err) {
    hasil = { ok: false, error: String(err && err.message ? err.message : err) };
  }
  return ContentService
    .createTextOutput(JSON.stringify(hasil))
    .setMimeType(ContentService.MimeType.JSON);
}

function prosesRequestGateway_(e) {
  if (!e || !e.postData || !e.postData.contents) {
    throw new Error('Request kosong / bukan POST JSON.');
  }

  const body = JSON.parse(e.postData.contents);

  if (!body.token || body.token !== SECRET_TOKEN) {
    throw new Error('Token tidak valid. Request ditolak.');
  }

  const action = String(body.action || '').trim();
  const handler = ACTION_HANDLERS[action];
  if (!handler) {
    throw new Error('Action "' + action + '" tidak dikenali gateway ini.');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ⬅️ FIX (performance, Oktober 2026): aksi "bacaSheet" itu READ-ONLY --
  // gak pernah nulis apapun ke sheet manapun (lihat aksiBacaSheet_, isinya
  // cuma getRange/getValues). Baca gak butuh saling eksklusif sama baca
  // lain -- cuma TULIS yang perlu dikunci biar gak saling timpa. Sebelum
  // fix ini, SEMUA aksi (termasuk baca) ngantre di LockService yang SAMA,
  // jadi satu proses simpan yang beberapa kali baca sheet besar (mis.
  // simpanTandaTerimaKePembelian() di inventory_SANFELO, yang baca
  // PEMBELIAN penuh 2-3x dalam satu kali jalan: cek duplikat, sync HUTANG,
  // generate NO PB) bisa nahan SEMUA request lain dari SELURUH sistem
  // (inventory_SANFELO maupun inventory_SL) sampai beres -- padahal baca
  // doang, gak ada resiko data ke-timpa. Sekarang 'bacaSheet' DILEWATKAN
  // dari lock sama sekali, cuma aksi TULIS yang tetap wajib antre di bawah.
  // ⬅️ BARU (Oktober 2026): "peekNomorPoMasukBersama" juga READ-ONLY (cuma
  // intip nilai counter, gak pernah nulis balik -- beda dari
  // "getNomorPoMasukBersama" yang BENERAN increment+klaim) -- sama alasannya
  // kayak "bacaSheet" di atas, gak perlu antre di belakang lock tulis.
  if (action === 'bacaSheet' || action === 'peekNomorPoMasukBersama') {
    return handler(ss, body);
  }

  // ⬅️ FIX (race condition): file ini adalah SATU-SATUNYA jalur tulis ke
  // database_SANFELO begitu pemisahan database jalan -- artinya SEMUA
  // aksi simpan dari inventory_SANFELO (dan nanti mungkin lebih dari satu
  // orang/tab sekaligus) berakhir lewat doPost() di sini. Web App Apps
  // Script MEMANG dirancang buat nanganin beberapa request bersamaan --
  // jadi tanpa lock, dua request yang hampir bersamaan (mis. dua admin
  // nge-Cancel PO beda, atau nge-post PO baru dari dua device) bisa
  // saling interleave baca-hitung-tulis-nya persis kayak bug poBaru()/
  // prosesPembayaranMultiPB() yang sudah diperbaiki di sisi inventory --
  // cuma di sini risikonya lebih besar karena ini titik tunggal yang
  // dipakai bersama. Sekarang tiap request dikunci dulu sebelum jalanin
  // handler-nya; request lain nunggu antre (maks 10 detik, gateway ini
  // cuma dipanggil sinkron per aksi simpan jadi wajar nunggu sebentar),
  // baru dapat error jelas kalau beneran macet -- bukan data ke-timpa
  // diam-diam.
  const lock = LockService.getScriptLock();
  const dapatLock = lock.tryLock(10000); // tunggu maks 10 detik

  if (!dapatLock) {
    throw new Error('Database sedang sibuk memproses aksi lain, coba lagi sebentar.');
  }

  try {
    return handler(ss, body);
  } finally {
    lock.releaseLock();
  }
}

/* ------------------------------------------------------
   AKSI GENERIK: tambah baris di sheet mana pun yang ada di
   daftar SHEET_YANG_DIIZINKAN.
   body: { sheet, rows: [[...], [...]] }
   ------------------------------------------------------ */
function aksiAppendRows_(ss, body) {
  const sheetName = String(body.sheet || '').trim();
  if (!sheetName) throw new Error('Nama sheet tujuan (sheet) tidak dikirim.');
  if (SHEET_YANG_DIIZINKAN.indexOf(sheetName) === -1) {
    throw new Error('Sheet "' + sheetName + '" tidak ada di daftar SHEET_YANG_DIIZINKAN.');
  }

  const rows = body.rows;
  if (!Array.isArray(rows) || rows.length === 0) {
    throw new Error('Data (rows) kosong / bukan array.');
  }

  const sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error('Sheet "' + sheetName + '" tidak ditemukan di spreadsheet database.');

  const startRow = sh.getLastRow() + 1;
  sh.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);

  return { ok: true, sheet: sheetName, baris_ditulis: rows.length, mulai_dari_baris: startRow };
}

/* ------------------------------------------------------
   AKSI: syncPoMasuk — versi database dari syncToHistory() +
   deleteHistoryPORows_() + applyHistoryPOActionValidation_() +
   bersihkanDuplikatUntukPO_() yang dulu jalan lokal di
   inventory_SANFELO. Sekarang jalan di sini, terhadap sheet
   HISTORY_PO_MASUK yang ASLI (bukan mirror).

   body: { jenis: "MASUK", poNo: "...", rows: [[18 kolom], ...] }
   ------------------------------------------------------ */
function aksiSyncPoMasuk_(ss, body) {
  const jenis = String(body.jenis || 'MASUK').trim().toUpperCase();
  const poNo = String(body.poNo || '').trim();
  const rows = body.rows;

  if (!poNo) throw new Error('poNo tidak dikirim.');
  if (!Array.isArray(rows)) throw new Error('rows harus array (boleh kosong kalau memang PO tidak punya item).');

  const hist = ss.getSheetByName(HISTORY_PO_MASUK_SHEET);
  if (!hist) throw new Error('Sheet "' + HISTORY_PO_MASUK_SHEET + '" tidak ditemukan.');

  // 1) Hapus baris lama untuk PO ini (kalau ada, misal ini re-sync/edit)
  hapusBarisHistoryPOByJenisPoNo_(hist, jenis, poNo);

  // 2) Insert baris baru
  // ⬅️ TAMBAHAN: startRow di-return ke pemanggil (lihat mulai_dari_baris di
  // bawah) -- sisi inventory_SANFELO butuh nomor baris ini buat nyiapin
  // dropdown ACTION (kolom Q) di mirror lokal, karena kolom Q sengaja TIDAK
  // ikut di-mirror lewat IMPORTRANGE (baca komentar di kepala file) jadi gak
  // bisa ditentukan dari hasil baca mirror -- harus dikasih tau eksplisit
  // dari sini, sumber kebenarannya.
  // ⬅️ FIX (Oktober 2026): kolom tanggal (C TGL PO, H TGL_KIRIM) di body.rows
  // dikirim dari sisi inventory sebagai object Date -- begitu payload ini
  // di-JSON.stringify() buat dikirim ke Web App, Date OTOMATIS berubah jadi
  // teks ISO ("2026-09-30T00:00:00.000Z") oleh JSON.stringify sendiri (ini
  // perilaku bawaan JavaScript, bukan bug di PO_HISTORY.js). Semua aksi sync
  // LAIN yang nulis tanggal (TEMPO TOKO/HUTANG/SPAREPART/PEMBELIAN) SUDAH
  // nangani ini -- selalu parse ulang pakai parseTanggalDariGateway_() +
  // setNumberFormat('dd/mm/yyyy') sebelum ditulis. Aksi syncPoMasuk ini
  // (salah satu yang PALING AWAL dimigrasi, sebelum pola itu baku) ternyata
  // KELEWATAN -- rows ditulis APA ADANYA, jadi teks ISO mentah itu nangkring
  // sebagai teks biasa di kolom C/H (bukan tanggal beneran), gak rapi &
  // "nempel waktu" yang gak perlu. Sekarang disamakan: parse ulang C & H ke
  // Date yang benar, format dd/mm/yyyy, SEBELUM ditulis -- biar konsisten
  // sama baris2 lama & sama semua tab lain.
  let startRow = null;
  if (rows.length > 0) {
    const rowsTglFixed = rows.map(function (r) {
      const copy = r.slice();
      copy[2] = parseTanggalDariGateway_(copy[2]); // C TGL PO
      copy[7] = parseTanggalDariGateway_(copy[7]); // H TGL_KIRIM
      return copy;
    });

    startRow = hist.getLastRow() + 1;
    hist.getRange(startRow, 1, rowsTglFixed.length, HISTORY_PO_MAIN_COLS).setValues(rowsTglFixed);
    hist.getRange(startRow, 3, rowsTglFixed.length, 1).setNumberFormat('dd/mm/yyyy'); // C
    hist.getRange(startRow, 8, rowsTglFixed.length, 1).setNumberFormat('dd/mm/yyyy'); // H
  }

  // 3) Bersihkan data double khusus utk PO ini (jaga-jaga kalau ada race)
  jalankanDedupUntukPO_(hist, jenis, poNo);

  return { ok: true, poNo: poNo, baris_ditulis: rows.length, mulai_dari_baris: startRow };
}

function hapusBarisHistoryPOByJenisPoNo_(hist, jenis, poNo) {
  const lastRow = hist.getLastRow();
  if (lastRow < 2) return;

  const data = hist.getRange(2, 1, lastRow - 1, 2).getValues();
  for (let i = data.length - 1; i >= 0; i--) {
    const rowJenis = String(data[i][0] || '').trim().toUpperCase();
    const rowPoNo = String(data[i][1] || '').trim();
    if (rowJenis === jenis && rowPoNo === poNo) {
      hist.deleteRow(i + 2);
    }
  }
}

/* ------------------------------------------------------
   AKSI: cancelPo — versi database dari handleHistoryCancel().
   Nyari semua baris JENIS+NO PO yang cocok, backup status lama
   ke kolom S, set STATUS jadi CANCELLED, kasih warna pink.
   (Kolom Q / ACTION TIDAK disentuh di sini — itu urusan lokal
   di sisi inventory karena kolom itu memang harus tetap bisa
   diklik manual oleh admin, jadi TIDAK ikut di-mirror.)

   body: { jenis, poNo }
   ------------------------------------------------------ */
function aksiCancelPo_(ss, body) {
  const jenis = String(body.jenis || '').trim().toUpperCase();
  const poNo = String(body.poNo || '').trim();
  if (!poNo) throw new Error('poNo tidak dikirim.');

  const hist = ss.getSheetByName(HISTORY_PO_MASUK_SHEET);
  if (!hist) throw new Error('Sheet "' + HISTORY_PO_MASUK_SHEET + '" tidak ditemukan.');

  const lastRow = hist.getLastRow();
  if (lastRow < 2) return { ok: true, updatedCount: 0 };

  // ⬅️ FIX (bug kritis): kolom S (HISTORY_PO_BACKUP_COL = 19) baru pernah
  // ditulis kalau sudah pernah ada PO yang di-cancel sebelumnya. Kalau
  // hist.getLastColumn() masih < 19, pakai itu langsung sebagai lebar
  // data/range di bawah bikin tulisan backup status ke data[i][18] HILANG
  // DIAM-DIAM pas write-back, karena setValues() akhirnya nulis ke range
  // yang lebih sempit dari kolom S. Dipaksa minimal sampai kolom S.
  const lastCol = Math.max(hist.getLastColumn(), HISTORY_PO_BACKUP_COL);
  const data = hist.getRange(2, 1, lastRow - 1, lastCol).getValues();

  // ⬅️ FIX (performa): dulu tiap baris yang match PO ini baca STATUS
  // sekarang (getValue) lalu nulis backup (setValue), STATUS baru
  // (setValue), & warna (setBackground) SATU-SATU DI DALAM LOOP -- sampai
  // 4 panggilan API per baris, padahal loop-nya scan SELURUH
  // HISTORY_PO_MASUK. Sama kayak versi lokalnya (handleHistoryCancel_ di
  // inventory) yang sudah dibatch, di sini juga dikonversi: seluruh sheet
  // dibaca sekali (sudah kebawa lewat `data` di atas), warna dibaca
  // sekali (getBackgrounds), baris yang match diubah di memory, baris
  // yang tidak match dibiarkan persis, baru ditulis balik SEKALI lewat
  // setValues()/setBackgrounds() -- hasil akhir identik, cuma jauh lebih
  // sedikit panggilan API.
  const warnaGrid = hist.getRange(2, 1, lastRow - 1, lastCol).getBackgrounds();
  let updatedCount = 0;

  for (let i = 0; i < data.length; i++) {
    const rowJenis = String(data[i][0] || '').trim().toUpperCase();
    const rowPoNo = String(data[i][1] || '').trim();
    if (rowJenis !== jenis || rowPoNo !== poNo) continue;

    const currentStatus = data[i][HISTORY_PO_STATUS_COL - 1];
    data[i][HISTORY_PO_BACKUP_COL - 1] = currentStatus;
    data[i][HISTORY_PO_STATUS_COL - 1] = 'CANCELLED';
    for (let c = 0; c < lastCol; c++) warnaGrid[i][c] = '#f4cccc';
    updatedCount++;
  }

  if (updatedCount > 0) {
    hist.getRange(2, 1, data.length, lastCol).setValues(data);
    hist.getRange(2, 1, warnaGrid.length, lastCol).setBackgrounds(warnaGrid);
  }

  return { ok: true, updatedCount: updatedCount };
}

/* ------------------------------------------------------
   AKSI: uncancelPo — versi database dari handleHistoryUncancel().
   body: { jenis, poNo }
   ------------------------------------------------------ */
function aksiUncancelPo_(ss, body) {
  const jenis = String(body.jenis || '').trim().toUpperCase();
  const poNo = String(body.poNo || '').trim();
  if (!poNo) throw new Error('poNo tidak dikirim.');

  const hist = ss.getSheetByName(HISTORY_PO_MASUK_SHEET);
  if (!hist) throw new Error('Sheet "' + HISTORY_PO_MASUK_SHEET + '" tidak ditemukan.');

  const lastRow = hist.getLastRow();
  if (lastRow < 2) return { ok: true, updatedCount: 0 };

  // ⬅️ FIX (performa, Oktober 2026 — ditemukan lewat audit menyeluruh):
  // dulu tiap baris yang match PO ini baca backup-status (getValue),
  // nulis STATUS balik (setValue), hapus backup (clearContent), & nulis
  // warna (setBackground) SATU-SATU DI DALAM LOOP -- 4 panggilan API per
  // baris, padahal loop-nya scan SELURUH HISTORY_PO_MASUK. Sekarang
  // dibatch sama persis seperti aksiCancelPo_ (saudara aksi ini, sudah
  // dibatch dari awal): seluruh sheet dibaca sekali, diubah di memory,
  // baru ditulis balik SEKALI lewat setValues()/setBackgrounds() -- hasil
  // akhir identik, cuma jauh lebih sedikit panggilan API.
  const lastCol = Math.max(hist.getLastColumn(), HISTORY_PO_BACKUP_COL);
  const data = hist.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const warnaGrid = hist.getRange(2, 1, lastRow - 1, lastCol).getBackgrounds();
  let updatedCount = 0;

  for (let i = 0; i < data.length; i++) {
    const rowJenis = String(data[i][0] || '').trim().toUpperCase();
    const rowPoNo = String(data[i][1] || '').trim();
    if (rowJenis !== jenis || rowPoNo !== poNo) continue;

    const backupStatus = data[i][HISTORY_PO_BACKUP_COL - 1];
    data[i][HISTORY_PO_STATUS_COL - 1] = backupStatus || '';
    data[i][HISTORY_PO_BACKUP_COL - 1] = '';
    for (let c = 0; c < lastCol; c++) warnaGrid[i][c] = '#ffffff';
    updatedCount++;
  }

  if (updatedCount > 0) {
    hist.getRange(2, 1, data.length, lastCol).setValues(data);
    hist.getRange(2, 1, warnaGrid.length, lastCol).setBackgrounds(warnaGrid);
  }

  return { ok: true, updatedCount: updatedCount };
}

/* ------------------------------------------------------
   AKSI BARU (Oktober 2026 — audit menyeluruh): perbaikiStatusHistoryPOMassalBatch
   — versi database dari perbaikiStatusHistoryPOMassal() (SyncDone.js sisi
   inventory_SANFELO, tombol menu admin "✅ Perbaiki Status HISTORY PO
   (dari SJ)"). Sejak HISTORY_PO_MASUK jadi mirror IMPORTRANGE (read-only)
   di inventory_SANFELO, versi lama fungsi itu (setValues()/
   setBackgrounds() LANGSUNG ke sheet lokal) AKAN ERROR persis saat admin
   mengklik tombol menu ini. Bagian BACA + HITUNG (scan TRANSAKSI/
   DETAIL_TRANSAKSI, bandingkan ke HISTORY_PO_MASUK, susun rencana
   perubahan, tampilkan preview + minta konfirmasi YES/NO) TETAP 100% di
   sisi inventory_SANFELO seperti sebelumnya (semua itu baca-saja, aman
   di atas mirror) -- CUMA langkah TULIS terakhir (sesudah admin klik YES)
   yang sekarang dikirim ke sini.

   body: { updates: [{ rowIdx, statusBaru, warnaBaru }, ...] }
   rowIdx = nomor baris APA ADANYA di mirror lokal inventory_SANFELO, yang
   SELALU sejajar 1:1 dengan baris di database_SANFELO (semua mirror
   IMPORTRANGE di sistem ini mulai dari baris 2 yang sama) -- pola yang
   sama dipakai aksi gateway lain yang terima rowIndex dari client, mis.
   bayarHutangMultiPB.
   ------------------------------------------------------ */
function aksiPerbaikiStatusHistoryPOMassalBatch_(ss, body) {
  const updates = Array.isArray(body.updates) ? body.updates : [];
  if (updates.length === 0) return { ok: true, updatedCount: 0 };

  const hist = ss.getSheetByName(HISTORY_PO_MASUK_SHEET);
  if (!hist) throw new Error('Sheet "' + HISTORY_PO_MASUK_SHEET + '" tidak ditemukan di database.');

  const lastCol = Math.max(hist.getLastColumn(), HISTORY_PO_BACKUP_COL);
  const urut = updates.slice().sort(function (a, b) { return Number(a.rowIdx) - Number(b.rowIdx); });

  let idx = 0;
  while (idx < urut.length) {
    const startRow = Number(urut[idx].rowIdx);
    let endRow = startRow;
    let next = idx + 1;
    while (next < urut.length && Number(urut[next].rowIdx) === endRow + 1) {
      endRow = Number(urut[next].rowIdx);
      next++;
    }

    const numRows = endRow - startRow + 1;
    const statusValues = [];
    const warnaGrid = [];
    for (let r = 0; r < numRows; r++) {
      const u = urut[idx + r];
      statusValues.push([String(u.statusBaru || '')]);
      const rowWarna = [];
      for (let c = 0; c < lastCol; c++) rowWarna.push(String(u.warnaBaru || '#ffffff'));
      warnaGrid.push(rowWarna);
    }

    hist.getRange(startRow, HISTORY_PO_STATUS_COL, numRows, 1).setValues(statusValues);
    hist.getRange(startRow, 1, numRows, lastCol).setBackgrounds(warnaGrid);

    idx = next;
  }

  return { ok: true, updatedCount: urut.length };
}

/* ------------------------------------------------------
   AKSI BARU (Oktober 2026 — audit menyeluruh): perbaikiWarnaStatusHistoryPOBatch
   — versi database dari perbaikiWarnaStatusHistoryPO() (SyncDone.js,
   tool manual Apps-Script-editor-only, TIDAK ada di menu admin). Sama
   alasannya dengan aksi di atas: HANYA mengembalikan warna latar ke
   putih untuk baris yang dikirim, teks STATUS SAMA SEKALI TIDAK disentuh.

   body: { rows: [rowIdx, ...] }
   ------------------------------------------------------ */
function aksiPerbaikiWarnaStatusHistoryPOBatch_(ss, body) {
  const rows = Array.isArray(body.rows)
    ? body.rows.map(Number).filter(function (n) { return n >= 2; })
    : [];
  if (rows.length === 0) return { ok: true, updatedCount: 0 };

  const hist = ss.getSheetByName(HISTORY_PO_MASUK_SHEET);
  if (!hist) throw new Error('Sheet "' + HISTORY_PO_MASUK_SHEET + '" tidak ditemukan di database.');

  const lastCol = hist.getLastColumn();
  const urut = rows.slice().sort(function (a, b) { return a - b; });

  let idx = 0;
  while (idx < urut.length) {
    const startRow = urut[idx];
    let endRow = startRow;
    let next = idx + 1;
    while (next < urut.length && urut[next] === endRow + 1) {
      endRow = urut[next];
      next++;
    }

    const numRows = endRow - startRow + 1;
    const warnaGrid = [];
    for (let r = 0; r < numRows; r++) {
      const rowWarna = [];
      for (let c = 0; c < lastCol; c++) rowWarna.push('#ffffff');
      warnaGrid.push(rowWarna);
    }
    hist.getRange(startRow, 1, numRows, lastCol).setBackgrounds(warnaGrid);

    idx = next;
  }

  return { ok: true, updatedCount: urut.length };
}

/* ------------------------------------------------------
   AKSI: bersihkanDuplikat — versi database dari
   bersihkanDuplikatHistoryPO() / bersihkanDuplikatUntukPO_().
   body: { jenis, poNo }  <- opsional, kalau dikosongin bersihin SEMUA sheet
   ------------------------------------------------------ */
function aksiBersihkanDuplikat_(ss, body) {
  const jenis = body.jenis ? String(body.jenis).trim().toUpperCase() : '';
  const poNo = body.poNo ? String(body.poNo).trim() : '';

  const hist = ss.getSheetByName(HISTORY_PO_MASUK_SHEET);
  if (!hist) throw new Error('Sheet "' + HISTORY_PO_MASUK_SHEET + '" tidak ditemukan.');

  const dihapus = jalankanDedupUntukPO_(hist, jenis, poNo, true);
  return { ok: true, baris_dihapus: dihapus };
}

function jalankanDedupUntukPO_(hist, jenisFilter, poNoFilter, hitungSaja) {
  const lastRow = hist.getLastRow();
  if (lastRow < 2) return 0;

  const data = hist.getRange(2, 1, lastRow - 1, 16).getValues();
  const map = {};

  data.forEach(function (row, i) {
    const jenis = String(row[0] || '').trim().toUpperCase();
    const poNo = String(row[1] || '').trim();
    if (!jenis || !poNo) return;

    if (jenisFilter && jenis !== jenisFilter) return;
    if (poNoFilter && poNo !== poNoFilter) return;

    const line = row[15];
    const ts = row[14];
    const key = jenis + '||' + poNo + '||' + String(line).trim();

    if (!map[key]) map[key] = [];
    map[key].push({ rowIdx: i + 2, timestamp: parseTimestampGateway_(ts) });
  });

  const rowsToDelete = [];
  Object.keys(map).forEach(function (key) {
    const entries = map[key];
    if (entries.length <= 1) return;
    entries.sort(function (a, b) { return b.timestamp - a.timestamp; });
    for (let i = 1; i < entries.length; i++) rowsToDelete.push(entries[i].rowIdx);
  });

  rowsToDelete.sort(function (a, b) { return b - a; });
  rowsToDelete.forEach(function (r) { hist.deleteRow(r); });

  return rowsToDelete.length;
}

function parseTimestampGateway_(value) {
  if (!value) return 0;
  if (value instanceof Date) return value.getTime();
  const str = String(value).trim();
  const m = str.match(/^(\d{2})\/(\d{2})\/(\d{4})\s+(\d{2}):(\d{2}):(\d{2})$/);
  if (!m) return 0;
  const [, dd, mm, yyyy, hh, mi, ss] = m;
  return new Date(yyyy, mm - 1, dd, hh, mi, ss).getTime();
}

/* ------------------------------------------------------
   AKSI: updateStatusDoneSjDelta — versi database dari
   updateStatusDoneSJHistoryPODelta_().
   body: { oldMappingList: [{poNo, line, qty}, ...], newMappingList: [...] }
   ------------------------------------------------------ */
function aksiUpdateStatusDoneSjDelta_(ss, body) {
  const oldMappingList = body.oldMappingList || [];
  const newMappingList = body.newMappingList || [];

  const hist = ss.getSheetByName(HISTORY_PO_MASUK_SHEET);
  if (!hist) throw new Error('Sheet "' + HISTORY_PO_MASUK_SHEET + '" tidak ditemukan.');

  const deltaMap = {};
  function tambahKe(list, tanda) {
    (list || []).forEach(function (m) {
      const key = String(m.poNo).trim() + '||' + String(m.line).trim();
      deltaMap[key] = (deltaMap[key] || 0) + tanda * (Number(m.qty) || 0);
    });
  }
  tambahKe(oldMappingList, -1);
  tambahKe(newMappingList, 1);

  const keys = Object.keys(deltaMap).filter(function (k) { return deltaMap[k] !== 0; });
  if (keys.length === 0) return { ok: true, baris_diubah: 0 };

  const lastRow = hist.getLastRow();
  if (lastRow < 2) return { ok: true, baris_diubah: 0 };

  const lastCol = hist.getLastColumn();
  const data = hist.getRange(2, 1, lastRow - 1, 16).getValues();
  let diubah = 0;

  for (let i = 0; i < data.length; i++) {
    const jenis = String(data[i][0] || '').trim().toUpperCase();
    if (jenis !== 'MASUK') continue;

    const poNo = String(data[i][1] || '').trim();
    const line = String(data[i][15] || '').trim();
    const key = poNo + '||' + line;
    if (!(key in deltaMap) || deltaMap[key] === 0) continue;

    const qtyAsliPO = Number(data[i][10]) || 0;
    const statusSaatIni = String(data[i][11] || '').trim().toUpperCase();
    const qtyDoneSebelumnya = parseDoneQtyGateway_(statusSaatIni, qtyAsliPO);

    let qtyDoneBaru = qtyDoneSebelumnya + deltaMap[key];
    if (qtyDoneBaru < 0) qtyDoneBaru = 0;
    if (qtyDoneBaru > qtyAsliPO) qtyDoneBaru = qtyAsliPO;

    const rowIdx = i + 2;

    if (qtyDoneBaru <= 0) {
      hist.getRange(rowIdx, HISTORY_PO_STATUS_COL).setValue('');
      hist.getRange(rowIdx, 1, 1, lastCol).setBackground('#ffffff');
    } else if (qtyDoneBaru >= qtyAsliPO) {
      hist.getRange(rowIdx, HISTORY_PO_STATUS_COL).setValue('DONE SJ');
      hist.getRange(rowIdx, 1, 1, lastCol).setBackground('#d9ead3');
    } else {
      hist.getRange(rowIdx, HISTORY_PO_STATUS_COL).setValue('DONE SJ ' + qtyDoneBaru);
      hist.getRange(rowIdx, 1, 1, lastCol).setBackground('#fff2cc');
    }
    diubah++;
  }

  return { ok: true, baris_diubah: diubah };
}

function parseDoneQtyGateway_(statusStr, qtyAsli) {
  const s = String(statusStr || '').trim().toUpperCase();
  if (s === 'DONE SJ') return qtyAsli;
  const m = s.match(/^DONE SJ\s+(\d+(\.\d+)?)$/);
  if (m) return Number(m[1]);
  return 0;
}

/* ------------------------------------------------------
   AKSI: syncPoKeluar — versi database dari simpanHistoryPOKeluar_() +
   hapusHistoryPOKeluarRows_() + ambilQtyTerkirimLama_() yang dulu jalan
   lokal di inventory_SANFELO. Sama pola-nya kayak syncPoMasuk: hapus
   baris lama punya PO ini, tulis baris baru -- bedanya di sini nilai
   kolom H (QTY TERKIRIM) yang SUDAH ADA buat tiap NAMA BARANG di PO ini
   dipertahankan (persis logic ambilQtyTerkirimLama_ yang lama), karena
   kolom H itu sumber kebenarannya sekarang di SINI (diisi dari sisi
   inventory lewat aksi updateQtyTerkirimPoKeluar_, bukan dikirim ulang
   tiap kali PO-nya di-resync).

   body: { poNo, tgl, items: [[namaToko, namaBarang, jumlah, keterangan,
           nilaiF], ...] }  <- nilaiF (lookup DATA_BERSIH) sudah dihitung
           di sisi inventory (DATA_BERSIH cuma ada di inventory, bukan
           di database), jadi tinggal dipakai apa adanya di sini.
   ------------------------------------------------------ */
function aksiSyncPoKeluar_(ss, body) {
  const poNo = String(body.poNo || '').trim();
  // ⬅️ FIX (Oktober 2026, sama kasusnya kayak aksiSyncPoMasuk_ di atas):
  // body.tgl dikirim dari sisi inventory sebagai Date, tapi JSON.stringify()
  // otomatis ubah itu jadi teks ISO ("2026-09-30T00:00:00.000Z") pas
  // dikirim lewat gateway -- kalau ditulis apa adanya, kolom B (TGL) di
  // HISTORY_PO_KELUAR jadi teks ISO mentah, bukan tanggal dd/mm/yyyy yang
  // rapi kayak kolom lain. Di-parse ulang di sini pakai
  // parseTanggalDariGateway_() (helper yang sama dipakai TEMPO TOKO/HUTANG/
  // SPAREPART/PEMBELIAN) sebelum dipakai buat nulis baris.
  const tgl = parseTanggalDariGateway_(body.tgl || '');
  const items = body.items;

  if (!poNo) throw new Error('poNo tidak dikirim.');
  if (!Array.isArray(items)) throw new Error('items harus array (boleh kosong).');

  let hist = ss.getSheetByName(HISTORY_PO_KELUAR_SHEET);
  if (!hist) {
    hist = ss.insertSheet(HISTORY_PO_KELUAR_SHEET);
    hist.appendRow(HISTORY_PO_KELUAR_HEADER);
  }

  // 1) Ambil QTY TERKIRIM (kolom H) lama yang sudah ada utk PO ini,
  // dipetakan per NAMA BARANG -- biar gak ke-reset waktu baris lama
  // dihapus & ditulis ulang.
  const mapQtyTerkirim = {};
  const lastRowLama = hist.getLastRow();
  if (lastRowLama >= 2) {
    const dataLama = hist.getRange(2, 1, lastRowLama - 1, HISTORY_PO_KELUAR_QTY_COL).getValues();
    dataLama.forEach(function (r) {
      const rowPoNo = String(r[0] || '').trim();
      const namaBarang = r[3]; // D
      const qtyTerkirim = r[7]; // H
      if (rowPoNo === poNo && namaBarang) {
        mapQtyTerkirim[String(namaBarang).trim().toUpperCase()] = qtyTerkirim;
      }
    });
  }

  // 2) Hapus baris lama utk PO ini
  const lastRowSebelumHapus = hist.getLastRow();
  if (lastRowSebelumHapus >= 2) {
    const dataA = hist.getRange(2, 1, lastRowSebelumHapus - 1, 1).getValues();
    for (let i = dataA.length - 1; i >= 0; i--) {
      if (String(dataA[i][0] || '').trim() === poNo) hist.deleteRow(i + 2);
    }
  }

  // 3) Tulis baris baru
  const timestamp = new Date();
  const rows = [];

  items.forEach(function (item) {
    const namaToko = item[0];
    const namaBarang = item[1];
    const jumlah = item[2];
    const keterangan = item[3];
    const nilaiF = item[4];

    if (!namaBarang) return;

    const key = String(namaBarang).trim().toUpperCase();
    const qtyTerkirim = mapQtyTerkirim.hasOwnProperty(key) ? mapQtyTerkirim[key] : '';

    rows.push([poNo, tgl, namaToko, namaBarang, jumlah, keterangan, timestamp, qtyTerkirim, nilaiF]);
  });

  let startRow = null;
  if (rows.length > 0) {
    startRow = hist.getLastRow() + 1;
    hist.getRange(startRow, 1, rows.length, HISTORY_PO_KELUAR_COLS).setValues(rows);
    hist.getRange(startRow, 2, rows.length, 1).setNumberFormat('dd/mm/yyyy'); // B TGL
  }

  return { ok: true, poNo: poNo, baris_ditulis: rows.length, mulai_dari_baris: startRow };
}

/* ------------------------------------------------------
   AKSI: updateQtyTerkirimPoKeluar — dipanggil tiap kali admin ngetik di
   kolom H (QTY TERKIRIM) lokal di HISTORY_PO_KELUAR sisi inventory.
   Kolom H sengaja TIDAK ikut di-mirror IMPORTRANGE (biar tetap bisa
   diketik manual), jadi satu-satunya cara database ikut punya nilai
   yang sama adalah lewat aksi ini, dikirim satu-satu tiap kali sel-nya
   diedit (lihat handler onEdit di PO_HISTORY.js sisi inventory).

   Dicocokkan lewat NO PO + NAMA BARANG (+ NAMA TOKO kalau dikirim, buat
   disambiguasi kalau ada nama barang yang sama di lebih dari satu toko
   dalam PO yang sama). Normalnya cuma ketemu 1 baris; kalau lebih dari
   1 yang cocok (edge case toko+barang sama persis), semuanya diupdate
   nilai yang sama -- lebih aman daripada nebak salah satu.

   body: { poNo, namaToko, namaBarang, qtyTerkirim }
   ------------------------------------------------------ */
function aksiUpdateQtyTerkirimPoKeluar_(ss, body) {
  const poNo = String(body.poNo || '').trim();
  const namaToko = body.namaToko !== undefined && body.namaToko !== null ? String(body.namaToko).trim() : '';
  const namaBarang = String(body.namaBarang || '').trim();
  const qtyTerkirim = body.qtyTerkirim;

  if (!poNo) throw new Error('poNo tidak dikirim.');
  if (!namaBarang) throw new Error('namaBarang tidak dikirim.');

  const hist = ss.getSheetByName(HISTORY_PO_KELUAR_SHEET);
  if (!hist) throw new Error('Sheet "' + HISTORY_PO_KELUAR_SHEET + '" tidak ditemukan.');

  const lastRow = hist.getLastRow();
  if (lastRow < 2) return { ok: true, updatedCount: 0 };

  const data = hist.getRange(2, 1, lastRow - 1, 4).getValues(); // A,B,C,D
  let updatedCount = 0;

  for (let i = 0; i < data.length; i++) {
    const rowPoNo = String(data[i][0] || '').trim();
    const rowToko = String(data[i][2] || '').trim();
    const rowBarang = String(data[i][3] || '').trim();

    if (rowPoNo !== poNo || rowBarang !== namaBarang) continue;
    if (namaToko && rowToko && rowToko !== namaToko) continue;

    hist.getRange(i + 2, HISTORY_PO_KELUAR_QTY_COL).setValue(qtyTerkirim);
    updatedCount++;
  }

  return { ok: true, updatedCount: updatedCount };
}

/* ------------------------------------------------------
   AKSI: syncNilaiDataBersihBatch — versi database dari
   syncNilaiDataBersihKeHistory_() yang lama. DATA_BERSIH cuma ada di
   inventory_SANFELO (belum dimigrasi), jadi peta nama barang -> nilai
   dihitung di sisi inventory dan dikirim utuh ke sini; server tinggal
   cocokkan ke semua baris HISTORY_PO_KELUAR (lintas PO, sama seperti
   perilaku fungsi lama) dan tulis ke kolom I.

   body: { map: { "NAMA BARANG UPPER": nilai, ... } }
   ------------------------------------------------------ */
function aksiSyncNilaiDataBersihBatch_(ss, body) {
  const map = body.map || {};

  const hist = ss.getSheetByName(HISTORY_PO_KELUAR_SHEET);
  if (!hist) throw new Error('Sheet "' + HISTORY_PO_KELUAR_SHEET + '" tidak ditemukan.');

  const lastRow = hist.getLastRow();
  if (lastRow < 2) return { ok: true, baris_diubah: 0 };

  const numRows = lastRow - 1;
  const namaBarangCol = hist.getRange(2, 4, numRows, 1).getValues(); // D
  const hasilI = [];
  let diubah = 0;

  for (let i = 0; i < numRows; i++) {
    const namaBarang = namaBarangCol[i][0];
    const key = namaBarang ? String(namaBarang).trim().toUpperCase() : '';
    const nilai = map.hasOwnProperty(key) ? map[key] : '';
    hasilI.push([nilai]);
    if (nilai !== '') diubah++;
  }

  hist.getRange(2, 9, numRows, 1).setValues(hasilI); // kolom I

  return { ok: true, baris_diubah: diubah };
}

/* ------------------------------------------------------
   AKSI: syncTempoTokoFull — versi database dari syncTempoToko_() +
   updateStatusPembelianAll_() sisi inventory. Kolom A-E TEMPO TOKO
   (TGL KIRIM, CUSTOMER, NO INV, TEMPO, STATUS PEMBELIAN) DIHITUNG ULANG
   SELURUHNYA tiap kali sync jalan (bukan cuma nambah baris baru kayak
   HISTORY_PO_MASUK/KELUAR), jadi pola aksinya "ganti semua": seluruh isi
   A:E lama dihapus, ditulis ulang dari rows yang dikirim.

   Kolom G:H (ringkasan per customer + dropdown UNLOCKED/LOCKED) SENGAJA
   TIDAK disentuh di sini sama sekali -- itu murni lokal di sisi
   inventory (dipakai juga buat cekTempoTokoBlokir_ yang butuh baca fresh
   tanpa delay), jadi gak pernah ikut dikirim/ditulis ke database.

   body: { rows: [[tglKirim, customer, noInv, tempo, status], ...] }
   tglKirim & tempo bisa datang sebagai Date object ATAU string ISO (hasil
   JSON.stringify Date di sisi inventory, atau apa adanya kalau row-nya
   cuma diteruskan lagi dari hasil bacaSheetGateway_ sebelumnya) atau ""
   kalau kosong -- dikonversi balik ke Date object di sini biar format
   dd/mm/yyyy tetap kepasang benar.
   ------------------------------------------------------ */
function aksiSyncTempoTokoFull_(ss, body) {
  const rows = body.rows;
  if (!Array.isArray(rows)) throw new Error('rows harus array (boleh kosong).');

  // ⬅️ Auto-bikin sheet kalau belum ada (sama kayak syncPoKeluar) -- biar
  // gak perlu bikin tab TEMPO TOKO manual dulu sebelum sync pertama jalan.
  // Kalau sudah ada tab TEMPO TOKO lama hasil duplikasi manual (sebelum
  // migrasi ini), isinya akan DITIMPA oleh sync pertama -- pastikan itu
  // memang sheet kosong/baru sebelum deploy, bukan salinan lama yang mau
  // dipertahankan.
  let sh = ss.getSheetByName(TEMPO_TOKO_SHEET);
  if (!sh) {
    sh = ss.insertSheet(TEMPO_TOKO_SHEET);
    sh.appendRow(['TGL KIRIM', 'CUSTOMER', 'NO INV', 'TEMPO', 'STATUS PEMBELIAN']);
  }

  const lastRow = sh.getLastRow();
  if (lastRow >= 2) {
    sh.getRange(2, 1, lastRow - 1, TEMPO_TOKO_COLS).clearContent();
  }

  if (rows.length === 0) {
    return { ok: true, baris_ditulis: 0 };
  }

  const converted = rows.map(function (r) {
    return [
      parseTanggalDariGateway_(r[0]),
      r[1] || '',
      r[2] || '',
      parseTanggalDariGateway_(r[3]),
      r[4] || ''
    ];
  });

  sh.getRange(2, 1, converted.length, TEMPO_TOKO_COLS).setValues(converted);
  sh.getRange(2, 1, converted.length, 1).setNumberFormat('dd/mm/yyyy');
  sh.getRange(2, 4, converted.length, 1).setNumberFormat('dd/mm/yyyy');

  // ⬅️ BARU: ikut kasih warna baris (hijau UNLOCKED / merah LOCKED / kuning
  // "-N HARI") persis kayak yang dipakai di sisi inventory, biar tampilan
  // di database ikut konsisten -- murni kosmetik, gak dibaca oleh logic
  // manapun (bacaSheetGateway_ cuma baca VALUE, bukan warna).
  const warnaGrid = converted.map(function (r) {
    const warna = warnaUntukStatusGateway_(r[4]);
    return [warna, warna, warna, warna, warna];
  });
  sh.getRange(2, 1, warnaGrid.length, TEMPO_TOKO_COLS).setBackgrounds(warnaGrid);

  return { ok: true, baris_ditulis: converted.length };
}

function parseTanggalDariGateway_(value) {
  if (!value) return '';
  if (value instanceof Date) return value;
  const d = new Date(value);
  return isNaN(d.getTime()) ? '' : d;
}

// Sama persis dengan warnaUntukStatus_() di TempoToko.js sisi inventory --
// UNLOCKED=hijau, LOCKED=merah, selain itu ("-N HARI" masa grace)=kuning.
function warnaUntukStatusGateway_(status) {
  const s = String(status || '').trim().toUpperCase();
  if (s === 'LOCKED') return '#f4cccc';
  if (s === 'UNLOCKED') return '#d9ead3';
  return '#fff2cc';
}

/* ------------------------------------------------------
   AKSI: perpanjangTempoPerCustomer — versi database dari bagian
   handlePerpanjangTempo_() yang nulis kolom D (TEMPO). Nyari semua baris
   CUSTOMER cocok DAN STATUS PEMBELIAN (kolom E) = "LOCKED", tambah TEMPO
   (kolom D) lama + jumlahHari.

   Status PEMBELIAN (kolom E) di sini masih nilai hasil sync terakhir
   (bisa beda momen sama ringkasan G:H lokal di inventory, tapi gak
   masalah -- sisi inventory recompute ulang status+G:H dari TEMPO baru
   lewat updateStatusPembelianAll_() begitu balik dari aksi ini).

   body: { customer, jumlahHari }
   ------------------------------------------------------ */
function aksiPerpanjangTempoPerCustomer_(ss, body) {
  const customer = String(body.customer || '').trim();
  const jumlahHari = Number(body.jumlahHari);

  if (!customer) throw new Error('customer tidak dikirim.');
  if (!jumlahHari || isNaN(jumlahHari) || jumlahHari <= 0) {
    throw new Error('jumlahHari harus angka lebih dari 0.');
  }

  const sh = ss.getSheetByName(TEMPO_TOKO_SHEET);
  if (!sh) throw new Error('Sheet "' + TEMPO_TOKO_SHEET + '" tidak ditemukan di database.');

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { ok: true, updatedCount: 0, rows: [] };

  const data = sh.getRange(2, 1, lastRow - 1, TEMPO_TOKO_COLS).getValues(); // A..E
  const updatedIdx = [];

  for (let i = 0; i < data.length; i++) {
    const rowCustomer = String(data[i][1] || '').trim();
    if (rowCustomer !== customer) continue;

    const status = String(data[i][4] || '').trim().toUpperCase();
    if (status !== 'LOCKED') continue;

    const tempoLama = data[i][3];
    if (!tempoLama) continue;

    const tempoBaru = new Date(tempoLama);
    tempoBaru.setDate(tempoBaru.getDate() + jumlahHari);
    data[i][3] = tempoBaru;

    updatedIdx.push(i);
  }

  if (updatedIdx.length > 0) {
    sh.getRange(2, 1, data.length, TEMPO_TOKO_COLS).setValues(data);
    sh.getRange(2, 4, data.length, 1).setNumberFormat('dd/mm/yyyy');
  }

  return {
    ok: true,
    updatedCount: updatedIdx.length,
    rows: updatedIdx.map(function (i) {
      return { noInv: String(data[i][2] || '').trim(), tempoBaru: data[i][3].toISOString() };
    })
  };
}

/* ------------------------------------------------------
   AKSI: appendEkspedisi — versi database dari bagian simpanBiayaTambahan()
   yang nulis baris baru ke sheet EKSPEDISI. Pure append (gak ada
   dedup/delete kayak HISTORY_PO_*), makanya action-nya ringan -- cuma
   nambah baris + format angka QTY/HARGA. Baris PENJUALAN (kolom H/I +
   SISA BAYAR/STATUS) yang diupdate di fungsi yang sama TETAP lokal di
   inventory (PENJUALAN belum dimigrasi), jadi TIDAK ikut lewat sini.

   body: { rows: [[13 kolom], ...] } -- normalnya cuma 1 baris per
   submit popup, tapi dibiarkan bentuk array biar konsisten & fleksibel.
   ------------------------------------------------------ */
function aksiAppendEkspedisi_(ss, body) {
  const rows = body.rows;
  if (!Array.isArray(rows) || rows.length === 0) throw new Error('rows kosong / bukan array.');

  const sh = ss.getSheetByName(EKSPEDISI_SHEET);
  if (!sh) throw new Error('Sheet "' + EKSPEDISI_SHEET + '" tidak ditemukan di database.');

  const startRow = sh.getLastRow() + 1;
  sh.getRange(startRow, 1, rows.length, EKSPEDISI_COLS).setValues(rows);
  sh.getRange(startRow, EKSPEDISI_QTY_COL, rows.length, 1).setNumberFormat('#,##0');
  sh.getRange(startRow, EKSPEDISI_HARGA_COL, rows.length, 1).setNumberFormat('#,##0');

  return { ok: true, baris_ditulis: rows.length, mulai_dari_baris: startRow };
}

/* ------------------------------------------------------
   HELPER HUTANG — duplikat sengaja dari fungsi yang sama di sisi
   inventory_SANFELO (Pembayaran.gs / PembayaranHutang.gs /
   InvoiceGroupingPopup.gs). Dua project Apps Script yang terpisah
   TIDAK bisa saling manggil fungsi, jadi harus di-duplikat di sini
   -- tapi logic & hasilnya harus tetap identik.
   ------------------------------------------------------ */
function formatRupiahGatewayHutang_(number) {
  return 'Rp' + Number(number || 0).toLocaleString('id-ID');
}

function parseRupiahTextGatewayHutang_(text) {
  const digits = String(text).replace(/[^0-9]/g, '');
  return Number(digits) || 0;
}

// Ngerti dua format riwayat cicilan: BARU "POST-NNN - tgl - Rp - metode"
// (dipakai cicilan PENJUALAN, bukan HUTANG) & LAMA "tgl - Rp - metode"
// (yang ini yang dipakai HUTANG) — sama persis kayak parseRiwayatCell_
// di Pembayaran.gs sisi inventory.
const HUTANG_NOPOST_PREFIX_ = 'POST-';
function parseRiwayatCellGatewayHutang_(text) {
  const str = String(text).trim();
  if (!str) return null;
  const parts = str.split(' - ');
  if (parts.length < 2) return null;

  let idx = 0;
  if (parts[0].trim().indexOf(HUTANG_NOPOST_PREFIX_) === 0) idx = 1;

  const jumlah = parseRupiahTextGatewayHutang_(parts[idx + 1]);
  return { jumlah: jumlah };
}

function sisaBayarNumericGatewayHutang_(raw) {
  if (typeof raw === 'number') return raw;
  const str = String(raw || '').trim();
  if (str.toUpperCase().indexOf('LEBIH') === 0) return 0;
  return Number(str) || 0;
}

function normalisasiBulanGatewayHutang_(val) {
  if (val instanceof Date) {
    const y = val.getFullYear();
    const m = String(val.getMonth() + 1).padStart(2, '0');
    return y + '-' + m;
  }
  let str = String(val || '').replace(/ /g, ' ').trim();
  if (!str) return '';
  const match = str.match(/^(\d{4})[-\/](\d{1,2})$/);
  if (match) {
    const y = match[1];
    const m = match[2].padStart(2, '0');
    return y + '-' + m;
  }
  return str;
}

function normalisasiSupplierGatewayHutang_(val) {
  return String(val || '').replace(/ /g, ' ').trim().toUpperCase();
}

function parseTanggalInputGatewayHutang_(str) {
  if (!str) return null;
  const parts = String(str).split('-'); // ["2026","07","31"]
  if (parts.length !== 3) return new Date(str);
  return new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
}

// Cari baris terakhir HUTANG yang BENERAN ada isi di kolom B (NO PB),
// bukan pakai hutang.getLastRow() mentah (bisa ke-fool formatting kosong).
function cariLastRowHutangGatewayBerisi_(hutang) {
  const lastRow = hutang.getLastRow();
  if (lastRow < 2) return 1;
  const colB = hutang.getRange(2, HUTANG_COL_PB, lastRow - 1, 1).getValues();
  let last = 1;
  colB.forEach(function (row, i) {
    if (String(row[0] || '').trim() !== '') last = i + 2;
  });
  return last;
}

// Hitung ulang SISA BAYAR 1 baris, murni dari TOTAL terbaru dikurangi
// TOTAL YANG SUDAH DIBAYAR (di-parse ulang dari riwayat asli kolom K dst
// -- bukan ngutak-atik SISA BAYAR lama yang bisa udah gak akurat).
function recomputeSisaBayarGatewayHutang_(hutang, rowIdx, totalBaru) {
  const lastCol = hutang.getLastColumn();
  let totalSudahDibayar = 0;

  if (lastCol >= HUTANG_RIWAYAT_START_COL) {
    const riwayatCells = hutang
      .getRange(rowIdx, HUTANG_RIWAYAT_START_COL, 1, lastCol - HUTANG_RIWAYAT_START_COL + 1)
      .getValues()[0];
    riwayatCells.forEach(function (cell) {
      const parsed = parseRiwayatCellGatewayHutang_(cell);
      if (parsed) totalSudahDibayar += parsed.jumlah;
    });
  }

  const sisaBaru = totalBaru - totalSudahDibayar;
  const sisaRange = hutang.getRange(rowIdx, HUTANG_COL_SISABAYAR);

  if (sisaBaru < 0) {
    sisaRange.setValue('LEBIH ' + formatRupiahGatewayHutang_(Math.abs(sisaBaru)));
  } else {
    sisaRange.setValue(sisaBaru).setNumberFormat('#,##0');
  }

  return sisaBaru;
}

// Refresh STATUS (kolom I) + warna baris A:J utk rowIndexes yang dikasih.
// Port persis dari updateStatusJatuhTempo_ (InvoiceGroupingPopup.gs sisi
// inventory, versi SEBELUM migrasi) -- logic & hasil akhirnya identik,
// cuma sekarang jalan terhadap sheet HUTANG yang asli.
function updateStatusJatuhTempoGatewayHutang_(hutang, rowIndexes) {
  if (!rowIndexes || rowIndexes.length === 0) return;

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const lastCol = HUTANG_COL_SISABAYAR; // kolom J, warna baris cuma A:J

  const sortedRows = rowIndexes.slice().sort(function (a, b) { return a - b; });

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
    const dataGroup = hutang.getRange(startRow, 1, numRows, lastCol).getValues();
    const statusValues = dataGroup.map(function (row) { return [row[HUTANG_COL_STATUS - 1]]; });
    const warnaGrid = hutang.getRange(startRow, 1, numRows, lastCol).getBackgrounds();
    const fontGrid = hutang.getRange(startRow, 1, numRows, lastCol).getFontColors();
    let adaPerubahan = false;

    for (let r = 0; r < numRows; r++) {
      const jatuhTempoVal = dataGroup[r][HUTANG_COL_JATUHTEMPO - 1];
      const sisaBayar = Number(dataGroup[r][HUTANG_COL_SISABAYAR - 1]) || 0;

      if (sisaBayar <= 0) {
        statusValues[r][0] = 'LUNAS';
        for (let c = 0; c < lastCol; c++) { warnaGrid[r][c] = HUTANG_STATUS_WARNA.LUNAS.bg; fontGrid[r][c] = HUTANG_STATUS_WARNA.LUNAS.font; }
        adaPerubahan = true;
        continue;
      }

      let jt = null;
      if (jatuhTempoVal instanceof Date) {
        jt = new Date(jatuhTempoVal);
      } else if (jatuhTempoVal) {
        jt = parseTanggalInputGatewayHutang_(String(jatuhTempoVal));
      }
      if (!jt || isNaN(jt.getTime())) continue; // belum ada jatuh tempo -> biarin kosong

      jt.setHours(0, 0, 0, 0);
      const selisihHari = Math.round((jt - today) / (1000 * 60 * 60 * 24));

      if (selisihHari > 0) {
        statusValues[r][0] = '-' + selisihHari + ' HARI';
        for (let c = 0; c < lastCol; c++) { warnaGrid[r][c] = HUTANG_STATUS_WARNA.MENDEKATI.bg; fontGrid[r][c] = HUTANG_STATUS_WARNA.MENDEKATI.font; }
      } else {
        statusValues[r][0] = 'JATUH TEMPO';
        for (let c = 0; c < lastCol; c++) { warnaGrid[r][c] = HUTANG_STATUS_WARNA.LEWAT.bg; fontGrid[r][c] = HUTANG_STATUS_WARNA.LEWAT.font; }
      }
      adaPerubahan = true;
    }

    if (adaPerubahan) {
      hutang.getRange(startRow, HUTANG_COL_STATUS, numRows, 1).setValues(statusValues);
      hutang.getRange(startRow, 1, numRows, lastCol).setBackgrounds(warnaGrid);
      hutang.getRange(startRow, 1, numRows, lastCol).setFontColors(fontGrid);
    }

    idx = next;
  }
}

/* ------------------------------------------------------
   AKSI: syncHutangDariPembelian — versi database dari
   syncHutangFromPembelian_() (PembelianHutangSync.gs lama).
   PEMBELIAN TETAP LOKAL (belum dimigrasi) — sisi inventory yang scan
   & grouping-nya, di sini cuma terima hasil groupingnya lalu
   insert/update ke HUTANG asli + recompute SISA BAYAR & STATUS.

   body: { groups: [{ noPB, noSJ, supplier, tglBeli, total }, ...] }
   tglBeli boleh Date/ISO-string/teks apa adanya, di-parse pakai
   parseTanggalDariGateway_ yang sudah ada (dipakai bareng TEMPO TOKO).
   ------------------------------------------------------ */
function aksiSyncHutangDariPembelian_(ss, body) {
  const groups = body.groups;
  if (!Array.isArray(groups)) throw new Error('groups harus array (boleh kosong).');

  const hutang = ss.getSheetByName(HUTANG_SHEET);
  if (!hutang) throw new Error('Sheet "' + HUTANG_SHEET + '" tidak ditemukan di database.');

  if (groups.length === 0) {
    return { ok: true, insertedCount: 0, updatedCount: 0, pbJadiLunasKarenaKoreksi: [] };
  }

  const lastRowHutangBerisi = cariLastRowHutangGatewayBerisi_(hutang);

  const existingMap = {};
  if (lastRowHutangBerisi >= 2) {
    const hutangData = hutang.getRange(2, HUTANG_COL_PB, lastRowHutangBerisi - 1, 1).getValues();
    hutangData.forEach(function (row, i) {
      const noPB = String(row[0] || '').trim();
      if (noPB) existingMap[noPB] = i + 2;
    });
  }

  let nextInsertRow = lastRowHutangBerisi + 1;
  const rowIndexesUpdated = [];
  const pbJadiLunasKarenaKoreksi = [];
  let insertedCount = 0;
  let updatedCount = 0;

  groups.forEach(function (g) {
    const noPB = String(g.noPB || '').trim();
    if (!noPB) return;
    const total = Number(g.total) || 0;
    const tglBeli = parseTanggalDariGateway_(g.tglBeli);

    if (existingMap[noPB]) {
      const rowIdx = existingMap[noPB];
      hutang.getRange(rowIdx, HUTANG_COL_PB, 1, 1).setValue(noPB);
      hutang.getRange(rowIdx, HUTANG_COL_NOSJ, 1, 3).setValues([[g.noSJ || '', g.supplier || '', tglBeli]]);
      // ⬅️ FIX (Oktober 2026 — audit menyeluruh): kolom TGL BELI (F) gak
      // pernah dipaksa format tanggal setelah ditulis -- sama persis
      // celah yang bikin HISTORY_PO_MASUK/HISTORY_PO_KELUAR sempat
      // nampilin teks ISO ("2026-09-30T00:00:...") bukan dd/mm/yyyy.
      // parseTanggalDariGateway_ di atas sudah mastiin nilainya Date
      // beneran, tapi kalau kolomnya KEBETULAN masih berformat "Automatic"/
      // teks dari sebelumnya, tampilannya tetap salah walau value-nya
      // udah benar -- makanya format dipaksa di sini juga, gak cukup cuma
      // parse ulang.
      hutang.getRange(rowIdx, HUTANG_COL_NOSJ + 2, 1, 1).setNumberFormat('dd/mm/yyyy');
      hutang.getRange(rowIdx, HUTANG_COL_TOTAL, 1, 1).setValue(total);

      const sisaBaru = recomputeSisaBayarGatewayHutang_(hutang, rowIdx, total);
      rowIndexesUpdated.push(rowIdx);
      updatedCount++;
      if (sisaBaru <= 0) pbJadiLunasKarenaKoreksi.push(noPB);
    } else {
      hutang.getRange(nextInsertRow, 1, 1, 10).setValues([[
        '', noPB, '', g.noSJ || '', g.supplier || '', tglBeli, '', total, '', total
      ]]);
      // ⬅️ FIX (Oktober 2026 — lihat komentar di atas): kolom F (TGL BELI, kolom ke-6).
      hutang.getRange(nextInsertRow, 6, 1, 1).setNumberFormat('dd/mm/yyyy');
      rowIndexesUpdated.push(nextInsertRow);
      insertedCount++;
      nextInsertRow++;
    }
  });

  try {
    updateStatusJatuhTempoGatewayHutang_(hutang, rowIndexesUpdated);
  } catch (err) {
    // non-blocking, sama kayak versi lokal lama
  }

  const pbUnik = pbJadiLunasKarenaKoreksi.filter(function (pb, i) { return pbJadiLunasKarenaKoreksi.indexOf(pb) === i; });

  return { ok: true, insertedCount: insertedCount, updatedCount: updatedCount, pbJadiLunasKarenaKoreksi: pbUnik };
}

/* ------------------------------------------------------
   AKSI: bayarHutangMultiPB — versi database dari
   prosesPembayaranMultiPB() (PembayaranHutang.gs lama). Lock lokal
   yang dulu dipakai di inventory TIDAK PERLU dibawa ke sini —
   prosesRequestGateway_ di atas SUDAH mengunci SEMUA aksi gateway
   pakai LockService sebelum manggil handler manapun (termasuk ini),
   jadi proteksi race condition antar-admin/tab tetap terjaga, cuma
   sekarang titiknya di sini (terpusat), bukan lokal per-device.

   body: { bulan, supplier, tglBayarText ("dd/MM/yyyy"), metodeBayar,
           payments: [{ rowIndex, pb, jumlahBayar }, ...] }
   ------------------------------------------------------ */
function aksiBayarHutangMultiPB_(ss, body) {
  const bulan = normalisasiBulanGatewayHutang_(body.bulan);
  const supplier = String(body.supplier || '').trim();
  const metodeBayar = String(body.metodeBayar || '').trim().toUpperCase();
  const tglBayarText = String(body.tglBayarText || '').trim();
  const payments = Array.isArray(body.payments) ? body.payments : [];

  if (!bulan) throw new Error('Bulan tidak boleh kosong.');
  if (!supplier) throw new Error('Supplier tidak boleh kosong.');
  if (!tglBayarText) throw new Error('Tanggal bayar belum diisi / tidak valid.');
  if (metodeBayar !== 'TUNAI' && metodeBayar !== 'TRANSFER' && metodeBayar !== 'ADJUSTMENT') {
    throw new Error('Metode pembayaran harus TUNAI, TRANSFER, atau ADJUSTMENT.');
  }
  if (payments.length === 0) throw new Error('Centang minimal 1 PB untuk dibayar.');

  const hutang = ss.getSheetByName(HUTANG_SHEET);
  if (!hutang) throw new Error('Sheet "' + HUTANG_SHEET + '" tidak ditemukan di database.');

  const parsed = [];
  for (let idx = 0; idx < payments.length; idx++) {
    const p = payments[idx];
    const rowIndex = Number(p.rowIndex);
    const pbExpected = String(p.pb || '').trim();
    const jumlahBayar = Number(p.jumlahBayar);

    if (!rowIndex || rowIndex < 2) throw new Error('Baris PB tidak valid.');
    if (!jumlahBayar || jumlahBayar <= 0) {
      throw new Error('Jumlah bayar untuk PB ' + (pbExpected || rowIndex) + ' tidak valid.');
    }

    // re-validasi baris masih cocok bulan+supplier+PB (jaga2 kalau sheet sempat berubah)
    const rowBulan = normalisasiBulanGatewayHutang_(hutang.getRange(rowIndex, HUTANG_COL_BULAN).getValue());
    const rowSupplier = normalisasiSupplierGatewayHutang_(hutang.getRange(rowIndex, HUTANG_COL_SUPPLIER).getValue());
    const pbActual = String(hutang.getRange(rowIndex, HUTANG_COL_PB).getValue() || '').trim();

    if (rowBulan !== bulan || rowSupplier !== normalisasiSupplierGatewayHutang_(supplier)) {
      throw new Error('Data hutang ini sudah berubah di sheet. Cari ulang bulan/supplier-nya sebelum bayar lagi.');
    }
    if (pbExpected && pbActual !== pbExpected) {
      throw new Error('Data PB ' + pbExpected + ' sudah berubah di sheet. Cari ulang datanya sebelum bayar lagi.');
    }

    parsed.push({ rowIndex: rowIndex, pb: pbActual, jumlahBayar: jumlahBayar });
  }

  // --- proses tiap PB SENDIRI-SENDIRI, tidak FIFO ke baris lain ---
  const rowIndexesUpdated = [];
  const pbBaruLunas = [];

  parsed.forEach(function (p) {
    const sisaLamaRaw = hutang.getRange(p.rowIndex, HUTANG_COL_SISABAYAR).getValue();
    const sisaLama = sisaBayarNumericGatewayHutang_(sisaLamaRaw);

    let sisaBaru;
    let jadiLunas = false;
    if (p.jumlahBayar >= sisaLama) {
      const lebih = p.jumlahBayar - sisaLama;
      sisaBaru = lebih > 0 ? ('LEBIH ' + formatRupiahGatewayHutang_(lebih)) : 0;
      jadiLunas = true;
    } else {
      sisaBaru = sisaLama - p.jumlahBayar;
    }

    const targetRange = hutang.getRange(p.rowIndex, HUTANG_COL_SISABAYAR);
    if (typeof sisaBaru === 'number') {
      targetRange.setValue(sisaBaru).setNumberFormat('#,##0');
    } else {
      targetRange.setValue(sisaBaru);
    }

    // riwayat DI BARIS PB INI SENDIRI, mulai kolom K, kolom kosong pertama
    let targetCol = HUTANG_RIWAYAT_START_COL;
    while (hutang.getRange(p.rowIndex, targetCol).getValue()) targetCol++;
    const cellText = tglBayarText + ' - ' + formatRupiahGatewayHutang_(p.jumlahBayar) + ' - ' + metodeBayar;
    hutang.getRange(p.rowIndex, targetCol).setValue(cellText);

    rowIndexesUpdated.push(p.rowIndex);
    if (jadiLunas && p.pb) pbBaruLunas.push(p.pb);
  });

  updateStatusJatuhTempoGatewayHutang_(hutang, rowIndexesUpdated);

  const pbUnik = pbBaruLunas.filter(function (pb, i) { return pbBaruLunas.indexOf(pb) === i; });

  return { ok: true, jumlahPBDibayar: parsed.length, pbBaruLunas: pbUnik };
}

/* ------------------------------------------------------
   AKSI: groupingInvoiceHutang — versi database dari
   simpanGroupingInvoice() (InvoiceGroupingPopup.gs lama).

   body: { rows: [rowIndex, ...], noInv, jatuhTempo, transaksiBulan }
   jatuhTempo disimpan APA ADANYA (biasanya teks "yyyy-MM-dd" dari
   <input type=date>), SAMA PERSIS kayak versi lokal lama — jangan
   di-parse jadi Date di sini, karena updateStatusJatuhTempoGatewayHutang_
   sudah tahu cara baca teks itu juga (lihat parseTanggalInputGatewayHutang_).
   ------------------------------------------------------ */
function aksiGroupingInvoiceHutang_(ss, body) {
  const hutang = ss.getSheetByName(HUTANG_SHEET);
  if (!hutang) throw new Error('Sheet "' + HUTANG_SHEET + '" tidak ditemukan di database.');

  const noInv = String(body.noInv || '').trim();
  const jatuhTempo = body.jatuhTempo;
  const transaksiBulan = String(body.transaksiBulan || '').trim();
  const rows = Array.isArray(body.rows) ? body.rows : [];

  if (!noInv) throw new Error('NO INV wajib diisi.');
  if (!rows.length) throw new Error('Pilih minimal 1 NO SJ.');

  const lastRow = hutang.getLastRow();
  if (lastRow >= 2) {
    const existingInv = hutang.getRange(2, HUTANG_COL_NOINV, lastRow - 1, 1).getValues();
    const dupe = existingInv.some(function (r) {
      return String(r[0]).trim().toUpperCase() === noInv.toUpperCase();
    });
    if (dupe) throw new Error('NO INV "' + noInv + '" sudah dipakai baris lain.');
  }

  rows.forEach(function (rowIndex) {
    hutang.getRange(rowIndex, HUTANG_COL_BULAN).setValue(transaksiBulan);
    hutang.getRange(rowIndex, HUTANG_COL_NOINV).setValue(noInv);
    hutang.getRange(rowIndex, HUTANG_COL_JATUHTEMPO).setValue(jatuhTempo);
  });

  updateStatusJatuhTempoGatewayHutang_(hutang, rows);

  return { ok: true, count: rows.length };
}

/* ------------------------------------------------------
   TRIGGER HARIAN untuk HUTANG — dipasang MANUAL di sisi
   database_SANFELO (BUKAN inventory), karena sheet HUTANG asli
   sekarang ada di sini:
     Apps Script database_SANFELO -> ikon jam (Triggers) di kiri ->
     Add Trigger -> Choose function: refreshStatusJatuhTempoHutangHarian,
     Select event source: Time-driven, Day timer, jam berapa saja yang
     nyaman (mis. 01:00-02:00) -> Save.
   Port persis dari refreshSemuaStatusJatuhTempo() lama (dulu di
   InvoiceGroupingPopup.gs sisi inventory) — logic identik, cuma
   sekarang jalan langsung di sheet HUTANG asli, jadi TIDAK PERLU
   lewat gateway/HTTP sama sekali buat trigger harian ini.
   ------------------------------------------------------ */
function refreshStatusJatuhTempoHutangHarian() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hutang = ss.getSheetByName(HUTANG_SHEET);
  if (!hutang) throw new Error('Sheet "' + HUTANG_SHEET + '" tidak ditemukan.');

  const lastRow = hutang.getLastRow();
  if (lastRow < 2) return;

  const rowIndexes = [];
  for (let r = 2; r <= lastRow; r++) rowIndexes.push(r);

  updateStatusJatuhTempoGatewayHutang_(hutang, rowIndexes);
}

/* ------------------------------------------------------
   AKSI: syncSparepart — versi database dari bagian tulis
   finalizeSparepartSave_() (SPAREPART_SAVE.gs lama). Delete-by-NO SJ
   lalu insert, sama pola kayak syncPoKeluar/syncPoMasuk. Baris final
   (termasuk STATUS ORDER yang tepat buat mode edit/baru) sudah
   dihitung di sisi inventory — gateway cuma eksekusi tulisnya.

   body: { noSJ, rows: [[19 kolom], ...] }
   Kolom D (TANGGAL) & E (TEMPO) boleh dikirim sebagai Date/ISO-string/
   teks apa adanya, di-parse pakai parseTanggalDariGateway_ yang sudah
   ada (dipakai bareng TEMPO TOKO/HUTANG).
   ------------------------------------------------------ */
function aksiSyncSparepart_(ss, body) {
  const noSJ = String(body.noSJ || '').trim();
  const rows = body.rows;
  if (!noSJ) throw new Error('noSJ tidak dikirim.');
  if (!Array.isArray(rows)) throw new Error('rows harus array (boleh kosong).');

  const sh = ss.getSheetByName(SPAREPART_SHEET);
  if (!sh) throw new Error('Sheet "' + SPAREPART_SHEET + '" tidak ditemukan di database.');

  // 1) Hapus baris lama utk NO SJ ini (no-op kalau memang NO SJ baru)
  const lastRowSebelumHapus = sh.getLastRow();
  let deletedCount = 0;
  if (lastRowSebelumHapus >= 2) {
    const dataA = sh.getRange(2, SPAREPART_COL_NOSJ, lastRowSebelumHapus - 1, 1).getValues();
    for (let i = dataA.length - 1; i >= 0; i--) {
      if (String(dataA[i][0] || '').trim() === noSJ) {
        sh.deleteRow(i + 2);
        deletedCount++;
      }
    }
  }

  if (rows.length === 0) {
    return { ok: true, deletedCount: deletedCount, insertedCount: 0, mulai_dari_baris: null };
  }

  // 2) Parse ulang kolom tanggal (index 3 = D, index 4 = E, 0-based)
  const converted = rows.map(function (r) {
    const copy = r.slice();
    copy[SPAREPART_COL_TANGGAL - 1] = parseTanggalDariGateway_(r[SPAREPART_COL_TANGGAL - 1]);
    copy[SPAREPART_COL_TEMPO - 1] = parseTanggalDariGateway_(r[SPAREPART_COL_TEMPO - 1]);
    return copy;
  });

  // 3) Tulis baris baru
  const startRow = sh.getLastRow() + 1;
  sh.getRange(startRow, 1, converted.length, SPAREPART_COLS).setValues(converted);
  sh.getRange(startRow, SPAREPART_COL_TANGGAL, converted.length, 1).setNumberFormat('dd/mm/yyyy');
  sh.getRange(startRow, SPAREPART_COL_TEMPO, converted.length, 1).setNumberFormat('dd/mm/yyyy');

  return { ok: true, deletedCount: deletedCount, insertedCount: converted.length, mulai_dari_baris: startRow };
}

/* ------------------------------------------------------
   AKSI: tandaiDoneSparepart — versi database dari
   tandaiDoneSparepartByNoInv_() (Pembayaran.gs lama), dipanggil
   otomatis begitu sebuah invoice PENJUALAN jadi LUNAS.

   ⚠️ FIX SEKALIAN (ditemukan pas migrasi): versi lokal lama nulis
   STATUS ORDER ke kolom 18 (R), padahal kolom 18 itu LABA (lihat
   struktur kolom di komentar atas) — STATUS ORDER yang benar itu
   kolom 19 (S). Jadi versi lama diam-diam NIMPA angka LABA dengan
   teks "DONE" tiap kali invoice sparepart lunas. Di sini sudah
   ditulis ke kolom yang BENAR (S = SPAREPART_COL_STATUS = 19).

   body: { noInvoice }
   ------------------------------------------------------ */
function aksiTandaiDoneSparepart_(ss, body) {
  const noInvTarget = String(body.noInvoice || '').trim();
  if (!noInvTarget) return { ok: true, updatedCount: 0 };

  const sh = ss.getSheetByName(SPAREPART_SHEET);
  if (!sh) throw new Error('Sheet "' + SPAREPART_SHEET + '" tidak ditemukan di database.');

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { ok: true, updatedCount: 0 };

  const lastCol = Math.max(sh.getLastColumn(), SPAREPART_COL_STATUS);
  const dataNoInv = sh.getRange(2, SPAREPART_COL_NOINV, lastRow - 1, 1).getValues();

  const matchedRows = [];
  for (let i = 0; i < dataNoInv.length; i++) {
    const noInv = String(dataNoInv[i][0] || '').trim();
    if (noInv === noInvTarget) matchedRows.push(i + 2);
  }

  const warnaLunas = typeof PENJUALAN_WARNA_LUNAS !== 'undefined' ? PENJUALAN_WARNA_LUNAS : '#d9ead3';

  let idx = 0;
  while (idx < matchedRows.length) {
    const startRow = matchedRows[idx];
    let endRow = startRow;
    let next = idx + 1;
    while (next < matchedRows.length && matchedRows[next] === endRow + 1) {
      endRow = matchedRows[next];
      next++;
    }

    const numRows = endRow - startRow + 1;
    const statusValues = [];
    const warnaGrid = [];
    for (let r = 0; r < numRows; r++) {
      statusValues.push(['DONE']);
      const rowWarna = [];
      for (let c = 0; c < lastCol; c++) rowWarna.push(warnaLunas);
      warnaGrid.push(rowWarna);
    }

    sh.getRange(startRow, SPAREPART_COL_STATUS, numRows, 1).setValues(statusValues);
    sh.getRange(startRow, 1, numRows, lastCol).setBackgrounds(warnaGrid);

    idx = next;
  }

  return { ok: true, updatedCount: matchedRows.length };
}

/* ------------------------------------------------------
   AKSI: syncPembelian — versi database dari bagian "TULIS KE
   PEMBELIAN" di simpanTandaTerimaKePembelian() (TANDA_TERIMA_BRG.gs
   lama), dipakai juga oleh hapusBarisPembelian_() buat delete-only
   (lihat 3 skenario di komentar struktur kolom PEMBELIAN di atas).

   body: { noSJ, supplier, rows: [[15 kolom], ...] }
   - noSJ+supplier kosong/gak match -> skip hapus, langsung insert.
   - rows kosong -> cuma hapus, gak ada yang diinsert (dipakai flow
     "Edit": hapus baris lama, user edit manual, simpan ulang nanti).
   Kolom E (TGL BELI) boleh dikirim Date/ISO-string/teks apa adanya,
   di-parse pakai parseTanggalDariGateway_ yang sudah ada.
   ------------------------------------------------------ */
function aksiSyncPembelian_(ss, body) {
  const noSJ = String(body.noSJ || '').trim();
  const supplier = String(body.supplier || '').trim().toUpperCase();
  const rows = body.rows;
  if (!Array.isArray(rows)) throw new Error('rows harus array (boleh kosong).');

  const sh = ss.getSheetByName(PEMBELIAN_SHEET);
  if (!sh) throw new Error('Sheet "' + PEMBELIAN_SHEET + '" tidak ditemukan di database.');

  // 1) Hapus baris lama yang match NO SJ + SUPPLIER (case-insensitive,
  //    sama persis kriteria cariBarisPembelianByNoSJ_ versi lokal lama)
  //    -- no-op kalau noSJ kosong atau memang kombinasi baru.
  let deletedCount = 0;
  if (noSJ) {
    const lastRowSebelumHapus = sh.getLastRow();
    if (lastRowSebelumHapus >= 2) {
      const dataCD = sh.getRange(2, PEMBELIAN_COL_NOSJ, lastRowSebelumHapus - 1, 2).getValues(); // C,D
      for (let i = dataCD.length - 1; i >= 0; i--) {
        const rowNoSJ = String(dataCD[i][0] || '').trim();
        const rowSupplier = String(dataCD[i][1] || '').trim().toUpperCase();
        if (rowNoSJ === noSJ && rowSupplier === supplier) {
          sh.deleteRow(i + 2);
          deletedCount++;
        }
      }
    }
  }

  if (rows.length === 0) {
    return { ok: true, deletedCount: deletedCount, insertedCount: 0, mulai_dari_baris: null };
  }

  // 2) Parse ulang kolom tanggal (index 4 = E, 0-based)
  const converted = rows.map(function (r) {
    const copy = r.slice();
    copy[PEMBELIAN_COL_TGLBELI - 1] = parseTanggalDariGateway_(r[PEMBELIAN_COL_TGLBELI - 1]);
    return copy;
  });

  // 3) Tulis baris baru
  const startRow = sh.getLastRow() + 1;
  sh.getRange(startRow, 1, converted.length, PEMBELIAN_COLS).setValues(converted);
  sh.getRange(startRow, PEMBELIAN_COL_TGLBELI, converted.length, 1).setNumberFormat('dd/mm/yyyy');

  return { ok: true, deletedCount: deletedCount, insertedCount: converted.length, mulai_dari_baris: startRow };
}

/* ------------------------------------------------------
   AKSI: tandaiLunasPembelian — versi database dari
   tandaiLunasPembelianByPB_() (PembayaranHutang.gs lama), dipanggil
   otomatis dari prosesPembayaranMultiPB() maupun syncHutangFromPembelian_()
   begitu sebuah NO PB nyampe SISA BAYAR = 0.

   body: { pb }
   ------------------------------------------------------ */
function aksiTandaiLunasPembelian_(ss, body) {
  const pbTarget = String(body.pb || '').trim();
  if (!pbTarget) return { ok: true, updatedCount: 0 };

  const sh = ss.getSheetByName(PEMBELIAN_SHEET);
  if (!sh) throw new Error('Sheet "' + PEMBELIAN_SHEET + '" tidak ditemukan di database.');

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { ok: true, updatedCount: 0 };

  const lastCol = Math.max(sh.getLastColumn(), PEMBELIAN_COL_STATUS_LUNAS);
  const dataPB = sh.getRange(2, PEMBELIAN_COL_PB, lastRow - 1, 1).getValues();

  const warnaLunas = typeof PENJUALAN_WARNA_LUNAS !== 'undefined' ? PENJUALAN_WARNA_LUNAS : '#d9ead3';

  const matchedRows = [];
  for (let i = 0; i < dataPB.length; i++) {
    const pb = String(dataPB[i][0] || '').trim();
    if (pb === pbTarget) matchedRows.push(i + 2);
  }

  let idx = 0;
  while (idx < matchedRows.length) {
    const startRow = matchedRows[idx];
    let endRow = startRow;
    let next = idx + 1;
    while (next < matchedRows.length && matchedRows[next] === endRow + 1) {
      endRow = matchedRows[next];
      next++;
    }

    const numRows = endRow - startRow + 1;
    const statusValues = [];
    const warnaGrid = [];
    for (let r = 0; r < numRows; r++) {
      statusValues.push(['LUNAS']);
      const rowWarna = [];
      for (let c = 0; c < lastCol; c++) rowWarna.push(warnaLunas);
      warnaGrid.push(rowWarna);
    }

    sh.getRange(startRow, PEMBELIAN_COL_STATUS_LUNAS, numRows, 1).setValues(statusValues);
    sh.getRange(startRow, 1, numRows, lastCol).setBackgrounds(warnaGrid);

    idx = next;
  }

  return { ok: true, updatedCount: matchedRows.length };
}

/* ------------------------------------------------------
   AKSI: tandaiLunasPembelianBatch — Oktober 2026, FIX PERFORMA.

   VERSI BATCH dari aksiTandaiLunasPembelian_ di atas. Sebelum ini,
   syncHutangFromPembelian_() (dipanggil tiap abis Simpan Tanda Terima
   Barang) dan prosesPembayaranMultiPB() (popup Bayar Hutang) SAMA-SAMA
   loop manggil aksi "tandaiLunasPembelian" SATU PER SATU PB yang jadi
   LUNAS -- kalau sekaligus ada puluhan PB (mis. 94 PB dalam satu sync),
   itu jadi PULUHAN panggilan jaringan BERURUTAN ke Web App ini, tiap
   panggilan antre LockService sendiri-sendiri -- total bisa bermenit-
   menit, dan kepadatan request yang tinggi dalam waktu singkat beresiko
   kena limit Google (manifes sebagai error 404 "Sorry, unable to open
   the file at this time").

   Fix: SATU panggilan ini terima LIST PB sekaligus, baca kolom PB di
   PEMBELIAN SATU KALI SAJA (bukan 1x per PB), terus tandain LUNAS +
   hijau buat SEMUA PB yang match dalam satu pass. Logic per-PB (cari
   baris, kelompokkan jadi rentang contiguous, tulis status + warna)
   PERSIS SAMA kayak aksiTandaiLunasPembelian_ di atas, cuma sekarang
   dijalankan buat banyak PB tanpa baca ulang kolom PB tiap kali.

   body: { pbList: [string, ...] }
   return: { ok: true, updatedCount: <total baris ke-update>,
             perPb: { "<pb>": <jumlah baris>, ... } }
   ------------------------------------------------------ */
function aksiTandaiLunasPembelianBatch_(ss, body) {
  const pbListRaw = Array.isArray(body.pbList) ? body.pbList : [];
  const pbTargets = pbListRaw
    .map(function(p) { return String(p || '').trim(); })
    .filter(function(p) { return p; });

  if (pbTargets.length === 0) return { ok: true, updatedCount: 0, perPb: {} };

  const sh = ss.getSheetByName(PEMBELIAN_SHEET);
  if (!sh) throw new Error('Sheet "' + PEMBELIAN_SHEET + '" tidak ditemukan di database.');

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { ok: true, updatedCount: 0, perPb: {} };

  const lastCol = Math.max(sh.getLastColumn(), PEMBELIAN_COL_STATUS_LUNAS);
  const dataPB = sh.getRange(2, PEMBELIAN_COL_PB, lastRow - 1, 1).getValues(); // ⬅️ dibaca SEKALI buat semua PB

  const warnaLunas = typeof PENJUALAN_WARNA_LUNAS !== 'undefined' ? PENJUALAN_WARNA_LUNAS : '#d9ead3';
  const pbTargetSet = {};
  pbTargets.forEach(function(pb) { pbTargetSet[pb] = true; });

  // Kelompokkan SEMUA baris (lintas PB) jadi { pb: [rowNum, ...] } dalam satu scan.
  const matchedRowsByPb = {};
  for (let i = 0; i < dataPB.length; i++) {
    const pb = String(dataPB[i][0] || '').trim();
    if (!pb || !pbTargetSet[pb]) continue;
    if (!matchedRowsByPb[pb]) matchedRowsByPb[pb] = [];
    matchedRowsByPb[pb].push(i + 2);
  }

  let totalUpdated = 0;
  const perPb = {};

  pbTargets.forEach(function(pb) {
    const matchedRows = matchedRowsByPb[pb] || [];
    perPb[pb] = matchedRows.length;
    if (matchedRows.length === 0) return;
    totalUpdated += matchedRows.length;

    let idx = 0;
    while (idx < matchedRows.length) {
      const startRow = matchedRows[idx];
      let endRow = startRow;
      let next = idx + 1;
      while (next < matchedRows.length && matchedRows[next] === endRow + 1) {
        endRow = matchedRows[next];
        next++;
      }

      const numRows = endRow - startRow + 1;
      const statusValues = [];
      const warnaGrid = [];
      for (let r = 0; r < numRows; r++) {
        statusValues.push(['LUNAS']);
        const rowWarna = [];
        for (let c = 0; c < lastCol; c++) rowWarna.push(warnaLunas);
        warnaGrid.push(rowWarna);
      }

      sh.getRange(startRow, PEMBELIAN_COL_STATUS_LUNAS, numRows, 1).setValues(statusValues);
      sh.getRange(startRow, 1, numRows, lastCol).setBackgrounds(warnaGrid);

      idx = next;
    }
  });

  return { ok: true, updatedCount: totalUpdated, perPb: perPb };
}

/* ------------------------------------------------------
   HELPER PENJUALAN — format/parse riwayat cicilan & rupiah. Duplikat
   sengaja dari Pembayaran.gs sisi inventory (parseRiwayatCell_,
   formatRupiahText_, parseRupiahText_) — dua project Apps Script
   terpisah gak bisa saling manggil fungsi, logic & hasilnya harus
   identik.
   ------------------------------------------------------ */
function formatRupiahGatewayPenjualan_(number) {
  return 'Rp' + Number(number || 0).toLocaleString('id-ID');
}

function parseRupiahTextGatewayPenjualan_(text) {
  const digits = String(text).replace(/[^0-9]/g, '');
  return Number(digits) || 0;
}

// Format BARU "POST-NNN - dd/MM/yyyy - RpX - METODE" & format LAMA
// "dd/MM/yyyy - RpX - METODE" dua-duanya kebaca benar, sama persis
// parseRiwayatCell_ di Pembayaran.gs.
function parseRiwayatCellGatewayPenjualan_(text) {
  const str = String(text).trim();
  if (!str) return null;
  const parts = str.split(' - ');
  if (parts.length < 2) return null;

  let idx = 0;
  let noPost = '';
  if (parts[0].trim().indexOf(PENJUALAN_NOPOST_PREFIX_) === 0) {
    noPost = parts[0].trim();
    idx = 1;
  }

  const tanggal = (parts[idx] || '').trim();
  const jumlah = parseRupiahTextGatewayPenjualan_(parts[idx + 1]);
  const metode = parts.length > idx + 2 ? parts[idx + 2].trim() : '';
  return { tanggal: tanggal, jumlah: jumlah, metode: metode, noPost: noPost };
}

/* ------------------------------------------------------
   HELPER PENJUALAN — grouping per customer, backfill SISA BAYAR,
   auto-LUNAS, re-color. Port PERSIS dari groupPenjualanByCustomer_ /
   backfillSisaBayarPenjualan_ / autoLunaskanSisaBayarNol_ /
   reapplyWarnaLunasPenjualan_ (PENJUALAN.gs sisi inventory, versi
   SEBELUM migrasi) -- logic & hasil akhirnya identik, cuma sekarang
   jalan terhadap sheet PENJUALAN yang asli di database. TIDAK ada lagi
   urusan baris TOTAL virtual di sini (konsepnya dihapus total buat
   PENJUALAN yang sudah termigrasi -- lihat catatan arsitektur di
   komentar PENJUALAN_SHEET di atas).
   ------------------------------------------------------ */
function groupPenjualanByCustomerGateway_(sh) {
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 2) return;

  const data = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();

  const customerOrder = [];
  const customerRows = {};
  let lastCustomer = '';

  data.forEach(function (row) {
    const rawCustomer = String(row[1] || '').trim();
    const customer = rawCustomer || lastCustomer;
    if (rawCustomer) lastCustomer = rawCustomer;
    if (!customerRows[customer]) {
      customerRows[customer] = [];
      customerOrder.push(customer);
    }
    customerRows[customer].push(row);
  });

  const flatRows = [];
  customerOrder.forEach(function (customer) {
    customerRows[customer].forEach(function (row) {
      const newRow = row.slice();
      newRow[1] = customer;
      flatRows.push(newRow);
    });
  });

  if (flatRows.length === 0) return;

  sh.getRange(2, 1, flatRows.length, lastCol).setValues(flatRows);
  sh.getRange(2, 1, flatRows.length, 1).setNumberFormat('dd/mm/yyyy');
  sh.getRange(2, 5, flatRows.length, 1).setNumberFormat('#,##0');

  if (flatRows.length < data.length) {
    const sisaBaris = data.length - flatRows.length;
    sh.getRange(2 + flatRows.length, 1, sisaBaris, lastCol).clearContent();
  }
}

function backfillSisaBayarPenjualanGateway_(sh) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  const colTotal = PENJUALAN_COL_TOTAL;
  const colStatus = PENJUALAN_COL_STATUS;
  const colSisa = PENJUALAN_COL_SISA;
  const colEksp = PENJUALAN_COL_EKSPEDISI;
  const colPalet = PENJUALAN_COL_PALET;

  const lastCol = Math.max(colSisa, colEksp, colPalet, colTotal, colStatus);
  const data = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();

  const sisaBaru = data.map(function (row) {
    const total = Number(row[colTotal - 1]) || 0;
    const status = String(row[colStatus - 1] || '').trim();
    const sisa = row[colSisa - 1];

    if (!status && row[colTotal - 1] !== '' && row[colTotal - 1] !== null) {
      return [total]; // G = E doang, gak ikut H & I
    }
    return [sisa];
  });

  sh.getRange(2, colSisa, sisaBaru.length, 1).setValues(sisaBaru);
  sh.getRange(2, colSisa, sisaBaru.length, 1).setNumberFormat('#,##0');
}

function autoLunaskanSisaBayarNolGateway_(sh) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  const colStatus = PENJUALAN_COL_STATUS;
  const colSisa = PENJUALAN_COL_SISA;
  const lastCol = Math.max(colStatus, colSisa);
  const data = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();

  const statusBaru = data.map(function (row) {
    const sisaRaw = row[colSisa - 1];
    const sisaKosong = sisaRaw === '' || sisaRaw === null;
    if (!sisaKosong && Number(sisaRaw) === 0) return ['LUNAS'];
    return [row[colStatus - 1]];
  });

  sh.getRange(2, colStatus, statusBaru.length, 1).setValues(statusBaru);
}

function reapplyWarnaLunasPenjualanGateway_(sh) {
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return;

  const colStatus = PENJUALAN_COL_STATUS;
  const warnaLunas = PENJUALAN_WARNA_LUNAS;
  if (colStatus > lastCol) return;

  const statusValues = sh.getRange(2, colStatus, lastRow - 1, 1).getValues();
  const numRows = statusValues.length;
  const warnaGrid = [];
  for (let i = 0; i < numRows; i++) {
    const status = String(statusValues[i][0] || '').trim().toUpperCase();
    const warnaBaris = status === 'LUNAS' ? warnaLunas : null;
    const row = [];
    for (let c = 0; c < lastCol; c++) row.push(warnaBaris);
    warnaGrid.push(row);
  }
  if (numRows > 0) {
    sh.getRange(2, 1, numRows, lastCol).setBackgrounds(warnaGrid);
  }
}

/* ------------------------------------------------------
   AKSI: upsertPenjualan — versi database dari bagian UPSERT+GROUPING+
   BACKFILL+AUTO-LUNAS+RECOLOR di syncPenjualanFromTransaksi_() &
   syncPenjualanFromSparepart_() (PENJUALAN.gs lama). Satu aksi dipakai
   BERSAMA oleh kedua caller itu -- bedanya cuma flag updateCustomer
   (lihat di bawah).

   body: {
     entries: [{ noInv, tanggal, customer, method, total }, ...],
     updateCustomer: boolean  -- kalau true, kolom B (CUSTOMER) ikut
       ditimpa tiap sync walau baris itu SUDAH ADA (dipakai sync dari
       SPAREPART, karena NAMA TOKO bisa diedit ulang di sana). Kalau
       false/gak dikirim, baris yang SUDAH ADA kolom B-nya TIDAK
       disentuh sama sekali, cuma baris BARU yang diisi dari entry
       (dipakai sync dari TRANSAKSI, sama persis perilaku lama).
   }
   tanggal boleh Date/ISO-string/teks apa adanya, di-parse pakai
   parseTanggalDariGateway_ yang sudah ada.
   ------------------------------------------------------ */
function aksiUpsertPenjualan_(ss, body) {
  const entries = body.entries;
  if (!Array.isArray(entries)) throw new Error('entries harus array (boleh kosong).');

  const updateCustomer = !!body.updateCustomer;

  const sh = ss.getSheetByName(PENJUALAN_SHEET);
  if (!sh) throw new Error('Sheet "' + PENJUALAN_SHEET + '" tidak ditemukan di database.');

  if (entries.length === 0) {
    return { ok: true, updatedCount: 0, insertedCount: 0 };
  }

  const entryMap = {};
  const orderList = [];
  entries.forEach(function (e) {
    const noInv = String(e.noInv || '').trim();
    if (!noInv) return;
    entryMap[noInv] = e;
    orderList.push(noInv);
  });

  const lastRowP = sh.getLastRow();
  const existingRow = {};
  if (lastRowP >= 2) {
    const dataP = sh.getRange(2, PENJUALAN_COL_NOINV, lastRowP - 1, 1).getValues();
    dataP.forEach(function (r, i) {
      const noInv = String(r[0] || '').trim();
      if (noInv) existingRow[noInv] = i + 2;
    });
  }

  // 1) UPDATE baris yang sudah ada (A, D, E -- B cuma kalau updateCustomer,
  // F/G ke kanan / data pembayaran SAMA SEKALI TIDAK DISENTUH, sama kayak
  // versi lokal lama).
  let updatedCount = 0;
  Object.keys(existingRow).forEach(function (noInv) {
    const e = entryMap[noInv];
    if (!e) return;
    const rowIdx = existingRow[noInv];

    // ⬅️ FIX yang sama kayak versi lokal lama: kalau TOTAL berubah DAN
    // baris ini sudah pernah disentuh pembayaran (STATUS terisi, SISA
    // BAYAR masih angka murni), selisih TOTAL ikut ditambah/dikurang ke
    // SISA BAYAR juga, biar piutang yang beneran belum lunas tetap
    // konsisten sama TOTAL terbaru.
    const totalBaru = Number(e.total) || 0;
    const totalLama = Number(sh.getRange(rowIdx, PENJUALAN_COL_TOTAL).getValue()) || 0;
    const selisihTotal = totalBaru - totalLama;

    if (selisihTotal !== 0) {
      const statusRow = String(sh.getRange(rowIdx, PENJUALAN_COL_STATUS).getValue() || '').trim();
      const sisaRowRaw = sh.getRange(rowIdx, PENJUALAN_COL_SISA).getValue();
      if (statusRow && typeof sisaRowRaw === 'number') {
        const sisaBaruRow = sisaRowRaw + selisihTotal;
        sh.getRange(rowIdx, PENJUALAN_COL_SISA)
          .setValue(sisaBaruRow < 0 ? 0 : sisaBaruRow)
          .setNumberFormat('#,##0');
      }
    }

    if (updateCustomer) {
      sh.getRange(rowIdx, PENJUALAN_COL_CUSTOMER).setValue(e.customer || '');
    }
    sh.getRange(rowIdx, PENJUALAN_COL_TGLKIRIM).setValue(parseTanggalDariGateway_(e.tanggal)).setNumberFormat('dd/mm/yyyy');
    sh.getRange(rowIdx, PENJUALAN_COL_METHOD).setValue(e.method || '');
    sh.getRange(rowIdx, PENJUALAN_COL_TOTAL).setValue(totalBaru).setNumberFormat('#,##0');

    updatedCount++;
  });

  // 2) TAMBAHKAN invoice baru di baris paling bawah
  const newRows = [];
  orderList.forEach(function (noInv) {
    if (existingRow[noInv]) return;
    const e = entryMap[noInv];
    newRows.push([
      parseTanggalDariGateway_(e.tanggal), e.customer || '', noInv, e.method || '', Number(e.total) || 0
    ]);
  });

  if (newRows.length > 0) {
    const startRow = sh.getLastRow() + 1;
    sh.getRange(startRow, 1, newRows.length, 5).setValues(newRows);
    sh.getRange(startRow, 1, newRows.length, 1).setNumberFormat('dd/mm/yyyy');
    sh.getRange(startRow, 5, newRows.length, 1).setNumberFormat('#,##0');
  }

  // 3) GROUPING + BACKFILL + AUTO-LUNAS + RE-COLOR
  groupPenjualanByCustomerGateway_(sh);
  backfillSisaBayarPenjualanGateway_(sh);
  autoLunaskanSisaBayarNolGateway_(sh);
  reapplyWarnaLunasPenjualanGateway_(sh);

  return { ok: true, updatedCount: updatedCount, insertedCount: newRows.length };
}

/* ------------------------------------------------------
   AKSI: sortPenjualanByCustomer — versi database dari
   sortPenjualanBySeries_() (FILTER_SIDEBAR.gs lama, dipanggil dari
   checkbox "Urutkan berdasarkan Customer" di sidebar Filter). Susun
   ulang SEMUA baris PENJUALAN per customer (grup diurutkan alfabetis),
   TANPA menyisipkan baris ringkasan/TOTAL apapun -- itu sudah
   dihilangkan total dari alur PENJUALAN (lihat catatan arsitektur di
   komentar PENJUALAN_SHEET di atas). Baris tanpa customer (kalau ada)
   dikelompokkan di bawah kunci "" biar jumlah baris sebelum/sesudah
   selalu sama persis, gak ada yang nyangkut jadi duplikat.

   body: {}  -- gak butuh parameter apa-apa, selalu urutkan SELURUH
   tabel PENJUALAN yang ada sekarang.
   ------------------------------------------------------ */
function aksiSortPenjualanByCustomer_(ss, body) {
  const sh = ss.getSheetByName(PENJUALAN_SHEET);
  if (!sh) throw new Error('Sheet "' + PENJUALAN_SHEET + '" tidak ditemukan di database.');

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 2) return { ok: true, sorted: false, rowCount: 0 };

  const data = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();

  const groups = {};
  const customerOrder = [];

  data.forEach(function (row) {
    const customer = String(row[1] || '').trim();
    if (!groups[customer]) {
      groups[customer] = [];
      customerOrder.push(customer);
    }
    groups[customer].push(row);
  });

  customerOrder.sort(function (a, b) { return a.localeCompare(b, 'id'); });

  const output = [];
  customerOrder.forEach(function (customer) {
    groups[customer].forEach(function (row) { output.push(row); });
  });

  if (output.length === 0) return { ok: true, sorted: false, rowCount: 0 };

  sh.getRange(2, 1, output.length, lastCol).setValues(output);
  sh.getRange(2, 1, output.length, 1).setNumberFormat('dd/mm/yyyy');
  sh.getRange(2, 5, output.length, 1).setNumberFormat('#,##0');

  if (output.length < data.length) {
    const sisaBaris = data.length - output.length;
    sh.getRange(2 + output.length, 1, sisaBaris, lastCol).clearContent();
  }

  reapplyWarnaLunasPenjualanGateway_(sh);

  return { ok: true, sorted: true, rowCount: output.length };
}

/* ------------------------------------------------------
   HELPER PENJUALAN — cari 1 invoice (port PERSIS dari cariDataInvoice()
   di Pembayaran.gs sisi inventory, cuma dijalankan di sini terhadap
   sheet PENJUALAN yang ASLI, bukan mirror -- harus akurat detik-itu-juga
   sebelum sebuah pembayaran diproses).
   ------------------------------------------------------ */
function cariDataInvoiceGateway_(sh, noInvoice) {
  const trimmed = String(noInvoice || '').trim();
  const formatValid = /^(INV|SPRIN)-\d{4}-\d{2}-\d{3}$/.test(trimmed);
  if (!formatValid) {
    return { found: false, message: 'Format No Invoice salah. Gunakan format: INV-2026-07-001' };
  }

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { found: false, message: 'Data PENJUALAN masih kosong.' };

  const lastCol = Math.max(sh.getLastColumn(), PENJUALAN_COL_CICILAN_START);
  const data = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();

  let idx = -1;
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][PENJUALAN_COL_NOINV - 1] || '').trim() === trimmed) {
      idx = i;
      break;
    }
  }

  if (idx === -1) {
    return { found: false, message: 'Invoice ' + trimmed + ' tidak ditemukan di PENJUALAN.' };
  }

  let namaToko = String(data[idx][PENJUALAN_COL_CUSTOMER - 1] || '').trim();
  let scan = idx;
  while (!namaToko && scan > 0) {
    scan--;
    namaToko = String(data[scan][PENJUALAN_COL_CUSTOMER - 1] || '').trim();
  }

  const totalDasar = Number(data[idx][PENJUALAN_COL_TOTAL - 1]) || 0;
  const ekspedisi = Number(data[idx][PENJUALAN_COL_EKSPEDISI - 1]) || 0;
  const palet = Number(data[idx][PENJUALAN_COL_PALET - 1]) || 0;
  const total = totalDasar; // G = E doang, gak ikut H & I

  const statusRaw = String(data[idx][PENJUALAN_COL_STATUS - 1] || '').trim().toUpperCase();
  const sisaCell = data[idx][PENJUALAN_COL_SISA - 1];

  const sisaTagihan = statusRaw ? Number(sisaCell) || 0 : total;
  const totalSudahDibayar = total - sisaTagihan;
  const status = statusRaw === 'LUNAS' ? 'Lunas' : 'Cicilan';

  const riwayatPembayaran = [];
  for (let c = PENJUALAN_COL_CICILAN_START - 1; c < data[idx].length; c++) {
    const parsed = parseRiwayatCellGatewayPenjualan_(data[idx][c]);
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
    palet: palet
  };
}

// Port PERSIS dari getNextNoPost_() di Pembayaran.gs: scan ULANG semua
// riwayat cicilan yang BENERAN ada di PENJUALAN (sumber kebenaran utama),
// dibandingkan ke cache SYSTEM!B4 (kalau sheet SYSTEM ada di database
// ini -- opsional, gak wajib dibikin, fallback aman ke scan doang kalau
// gak ada), ambil yang PALING BESAR +1.
function getNextNoPostGateway_(ss, sh) {
  const system = ss.getSheetByName('SYSTEM');

  let maxDariData = 0;
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow >= 2 && lastCol >= PENJUALAN_COL_CICILAN_START) {
    const data = sh
      .getRange(2, PENJUALAN_COL_CICILAN_START, lastRow - 1, lastCol - PENJUALAN_COL_CICILAN_START + 1)
      .getValues();

    data.forEach(function (row) {
      row.forEach(function (cell) {
        const str = String(cell || '').trim();
        const m = str.match(/^POST-(\d+) - /);
        if (!m) return;
        const num = parseInt(m[1], 10);
        if (!isNaN(num) && num > maxDariData) maxDariData = num;
      });
    });
  }

  let counterCache = 0;
  if (system) {
    counterCache = Number(system.getRange(PENJUALAN_SYSTEM_NOPOST_CELL_).getValue()) || 0;
  }

  const nextNum = Math.max(maxDariData, counterCache) + 1;

  if (system) {
    system.getRange(PENJUALAN_SYSTEM_NOPOST_CELL_).setValue(nextNum);
  }

  return PENJUALAN_NOPOST_PREFIX_ + Utilities.formatString('%03d', nextNum);
}

/* ------------------------------------------------------
   AKSI: prosesPembayaranPenjualan — versi database dari prosesPembayaran_()
   (popup Pembayaran, Pembayaran.gs lama). Lock lokal yang dulu dipakai di
   inventory TIDAK PERLU dibawa ke sini -- prosesRequestGateway_ di atas
   SUDAH mengunci SEMUA aksi gateway sebelum manggil handler manapun
   (termasuk ini), jadi proteksi race condition (2 admin posting hampir
   bersamaan) tetap terjaga, cuma titiknya sekarang di sini.

   body: { noInvoice, jumlahBayar, metodeBayar ("TUNAI"/"TRANSFER"/
   "ADJUSTMENT"), tglBayarText ("dd/MM/yyyy", sudah diformat di sisi
   inventory lewat formatTglBayarInput_) }

   Validasi gagal -> throw Error (sama kayak aksi gateway lain), BUKAN
   return {ok:false,...} -- pemanggil (prosesPembayaran_ versi baru di
   Pembayaran.gs sisi inventory) yang nangkep & convert balik ke bentuk
   {success:false, message:...} biar kontrak popup UI gak berubah.
   ------------------------------------------------------ */
function aksiProsesPembayaranPenjualan_(ss, body) {
  const noInvoice = String(body.noInvoice || '').trim();
  const jumlahBayar = Number(body.jumlahBayar);
  const metodeBayar = String(body.metodeBayar || '').trim().toUpperCase();
  const tglBayarText = String(body.tglBayarText || '').trim();

  if (!jumlahBayar || jumlahBayar <= 0) throw new Error('Jumlah bayar tidak valid.');
  if (metodeBayar !== 'TUNAI' && metodeBayar !== 'TRANSFER' && metodeBayar !== 'ADJUSTMENT') {
    throw new Error('Metode pembayaran harus TUNAI atau TRANSFER atau ADJUSTMENT.');
  }
  if (!tglBayarText) throw new Error('Tanggal pembayaran tidak valid.');

  const sh = ss.getSheetByName(PENJUALAN_SHEET);
  if (!sh) throw new Error('Sheet "' + PENJUALAN_SHEET + '" tidak ditemukan di database.');

  const current = cariDataInvoiceGateway_(sh, noInvoice);
  if (!current.found) throw new Error(current.message || 'Invoice tidak ditemukan.');

  const bayarBulat = Math.round(jumlahBayar * 100);
  const sisaBulat = Math.round(current.sisaTagihan * 100);
  const isLebih = bayarBulat > sisaBulat;
  const lebih = isLebih ? (jumlahBayar - current.sisaTagihan) : 0;

  const rowIndex = current.rowIndex;
  const sisaBaru = current.sisaTagihan - jumlahBayar;
  const statusBaru = sisaBaru <= 0 ? 'LUNAS' : 'CICIL';

  sh.getRange(rowIndex, PENJUALAN_COL_STATUS).setValue(statusBaru);

  const sisaRange = sh.getRange(rowIndex, PENJUALAN_COL_SISA);
  if (isLebih) {
    sisaRange.setValue('LEBIH ' + formatRupiahGatewayPenjualan_(lebih));
  } else {
    sisaRange.setValue(sisaBaru).setNumberFormat('#,##0');
  }

  let targetCol = PENJUALAN_COL_CICILAN_START;
  while (sh.getRange(rowIndex, targetCol).getValue()) {
    targetCol++;
  }

  const noPost = getNextNoPostGateway_(ss, sh);

  const cellText = noPost + ' - ' + tglBayarText + ' - ' + formatRupiahGatewayPenjualan_(jumlahBayar) + ' - ' + metodeBayar;
  sh.getRange(rowIndex, targetCol).setValue(cellText);

  const lastColSheet = Math.max(sh.getLastColumn(), targetCol);
  if (statusBaru === 'LUNAS') {
    sh.getRange(rowIndex, 1, 1, lastColSheet).setBackground(PENJUALAN_WARNA_LUNAS);
  }
  if (isLebih) {
    sisaRange.setBackground(PENJUALAN_WARNA_LEBIH);
  }

  const riwayatBaru = current.riwayatPembayaran.slice();
  riwayatBaru.push({ noPost: noPost, tanggal: tglBayarText, jumlah: jumlahBayar, metode: metodeBayar });

  return {
    ok: true,
    noInvoice: current.noInvoice,
    namaToko: current.namaToko,
    totalTagihan: current.totalTagihan,
    totalSudahDibayar: current.totalTagihan - sisaBaru,
    sisaTagihan: sisaBaru,
    status: statusBaru === 'LUNAS' ? 'Lunas' : 'Cicilan',
    noPost: noPost,
    riwayatPembayaran: riwayatBaru,
    jadiLunas: statusBaru === 'LUNAS'
  };
}

/* ------------------------------------------------------
   AKSI: updateBiayaTambahanPenjualan — versi database dari bagian
   PENJUALAN di simpanBiayaTambahan() + hitungUlangSisaBayar_()
   (BiayaTambahan.gs lama). Update kolom H (EKSPEDISI) atau I (PALET)
   = nominal (DITIMPA, bukan ditambah), lalu recompute SISA BAYAR (G) &
   STATUS (F) 1 baris yang match REFERENSI (= kolom NO INV).

   body: { referensi, jenisBiaya ("EKSPEDISI"/"PALET"), nominal }
   ------------------------------------------------------ */
function aksiUpdateBiayaTambahanPenjualan_(ss, body) {
  const referensi = String(body.referensi || '').trim();
  const jenisBiaya = String(body.jenisBiaya || '').trim().toUpperCase();
  const nominal = Number(body.nominal);

  if (!referensi) throw new Error('referensi tidak dikirim.');
  if (jenisBiaya !== 'EKSPEDISI' && jenisBiaya !== 'PALET') {
    throw new Error('jenisBiaya harus EKSPEDISI atau PALET.');
  }
  if (!nominal || nominal <= 0) throw new Error('nominal tidak valid.');

  const sh = ss.getSheetByName(PENJUALAN_SHEET);
  if (!sh) throw new Error('Sheet "' + PENJUALAN_SHEET + '" tidak ditemukan di database.');

  const lastRow = sh.getLastRow();
  if (lastRow < 2) throw new Error('Data PENJUALAN masih kosong.');

  const lastCol = Math.max(sh.getLastColumn(), PENJUALAN_COL_CICILAN_START);
  const dataRange = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();

  let idx = -1;
  for (let i = 0; i < dataRange.length; i++) {
    if (String(dataRange[i][PENJUALAN_COL_NOINV - 1] || '').trim() === referensi) {
      idx = i;
      break;
    }
  }

  if (idx === -1) {
    throw new Error('Referensi ' + referensi + ' tidak ditemukan di kolom NO INV PENJUALAN.');
  }

  const rowIndex = idx + 2;
  const targetCol = jenisBiaya === 'EKSPEDISI' ? PENJUALAN_COL_EKSPEDISI : PENJUALAN_COL_PALET;
  sh.getRange(rowIndex, targetCol).setValue(nominal).setNumberFormat('#,##0');

  const hasil = hitungUlangSisaBayarGateway_(sh, rowIndex);

  return {
    ok: true,
    referensi: referensi,
    jenisBiaya: jenisBiaya,
    nominal: nominal,
    totalTagihan: hasil.totalTagihan,
    sisaBayar: hasil.sisaBayar,
    status: hasil.status
  };
}

// Port PERSIS dari hitungUlangSisaBayar_() di BiayaTambahan.gs sisi
// inventory -- logic & hasil akhirnya identik, cuma sekarang jalan
// terhadap sheet PENJUALAN yang asli.
function hitungUlangSisaBayarGateway_(sh, rowIndex) {
  const lastCol = Math.max(sh.getLastColumn(), PENJUALAN_COL_CICILAN_START);
  const rowValues = sh.getRange(rowIndex, 1, 1, lastCol).getValues()[0];

  const totalTagihan = Number(rowValues[PENJUALAN_COL_TOTAL - 1]) || 0;

  let totalSudahDibayar = 0;
  for (let c = PENJUALAN_COL_CICILAN_START - 1; c < rowValues.length; c++) {
    const parsed = parseRiwayatCellGatewayPenjualan_(rowValues[c]);
    if (parsed) totalSudahDibayar += parsed.jumlah;
  }

  const statusLama = String(rowValues[PENJUALAN_COL_STATUS - 1] || '').trim().toUpperCase();
  const sisaRange = sh.getRange(rowIndex, PENJUALAN_COL_SISA);
  const statusRange = sh.getRange(rowIndex, PENJUALAN_COL_STATUS);

  if (totalSudahDibayar <= 0) {
    statusRange.setValue('');
    sisaRange.setValue(totalTagihan).setNumberFormat('#,##0').setBackground(null);
    if (statusLama === 'LUNAS') {
      sh.getRange(rowIndex, 1, 1, lastCol).setBackground(null);
    }
    return { totalTagihan: totalTagihan, sisaBayar: totalTagihan, status: '' };
  }

  const sisaBaru = totalTagihan - totalSudahDibayar;

  if (sisaBaru <= 0) {
    statusRange.setValue('LUNAS');
    sh.getRange(rowIndex, 1, 1, lastCol).setBackground(PENJUALAN_WARNA_LUNAS);
    if (sisaBaru < 0) {
      sisaRange.setValue('LEBIH ' + formatRupiahGatewayPenjualan_(-sisaBaru)).setBackground(PENJUALAN_WARNA_LEBIH);
    } else {
      sisaRange.setValue(0).setNumberFormat('#,##0').setBackground(PENJUALAN_WARNA_LUNAS);
    }
    return { totalTagihan: totalTagihan, sisaBayar: sisaBaru, status: 'LUNAS' };
  }

  statusRange.setValue('CICIL');
  sisaRange.setValue(sisaBaru).setNumberFormat('#,##0').setBackground(null);
  if (statusLama === 'LUNAS') {
    sh.getRange(rowIndex, 1, 1, lastCol).setBackground(null);
  }
  return { totalTagihan: totalTagihan, sisaBayar: sisaBaru, status: 'CICIL' };
}

/* ------------------------------------------------------
   AKSI: getNomorPoMasukBersama — SATU counter "No PO Masuk" yang
   dipakai BERSAMA oleh inventory_SANFELO dan inventory_SL, supaya
   nomor PO di kedua sistem itu saling melanjutkan (gak ada 2 sistem
   yang pakai nomor yang sama, dan gak ada nomor yang kelewat/bentrok).

   Ditemukan dari insiden 2026-10-01: data dobel di HISTORY_PO_MASUK
   (PO #2026-223/#2026-224 orphan) — salah satu akar masalahnya adalah
   SANFELO dan SL masing-masing punya counter SENDIRI-SENDIRI
   (SYSTEM!B3/C3 di sheet-nya masing-masing) yang kebetulan aja gak
   pernah tabrakan. User minta supaya ini jadi SATU counter yang
   beneran sinkron, dicek dari gateway bersama ini, bukan cuma
   kebetulan gak tabrakan.

   Disimpan di SHEET "SYSTEM" milik database_SANFELO sendiri (BUKAN
   SYSTEM!B3/C3 yang lama — itu tetap dipakai programnya sanfelo utk
   keperluan lain yang gak terkait no PO Masuk bersama ini, dibiarkan
   apa adanya), supaya tetap 1 lokasi yang gampang diaudit, gak perlu
   nambah Web App/Script baru. Dilindungi otomatis sama LockService
   global yang sudah bungkus SEMUA aksi di prosesRequestGateway_, jadi
   gak akan ada 2 request masuk bareng dapat nomor yang sama.

   body: { asal }  -- asal cuma buat log, isi "SANFELO" atau "SL".
   return: { ok: true, nomor: "#2026-225" }

   SETUP SEKALI DI AWAL (WAJIB, baca juga di pesan error di bawah):
   1. Buat sheet baru di database_SANFELO, namanya PERSIS "SYSTEM"
      (kalau belum ada).
   2. Isi cell B6 dengan ANGKA PO MASUK TERTINGGI yang sudah pernah
      kepake tahun ini -- CEK KEDUA SISI (SANFELO dan SL), ambil mana
      yang lebih besar.
   3. Isi cell C6 dengan tahun berjalan (contoh: 2026).
   4. (Opsional) buat sheet "LOG_PO_BERSAMA" kalau mau ada jejak audit
      asal tiap nomor PO (SANFELO atau SL) -- kalau sheet ini gak ada,
      logging dilewati diam-diam, gak bikin request gagal.
   Setelah seeding awal ini, counter jalan otomatis sendiri dan gak
   perlu disentuh manual lagi kecuali mau koreksi manual.
   ------------------------------------------------------ */
function aksiGetNomorPoMasukBersama_(ss, body) {
  const system = ss.getSheetByName('SYSTEM');
  if (!system) {
    throw new Error(
      'Sheet "SYSTEM" belum ada di database_SANFELO. Buat dulu sheet baru ' +
      'bernama persis "SYSTEM", isi cell ' + SHARED_PO_MASUK_COUNTER_CELL_ +
      ' dengan nomor PO MASUK tertinggi yang pernah kepake tahun ini (cek ' +
      'SANFELO maupun SL, mana yang lebih besar) dan cell ' +
      SHARED_PO_MASUK_PERIODE_CELL_ + ' dengan tahun berjalan, baru coba lagi.'
    );
  }

  const asal = String(body.asal || '').trim().toUpperCase() || 'TIDAK DIKETAHUI';
  const tahunSekarang = String(new Date().getFullYear());

  const periodeTersimpan = String(system.getRange(SHARED_PO_MASUK_PERIODE_CELL_).getValue() || '').trim();
  let counterSaatIni = 0;
  if (periodeTersimpan === tahunSekarang) {
    counterSaatIni = Number(system.getRange(SHARED_PO_MASUK_COUNTER_CELL_).getValue()) || 0;
  }

  const nextNum = counterSaatIni + 1;
  system.getRange(SHARED_PO_MASUK_COUNTER_CELL_).setValue(nextNum);
  system.getRange(SHARED_PO_MASUK_PERIODE_CELL_).setValue(tahunSekarang);

  const nomor = '#' + tahunSekarang + '-' + Utilities.formatString('%03d', nextNum);

  try {
    const log = ss.getSheetByName(SHARED_PO_MASUK_LOG_SHEET_);
    if (log) {
      log.appendRow([new Date(), asal, nomor]);
    }
  } catch (errLog) {
    // gak boleh gagalin permintaan nomor cuma gara-gara logging gagal
  }

  return { ok: true, nomor: nomor };
}

/* ------------------------------------------------------
   AKSI BARU (Oktober 2026 — perbaikan alur, atas permintaan user):
   peekNomorPoMasukBersama — INTIP nomor PO Masuk berikutnya TANPA
   mengklaim/nge-increment apapun. Read-only murni.

   Kenapa ini ditambahkan: dulu poBaru() (tombol "PO Masuk Baru") LANGSUNG
   manggil getNomorPoMasukBersama_ (di atas) begitu form dibuka -- jadi
   nomor itu "kebakar"/kepake SAAT ITU JUGA, walau admin belum tentu jadi
   nyimpen PO-nya. Akibatnya kalau SANFELO & SL sama-sama buka "PO Masuk
   Baru" hampir bersamaan, keduanya langsung dapat nomor BEDA (mis. SL
   dapat #231 duluan, begitu SANFELO buka form-nya kebagian #232) --
   padahal belum tentu dua-duanya jadi dipakai.

   Sekarang: poBaru() di KEDUA sistem (SANFELO & SL) pakai aksi ini buat
   nampilin nomor CALON di form (gak nyentuh/nambah counter), jadi kalau
   dibuka bersamaan dua-duanya nampilin nomor yang SAMA. Nomor itu baru
   BENERAN diklaim (lewat getNomorPoMasukBersama_ di atas) pas admin
   beneran mulai nyimpen (baris item pertama ke-sync ke HISTORY_PO_MASUK,
   lihat perubahan di syncToHistory() kedua sistem) -- siapa yang nyimpen
   duluan itu yang dapat nomornya; sistem yang satunya, begitu DIA juga
   mulai nyimpen, otomatis ke-deteksi nomor previewnya sudah kepake orang
   lain dan diganti ke nomor berikutnya + admin dikasih tau lewat alert.
   ------------------------------------------------------ */
function aksiPeekNomorPoMasukBersama_(ss, body) {
  const system = ss.getSheetByName('SYSTEM');
  if (!system) {
    throw new Error(
      'Sheet "SYSTEM" belum ada di database_SANFELO. Buat dulu sheet baru ' +
      'bernama persis "SYSTEM", isi cell ' + SHARED_PO_MASUK_COUNTER_CELL_ +
      ' dengan nomor PO MASUK tertinggi yang pernah kepake tahun ini dan ' +
      'cell ' + SHARED_PO_MASUK_PERIODE_CELL_ + ' dengan tahun berjalan, ' +
      'baru coba lagi.'
    );
  }

  const tahunSekarang = String(new Date().getFullYear());
  const periodeTersimpan = String(system.getRange(SHARED_PO_MASUK_PERIODE_CELL_).getValue() || '').trim();
  let counterSaatIni = 0;
  if (periodeTersimpan === tahunSekarang) {
    counterSaatIni = Number(system.getRange(SHARED_PO_MASUK_COUNTER_CELL_).getValue()) || 0;
  }

  const nextNum = counterSaatIni + 1;
  const nomor = '#' + tahunSekarang + '-' + Utilities.formatString('%03d', nextNum);

  return { ok: true, nomor: nomor };
}

/* ------------------------------------------------------
   AKSI: bacaSheet — baca data FRESH langsung dari sheet asli di
   database (bukan lewat mirror IMPORTRANGE yang ada delay-nya).
   Dipakai cuma buat pengecekan yang harus akurat detik-itu-juga,
   misalnya cek toko LOCKED sebelum bikin Surat Jalan.

   body: { sheet, startRow, startCol, numCols }
   startRow/startCol default ke 2/1 (baris data pertama, kolom A).
   Baris terakhir otomatis ngikutin lastRow sheet itu saat ini juga.
   ------------------------------------------------------ */
function aksiBacaSheet_(ss, body) {
  const sheetName = String(body.sheet || '').trim();
  if (!sheetName) throw new Error('Nama sheet (sheet) tidak dikirim.');
  if (SHEET_BACA_DIIZINKAN.indexOf(sheetName) === -1) {
    throw new Error('Sheet "' + sheetName + '" tidak ada di daftar SHEET_BACA_DIIZINKAN.');
  }

  const startRow = Number(body.startRow) || 2;
  const startCol = Number(body.startCol) || 1;

  const sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error('Sheet "' + sheetName + '" tidak ditemukan di database.');

  const lastRow = sh.getLastRow();
  if (lastRow < startRow) {
    return { ok: true, rows: [] };
  }

  // ⬅️ BARU (migrasi PENJUALAN, Oktober 2026): numCols sekarang OPSIONAL.
  // Kalau gak dikirim (atau dikirim 0), baca "selebar sheet-nya saat ini"
  // (sh.getLastColumn()) -- dipakai buat sheet dengan kolom yang terus
  // nambah ke kanan tanpa batas pasti (riwayat cicilan PENJUALAN mulai
  // kolom J, mirip kolom riwayat HUTANG mulai K), biar caller gak perlu
  // nebak/pasang buffer angka kolom yang mungkin kurang lebar ke depannya.
  // numCols eksplisit (seperti yang sudah dipakai HUTANG/SPAREPART/
  // PEMBELIAN) tetap jalan persis seperti sebelumnya, gak ada yang berubah.
  const numColsRaw = Number(body.numCols) || 0;
  const numCols = numColsRaw > 0 ? numColsRaw : Math.max(1, sh.getLastColumn() - startCol + 1);

  const numRows = lastRow - startRow + 1;
  const rows = sh.getRange(startRow, startCol, numRows, numCols).getValues();

  return { ok: true, rows: rows };
}

/* ------------------------------------------------------
   Fungsi uji coba manual (jalankan dari editor Apps Script)
   ------------------------------------------------------ */
function testLangsungDariDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('TEST');
  if (!sh) {
    SpreadsheetApp.getUi().alert('Bikin dulu sheet baru namanya "TEST" di file ini, lalu run lagi.');
    return;
  }
  const hasil = aksiAppendRows_(ss, { sheet: 'TEST', rows: [[new Date(), 'Test langsung dari database', 123]] });
  SpreadsheetApp.getUi().alert('Berhasil: ' + JSON.stringify(hasil));
}

/* ------------------------------------------------------
   PERBAIKAN SEKALI JALAN (Oktober 2026) — jalankan MANUAL dari editor
   Apps Script (pilih function ini di dropdown atas, klik Run), BUKAN
   dari gateway/Web App. Cuma perlu dijalankan SEKALI.

   Ini bukan aksi gateway baru -- ini pembersih utk baris yang SUDAH
   KETERLANJUR tersimpan dengan tanggal berformat teks ISO mentah
   ("2026-09-30T00:00:00.000Z") gara-gara bug di aksiSyncPoMasuk_ /
   aksiSyncPoKeluar_ (lihat komentar FIX di masing-masing fungsi itu) --
   fix di fungsi-fungsi itu cuma mencegah data BARU rusak, TIDAK
   otomatis membetulkan data LAMA yang sudah terlanjur tersimpan rusak.

   Yang di-scan & dibetulkan (ubah jadi Date asli + format dd/mm/yyyy,
   NILAI/arti tanggalnya TIDAK berubah, cuma bentuk penyimpanannya):
   - HISTORY_PO_MASUK kolom C (TGL PO) & H (TGL_KIRIM)
   - HISTORY_PO_KELUAR kolom B (TGL)
   Sel yang sudah berupa Date asli (baris lama, sebelum bug ini
   muncul) dilewati apa adanya -- aman dijalankan berkali-kali kalau
   ragu, gak bakal ngerusak yang sudah benar.
   ------------------------------------------------------ */
function perbaikiTanggalRusakHistoryPO() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let totalDiperbaiki = 0;

  function perbaikiKolom(sheetName, kolom) {
    const sh = ss.getSheetByName(sheetName);
    if (!sh) return;

    const lastRow = sh.getLastRow();
    if (lastRow < 2) return;

    const range = sh.getRange(2, kolom, lastRow - 1, 1);
    const values = range.getValues();
    let adaPerubahanNilai = false;

    for (let i = 0; i < values.length; i++) {
      const v = values[i][0];
      if (!v || v instanceof Date) continue; // kosong / sudah Date asli -- cuma betulin teks ISO
      const str = String(v).trim();
      if (!/^\d{4}-\d{2}-\d{2}/.test(str)) continue; // bukan pola ISO -> jangan diotak-atik

      const parsed = new Date(str);
      if (isNaN(parsed.getTime())) continue; // gak valid -> lewati, jangan maksa

      values[i][0] = parsed;
      adaPerubahanNilai = true;
      totalDiperbaiki++;
    }

    if (adaPerubahanNilai) {
      range.setValues(values);
    }

    // ⬅️ FIX: SEBELUMNYA setNumberFormat('dd/mm/yyyy') cuma dipanggil kalau
    // ada nilai TEKS yang kekonversi (adaPerubahanNilai true) -- jadi kalau
    // ternyata sel-selnya SUDAH berupa Date asli dari awal (bukan teks),
    // tapi FORMAT TAMPILAN kolomnya kebetulan ke-set "yyyy-mm-dd" (bukan
    // "dd/mm/yyyy") -- entah dari sononya atau gara-gara setNumberFormat
    // yang gak lengkap sebelumnya -- makanya tetap keliatan "2026-09-30"
    // padahal datanya sendiri sudah benar. Sekarang format SELALU dipaksa
    // ulang ke dd/mm/yyyy di SELURUH kolom, gak peduli ada nilai teks yang
    // dibetulin atau nggak -- biar kasus "datanya benar tapi tampilannya
    // salah" ini ikut kebetulin juga.
    range.setNumberFormat('dd/mm/yyyy');
  }

  perbaikiKolom(HISTORY_PO_MASUK_SHEET, 3); // C TGL PO
  perbaikiKolom(HISTORY_PO_MASUK_SHEET, 8); // H TGL_KIRIM
  perbaikiKolom(HISTORY_PO_KELUAR_SHEET, 2); // B TGL

  SpreadsheetApp.getUi().alert('Selesai. Total sel tanggal yang diperbaiki: ' + totalDiperbaiki);
}