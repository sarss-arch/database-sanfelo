/* ======================================================
   WEB_APP.js   (PROJECT: database_SANFELO)
   ------------------------------------------------------
   Web admin SANFELO (tahap 1): dashboard, data, pembayaran piutang,
   bayar hutang, detail invoice.

   PRINSIP KEAMANAN DATA:
   - Data tetap di spreadsheet database_SANFELO (file ini menempel di sana).
   - SEMUA penulisan lewat prosesRequestGateway_ (DatabaseGateway.js) --
     jalur, kunci (LockService), dan logika yang SAMA PERSIS dengan yang
     dipakai inventory_SANFELO. Web ini tidak punya logika tulis sendiri,
     jadi hasilnya sama dengan posting dari spreadsheet inventory, dan
     posting dari web & dari inventory tidak bisa saling tabrak.
   - Pembacaan langsung dari sheet database (selalu data terbaru).
   - Halaman hanya terbuka dengan link berkunci (?k=...) + login nama & PIN.
   - inventory_SANFELO TIDAK diubah sama sekali dan tetap bisa dipakai.

   SETUP (pemilik, sekali): lihat PANDUAN di WEB_SETUP.js.
   ====================================================== */

const WEB_PROP_KUNCI = 'WEB_KUNCI_AKSES';
const WEB_PREFIX_PIN = 'WEB_PIN_';
const WEB_SESI_DETIK = 6 * 3600;
const WEB_LOG_SHEET = 'LOG_WEB';
const WEB_TZ = 'Asia/Jakarta';

function doGet(e) {
  const kunci = PropertiesService.getScriptProperties().getProperty(WEB_PROP_KUNCI);
  const k = e && e.parameter ? e.parameter.k : '';
  if (!kunci || k !== kunci) {
    return HtmlService.createHtmlOutput('<p style="font-family:sans-serif;padding:24px">Halaman tidak tersedia.</p>')
      .setTitle('SANFELO');
  }
  const t = HtmlService.createTemplateFromFile('WebIndex');
  t.konfig = JSON.stringify({
    merek: 'SANFELO', merekPanjang: 'SANFELO LIVING', ketAdj: true, fitur: ['po', 'sj', 'tt', 'biaya', 'spr', 'stok', 'kas', 'laporan', 'cetak', 'latihan', 'alat', 'pok'],
    catatan: 'di database_SANFELO, lewat jalur yang sama dengan spreadsheet inventory'
  });
  return t.evaluate()
    .setTitle('SANFELO Admin')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function webInclude_(nama) {
  return HtmlService.createHtmlOutputFromFile(nama).getContent();
}

/* ---------------- login & sesi ---------------- */

function webLogin(nama, pin) {
  const namaBersih = String(nama || '').trim().toUpperCase();
  const cache = CacheService.getScriptCache();
  const kunciGagal = 'WEBGAGAL_' + namaBersih;
  const gagal = Number(cache.get(kunciGagal) || 0);
  if (gagal >= 5) return { ok: false, pesan: 'Terlalu banyak percobaan. Coba lagi 10 menit lagi.' };

  const pinBenar = PropertiesService.getScriptProperties().getProperty(WEB_PREFIX_PIN + namaBersih);
  if (!namaBersih || !pinBenar || String(pin || '').trim() !== pinBenar) {
    cache.put(kunciGagal, String(gagal + 1), 600);
    return { ok: false, pesan: 'Nama atau PIN salah.' };
  }
  cache.remove(kunciGagal);
  const token = Utilities.getUuid();
  cache.put('WEBSES_' + token, namaBersih, WEB_SESI_DETIK);
  webCatatLog_(namaBersih, 'LOGIN', '');
  return { ok: true, token: token, nama: namaBersih };
}

function webLogout(token) {
  CacheService.getScriptCache().remove('WEBSES_' + token);
  return { ok: true };
}

function webCekSesi_(token) {
  const cache = CacheService.getScriptCache();
  const nama = token ? cache.get('WEBSES_' + token) : null;
  if (!nama) throw new Error('SESI_HABIS');
  cache.put('WEBSES_' + token, nama, WEB_SESI_DETIK); // diperpanjang tiap dipakai
  return nama;
}

/**
 * Satu-satunya pintu dari halaman web ke server (selain login/logout).
 * Semua jawaban diubah ke bentuk yang aman dikirim ke browser.
 */
// MODE LATIHAN: semua alur dijalankan, tapi TIDAK ADA yang ditulis ke database
// (aksi tulis gateway hanya dicatat & dibalas "seolah berhasil").
var WEB_LATIHAN_ = false;
var WEB_LATIHAN_CATAT_ = [];
const WEB_AKSI_BACA_ = ['bacaSheet', 'peekNomorPoMasukBersama', 'cekStatusInvoiceCetak'];
const WEB_AKSI_ASLI_SAJA_ = ['catatInvoiceTercetak', 'batalInvoiceTercetak']; // handler-nya selalu menulis database asli

function webApi(token, aksi, args) {
  WEB_LATIHAN_ = !!(args && args.latihan);
  WEB_LATIHAN_CATAT_ = [];
  try {
    const admin = webCekSesi_(token);
    const fn = WEB_API_[aksi];
    if (!fn) throw new Error('Aksi "' + aksi + '" tidak dikenal.');
    const hasil = webJalankanAksi_(aksi, fn, args || {}, admin); // WEB_PERCEPAT.js: simpanan baca + versi data
    return webAman_({ ok: true, data: hasil, latihan: WEB_LATIHAN_ ? WEB_LATIHAN_CATAT_ : null });
  } catch (err) {
    const pesan = String(err && err.message ? err.message : err).replace(/^Gateway menolak: /, '');
    return { ok: false, pesan: pesan, sesiHabis: pesan === 'SESI_HABIS' };
  }
}

const WEB_API_ = {
  dashboard: function (a) { return webDashboard_(); },
  daftarData: function (a) { return webDaftarData_(a); },
  daftarPiutang: function (a) { return webDaftarPiutang_(a); },
  cariInvoice: function (a) { return webCariDataInvoice_(a.noInvoice); },
  bayarPiutang: function (a, admin) { return webProsesPembayaran_(a, admin); },
  ubahRiwayatBayar: function (a, admin) { return webUbahRiwayatBayar_(a, admin); },
  supplierHutang: function (a) { return webSupplierListByBulan_(a.bulan); },
  daftarPBHutang: function (a) { return webPBListByBulanSupplier_(a.bulan, a.supplier); },
  bayarHutang: function (a, admin) { return webProsesPembayaranMultiPB_(a, admin); },
  bulanHutang: function () { return webDaftarBulanHutang_(); },
  detailInvoice: function (a) { return webDetailInvoice_(a.noInvoice); },
  logWeb: function (a) { return webBacaLog_(a); },
  daftarAdmin: function () { return webDaftarAdmin_(); },
  versiData: function () { return webVersiData_(); },
  semuaData: function (a) { return webDaftarData_({ sheet: a.sheet, semua: true }); },
  dokumen: function (a) { return webDokumen_(a); },
  cekHutang: function () { return { tanpaBulan: [], yatim: [], sisaAneh: [] }; }, // pemeriksaan HUTANG baru ada di web SL
  aturPinAdmin: function (a, admin) { return webAturPinAdmin_(a, admin); },
  // WEB_KAS_CRUD.js: ubah / hapus baris petty cash & buku bank
  ubahKas: function (a, admin) { return webUbahKas_(a, admin); },
  hapusKas: function (a, admin) { return webHapusKas_(a, admin); },
  // WEB_PO_MASUK.js
  masterPO: function () { return webMasterPO_(); },
  tambahCustomer: function (a, admin) { return webTambahCustomer_(a, admin); },
  hapusCustomer: function (a, admin) { return webHapusCustomer_(a, admin); },
  ubahCustomer: function (a, admin) { return webUbahCustomer_(a, admin); }, // WEB_CUSTOMER.js
  gantiCustomerInvoice: function (a, admin) { return webGantiCustomerInvoice_(a, admin); },
  hitungPO: function (a) { return webHitungPO_(a); },
  simpanPO: function (a, admin) { return webSimpanPOMasuk_(a, admin); },
  bukaPO: function (a) { return webBukaPO_(a); },
  ubahPO: function (a, admin) { return webUbahPOMasuk_(a, admin); },
  // WEB_SURAT_JALAN.js
  daftarPOTerbuka: function (a) { return webDaftarPOTerbuka_(a); },
  muatPOSJ: function (a) { return webMuatPOuntukSJ_(a); },
  pratinjauSJ: function (a) { return webPratinjauSJ_(a); },
  simpanSJ: function (a, admin) { return webSimpanSJ_(a, admin); },
  // WEB_POSTING_LAIN.js
  masterTT: function () { return webMasterTT_(); },
  muatPOKeluarTT: function (a) { return webMuatPOKeluarTT_(a); },
  hitungTT: function (a) { return webHitungTT_(a); },
  simpanTT: function (a, admin) { return webSimpanTT_(a, admin); },
  invoiceUntukBiaya: function (a) { return webInvoiceUntukBiaya_(a.noInvoice); },
  simpanBiaya: function (a, admin) { return webSimpanBiaya_(a, admin); },
  daftarPalet: function () { return webDaftarPalet_(); },
  detailPalet: function (a) { return webDetailPalet_(a); },
  pratinjauSparepart: function (a) { return webPratinjauSparepart_(a); },
  simpanSparepart: function (a, admin) { return webSimpanSparepart_(a, admin); },
  // WEB_KAS_LAPORAN.js
  stok: function () { return webStok_(); },
  daftarKas: function (a) { return webDaftarKas_(a); },
  simpanPettyCash: function (a, admin) { return webSimpanPettyCash_(a, admin); },
  simpanBukuBank: function (a, admin) { return webSimpanBukuBank_(a, admin); },
  laporan: function (a) { return webLaporan_(a); },
  // WEB_CETAK.js
  cetak: function (a, admin) { return webCetak_(a, admin); },
  // WEB_ALAT.js
  siapkanLatihan: function (a) { return webSiapkanLatihan_(a); },
  alat: function (a, admin) { return webAlat_(a, admin); },
  tempoTokoDaftar: function () { return webTempoTokoDaftar_(); },
  perpanjangTempo: function (a, admin) { return webPerpanjangTempo_(a, admin); },
  batalTercetak: function (a, admin) { return webBatalTercetak_(a, admin); },
  susulanKandidat: function (a) { return webSusulanKandidat_(String(a.bulan || '')); },
  susulanSimpan: function (a, admin) { return webSusulanSimpan_(a, admin); },
  statusSJ: function (a) { return webDetailStatusSJ_(a); },
  ubahStatus: function (a, admin) { return webUbahStatus_(a, admin); },
  groupingDaftar: function () { return webGroupingDaftar_(); },
  groupingSimpan: function (a, admin) { return webGroupingSimpan_(a, admin); },
  poKeluarMuat: function (a) { return webPOKeluarMuat_(a); },
  poKeluarDariPOMasuk: function (a) { return webPOKeluarDariPOMasuk_(a); },
  poKeluarSimpan: function (a, admin) { return webPOKeluarSimpan_(a, admin); },
  bukaSJ: function (a) { return webBukaSJ_(a); },
  bukaTT: function (a) { return webBukaTT_(a); },
  bukaSpr: function (a) { return webBukaSpr_(a); }
};

/* ---------------- bantuan umum ---------------- */

// mode latihan: semua baca & tulis diarahkan ke FILE LATIHAN (salinan database_SANFELO),
// jadi simpan & cetak benar-benar berjalan tanpa menyentuh data asli
function webSS_() { return WEB_LATIHAN_ ? webLatihanSS_() : SpreadsheetApp.getActiveSpreadsheet(); }

const WEB_PROP_LATIHAN = 'WEB_FILE_LATIHAN_ID';
const WEB_PROP_LATIHAN_WAKTU = 'WEB_FILE_LATIHAN_WAKTU';
const WEB_LATIHAN_UMUR_JAM_ = 12; // file latihan otomatis dibuat ulang kalau lebih tua dari ini
var WEB_LATIHAN_SS_ = null;

function webLatihanSS_() {
  if (WEB_LATIHAN_SS_) return WEB_LATIHAN_SS_;
  const id = PropertiesService.getScriptProperties().getProperty(WEB_PROP_LATIHAN);
  if (!id) throw new Error('File latihan belum siap. Matikan lalu nyalakan lagi Mode latihan.');
  const asli = SpreadsheetApp.getActiveSpreadsheet().getId();
  if (id === asli) throw new Error('File latihan tidak boleh sama dengan database asli.');
  WEB_LATIHAN_SS_ = SpreadsheetApp.openById(id);
  return WEB_LATIHAN_SS_;
}

// dipanggil saat Mode latihan dinyalakan / tombol "Segarkan data latihan"
function webSiapkanLatihan_(a) {
  const props = PropertiesService.getScriptProperties();
  const perluBaru = function () {
    const waktu = Number(props.getProperty(WEB_PROP_LATIHAN_WAKTU)) || 0;
    return !props.getProperty(WEB_PROP_LATIHAN) || (Date.now() - waktu) > WEB_LATIHAN_UMUR_JAM_ * 3600 * 1000;
  };
  const segarkan = !!(a && a.segarkan);
  if (segarkan || perluBaru()) {
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(30000)) throw new Error('Database sedang sibuk, coba lagi sebentar.');
    try {
      if (segarkan || perluBaru()) { // dicek ulang: admin lain mungkin baru saja membuat salinan
        const idLama = props.getProperty(WEB_PROP_LATIHAN);
        const salinan = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId())
          .makeCopy('SANFELO LATIHAN WEB - BUKAN DATA ASLI');
        props.setProperty(WEB_PROP_LATIHAN, salinan.getId());
        props.setProperty(WEB_PROP_LATIHAN_WAKTU, String(Date.now()));
        WEB_LATIHAN_SS_ = null;
        if (idLama) { try { DriveApp.getFileById(idLama).setTrashed(true); } catch (err) { /* sudah tidak ada */ } }
      }
    } finally {
      lock.releaseLock();
    }
  }
  const w = Number(props.getProperty(WEB_PROP_LATIHAN_WAKTU)) || Date.now();
  return { waktu: Utilities.formatDate(new Date(w), WEB_TZ, 'dd/MM/yyyy HH:mm') };
}

// DATA_BERSIH & DISKON yang dipakai admin sehari-hari masih sheet LOKAL di inventory_SANFELO
// (salinan di database_SANFELO belum diperbarui), jadi web membaca dari inventory --
// sama dengan rumus PO_MASUK / SURAT JALAN / HARGATOKO di spreadsheet inventory.
const WEB_SHEET_MASTER_ = ['DATA_BERSIH', 'DISKON'];
// PERCEPAT: file inventory_SANFELO cukup dibuka SEKALI per permintaan (dulu dibuka ulang
// setiap kali DATA_BERSIH / DISKON / form dibaca -- tiap buka bisa 1-3 detik)
var WEB_INV_SS_ = null;
function webInventorySS_() {
  if (!WEB_INV_SS_) WEB_INV_SS_ = SpreadsheetApp.openById(INVENTORY_SS_ID_);
  return WEB_INV_SS_;
}
// nilai sel di form inventory (dipakai supaya nomor yang sedang tampil di form inventory tidak dipakai web)
function webNilaiFormInventory_(sheet, a1) {
  try {
    const sh = webInventorySS_().getSheetByName(sheet);
    return sh ? String(sh.getRange(a1).getValue() || '').trim() : '';
  } catch (err) {
    return '';
  }
}

function webMaster_(nama) {
  const sh = webInventorySS_().getSheetByName(nama); // INVENTORY_SS_ID_: GatewayTransaksi.js
  if (!sh) throw new Error('Sheet "' + nama + '" tidak ada di inventory_SANFELO.');
  return sh;
}

function webAman_(v) {
  if (v instanceof Date) return isNaN(v.getTime()) ? '' : v.toISOString();
  if (Array.isArray(v)) return v.map(webAman_);
  if (v && typeof v === 'object') {
    const o = {};
    Object.keys(v).forEach(function (k) { o[k] = webAman_(v[k]); });
    return o;
  }
  if (typeof v === 'number' && !isFinite(v)) return null;
  return v;
}

function webTgl_(v) {
  if (typeof v === 'string' && /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}/.test(v)) v = new Date(v); // tanggal dari gateway (JSON)
  if (v instanceof Date && !isNaN(v.getTime())) return Utilities.formatDate(v, WEB_TZ, 'dd/MM/yyyy');
  return String(v === null || v === undefined ? '' : v).trim();
}

function webBulanDari_(v) {
  if (v instanceof Date && !isNaN(v.getTime())) return Utilities.formatDate(v, WEB_TZ, 'yyyy-MM');
  const s = String(v || '').trim();
  let m = s.match(/^(\d{4})[-\/](\d{1,2})/);
  if (m) return m[1] + '-' + ('0' + m[2]).slice(-2);
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return m[3] + '-' + ('0' + m[2]).slice(-2);
  return '';
}

function webAngka_(v) {
  if (typeof v === 'number') return v;
  const n = Number(String(v || '').replace(/[^\d.-]/g, ''));
  return isNaN(n) ? 0 : n;
}

// semua baris data (mulai baris 2) selebar kolom terisi
function webBaca_(nama, numCols) {
  const sh = webSS_().getSheetByName(nama);
  if (!sh) throw new Error('Sheet "' + nama + '" tidak ada di database.');
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  const lebar = numCols || Math.max(1, sh.getLastColumn());
  return sh.getRange(2, 1, lastRow - 1, lebar).getValues();
}

// jalankan aksi gateway lewat jalur yang sama persis dengan inventory_SANFELO
var WEB_LOCK_DIPEGANG_ = false; // true selama kode salinan inventory sedang memegang kunci (lihat WEB_INV_SALINAN.js)

function webGateway_(action, payload) {
  const body = Object.assign({}, payload || {}, { token: SECRET_TOKEN, action: action });
  if (WEB_LATIHAN_) {
    // LOG_INVOICE_CETAK dibaca/ditulis gateway langsung ke database asli -> catat/batal hanya disimulasikan,
    // cek status boleh membaca data asli
    if (WEB_AKSI_ASLI_SAJA_.indexOf(action) !== -1) return webSimulasi_(action, payload || {});
    if (action !== 'cekStatusInvoiceCetak') {
      const handlerL = ACTION_HANDLERS[action];
      if (!handlerL) throw new Error('Action "' + action + '" tidak dikenali gateway ini.');
      if (WEB_AKSI_BACA_.indexOf(action) === -1) WEB_LATIHAN_CATAT_.push(action);
      const hl = handlerL(webLatihanSS_(), JSON.parse(JSON.stringify(body)));
      if (!hl || hl.ok === false) throw new Error((hl && hl.error) || ('Aksi ' + action + ' gagal.'));
      return hl;
    }
  }
  if (WEB_LOCK_DIPEGANG_) {
    // kunci sudah dipegang eksekusi ini -> jalankan aksi yang sama langsung (isi tetap lewat JSON)
    const handler = ACTION_HANDLERS[action];
    if (!handler) throw new Error('Action "' + action + '" tidak dikenali gateway ini.');
    const h = handler(webSS_(), JSON.parse(JSON.stringify(body)));
    if (!h || h.ok === false) throw new Error((h && h.error) || ('Aksi ' + action + ' gagal.'));
    return h;
  }
  const hasil = prosesRequestGateway_({ postData: { contents: JSON.stringify(body) } });
  if (!hasil || hasil.ok === false) throw new Error((hasil && hasil.error) || ('Aksi ' + action + ' gagal.'));
  return hasil;
}

function webCatatLog_(admin, aksi, ket) {
  WEB_LOG_DICATAT_ = true; // WEB_PERCEPAT.js: aksi ini sudah tercatat
  try { // di mode latihan masuk ke LOG_WEB file latihan
    const ss = webSS_();
    let sh = ss.getSheetByName(WEB_LOG_SHEET);
    if (!sh) {
      sh = ss.insertSheet(WEB_LOG_SHEET);
      sh.getRange(1, 1, 1, 4).setValues([['WAKTU', 'ADMIN', 'AKSI', 'KETERANGAN']]).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
    sh.appendRow([new Date(), admin, aksi, String(ket || '').slice(0, 500)]);
  } catch (err) {
    Logger.log('Log web gagal: ' + err);
  }
}

function webSimulasi_(action, payload) {
  const baris = Array.isArray(payload.rows) ? payload.rows.length : (Array.isArray(payload.items) ? payload.items.length : 0);
  WEB_LATIHAN_CATAT_.push(action + (baris ? ' (' + baris + ' baris)' : ''));
  if (action === 'getNomorPoMasukBersama') {
    // nomor yang AKAN didapat, tanpa memakai nomor asli
    const h = prosesRequestGateway_({ postData: { contents: JSON.stringify({ token: SECRET_TOKEN, action: 'peekNomorPoMasukBersama', asal: payload.asal }) } });
    return { ok: true, nomor: h.nomor, simulasi: true };
  }
  return { ok: true, simulasi: true, baris_ditulis: baris, mulai_dari_baris: null, insertedCount: 0, pbBaruLunas: [], jumlahPBDibayar: (payload.payments || []).length };
}

function webBukanLatihan_(apa) {
  if (WEB_LATIHAN_) throw new Error('Mode latihan: ' + apa + ' tidak bisa dicoba karena langsung mengubah spreadsheet inventory. Matikan mode latihan untuk menyimpan sungguhan.');
}

// a.admin = nama admin (kosong = semua), a.cari = teks; membaca 3000 baris terakhir, tampil maks 300
function webBacaLog_(a) {
  a = a || {};
  const sh = webSS_().getSheetByName(WEB_LOG_SHEET);
  if (!sh || sh.getLastRow() < 2) return [];
  const filterAdmin = String(a.admin || '').trim().toUpperCase();
  const cari = String(a.cari || '').trim().toUpperCase();
  const n = Math.min(3000, sh.getLastRow() - 1);
  const hasil = [];
  const rows = sh.getRange(sh.getLastRow() - n + 1, 1, n, 4).getValues().reverse();
  for (let i = 0; i < rows.length && hasil.length < 300; i++) {
    const r = rows[i];
    if (filterAdmin && String(r[1] || '').trim().toUpperCase() !== filterAdmin) continue;
    if (cari && (String(r[2]) + ' ' + String(r[3])).toUpperCase().indexOf(cari) === -1) continue;
    hasil.push({
      waktu: r[0] instanceof Date ? Utilities.formatDate(r[0], WEB_TZ, 'dd/MM/yyyy HH:mm') : String(r[0]),
      admin: r[1], aksi: r[2], ket: r[3]
    });
  }
  return hasil;
}

// hanya ADMIN (pemilik) yang boleh membuat / mengganti PIN admin lain (mis. NUR, SILVA)
function webAturPinAdmin_(a, admin) {
  if (admin !== 'ADMIN') throw new Error('Hanya akun ADMIN yang boleh mengatur PIN admin.');
  const nama = String(a.nama || '').trim().toUpperCase();
  const pin = String(a.pinBaru || '').trim();
  if (!/^[A-Z]{2,20}$/.test(nama) || nama === 'ADMIN') throw new Error('Nama admin hanya huruf (2-20), bukan ADMIN.');
  if (!/^\d{4,8}$/.test(pin)) throw new Error('PIN harus 4-8 angka.');
  const props = PropertiesService.getScriptProperties();
  const baru = !props.getProperty(WEB_PREFIX_PIN + nama);
  props.setProperty(WEB_PREFIX_PIN + nama, pin);
  CacheService.getScriptCache().remove('WEBGAGAL_' + nama);
  webCatatLog_(admin, baru ? 'TAMBAH ADMIN' : 'GANTI PIN ADMIN', nama);
  return { nama: nama, baru: baru };
}

// nama admin yang terdaftar (untuk filter Aktivitas Admin)
function webDaftarAdmin_() {
  return PropertiesService.getScriptProperties().getKeys()
    .filter(function (k) { return k.indexOf(WEB_PREFIX_PIN) === 0; })
    .map(function (k) { return k.slice(WEB_PREFIX_PIN.length); })
    .sort();
}

/* ---------------- riwayat pembayaran (format sama dengan inventory) ---------------- */

const WEB_ADJ_MIN_ = 3;
const WEB_ADJ_MAX_ = 60;

function webBersihkanKetAdj_(ket) {
  return String(ket || '')
    .replace(/[\r\n\t]+/g, ' ')
    .replace(/\s+-\s+/g, ' ')
    .replace(/[()]/g, ' ')
    .replace(/\s{2,}/g, ' ')
    .trim()
    .substring(0, WEB_ADJ_MAX_)
    .trim();
}

function webValidasiKetAdj_(metode, ketMentah) {
  if (metode !== 'ADJUSTMENT') return { ok: true, keterangan: '' };
  const ket = webBersihkanKetAdj_(ketMentah);
  if (ket.length < WEB_ADJ_MIN_) {
    return {
      ok: false, keterangan: '',
      message: 'Metode Adjustment wajib diisi keterangannya (minimal ' + WEB_ADJ_MIN_ +
        ' huruf), contoh: pembulatan, biaya admin, sparepart.'
    };
  }
  return { ok: true, keterangan: ket };
}

function webSusunMetode_(metode, ket) {
  return metode === 'ADJUSTMENT' && ket ? 'ADJUSTMENT (' + ket + ')' : metode;
}

function webPecahMetode_(teks) {
  const s = String(teks || '').trim();
  const m = s.match(/^ADJUSTMENT\s*\(([\s\S]*)\)\s*$/i);
  return m ? { metode: 'ADJUSTMENT', keterangan: m[1].trim() } : { metode: s, keterangan: '' };
}

function webParseRiwayat_(text) {
  const str = String(text || '').trim();
  if (!str) return null;
  const parts = str.split(' - ');
  if (parts.length < 2) return null;
  let idx = 0;
  let noPost = '';
  if (parts[0].trim().indexOf('POST-') === 0) {
    noPost = parts[0].trim();
    idx = 1;
  }
  const metodeTeks = parts.length > idx + 2 ? parts.slice(idx + 2).join(' - ').trim() : '';
  const pecah = webPecahMetode_(metodeTeks);
  return {
    noPost: noPost,
    tanggal: (parts[idx] || '').trim(),
    jumlah: Number(String(parts[idx + 1] || '').replace(/[^0-9]/g, '')) || 0,
    metode: pecah.metode,
    keterangan: pecah.keterangan
  };
}

// "yyyy-MM-dd" (input tanggal) -> "dd/MM/yyyy"
function webTglInput_(s) {
  const p = String(s || '').split('-');
  if (p.length !== 3 || !p[0] || !p[1] || !p[2]) return null;
  return p[2] + '/' + p[1] + '/' + p[0];
}

/* ---------------- dashboard ---------------- */

function webDashboard_() {
  const sekarang = new Date();
  const bulanIni = Utilities.formatDate(sekarang, WEB_TZ, 'yyyy-MM');
  const hariIni = new Date(Utilities.formatDate(sekarang, WEB_TZ, 'yyyy-MM-dd') + 'T00:00:00+07:00');

  const piutang = webDaftarPiutang_({});
  const totalPiutang = piutang.reduce(function (s, p) { return s + p.sisa; }, 0);

  let totalHutang = 0;
  let hutangLewat = 0;
  let hutangSegera = 0;
  const pbBelumLunas = [];
  webBaca_('HUTANG', 10).forEach(function (r) {
    const pb = String(r[1] || '').trim();
    if (!pb) return;
    const status = String(r[8] || '').trim().toUpperCase();
    const sisaRaw = r[9];
    const sisa = typeof sisaRaw === 'number' ? sisaRaw : (String(sisaRaw || '').toUpperCase().indexOf('LEBIH') === 0 ? 0 : webAngka_(sisaRaw));
    if (status === 'LUNAS' || sisa <= 0) return;
    totalHutang += sisa;
    const jt = r[6] instanceof Date ? r[6] : null;
    let hari = null;
    if (jt) {
      hari = Math.round((jt.getTime() - hariIni.getTime()) / 86400000);
      if (hari < 0) hutangLewat++;
      else if (hari <= 7) hutangSegera++;
    }
    pbBelumLunas.push({ pb: pb, supplier: String(r[4] || '').trim(), jatuhTempo: webTgl_(r[6]), hari: hari, sisa: sisa });
  });
  pbBelumLunas.sort(function (a, b) { return (a.hari === null ? 9999 : a.hari) - (b.hari === null ? 9999 : b.hari); });

  let jmlBulanIni = 0;
  let nilaiBulanIni = 0;
  const trx = webBaca_('TRANSAKSI', 17).filter(function (r) { return String(r[1] || '').trim(); });
  trx.forEach(function (r) {
    if (webBulanDari_(r[3]) === bulanIni) {
      jmlBulanIni++;
      nilaiBulanIni += webAngka_(r[12]);
    }
  });
  const terbaru = trx.slice(-8).reverse().map(function (r) {
    return { noSJ: r[1], noInv: r[2], tanggal: webTgl_(r[3]), customer: r[4], total: webAngka_(r[12]), status: r[13] };
  });

  // status toko dihitung seperti ringkasan G:H TEMPO TOKO di inventory
  const locked = webInv_().statusRingkasanToko_()
    .filter(function (x) { return x.status === 'LOCKED'; })
    .map(function (x) { return x.customer; });

  let ringkasPO = null;
  try { ringkasPO = webRingkasPO_(); } catch (err) { Logger.log('ringkasPO: ' + err); } // WEB_DOKUMEN.js

  return {
    bulanIni: bulanIni,
    totalPiutang: totalPiutang,
    jumlahPiutang: piutang.length,
    piutangTerlama: piutang.slice().sort(function (a, b) { return (a.tglMs || 0) - (b.tglMs || 0); }).slice(0, 6),
    totalHutang: totalHutang,
    jumlahPBHutang: pbBelumLunas.length,
    hutangLewat: hutangLewat,
    hutangSegera: hutangSegera,
    hutangTerdekat: pbBelumLunas.slice(0, 6),
    transaksiBulanIni: jmlBulanIni,
    nilaiBulanIni: nilaiBulanIni,
    transaksiTerbaru: terbaru,
    ringkasPO: ringkasPO,
    tokoLocked: locked
  };
}

/* ---------------- piutang & pembayaran (port Pembayaran.js) ---------------- */

function webDaftarPiutang_(a) {
  const cari = String(a.cari || '').trim().toUpperCase();
  const data = webBaca_('PENJUALAN');
  const hasil = [];
  let tokoTerakhir = '';
  data.forEach(function (r) {
    const customer = String(r[1] || '').trim();
    if (customer) tokoTerakhir = customer;
    const noInv = String(r[2] || '').trim();
    if (!noInv) return;
    const status = String(r[5] || '').trim().toUpperCase();
    if (status === 'LUNAS') return;
    const total = Number(r[4]) || 0;
    const sisa = status ? (Number(r[6]) || 0) : total;
    if (sisa <= 0) return;
    const toko = customer || tokoTerakhir;
    if (cari && (noInv + ' ' + toko).toUpperCase().indexOf(cari) === -1) return;
    hasil.push({
      noInv: noInv, toko: toko, tanggal: webTgl_(r[0]),
      tglMs: r[0] instanceof Date ? r[0].getTime() : 0,
      method: String(r[3] || '').trim(), total: total, sisa: sisa,
      status: status || 'BELUM BAYAR'
    });
  });
  return hasil;
}

function webCariDataInvoice_(noInvoice) {
  const trimmed = String(noInvoice || '').trim();
  if (!trimmed) return { found: false, message: 'No Invoice tidak boleh kosong.' };
  if (!/^(INV|SPRIN)-\d{4}-\d{2}-\d{3}$/.test(trimmed)) {
    return { found: false, message: 'Format No Invoice salah. Gunakan format: INV-2026-07-001' };
  }
  const data = webBaca_('PENJUALAN').map(function (r) {
    const row = r.slice();
    while (row.length < 10) row.push('');
    return row;
  });
  if (!data.length) return { found: false, message: 'Data PENJUALAN masih kosong.' };

  let idx = -1;
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][2] || '').trim() === trimmed) { idx = i; break; }
  }
  if (idx === -1) return { found: false, message: 'Invoice ' + trimmed + ' tidak ditemukan di PENJUALAN.' };

  let namaToko = String(data[idx][1] || '').trim();
  let scan = idx;
  while (!namaToko && scan > 0) {
    scan--;
    namaToko = String(data[scan][1] || '').trim();
  }
  const total = Number(data[idx][4]) || 0;
  const statusRaw = String(data[idx][5] || '').trim().toUpperCase();
  const sisaTagihan = statusRaw ? Number(data[idx][6]) || 0 : total;
  const riwayat = [];
  for (let c = 9; c < data[idx].length; c++) {
    const p = webParseRiwayat_(data[idx][c]);
    if (p) riwayat.push(p);
  }
  return {
    found: true,
    noInvoice: String(data[idx][2]).trim(),
    namaToko: namaToko,
    tanggal: webTgl_(data[idx][0]),
    method: String(data[idx][3] || '').trim(),
    totalTagihan: total,
    totalSudahDibayar: total - sisaTagihan,
    sisaTagihan: sisaTagihan,
    status: statusRaw === 'LUNAS' ? 'Lunas' : 'Cicilan',
    riwayatPembayaran: riwayat,
    ekspedisi: Number(data[idx][7]) || 0,
    palet: Number(data[idx][8]) || 0
  };
}

function webProsesPembayaran_(a, admin) {
  const noInvoice = String(a.noInvoice || '').trim();
  const jumlahBayar = Number(a.jumlahBayar);
  const metodeBayar = String(a.metodeBayar || '').trim().toUpperCase();
  const tglBayarText = webTglInput_(a.tglBayar);

  if (!jumlahBayar || jumlahBayar <= 0) return { success: false, message: 'Jumlah bayar tidak valid.' };
  if (['TUNAI', 'TRANSFER', 'ADJUSTMENT'].indexOf(metodeBayar) === -1) {
    return { success: false, message: 'Metode pembayaran harus TUNAI atau TRANSFER atau ADJUSTMENT.' };
  }
  const ketAdj = webValidasiKetAdj_(metodeBayar, a.keteranganAdjustment);
  if (!ketAdj.ok) return { success: false, message: ketAdj.message };
  if (!tglBayarText) return { success: false, message: 'Tanggal pembayaran tidak valid.' };

  // pengaman klik 2x: pembayaran dengan tanggal, nominal & metode sama persis sudah tercatat
  const sebelum = webCariDataInvoice_(noInvoice);
  if (!sebelum.found) return { success: false, message: sebelum.message };
  const metodeTeks = webSusunMetode_(metodeBayar, ketAdj.keterangan);
  const sudahAda = sebelum.riwayatPembayaran.some(function (r) {
    return r.tanggal === tglBayarText && r.jumlah === Math.round(jumlahBayar) &&
      webSusunMetode_(r.metode, r.keterangan).toUpperCase() === metodeTeks.toUpperCase();
  });
  if (sudahAda && !a.tetapSimpan) {
    return {
      success: false, perluKonfirmasi: true,
      message: 'Pembayaran ' + noInvoice + ' tanggal ' + tglBayarText + ' sebesar ' + Math.round(jumlahBayar) +
        ' (' + metodeTeks + ') SUDAH tercatat. Yakin mau mencatat lagi?'
    };
  }

  let hasil;
  try {
    hasil = webGateway_('prosesPembayaranPenjualan', {
      noInvoice: noInvoice,
      jumlahBayar: jumlahBayar,
      metodeBayar: metodeBayar,
      keteranganAdjustment: ketAdj.keterangan,
      tglBayarText: tglBayarText
    });
  } catch (err) {
    return { success: false, message: String(err.message || err) };
  }

  const catatan = [];
  if (hasil.jadiLunas) {
    try { webGateway_('tandaiDoneTransaksi', { noInvs: [hasil.noInvoice] }); }
    catch (err) { catatan.push('Tandai DONE transaksi gagal: ' + err.message); }
    try { webGateway_('tandaiDoneSparepart', { noInvoice: hasil.noInvoice }); }
    catch (err) { catatan.push('Tandai DONE sparepart gagal: ' + err.message); }
  }
  webCatatLog_(admin, 'BAYAR PIUTANG', noInvoice + ' | ' + tglBayarText + ' | ' + Math.round(jumlahBayar) + ' | ' +
    metodeTeks + ' | ' + (hasil.noPost || '') + (catatan.length ? ' | ' + catatan.join('; ') : ''));

  const sesudah = webCariDataInvoice_(noInvoice);
  sesudah.success = true;
  sesudah.noPost = hasil.noPost;
  sesudah.catatan = catatan;
  return sesudah;
}

/* ---------------- ubah tanggal / keterangan adjustment riwayat pembayaran ----------------
   Lewat aksi gateway "ubahRiwayatPembayaran" (GatewayEditPembayaran.js di project ini).
   Nominal, metode, sisa & status tidak berubah. Dicatat di LOG_EDIT_PEMBAYARAN + LOG_WEB. */
function webUbahRiwayatBayar_(a, admin) {
  const noInvoice = String(a.noInvoice || '').trim().toUpperCase();
  let h;
  try {
    h = webGateway_('ubahRiwayatPembayaran', {
      noInvoice: noInvoice, urutan: Number(a.urutan), noPost: String(a.noPost || ''),
      tanggalLama: String(a.tanggalLama || ''), jumlahLama: Number(a.jumlahLama) || 0,
      tglBaru: String(a.tglBaru || ''), keteranganBaru: String(a.keteranganBaru || ''),
      alasan: String(a.alasan || ''), oleh: 'WEB ' + admin
    });
  } catch (err) {
    const pesan = String(err.message || err);
    if (/tidak dikenali/.test(pesan)) {
      return { success: false, message: 'Fitur ubah riwayat belum dipasang di database: tambahkan GatewayEditPembayaran.js + 1 baris ubahRiwayatPembayaran di ACTION_HANDLERS.' };
    }
    return { success: false, message: pesan };
  }
  webCatatLog_(admin, 'UBAH RIWAYAT BAYAR', noInvoice + ' | ' + h.lama + ' -> ' + h.baru + ' | alasan: ' + String(a.alasan || ''));
  const d = webCariDataInvoice_(noInvoice);
  d.success = true;
  d.pesan = 'Riwayat ' + (h.noPost || '') + ' diubah menjadi: ' + h.baru;
  return d;
}

/* ---------------- hutang (port PembayaranHutang.js) ---------------- */

const WEB_HUTANG_ = { BULAN: 1, PB: 2, NOINV: 3, NOSJ: 4, SUPPLIER: 5, TGLBELI: 6, JT: 7, TOTAL: 8, STATUS: 9, SISA: 10, RIWAYAT: 11 };

function webNormBulan_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, WEB_TZ, 'yyyy-MM');
  const s = String(v || '').replace(/ /g, ' ').trim();
  const m = s.match(/^(\d{4})[-\/](\d{1,2})$/);
  return m ? m[1] + '-' + ('0' + m[2]).slice(-2) : s;
}

function webNormSupplier_(v) { return String(v || '').replace(/ /g, ' ').trim().toUpperCase(); }

function webSisaHutang_(raw) {
  if (typeof raw === 'number') return raw;
  const s = String(raw || '').trim();
  if (s.toUpperCase().indexOf('LEBIH') === 0) return 0;
  return Number(s) || 0;
}

function webDaftarBulanHutang_() {
  const ada = {};
  webBaca_('HUTANG', WEB_HUTANG_.BULAN).forEach(function (r) {
    const b = webNormBulan_(r[0]);
    if (/^\d{4}-\d{2}$/.test(b)) ada[b] = true;
  });
  return Object.keys(ada).sort().reverse();
}

function webSupplierListByBulan_(bulanInput) {
  const bulan = webNormBulan_(bulanInput);
  if (!bulan) return { success: false, message: 'Bulan tidak boleh kosong.', list: [] };
  const seen = {};
  const list = [];
  webBaca_('HUTANG', WEB_HUTANG_.SUPPLIER).forEach(function (r) {
    if (webNormBulan_(r[0]) !== bulan) return;
    const s = String(r[4] || '').trim();
    if (!s) return;
    const k = webNormSupplier_(s);
    if (!seen[k]) { seen[k] = true; list.push(s); }
  });
  list.sort(function (a, b) { return a.localeCompare(b); });
  return { success: true, bulan: bulan, list: list };
}

function webPBListByBulanSupplier_(bulanInput, supplierInput) {
  const bulan = webNormBulan_(bulanInput);
  const supplier = String(supplierInput || '').trim();
  if (!bulan) return { success: false, message: 'Bulan tidak boleh kosong.' };
  if (!supplier) return { success: false, message: 'Supplier tidak boleh kosong.' };
  const target = webNormSupplier_(supplier);
  const daftarPB = [];
  webBaca_('HUTANG', WEB_HUTANG_.SISA).forEach(function (r, i) {
    if (webNormBulan_(r[0]) !== bulan || webNormSupplier_(r[4]) !== target) return;
    const sisaRaw = r[9];
    daftarPB.push({
      rowIndex: i + 2,
      pb: String(r[1] || '').trim(),
      noInv: String(r[2] || '').trim() || '-',
      noSJ: String(r[3] || '').trim(),
      tglBeli: webTgl_(r[5]),
      jatuhTempo: webTgl_(r[6]),
      total: Number(r[7]) || 0,
      sisaBayar: webSisaHutang_(sisaRaw),
      isLebih: String(sisaRaw || '').toUpperCase().indexOf('LEBIH') === 0,
      status: String(r[8] || '').trim()
    });
  });
  if (!daftarPB.length) return { success: false, message: 'Tidak ada data hutang untuk ' + supplier + ' di bulan ' + bulan + '.' };
  const totalTagihan = daftarPB.reduce(function (s, p) { return s + p.total; }, 0);
  const sisaTagihan = daftarPB.reduce(function (s, p) { return s + p.sisaBayar; }, 0);
  return {
    success: true, bulan: bulan, supplier: supplier, daftarPB: daftarPB,
    totalTagihan: totalTagihan, sisaTagihan: sisaTagihan, status: sisaTagihan <= 0 ? 'Lunas' : 'Cicilan'
  };
}

function webProsesPembayaranMultiPB_(a, admin) {
  const bulan = webNormBulan_(a.bulan);
  const supplier = String(a.supplier || '').trim();
  const metodeBayar = String(a.metodeBayar || '').trim().toUpperCase();
  const tglBayarText = webTglInput_(a.tglBayar);
  const payments = Array.isArray(a.payments) ? a.payments : [];

  if (!bulan) return { success: false, message: 'Bulan tidak boleh kosong.' };
  if (!supplier) return { success: false, message: 'Supplier tidak boleh kosong.' };
  if (!tglBayarText) return { success: false, message: 'Tanggal bayar belum diisi / tidak valid.' };
  if (['TUNAI', 'TRANSFER', 'ADJUSTMENT'].indexOf(metodeBayar) === -1) {
    return { success: false, message: 'Metode pembayaran harus TUNAI, TRANSFER, atau ADJUSTMENT.' };
  }
  const ketAdj = webValidasiKetAdj_(metodeBayar, a.keteranganAdjustment);
  if (!ketAdj.ok) return { success: false, message: ketAdj.message };
  const metodeTeks = webSusunMetode_(metodeBayar, ketAdj.keterangan);
  if (!payments.length) return { success: false, message: 'Centang minimal 1 PB untuk dibayar.' };

  const paymentsPayload = payments.map(function (p) {
    return { rowIndex: Number(p.rowIndex), pb: String(p.pb || '').trim(), jumlahBayar: Number(p.jumlahBayar) };
  });

  // pengaman klik 2x (sama dengan cariPembayaranHutangDobel_ di inventory)
  const rows = webBaca_('HUTANG');
  const dobel = [];
  paymentsPayload.forEach(function (p) {
    const row = rows[p.rowIndex - 2];
    if (!row) return;
    const sudah = row.slice(WEB_HUTANG_.RIWAYAT - 1).some(function (cell) {
      const b = String(cell || '').split(' - ');
      if (b.length < 3) return false;
      return b[0].trim() === tglBayarText &&
        Number(String(b[1]).replace(/[^\d]/g, '')) === Math.round(p.jumlahBayar) &&
        b.slice(2).join(' - ').trim().toUpperCase() === metodeTeks.toUpperCase();
    });
    if (sudah) dobel.push(p.pb || ('baris ' + p.rowIndex));
  });
  if (dobel.length) {
    return {
      success: false,
      message: 'Pembayaran ini SUDAH TERCATAT sebelumnya (' + dobel.join(', ') + ', ' + tglBayarText + ', ' + metodeTeks + '). ' +
        'Kalau memang ada cicilan kedua dengan nominal & tanggal sama, ganti tanggal bayarnya.'
    };
  }

  let hasil;
  try {
    hasil = webGateway_('bayarHutangMultiPB', {
      bulan: bulan, supplier: supplier, tglBayarText: tglBayarText, metodeBayar: metodeBayar,
      keteranganAdjustment: ketAdj.keterangan, payments: paymentsPayload
    });
  } catch (err) {
    return { success: false, message: String(err.message || err) };
  }

  const catatan = [];
  const lunas = (hasil.pbBaruLunas || []).map(function (p) { return String(p || '').trim(); }).filter(Boolean);
  if (lunas.length) {
    try { webGateway_('tandaiLunasPembelianBatch', { pbList: lunas }); }
    catch (err) { catatan.push('Tandai LUNAS di PEMBELIAN gagal: ' + err.message); }
  }
  webCatatLog_(admin, 'BAYAR HUTANG', supplier + ' ' + bulan + ' | ' + tglBayarText + ' | ' + metodeTeks + ' | ' +
    paymentsPayload.map(function (p) { return p.pb + '=' + Math.round(p.jumlahBayar); }).join(', ') +
    (catatan.length ? ' | ' + catatan.join('; ') : ''));

  const sesudah = webPBListByBulanSupplier_(bulan, supplier);
  sesudah.success = true;
  sesudah.jumlahPBDibayar = hasil.jumlahPBDibayar;
  sesudah.pbBaruLunas = lunas;
  sesudah.catatan = catatan;
  return sesudah;
}

/* ---------------- data (baca saja) ---------------- */

// kolom: jumlah kolom yang ditampilkan; tgl: kolom tanggal untuk filter bulan; uang: kolom rupiah
const WEB_DATA_ = {
  'TRANSAKSI': { kolom: 17, tgl: 4, uang: [13] },
  'DETAIL_TRANSAKSI': { kolom: 20, tgl: 2, uang: [13, 14, 15, 18, 19] },
  'PENJUALAN': { kolom: 0, tgl: 1, uang: [5, 7, 8, 9] },
  'HISTORY_PO_MASUK': { kolom: 19, tgl: 3, uang: [13, 14], sembunyi: [17] },
  'HISTORY_PO_KELUAR': { kolom: 9, tgl: 2, uang: [9] },
  'PEMBELIAN': { kolom: 15, tgl: 5, uang: [10] },
  'HUTANG': { kolom: 0, tgl: 6, uang: [8, 10] },
  'TEMPO TOKO': { kolom: 5, tgl: 1, uang: [] },
  'SPAREPART': { kolom: 19, tgl: 4, uang: [14, 15, 18] },
  'EKSPEDISI': { kolom: 14, tgl: 4, uang: [12] },
  'LOG_INVOICE_CETAK': { kolom: 10, tgl: 2, uang: [4] },
  'DISKON': { kolom: 10, tgl: 0, uang: [], master: true, urutSheet: true },
  // kolom I (BARCODE) berisi gambar di dalam sel -> tidak bisa dibaca sebagai teks, disembunyikan
  'DATA_BERSIH': { kolom: 11, tgl: 0, uang: [4, 11], master: true, urutSheet: true, sembunyi: [9] }
};

function webDaftarData_(a) {
  const nama = String(a.sheet || '');
  const cfg = WEB_DATA_[nama];
  if (!cfg) throw new Error('Sheet tidak tersedia.');
  const sh = cfg.master ? webMaster_(nama) : webSS_().getSheetByName(nama); // DISKON & DATA_BERSIH dari inventory
  if (!sh) throw new Error('Sheet "' + nama + '" tidak ada di database.');
  const lebar = cfg.kolom || Math.max(1, sh.getLastColumn());
  const header = sh.getRange(1, 1, 1, lebar).getValues()[0].map(function (h, i) {
    return String(h || '').trim() || ('KOLOM ' + (i + 1));
  });
  const cari = String(a.cari || '').trim().toUpperCase();
  const bulan = String(a.bulan || '').trim();
  const semua = sh.getLastRow() < 2 ? [] : sh.getRange(2, 1, sh.getLastRow() - 1, lebar).getValues();

  let baris = [];
  semua.forEach(function (r, i) {
    // baris kosong = tidak ada isi selain kotak centang FALSE (mis. kolom BIAYA TAMBAH di TRANSAKSI)
    if (!r.some(function (v) { return v !== '' && v !== null && v !== false; })) return;
    const tampil = r.map(function (v) {
      if (v instanceof Date) return webTgl_(v);
      return (v && typeof v === 'object') ? '' : v; // gambar di dalam sel (CellImage) tidak dikirim ke web
    });
    if (bulan && cfg.tgl && webBulanDari_(r[cfg.tgl - 1]) !== bulan) return;
    if (cari && tampil.join(' ').toUpperCase().indexOf(cari) === -1) return;
    baris.push({ no: i + 2, v: tampil });
  });
  const urutSheet = !!cfg.urutSheet; // master: urutan sama dengan spreadsheet (atas ke bawah)
  if (!urutSheet) baris.reverse();

  const total = baris.length;
  const per = a.semua ? Math.max(1, total) : Math.min(Number(a.per) || 50, 200); // semua: halaman daftar web (cari & saring di browser)
  const halaman = a.semua ? 1 : Math.max(1, Number(a.halaman) || 1);
  baris = baris.slice((halaman - 1) * per, halaman * per);

  const sembunyi = cfg.sembunyi || [];
  return {
    sheet: nama, header: header, uang: cfg.uang, sembunyi: sembunyi, urutSheet: urutSheet,
    baris: baris, total: total, halaman: halaman, per: per
  };
}

/* ---------------- detail invoice (baca saja) ---------------- */

function webDetailInvoice_(noInvoice) {
  const noInv = String(noInvoice || '').trim().toUpperCase();
  if (!noInv) throw new Error('No invoice kosong.');

  const trx = webBaca_('TRANSAKSI', 17).filter(function (r) { return String(r[2] || '').trim().toUpperCase() === noInv; });
  if (!trx.length) throw new Error('Invoice ' + noInv + ' tidak ditemukan di TRANSAKSI.');
  const t = trx[trx.length - 1];

  const items = webBaca_('DETAIL_TRANSAKSI', 20)
    .filter(function (r) { return String(r[2] || '').trim().toUpperCase() === noInv; })
    .sort(function (a, b) { return (Number(a[15]) || 0) - (Number(b[15]) || 0); })
    .map(function (r) {
      return {
        kode: r[5], barang: r[6], qty: r[7], sat: r[8], ket: r[9],
        harga: webAngka_(r[12]), diskon: webAngka_(r[13]), net: webAngka_(r[14]), status: r[19]
      };
    });

  const bayar = webCariDataInvoice_(noInv);
  const logCetak = webBaca_('LOG_INVOICE_CETAK', 10).filter(function (r) {
    return String(r[0] || '').trim().toUpperCase() === noInv;
  })[0];

  return {
    noSJ: t[1], noInv: t[2], tanggal: webTgl_(t[3]), customer: t[4], alamat: t[5],
    tempo: webTgl_(t[6]), method: t[9], noPol: t[10], sales: t[11], total: webAngka_(t[12]),
    statusOrder: t[13], keterangan: t[15], noPO: t[16],
    items: items,
    pembayaran: bayar.found ? bayar : null,
    cetak: logCetak ? { status: logCetak[8], jumlah: logCetak[7], terakhir: webTgl_(logCetak[6]) } : null
  };
}
