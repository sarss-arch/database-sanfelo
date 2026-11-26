/* ======================================================
   WEB_PERCEPAT.js   (dipakai web SL & web SANFELO, isi file sama)
   ------------------------------------------------------
   Supaya web tidak membaca ulang sheet besar setiap kali halaman dibuka.

   CARA KERJA (data TIDAK pernah diubah oleh file ini):
   - Hasil aksi BACA yang berat (dashboard, master PO, master TT, stok, ...)
     disimpan sementara di CacheService (memori Google, bukan sheet).
   - Setiap aksi TULIS dari web (simpan PO, SJ, bayar, petty cash, ...)
     menaikkan "versi data" -> semua simpanan lama otomatis tidak dipakai,
     jadi sesudah simpan web selalu membaca data terbaru.
   - Perubahan langsung di spreadsheet tetap terbaca paling lambat sesuai
     umur simpanan (WEB_CACHE_TTL_), atau segera dengan tombol 🔄 Segarkan.
   - Mode latihan tidak memakai simpanan sama sekali.
   - Kalau CacheService bermasalah, web jalan seperti biasa (baca langsung).
   ====================================================== */

// umur simpanan per aksi (detik)
const WEB_CACHE_TTL_ = {
  dashboard: 180,
  masterPO: 600,
  masterTT: 300,
  stok: 300,
  tempoTokoDaftar: 300,
  daftarPOTerbuka: 180,
  bulanHutang: 900,
  daftarPalet: 300,
  laporan: 300,
  groupingDaftar: 300
};

// aksi yang hanya membaca (tidak menaikkan versi data). Aksi lain dianggap menulis.
const WEB_AKSI_BACA_SAJA_ = {
  dashboard: 1, daftarData: 1, daftarPiutang: 1, cariInvoice: 1, supplierHutang: 1, daftarPBHutang: 1,
  bulanHutang: 1, detailInvoice: 1, logWeb: 1, masterPO: 1, hitungPO: 1, daftarPOTerbuka: 1, muatPOSJ: 1,
  pratinjauSJ: 1, masterTT: 1, muatPOKeluarTT: 1, hitungTT: 1, invoiceUntukBiaya: 1, daftarPalet: 1,
  detailPalet: 1, pratinjauSparepart: 1, stok: 1, daftarKas: 1, laporan: 1, tempoTokoDaftar: 1,
  susulanKandidat: 1, statusSJ: 1, groupingDaftar: 1, poKeluarMuat: 1, poKeluarDariPOMasuk: 1,
  bukaSJ: 1, bukaTT: 1, bukaSpr: 1, siapkanLatihan: 1, daftarAdmin: 1, cekHutang: 1, versiData: 1
};

const WEB_CACHE_POTONG_ = 90000; // batas 1 nilai CacheService = 100 KB
const WEB_CACHE_MAKS_POTONG_ = 60;

var WEB_LOG_DICATAT_ = false; // true kalau aksi ini sudah menulis LOG_WEB sendiri

function webCacheVersi_() {
  const c = CacheService.getScriptCache();
  let v = c.get('WEBC_VERSI');
  if (!v) {
    v = String(Date.now());
    c.put('WEBC_VERSI', v, 21600);
  }
  return v;
}

function webCacheNaikkanVersi_() {
  try {
    CacheService.getScriptCache().put('WEBC_VERSI', Date.now() + Math.random().toString(36).slice(2, 6), 21600);
  } catch (err) {
    Logger.log('Versi cache gagal dinaikkan: ' + err);
  }
}

function webCacheKunci_(aksi, args) {
  const bersih = {};
  Object.keys(args || {}).sort().forEach(function (k) {
    if (['latihan', 'jawabYa', 'segar'].indexOf(k) === -1) bersih[k] = args[k];
  });
  const d = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, aksi + '|' + JSON.stringify(bersih), Utilities.Charset.UTF_8);
  return 'WEBC_' + webCacheVersi_() + '_' + Utilities.base64EncodeWebSafe(d).replace(/=+$/, '');
}

function webCacheAmbil_(kunci) {
  const c = CacheService.getScriptCache();
  const n = Number(c.get(kunci + '_n'));
  if (!n) return null;
  const ks = [];
  for (let i = 0; i < n; i++) ks.push(kunci + '_' + i);
  const m = c.getAll(ks);
  let s = '';
  for (let i = 0; i < n; i++) {
    if (m[ks[i]] === undefined || m[ks[i]] === null) return null; // sebagian sudah hilang -> baca ulang
    s += m[ks[i]];
  }
  const blob = Utilities.newBlob(Utilities.base64Decode(s), 'application/x-gzip', 'c.gz');
  return JSON.parse(Utilities.ungzip(blob).getDataAsString('UTF-8'));
}

function webCacheSimpan_(kunci, data, ttl) {
  const json = JSON.stringify(data);
  const b64 = Utilities.base64Encode(Utilities.gzip(Utilities.newBlob(json, 'application/json', 'c.json')).getBytes());
  const n = Math.ceil(b64.length / WEB_CACHE_POTONG_);
  if (!n || n > WEB_CACHE_MAKS_POTONG_) return;
  const o = {};
  for (let i = 0; i < n; i++) o[kunci + '_' + i] = b64.substr(i * WEB_CACHE_POTONG_, WEB_CACHE_POTONG_);
  o[kunci + '_n'] = String(n);
  CacheService.getScriptCache().putAll(o, ttl);
}

/**
 * Menjalankan satu aksi web dengan simpanan sementara (aksi baca berat) dan
 * menaikkan versi data sesudah aksi tulis. Dipanggil dari webApi.
 */
function webJalankanAksi_(aksi, fn, args, admin) {
  WEB_LOG_DICATAT_ = false;
  const ttl = WEB_LATIHAN_ ? 0 : WEB_CACHE_TTL_[aksi];
  if (ttl) {
    let kunci = null;
    try {
      kunci = webCacheKunci_(aksi, args);
      if (!args.segar) {
        const ada = webCacheAmbil_(kunci);
        if (ada !== null) return ada;
      }
    } catch (err) {
      Logger.log('Cache baca gagal (' + aksi + '): ' + err);
    }
    const hasil = webAman_(fn(args, admin));
    try { if (kunci) webCacheSimpan_(kunci, hasil, ttl); } catch (err) { Logger.log('Cache simpan gagal (' + aksi + '): ' + err); }
    return hasil;
  }
  const tulis = !WEB_AKSI_BACA_SAJA_[aksi];
  try {
    const hasil = fn(args, admin);
    // jejak admin: aksi tulis yang belum mencatat sendiri tetap masuk LOG_WEB
    if (tulis && !WEB_LOG_DICATAT_ && !(['hapusCustomer', 'ubahCustomer', 'gantiCustomerInvoice'].indexOf(aksi) !== -1 && !args.yakin)) webCatatLog_(admin, String(aksi).toUpperCase(), webRingkasArgs_(args));
    return hasil;
  } finally {
    // tetap dinaikkan walau gagal di tengah jalan: bisa saja sebagian sudah tertulis
    if (tulis && !WEB_LATIHAN_) webCacheNaikkanVersi_();
  }
}

/**
 * Cek ringan untuk web (dipanggil tiap ±45 detik dari browser yang sedang terbuka).
 * Mengembalikan "versi data". Versi naik bila:
 *  - ada aksi tulis dari web / gateway (webCacheNaikkanVersi_), atau
 *  - jumlah baris salah satu sheet di database berubah (mis. baris baru dari popup spreadsheet).
 * Hanya membaca jumlah baris tiap sheet, tidak membaca isi dan tidak menulis apa pun ke sheet.
 */
function webVersiData_() {
  const c = CacheService.getScriptCache();
  try {
    const sidik = webSS_().getSheets().map(function (sh) { return sh.getLastRow(); }).join('.');
    const lama = c.get('WEBC_SIDIK');
    if (lama !== sidik) {
      if (lama) webCacheNaikkanVersi_(); // ada baris baru/terhapus -> simpanan server dibuang juga
      c.put('WEBC_SIDIK', sidik, 21600);
    }
  } catch (err) {
    Logger.log('versiData: ' + err);
  }
  return { v: webCacheVersi_() };
}

// ringkasan isian untuk LOG_WEB (tanpa password / PIN)
function webRingkasArgs_(args) {
  const o = {};
  Object.keys(args || {}).forEach(function (k) {
    if (['latihan', 'jawabYa', 'segar'].indexOf(k) !== -1) return;
    if (/pass|pin|token/i.test(k)) { o[k] = '***'; return; }
    o[k] = args[k];
  });
  let s = '';
  try { s = JSON.stringify(o); } catch (err) { s = ''; }
  return s.length > 450 ? s.slice(0, 447) + '...' : s;
}
