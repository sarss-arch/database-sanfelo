/* ======================================================
   WEB_PO_MASUK.js   (PROJECT: database_SANFELO)
   ------------------------------------------------------
   Web admin: tambah customer (sheet DISKON) dan posting PO Masuk.

   PO Masuk dihitung PERSIS seperti form PO_MASUK di inventory_SANFELO:
     NAMA   = DATA_BERSIH kolom B (VLOOKUP kode)
     STATUS = STOK DATA_BERSIH kolom J >= qty ? "GUDANG" : "ORDER <kurang>"
     HARGA  = HARGA DATA_BERSIH kolom D x qty           (kolom F form)
     DISKON = HARGATOKO(harga, toko, method, diskon)     (kolom G form)
     TOTAL  = HARGA - DISKON                             (kolom H form)
   Nomor PO diambil dari penghitung BERSAMA SANFELO & SL (aksi
   getNomorPoMasukBersama), lalu disimpan lewat aksi syncPoMasuk --
   jalur yang sama dengan inventory_SANFELO.
   ====================================================== */

const WEB_PILIHAN_METHOD_ = ['CREDIT', 'CASH', 'DISPLAY'];
const WEB_PILIHAN_DISKON_ = ['', '3SET', '5SET', 'SALE', 'NODISC', 'DISPLAY'];

function webRapikanNama_(v) {
  return String(v === null || v === undefined ? '' : v).replace(/[\s ​]+/g, ' ').trim();
}

/* ---------------- data master ---------------- */

// total diskon bertingkat, sama dengan rumus kolom D (TOTAL) di DISKON: (1 - (1-35%)(1-5%)(1-5%)) x 100 = 41,3375
function webTotalDiskon_(teks) {
  const bagian = String(teks || '').replace(/\s/g, '').replace(/,/g, '.').split('+').filter(String);
  if (!bagian.length || bagian.some(function (x) { return isNaN(Number(x)); })) return '';
  const sisa = bagian.reduce(function (s, x) { return s * (1 - Number(x) / 100); }, 1);
  return Math.round((1 - sisa) * 100 * 1e6) / 1e6;
}

function webBacaDiskon_() {
  const sh = webMaster_('DISKON');
  const last = sh.getLastRow();
  return last < 2 ? [] : sh.getRange(2, 1, last - 1, 10).getValues();
}

function webBacaBarang_() {
  const sh = webMaster_('DATA_BERSIH');
  const last = sh.getLastRow();
  return (last < 2 ? [] : sh.getRange(2, 1, last - 1, 11).getValues())
    .filter(function (r) { return String(r[0] || '').trim(); })
    .map(function (r) {
      return {
        kode: String(r[0]).trim(), nama: String(r[1] || '').trim(), ukuran: r[2],
        harga: Number(r[3]) || 0, stok: Number(r[9]) || 0, ket: String(r[5] || '').trim()
      };
    });
}

function webMasterPO_() {
  const diskon = webBacaDiskon_();
  const customer = diskon
    .filter(function (r) { return String(r[0] || '').trim(); })
    .map(function (r) {
      return {
        nama: String(r[0]), wilayah: String(r[1] || '').trim(), diskon: String(r[2] || '').trim(),
        totalDiskon: Number(r[3]) || 0, metode: String(r[4] || '').trim(),
        alamat: String(r[5] || '').trim(), telp: String(r[6] || '').trim()
      };
    });
  const paket = diskon
    .filter(function (r) { return String(r[8] || '').trim() && String(r[9] || '').trim(); })
    .map(function (r) { return { diskon: String(r[8]).trim(), ket: String(r[9]).trim(), total: webTotalDiskon_(r[8]) }; });
  const unik = function (arr) { return arr.filter(function (v, i) { return v && arr.indexOf(v) === i; }).sort(); };
  let nomor = '';
  try { nomor = webGateway_('peekNomorPoMasukBersama', { asal: 'SANFELO' }).nomor; } catch (err) { nomor = ''; }
  return {
    customer: customer,
    barang: webBacaBarang_(),
    paketDiskon: paket,
    wilayah: unik(customer.map(function (c) { return c.wilayah; })),
    metodeBayar: unik(customer.map(function (c) { return c.metode; })),
    pilihanMethod: WEB_PILIHAN_METHOD_,
    pilihanDiskon: WEB_PILIHAN_DISKON_,
    nomorPerkiraan: nomor
  };
}

/* ---------------- tambah customer ---------------- */

function webTambahCustomer_(a, admin) {
  const nama = webRapikanNama_(a.nama).toUpperCase();
  const wilayah = webRapikanNama_(a.wilayah).toUpperCase();
  const diskon = String(a.diskon || '').replace(/\s/g, '').replace(/,/g, '.');
  const metode = webRapikanNama_(a.metode).toUpperCase();
  const alamat = webRapikanNama_(a.alamat);
  let telp = String(a.telp || '').replace(/\D/g, '');
  if (telp.indexOf('62') === 0 && telp.length >= 11) telp = '0' + telp.substring(2);
  if (telp && telp.charAt(0) !== '0') telp = '0' + telp;

  webBukanLatihan_('tambah customer');
  if (nama.length < 3) throw new Error('Nama toko minimal 3 huruf.');
  if (!/^\d+(\.\d+)?(\+\d+(\.\d+)?)*$/.test(diskon)) throw new Error('Format diskon salah. Contoh: 35+5+5');
  if (!metode) throw new Error('Metode bayar wajib diisi (mis. 30 HARI).');

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new Error('Database sedang sibuk, coba lagi sebentar.');
  try {
    const sh = webMaster_('DISKON'); // DISKON di inventory_SANFELO (yang dipakai rumus PO_MASUK)
    const data = webBacaDiskon_();
    const kembar = data.some(function (r) { return webRapikanNama_(r[0]).toUpperCase() === nama; });
    if (kembar) throw new Error('Toko "' + nama + '" sudah ada di DISKON.');

    const baris = barisTerakhirBerisiDiskon_(sh, 1) + 1; // GatewayDiskonDataBersih.js
    if (baris > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), 1);
    // kolom D TOTAL: rumus yang sama dengan baris lain (disalin dari baris 2)
    const rumusD = sh.getRange(2, 4).getFormulaR1C1();
    sh.getRange(baris, 7).setNumberFormat('@');
    sh.getRange(baris, 1, 1, 7).setValues([[nama, wilayah, diskon, '', metode, alamat, telp]]);
    if (rumusD) sh.getRange(baris, 4).setFormulaR1C1(rumusD);
    else sh.getRange(baris, 4).setValue(webTotalDiskon_(diskon)); // kolom TOTAL berisi angka (bukan rumus)
    // urut A-Z kolom A-G, sama dengan autoSortDiskon_ (onEdit DISKON di inventory)
    const lastRow = sh.getLastRow();
    if (lastRow >= 3) sh.getRange(2, 1, lastRow - 1, 7).sort({ column: 1, ascending: true });
    webCatatLog_(admin, 'TAMBAH CUSTOMER', nama + ' | ' + wilayah + ' | ' + diskon + ' | ' + metode);
  } finally {
    lock.releaseLock();
  }
  return { ok: true, nama: nama };
}

/* ---------------- hitung harga toko (port persis HARGATOKO di Kode.js) ---------------- */

function webHargaToko_(h, toko, met, disp, dataDiskon) {
  if (h === '' || h == null) return '';
  if (toko === '' || toko == null) return '';
  h = Number(h);
  const dispStr = disp ? disp.toString().trim().toUpperCase() : '';
  if (dispStr === 'SALE') return 0;
  if (dispStr === 'NODISC') return 0;

  const tokoStr = toko.toString().trim();
  let diskonRaw = null;
  for (let i = 0; i < dataDiskon.length; i++) {
    const n = dataDiskon[i][0];
    if (n !== '' && n !== null && n.toString().trim() === tokoStr) { diskonRaw = dataDiskon[i][2]; break; }
  }
  if (diskonRaw === null || diskonRaw === '') return '';

  const diskonArr = diskonRaw.toString().split('+');
  if (met && met.toString().trim().toUpperCase() === 'CASH') diskonArr.push('1,5');
  if (dispStr === 'DISPLAY') diskonArr.push('2,5');
  if (dispStr === '3SET') diskonArr.push('5');
  if (dispStr === '5SET') diskonArr.push('7,5');

  let sisaPersen = 1;
  for (let j = 0; j < diskonArr.length; j++) {
    const d = parseFloat(diskonArr[j].toString().trim().replace(',', '.'));
    sisaPersen *= (1 - d / 100);
  }
  const totalDiskonPersen = Math.round((1 - sisaPersen) * 100 * 10000 + 1e-9) / 10000;
  return h * (totalDiskonPersen / 100);
}

/* ---------------- hapus customer ---------------- */

// Hapus satu customer dari DISKON: hanya kolom A-G baris itu yang dihapus (baris di bawahnya naik),
// tabel paket diskon di kolom I-J tidak ikut bergeser. Data lama (PO, transaksi, piutang) tidak diubah.
function webHapusCustomer_(a, admin) {
  const nama = webRapikanNama_(a.nama).toUpperCase();
  const baris = Number(a.baris);
  if (!nama || !(baris >= 2)) throw new Error('Customer belum dipilih.');
  webBukanLatihan_('hapus customer');
  const sh = webMaster_('DISKON');
  const cek = function () {
    if (baris > sh.getLastRow() || webRapikanNama_(sh.getRange(baris, 1).getValue()).toUpperCase() !== nama) {
      throw new Error('Data DISKON sudah berubah (baris ' + baris + ' bukan ' + nama + '). Muat ulang halaman lalu coba lagi.');
    }
  };
  cek();
  // pemakaian customer ini di data (hanya untuk diberitahukan ke admin)
  const sama = function (v) { return webRapikanNama_(v).toUpperCase() === nama; };
  const po = {};
  webBaca_('HISTORY_PO_MASUK', 4).forEach(function (r) { if (String(r[0]).trim().toUpperCase() === 'MASUK' && sama(r[3])) po[String(r[1]).trim()] = true; });
  const trx = webBaca_('TRANSAKSI', 5).filter(function (r) { return r[1] && sama(r[4]); }).length;
  const piutang = webDaftarPiutang_({}).filter(function (p) { return sama(p.toko); });
  const sisa = piutang.reduce(function (s, p) { return s + p.sisa; }, 0);
  if (!a.yakin) {
    return {
      perluKonfirmasi: true, nama: nama, po: Object.keys(po).length, transaksi: trx, piutang: piutang.length, sisaPiutang: sisa,
      pesan: 'Hapus ' + nama + ' dari sheet DISKON? Customer ini punya ' + Object.keys(po).length + ' PO, ' + trx + ' transaksi' +
        (piutang.length ? ', dan ' + piutang.length + ' invoice belum lunas' : '') + '. Data lama tidak ikut terhapus, ' +
        'tapi toko ini tidak bisa dipilih lagi di PO Masuk dan diskonnya tidak terbaca untuk PO baru.'
    };
  }
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new Error('Spreadsheet sedang sibuk, coba lagi sebentar.');
  try {
    cek();
    sh.getRange(baris, 1, 1, 7).deleteCells(SpreadsheetApp.Dimension.ROWS);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  webCatatLog_(admin, 'HAPUS CUSTOMER', nama + ' | ' + Object.keys(po).length + ' PO | ' + trx + ' transaksi | sisa piutang ' + Math.round(sisa));
  return { ok: true, nama: nama };
}

/* ---------------- PO Masuk ---------------- */

// tanggal "yyyy-MM-dd" -> Date jam 00:00 WIB (sama dengan tanggal yang diketik di form)
function webTanggalDariInput_(s) {
  const p = String(s || '').split('-');
  if (p.length !== 3) return '';
  return new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]));
}

function webHitungPO_(a) {
  const customer = String(a.customer || '');
  const method = String(a.method || '').trim().toUpperCase();
  const diskonPO = String(a.diskon || '').trim().toUpperCase();
  const dataDiskon = webBacaDiskon_();
  const barang = {};
  webBacaBarang_().forEach(function (b) { barang[b.kode.toUpperCase()] = b; });

  const salah = [];
  const baris = [];
  (a.items || []).forEach(function (it, i) {
    const kode = String(it.kode || '').trim().toUpperCase();
    if (!kode) return;
    const qty = Number(it.qty);
    const b = barang[kode];
    if (!b) { salah.push('Baris ' + (i + 1) + ': kode ' + kode + ' tidak ada di DATA_BERSIH.'); return; }
    if (!(qty > 0)) { salah.push('Baris ' + (i + 1) + ': qty ' + kode + ' belum diisi.'); return; }
    // harga satuan khusus: hanya dikirim web saat ✏️ Ubah PO (harga PO lama dipertahankan)
    const satuanKhusus = it.hargaSatuan === '' || it.hargaSatuan == null ? null : Number(it.hargaSatuan);
    const harga = qty * (satuanKhusus !== null && !isNaN(satuanKhusus) ? satuanKhusus : b.harga);
    const diskon = webHargaToko_(harga, customer, method, diskonPO, dataDiskon);
    baris.push({
      line: baris.length + 1, kode: b.kode, nama: b.nama, qty: qty,
      status: b.stok >= qty ? 'GUDANG' : 'ORDER ' + Math.abs(b.stok - qty),
      harga: harga, diskon: diskon, total: diskon === '' ? '' : harga - diskon
    });
  });
  const tokoAda = dataDiskon.some(function (r) { return String(r[0] || '').trim() === customer.trim(); });
  if (customer && !tokoAda) salah.push('Customer "' + customer + '" tidak ada di DISKON.');
  if (customer && baris.some(function (b) { return b.diskon === ''; }) && diskonPO !== 'SALE' && diskonPO !== 'NODISC') {
    salah.push('Diskon toko ' + customer + ' kosong di sheet DISKON.');
  }
  const total = baris.reduce(function (s, b) { return s + (Number(b.total) || 0); }, 0);
  return { baris: baris, total: total, salah: salah };
}

function webSimpanPOMasuk_(a, admin) {
  const customer = String(a.customer || '').trim();
  const tglPO = webTanggalDariInput_(a.tglPO);
  const tglKirim = a.tglKirim ? webTanggalDariInput_(a.tglKirim) : '';
  const method = String(a.method || '').trim().toUpperCase();
  const diskonPO = String(a.diskon || '').trim().toUpperCase();

  if (!customer) throw new Error('Customer belum dipilih.');
  if (!tglPO) throw new Error('Tanggal PO belum diisi.');
  if (WEB_PILIHAN_METHOD_.indexOf(method) === -1) throw new Error('Method harus ' + WEB_PILIHAN_METHOD_.join(' / ') + '.');
  if (WEB_PILIHAN_DISKON_.indexOf(diskonPO) === -1) throw new Error('Pilihan diskon tidak dikenal.');

  const hitung = webHitungPO_(a);
  if (hitung.salah.length) throw new Error(hitung.salah.join('\n'));
  if (!hitung.baris.length) throw new Error('Belum ada barang.');

  // alamat & telp dari DISKON (sama dengan rumus B3/B4 di form PO_MASUK)
  let alamat = '';
  let telp = '';
  webBacaDiskon_().some(function (r) {
    if (String(r[0] || '').trim().toUpperCase() === customer.toUpperCase()) { alamat = r[5]; telp = r[6]; return true; }
    return false;
  });

  const timestamp = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy HH:mm:ss');
  const poNo = webGateway_('getNomorPoMasukBersama', { asal: 'SANFELO' }).nomor; // klaim nomor bersama SANFELO & SL
  const rows = hitung.baris.map(function (b) {
    return ['MASUK', poNo, tglPO, customer, diskonPO, alamat, telp, tglKirim, b.kode, b.nama, b.qty,
      b.status, b.harga, b.total, timestamp, b.line, '', method];
  });
  const hasil = webGateway_('syncPoMasuk', { jenis: 'MASUK', poNo: poNo, rows: rows });
  webCatatLog_(admin, 'PO MASUK', poNo + ' | ' + customer + ' | ' + rows.length + ' barang | ' + Math.round(hitung.total));
  return { poNo: poNo, barisDitulis: hasil.baris_ditulis, total: hitung.total, baris: hitung.baris };
}

/**
 * Tulis hasil ✏️ Ubah PO (WEB_PO_UBAH.js) lewat aksi gateway syncPoMasuk (jalur yang sama dengan
 * simpan PO), lalu warna status & kolom S (barang CANCELLED) di HISTORY_PO_MASUK database.
 * susunFn dijalankan DI DALAM kunci yang sama dengan gateway.
 */
function webTulisUbahPO_(poNo, susunFn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new Error('Database sedang sibuk, coba lagi sebentar.');
  WEB_LOCK_DIPEGANG_ = true; // webGateway_ menjalankan aksi langsung (kunci sudah dipegang)
  try {
    const s = susunFn();
    const hasil = webGateway_('syncPoMasuk', { jenis: 'MASUK', poNo: poNo, rows: s.rows });
    const hist = webSS_().getSheetByName('HISTORY_PO_MASUK');
    const startRow = Number(hasil.mulai_dari_baris) || 0;
    if (hist && startRow >= 2) {
      const lebar = Math.max(hist.getLastColumn(), 19);
      hist.getRange(startRow, 1, s.rows.length, lebar).setBackground('#ffffff');
      s.susulan.forEach(function (x) {
        const b = startRow + x.offset;
        hist.getRange(b, 1, 1, lebar).setBackground(x.warna);
        if (x.cancelled) hist.getRange(b, 19).setValue(x.backup || '');
      });
    }
    SpreadsheetApp.flush();
    return { barisDitulis: s.rows.length };
  } finally {
    WEB_LOCK_DIPEGANG_ = false;
    lock.releaseLock();
  }
}
