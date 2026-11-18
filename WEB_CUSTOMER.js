/* ======================================================
   WEB_CUSTOMER.js   (dipakai web SL & web SANFELO, isi file sama)
   ------------------------------------------------------
   CRUD customer dari web:
   - Tambah & hapus customer: sudah ada (WEB_PO_MASUK.js).
   - UBAH customer (sheet DISKON): wilayah, diskon, metode, alamat, telp,
     dan NAMA. Kalau nama diganti, boleh sekalian mengganti nama itu di
     SEMUA data lama (PO, transaksi, detail, piutang, tempo, log cetak,
     sparepart, ekspedisi, PO keluar, toko asal pembelian).
   - GANTI CUSTOMER SATU INVOICE: invoice dipindah ke customer lain
     (mis. INV-2026-09-035 MEBEL MAS AGUNG -> HOKI MAPAN) di semua sheet
     yang memuat no invoice itu, plus PO sumbernya kalau PO itu hanya
     dipakai invoice ini. Sesudahnya TEMPO TOKO disinkron ulang dengan
     fungsi yang sama dengan Simpan Draft.

   PENGAMAN DATA:
   - Kolom dicari dari JUDUL kolom (baris 1), bukan nomor kolom, jadi
     aman untuk susunan sheet SL maupun SANFELO.
   - Hanya sel nama/alamat yang cocok yang ditulis; sel berisi rumus
     tidak pernah ditimpa. Nominal, tanggal, pembayaran tidak disentuh.
   - Dua langkah: pratinjau (tidak menulis apa pun) -> admin setuju ->
     tulis. Semua dicatat di LOG_WEB dengan nama admin.
   ====================================================== */

// ganti customer SATU invoice: sheet yang punya kolom NO INV
const WEB_CUST_PER_INV_ = [
  { sheet: 'TRANSAKSI', nama: 'CUSTOMER', alamat: 'ALAMAT' },
  { sheet: 'DETAIL_TRANSAKSI', nama: 'NAMA TOKO' },
  { sheet: 'PENJUALAN', nama: 'CUSTOMER' },
  { sheet: 'TEMPO TOKO', nama: 'CUSTOMER', maks: 5 }, // A..E (G:H = ringkasan, disusun ulang oleh sinkron)
  { sheet: 'LOG_INVOICE_CETAK', nama: 'CUSTOMER' },
  { sheet: 'SPAREPART', nama: 'NAMA TOKO', alamat: 'ALAMAT' },
  { sheet: 'EKSPEDISI', nama: 'NAMA TOKO', alamat: 'ALAMAT' }
];

// ganti NAMA customer di semua data lama (kolom nama saja; alamat lama tetap)
const WEB_CUST_SEMUA_ = [
  { sheet: 'HISTORY_PO_MASUK', nama: 'CUSTOMER' },
  { sheet: 'TRANSAKSI', nama: 'CUSTOMER' },
  { sheet: 'DETAIL_TRANSAKSI', nama: 'NAMA TOKO' },
  { sheet: 'PENJUALAN', nama: 'CUSTOMER' },
  { sheet: 'TEMPO TOKO', nama: 'CUSTOMER' },
  { sheet: 'LOG_INVOICE_CETAK', nama: 'CUSTOMER' },
  { sheet: 'LOG_INVOICE_RIWAYAT', nama: 'CUSTOMER' },
  { sheet: 'SPAREPART', nama: 'NAMA TOKO' },
  { sheet: 'EKSPEDISI', nama: 'NAMA TOKO' },
  { sheet: 'HISTORY_PO_KELUAR', nama: 'NAMA TOKO' },
  { sheet: 'PEMBELIAN', nama: 'KETERANGAN' } // toko asal barang (isi persis nama toko)
];

function webCustNorm_(v) { return String(v === null || v === undefined ? '' : v).replace(/ /g, ' ').replace(/\s+/g, ' ').trim().toUpperCase(); }

// nomor kolom (1-based) yang judulnya = judul; maks = batas kolom
function webCustKolom_(sh, judul, maks) {
  const lebar = Math.min(sh.getLastColumn(), maks || 1000);
  if (lebar < 1) return [];
  const h = sh.getRange(1, 1, 1, lebar).getValues()[0];
  const hasil = [];
  h.forEach(function (x, i) { if (webCustNorm_(x) === judul) hasil.push(i + 1); });
  return hasil;
}

function webCustBacaKolom_(sh, kol) {
  const n = sh.getLastRow() - 1;
  return n < 1 ? [] : sh.getRange(2, kol, n, 1).getValues().map(function (r) { return r[0]; });
}

// tulis sel-sel (baris -> nilai) di satu kolom; sel berumus dilewati. Mengembalikan jumlah sel ditulis.
function webCustTulis_(sh, kol, ubah) {
  const baris = Object.keys(ubah).map(Number).sort(function (a, b) { return a - b; });
  let ditulis = 0;
  let i = 0;
  while (i < baris.length) {
    let j = i;
    while (j + 1 < baris.length && baris[j + 1] === baris[j] + 1) j++;
    const awal = baris[i], n = j - i + 1;
    const rg = sh.getRange(awal, kol, n, 1);
    const rumus = rg.getFormulas();
    if (rumus.every(function (f) { return !f[0]; })) {
      rg.setValues(baris.slice(i, j + 1).map(function (b) { return [ubah[b]]; }));
      ditulis += n;
    } else {
      for (let k = 0; k < n; k++) if (!rumus[k][0]) { sh.getRange(awal + k, kol).setValue(ubah[awal + k]); ditulis++; }
    }
    i = j + 1;
  }
  return ditulis;
}

function webCustSheet_(nama) { return webSS_().getSheetByName(nama); }

// kode inventory yang dipakai web (SANFELO: webInv_, SL: webSl_)
function webCustKode_() {
  if (typeof webInv_ === 'function') return webInv_();
  return webSl_();
}

// TEMPO TOKO (termasuk ringkasan LOCKED/UNLOCKED per toko) disusun ulang seperti sesudah Simpan Draft
function webCustSinkronTempo_(catatan) {
  try {
    const kode = webCustKode_();
    if (typeof webJalankanSl_ === 'function' && typeof webInv_ !== 'function') webJalankanSl_(function () { kode.syncTempoToko_(false); });
    else kode.syncTempoToko_(false);
  } catch (err) {
    catatan.push('Sinkron TEMPO TOKO gagal: ' + (err.message || err) + ' (jalankan Sinkron Tempo Toko di menu Alat).');
  }
}

/* ======================================================
   GANTI CUSTOMER SATU INVOICE
   a = { noInv, customer, alamat, yakin }
   ====================================================== */
function webGantiCustomerInvoice_(a, admin) {
  const noInv = webCustNorm_(a.noInv);
  const baru = webCustNorm_(a.customer);
  const alamatBaru = String(a.alamat || '').replace(/\s+/g, ' ').trim();
  if (!noInv) throw new Error('No invoice kosong.');
  if (baru.length < 3) throw new Error('Pilih customer baru.');
  // customer baru harus ada di DISKON (supaya diskon, metode & tempo tokonya terbaca)
  const diskon = webBacaDiskon_();
  const rowCust = diskon.filter(function (r) { return webCustNorm_(r[0]) === baru; })[0];
  if (!rowCust) throw new Error('Customer "' + baru + '" belum ada di DISKON. Tambahkan dulu lewat ＋ Customer baru.');
  const namaBaru = String(rowCust[0]).trim(); // tulisan persis seperti di DISKON

  const rencana = webCustRencanaInvoice_(noInv, namaBaru, alamatBaru);
  if (!rencana.ditemukan) throw new Error('Invoice ' + noInv + ' tidak ditemukan di TRANSAKSI / SPAREPART.');
  if (!rencana.tugas.length) throw new Error('Invoice ' + noInv + ' sudah atas nama ' + namaBaru + (alamatBaru ? ' dengan alamat itu' : '') + '. Tidak ada yang berubah.');
  if (!a.yakin) return webCustRingkasRencana_(rencana, noInv, namaBaru, alamatBaru);

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('Database sedang sibuk, coba lagi sebentar.');
  let total = 0;
  const per = [];
  try {
    // dihitung ulang di dalam kunci (data terbaru)
    const r2 = webCustRencanaInvoice_(noInv, namaBaru, alamatBaru);
    r2.tugas.forEach(function (t) {
      const sh = webCustSheet_(t.sheet);
      const n = webCustTulis_(sh, t.kol, t.ubah);
      if (n) { total += n; per.push(t.sheet + ' ' + t.jenis + ' ' + n); }
    });
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  const catatan = [];
  webCustSinkronTempo_(catatan);
  webCatatLog_(admin, 'GANTI CUSTOMER INVOICE', noInv + ' | ' + rencana.lama.join('/') + ' -> ' + namaBaru +
    (alamatBaru ? ' | alamat: ' + alamatBaru : '') + ' | ' + per.join(', ') + (rencana.poDilewati.length ? ' | PO tidak diubah: ' + rencana.poDilewati.join(', ') : '') +
    (catatan.length ? ' | ' + catatan.join('; ') : ''));
  return { ok: true, noInv: noInv, customer: namaBaru, sel: total, rincian: per, catatan: catatan, poDilewati: rencana.poDilewati };
}

// susun daftar sel yang akan diubah (tidak menulis apa pun)
function webCustRencanaInvoice_(noInv, namaBaru, alamatBaru) {
  const tugas = [];
  const lama = {};
  let ditemukan = false;
  let adaAlamat = false;
  let noPO = '';
  WEB_CUST_PER_INV_.forEach(function (cfg) {
    const sh = webCustSheet_(cfg.sheet);
    if (!sh || sh.getLastRow() < 2) return;
    const kInv = webCustKolom_(sh, 'NO INV', cfg.maks)[0];
    const kNama = webCustKolom_(sh, cfg.nama, cfg.maks)[0];
    if (!kInv || !kNama) return;
    const inv = webCustBacaKolom_(sh, kInv);
    const nama = webCustBacaKolom_(sh, kNama);
    const ubahNama = {};
    let efektif = '';
    for (let i = 0; i < inv.length; i++) {
      if (String(nama[i] || '').trim()) efektif = String(nama[i]).trim();
      if (webCustNorm_(inv[i]) !== noInv) continue;
      if (cfg.sheet === 'TRANSAKSI' || cfg.sheet === 'SPAREPART') ditemukan = true;
      if (efektif) lama[efektif] = true;
      if (String(nama[i] || '').trim() !== namaBaru) ubahNama[i + 2] = namaBaru;
      // PENJUALAN bergaya grup (nama hanya di baris pertama): baris di bawahnya tetap milik customer lama
      if (cfg.sheet === 'PENJUALAN' && i + 1 < inv.length && String(inv[i + 1] || '').trim() && !String(nama[i + 1] || '').trim() && efektif) {
        ubahNama[i + 3] = efektif;
      }
    }
    if (Object.keys(ubahNama).length) tugas.push({ sheet: cfg.sheet, kol: kNama, ubah: ubahNama, jenis: 'nama' });
    if (cfg.alamat && alamatBaru) {
      const kAl = webCustKolom_(sh, cfg.alamat, cfg.maks)[0];
      if (kAl) {
        const al = webCustBacaKolom_(sh, kAl);
        const ubahAl = {};
        for (let i = 0; i < inv.length; i++) {
          if (webCustNorm_(inv[i]) === noInv && String(al[i] || '').trim() !== alamatBaru) ubahAl[i + 2] = alamatBaru;
        }
        if (Object.keys(ubahAl).length) { adaAlamat = true; tugas.push({ sheet: cfg.sheet, kol: kAl, ubah: ubahAl, jenis: 'alamat' }); }
      }
    }
    if (cfg.sheet === 'TRANSAKSI') {
      const kPO = webCustKolom_(sh, 'NO PO')[0];
      if (kPO) {
        const po = webCustBacaKolom_(sh, kPO);
        for (let i = 0; i < inv.length; i++) if (webCustNorm_(inv[i]) === noInv && String(po[i] || '').trim()) noPO = String(po[i]).trim();
        // PO yang juga dipakai invoice lain tidak diubah
        const dipakaiLain = {};
        for (let i = 0; i < inv.length; i++) {
          if (webCustNorm_(inv[i]) === noInv) continue;
          String(po[i] || '').split(',').forEach(function (p) { p = p.trim(); if (p) dipakaiLain[p] = true; });
        }
        tugas.poList = String(noPO).split(',').map(function (p) { return p.trim(); }).filter(Boolean);
        tugas.poLain = dipakaiLain;
      }
    }
  });

  // PO sumber (HISTORY_PO_MASUK): ikut diganti hanya kalau PO itu khusus invoice ini
  const poDiganti = [];
  const poDilewati = [];
  const poList = tugas.poList || [];
  const poLain = tugas.poLain || {};
  const shPO = webCustSheet_('HISTORY_PO_MASUK');
  if (poList.length && shPO && shPO.getLastRow() >= 2) {
    const kNo = webCustKolom_(shPO, 'NO PO')[0];
    const kNama = webCustKolom_(shPO, 'CUSTOMER')[0];
    const kAl = webCustKolom_(shPO, 'ALAMAT')[0];
    if (kNo && kNama) {
      const no = webCustBacaKolom_(shPO, kNo);
      const nm = webCustBacaKolom_(shPO, kNama);
      const al = kAl ? webCustBacaKolom_(shPO, kAl) : [];
      const ubahNama = {}, ubahAl = {};
      poList.forEach(function (p) {
        if (poLain[p]) { poDilewati.push(p + ' (dipakai invoice lain)'); return; }
        let ada = false;
        for (let i = 0; i < no.length; i++) {
          if (String(no[i] || '').trim() !== p) continue;
          ada = true;
          if (String(nm[i] || '').trim() !== namaBaru) ubahNama[i + 2] = namaBaru;
          if (kAl && alamatBaru && String(al[i] || '').trim() !== alamatBaru) ubahAl[i + 2] = alamatBaru;
        }
        if (ada) poDiganti.push(p); else poDilewati.push(p + ' (tidak ditemukan)');
      });
      if (Object.keys(ubahNama).length) tugas.push({ sheet: 'HISTORY_PO_MASUK', kol: kNama, ubah: ubahNama, jenis: 'nama' });
      if (Object.keys(ubahAl).length) tugas.push({ sheet: 'HISTORY_PO_MASUK', kol: kAl, ubah: ubahAl, jenis: 'alamat' });
    }
  }
  return { tugas: tugas, lama: Object.keys(lama), ditemukan: ditemukan, adaAlamat: adaAlamat, poDiganti: poDiganti, poDilewati: poDilewati };
}

function webCustRingkasRencana_(r, noInv, namaBaru, alamatBaru) {
  const per = {};
  r.tugas.forEach(function (t) {
    const k = t.sheet;
    per[k] = per[k] || { sheet: k, nama: 0, alamat: 0 };
    per[k][t.jenis] += Object.keys(t.ubah).length;
  });
  return {
    perluKonfirmasi: true, noInv: noInv, lama: r.lama, baru: namaBaru, alamat: alamatBaru,
    rincian: Object.keys(per).map(function (k) { return per[k]; }),
    poDiganti: r.poDiganti, poDilewati: r.poDilewati
  };
}

/* ======================================================
   UBAH CUSTOMER (sheet DISKON) + ganti nama di semua data lama
   a = { baris, nama (lama), baru: { nama, wilayah, diskon, metode, alamat, telp }, gantiSemua, yakin }
   ====================================================== */
function webUbahCustomer_(a, admin) {
  webBukanLatihan_('ubah customer');
  const lamaNama = webCustNorm_(a.nama);
  const baris = Number(a.baris);
  if (!lamaNama || !(baris >= 2)) throw new Error('Customer belum dipilih.');
  const b = a.baru || {};
  const nama = webRapikanNama_(b.nama).toUpperCase();
  const wilayah = webRapikanNama_(b.wilayah).toUpperCase();
  const diskon = String(b.diskon || '').replace(/\s/g, '').replace(/,/g, '.');
  const metode = webRapikanNama_(b.metode).toUpperCase();
  const alamat = webRapikanNama_(b.alamat);
  let telp = String(b.telp || '').replace(/\D/g, '');
  if (telp.indexOf('62') === 0 && telp.length >= 11) telp = '0' + telp.substring(2);
  if (telp && telp.charAt(0) !== '0') telp = '0' + telp;
  if (nama.length < 3) throw new Error('Nama toko minimal 3 huruf.');
  if (!/^\d+(\.\d+)?(\+\d+(\.\d+)?)*$/.test(diskon)) throw new Error('Format diskon salah. Contoh: 35+5+5');
  if (!metode) throw new Error('Metode bayar wajib diisi (mis. 30 HARI).');
  const gantiNama = nama !== lamaNama;
  const gantiSemua = gantiNama && !!a.gantiSemua;

  const sh = webMaster_('DISKON');
  const cek = function () {
    if (baris > sh.getLastRow() || webCustNorm_(sh.getRange(baris, 1).getValue()) !== lamaNama) {
      throw new Error('Data DISKON sudah berubah (baris ' + baris + ' bukan ' + lamaNama + '). Muat ulang lalu coba lagi.');
    }
    if (gantiNama && webBacaDiskon_().some(function (r, i) { return i + 2 !== baris && webCustNorm_(r[0]) === nama; })) {
      throw new Error('Nama "' + nama + '" sudah dipakai customer lain di DISKON.');
    }
  };
  cek();
  const rencana = gantiSemua ? webCustRencanaSemua_(lamaNama, nama) : [];
  if (!a.yakin) {
    return {
      perluKonfirmasi: true, lama: lamaNama, baru: nama, gantiNama: gantiNama, gantiSemua: gantiSemua,
      rincian: rencana.map(function (t) { return { sheet: t.sheet, nama: Object.keys(t.ubah).length }; })
    };
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('Database sedang sibuk, coba lagi sebentar.');
  const per = [];
  try {
    cek();
    const lamaBaris = sh.getRange(baris, 1, 1, 7).getValues()[0];
    sh.getRange(baris, 7).setNumberFormat('@');
    sh.getRange(baris, 1, 1, 3).setValues([[nama, wilayah, diskon]]);
    sh.getRange(baris, 5, 1, 3).setValues([[metode, alamat, telp]]);
    if (!sh.getRange(baris, 4).getFormula()) sh.getRange(baris, 4).setValue(webTotalDiskon_(diskon));
    per.push('DISKON: ' + lamaBaris.slice(0, 7).join(' / ') + ' -> ' + [nama, wilayah, diskon, '', metode, alamat, telp].join(' / '));
    // salinan DISKON di file database (SANFELO) ikut disamakan kalau ada
    const db = webSS_().getSheetByName('DISKON');
    if (db && db.getParent().getId() !== sh.getParent().getId() && db.getLastRow() >= 2) {
      const nm = webCustBacaKolom_(db, 1);
      for (let i = 0; i < nm.length; i++) {
        if (webCustNorm_(nm[i]) !== lamaNama) continue;
        db.getRange(i + 2, 7).setNumberFormat('@');
        db.getRange(i + 2, 1, 1, 3).setValues([[nama, wilayah, diskon]]);
        db.getRange(i + 2, 5, 1, 3).setValues([[metode, alamat, telp]]);
        if (!db.getRange(i + 2, 4).getFormula()) db.getRange(i + 2, 4).setValue(webTotalDiskon_(diskon));
        per.push('DISKON database');
      }
    }
    if (gantiNama) webCustUrutDiskon_(sh);
    if (gantiSemua) {
      webCustRencanaSemua_(lamaNama, nama).forEach(function (t) {
        const n = webCustTulis_(webCustSheet_(t.sheet), t.kol, t.ubah);
        if (n) per.push(t.sheet + ' ' + n);
      });
    }
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  const catatan = [];
  if (gantiSemua) webCustSinkronTempo_(catatan);
  webCatatLog_(admin, 'UBAH CUSTOMER', lamaNama + (gantiNama ? ' -> ' + nama : '') + ' | ' + per.join(' | ') + (catatan.length ? ' | ' + catatan.join('; ') : ''));
  return { ok: true, nama: nama, rincian: per, catatan: catatan };
}

function webCustRencanaSemua_(lamaNama, namaBaru) {
  const tugas = [];
  WEB_CUST_SEMUA_.forEach(function (cfg) {
    const sh = webCustSheet_(cfg.sheet);
    if (!sh || sh.getLastRow() < 2) return;
    webCustKolom_(sh, cfg.nama).forEach(function (kol) {
      const v = webCustBacaKolom_(sh, kol);
      const ubah = {};
      v.forEach(function (x, i) { if (webCustNorm_(x) === lamaNama) ubah[i + 2] = namaBaru; });
      if (Object.keys(ubah).length) tugas.push({ sheet: cfg.sheet, kol: kol, ubah: ubah });
    });
  });
  return tugas;
}

// urut A-Z seperti waktu tambah customer
function webCustUrutDiskon_(sh) {
  if (typeof webSl_ === 'function') { webSl_().autoSortDiskon_(sh); return; }
  const lastRow = sh.getLastRow();
  if (lastRow >= 3) sh.getRange(2, 1, lastRow - 1, 7).sort({ column: 1, ascending: true });
}
