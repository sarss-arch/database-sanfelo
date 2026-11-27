/* ======================================================
   WEB_PO_UBAH.js   (dipakai web SL & web SANFELO, isi file sama)
   ------------------------------------------------------
   ✏️ Ubah PO Masuk yang sudah tersimpan, dari web.
   Aturannya sama dengan EDIT dari HISTORY_PO_MASUK di Inventory_SL
   (syncToHistoryInternal_):
   - Nomor PO TETAP (tidak mengambil nomor baru).
   - LINE stabil per barang: barang yang sudah ada memakai LINE lamanya,
     jadi Surat Jalan / invoice yang menunjuk PO + LINE tetap benar.
   - Status "DONE SJ" / "DONE SJ n" / "CANCELLED" per barang dibawa ke baris baru.
   - Barang yang SUDAH dikirim tidak bisa dihapus (dipertahankan) dan
     qty-nya tidak boleh di bawah qty yang sudah dikirim.
   - Harga barang lama tetap memakai harga PO lama (kecuali diminta pakai
     price list terbaru).
   - Sebelum menulis, isi PO di sheet dicek sama persis dengan saat dibuka
     (sidik). Kalau sudah berubah (mis. baru dibuat SJ), simpan ditolak.
   - Isi lama & baru dicatat di LOG_WEB.
   Penulisan ke sheet ada di webTulisUbahPO_ (WEB_PO_MASUK.js tiap web).
   ====================================================== */

const WEB_HPO_ = { JENIS: 0, PO: 1, TGL: 2, CUST: 3, DISKON: 4, ALAMAT: 5, TELP: 6, KIRIM: 7, KODE: 8, NAMA: 9, QTY: 10, STATUS: 11, HARGA: 12, TOTAL: 13, WAKTU: 14, LINE: 15, AKSI: 16, METHOD: 17, BACKUP: 18 };

function webTglIsoPO_(v) {
  if (v instanceof Date && !isNaN(v.getTime())) return Utilities.formatDate(v, WEB_TZ, 'yyyy-MM-dd');
  const s = String(v || '').trim();
  let m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
  m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  return m ? m[0] : '';
}

// jumlah yang sudah dikirim dari teks status (sama dengan parseDoneQty_ SL)
function webQtyDonePO_(status, qty) {
  const s = String(status || '').trim().toUpperCase();
  if (s.indexOf('DONE SJ') !== 0) return 0;
  const n = Number(s.replace('DONE SJ', '').trim());
  return n > 0 ? n : (Number(qty) || 0);
}

// semua baris HISTORY_PO_MASUK untuk 1 PO: [{ baris, line, kode, qty, status, backup, nilai }]
function webBacaLamaPO_(poNo) {
  const sh = webSS_().getSheetByName('HISTORY_PO_MASUK');
  if (!sh) throw new Error('Sheet "HISTORY_PO_MASUK" tidak ada.');
  const last = sh.getLastRow();
  const list = [];
  if (last < 2) return list;
  const target = String(poNo).trim();
  sh.getRange(2, 1, last - 1, 19).getValues().forEach(function (r, i) {
    if (String(r[0] || '').trim().toUpperCase() !== 'MASUK' || String(r[1] || '').trim() !== target) return;
    list.push({
      baris: i + 2, line: String(r[WEB_HPO_.LINE]).trim(), kode: String(r[WEB_HPO_.KODE] || '').trim().toUpperCase(),
      qty: Number(r[WEB_HPO_.QTY]) || 0, status: String(r[WEB_HPO_.STATUS] || '').trim().toUpperCase(), backup: r[WEB_HPO_.BACKUP], nilai: r
    });
  });
  return list;
}

// sidik isi PO (tanpa kolom Q ACTION & nomor baris) untuk mendeteksi perubahan sejak dibuka
function webSidikPO_(lamaList) {
  const isi = lamaList.map(function (l) {
    return l.nilai.map(function (v, i) { return i === WEB_HPO_.AKSI ? '' : (v instanceof Date ? v.getTime() : String(v)); }).join('\u0001');
  }).sort().join('\u0002');
  return Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, isi, Utilities.Charset.UTF_8));
}

/** Buka PO tersimpan untuk diubah di form web. a = { poNo } */
function webBukaPO_(a) {
  const poNo = String(a.poNo || '').trim();
  if (!poNo) throw new Error('Isi nomor PO.');
  const lama = webBacaLamaPO_(poNo);
  if (!lama.length) throw new Error('PO ' + poNo + ' tidak ditemukan di HISTORY_PO_MASUK.');
  if (lama.every(function (l) { return l.status === 'CANCELLED'; })) {
    throw new Error('PO ' + poNo + ' sudah di-CANCEL. Aktifkan dulu lewat ACTION "BATAL CANCEL" di HISTORY_PO_MASUK.');
  }
  const h = lama[0].nilai;
  const hargaList = {};
  webBacaBarang_().forEach(function (b) { const k = b.kode.toUpperCase(); if (hargaList[k] === undefined) hargaList[k] = Number(b.harga) || 0; });
  const items = lama.slice().sort(function (x, y) { return (Number(x.line) || 0) - (Number(y.line) || 0); }).map(function (l) {
    const harga = Number(l.nilai[WEB_HPO_.HARGA]);
    const satuan = l.qty > 0 && !isNaN(harga) ? harga / l.qty : null;
    const beda = satuan !== null && hargaList[l.kode] !== undefined && Math.abs(satuan - hargaList[l.kode]) >= 1;
    return {
      line: l.line, kode: l.kode, qty: l.qty, status: l.status, qtyDone: webQtyDonePO_(l.status, l.qty),
      hargaSatuan: beda ? satuan : '', hargaList: hargaList[l.kode] === undefined ? '' : hargaList[l.kode]
    };
  });
  return {
    poNo: poNo, customer: String(h[WEB_HPO_.CUST] || '').trim(), tglPO: webTglIsoPO_(h[WEB_HPO_.TGL]), tglKirim: webTglIsoPO_(h[WEB_HPO_.KIRIM]),
    diskon: String(h[WEB_HPO_.DISKON] || '').trim().toUpperCase(), method: String(h[WEB_HPO_.METHOD] || '').trim().toUpperCase(),
    items: items, sidik: webSidikPO_(lama)
  };
}

/**
 * Port tetapkanLineStabilPO_ + bawa status lama + pertahankan barang terkirim (syncToHistoryInternal_ SL).
 * rows: baris baru 18 kolom (LINE = urutan form). Hasil: { rows, susulan, dipertahankan }
 */
function webSusunUbahPO_(rows, lamaList, poNo) {
  const kodeOf = function (r) { return String(r[WEB_HPO_.KODE] || '').trim().toUpperCase(); };
  const terpakai = {}, lamaTerpakai = {}, selesai = [];
  rows.forEach(function (r, i) {
    const line = String(r[WEB_HPO_.LINE]).trim();
    if (lamaList.some(function (l) { return l.line === line && l.kode === kodeOf(r); }) && !terpakai[line]) {
      terpakai[line] = true; lamaTerpakai[line] = true; selesai[i] = true;
    }
  });
  rows.forEach(function (r, i) {
    if (selesai[i]) return;
    const lama = lamaList.filter(function (l) { return l.kode === kodeOf(r) && !lamaTerpakai[l.line] && !terpakai[l.line]; })[0];
    if (!lama) return;
    r[WEB_HPO_.LINE] = isNaN(Number(lama.line)) ? lama.line : Number(lama.line);
    terpakai[lama.line] = true; lamaTerpakai[lama.line] = true; selesai[i] = true;
  });
  let maks = 0;
  lamaList.forEach(function (l) { maks = Math.max(maks, Number(l.line) || 0); });
  rows.forEach(function (r) { maks = Math.max(maks, Number(r[WEB_HPO_.LINE]) || 0); });
  const lineLama = {};
  lamaList.forEach(function (l) { lineLama[l.line] = true; });
  rows.forEach(function (r, i) {
    if (selesai[i]) return;
    const line = String(r[WEB_HPO_.LINE]).trim();
    if (line && !terpakai[line] && !lineLama[line]) { terpakai[line] = true; return; }
    maks++; r[WEB_HPO_.LINE] = maks; terpakai[String(maks)] = true;
  });

  const statusLama = {};
  lamaList.forEach(function (l) { if (!statusLama[l.line]) statusLama[l.line] = l; });
  const lineBaru = {};
  rows.forEach(function (r) { lineBaru[String(r[WEB_HPO_.LINE]).trim()] = true; });
  const susulan = [];
  const salah = [];
  rows.forEach(function (r, offset) {
    const lama = statusLama[String(r[WEB_HPO_.LINE]).trim()];
    if (!lama || lama.kode !== kodeOf(r)) return;
    if (lama.status === 'CANCELLED') {
      r[WEB_HPO_.STATUS] = 'CANCELLED';
      susulan.push({ offset: offset, warna: '#f4cccc', backup: lama.backup, cancelled: true });
      return;
    }
    const qtyDone = webQtyDonePO_(lama.status, lama.qty);
    if (!(qtyDone > 0)) return;
    const qtyBaru = Number(r[WEB_HPO_.QTY]) || 0;
    if (qtyBaru < qtyDone) { salah.push(lama.kode + ': qty ' + qtyBaru + ' lebih kecil dari yang sudah dikirim (' + qtyDone + ').'); return; }
    if (qtyDone >= qtyBaru) { r[WEB_HPO_.STATUS] = 'DONE SJ'; susulan.push({ offset: offset, warna: '#d9ead3' }); }
    else { r[WEB_HPO_.STATUS] = 'DONE SJ ' + qtyDone; susulan.push({ offset: offset, warna: '#fff2cc' }); }
  });
  if (salah.length) throw new Error('Qty tidak boleh di bawah yang sudah dikirim:\n' + salah.join('\n') + '\nKoreksi dulu Surat Jalan / invoice-nya.');

  const dipertahankan = [];
  lamaList.forEach(function (l) {
    if (lineBaru[l.line]) return;
    if (!(webQtyDonePO_(l.status, l.qty) > 0)) return;
    if (dipertahankan.some(function (d) { return d.line === l.line; })) return;
    dipertahankan.push(l);
  });
  dipertahankan.forEach(function (l) {
    const r = l.nilai.slice(0, 18);
    r[WEB_HPO_.PO] = poNo;
    r[WEB_HPO_.AKSI] = '';
    susulan.push({ offset: rows.length, warna: l.status === 'DONE SJ' ? '#d9ead3' : '#fff2cc' });
    rows.push(r);
  });
  return { rows: rows, susulan: susulan, dipertahankan: dipertahankan };
}

/** Simpan perubahan PO. a = isian simpanPO + { poNo, sidik } */
function webUbahPOMasuk_(a, admin) {
  const poNo = String(a.poNo || '').trim();
  if (!poNo) throw new Error('Nomor PO tidak ada.');
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
  if (!hitung.baris.length) throw new Error('Belum ada barang. Untuk membatalkan PO pakai CANCEL di HISTORY_PO_MASUK.');

  let alamat = '', telp = '';
  webBacaDiskon_().some(function (r) {
    if (String(r[0] || '').trim().toUpperCase() === customer.toUpperCase()) { alamat = r[5]; telp = r[6]; return true; }
    return false;
  });
  const timestamp = Utilities.formatDate(new Date(), WEB_TZ, 'dd/MM/yyyy HH:mm:ss');
  const rows = hitung.baris.map(function (b) {
    return ['MASUK', poNo, tglPO, customer, diskonPO, alamat, telp, tglKirim, b.kode, b.nama, b.qty,
      b.status, b.harga, b.total, timestamp, b.line, '', method];
  });

  let ringkasLama = '', susun = null;
  // dikerjakan di dalam kunci tiap web: baca ulang -> cek sidik -> susun -> tulis
  const tulis = webTulisUbahPO_(poNo, function () {
    const lama = webBacaLamaPO_(poNo);
    if (!lama.length) throw new Error('PO ' + poNo + ' tidak ditemukan lagi di HISTORY_PO_MASUK.');
    if (webSidikPO_(lama) !== String(a.sidik || '')) {
      throw new Error('PO ' + poNo + ' baru saja berubah (mis. dibuat Surat Jalan atau diedit admin lain). Buka ulang PO-nya lalu ubah lagi.');
    }
    if (lama.every(function (l) { return l.status === 'CANCELLED'; })) throw new Error('PO ' + poNo + ' sudah di-CANCEL.');
    ringkasLama = lama.map(function (l) { return l.kode + ' x' + l.qty + (l.status ? ' ' + l.status : ''); }).join(', ');
    susun = webSusunUbahPO_(rows, lama, poNo);
    return susun;
  });

  const pesan = [];
  if (susun.dipertahankan.length) {
    pesan.push('Barang yang sudah dikirim tetap ada di PO: ' + susun.dipertahankan.map(function (l) { return l.kode + ' (' + l.status + ')'; }).join(', ') + '.');
  }
  webCatatLog_(admin, 'UBAH PO MASUK', poNo + ' | ' + customer + ' | lama: ' + ringkasLama + ' | baru: ' +
    susun.rows.map(function (r) { return r[WEB_HPO_.KODE] + ' x' + r[WEB_HPO_.QTY]; }).join(', ') + ' | total ' + Math.round(hitung.total));
  return { poNo: poNo, barisDitulis: tulis.barisDitulis, total: hitung.total, pesan: pesan };
}
