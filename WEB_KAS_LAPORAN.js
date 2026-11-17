/* ======================================================
   WEB_KAS_LAPORAN.js   (PROJECT: database_SANFELO)
   ------------------------------------------------------
   Web admin: Stok Barang, Petty Cash, Buku Bank, Laporan Invoice.
   - Stok Barang  : fungsi popup Stok Barang inventory (ambilDataStokBarang).
   - Petty Cash   : fungsi popup Petty Cash inventory (processTransactionBatch),
                    yang memang menulis langsung ke PETTY CASH di database.
   - Buku Bank    : baris baru ditambahkan di bawah BUKU BANK database dengan
                    rumus SALDO yang sama dengan baris di atasnya.
   - Laporan      : hitungan yang sama dengan tab LAPORAN INVOICE inventory
                    (LOG_INVOICE_CETAK + PO pending dari HISTORY_PO_MASUK).
   Inventory tetap melihat semuanya lewat salinan IMPORTRANGE seperti biasa.
   ====================================================== */

function webStok_() {
  return webInv_().ambilDataStokBarang();
}

/* ---------------- PETTY CASH ---------------- */

const WEB_KAS_KATEGORI_ = {
  'PETTY CASH': [['PIUTANG_USAHA', 'Piutang usaha', 'DEBET'], ['LAINNYA', 'Lainnya', 'DEBET'],
    ['ONGKOS_KIRIM', 'Ongkos kirim', 'KREDIT'], ['OPERASIONAL', 'Operasional', 'KREDIT'], ['ADMIN', 'Admin', 'KREDIT']],
  'BUKU BANK': [['PIUTANG_USAHA', 'Piutang usaha', 'DEBET'], ['LAINNYA', 'Lainnya', 'DEBET'],
    ['HUTANG_USAHA', 'Hutang usaha', 'KREDIT'], ['OPERASIONAL', 'Operasional', 'KREDIT'], ['ADMIN', 'Admin', 'KREDIT']]
};
const WEB_KAS_MULAI_ = { 'PETTY CASH': 4, 'BUKU BANK': 3 };

function webDaftarKas_(a) {
  const nama = a.sheet === 'BUKU BANK' ? 'BUKU BANK' : 'PETTY CASH';
  const sh = webSS_().getSheetByName(nama);
  const mulai = WEB_KAS_MULAI_[nama];
  const last = sh.getLastRow();
  const kat = WEB_KAS_KATEGORI_[nama];
  const baris = [];
  let tglTerakhir = '';
  if (last >= mulai) {
    sh.getRange(mulai, 1, last - mulai + 1, 9).getValues().forEach(function (r, i) {
      if (!r[1] && r[7] === '') return;
      if (r[0] instanceof Date) tglTerakhir = webTgl_(r[0]);
      let kategori = '';
      let nominal = 0;
      for (let c = 2; c <= 6; c++) {
        if (r[c] !== '' && r[c] !== null) { kategori = kat[c - 2][1]; nominal = Number(r[c]) || 0; }
      }
      baris.push({ no: mulai + i, tanggal: r[0] instanceof Date ? webTgl_(r[0]) : '', tglGrup: tglTerakhir, ket: r[1],
        kategori: kategori, debet: c2Debet_(kat, r), kredit: c2Kredit_(kat, r), saldo: Number(r[7]) || 0, nominal: nominal });
    });
  }
  const bulan = String(a.bulan || '');
  const tampil = bulan ? baris.filter(function (b) { return webBulanDari_(b.tglGrup) === bulan; }) : baris.slice(-150);
  return { sheet: nama, kategori: kat, baris: tampil.reverse(), saldo: baris.length ? baris[baris.length - 1].saldo : 0 };
}

function c2Debet_(kat, r) { let s = 0; for (let c = 2; c <= 6; c++) if (kat[c - 2][2] === 'DEBET') s += Number(r[c]) || 0; return s; }
function c2Kredit_(kat, r) { let s = 0; for (let c = 2; c <= 6; c++) if (kat[c - 2][2] === 'KREDIT') s += Number(r[c]) || 0; return s; }

function webSimpanPettyCash_(a, admin) {
  // fungsi popup Petty Cash inventory: kunci, urut tanggal, saldo, gabung sel tanggal -- persis sama
  const h = webInv_().processTransactionBatch({ tanggal: a.tanggal, items: a.items || [] });
  if (h && h.success) {
    webCatatLog_(admin, 'PETTY CASH', a.tanggal + ' | ' + (a.items || []).map(function (i) { return i.kategori + ' ' + i.nominal + ' ' + i.keterangan; }).join('; '));
  }
  return h;
}

/* ---------------- BUKU BANK ---------------- */

function webSimpanBukuBank_(a, admin) {
  const kat = WEB_KAS_KATEGORI_['BUKU BANK'];
  const tgl = webTanggalDariInput_(a.tanggal);
  if (!tgl) throw new Error('Tanggal wajib diisi.');
  const items = (a.items || []).map(function (it, i) {
    const k = kat.map(function (x) { return x[0]; }).indexOf(it.kategori);
    const nominal = Number(it.nominal);
    if (k === -1) throw new Error('Baris ' + (i + 1) + ': kategori tidak dikenali.');
    if (!(nominal > 0)) throw new Error('Baris ' + (i + 1) + ': nominal tidak valid.');
    if (!String(it.keterangan || '').trim()) throw new Error('Baris ' + (i + 1) + ': keterangan wajib diisi.');
    return { kolom: 3 + k, nominal: nominal, ket: String(it.keterangan).trim() };
  });
  if (!items.length) throw new Error('Belum ada transaksi.');

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new Error('Database sedang sibuk, coba lagi sebentar.');
  try {
    const sh = webSS_().getSheetByName('BUKU BANK');
    const mulai = WEB_KAS_MULAI_['BUKU BANK'];
    // baris terakhir yang berisi tanggal/keterangan (kolom H berisi rumus sampai bawah)
    const last = sh.getLastRow();
    const ab = last >= mulai ? sh.getRange(mulai, 1, last - mulai + 1, 2).getValues() : [];
    let akhir = mulai - 1;
    let tglAkhir = null;
    ab.forEach(function (r, i) {
      if (r[0] !== '' || String(r[1] || '').trim()) akhir = mulai + i;
      if (r[0] instanceof Date) tglAkhir = r[0];
    });
    const pertama = akhir + 1;
    const butuh = pertama + items.length - 1;
    if (butuh > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), butuh - sh.getMaxRows());
    const nilai = items.map(function (it) {
      const row = [tgl, it.ket, '', '', '', '', ''];
      row[it.kolom - 1] = it.nominal;
      return row;
    });
    sh.getRange(pertama, 1, nilai.length, 7).setValues(nilai);
    const rumus = items.map(function (it, i) {
      const r = pertama + i;
      return [r === mulai ? '=C' + r + '+D' + r + '-E' + r + '-F' + r + '-G' + r : '=H' + (r - 1) + '+C' + r + '+D' + r + '-E' + r + '-F' + r + '-G' + r];
    });
    sh.getRange(pertama, 8, rumus.length, 1).setFormulas(rumus);
    if (akhir >= mulai) sh.getRange(akhir, 1, 1, 8).copyTo(sh.getRange(pertama, 1, items.length, 8), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    SpreadsheetApp.flush();
    const saldo = sh.getRange(butuh, 8).getValue();
    webCatatLog_(admin, 'BUKU BANK', Utilities.formatDate(tgl, WEB_TZ, 'dd/MM/yyyy') + ' | ' + items.map(function (i) { return i.ket + ' ' + i.nominal; }).join('; '));
    return {
      success: true, baris: items.length, saldo: saldo,
      peringatan: tglAkhir && tgl.getTime() < tglAkhir.getTime() ? 'Tanggal ini lebih lama dari baris terakhir (' + webTgl_(tglAkhir) + '); baris tetap ditambahkan di paling bawah.' : ''
    };
  } finally {
    lock.releaseLock();
  }
}

/* ---------------- LAPORAN INVOICE ---------------- */

function webLaporan_(a) {
  const bulan = String(a.bulan || '') || Utilities.formatDate(new Date(), WEB_TZ, 'yyyy-MM');
  const semua = bulan === 'SEMUA';
  const cocok = function (b) { return semua || b === bulan; };

  // LOG_INVOICE_CETAK: A NO INV, B TGL, C CUSTOMER, D TOTAL, E BULAN, F/G TANDA, H JML CETAK, I STATUS, J ALASAN
  const log = webBaca_('LOG_INVOICE_CETAK', 10).filter(function (r) { return String(r[0] || '').trim(); });
  const bulanSet = {};
  const tercetak = [];
  let omzet = 0;
  let jmlCetak = 0;
  let jmlBatal = 0;
  const sudahDilog = {};
  log.forEach(function (r) {
    const b = webBulanDari_(r[4]) || String(r[4] || '');
    if (b) bulanSet[b] = true;
    sudahDilog[String(r[0]).trim()] = true;
    if (!cocok(b)) return;
    const st = String(r[8] || '').trim().toUpperCase();
    if (st === 'TERCETAK') { omzet += Number(r[3]) || 0; jmlCetak++; }
    if (st === 'BATAL') jmlBatal++;
    tercetak.push({ noInv: r[0], tgl: webTgl_(r[1]), tglMs: r[1] instanceof Date ? r[1].getTime() : 0, customer: r[2], total: Number(r[3]) || 0,
      jml: r[7], status: st, terakhir: r[6] instanceof Date ? Utilities.formatDate(r[6], WEB_TZ, 'dd/MM/yyyy HH:mm') : '', alasan: r[9] });
  });
  tercetak.sort(function (x, y) { return (y.tglMs - x.tglMs) || String(y.noInv).localeCompare(String(x.noInv)); });

  // PO pending: sisa qty = QUANTITY - qty DONE SJ, nilai = TOTAL x sisa / QUANTITY
  const perPO = {};
  webBaca_('HISTORY_PO_MASUK', 14).forEach(function (r) {
    const po = String(r[1] || '').trim();
    if (!po) return;
    const b = r[2] instanceof Date ? Utilities.formatDate(r[2], WEB_TZ, 'yyyy-MM') : '';
    if (b) bulanSet[b] = true;
    const qty = Number(r[10]) || 0;
    const st = String(r[11] || '').trim().toUpperCase();
    const done = st === 'DONE SJ' ? qty : (st.indexOf('DONE SJ ') === 0 ? (parseFloat(st.substring(8)) || 0) : 0);
    const sisa = Math.max(0, qty - done);
    if (!(sisa > 0) || !cocok(b)) return;
    const k = po + '||' + r[3];
    if (!perPO[k]) perPO[k] = { poNo: po, customer: r[3], tglMs: Infinity, sisa: 0, nilai: 0 };
    const p = perPO[k];
    if (r[2] instanceof Date && r[2].getTime() < p.tglMs) p.tglMs = r[2].getTime();
    p.sisa += sisa;
    p.nilai += qty > 0 ? (Number(r[13]) || 0) * sisa / qty : 0;
  });
  const hariIni = new Date();
  const pending = Object.keys(perPO).map(function (k) { return perPO[k]; }).sort(function (x, y) { return x.tglMs - y.tglMs; }).map(function (p) {
    return { poNo: p.poNo, customer: p.customer, tgl: isFinite(p.tglMs) ? webTgl_(new Date(p.tglMs)) : '', sisa: p.sisa, nilai: p.nilai,
      umur: isFinite(p.tglMs) ? Math.floor((hariIni.getTime() - p.tglMs) / 86400000) : '' };
  });

  // sudah diposting di PENJUALAN tapi belum ditandai tercetak
  const belum = [];
  webBaca_('PENJUALAN', 5).forEach(function (r) {
    const noInv = String(r[2] || '').trim();
    if (!noInv || sudahDilog[noInv] || !(r[0] instanceof Date)) return;
    if (!cocok(Utilities.formatDate(r[0], WEB_TZ, 'yyyy-MM'))) return;
    belum.push({ tgl: webTgl_(r[0]), customer: r[1], noInv: noInv, total: Number(r[4]) || 0 });
  });

  return {
    bulan: bulan, daftarBulan: Object.keys(bulanSet).sort().reverse(),
    omzet: omzet, jmlCetak: jmlCetak, jmlBatal: jmlBatal,
    totalPending: pending.reduce(function (s, p) { return s + p.nilai; }, 0), jmlPOPending: pending.length,
    tercetak: tercetak, pending: pending, belumDitandai: belum
  };
}
