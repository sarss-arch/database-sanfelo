/* ======================================================
   WEB_ALAT.js   (PROJECT: database_SANFELO)
   ------------------------------------------------------
   Menu inventory yang tersisa, versi web:
   - Ubah dokumen tersimpan (pengganti 📂 Buka Draft): Surat Jalan/Invoice
     (sekaligus 🔄 Ambil Ulang Harga dari PO), SJ Beli, Sparepart
   - PO Keluar (➕ Ke PO Keluar, 🆕 PO Keluar Baru, 📦 Export PO Keluar)
   - Alat: ↩️ Batalkan Tanda Tercetak, 📥 Catat Susulan, ✏️ Ubah Status
     Transaksi, 🗓️ Sinkron Tempo Toko, 🔧 Perbaiki Tempo Transaksi Lama,
     ✅ Perbaiki Status HISTORY PO, 🔁 Hitung Ulang QTY Terkirim, 👥 Grouping HUTANG
   Semuanya memakai fungsi inventory yang sama (WEB_INV_SALINAN.js) dan aksi
   gateway yang sama.
   ====================================================== */

var WEB_UI_KUMPUL_ = null; // pesan alert dari fungsi inventory dikumpulkan di sini saat alat dijalankan

function webJalankanAlat_(namaAlat, admin, fn) {
  WEB_UI_KUMPUL_ = [];
  try {
    fn();
    const pesan = WEB_UI_KUMPUL_.slice();
    webCatatLog_(admin, 'ALAT ' + namaAlat, pesan.join(' | ').slice(0, 400));
    return { pesan: pesan.length ? pesan : ['Selesai.'] };
  } finally {
    WEB_UI_KUMPUL_ = null;
  }
}

/* ======================= ALAT ======================= */

function webAlat_(a, admin) {
  const inv = webInv_();
  const alat = String(a.alat || '');
  if (alat === 'sinkronTempo') return webJalankanAlat_('Sinkron Tempo Toko', admin, function () { inv.syncTempoToko_(true); });
  if (alat === 'perbaikiTempoLama') return webJalankanAlat_('Perbaiki Tempo Transaksi Lama', admin, function () { inv.perbaikiTempoTransaksiLama(); });
  if (alat === 'perbaikiStatusPO') return webJalankanAlat_('Perbaiki Status HISTORY PO', admin, function () { inv.perbaikiStatusHistoryPOMassal(); });
  if (alat === 'hitungUlangQty') return webJalankanAlat_('Hitung Ulang QTY Terkirim', admin, function () { inv.hitungUlangQtyTerkirimPOKeluar_(true); });
  throw new Error('Alat tidak dikenal.');
}

function webBatalTercetak_(a, admin) {
  const noInv = String(a.noInv || '').trim().toUpperCase();
  const alasan = String(a.alasan || '').trim();
  if (!noInv) throw new Error('No invoice kosong.');
  if (!alasan) throw new Error('Alasan wajib diisi.');
  webGateway_('batalInvoiceTercetak', { noInv: noInv, alasan: alasan });
  webCatatLog_(admin, 'BATAL TERCETAK', noInv + ' | ' + alasan);
  return { pesan: '↩️ Tanda tercetak ' + noInv + ' dibatalkan.' };
}

// sama dengan catatSusulanInvoiceTercetak: invoice bulan itu yang belum ada di LOG_INVOICE_CETAK
const WEB_SUSULAN_ALASAN_ = 'Dicatat susulan: invoice sudah DONE SJ sebelum ada tombol SIAP CETAK';

function webSusulanKandidat_(bulan) {
  if (!/^\d{4}-\d{2}$/.test(bulan)) throw new Error('Format bulan harus yyyy-MM.');
  const tz = Session.getScriptTimeZone();
  const sudah = {};
  webBaca_('LOG_INVOICE_CETAK', 1).forEach(function (r) { const k = String(r[0] || '').trim(); if (k) sudah[k] = true; });
  const kandidat = [];
  const dilihat = {};
  let dilewati = 0;
  webBaca_('TRANSAKSI', 13).forEach(function (r) {
    const noInv = String(r[2] || '').trim();
    if (!noInv || dilihat[noInv]) return;
    const tgl = r[3] instanceof Date ? r[3] : null;
    if (!tgl || Utilities.formatDate(tgl, tz, 'yyyy-MM') !== bulan) return;
    dilihat[noInv] = true;
    if (sudah[noInv]) { dilewati++; return; }
    kandidat.push({ noInv: noInv, tglInvoice: Utilities.formatDate(tgl, tz, 'yyyy-MM-dd'), customer: String(r[4] || '').trim(), total: Number(r[12]) || 0 });
  });
  kandidat.sort(function (x, y) { return x.noInv < y.noInv ? -1 : 1; });
  return { bulan: bulan, kandidat: kandidat, dilewati: dilewati };
}

function webSusulanSimpan_(a, admin) {
  const k = webSusulanKandidat_(String(a.bulan || ''));
  const mulai = Date.now();
  const berhasil = [], gagal = [], sudahAda = [];
  let berhenti = false;
  for (let i = 0; i < k.kandidat.length; i++) {
    if (Date.now() - mulai > 4.5 * 60 * 1000) { berhenti = true; break; }
    const x = k.kandidat[i];
    try {
      if (webGateway_('cekStatusInvoiceCetak', { noInv: x.noInv }).ada) { sudahAda.push(x.noInv); continue; }
      webGateway_('catatInvoiceTercetak', { noInv: x.noInv, tglInvoice: x.tglInvoice, customer: x.customer, total: x.total, alasan: WEB_SUSULAN_ALASAN_ });
      berhasil.push(x.noInv);
    } catch (err) {
      gagal.push(x.noInv + ' (' + err.message + ')');
    }
  }
  webCatatLog_(admin, 'CATAT SUSULAN', k.bulan + ' | ' + berhasil.length + ' invoice');
  return { berhasil: berhasil, gagal: gagal, sudahAda: sudahAda, berhenti: berhenti };
}

function webDetailStatusSJ_(a) {
  const noSJ = String(a.noSJ || '').trim().toUpperCase();
  const t = webBaca_('TRANSAKSI', 17).filter(function (r) { return String(r[1] || '').trim().toUpperCase() === noSJ; }).pop();
  if (!t) throw new Error(noSJ + ' tidak ditemukan di TRANSAKSI.');
  return {
    noSJ: t[1], noInv: t[2], customer: t[4], status: t[13],
    items: webBaca_('DETAIL_TRANSAKSI', 20).filter(function (r) { return String(r[0] || '').trim().toUpperCase() === noSJ; })
      .map(function (r) { return { baris: r[15], kode: r[5], barang: r[6], qty: r[7], status: r[19] }; })
  };
}

function webUbahStatus_(a, admin) {
  const payload = a.sheet === 'DETAIL_TRANSAKSI'
    ? { sheet: 'DETAIL_TRANSAKSI', noSJ: a.noSJ, baris: a.baris, kode: a.kode }
    : { sheet: 'TRANSAKSI', noSJ: a.noSJ };
  payload.status = String(a.status || '').trim().toUpperCase();
  if (!payload.status) throw new Error('Status baru kosong.');
  const h = webGateway_('ubahStatusTransaksi', payload); // aksi yang sama dengan menu ✏️ Ubah Status Transaksi
  webCatatLog_(admin, 'UBAH STATUS', payload.sheet + ' ' + a.noSJ + (a.kode ? ' ' + a.kode : '') + ' | ' + (h.statusLama || '') + ' -> ' + h.statusBaru);
  return { pesan: 'Status diubah: ' + (h.statusLama || '(kosong)') + ' → ' + h.statusBaru };
}

function webGroupingDaftar_() {
  return webInv_().getSJBelumInvoice();
}

function webGroupingSimpan_(a, admin) {
  if (!a.jatuhTempo) throw new Error('JATUH TEMPO wajib diisi.');
  const h = webInv_().simpanGroupingInvoice({ noInv: a.noInv, jatuhTempo: a.jatuhTempo, transaksiBulan: a.transaksiBulan, rows: a.rows || [] });
  webCatatLog_(admin, 'GROUPING HUTANG', a.noInv + ' | ' + (a.rows || []).length + ' SJ | ' + a.transaksiBulan);
  return h;
}

/* ======================= PO KELUAR ======================= */

function webPOKeluarMuat_(a) {
  const inv = webInv_();
  const tgl = webTanggalDariInput_(a.tgl);
  if (!tgl) throw new Error('Tanggal PO keluar belum diisi.');
  const poNo = inv.getNoPOByTanggal_(tgl); // satu nomor per tanggal, sama dengan inventory
  const items = inv.ambilItemHistoryPOKeluar_(null, poNo);
  return { poNo: poNo, items: items, formInventory: webNilaiFormInventory_('PO_KELUAR', 'B4') };
}

function webPOKeluarDariPOMasuk_(a) {
  const inv = webInv_();
  const peta = inv.buatMapDataBersih_();
  const poNo = String(a.poNo || '').trim();
  const rows = webBaca_('HISTORY_PO_MASUK', 11).filter(function (r) {
    return String(r[0] || '').trim().toUpperCase() === 'MASUK' && String(r[1] || '').trim() === poNo;
  });
  if (!rows.length) throw new Error('PO masuk ' + poNo + ' tidak ditemukan.');
  return rows.filter(function (r) { return r[9] && r[10]; }).map(function (r) {
    return { toko: r[3], barang: r[9], jumlah: r[10], keterangan: '', nilaiF: inv.cariNilaiDataBersih_(peta, r[9]) };
  });
}

function webPOKeluarSimpan_(a, admin) {
  const inv = webInv_();
  const tgl = webTanggalDariInput_(a.tgl);
  if (!tgl) throw new Error('Tanggal PO keluar belum diisi.');
  const poNo = inv.getNoPOByTanggal_(tgl);
  if (String(a.poNo || '') !== poNo) throw new Error('Nomor PO berubah (' + a.poNo + ' -> ' + poNo + '). Muat ulang tanggalnya.');
  const diForm = webNilaiFormInventory_('PO_KELUAR', 'B4');
  if (diForm && diForm === poNo && !a.abaikanForm) {
    throw new Error('PO ' + poNo + ' sedang terbuka di form PO_KELUAR spreadsheet inventory. Selesaikan di sana dulu (🆕 PO Keluar Baru) supaya tidak saling timpa.');
  }
  const items = (a.items || []).filter(function (it) { return String(it.barang || '').trim(); }).map(function (it) {
    return [String(it.toko || '').trim(), String(it.barang).trim(), Number(it.jumlah) || 0, String(it.keterangan || ''), it.nilaiF === '' ? '' : Number(it.nilaiF) || 0];
  });
  // sama dengan simpanHistoryPOKeluar_ (dipanggil 🆕 PO Keluar Baru)
  webGateway_('syncPoKeluar', { poNo: poNo, tgl: Utilities.formatDate(tgl, 'Asia/Jakarta', 'yyyy-MM-dd'), items: items });
  const catatan = [];
  try { inv.hitungUlangQtyTerkirimPOKeluar_(false); } catch (err) { catatan.push('Hitung ulang QTY TERKIRIM gagal: ' + err.message); }
  webCatatLog_(admin, 'PO KELUAR', poNo + ' | ' + items.length + ' barang');
  return { poNo: poNo, barang: items.length, catatan: catatan };
}

/* ======================= UBAH DOKUMEN TERSIMPAN ======================= */

// Surat Jalan: barang PO + qty yang sudah ada di SJ ini (sisa PO ditambah qty SJ ini)
function webBukaSJ_(a) {
  const inv = webInv_();
  const noSJ = String(a.noSJ || '').trim().toUpperCase();
  const t = webBaca_('TRANSAKSI', 17).filter(function (r) { return String(r[1] || '').trim().toUpperCase() === noSJ; }).pop();
  if (!t) throw new Error(noSJ + ' tidak ditemukan di TRANSAKSI.');
  const det = webBaca_('DETAIL_TRANSAKSI', 21).filter(function (r) { return String(r[0] || '').trim().toUpperCase() === noSJ && String(r[5] || '').trim(); });
  const qtyIni = {};
  const poNos = [];
  const satKet = {};
  det.forEach(function (r) {
    let src = [];
    try { src = JSON.parse(r[20] || '[]'); } catch (err) { src = []; }
    if (!src.length) throw new Error(noSJ + ' berisi barang tanpa sumber PO (' + r[5] + '). SJ ini diubah lewat 📂 Buka Draft di spreadsheet.');
    src.forEach(function (s) {
      const k = String(s.poNo).trim() + '||' + String(s.line).trim();
      qtyIni[k] = (qtyIni[k] || 0) + (Number(s.qty) || 0);
      if (poNos.indexOf(String(s.poNo).trim()) === -1) poNos.push(String(s.poNo).trim());
    });
    satKet[String(r[5]).trim()] = { sat: r[8], ket: r[9] };
  });
  const data = webBaca_('HISTORY_PO_MASUK', 18);
  const items = [];
  poNos.forEach(function (poNo) {
    data.filter(function (r) { return String(r[0] || '').trim().toUpperCase() === 'MASUK' && String(r[1] || '').trim() === poNo; }).forEach(function (r) {
      const k = poNo + '||' + String(r[15]).trim();
      const ini = qtyIni[k] || 0;
      const qtyAsli = Number(r[10]) || 0;
      items.push({ poNo: poNo, line: r[15], kode: r[8], nama: r[9], qtyAsli: qtyAsli, sisaQty: webSisaPO_(r, inv) + ini, qty: ini, qtyIni: ini,
        status: String(r[11] || '').trim().toUpperCase(), harga: Number(r[12]) || 0, total: Number(r[13]) || 0 });
    });
  });
  return {
    edit: { noSJ: String(t[1]).trim(), noInv: String(t[2]).trim() }, poNos: poNos, items: items, satKet: satKet,
    header: { toko: t[4], alamat: t[5], tglKirim: t[3] instanceof Date ? Utilities.formatDate(t[3], WEB_TZ, 'yyyy-MM-dd') : '', method: t[9], noPol: t[10], sales: t[11], keterangan: t[15] },
    totalLama: webAngka_(t[12])
  };
}

function webBukaTT_(a) {
  const noPB = String(a.noPB || '').trim().toUpperCase();
  const rows = webBaca_('PEMBELIAN', 15).filter(function (r) { return String(r[0] || '').trim().toUpperCase() === noPB; });
  if (!rows.length) throw new Error(noPB + ' tidak ditemukan di PEMBELIAN.');
  const inv = webInv_();
  const r0 = rows[0];
  return {
    editPB: { noPB: String(r0[0]).trim(), noSJ: String(r0[2]).trim(), supplier: String(r0[3]).trim() },
    header: { supplier: r0[3], noPO: r0[1], noSJ: r0[2], tglBeli: r0[4] instanceof Date ? Utilities.formatDate(r0[4], WEB_TZ, 'yyyy-MM-dd') : '', noKend: r0[5], penerima: r0[11], pengirim: r0[12] },
    items: rows.filter(function (r) { return r[6]; }).map(function (r) {
      const poList = inv.parsePoListTT_(r[1]);
      return { kode: String(r[6]), qty: r[8], tokoAsal: r[10] || '', poAsal: poList.length === 1 ? poList[0] : '', setDiskon: String(r[14] || '').trim().toUpperCase() };
    })
  };
}

function webBukaSpr_(a) {
  const noSJ = String(a.noSJ || '').trim().toUpperCase();
  const rows = webBaca_('SPAREPART', 19).filter(function (r) { return String(r[0] || '').trim().toUpperCase() === noSJ; });
  if (!rows.length) throw new Error(noSJ + ' tidak ditemukan di SPAREPART.');
  const r0 = rows[0];
  const tgl = function (v) { return v instanceof Date ? Utilities.formatDate(v, WEB_TZ, 'yyyy-MM-dd') : ''; };
  return {
    edit: { noSJ: String(r0[0]).trim(), noInv: String(r0[1]).trim(), status: r0[18] || 'NOT DONE' },
    header: { refrensi: r0[2], tglSJ: tgl(r0[3]), tempo: tgl(r0[4]), noKend: r0[5], namaToko: String(r0[6] || '').trim(), alamat: r0[7], ketInv: r0[12], method: r0[15], sales: r0[16] },
    items: rows.filter(function (r) { return r[8]; }).map(function (r) { return { barang: r[8], qty: r[9], sat: r[10], ketBrg: r[11], harga: r[13] }; })
  };
}

/* ============ PERPANJANG TEMPO TOKO (= pilih UNLOCKED di TEMPO TOKO!H inventory) ============
   Sama dengan handlePerpanjangTempo_ di inventory: password dulu, lalu jumlah hari,
   lalu aksi gateway perpanjangTempoPerCustomer (TEMPO lama + N hari untuk SEMUA invoice
   LOCKED milik toko itu), lalu status dihitung ulang (updateStatusPembelianAll_). */

const WEB_TEMPO_SALAH_MAKS_ = 5; // salah password 5x -> dikunci 15 menit per admin

function webTempoTokoDaftar_() {
  const inv = webInv_();
  const m = inv.ringkasanTokoMentah_();
  const toko = m.customerOrder.map(function (c) {
    return {
      customer: c,
      status: inv.statusFinalCustomer_(m.customerInfo[c]),
      invoice: (m.lockedPerToko[c] || []).map(function (x) {
        const t = x.tempo instanceof Date ? x.tempo : new Date(x.tempo);
        return { noInv: x.noInv, tempo: webTgl_(x.tempo), tempoIso: isNaN(t.getTime()) ? '' : Utilities.formatDate(t, WEB_TZ, 'yyyy-MM-dd') };
      })
    };
  });
  return { toko: toko };
}

function webPerpanjangTempo_(a, admin) {
  const inv = webInv_();
  const customer = String(a.customer || '').trim();
  if (!customer) throw new Error('Pilih toko dulu.');

  const cache = CacheService.getScriptCache();
  const kunciSalah = 'WEBTEMPO_SALAH_' + String(admin || '').toUpperCase();
  const salah = Number(cache.get(kunciSalah)) || 0;
  if (salah >= WEB_TEMPO_SALAH_MAKS_) throw new Error('Password salah ' + salah + ' kali. Coba lagi 15 menit lagi.');
  const passBenar = PropertiesService.getScriptProperties().getProperty('WEB_PASSWORD_TEMPO') || inv.TEMPO_PASSWORD;
  if (String(a.password || '').trim() !== String(passBenar)) {
    cache.put(kunciSalah, String(salah + 1), 900);
    webCatatLog_(admin, 'PERPANJANG TEMPO DITOLAK', customer + ' | password salah');
    throw new Error('Password salah. Perpanjangan dibatalkan.');
  }
  cache.remove(kunciSalah);

  const jumlahHari = parseInt(String(a.jumlahHari || '').trim(), 10);
  if (isNaN(jumlahHari) || jumlahHari <= 0) throw new Error('Jumlah hari harus angka lebih dari 0. Perpanjangan dibatalkan.');
  if (jumlahHari > 365) throw new Error('Jumlah hari maksimal 365. Periksa lagi angkanya.');

  // 1) status TEMPO TOKO di database disegarkan dulu (sama dengan yang terlihat di ringkasan
  //    inventory), karena gateway hanya memperpanjang baris yang kolom E-nya LOCKED
  inv.updateStatusPembelianAll_();
  const sebelum = (inv.ringkasanTokoMentah_().lockedPerToko[customer] || []);
  if (!sebelum.length) throw new Error('Toko "' + customer + '" tidak punya invoice LOCKED. Tidak ada tempo yang perlu diperpanjang.');

  // 2) aksi gateway yang sama persis dengan inventory
  const h = webGateway_('perpanjangTempoPerCustomer', { customer: customer, jumlahHari: jumlahHari });

  // 3) hitung ulang status (database kolom E) seperti inventory
  inv.updateStatusPembelianAll_();

  const catatan = [];
  if (!WEB_LATIHAN_) {
    // ringkasan G:H di TEMPO TOKO inventory ikut diperbarui (di inventory ditulis oleh updateStatusPembelianAll_)
    try {
      const shInv = SpreadsheetApp.openById(INVENTORY_SS_ID_).getSheetByName('TEMPO TOKO');
      if (shInv) {
        const m2 = inv.ringkasanTokoMentah_();
        inv.writeRingkasanToko_(shInv, m2.customerOrder, m2.customerInfo);
        inv.paksaRefreshMirrorTempoToko_(shInv);
      }
    } catch (err) {
      catatan.push('Ringkasan TEMPO TOKO di inventory belum ter-update (' + err.message + '); akan ikut ter-update saat Sinkron Tempo berikutnya.');
    }
  }

  const tambah = function (iso) { const d = new Date(iso); d.setDate(d.getDate() + jumlahHari); return d; };
  const rows = h.simulasi
    ? sebelum.map(function (x) { return { noInv: x.noInv, tempoLama: webTgl_(x.tempo), tempoBaru: webTgl_(tambah(x.tempo)) }; })
    : (h.rows || []).map(function (r) {
        const lama = sebelum.filter(function (x) { return x.noInv === r.noInv; })[0];
        return { noInv: r.noInv, tempoLama: lama ? webTgl_(lama.tempo) : '', tempoBaru: webTgl_(r.tempoBaru) };
      });
  const statusBaru = h.simulasi ? '' : (inv.statusRingkasanToko_().filter(function (s) { return s.customer === customer; })[0] || {}).status || '';

  webCatatLog_(admin, 'PERPANJANG TEMPO', customer + ' | +' + jumlahHari + ' hari | ' + rows.map(function (r) { return r.noInv + ' ' + r.tempoLama + '->' + r.tempoBaru; }).join(', ').slice(0, 350));
  return {
    customer: customer, jumlahHari: jumlahHari, jumlah: h.simulasi ? rows.length : h.updatedCount,
    rows: rows, statusBaru: statusBaru, catatan: catatan
  };
}
