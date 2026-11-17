/* ======================================================
   WEB_POSTING_LAIN.js   (PROJECT: database_SANFELO)
   ------------------------------------------------------
   Web admin: Tanda Terima Barang (SJ Beli), Biaya Tambah (ekspedisi/palet)
   + Invoice Palet, dan SJ/Invoice Sparepart.
   Semua hitungan & penulisan memakai SALINAN PERSIS fungsi inventory
   (WEB_INV_SALINAN.js) dan aksi gateway yang sama dengan inventory.
   Nomor PB / SPR / SPRIN yang sedang tampil di form inventory dilewati,
   supaya simpanan dari inventory tidak menimpa simpanan dari web.
   ====================================================== */

function webMaksNomor_(daftar, prefix) {
  let m = 0;
  daftar.forEach(function (v) {
    const s = String(v || '').trim();
    if (s.indexOf(prefix) === 0) {
      const n = parseInt(s.substring(prefix.length), 10);
      if (!isNaN(n) && n > m) m = n;
    }
  });
  return m;
}

function webBarangMap_() {
  const sh = webMaster_('DATA_BERSIH');
  const data = sh.getRange(1, 1, Math.max(1, sh.getLastRow()), 11).getValues();
  const perKode = {};
  const perNama = {};
  data.forEach(function (r) {
    const k = String(r[0] || '').trim();
    if (!k) return;
    if (perKode[k.toUpperCase()] === undefined) perKode[k.toUpperCase()] = { kode: k, nama: r[1], hpp: Number(r[10]) || 0 };
    const n = String(r[1] || '').trim().toUpperCase();
    if (n) perNama[n] = r[0];
  });
  return { perKode: perKode, perNama: perNama };
}

/* ======================= TANDA TERIMA BARANG ======================= */

function webNomorPB_() {
  const periode = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM');
  const prefix = 'PB-' + periode + '-';
  const daftar = webBaca_('PEMBELIAN', 1).map(function (r) { return r[0]; });
  daftar.push(webNilaiFormInventory_('TANDA TERIMA BRG', 'B4')); // nomor yang sedang tampil di form inventory
  return prefix + String(webMaksNomor_(daftar, prefix) + 1).padStart(3, '0');
}

function webMasterTT_() {
  const supplier = {};
  webBaca_('PEMBELIAN', 4).forEach(function (r) { const s = String(r[3] || '').trim(); if (s) supplier[s] = true; });
  const per = {};
  const urut = [];
  webBaca_('HISTORY_PO_KELUAR', 9).forEach(function (r) {
    const po = String(r[0] || '').trim();
    if (!po || !r[3]) return;
    if (!per[po]) { per[po] = { poNo: po, tgl: webTgl_(r[1]), supplier: String(r[8] || '').trim(), sisa: 0, toko: {} }; urut.push(po); }
    per[po].sisa += Math.max(0, (Number(r[4]) || 0) - (Number(r[7]) || 0));
    if (r[2]) per[po].toko[String(r[2]).trim()] = true;
  });
  const bm = webBarangMap_();
  return {
    supplier: Object.keys(supplier).sort(),
    poKeluar: urut.reverse().map(function (k) { const p = per[k]; p.toko = Object.keys(p.toko).join(', '); return p; })
      .filter(function (p) { return p.sisa > 0; }),
    barang: Object.keys(bm.perKode).map(function (k) { return { kode: bm.perKode[k].kode, nama: bm.perKode[k].nama }; }),
    nomorPB: webNomorPB_()
  };
}

// sama dengan loadTandaTerimaFromHistoryInternal_ + terapkanPilihanTandaTerimaPopup
function webMuatPOKeluarTT_(a) {
  const inv = webInv_();
  const poList = inv.parsePoListTT_((a.poNos || []).join(','));
  const data = webBaca_('HISTORY_PO_KELUAR', 8);
  const bm = webBarangMap_();
  let tgl = null;
  const items = [];
  const gagal = [];
  poList.forEach(function (poNo) {
    const cocok = data.filter(function (r) { return String(r[0] || '').trim() === poNo; });
    if (!cocok.length) { gagal.push(poNo + ' (tidak ditemukan)'); return; }
    if (tgl === null) tgl = cocok[0][1];
    cocok.forEach(function (r) {
      if (!r[3]) return;
      const jumlah = Number(r[4]) || 0;
      items.push({
        poNo: poNo, namaToko: r[2], namaBarang: r[3], qtyAsli: jumlah,
        sisaQty: Math.max(0, jumlah - (Number(r[7]) || 0)),
        kode: bm.perNama[String(r[3]).trim().toUpperCase()] || r[3]
      });
    });
  });
  return { items: items, gagal: gagal, tgl: tgl instanceof Date ? Utilities.formatDate(tgl, WEB_TZ, 'yyyy-MM-dd') : '' };
}

// harga per baris persis hitungUlangHargaTandaTerima_ (HPP x qty x (1 - diskon SET))
function webHitungTT_(a) {
  const inv = webInv_();
  const hppMap = inv.buildHppMapDariDataBersih_();
  const bm = webBarangMap_();
  const baris = (a.items || []).filter(function (it) {
    return String(it.kode || '').trim() || String(it.qty || '').trim() || String(it.tokoAsal || '').trim();
  }).map(function (it) {
    const kode = String(it.kode || '').trim();
    const qty = Number(it.qty) || 0;
    const b = bm.perKode[kode.toUpperCase()];
    return {
      kode: kode, qty: it.qty, nama: b ? b.nama : '', tokoAsal: String(it.tokoAsal || '').trim(),
      poAsal: String(it.poAsal || '').trim(), setDiskon: String(it.setDiskon || '').trim().toUpperCase(),
      harga: kode ? (hppMap[kode] || 0) * (1 - inv.persenDiskonSetTandaTerima_(it.setDiskon)) * qty : ''
    };
  });
  return { baris: baris, total: baris.reduce(function (s, b) { return s + (Number(b.harga) || 0); }, 0) };
}

function webSimpanTT_(a, admin) {
  const inv = webInv_();
  const supplier = String(a.supplier || '').trim();
  const noSJ = String(a.noSJ || '').trim();
  const noPO = String(a.noPO || '');
  const tglBeli = a.tglBeli ? webTanggalDariInput_(a.tglBeli) : '';
  if (!noSJ) throw new Error('NO SJ masih kosong.');
  if (!supplier) throw new Error('Supplier masih kosong.');
  if (!tglBeli) throw new Error('Tanggal beli masih kosong.');

  const h = webHitungTT_(a);
  const items = h.baris.length ? h.baris : [{ kode: '', qty: '', nama: '', harga: '', tokoAsal: '', poAsal: '', setDiskon: '' }];
  const kodeTakDikenal = items.filter(function (it) { return it.kode && !it.nama; }).map(function (it) { return it.kode; });
  if (kodeTakDikenal.length) throw new Error('Kode tidak ada di DATA_BERSIH: ' + kodeTakDikenal.join(', '));

  // ubah SJ Beli yang sudah tersimpan: kalau No SJ / supplier diganti, baris lama dihapus dulu (sama dengan NO = Edit di inventory)
  const editPB = a.editPB && a.editPB.noPB ? a.editPB : null;
  if (editPB && (editPB.noSJ !== noSJ || String(editPB.supplier).trim().toUpperCase() !== supplier.toUpperCase())) {
    webGateway_('syncPembelian', { noSJ: editPB.noSJ, supplier: editPB.supplier, rows: [] });
  }
  const existing = inv.cariBarisPembelianByNoSJ_(noSJ, supplier);
  let noPB;
  if (editPB) a.timpa = true;
  if (existing.length) {
    if (!a.timpa) {
      return {
        perluTimpa: true, noPB: existing[0].noPB,
        pesan: 'NO SJ "' + noSJ + '" untuk supplier "' + existing[0].supplier + '" sudah pernah disimpan (' + existing[0].noPB + ', ' +
          existing.length + ' barang). Timpa dengan data baru?'
      };
    }
    noPB = existing[0].noPB; // timpa: nomor PB lama dipakai lagi supaya HUTANG & riwayat bayarnya tetap nyambung
    inv.hapusBarisPembelian_(null, existing, true);
  } else {
    noPB = editPB ? editPB.noPB : webNomorPB_();
  }

  const headerPOList = inv.parsePoListTT_(noPO);
  const ambigu = [];
  const perluHistory = headerPOList.length > 1 && items.some(function (it) { return !it.poAsal; });
  const hist = perluHistory ? inv.bacaHistoryPOKeluarData_(webSS_()) : [];
  items.forEach(function (it) {
    if (it.poAsal || headerPOList.length <= 1) return;
    const hasil = inv.cariPoItemDiHistory_(hist, headerPOList, it.nama || it.kode, it.tokoAsal, 'inc', it.qty);
    if (hasil.ambiguous && String(it.nama || it.kode || '').trim() !== '') ambigu.push(String(it.nama || it.kode));
  });
  const noPoGabungan = headerPOList.length ? headerPOList.join(', ') :
    inv.parsePoListTT_(items.filter(function (it) { return it.poAsal; }).map(function (it) { return it.poAsal; }).join(',')).join(', ');

  const rows = items.map(function (it) {
    return [noPB, noPoGabungan, noSJ, supplier, tglBeli, String(a.noKend || ''), it.kode, it.nama, it.qty, it.harga,
      it.tokoAsal, String(a.penerima || ''), String(a.pengirim || ''), '', it.setDiskon];
  });
  webGateway_('syncPembelian', {
    noSJ: noSJ, supplier: supplier,
    rows: rows.map(function (r) { return r.map(function (v) { return v instanceof Date ? v.toISOString() : v; }); })
  });
  const catatan = [];
  try { inv.hitungUlangQtyTerkirimPOKeluar_(false); } catch (err) { catatan.push('QTY TERKIRIM gagal diupdate: ' + err.message); }
  try { inv.syncHutangFromPembelian_(); } catch (err) { catatan.push('Sync HUTANG gagal: ' + err.message); }
  if (ambigu.length) catatan.push(ambigu.length + ' barang (' + ambigu.join(', ') + ') ada di lebih dari satu PO -- cek QTY TERKIRIM di HISTORY_PO_KELUAR.');
  webCatatLog_(admin, 'TANDA TERIMA', noPB + ' | ' + supplier + ' | SJ ' + noSJ + ' | ' + rows.length + ' barang | ' + Math.round(h.total) +
    (existing.length ? ' | TIMPA' : ''));
  return { noPB: noPB, barang: rows.length, total: h.total, timpa: existing.length > 0, catatan: catatan };
}

/* ======================= BIAYA TAMBAH & INVOICE PALET ======================= */

function webInvoiceUntukBiaya_(noInv) {
  const p = webCariDataInvoice_(noInv);
  if (!p.found) throw new Error(p.message);
  const trx = webBaca_('TRANSAKSI', 17).filter(function (r) { return String(r[2] || '').trim() === p.noInvoice; })[0];
  return {
    noInv: p.noInvoice, toko: p.namaToko, alamat: trx ? trx[5] : '', method: p.method, tanggal: trx && trx[3] instanceof Date ? Utilities.formatDate(trx[3], WEB_TZ, 'yyyy-MM-dd') : '',
    ekspedisi: p.ekspedisi, palet: p.palet,
    barang: webBaca_('DETAIL_TRANSAKSI', 9).filter(function (r) { return String(r[2] || '').trim() === p.noInvoice; })
      .map(function (r) { return { nama: r[6], qty: r[7], sat: r[8] }; })
  };
}

function webSimpanBiaya_(a, admin) {
  // fungsi popup Biaya Tambah inventory (validasi, kunci, EKSPEDISI + PENJUALAN H/I)
  const h = webInv_().simpanBiayaTambahan(a);
  if (h && h.success) {
    webCatatLog_(admin, 'BIAYA TAMBAH', String(a.referensi) + ' | ' + a.jenisBiaya + ' | ' + a.namaEkspedisi + ' | ' + a.qty + ' x ' + a.harga);
  }
  return h;
}

function webDaftarPalet_() {
  return webInv_().getDaftarInvoicePalet();
}

function webDetailPalet_(a) {
  const inv = webInv_();
  const ref = String(a.referensi || '').trim();
  const rows = inv.bacaEkspedisiDb_().filter(function (r) {
    return String(r[1] || '').trim() === ref && inv.jenisBarisEkspedisi_(r);
  });
  if (!rows.length) throw new Error('Tidak ada biaya EKSPEDISI/PALET untuk ' + ref + '.');
  const r0 = rows[0];
  const items = rows.map(function (r) {
    return { barang: r[7], jenis: inv.jenisBarisEkspedisi_(r), qty: r[8], sat: r[9], ket: r[10], jumlah: Number(r[11]) || 0, ekspedisi: r[2] };
  });
  const jumlah = function (j) { return items.filter(function (i) { return i.jenis === j; }).reduce(function (s, i) { return s + i.jumlah; }, 0); };
  return {
    referensi: ref, toko: r0[5], alamat: r0[6], tanggal: webTgl_(r0[3]), tempo: webTgl_(r0[4]), method: r0[12],
    ekspedisi: items.map(function (i) { return i.ekspedisi; }).filter(function (v, i, a) { return v && a.indexOf(v) === i; }).join(', '),
    items: items, biayaEkspedisi: jumlah('EKSPEDISI'), biayaPalet: jumlah('PALET')
  };
}

/* ======================= SPAREPART ======================= */

function webNomorSparepart_() {
  const bulan = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM');
  const data = webBaca_('SPAREPART', 2);
  const ambil = function (kol, prefix, diForm) {
    const p = prefix + '-' + bulan + '-';
    const daftar = data.map(function (r) { return r[kol]; });
    daftar.push(diForm);
    return p + String(webMaksNomor_(daftar, p) + 1).padStart(3, '0');
  };
  return {
    noSJ: ambil(0, 'SPR', webNilaiFormInventory_('SJ_SPR', 'H6')),
    noInv: ambil(1, 'SPRIN', webNilaiFormInventory_('INV_SPR', 'H5'))
  };
}

// HPP per barang dari PEMBELIAN (sama dengan simpanSparepart): 1 data -> langsung, >1 data -> admin memilih
function webPratinjauSparepart_(a) {
  const peta = webInv_().buildPembelianMap_();
  const ambigu = [];
  const seen = {};
  (a.items || []).forEach(function (it) {
    const key = String(it.barang || '').trim().toUpperCase();
    if (!key || seen[key]) return;
    seen[key] = true;
    const e = peta[key] || [];
    if (e.length > 1) ambigu.push({ key: key, namaBarang: e[0].namaBarang, entries: e.map(function (x) { return { supplier: x.supplier, total: x.total, qty: x.qty, unitPrice: x.unitPrice }; }) });
  });
  return { ambigu: ambigu, nomor: webNomorSparepart_() };
}

function webSimpanSparepart_(a, admin) {
  const inv = webInv_();
  const items = (a.items || []).filter(function (it) { return String(it.barang || '').trim(); }).map(function (it) {
    return { barang: String(it.barang).trim(), qty: Number(it.qty) || 0, sat: String(it.sat || ''), ketBrg: String(it.ketBrg || ''), harga: Number(it.harga) || 0 };
  });
  if (!items.length) throw new Error('Belum ada barang.');
  if (items.length > 8) throw new Error('Maksimal 8 barang per SJ sparepart (sama dengan form SJ_SPR).');
  if (!String(a.namaToko || '').trim()) throw new Error('Nama toko masih kosong.');
  const tglSJ = webTanggalDariInput_(a.tglSJ);
  if (!tglSJ) throw new Error('Tanggal SJ masih kosong.');
  const tempo = a.tempo ? webTanggalDariInput_(a.tempo) : '';

  const peta = inv.buildPembelianMap_();
  const pilihan = a.pilihanHpp || {};
  const hpp = {};
  items.forEach(function (it) {
    const key = it.barang.toUpperCase();
    const e = peta[key] || [];
    if (!e.length) hpp[key] = '';
    else if (e.length === 1) hpp[key] = e[0].unitPrice;
    else hpp[key] = (e[Number(pilihan[key]) || 0] || e[0]).unitPrice;
  });

  const edit = a.edit && a.edit.noSJ ? a.edit : null; // ubah sparepart tersimpan: nomor & STATUS ORDER lama dipertahankan
  const nomor = edit ? { noSJ: edit.noSJ, noInv: edit.noInv } : webNomorSparepart_();
  const rows = items.map(function (it) {
    const unit = hpp[it.barang.toUpperCase()];
    const h = typeof unit === 'number' ? unit * it.qty : '';
    const laba = typeof h === 'number' ? it.harga - h : '';
    return [nomor.noSJ, nomor.noInv, String(a.refrensi || ''), tglSJ, tempo, String(a.noKend || ''), String(a.namaToko).trim(),
      String(a.alamat || ''), it.barang, it.qty, it.sat, it.ketBrg, String(a.ketInv || ''), it.harga, h,
      String(a.method || ''), String(a.sales || ''), laba, edit ? (edit.status || 'NOT DONE') : 'NOT DONE'];
  });
  webGateway_('syncSparepart', {
    noSJ: nomor.noSJ,
    rows: rows.map(function (r) { return r.map(function (v) { return v instanceof Date ? v.toISOString() : v; }); })
  });
  const catatan = [];
  try { inv.syncPenjualanFromSparepart_(false); } catch (err) { catatan.push('Invoice ini BELUM masuk PENJUALAN (' + err.message + ').'); }
  webCatatLog_(admin, 'SPAREPART', nomor.noSJ + ' / ' + nomor.noInv + ' | ' + a.namaToko + ' | ' + rows.length + ' barang');
  return { noSJ: nomor.noSJ, noInv: nomor.noInv, barang: rows.length, total: items.reduce(function (s, i) { return s + i.harga; }, 0), catatan: catatan };
}
