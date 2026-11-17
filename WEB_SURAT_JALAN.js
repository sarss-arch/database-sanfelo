/* ======================================================
   WEB_SURAT_JALAN.js   (PROJECT: database_SANFELO)
   ------------------------------------------------------
   Web admin: buat Surat Jalan + Invoice dari PO Masuk.

   Alurnya sama dengan spreadsheet inventory_SANFELO:
     isi No PO (H3)  -> cek toko LOCKED -> popup pilih barang & qty
     (terapkanPilihanSJPopup) -> isi SAT / KET / No Kend / Sales /
     Keterangan -> 💾 Simpan Draft.
   Hitungan Simpan Draft memakai SALINAN PERSIS fungsi inventory
   (WEB_INV_SALINAN.js): tempo, potongan 3SET/5SET, HPP, status DONE SJ
   di HISTORY_PO_MASUK, PENJUALAN, dan TEMPO TOKO. Penulisan lewat aksi
   gateway yang sama (simpanTransaksiDraft, updateStatusDoneSjDelta,
   upsertPenjualan, syncTempoTokoFull).

   Satu perbedaan yang DISENGAJA: total invoice dihitung dari SEMUA baris
   barang. Di sheet INVOICE inventory rumus Discount (H20) =SUM(H11:H17)
   tidak ikut baris ke-8, jadi invoice 8 barang di sana totalnya lebih besar
   dari seharusnya.
   ====================================================== */

const WEB_SJ_BARIS_AWAL_ = 9; // SJ_ITEM_START_ROW di inventory (kolom P "BARIS" DETAIL_TRANSAKSI)

// sisa qty per baris PO (sama dengan parseDoneQty_ inventory)
function webSisaPO_(r, inv) {
  const qtyAsli = Number(r[10]) || 0;
  const statusRaw = String(r[11] || '').trim().toUpperCase();
  return Math.max(0, qtyAsli - inv.parseDoneQty_(statusRaw, qtyAsli));
}

function webDaftarPOTerbuka_(a) {
  const inv = webInv_();
  const cari = String(a.cari || '').trim().toUpperCase();
  const per = {};
  const urut = [];
  webBaca_('HISTORY_PO_MASUK', 18).forEach(function (r) {
    if (String(r[0] || '').trim().toUpperCase() !== 'MASUK') return;
    const poNo = String(r[1] || '').trim();
    if (!poNo) return;
    if (!per[poNo]) {
      per[poNo] = { poNo: poNo, customer: String(r[3] || '').trim(), tglPO: webTgl_(r[2]), tglKirim: webTgl_(r[7]),
        method: String(r[17] || '').trim(), cancel: false, sisaBarang: 0, sisaNilai: 0, barang: 0 };
      urut.push(poNo);
    }
    const p = per[poNo];
    if (String(r[11] || '').trim().toUpperCase() === 'CANCELLED') p.cancel = true;
    const qtyAsli = Number(r[10]) || 0;
    const sisa = webSisaPO_(r, inv);
    p.barang++;
    if (sisa > 0) {
      p.sisaBarang += sisa;
      p.sisaNilai += qtyAsli > 0 ? (Number(r[13]) || 0) * sisa / qtyAsli : 0;
    }
  });
  return urut.reverse().map(function (k) { return per[k]; }).filter(function (p) {
    if (p.cancel || p.sisaBarang <= 0) return false;
    return !cari || (p.poNo + ' ' + p.customer).toUpperCase().indexOf(cari) !== -1;
  });
}

// toko LOCKED tidak boleh dibuatkan SJ (sama dengan cekTempoTokoBlokir_). Status toko dihitung
// ulang persis seperti ringkasan G:H TEMPO TOKO di inventory (statusRingkasanToko_), karena
// kolom G:H di database bukan sumber yang diperbarui.
function webCekTokoLocked_(poNoList) {
  const hist = webBaca_('HISTORY_PO_MASUK', 4);
  const toko = {};
  poNoList.forEach(function (poNo) {
    for (let i = 0; i < hist.length; i++) {
      if (String(hist[i][0] || '').trim().toUpperCase() === 'MASUK' && String(hist[i][1] || '').trim() === poNo) {
        const n = String(hist[i][3] || '').trim();
        if (n) toko[n] = true;
        break;
      }
    }
  });
  const nama = Object.keys(toko);
  if (!nama.length) return '';
  const ringkasan = webInv_().statusRingkasanToko_();
  if (!ringkasan.length) return 'Sheet "TEMPO TOKO" masih kosong, tidak bisa cek status toko. Surat Jalan tidak bisa dibuat.';
  for (let t = 0; t < nama.length; t++) {
    const r = ringkasan.filter(function (x) { return x.customer === nama[t]; })[0];
    if (r && r.status === 'LOCKED') return 'Toko "' + nama[t] + '" punya pembayaran yang jatuh tempo (LOCKED). Surat Jalan tidak bisa dibuat. Catat pembayarannya, atau perpanjang tempo di menu Alat & Perbaikan (perlu password).';
  }
  return '';
}

// sama dengan loadSuratJalanFromHistory: barang dari PO (boleh beberapa PO)
function webMuatPOuntukSJ_(a) {
  const inv = webInv_();
  const poNoList = (a.poNos || []).map(function (s) { return String(s).trim(); }).filter(Boolean);
  if (!poNoList.length) throw new Error('Pilih minimal 1 No PO.');
  const blokir = webCekTokoLocked_(poNoList);
  if (blokir) throw new Error(blokir);

  const data = webBaca_('HISTORY_PO_MASUK', 18);
  let header = null;
  const items = [];
  const gagal = [];
  poNoList.forEach(function (poNo) {
    const cocok = data.filter(function (r) {
      return String(r[0] || '').trim().toUpperCase() === 'MASUK' && String(r[1] || '').trim() === poNo;
    });
    if (!cocok.length) { gagal.push(poNo + ' (tidak ditemukan)'); return; }
    if (String(cocok[0][11] || '').trim().toUpperCase() === 'CANCELLED') { gagal.push(poNo + ' (CANCELLED)'); return; }
    if (!header) {
      header = { toko: cocok[0][3], alamat: cocok[0][5], tglKirim: cocok[0][7] instanceof Date ? Utilities.formatDate(cocok[0][7], WEB_TZ, 'yyyy-MM-dd') : '', method: cocok[0][17] };
    }
    cocok.forEach(function (r) {
      const qtyAsli = Number(r[10]) || 0;
      const sisa = webSisaPO_(r, inv);
      items.push({ poNo: poNo, line: r[15], kode: r[8], nama: r[9], qtyAsli: qtyAsli, sisaQty: sisa, qty: sisa,
        status: String(r[11] || '').trim().toUpperCase(), harga: Number(r[12]) || 0, total: Number(r[13]) || 0, customer: String(r[3] || '').trim() });
    });
  });
  if (!items.length) throw new Error('Tidak ada barang yang bisa dimuat. ' + gagal.join(', '));
  const tokoBeda = items.filter(function (it) { return it.customer !== String(header.toko || '').trim(); });
  return { header: header, items: items, gagal: gagal, tokoBeda: tokoBeda.length > 0, nomor: webNomorSJBerikut_() };
}

// nomor SJ / INV berikutnya: sama dengan transaksiBaru (dari TRANSAKSI di database, per bulan berjalan)
function webNomorSJBerikut_() {
  const sekarang = new Date();
  const tahun = sekarang.getFullYear();
  const bulan = Utilities.formatString('%02d', sekarang.getMonth() + 1);
  const periode = tahun + '-' + bulan;
  const nomor = webBaca_('TRANSAKSI', 3);
  // nomor yang sedang tampil di form SURAT JALAN / INVOICE inventory (sudah "dipesan" admin lewat Transaksi Baru)
  const diForm = ['', webNilaiFormInventory_('SURAT JALAN', 'H6'), webNilaiFormInventory_('INVOICE', 'H6')];
  const maks = function (kol, prefix) {
    let m = 0;
    const cari = prefix + periode + '-';
    nomor.concat([diForm]).forEach(function (r) {
      const v = String(r[kol] || '').trim();
      if (v.indexOf(cari) === 0) {
        const n = parseInt(v.substring(cari.length), 10);
        if (!isNaN(n) && n > m) m = n;
      }
    });
    return m + 1;
  };
  return {
    noSJ: 'SJ-' + tahun + '-' + bulan + '-' + Utilities.formatString('%03d', maks(1, 'SJ-')),
    noInv: 'INV-' + tahun + '-' + bulan + '-' + Utilities.formatString('%03d', maks(2, 'INV-'))
  };
}

/**
 * Gabung pilihan per KODE + harga invoice (persis terapkanPilihanSJPopup),
 * lalu baris DETAIL (persis simpanDraft). Dipakai untuk pratinjau & simpan.
 */
function webSusunSJ_(a) {
  const inv = webInv_();
  const histMap = {};
  webBaca_('HISTORY_PO_MASUK', 16).forEach(function (row) {
    if (String(row[0] || '').trim().toUpperCase() !== 'MASUK') return;
    histMap[String(row[1] || '').trim() + '||' + String(row[15] || '').trim()] = {
      qtyAsli: Number(row[10]) || 0, hargaUnit: Number(row[12]) || 0, total: Number(row[13]) || 0
    };
  });

  const mergedMap = {};
  const merged = [];
  (a.pilihan || []).forEach(function (it) {
    const qty = Number(it.qty) || 0;
    if (qty <= 0) return;
    const kode = String(it.kode || '').trim();
    if (!mergedMap[kode]) { mergedMap[kode] = { kode: kode, qty: 0, sources: [] }; merged.push(mergedMap[kode]); }
    mergedMap[kode].qty += qty;
    mergedMap[kode].sources.push({ poNo: it.poNo, line: it.line, qty: qty });
  });

  const namaBarang = {};
  const shBarang = webMaster_('DATA_BERSIH');
  shBarang.getRange(2, 1, Math.max(1, shBarang.getLastRow() - 1), 2).getValues().forEach(function (r) {
    const k = String(r[0] || '').trim().toUpperCase();
    if (k && namaBarang[k] === undefined) namaBarang[k] = r[1];
  });
  const satKet = a.satKet || {};
  const historyPOMap = inv.buildHistoryPOStatusMap_();

  const baris = merged.map(function (item, i) {
    let sumHarga = 0;
    let sumTotal = 0;
    item.sources.forEach(function (src) {
      const fresh = histMap[String(src.poNo).trim() + '||' + String(src.line).trim()];
      if (fresh && fresh.qtyAsli > 0) {
        const ratio = src.qty / fresh.qtyAsli;
        sumHarga += Math.round(fresh.hargaUnit * ratio);
        sumTotal += Math.round(fresh.total * ratio);
      }
    });
    const sumber = JSON.stringify(item.sources);
    const persenSet = inv.hitungPersenPotonganSetPO_(sumber, historyPOMap);
    const net = persenSet > 0 ? sumTotal * (1 - persenSet) : sumTotal;
    const sk = satKet[item.kode] || {};
    return {
      no: i + 1, kode: item.kode, barang: namaBarang[item.kode.toUpperCase()] || '', qty: item.qty,
      sat: String(sk.sat || '').trim(), ket: String(sk.ket || '').trim(),
      harga: sumHarga, diskon: sumHarga - sumTotal, total: sumTotal, net: net, persenSet: persenSet,
      sumber: sumber, sources: item.sources
    };
  });
  const totalHarga = baris.reduce(function (s, b) { return s + b.harga; }, 0);
  const totalDiskon = baris.reduce(function (s, b) { return s + b.diskon; }, 0);
  return { baris: baris, totalHarga: totalHarga, totalDiskon: totalDiskon, totalBayar: totalHarga - totalDiskon };
}

function webPratinjauSJ_(a) {
  const inv = webInv_();
  const s = webSusunSJ_(a);
  const noPO = (a.poNos || []).join(',');
  const tgl = a.tgl ? webTanggalDariInput_(a.tgl) : null;
  const tempoHari = inv.hitungTempoHariUntukTransaksi_(noPO, String(a.toko || ''), String(a.keterangan || ''));
  s.tempoHari = tempoHari;
  s.tempo = tgl ? webTgl_(inv.hitungTanggalTempoTransaksi_(tgl, tempoHari)) : '';
  s.nomor = webNomorSJBerikut_();
  return s;
}

function webSimpanSJ_(a, admin) {
  const inv = webInv_();
  const poNos = (a.poNos || []).map(function (x) { return String(x).trim(); }).filter(Boolean);
  const toko = String(a.toko || '').trim();
  const daerah = String(a.alamat || '');
  const method = String(a.method || '').trim();
  const noPol = String(a.noPol || '').trim();
  const sales = String(a.sales || '').trim();
  const keterangan = String(a.keterangan || '');
  const tgl = inv.hanyaTanggal_(webTanggalDariInput_(a.tgl));

  if (!poNos.length) throw new Error('No PO kosong.');
  if (!(tgl instanceof Date) || isNaN(tgl.getTime())) throw new Error('Tanggal Surat Jalan belum diisi.');
  if (!toko) throw new Error('Nama toko kosong.');
  const edit = a.edit && a.edit.noSJ ? a.edit : null; // ubah SJ yang sudah tersimpan (Buka Draft)
  if (!edit) {
    const blokir = webCekTokoLocked_(poNos);
    if (blokir) throw new Error(blokir);
  }
  const oldMapping = edit ? inv.ambilMappingLamaDariDetailTransaksi_(inv.bacaDetailDB_(), edit.noSJ) : [];
  const qtyLama = {};
  oldMapping.forEach(function (m) { const k = String(m.poNo).trim() + '||' + String(m.line).trim(); qtyLama[k] = (qtyLama[k] || 0) + (Number(m.qty) || 0); });

  const s = webSusunSJ_(a);
  if (!s.baris.length) throw new Error('Tidak ada barang terisi.');
  // qty tidak boleh melebihi sisa PO (sama dengan batas di popup pilih barang)
  const sisa = {};
  webBaca_('HISTORY_PO_MASUK', 18).forEach(function (r) {
    if (String(r[0] || '').trim().toUpperCase() === 'MASUK') { const k = String(r[1]).trim() + '||' + String(r[15]).trim(); sisa[k] = webSisaPO_(r, inv) + (qtyLama[k] || 0); }
  });
  s.baris.forEach(function (b) {
    b.sources.forEach(function (src) {
      const k = String(src.poNo).trim() + '||' + String(src.line).trim();
      if (sisa[k] === undefined) throw new Error('Baris PO ' + src.poNo + ' line ' + src.line + ' tidak ditemukan.');
      if (src.qty > sisa[k]) throw new Error(b.kode + ' (' + src.poNo + '): qty ' + src.qty + ' melebihi sisa PO ' + sisa[k] + '. Muat ulang PO-nya.');
    });
  });

  const noPOGabungan = poNos.join(',');
  const tempoHari = inv.hitungTempoHariUntukTransaksi_(noPOGabungan, toko, keterangan);
  const tempo = inv.hitungTanggalTempoTransaksi_(tgl, tempoHari);

  // nomor diambil sesaat sebelum simpan dan dicek belum dipakai siapa pun (mode ubah: nomor lama)
  const nomor = edit ? { noSJ: edit.noSJ, noInv: edit.noInv } : webNomorSJBerikut_();
  if (!edit) {
    const sudahAda = webBaca_('TRANSAKSI', 3).some(function (r) {
      return String(r[1] || '').trim() === nomor.noSJ || String(r[2] || '').trim() === nomor.noInv;
    });
    if (sudahAda) throw new Error('Nomor ' + nomor.noSJ + ' baru saja dipakai. Klik simpan lagi.');
  }

  const now = inv.getTimestamp();
  const rows = s.baris.map(function (b, i) {
    const hpp = (Number(inv.getExtraData(b.kode)) || 0) * (Number(b.qty) || 0);
    return [nomor.noSJ, tgl, nomor.noInv, toko, tempo, b.kode, b.barang, b.qty, b.sat, b.ket, now, now,
      b.harga, b.diskon, b.net, WEB_SJ_BARIS_AWAL_ + i, method, hpp];
  });
  const mappingBaru = [];
  s.baris.forEach(function (b) { b.sources.forEach(function (src) { mappingBaru.push({ poNo: src.poNo, line: src.line, qty: src.qty }); }); });

  inv.updateStatusDoneSJHistoryPODelta_(oldMapping, mappingBaru);
  webGateway_('simpanTransaksiDraft', {
    header: {
      noSJ: nomor.noSJ, noInv: nomor.noInv, tgl: tgl, toko: toko, daerah: daerah, tempo: tempo, now: now,
      method: method, noPol: noPol, sales: sales, total: s.totalBayar, keterangan: keterangan, noPO: noPOGabungan
    },
    items: rows,
    sourcePO: s.baris.map(function (b) { return b.sumber; })
  });

  const catatan = [];
  try {
    inv.syncPenjualanFromTransaksi_(false);
  } catch (err) {
    catatan.push('Sinkron PENJUALAN gagal: ' + err.message + ' (jalankan Sync Penjualan dari Transaksi di spreadsheet).');
    try { inv.syncTempoToko_(false); } catch (err2) { catatan.push('Sinkron TEMPO TOKO gagal: ' + err2.message); }
  }
  webCatatLog_(admin, edit ? 'UBAH SURAT JALAN' : 'SURAT JALAN', nomor.noSJ + ' / ' + nomor.noInv + ' | ' + toko + ' | PO ' + noPOGabungan + ' | ' +
    rows.length + ' barang | ' + Math.round(s.totalBayar) + (catatan.length ? ' | ' + catatan.join('; ') : ''));
  return { noSJ: nomor.noSJ, noInv: nomor.noInv, tempo: webTgl_(tempo), total: s.totalBayar, barang: rows.length, catatan: catatan };
}
