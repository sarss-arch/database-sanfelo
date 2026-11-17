/* ======================================================
   GatewayEditPembayaran.js   (PROJECT: database_SANFELO)
   ------------------------------------------------------
   Aksi gateway "ubahRiwayatPembayaran": ubah TANGGAL dan/atau KETERANGAN
   ADJUSTMENT satu riwayat pembayaran invoice di PENJUALAN.
   Nominal, metode, No Post, SISA BAYAR dan STATUS TIDAK berubah.
   Semua perubahan dicatat di sheet LOG_EDIT_PEMBAYARAN.

   PASANG (sekali):
   1. Tambah file skrip baru bernama GatewayEditPembayaran, isi dengan file ini.
   2. Di Databasegateway.js, di dalam const ACTION_HANDLERS = { ... }
      tambahkan SATU baris ini (misalnya tepat di bawah ubahStatusTransaksi):

        ubahRiwayatPembayaran: function (ss, body) { return aksiUbahRiwayatPembayaran_(ss, body); },

   3. Terapkan > Kelola deployment > edit deployment gateway > Versi: Versi baru > Terapkan
      (URL tetap sama). Web SANFELO di project ini ikut memakai aksi ini.
   ====================================================== */

function aksiUbahRiwayatPembayaran_(ss, body) {
  // gateway sudah memegang LockService (prosesRequestGateway_) -> sudahTerkunci = true
  const h = ubahRiwayatPembayaranSS_(ss, body, body.oleh || 'INVENTORY', true);
  return { ok: true, noInvoice: h.noInvoice, noPost: h.noPost, lama: h.lama, baru: h.baru };
}

const EDIT_BAYAR_LOG_SHEET_ = 'LOG_EDIT_PEMBAYARAN';
const EDIT_BAYAR_KOL_NOINV_ = 3;       // C
const EDIT_BAYAR_KOL_CICILAN_ = 10;    // J
const EDIT_BAYAR_KET_MIN_ = 3;
const EDIT_BAYAR_KET_MAX_ = 60;

// "POST-115 - 31/08/2026 - Rp1.997.500 - METODE" -> bagian-bagiannya (format lama tanpa POST juga)
function pecahSelBayarEdit_(teks) {
  const s = String(teks || '').trim();
  if (!s) return null;
  const p = s.split(' - ');
  if (p.length < 2) return null;
  const i = p[0].trim().indexOf('POST-') === 0 ? 1 : 0;
  const metodeTeks = p.length > i + 2 ? p.slice(i + 2).join(' - ').trim() : '';
  const m = metodeTeks.match(/^ADJUSTMENT\s*\(([\s\S]*)\)\s*$/i);
  return {
    noPost: i ? p[0].trim() : '', tanggal: (p[i] || '').trim(), jumlahTeks: (p[i + 1] || '').trim(),
    jumlah: Number(String(p[i + 1] || '').replace(/[^0-9]/g, '')) || 0,
    metodeTeks: metodeTeks,
    metode: m || /^ADJUSTMENT$/i.test(metodeTeks) ? 'ADJUSTMENT' : metodeTeks.toUpperCase(),
    keterangan: m ? m[1].trim() : ''
  };
}

function bersihkanKetEdit_(ket) {
  return String(ket || '').replace(/[\r\n\t]+/g, ' ').replace(/\s+-\s+/g, ' ').replace(/[()]/g, ' ')
    .replace(/\s{2,}/g, ' ').trim().substring(0, EDIT_BAYAR_KET_MAX_).trim();
}

// "yyyy-MM-dd" (input tanggal) -> "dd/MM/yyyy"; null kalau tidak valid / di masa depan
function tglEditKeTeks_(iso) {
  const m = String(iso || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return null;
  const d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  if (d.getFullYear() !== Number(m[1]) || d.getMonth() !== Number(m[2]) - 1 || d.getDate() !== Number(m[3])) return null;
  const hariIni = new Date();
  if (d.getTime() > new Date(hariIni.getFullYear(), hariIni.getMonth(), hariIni.getDate()).getTime()) return 'MASA_DEPAN';
  return m[3] + '/' + m[2] + '/' + m[1];
}

/**
 * Dipakai bersama oleh: gateway database SANFELO, Inventory_SL, web SANFELO & web SL.
 * p: { noInvoice, urutan (0 = riwayat pertama), noPost, tanggalLama, jumlahLama,
 *      tglBaru ("yyyy-MM-dd", boleh kosong), keteranganBaru (boleh kosong), alasan }
 */
function ubahRiwayatPembayaranSS_(ss, p, oleh, sudahTerkunci) {
  const noInvoice = String(p.noInvoice || '').trim();
  const alasan = String(p.alasan || '').replace(/\s+/g, ' ').trim();
  if (!noInvoice) throw new Error('No invoice kosong.');
  if (alasan.length < 3) throw new Error('Alasan perubahan wajib diisi (minimal 3 huruf).');

  let tglBaru = '';
  if (p.tglBaru) {
    tglBaru = tglEditKeTeks_(p.tglBaru);
    if (!tglBaru) throw new Error('Tanggal baru tidak valid.');
    if (tglBaru === 'MASA_DEPAN') throw new Error('Tanggal bayar tidak boleh setelah hari ini.');
  }
  const ubahKet = p.keteranganBaru !== undefined && p.keteranganBaru !== null && String(p.keteranganBaru).trim() !== '';

  const sh = ss.getSheetByName('PENJUALAN');
  if (!sh) throw new Error('Sheet PENJUALAN tidak ditemukan.');
  // sudahTerkunci: dipanggil dari gateway database yang SUDAH memegang LockService
  const lock = sudahTerkunci ? null : LockService.getScriptLock();
  if (lock && !lock.tryLock(20000)) throw new Error('Spreadsheet sedang sibuk, coba lagi sebentar.');
  try {
    const last = sh.getLastRow();
    if (last < 2) throw new Error('PENJUALAN masih kosong.');
    const noInvs = sh.getRange(2, EDIT_BAYAR_KOL_NOINV_, last - 1, 1).getValues();
    let row = -1;
    for (let i = 0; i < noInvs.length; i++) if (String(noInvs[i][0] || '').trim() === noInvoice) { row = i + 2; break; }
    if (row === -1) throw new Error('Invoice ' + noInvoice + ' tidak ditemukan di PENJUALAN.');

    const lebar = Math.max(1, sh.getLastColumn() - EDIT_BAYAR_KOL_CICILAN_ + 1);
    const sel = sh.getRange(row, EDIT_BAYAR_KOL_CICILAN_, 1, lebar).getValues()[0];
    const daftar = [];
    sel.forEach(function (v, j) {
      const x = pecahSelBayarEdit_(v);
      if (x) { x.kolom = EDIT_BAYAR_KOL_CICILAN_ + j; x.teks = String(v).trim(); daftar.push(x); }
    });
    const noPost = String(p.noPost || '').trim();
    const t = noPost ? daftar.filter(function (x) { return x.noPost === noPost; })[0] : daftar[Number(p.urutan)];
    if (!t) throw new Error('Riwayat pembayaran tidak ditemukan. Cari ulang invoice-nya.');
    // pengaman: isi yang diubah harus sama dengan yang dilihat admin
    if (t.tanggal !== String(p.tanggalLama || '').trim() || Math.round(t.jumlah) !== Math.round(Number(p.jumlahLama) || 0)) {
      throw new Error('Riwayat ini baru saja berubah. Cari ulang invoice-nya lalu ubah lagi.');
    }

    let metodeTeks = t.metodeTeks;
    if (ubahKet) {
      if (t.metode !== 'ADJUSTMENT') throw new Error('Keterangan hanya bisa diubah untuk pembayaran ADJUSTMENT.');
      const ket = bersihkanKetEdit_(p.keteranganBaru);
      if (ket.length < EDIT_BAYAR_KET_MIN_) throw new Error('Keterangan adjustment minimal ' + EDIT_BAYAR_KET_MIN_ + ' huruf, contoh: pembulatan, biaya admin, sparepart.');
      metodeTeks = 'ADJUSTMENT (' + ket + ')';
    }
    const baru = (t.noPost ? t.noPost + ' - ' : '') + (tglBaru || t.tanggal) + ' - ' + t.jumlahTeks + ' - ' + metodeTeks;
    if (baru === t.teks) throw new Error('Tidak ada yang berubah.');

    sh.getRange(row, t.kolom).setValue(baru);
    let log = ss.getSheetByName(EDIT_BAYAR_LOG_SHEET_);
    if (!log) {
      log = ss.insertSheet(EDIT_BAYAR_LOG_SHEET_);
      log.getRange(1, 1, 1, 7).setValues([['WAKTU', 'NO INV', 'NO POST', 'ISI LAMA', 'ISI BARU', 'ALASAN', 'OLEH']]).setFontWeight('bold');
      log.setFrozenRows(1);
    }
    log.appendRow([new Date(), noInvoice, t.noPost, t.teks, baru, alasan, String(oleh || '')]);
    SpreadsheetApp.flush();
    return { ok: true, noInvoice: noInvoice, noPost: t.noPost, lama: t.teks, baru: baru };
  } finally {
    if (lock) lock.releaseLock();
  }
}
