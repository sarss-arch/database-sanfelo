/* ======================================================
   SYNC_BIAYA_TAMBAHAN.js   (PROJECT: database_SANFELO)
   ------------------------------------------------------
   Menyamakan biaya EKSPEDISI & PALET di semua tempat:
     EKSPEDISI (log per barang)  <->  PENJUALAN H (EKSPEDISI) & I (PALET)
   Aturan: PENJUALAN H = jumlah HARGA baris EKSPEDISI ber-JENIS "EKSPEDISI"
           PENJUALAN I = jumlah HARGA baris EKSPEDISI ber-JENIS "PALET"
           untuk NO INV (= REFERENSI) yang sama.

   Kolom baru EKSPEDISI N = JENIS. Baris baru dari "🚚 Biaya Tambah" otomatis
   terisi; baris lama diisi lewat langkah 1 di bawah.

   LANGKAH (jalankan dari editor, pilih fungsi -> Run, lihat "Log eksekusi"):
     1. isiJenisBiayaEkspedisi
        Isi JENIS baris lama yang bisa dipastikan dari PENJUALAN H/I (mis.
        H 145.000 & I 65.000, barisnya 145.000 & 65.000 -> jelas). Yang tidak
        bisa dipastikan dibiarkan KOSONG = tidak pakai palet/ekspedisi (sesuai
        admin). Mau diubah? pilih manual di kolom N (ada dropdown).
        (cekJenisBiayaEkspedisi = sama, tapi cuma laporan)
     2. syncBiayaTambahanKePenjualan
        Tulis PENJUALAN H/I = jumlah dari EKSPEDISI untuk NO INV yang punya
        minimal 1 baris ber-JENIS. Baris JENIS kosong tidak dihitung. NO INV
        yang punya H/I di PENJUALAN tapi tidak ada satu pun baris ber-JENIS di
        EKSPEDISI cuma dilaporkan, tidak diubah.
        (cekSyncBiayaTambahan = sama, tapi cuma laporan)
   Aman diulang. SISA BAYAR tidak berubah (H/I memang tidak ikut SISA BAYAR).
   ====================================================== */

const EKSPEDISI_COL_REF_SYNC_ = 2;     // B REFERENSI
const EKSPEDISI_COL_HARGA_SYNC_ = 12;  // L HARGA
const EKSPEDISI_COL_JENIS_SYNC_ = 14;  // N JENIS (baru)
const JENIS_BIAYA_SYNC_ = ['EKSPEDISI', 'PALET'];

function cekJenisBiayaEkspedisi() { isiJenisBiayaInti_(true); }
function isiJenisBiayaEkspedisi() { isiJenisBiayaInti_(false); }
function cekSyncBiayaTambahan() { syncBiayaTambahanInti_(true); }
function syncBiayaTambahanKePenjualan() { syncBiayaTambahanInti_(false); }

function rpSync_(n) { return 'Rp' + Number(n || 0).toLocaleString('id-ID'); }

function bacaEkspedisiSync_(ss) {
  const sh = ss.getSheetByName(EKSPEDISI_SHEET);
  const lastRow = sh.getLastRow();
  const rows = lastRow >= 2 ? sh.getRange(2, 1, lastRow - 1, EKSPEDISI_COL_JENIS_SYNC_).getValues() : [];
  return { sh: sh, rows: rows };
}

// NO INV -> { h, i, row } dari PENJUALAN
function bacaPenjualanSync_(ss) {
  const sh = ss.getSheetByName(PENJUALAN_SHEET);
  const lastRow = sh.getLastRow();
  const map = {};
  if (lastRow < 2) return { sh: sh, map: map };
  sh.getRange(2, 1, lastRow - 1, PENJUALAN_COL_PALET).getValues().forEach(function (r, i) {
    const no = String(r[PENJUALAN_COL_NOINV - 1] || '').trim().toUpperCase();
    if (no && !map[no]) map[no] = { h: Number(r[PENJUALAN_COL_EKSPEDISI - 1]) || 0, i: Number(r[PENJUALAN_COL_PALET - 1]) || 0, row: i + 2 };
  });
  return { sh: sh, map: map };
}

function jenisSync_(v) {
  const j = String(v || '').trim().toUpperCase();
  return JENIS_BIAYA_SYNC_.indexOf(j) !== -1 ? j : '';
}

/**
 * Inti tebakan JENIS (tanpa sheet).
 * baris  : [{ idx, harga, jenis }] satu NO INV
 * target : { h, i } dari PENJUALAN
 * return : { idx: "EKSPEDISI"/"PALET" } hanya untuk baris yang PASTI
 *          (semua kombinasi yang cocok sepakat).
 */
function tebakJenisBiaya_(baris, target) {
  const kosong = baris.filter(function (b) { return !b.jenis; });
  if (!kosong.length || kosong.length > 12) return {};
  let sisaH = target.h, sisaI = target.i;
  baris.forEach(function (b) {
    if (b.jenis === 'EKSPEDISI') sisaH -= b.harga;
    if (b.jenis === 'PALET') sisaI -= b.harga;
  });
  const sepakat = {};
  let jumlahSolusi = 0;
  for (let mask = 0; mask < (1 << kosong.length); mask++) {
    let palet = 0, eks = 0;
    kosong.forEach(function (b, k) { if (mask & (1 << k)) palet += b.harga; else eks += b.harga; });
    if (palet !== sisaI || eks !== sisaH) continue;
    jumlahSolusi++;
    kosong.forEach(function (b, k) {
      const j = (mask & (1 << k)) ? 'PALET' : 'EKSPEDISI';
      if (sepakat[b.idx] === undefined) sepakat[b.idx] = j;
      else if (sepakat[b.idx] !== j) sepakat[b.idx] = null;
    });
  }
  const hasil = {};
  if (!jumlahSolusi) return hasil;
  Object.keys(sepakat).forEach(function (k) { if (sepakat[k]) hasil[k] = sepakat[k]; });
  return hasil;
}

function isiJenisBiayaInti_(cumaCek) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const eks = bacaEkspedisiSync_(ss);
  const pj = bacaPenjualanSync_(ss).map;
  const log = [];

  if (!cumaCek) {
    eks.sh.getRange(1, EKSPEDISI_COL_JENIS_SYNC_).setValue('JENIS').setFontWeight('bold');
    const rule = SpreadsheetApp.newDataValidation().requireValueInList(JENIS_BIAYA_SYNC_, true).setAllowInvalid(false).build();
    eks.sh.getRange(2, EKSPEDISI_COL_JENIS_SYNC_, Math.max(eks.sh.getMaxRows() - 1, 1), 1).setDataValidation(rule);
  }

  // kelompokkan per NO INV
  const grup = {}, urutan = [];
  eks.rows.forEach(function (r, i) {
    const ref = String(r[EKSPEDISI_COL_REF_SYNC_ - 1] || '').trim();
    if (!ref) return;
    const k = ref.toUpperCase();
    if (!grup[k]) { grup[k] = { ref: ref, baris: [] }; urutan.push(k); }
    grup[k].baris.push({ idx: i, harga: Number(r[EKSPEDISI_COL_HARGA_SYNC_ - 1]) || 0, jenis: jenisSync_(r[EKSPEDISI_COL_JENIS_SYNC_ - 1]) });
  });

  const jenisBaru = eks.rows.map(function (r) { return [r[EKSPEDISI_COL_JENIS_SYNC_ - 1]]; });
  let terisi = 0, manual = 0;
  urutan.forEach(function (k) {
    const g = grup[k];
    const kosong = g.baris.filter(function (b) { return !b.jenis; });
    if (!kosong.length) return;
    const p = pj[k];
    const tebak = p ? tebakJenisBiaya_(g.baris, p) : {};
    kosong.forEach(function (b) {
      if (tebak[b.idx]) { jenisBaru[b.idx] = [tebak[b.idx]]; terisi++; }
    });
    const sisa = kosong.filter(function (b) { return !tebak[b.idx]; });
    if (Object.keys(tebak).length) {
      log.push('✅ ' + g.ref + ': ' + Object.keys(tebak).map(function (i) { return 'baris ' + (Number(i) + 2) + ' = ' + tebak[i]; }).join(', '));
    }
    if (sisa.length) {
      manual += sisa.length;
      const alasan = !p ? 'NO INV tidak ada di PENJUALAN (salah ketik REFERENSI?)'
        : 'PENJUALAN H ' + rpSync_(p.h) + ' / I ' + rpSync_(p.i) + ' tidak cocok dengan harga barisnya';
      log.push('➖ ' + g.ref + ' -> dibiarkan KOSONG (tidak pakai palet/ekspedisi): ' +
        sisa.map(function (b) { return 'baris ' + (b.idx + 2) + ' (' + rpSync_(b.harga) + ')'; }).join(', ') + ' — ' + alasan);
    }
  });

  if (!cumaCek && terisi) eks.sh.getRange(2, EKSPEDISI_COL_JENIS_SYNC_, jenisBaru.length, 1).setValues(jenisBaru);
  Logger.log((cumaCek ? 'CEK SAJA (belum ada yang diubah)\n' : 'SELESAI\n') +
    'JENIS ' + (cumaCek ? 'bisa diisi otomatis' : 'terisi otomatis') + ': ' + terisi + ' baris. Dibiarkan kosong: ' + manual + ' baris.\n' +
    log.join('\n'));
}

function syncBiayaTambahanInti_(cumaCek) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const eks = bacaEkspedisiSync_(ss);
  const pj = bacaPenjualanSync_(ss);
  const log = [];

  const jumlah = {};
  eks.rows.forEach(function (r) {
    const k = String(r[EKSPEDISI_COL_REF_SYNC_ - 1] || '').trim().toUpperCase();
    const j = jenisSync_(r[EKSPEDISI_COL_JENIS_SYNC_ - 1]);
    if (!k || !j) return; // JENIS kosong = tidak pakai palet/ekspedisi
    if (!jumlah[k]) jumlah[k] = { h: 0, i: 0 };
    const harga = Number(r[EKSPEDISI_COL_HARGA_SYNC_ - 1]) || 0;
    if (j === 'EKSPEDISI') jumlah[k].h += harga;
    else jumlah[k].i += harga;
  });

  let diubah = 0;
  Object.keys(jumlah).forEach(function (k) {
    const p = pj.map[k];
    if (!p) { log.push('⚠️ ' + k + ': ada di EKSPEDISI tapi NO INV-nya tidak ada di PENJUALAN (salah ketik REFERENSI?) — dilewati.'); return; }
    const s = jumlah[k];
    if (s.h === p.h && s.i === p.i) return;
    diubah++;
    log.push((cumaCek ? '🔎 ' : '✅ ') + k + ' (PENJUALAN baris ' + p.row + '): H ' + rpSync_(p.h) + ' -> ' + rpSync_(s.h) +
      ', I ' + rpSync_(p.i) + ' -> ' + rpSync_(s.i));
    if (cumaCek) return;
    pj.sh.getRange(p.row, PENJUALAN_COL_EKSPEDISI, 1, 2).setValues([[s.h || '', s.i || '']]).setNumberFormat('#,##0');
  });

  Object.keys(pj.map).forEach(function (k) {
    const p = pj.map[k];
    if ((p.h || p.i) && !jumlah[k]) {
      log.push('❓ ' + k + ' (PENJUALAN baris ' + p.row + '): H ' + rpSync_(p.h) + ', I ' + rpSync_(p.i) +
        ' tapi tidak ada baris ber-JENIS di EKSPEDISI — tidak diubah, cek manual.');
    }
  });

  Logger.log((cumaCek ? 'CEK SAJA (belum ada yang diubah)\n' : 'SELESAI\n') +
    'PENJUALAN H/I ' + (cumaCek ? 'akan diubah' : 'diubah') + ': ' + diubah + ' NO INV.\n' + log.join('\n'));
}
