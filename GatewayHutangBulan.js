/* ======================================================
   GatewayHutangBulan.js   (PROJECT: database_SANFELO)
   ------------------------------------------------------
   Isi otomatis kolom A (TRANSAKSI BULAN) di HUTANG yang masih KOSONG,
   pakai bulan dari TGL BELI (kolom F), format teks "yyyy-MM".

   Kenapa: sebelum migrasi, BULAN diketik langsung di sheet HUTANG
   inventory. Sekarang sheet itu cuma salinan IMPORTRANGE (tidak bisa
   diketik), dan baris baru dari syncHutangDariPembelian ditulis dengan
   BULAN kosong -- jadi sejak September BULAN tidak pernah terisi, dan
   hutang-nya TIDAK MUNCUL di popup Bayar Hutang (yang mencari per BULAN).
   BULAN yang sudah terisi (termasuk hasil Grouping HUTANG) TIDAK diubah.

   PASANG (sekali):
   1. Tambahkan file ini ke project database_SANFELO.
   2. Di Databasegateway.js, di dalam const ACTION_HANDLERS = { ... },
      tambahkan satu baris:
        isiBulanHutangKosong: function (ss, body) { return aksiIsiBulanHutangKosong_(ss, body); },
   3. Terapkan > Kelola deployment > edit > Versi baru > Terapkan.
   4. Jalankan isiBulanHutangKosong dari editor SEKALI untuk mengisi
      baris lama (Sep/Okt). Setelah itu inventory memanggilnya otomatis
      tiap Simpan SJ Beli.
   ====================================================== */

function isiBulanHutangKosong() {
  const hasil = aksiIsiBulanHutangKosong_(SpreadsheetApp.getActiveSpreadsheet(), {});
  Logger.log('BULAN HUTANG diisi: ' + hasil.diisi + ' baris. ' + (hasil.contoh || []).join(', '));
}

function aksiIsiBulanHutangKosong_(ss, body) {
  const sh = ss.getSheetByName(HUTANG_SHEET);
  if (!sh) throw new Error('Sheet "' + HUTANG_SHEET + '" tidak ditemukan di database.');
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { ok: true, diisi: 0 };

  const tz = Session.getScriptTimeZone();
  const data = sh.getRange(2, 1, lastRow - 1, HUTANG_COL_TGLBELI).getValues(); // A..F
  const contoh = [];
  let diisi = 0;
  data.forEach(function (r, i) {
    if (String(r[HUTANG_COL_BULAN - 1] || '').trim() !== '' || r[HUTANG_COL_BULAN - 1] instanceof Date) return;
    if (!String(r[HUTANG_COL_PB - 1] || '').trim()) return;
    const tgl = parseTanggalDariGateway_(r[HUTANG_COL_TGLBELI - 1]);
    if (!(tgl instanceof Date) || isNaN(tgl.getTime())) return;
    const bulan = Utilities.formatDate(tgl, tz, 'yyyy-MM');
    sh.getRange(i + 2, HUTANG_COL_BULAN).setNumberFormat('@').setValue(bulan);
    diisi++;
    if (contoh.length < 5) contoh.push(r[HUTANG_COL_PB - 1] + '=' + bulan);
  });
  return { ok: true, diisi: diisi, contoh: contoh };
}
