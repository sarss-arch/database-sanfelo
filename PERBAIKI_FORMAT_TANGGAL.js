/* ======================================================
   PERBAIKI_FORMAT_TANGGAL.js   (PROJECT: database_SANFELO)
   ------------------------------------------------------
   JALANKAN MANUAL SEKALI dari editor: pilih perbaikiFormatTanggalDetail, Run.

   Masalah: di DETAIL_TRANSAKSI baris 98-346 kolom B (TANGGAL) & E (TEMPO),
   juga TRANSAKSI kolom I (LAST UPDATE), isinya tanggal yang tersimpan
   sebagai ANGKA (mis. 46221,29 = 18/07/2026) dengan format "Otomatis".
   Waktu salinDataAwalTransaksiPenjualan() menyalin dari inventory, format
   tanggal sel-sel itu tidak ikut, jadi di database kelihatan seperti angka
   / tanggalnya "hilang". Di inventory tetap tampil karena kolom salinan
   IMPORTRANGE di sana sudah diformat dd/mm/yyyy.

   Fungsi ini HANYA memasang format tanggal per kolom (isi sel tidak diubah
   sama sekali), jadi aman dan bisa diulang.
   ====================================================== */

function perbaikiFormatTanggalDetail() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const target = [
    { sheet: 'DETAIL_TRANSAKSI', kolom: { 2: 'dd/mm/yyyy', 5: 'dd/mm/yyyy', 11: 'dd/mm/yyyy hh:mm:ss', 12: 'dd/mm/yyyy hh:mm:ss' } },
    { sheet: 'TRANSAKSI',        kolom: { 4: 'dd/mm/yyyy', 7: 'dd/mm/yyyy', 9: 'dd/mm/yyyy hh:mm:ss' } }
  ];
  const hasil = [];
  target.forEach(function (t) {
    const sh = ss.getSheetByName(t.sheet);
    if (!sh) { hasil.push('❌ ' + t.sheet + ' tidak ditemukan'); return; }
    const n = sh.getMaxRows() - 1;
    Object.keys(t.kolom).forEach(function (k) {
      sh.getRange(2, Number(k), n, 1).setNumberFormat(t.kolom[k]);
    });
    hasil.push('✅ ' + t.sheet + ': format tanggal dipasang di kolom ' +
      Object.keys(t.kolom).map(function (k) { return String.fromCharCode(64 + Number(k)); }).join(', '));
  });
  Logger.log(hasil.join('\n'));
  ss.toast(hasil.join(' | '), 'Perbaiki format tanggal', 10);
}
