/* ======================================================
   PERBAIKI_HUTANG.js   (PROJECT: database_SANFELO)
   ------------------------------------------------------
   PERBAIKAN SEKALI JALAN (aman diulang), Oktober 2026:

   1. PEMBAYARAN HUTANG KECATAT DOBEL (mis. PB-2026-07-046/047/048):
      admin klik Bayar lagi setelah muncul "database sedang sibuk",
      padahal pembayaran pertama sebenarnya SUDAH tersimpan. Hasilnya
      riwayat yang sama persis tertulis 2x (kolom K & L) dan SISA BAYAR
      jadi "LEBIH RpX".
      Diperbaiki: di baris yang SISA BAYAR-nya "LEBIH", riwayat yang
      teksnya SAMA PERSIS dibuang (yang pertama disimpan), lalu SISA
      BAYAR dihitung ulang = TOTAL - jumlah riwayat yang tersisa, dan
      STATUS + warna di-refresh. Baris yang tidak LEBIH tidak disentuh
      (supaya cicilan asli dengan nominal & tanggal sama tidak ikut hilang).

   2. TYPO NAMA SUPPLIER: "PT ANUGERAH PUYRA SIANTAN" & "PT ANUGERAH PUTRA
      SIANTANO" -> "PT ANUGERAH PUTRA SIANTAN", di HUTANG (kolom E) DAN
      PEMBELIAN (kolom D). PEMBELIAN harus ikut, karena supplier di HUTANG
      ditulis ulang dari PEMBELIAN tiap kali Simpan SJ Beli.

   CARA PAKAI: di editor, pilih perbaikiHutangDobelDanTypo -> Run, lalu
   lihat hasilnya di "Log eksekusi". Mau lihat dulu tanpa mengubah apa-apa:
   jalankan cekHutangDobelDanTypo (cuma laporan).
   ====================================================== */

const PERBAIKI_SUPPLIER_TYPO_ = {
  'PT ANUGERAH PUYRA SIANTAN': 'PT ANUGERAH PUTRA SIANTAN',
  'PT ANUGERAH PUTRA SIANTANO': 'PT ANUGERAH PUTRA SIANTAN'
};

function cekHutangDobelDanTypo() { perbaikiHutangInti_(true); }
function perbaikiHutangDobelDanTypo() { perbaikiHutangInti_(false); }

function perbaikiHutangInti_(cumaCek) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const log = [];

  /* ---------- 1. pembayaran dobel ---------- */
  const hutang = ss.getSheetByName(HUTANG_SHEET);
  const lastRow = hutang.getLastRow();
  const lastCol = Math.max(hutang.getLastColumn(), HUTANG_RIWAYAT_START_COL);
  const data = lastRow >= 2 ? hutang.getRange(2, 1, lastRow - 1, lastCol).getValues() : [];
  const barisDiubah = [];

  data.forEach(function (r, i) {
    const sisaRaw = r[HUTANG_COL_SISABAYAR - 1];
    if (String(sisaRaw || '').trim().toUpperCase().indexOf('LEBIH') !== 0) return;

    const riwayat = r.slice(HUTANG_RIWAYAT_START_COL - 1).map(function (v) { return String(v || '').trim(); }).filter(String);
    const unik = riwayat.filter(function (t, k) { return riwayat.indexOf(t) === k; });
    if (unik.length === riwayat.length) {
      log.push('⚠️ ' + r[HUTANG_COL_PB - 1] + ': SISA ' + sisaRaw + ' tapi tidak ada riwayat dobel -- cek manual.');
      return;
    }

    const total = Number(r[HUTANG_COL_TOTAL - 1]) || 0;
    const dibayar = unik.reduce(function (s, t) {
      const bagian = t.split(' - ');
      return s + (Number(String(bagian[1] || '').replace(/[^\d]/g, '')) || 0);
    }, 0);
    const sisa = total - dibayar;
    const sisaBaru = sisa < 0 ? 'LEBIH ' + formatRupiahGatewayHutang_(-sisa) : sisa;
    const rowIdx = i + 2;
    log.push((cumaCek ? '🔎 ' : '✅ ') + r[HUTANG_COL_PB - 1] + ' (baris ' + rowIdx + '): ' + (riwayat.length - unik.length) +
             ' riwayat dobel dibuang, SISA ' + sisaRaw + ' -> ' + (typeof sisaBaru === 'number' ? sisaBaru : sisaBaru));

    if (cumaCek) return;
    const isiRiwayat = [];
    for (let c = HUTANG_RIWAYAT_START_COL; c <= lastCol; c++) isiRiwayat.push(unik[c - HUTANG_RIWAYAT_START_COL] || '');
    hutang.getRange(rowIdx, HUTANG_RIWAYAT_START_COL, 1, isiRiwayat.length).setValues([isiRiwayat]);
    const sel = hutang.getRange(rowIdx, HUTANG_COL_SISABAYAR).setValue(sisaBaru);
    if (typeof sisaBaru === 'number') sel.setNumberFormat('#,##0');
    barisDiubah.push(rowIdx);
  });
  if (barisDiubah.length) updateStatusJatuhTempoGatewayHutang_(hutang, barisDiubah);

  /* ---------- 2. typo supplier di HUTANG (E) & PEMBELIAN (D) ---------- */
  [[HUTANG_SHEET, HUTANG_COL_SUPPLIER], [PEMBELIAN_SHEET, PEMBELIAN_COL_SUPPLIER]].forEach(function (t) {
    const sh = ss.getSheetByName(t[0]);
    if (!sh || sh.getLastRow() < 2) return;
    const rg = sh.getRange(2, t[1], sh.getLastRow() - 1, 1);
    const vals = rg.getValues();
    let n = 0;
    const baru = vals.map(function (v) {
      const k = String(v[0] || '').replace(/ /g, ' ').trim().toUpperCase();
      if (PERBAIKI_SUPPLIER_TYPO_[k]) { n++; return [PERBAIKI_SUPPLIER_TYPO_[k]]; }
      return [v[0]];
    });
    if (n && !cumaCek) rg.setValues(baru);
    log.push((cumaCek ? '🔎 ' : '✅ ') + t[0] + ': ' + n + ' nama supplier typo' + (cumaCek ? ' ditemukan' : ' dibetulkan') + '.');
  });

  Logger.log((cumaCek ? 'CEK SAJA (belum ada yang diubah)\n' : 'SELESAI\n') + log.join('\n'));
}
