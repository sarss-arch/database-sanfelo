/* ======================================================
   GatewayQtyTerkirim.js   (PROJECT: database_SANFELO)
   ------------------------------------------------------
   AKSI: setQtyTerkirimPoKeluarBatch -- tulis ulang kolom H (QTY TERKIRIM)
   HISTORY_PO_KELUAR di database dengan hasil hitung ulang dari inventory
   (QTY_TERKIRIM_ULANG.js). Dipakai supaya angka di database sama dengan
   yang tampil di inventory (di inventory kolom H tidak ikut IMPORTRANGE).

   Beda dengan aksi lama updateQtyTerkirimPoKeluar: aksi lama mencari per
   NO PO + TOKO + BARANG dan menulis angka yang SAMA ke semua baris yang
   cocok -- salah untuk barang yang dipecah per warna (KS 001 BIRU/MERAH/
   COKLAT). Aksi ini menulis PER BARIS, berurutan dari baris 2.

   PENGAMAN: tiap baris dikirim beserta NO PO + NAMA BARANG + KETERANGAN-
   nya. Kalau ada satu saja yang tidak cocok dengan isi database saat ini
   (mis. PO keluar baru saja disimpan ulang di sela-sela), TIDAK ADA yang
   ditulis -- inventory tinggal hitung ulang lagi.

   PASANG (sekali):
   1. Tambahkan file ini ke project database_SANFELO.
   2. Di Databasegateway.js, di dalam const ACTION_HANDLERS = { ... },
      tambahkan satu baris:
        setQtyTerkirimPoKeluarBatch: function (ss, body) { return aksiSetQtyTerkirimPoKeluarBatch_(ss, body); },
   3. Terapkan > Kelola deployment > edit > Versi: Versi baru > Terapkan.

   body: { rows: [{ po, barang, ket, qty }, ...] }  -- rows[0] = baris 2
   ====================================================== */

function aksiSetQtyTerkirimPoKeluarBatch_(ss, body) {
  const rows = body && body.rows;
  if (!Array.isArray(rows)) throw new Error('rows harus array.');

  const sh = ss.getSheetByName(HISTORY_PO_KELUAR_SHEET);
  if (!sh) throw new Error('Sheet "' + HISTORY_PO_KELUAR_SHEET + '" tidak ditemukan di database.');

  const lastRow = sh.getLastRow();
  const n = Math.max(lastRow - 1, 0);
  if (rows.length !== n) {
    throw new Error('Jumlah baris berubah (inventory ' + rows.length + ', database ' + n + '). Hitung ulang lagi.');
  }
  if (n === 0) return { ok: true, baris_diubah: 0 };

  const U = function (s) { return String(s === null || s === undefined ? '' : s).trim().toUpperCase(); };
  const data = sh.getRange(2, 1, n, HISTORY_PO_KELUAR_QTY_COL).getValues(); // A..H
  for (let i = 0; i < n; i++) {
    const r = rows[i] || {};
    if (U(r.po) !== U(data[i][0]) || U(r.barang) !== U(data[i][3]) || U(r.ket) !== U(data[i][5])) {
      throw new Error('Baris ' + (i + 2) + ' tidak cocok (database: ' + data[i][0] + ' ' + data[i][3] +
                      ', inventory: ' + r.po + ' ' + r.barang + '). Tidak ada yang diubah, hitung ulang lagi.');
    }
  }

  let diubah = 0;
  const out = data.map(function (d, i) {
    const baru = Number(rows[i].qty) > 0 ? Number(rows[i].qty) : '';
    if ((Number(d[HISTORY_PO_KELUAR_QTY_COL - 1]) || 0) !== (Number(baru) || 0)) diubah++;
    return [baru];
  });
  sh.getRange(2, HISTORY_PO_KELUAR_QTY_COL, n, 1).setValues(out);
  return { ok: true, baris_diubah: diubah };
}
