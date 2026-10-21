/* ======================================================
   GatewayTransaksi.js   (PROJECT: database_SANFELO)
   ------------------------------------------------------
   FILE BARU (migrasi TRANSAKSI + DETAIL_TRANSAKSI, Oktober 2026).
   Tambahkan sebagai file terpisah di project Apps Script database_SANFELO,
   di sebelah DatabaseGateway.js & GatewayInvoiceCetak.js.

   Isinya:
     A) 4 aksi gateway baru (didaftarkan di ACTION_HANDLERS DatabaseGateway.js):
        - simpanTransaksiDraft     : versi database dari bagian tulis di
                                     simpanDraft() (Transaction.js inventory)
        - tandaiDoneTransaksi      : versi database dari tandaiDoneByNoInv_()
                                     (StatusDone.js) -- dipanggil saat LUNAS
        - ubahStatusTransaksi      : menu baru "✏️ Ubah Status Transaksi"
                                     (pengganti dropdown STATUS ORDER yang
                                     dulu diketik langsung di sheet)
        - updateTempoTransaksiBatch: versi database dari bagian tulis di
                                     perbaikiTempoTransaksiLama()
     B) salinDataAwalTransaksiPenjualan() -- JALANKAN MANUAL SEKALI dari
        editor database_SANFELO. Menyalin TRANSAKSI, DETAIL_TRANSAKSI, dan
        PENJUALAN yang ASLI dari inventory_SANFELO ke database ini (salinan
        lama di database disimpan dulu sebagai *_LAMA_<tanggal>).

   STRUKTUR KOLOM (sama persis dengan sheet di inventory):
   TRANSAKSI (A..Q, 17 kolom):
     A ID, B NO SJ, C NO INV, D TANGGAL, E CUSTOMER, F ALAMAT, G TEMPO,
     H STATUS, I LAST UPDATE, J METHOD, K NO POL, L SALES, M TOTAL,
     N STATUS ORDER, O BIAYA TAMBAH, P KETERANGAN, Q NO PO
   DETAIL_TRANSAKSI (A..U, 21 kolom):
     A NO SJ, B TANGGAL, C NO INV, D NAMA TOKO, E TEMPO, F KODE, G BARANG,
     H QTY, I SAT, J KETERANGAN, K CREATED_AT, L UPDATED_AT, M HARGA,
     N DISKON, O HARGA NET, P BARIS, Q METHOD, R HPP BARANG,
     S LABA (RUMUS =O-R), T STATUS ORDER, U SOURCE_PO (JSON)
   ====================================================== */

const TRX_SHEET_ = 'TRANSAKSI';
const TRX_COLS_ = 17;            // A..Q
const TRX_COL_ID_ = 1;           // A
const TRX_COL_NOSJ_ = 2;         // B
const TRX_COL_NOINV_ = 3;        // C
const TRX_COL_TGL_ = 4;          // D
const TRX_COL_TEMPO_ = 7;        // G
const TRX_COL_LASTUPDATE_ = 9;   // I
const TRX_COL_STATUS_ORDER_ = 14; // N
const TRX_COL_KETERANGAN_ = 16;  // P
const TRX_COL_NOPO_ = 17;        // Q
const TRX_STATUS_BOLEH_ = ['PROGRESS', 'DONE', 'REVISI'];

const DET_SHEET_ = 'DETAIL_TRANSAKSI';
const DET_COLS_ = 21;            // A..U
const DET_COLS_DATA_ = 18;       // A..R (yang dikirim dari inventory)
const DET_COL_NOSJ_ = 1;         // A
const DET_COL_NOINV_ = 3;        // C
const DET_COL_KODE_ = 6;         // F
const DET_COL_BARIS_ = 16;       // P
const DET_COL_LABA_ = 19;        // S
const DET_COL_STATUS_ORDER_ = 20; // T
const DET_COL_SOURCE_PO_ = 21;   // U
const DET_STATUS_BOLEH_ = ['PO', 'DONE'];

const TRX_WARNA_DONE_ = '#d9ead3';

// ID spreadsheet inventory_SANFELO (sumber salinan awal). Diambil dari URL
// spreadsheet inventory: docs.google.com/spreadsheets/d/<ID>/edit
const INVENTORY_SS_ID_ = '1Yb6UPuTxJ9CwhRRGpLfBNXnlDKgffkgzuAnetVUcsrw';

/* ---------- helper ---------- */

function cariBarisTerakhirBerisi_(sh, kolom) {
  const last = sh.getLastRow();
  if (last < 2) return 1;
  const vals = sh.getRange(1, kolom, last, 1).getValues();
  for (let i = vals.length - 1; i >= 0; i--) {
    if (String(vals[i][0]).trim() !== '') return i + 1;
  }
  return 1;
}

function rumusLabaDetail_(row) {
  return '=IFERROR(O' + row + '-R' + row + ';"")';
}

function statusPenjualanDiDatabase_(ss, noInv) {
  const sh = ss.getSheetByName(PENJUALAN_SHEET);
  if (!sh || sh.getLastRow() < 2) return '';
  const data = sh.getRange(2, PENJUALAN_COL_NOINV, sh.getLastRow() - 1,
                           PENJUALAN_COL_STATUS - PENJUALAN_COL_NOINV + 1).getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0] || '').trim() === noInv) {
      return String(data[i][PENJUALAN_COL_STATUS - PENJUALAN_COL_NOINV] || '').trim().toUpperCase();
    }
  }
  return '';
}

/* ------------------------------------------------------
   AKSI: simpanTransaksiDraft
   body: {
     header: { noSJ, noInv, tgl, toko, daerah, tempo, now, method, noPol,
               sales, total, keterangan, noPO },
     items:    [[18 nilai A..R], ...],   // urutan sama dgn simpanDraft() lama
     sourcePO: ["<json>", ...]           // sejajar dengan items -> kolom U
   }
   Perilaku sama persis dengan simpanDraft() lama:
   - NO SJ belum ada  -> baris TRANSAKSI baru (ID = max ID + 1),
                         H = "DRAFT", N = "PROGRESS"
   - NO SJ sudah ada  -> B..N ditimpa (N kembali "PROGRESS"), P & Q ditimpa,
                         semua baris DETAIL_TRANSAKSI NO SJ itu dihapus
   - item ditulis di bawah DETAIL_TRANSAKSI, S = rumus LABA, U = SOURCE_PO
   ------------------------------------------------------ */
function aksiSimpanTransaksiDraft_(ss, body) {
  const h = body.header || {};
  const items = body.items;
  const sourcePO = Array.isArray(body.sourcePO) ? body.sourcePO : [];

  const noSJ = String(h.noSJ || '').trim();
  const noInv = String(h.noInv || '').trim();
  if (!noSJ) throw new Error('No SJ kosong.');
  if (!noInv) throw new Error('No Invoice kosong.');
  if (!String(h.toko || '').trim()) throw new Error('Nama toko kosong.');
  if (!Array.isArray(items) || items.length === 0) {
    throw new Error('Tidak ada barang terisi. Draft dibatalkan agar data lama tidak terhapus.');
  }

  const trx = ss.getSheetByName(TRX_SHEET_);
  const det = ss.getSheetByName(DET_SHEET_);
  if (!trx) throw new Error('Sheet "' + TRX_SHEET_ + '" tidak ditemukan di database.');
  if (!det) throw new Error('Sheet "' + DET_SHEET_ + '" tidak ditemukan di database.');

  const tgl = parseTanggalDariGateway_(h.tgl);
  if (!(tgl instanceof Date)) throw new Error('Tanggal transaksi tidak valid.');
  const tempo = parseTanggalDariGateway_(h.tempo);
  const now = parseTanggalDariGateway_(h.now) || new Date();

  /* ---- TRANSAKSI ---- */
  const lastTrx = trx.getLastRow();
  const dataTrx = lastTrx >= 2 ? trx.getRange(2, 1, lastTrx - 1, TRX_COL_NOSJ_).getValues() : [];
  let rowTrx = -1;
  let maxId = 0;
  dataTrx.forEach(function (r, i) {
    const id = Number(r[TRX_COL_ID_ - 1]);
    if (!isNaN(id) && id > maxId) maxId = id;
    if (rowTrx === -1 && String(r[TRX_COL_NOSJ_ - 1] || '').trim() === noSJ) rowTrx = i + 2;
  });

  const isiBN = [
    noSJ, noInv, tgl, h.toko || '', h.daerah || '', tempo || '', 'DRAFT', now,
    h.method || '', h.noPol || '', h.sales || '', h.total === '' ? '' : (Number(h.total) || 0), 'PROGRESS'
  ];

  let hasil;
  if (rowTrx === -1) {
    rowTrx = cariBarisTerakhirBerisi_(trx, TRX_COL_NOSJ_) + 1;
    trx.getRange(rowTrx, 1, 1, 14).setValues([[maxId + 1].concat(isiBN)]);
    hasil = 'baru';
  } else {
    trx.getRange(rowTrx, 2, 1, 13).setValues([isiBN]);
    hasil = 'update';
  }
  trx.getRange(rowTrx, TRX_COL_KETERANGAN_).setValue(h.keterangan || '');
  trx.getRange(rowTrx, TRX_COL_NOPO_).setValue(h.noPO || '');
  trx.getRange(rowTrx, TRX_COL_TGL_).setNumberFormat('dd/mm/yyyy');
  trx.getRange(rowTrx, TRX_COL_TEMPO_).setNumberFormat('dd/mm/yyyy');
  trx.getRange(rowTrx, TRX_COL_LASTUPDATE_).setNumberFormat('dd/mm/yyyy hh:mm:ss');
  trx.getRange(rowTrx, 1, 1, TRX_COL_STATUS_ORDER_).setBackground(null); // status balik PROGRESS -> bukan hijau

  /* ---- DETAIL_TRANSAKSI: hapus item lama NO SJ ini ---- */
  const lastDet = det.getLastRow();
  if (lastDet >= 2) {
    const kolA = det.getRange(2, DET_COL_NOSJ_, lastDet - 1, 1).getValues();
    for (let i = kolA.length - 1; i >= 0; i--) {
      if (String(kolA[i][0] || '').trim() === noSJ) det.deleteRow(i + 2);
    }
  }

  /* ---- DETAIL_TRANSAKSI: tulis item baru ---- */
  const rows = items.map(function (r) {
    const row = r.slice(0, DET_COLS_DATA_);
    while (row.length < DET_COLS_DATA_) row.push('');
    row[1] = parseTanggalDariGateway_(row[1]);   // B TANGGAL
    row[4] = parseTanggalDariGateway_(row[4]);   // E TEMPO
    row[10] = parseTanggalDariGateway_(row[10]); // K CREATED_AT
    row[11] = parseTanggalDariGateway_(row[11]); // L UPDATED_AT
    return row.map(function (v) { return v === null || v === undefined ? '' : v; });
  });

  const startDet = cariBarisTerakhirBerisi_(det, DET_COL_NOSJ_) + 1;
  det.getRange(startDet, 1, rows.length, DET_COLS_DATA_).setValues(rows);
  det.getRange(startDet, 2, rows.length, 1).setNumberFormat('dd/mm/yyyy');
  det.getRange(startDet, 5, rows.length, 1).setNumberFormat('dd/mm/yyyy');
  det.getRange(startDet, 11, rows.length, 2).setNumberFormat('dd/mm/yyyy hh:mm:ss');
  det.getRange(startDet, 1, rows.length, DET_COL_STATUS_ORDER_).setBackground(null);

  const rumus = [], src = [];
  for (let j = 0; j < rows.length; j++) {
    rumus.push([rumusLabaDetail_(startDet + j)]);
    src.push([sourcePO[j] || '']);
  }
  det.getRange(startDet, DET_COL_LABA_, rows.length, 1).setFormulas(rumus);
  det.getRange(startDet, DET_COL_STATUS_ORDER_, rows.length, 1).clearContent();
  det.getRange(startDet, DET_COL_SOURCE_PO_, rows.length, 1).setValues(src);

  return { ok: true, hasil: hasil, baris: rowTrx, jumlahItem: rows.length };
}

/* ------------------------------------------------------
   AKSI: tandaiDoneTransaksi
   body: { noInvs: ["INV-...", ...] }
   Untuk tiap NO INV: TRANSAKSI kolom N = "DONE" (+ hijau A..N) dan semua
   baris DETAIL_TRANSAKSI kolom T = "DONE" (+ hijau A..T).
   ------------------------------------------------------ */
function aksiTandaiDoneTransaksi_(ss, body) {
  const daftar = (Array.isArray(body.noInvs) ? body.noInvs : [body.noInv])
    .map(function (v) { return String(v || '').trim(); })
    .filter(String);
  if (daftar.length === 0) throw new Error('noInvs kosong.');
  const set = {};
  daftar.forEach(function (v) { set[v] = true; });

  const hitung = { transaksi: 0, detail: 0 };
  [[TRX_SHEET_, TRX_COL_NOINV_, TRX_COL_STATUS_ORDER_, 'transaksi'],
   [DET_SHEET_, DET_COL_NOINV_, DET_COL_STATUS_ORDER_, 'detail']].forEach(function (cfg) {
    const sh = ss.getSheetByName(cfg[0]);
    if (!sh || sh.getLastRow() < 2) return;
    const vals = sh.getRange(2, cfg[1], sh.getLastRow() - 1, 1).getValues();
    vals.forEach(function (r, i) {
      if (!set[String(r[0] || '').trim()]) return;
      sh.getRange(i + 2, cfg[2]).setValue('DONE');
      sh.getRange(i + 2, 1, 1, cfg[2]).setBackground(TRX_WARNA_DONE_);
      hitung[cfg[3]]++;
    });
  });

  return { ok: true, transaksi: hitung.transaksi, detail: hitung.detail };
}

/* ------------------------------------------------------
   AKSI: ubahStatusTransaksi  (menu "✏️ Ubah Status Transaksi")
   body: { sheet: "TRANSAKSI" | "DETAIL_TRANSAKSI", noSJ, baris, kode, status }
   - TRANSAKSI        : baris dicari dari NO SJ (kolom B), status PROGRESS/DONE/REVISI
   - DETAIL_TRANSAKSI : baris dicari dari NO SJ (A) + BARIS (P) + KODE (F),
                        status PO/DONE
   Aturan lama tetap: DONE hanya boleh kalau invoice-nya sudah LUNAS di PENJUALAN.
   ------------------------------------------------------ */
function aksiUbahStatusTransaksi_(ss, body) {
  const namaSheet = String(body.sheet || '').trim();
  const noSJ = String(body.noSJ || '').trim();
  const status = String(body.status || '').trim().toUpperCase();
  if (!noSJ) throw new Error('No SJ kosong.');

  let sh, colStatus, boleh, cocok;
  if (namaSheet === TRX_SHEET_) {
    sh = ss.getSheetByName(TRX_SHEET_);
    colStatus = TRX_COL_STATUS_ORDER_;
    boleh = TRX_STATUS_BOLEH_;
    cocok = function (r) { return String(r[TRX_COL_NOSJ_ - 1] || '').trim() === noSJ; };
  } else if (namaSheet === DET_SHEET_) {
    const baris = String(body.baris === undefined ? '' : body.baris).trim();
    const kode = String(body.kode || '').trim();
    sh = ss.getSheetByName(DET_SHEET_);
    colStatus = DET_COL_STATUS_ORDER_;
    boleh = DET_STATUS_BOLEH_;
    cocok = function (r) {
      return String(r[DET_COL_NOSJ_ - 1] || '').trim() === noSJ &&
             String(r[DET_COL_BARIS_ - 1]).trim() === baris &&
             String(r[DET_COL_KODE_ - 1] || '').trim() === kode;
    };
  } else {
    throw new Error('sheet harus TRANSAKSI atau DETAIL_TRANSAKSI.');
  }
  if (boleh.indexOf(status) === -1) throw new Error('Status harus salah satu dari: ' + boleh.join(' / ') + '.');
  if (!sh || sh.getLastRow() < 2) throw new Error('Sheet ' + namaSheet + ' kosong.');

  const data = sh.getRange(2, 1, sh.getLastRow() - 1, colStatus).getValues();
  let rowIdx = -1;
  for (let i = 0; i < data.length; i++) {
    if (cocok(data[i])) { rowIdx = i + 2; break; }
  }
  if (rowIdx === -1) throw new Error('Baris tidak ditemukan di ' + namaSheet + ' database (No SJ ' + noSJ + ').');

  const kolNoInv = namaSheet === TRX_SHEET_ ? TRX_COL_NOINV_ : DET_COL_NOINV_;
  const noInv = String(data[rowIdx - 2][kolNoInv - 1] || '').trim();
  const statusLama = String(data[rowIdx - 2][colStatus - 1] || '').trim();

  if (status === 'DONE' && statusPenjualanDiDatabase_(ss, noInv) !== 'LUNAS') {
    throw new Error('Invoice ' + noInv + ' belum LUNAS di PENJUALAN. Status tidak bisa diubah jadi DONE.');
  }

  sh.getRange(rowIdx, colStatus).setValue(status);
  sh.getRange(rowIdx, 1, 1, colStatus).setBackground(status === 'DONE' ? TRX_WARNA_DONE_ : null);
  return { ok: true, noInv: noInv, statusLama: statusLama, statusBaru: status };
}

/* ------------------------------------------------------
   AKSI: updateTempoTransaksiBatch  (perbaikiTempoTransaksiLama)
   body: { updates: [{ noSJ, tempo }, ...] }  -> TRANSAKSI kolom G
   ------------------------------------------------------ */
function aksiUpdateTempoTransaksiBatch_(ss, body) {
  const updates = Array.isArray(body.updates) ? body.updates : [];
  const sh = ss.getSheetByName(TRX_SHEET_);
  if (!sh || sh.getLastRow() < 2 || updates.length === 0) return { ok: true, updatedCount: 0 };

  const posisi = {};
  sh.getRange(2, TRX_COL_NOSJ_, sh.getLastRow() - 1, 1).getValues().forEach(function (r, i) {
    const k = String(r[0] || '').trim();
    if (k && !posisi[k]) posisi[k] = i + 2;
  });

  let n = 0;
  updates.forEach(function (u) {
    const row = posisi[String(u.noSJ || '').trim()];
    const tempo = parseTanggalDariGateway_(u.tempo);
    if (!row || !(tempo instanceof Date)) return;
    sh.getRange(row, TRX_COL_TEMPO_).setValue(tempo).setNumberFormat('dd/mm/yyyy');
    n++;
  });
  return { ok: true, updatedCount: n };
}

/* ======================================================
   B) SALIN DATA AWAL -- jalankan MANUAL SEKALI dari editor database_SANFELO
   ------------------------------------------------------
   Pilih fungsi salinDataAwalTransaksiPenjualan di dropdown editor, klik Run.
   - Data dibaca dari inventory_SANFELO (sheet TRANSAKSI, DETAIL_TRANSAKSI,
     PENJUALAN yang MASIH data asli / belum jadi salinan IMPORTRANGE).
   - Sheet lama di database disalin dulu jadi <NAMA>_LAMA_<yyyyMMdd_HHmm>
     (disembunyikan), lalu ditimpa dengan data inventory: isi, format angka,
     dan warna latar. Kolom LABA di DETAIL_TRANSAKSI ditulis ulang sebagai rumus.
   - Di akhir muncul ringkasan jumlah baris & total, inventory vs database.
   Kalau sheet inventory ternyata sudah berupa salinan IMPORTRANGE, fungsi
   berhenti (artinya migrasi sudah pernah dijalankan -- jangan diulang).
   ====================================================== */
function salinDataAwalTransaksiPenjualan() {
  // Dijalankan dari EDITOR Apps Script, jadi SpreadsheetApp.getUi() tidak tersedia
  // ("Cannot call SpreadsheetApp.getUi() from this context"). Dipakai uiAman_():
  // kalau ada UI (dijalankan dari menu spreadsheet) tampil dialog, kalau tidak
  // (dari editor) cukup dicatat di Log eksekusi. Menjalankan fungsi ini dari editor
  // sudah merupakan tindakan sengaja, jadi konfirmasi Yes/No dilewati di sana.
  const ui = uiAman_();
  if (ui) {
    const ok = ui.alert('Salin data awal dari inventory?',
      'TRANSAKSI, DETAIL_TRANSAKSI, dan PENJUALAN di database ini akan DITIMPA dengan data dari inventory_SANFELO.\n' +
      'Salinan lama di database disimpan dulu sebagai sheet *_LAMA_<tanggal>.\n\n' +
      'Pastikan admin sedang TIDAK memakai inventory. Lanjutkan?', ui.ButtonSet.YES_NO);
    if (ok !== ui.Button.YES) return;
  } else {
    Logger.log('Dijalankan dari editor: konfirmasi dilewati. Pastikan admin TIDAK sedang memakai inventory.');
  }

  const db = SpreadsheetApp.getActiveSpreadsheet();
  const src = SpreadsheetApp.openById(INVENTORY_SS_ID_);
  const cap = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmm');

  const cfgs = [
    { nama: TRX_SHEET_, kolKunci: TRX_COL_NOSJ_, lebar: TRX_COLS_, kolTotal: 13 },
    { nama: DET_SHEET_, kolKunci: DET_COL_NOSJ_, lebar: DET_COLS_, kolTotal: 15 },
    { nama: PENJUALAN_SHEET, kolKunci: PENJUALAN_COL_NOINV, lebar: 0, kolTotal: PENJUALAN_COL_TOTAL }
  ];

  // cek dulu semua sumber masih data asli (bukan IMPORTRANGE) sebelum menyentuh apa pun
  cfgs.forEach(function (c) {
    const s = src.getSheetByName(c.nama);
    if (!s) throw new Error('Sheet ' + c.nama + ' tidak ada di inventory_SANFELO.');
    if (String(s.getRange('A2').getFormula()).toUpperCase().indexOf('IMPORTRANGE') !== -1) {
      throw new Error('Sheet ' + c.nama + ' di inventory sudah berupa salinan IMPORTRANGE -- migrasi sudah pernah dijalankan. Dibatalkan, tidak ada yang diubah.');
    }
  });

  const laporan = [];
  cfgs.forEach(function (c) {
    const s = src.getSheetByName(c.nama);
    const lastRow = cariBarisTerakhirBerisi_(s, c.kolKunci);
    const lebar = c.lebar || Math.max(s.getLastColumn(), PENJUALAN_COL_CICILAN_START);
    const rng = s.getRange(1, 1, lastRow, lebar);
    const values = rng.getValues();
    const formats = rng.getNumberFormats();
    const warna = rng.getBackgrounds();

    let dest = db.getSheetByName(c.nama);
    if (dest) {
      dest.copyTo(db).setName(c.nama + '_LAMA_' + cap).hideSheet();
    } else {
      dest = db.insertSheet(c.nama);
    }
    dest.clear();
    dest.clearConditionalFormatRules();
    dest.getRange(1, 1, dest.getMaxRows(), dest.getMaxColumns()).clearDataValidations();
    if (dest.getMaxRows() < lastRow + 200) dest.insertRowsAfter(dest.getMaxRows(), lastRow + 200 - dest.getMaxRows());
    if (dest.getMaxColumns() < lebar) dest.insertColumnsAfter(dest.getMaxColumns(), lebar - dest.getMaxColumns());

    const out = dest.getRange(1, 1, lastRow, lebar);
    out.setValues(values);
    out.setNumberFormats(formats);
    out.setBackgrounds(warna);
    dest.getRange(1, 1, 1, lebar).setFontWeight('bold');
    dest.setFrozenRows(1);

    if (c.nama === DET_SHEET_ && lastRow >= 2) {
      const rumus = [];
      for (let i = 1; i < values.length; i++) {
        const adaIsi = String(values[i][0] || '').trim() !== '' || String(values[i][DET_COL_KODE_ - 1] || '').trim() !== '';
        rumus.push([adaIsi ? rumusLabaDetail_(i + 1) : '']);
      }
      dest.getRange(2, DET_COL_LABA_, rumus.length, 1).setFormulas(rumus);
    }
    // dropdown status di database (untuk koreksi manual oleh IT/owner)
    if (c.nama === TRX_SHEET_) {
      dest.getRange(2, TRX_COL_STATUS_ORDER_, dest.getMaxRows() - 1, 1).setDataValidation(
        SpreadsheetApp.newDataValidation().requireValueInList(TRX_STATUS_BOLEH_, true).setAllowInvalid(true).build());
    }
    if (c.nama === DET_SHEET_) {
      dest.getRange(2, DET_COL_STATUS_ORDER_, dest.getMaxRows() - 1, 1).setDataValidation(
        SpreadsheetApp.newDataValidation().requireValueInList(DET_STATUS_BOLEH_, true).setAllowInvalid(true).build());
    }

    // verifikasi: jumlah baris berisi & total, sumber vs tujuan
    SpreadsheetApp.flush();
    const ringkas = function (sh) {
      const n = cariBarisTerakhirBerisi_(sh, c.kolKunci);
      if (n < 2) return { baris: 0, total: 0 };
      const v = sh.getRange(2, 1, n - 1, Math.max(c.kolKunci, c.kolTotal)).getValues();
      let baris = 0, total = 0;
      v.forEach(function (r) {
        if (String(r[c.kolKunci - 1] || '').trim() === '') return;
        baris++;
        total += Number(r[c.kolTotal - 1]) || 0;
      });
      return { baris: baris, total: Math.round(total) };
    };
    const a = ringkas(s), b = ringkas(dest);
    const cocok = a.baris === b.baris && a.total === b.total;
    laporan.push((cocok ? '✅ ' : '❌ ') + c.nama + ': inventory ' + a.baris + ' baris / total ' + a.total +
                 '  vs  database ' + b.baris + ' baris / total ' + b.total);
  });

  const ringkasan = laporan.join('\n') +
    '\n\nKalau semua ✅, lanjut ke langkah berikutnya (pasang salinan di inventory).\n' +
    'Kalau ada ❌, JANGAN lanjut -- kabari dulu.';
  Logger.log('Salin data awal selesai\n' + ringkasan);   // dari editor: lihat di "Log eksekusi"
  if (ui) ui.alert('Salin data awal selesai', ringkasan, ui.ButtonSet.OK);
  else SpreadsheetApp.getActiveSpreadsheet().toast('Selesai. Lihat hasilnya di Log eksekusi.', 'Salin data awal', 10);
}

/** Ambil UI kalau tersedia (dijalankan dari spreadsheet), null kalau dari editor. */
function uiAman_() {
  try { return SpreadsheetApp.getUi(); } catch (e) { return null; }
}