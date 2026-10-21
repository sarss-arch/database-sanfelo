/* ======================================================
   GatewayInvoiceCetak.js   (PROJECT: database_SANFELO)
   ------------------------------------------------------
   FILE BARU -- tambahkan sebagai file terpisah di project Apps Script
   database_SANFELO (sebelah DatabaseGateway.js). Isinya 3 aksi gateway
   untuk mencatat invoice yang "SIAP CETAK":

     - cekStatusInvoiceCetak  (baca saja)
     - catatInvoiceTercetak   (catat cetak pertama / cetak ulang)
     - batalInvoiceTercetak   (batalkan tanda tercetak)

   Data disimpan di 2 sheet (dibuat otomatis kalau belum ada):
     - LOG_INVOICE_CETAK   : 1 baris per invoice (dibaca tab LAPORAN INVOICE)
     - LOG_INVOICE_RIWAYAT : 1 baris per klik (jejak audit)

   SETELAH file ini disimpan, 2 hal WAJIB:
   1) Daftarkan 3 aksi di ACTION_HANDLERS (DatabaseGateway.js), ikuti
      format baris lain yang sudah ada di situ:
          cekStatusInvoiceCetak: aksiCekStatusInvoiceCetak_,
          catatInvoiceTercetak:  aksiCatatInvoiceTercetak_,
          batalInvoiceTercetak:  aksiBatalInvoiceTercetak_,
   2) Redeploy gateway sebagai New version (Web App).
   ====================================================== */

const LOG_INV_SHEET_          = 'LOG_INVOICE_CETAK';
const LOG_INV_RIWAYAT_SHEET_  = 'LOG_INVOICE_RIWAYAT';

// LOG_INVOICE_CETAK: A NO INV, B TGL INVOICE, C CUSTOMER, D TOTAL, E BULAN (teks yyyy-MM),
//   F TANDA PERTAMA, G TANDA TERAKHIR, H JUMLAH CETAK, I STATUS, J ALASAN TERAKHIR
const LOG_INV_HEADERS_ = ['NO INV', 'TGL INVOICE', 'CUSTOMER', 'TOTAL', 'BULAN',
                          'TANDA PERTAMA', 'TANDA TERAKHIR', 'JUMLAH CETAK', 'STATUS', 'ALASAN TERAKHIR'];
// LOG_INVOICE_RIWAYAT: A WAKTU, B NO INV, C AKSI, D ALASAN, E TOTAL SEBELUM, F TOTAL SESUDAH, G CUSTOMER
const LOG_INV_RIWAYAT_HEADERS_ = ['WAKTU', 'NO INV', 'AKSI', 'ALASAN', 'TOTAL SEBELUM', 'TOTAL SESUDAH', 'CUSTOMER'];

/**
 * Jalankan SEKALI dari editor database_SANFELO (opsional) supaya kedua
 * sheet log langsung ada sebelum tab LAPORAN INVOICE dipasang di inventory.
 */
function setupLogInvoiceCetak() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  siapkanSheetLogInvoice_(ss, LOG_INV_SHEET_, LOG_INV_HEADERS_);
  siapkanSheetLogInvoice_(ss, LOG_INV_RIWAYAT_SHEET_, LOG_INV_RIWAYAT_HEADERS_);
  SpreadsheetApp.getUi().alert('Sheet LOG_INVOICE_CETAK dan LOG_INVOICE_RIWAYAT siap.');
}

/* ---------- helper ---------- */

// Handler gateway bisa menerima (body) atau (ss, body) tergantung konvensi
// DatabaseGateway.js -- ambil argumen yang berupa body (bukan Spreadsheet).
function ambilBodyInvoiceCetak_(a, b) {
  if (b !== undefined && b !== null && typeof b === 'object' && typeof b.getSheets !== 'function') return b;
  return a || {};
}

function siapkanSheetLogInvoice_(ss, nama, headers) {
  let sh = ss.getSheetByName(nama);
  if (sh) return sh;
  sh = ss.insertSheet(nama);
  sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#e8eaed');
  sh.setFrozenRows(1);
  const maxRows = sh.getMaxRows();
  if (nama === LOG_INV_SHEET_) {
    sh.getRange(1, 1, maxRows, 1).setNumberFormat('@');           // NO INV teks
    sh.getRange(1, 5, maxRows, 1).setNumberFormat('@');           // BULAN teks ("2026-10" jangan jadi tanggal)
    sh.getRange(2, 2, maxRows - 1, 1).setNumberFormat('dd/mm/yyyy');
    sh.getRange(2, 4, maxRows - 1, 1).setNumberFormat('#,##0');
    sh.getRange(2, 6, maxRows - 1, 2).setNumberFormat('dd/mm/yyyy hh:mm:ss');
  } else {
    sh.getRange(2, 1, maxRows - 1, 1).setNumberFormat('dd/mm/yyyy hh:mm:ss');
    sh.getRange(1, 2, maxRows, 1).setNumberFormat('@');
    sh.getRange(2, 5, maxRows - 1, 2).setNumberFormat('#,##0');
  }
  return sh;
}

function cariBarisInvoiceLog_(sh, noInv) {
  const last = sh.getLastRow();
  if (last < 2) return -1;
  const vals = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === noInv) return i + 2;
  }
  return -1;
}

function parseTglInvoice_(s) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(s || ''));
  if (!m) throw new Error('Format tanggal invoice tidak valid: ' + s);
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
}

function catatRiwayatInvoice_(ss, noInv, aksi, alasan, totalSebelum, totalSesudah, customer) {
  const sh = siapkanSheetLogInvoice_(ss, LOG_INV_RIWAYAT_SHEET_, LOG_INV_RIWAYAT_HEADERS_);
  const r = sh.getLastRow() + 1;
  sh.getRange(r, 1, 1, 7).setValues([[new Date(), noInv, aksi, alasan || '', totalSebelum, totalSesudah, customer || '']]);
}

function fmtWaktuInvoice_(d) {
  if (!(d instanceof Date)) return '';
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
}

/* ---------- aksi ---------- */

function aksiCekStatusInvoiceCetak_(a, b) {
  const body = ambilBodyInvoiceCetak_(a, b);
  const noInv = String(body.noInv || '').trim();
  if (!noInv) throw new Error('noInv kosong.');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(LOG_INV_SHEET_);
  if (!sh) return { ok: true, ada: false };
  const row = cariBarisInvoiceLog_(sh, noInv);
  if (row === -1) return { ok: true, ada: false };
  const v = sh.getRange(row, 1, 1, 10).getValues()[0];
  return {
    ok: true, ada: true,
    status: String(v[8]),
    jumlah: Number(v[7]) || 0,
    total: Number(v[3]) || 0,
    tandaTerakhir: fmtWaktuInvoice_(v[6])
  };
}

function aksiCatatInvoiceTercetak_(a, b) {
  const body = ambilBodyInvoiceCetak_(a, b);
  const noInv    = String(body.noInv || '').trim();
  const customer = String(body.customer || '').trim();
  const total    = Number(body.total);
  const alasan   = String(body.alasan || '').trim();
  if (!noInv)    throw new Error('NO INV kosong.');
  if (!customer) throw new Error('Customer kosong.');
  if (!(total > 0)) throw new Error('Total invoice harus lebih dari 0.');
  const tgl = parseTglInvoice_(body.tglInvoice);
  const bulan = Utilities.formatDate(tgl, Session.getScriptTimeZone(), 'yyyy-MM');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = siapkanSheetLogInvoice_(ss, LOG_INV_SHEET_, LOG_INV_HEADERS_);
  const row = cariBarisInvoiceLog_(sh, noInv);
  const now = new Date();

  if (row === -1) {
    const r = sh.getLastRow() + 1;
    // ⬅️ FIX: alasan ikut disimpan (dulu selalu ''), supaya catatan seperti
    // "Dicatat susulan: ..." dari menu catat susulan tidak hilang.
    sh.getRange(r, 1, 1, 10).setValues([[noInv, tgl, customer, total, bulan, now, now, 1, 'TERCETAK', alasan]]);
    catatRiwayatInvoice_(ss, noInv, 'CETAK', alasan, 0, total, customer);
    return { ok: true, hasil: 'baru', jumlah: 1 };
  }

  const lama = sh.getRange(row, 1, 1, 10).getValues()[0];
  const totalSebelum = Number(lama[3]) || 0;
  const statusLama = String(lama[8]);
  const jumlahBaru = (Number(lama[7]) || 0) + 1;

  if (statusLama === 'TERCETAK' && !alasan) {
    throw new Error('Invoice ini sudah pernah ditandai tercetak. Alasan cetak ulang wajib diisi.');
  }

  // pertahankan TANDA PERTAMA (kolom F), perbarui sisanya
  sh.getRange(row, 2, 1, 4).setValues([[tgl, customer, total, bulan]]);       // B:E
  sh.getRange(row, 7, 1, 4).setValues([[now, jumlahBaru, 'TERCETAK', alasan]]); // G:J
  catatRiwayatInvoice_(ss, noInv, statusLama === 'BATAL' ? 'CETAK LAGI (setelah batal)' : 'CETAK ULANG',
                       alasan, totalSebelum, total, customer);
  return { ok: true, hasil: statusLama === 'BATAL' ? 'aktif_lagi' : 'ulang', jumlah: jumlahBaru, totalSebelum: totalSebelum };
}

function aksiBatalInvoiceTercetak_(a, b) {
  const body = ambilBodyInvoiceCetak_(a, b);
  const noInv  = String(body.noInv || '').trim();
  const alasan = String(body.alasan || '').trim();
  if (!noInv)  throw new Error('NO INV kosong.');
  if (!alasan) throw new Error('Alasan pembatalan wajib diisi.');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(LOG_INV_SHEET_);
  const row = sh ? cariBarisInvoiceLog_(sh, noInv) : -1;
  if (row === -1) throw new Error('Invoice ' + noInv + ' belum pernah ditandai tercetak.');

  const v = sh.getRange(row, 1, 1, 10).getValues()[0];
  if (String(v[8]) !== 'TERCETAK') throw new Error('Invoice ' + noInv + ' sudah berstatus ' + v[8] + '.');

  sh.getRange(row, 7, 1, 4).setValues([[new Date(), v[7], 'BATAL', alasan]]);
  catatRiwayatInvoice_(ss, noInv, 'BATAL', alasan, Number(v[3]) || 0, 0, String(v[2]));
  return { ok: true };
}