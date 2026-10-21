/* ======================================================
   GatewayDiskonDataBersih.js   (PROJECT: database_SANFELO)
   ------------------------------------------------------
   FILE BARU -- migrasi sheet DISKON dan DATA_BERSIH (Oktober 2026).
   Sejak migrasi ini, kedua sheet itu di inventory cuma SALINAN
   (IMPORTRANGE) dari database. Yang mengubah isinya IT/owner langsung
   di sheet database ini.

   TIDAK perlu deploy ulang Web App: file ini cuma fungsi yang dijalankan
   manual dari editor (dan satu onEdit sederhana untuk sort DISKON).

   KETERGANTUNGAN: memakai konstanta INVENTORY_SS_ID_ yang sudah ada di
   GatewayTransaksi.js (project yang sama).

   FUNGSI (jalankan dari editor, urutan A -> B):
   A. cekSelisihDiskonDataBersih()  : BACA SAJA. Membandingkan inventory vs
      database, tidak mengubah apa pun. Jalankan dulu untuk lihat bedanya.
   B. salinDataAwalDiskonDataBersih(): menyalin DISKON & DATA_BERSIH dari
      inventory ke database (rumus ikut tersalin sebagai rumus), dengan
      cadangan *_LAMA_<tanggal> dulu, lalu cek ulang sel demi sel.
   C. rapikanNomorTelpDiskonDB()    : rapikan kolom NO TELP DISKON di database
      (pengganti rapikanNomorTelpDiskon yang dulu jalan di inventory).
   D. sortDiskonDiDatabase()        : sort DISKON A-Z manual (onEdit di bawah
      otomatis tiap ada ketikan).
   ====================================================== */

const DB_MIGRASI_SHEETS_ = ['DISKON', 'DATA_BERSIH'];

/** UI kalau ada (dijalankan dari spreadsheet), null kalau dari editor. */
function uiTersediaDiskon_() {
  try { return SpreadsheetApp.getUi(); } catch (e) { return null; }
}

function laporkanDiskon_(judul, teks) {
  Logger.log(judul + '\n' + teks);
  const ui = uiTersediaDiskon_();
  if (ui) {
    ui.alert(judul, teks, ui.ButtonSet.OK);
  } else {
    try { SpreadsheetApp.getActiveSpreadsheet().toast('Selesai. Lihat hasilnya di Log eksekusi.', judul, 10); } catch (e) { /* abaikan */ }
  }
}

/** Baris terakhir yang kolom kunci (A) terisi. */
function barisTerakhirBerisiDiskon_(sh, kol) {
  const last = sh.getLastRow();
  if (last < 1) return 0;
  const v = sh.getRange(1, kol, last, 1).getValues();
  for (let i = v.length - 1; i >= 0; i--) {
    if (String(v[i][0] === null || v[i][0] === undefined ? '' : v[i][0]).trim() !== '') return i + 1;
  }
  return 0;
}

/** Isi sel berupa objek (mis. gambar di dalam sel / CellImage), bukan teks/angka/tanggal. */
function adalahObjek_(v) {
  return v !== null && v !== undefined && typeof v === 'object' && !(v instanceof Date);
}

function sama_(a, b) {
  if (a === b) return true;
  // Gambar di dalam sel tidak bisa ditulis lewat setValues: dianggap sama dengan sel gambar lain / sel kosong.
  if (adalahObjek_(a) && (adalahObjek_(b) || b === '' || b === null || b === undefined)) return true;
  if (adalahObjek_(b) && (a === '' || a === null || a === undefined)) return true;
  const sa = (a === null || a === undefined) ? '' : a;
  const sb = (b === null || b === undefined) ? '' : b;
  if (sa === '' && sb === '') return true;
  if (typeof sa === 'number' && typeof sb === 'number') return Math.abs(sa - sb) < 1e-9;
  if (sa instanceof Date && sb instanceof Date) return sa.getTime() === sb.getTime();
  return String(sa) === String(sb);
}

/** Bandingkan nilai (hasil hitung) dua sheet sel demi sel. */
function bandingkanSheet_(sA, sB) {
  const lastA = barisTerakhirBerisiDiskon_(sA, 1);
  const lastB = barisTerakhirBerisiDiskon_(sB, 1);
  const baris = Math.max(lastA, lastB);
  const lebar = Math.max(sA.getLastColumn(), sB.getLastColumn(), 1);
  const res = { barisA: lastA - 1, barisB: lastB - 1, beda: 0, contoh: [] };
  if (baris < 1) return res;
  const va = sA.getRange(1, 1, baris, lebar).getValues();
  const vb = sB.getRange(1, 1, baris, lebar).getValues();
  for (let i = 0; i < baris; i++) {
    for (let j = 0; j < lebar; j++) {
      if (!sama_(va[i][j], vb[i][j])) {
        res.beda++;
        if (res.contoh.length < 5) {
          res.contoh.push(sA.getRange(i + 1, j + 1).getA1Notation() + ' (inventory: "' + va[i][j] + '" | database: "' + vb[i][j] + '")');
        }
      }
    }
  }
  return res;
}

/* ------------------------------------------------------
   A) CEK SELISIH -- baca saja, tidak mengubah apa pun
   ------------------------------------------------------ */
function cekSelisihDiskonDataBersih() {
  const db = SpreadsheetApp.getActiveSpreadsheet();
  const src = SpreadsheetApp.openById(INVENTORY_SS_ID_);
  const baris = [];
  DB_MIGRASI_SHEETS_.forEach(function (nama) {
    const s = src.getSheetByName(nama);
    const d = db.getSheetByName(nama);
    if (!s) { baris.push('❌ ' + nama + ': tidak ada di inventory'); return; }
    if (!d) { baris.push('⚠️ ' + nama + ': belum ada di database (akan dibuat saat disalin)'); return; }
    if (String(s.getRange('A2').getFormula()).toUpperCase().indexOf('IMPORTRANGE') !== -1) {
      baris.push('ℹ️ ' + nama + ': di inventory sudah berupa salinan IMPORTRANGE (migrasi sudah dijalankan)');
      return;
    }
    const r = bandingkanSheet_(s, d);
    baris.push((r.beda === 0 ? '✅ ' : '⚠️ ') + nama + ': inventory ' + r.barisA + ' baris, database ' + r.barisB +
      ' baris, ' + r.beda + ' sel berbeda' + (r.contoh.length ? '\n     contoh: ' + r.contoh.join('\n             ') : ''));
  });
  laporkanDiskon_('Cek selisih DISKON & DATA_BERSIH (tidak ada yang diubah)', baris.join('\n'));
}

/* ------------------------------------------------------
   B) SALIN DATA AWAL
   ------------------------------------------------------ */
function salinDataAwalDiskonDataBersih() {
  const ui = uiTersediaDiskon_();
  if (ui) {
    const ok = ui.alert('Salin DISKON & DATA_BERSIH dari inventory?',
      'DISKON dan DATA_BERSIH di database ini akan DITIMPA dengan data dari inventory_SANFELO.\n' +
      'Salinan lama disimpan dulu sebagai sheet *_LAMA_<tanggal>.\n\nPastikan admin sedang TIDAK mengubah DISKON / DATA_BERSIH. Lanjutkan?',
      ui.ButtonSet.YES_NO);
    if (ok !== ui.Button.YES) return;
  } else {
    Logger.log('Dijalankan dari editor: konfirmasi dilewati. Pastikan tidak ada yang sedang mengubah DISKON / DATA_BERSIH di inventory.');
  }

  const db = SpreadsheetApp.getActiveSpreadsheet();
  const src = SpreadsheetApp.openById(INVENTORY_SS_ID_);
  const cap = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmm');

  // 1) Cek semua sumber dulu -- sebelum menyentuh apa pun.
  const rencana = [];
  DB_MIGRASI_SHEETS_.forEach(function (nama) {
    const s = src.getSheetByName(nama);
    if (!s) throw new Error('Sheet ' + nama + ' tidak ada di inventory_SANFELO.');
    if (String(s.getRange('A2').getFormula()).toUpperCase().indexOf('IMPORTRANGE') !== -1) {
      throw new Error('Sheet ' + nama + ' di inventory sudah berupa salinan IMPORTRANGE -- migrasi sudah pernah dijalankan. Dibatalkan, tidak ada yang diubah.');
    }
    const last = barisTerakhirBerisiDiskon_(s, 1);
    if (last < 2) throw new Error('Sheet ' + nama + ' di inventory kosong. Dibatalkan, tidak ada yang diubah.');
    const lebar = Math.max(s.getLastColumn(), 1);
    const rng = s.getRange(1, 1, last, lebar);
    const formulas = rng.getFormulas();
    // Rumus yang menunjuk sheet lain / IMPORTRANGE tidak aman disalin apa adanya.
    for (let i = 0; i < formulas.length; i++) {
      for (let j = 0; j < formulas[i].length; j++) {
        const f = String(formulas[i][j] || '');
        if (f && (f.indexOf('!') !== -1 || f.toUpperCase().indexOf('IMPORTRANGE') !== -1)) {
          throw new Error('Sel ' + rng.getCell(i + 1, j + 1).getA1Notation() + ' di ' + nama + ' berisi rumus yang merujuk sheet lain (' + f.substring(0, 60) + '). Dibatalkan, tidak ada yang diubah. Kabari IT.');
        }
      }
    }
    rencana.push({ nama: nama, sheet: s, last: last, lebar: lebar, rng: rng, formulas: formulas });
  });

  // 2) Salin.
  const laporan = [];
  rencana.forEach(function (c) {
    let dest = db.getSheetByName(c.nama);

    // Sudah sama persis dengan inventory -> dilewati (aman dijalankan ulang).
    if (dest) {
      const pra = bandingkanSheet_(c.sheet, dest);
      if (pra.beda === 0 && pra.barisA > 0 && pra.barisA === pra.barisB) {
        laporan.push('✅ ' + c.nama + ': database sudah SAMA dengan inventory (' + pra.barisA + ' baris), dilewati');
        return;
      }
    }

    const values = c.rng.getValues();
    const formats = c.rng.getNumberFormats();
    const warna = c.rng.getBackgrounds();

    // Semua TEKS ditulis dengan format teks (@) supaya Sheets tidak menafsirkannya:
    // "+kaca", "-", "35+5+5", "0813..." dst. akan dianggap rumus/angka kalau tidak begini.
    // Sel berumus dikosongkan dulu di tabel nilai, lalu rumusnya ditulis terpisah.
    const isi = [];
    const gambar = [];         // sel berisi gambar (tidak bisa ditulis lewat setValues)
    const barisFormula = [];   // [{row, col, list:[[f],[f],...]}]  (potongan kolom yang berurutan)
    for (let i = 0; i < c.last; i++) {
      const baris = [];
      for (let j = 0; j < c.lebar; j++) {
        const f = c.formulas[i][j];
        const v = values[i][j];
        if (f) {
          baris.push('');
        } else if (adalahObjek_(v)) {
          gambar.push(c.rng.getCell(i + 1, j + 1).getA1Notation());   // gambar di dalam sel
          baris.push('');
        } else {
          if (typeof v === 'string' && v !== '') formats[i][j] = '@';
          baris.push(v);
        }
      }
      isi.push(baris);
    }
    for (let j = 0; j < c.lebar; j++) {
      let run = null;
      for (let i = 0; i < c.last; i++) {
        const f = c.formulas[i][j];
        if (f) {
          if (!run) { run = { row: i + 1, col: j + 1, list: [] }; barisFormula.push(run); }
          run.list.push([f]);
        } else {
          run = null;
        }
      }
    }

    let barisLama = 0;
    let namaBackup = '';
    const namaLamaSet = {};
    if (dest) {
      barisLama = Math.max(0, barisTerakhirBerisiDiskon_(dest, 1) - 1);
      if (barisLama > 0) {
        if (c.nama === 'DISKON') {
          dest.getRange(2, 1, barisLama, 1).getValues().forEach(function (r) {
            const k = String(r[0] || '').trim(); if (k) namaLamaSet[k] = true;
          });
        }
        namaBackup = c.nama + '_LAMA_' + cap;
        dest.copyTo(db).setName(namaBackup).hideSheet();
      }
      dest.clear();
    } else {
      dest = db.insertSheet(c.nama);
    }
    if (dest.getMaxRows() < c.last + 200) dest.insertRowsAfter(dest.getMaxRows(), c.last + 200 - dest.getMaxRows());
    if (dest.getMaxColumns() < c.lebar) dest.insertColumnsAfter(dest.getMaxColumns(), c.lebar - dest.getMaxColumns());

    const tujuan = dest.getRange(1, 1, c.last, c.lebar);
    try {
      tujuan.setNumberFormats(formats);
      tujuan.setValues(isi);
      barisFormula.forEach(function (r) {
        dest.getRange(r.row, r.col, r.list.length, 1).setFormulas(r.list);
      });
      tujuan.setBackgrounds(warna);
    } catch (err) {
      // Cari baris/sel penyebabnya (untuk laporan), lalu KEMBALIKAN database seperti semula.
      const gagal = [];
      try {
        dest.getRange(1, 1, c.last, c.lebar).clearContent();
        for (let i = 0; i < c.last && gagal.length < 3; i++) {
          try {
            dest.getRange(i + 1, 1, 1, c.lebar).setValues([isi[i]]);
          } catch (e2) {
            for (let j = 0; j < c.lebar; j++) {
              try { dest.getRange(i + 1, j + 1).setValue(isi[i][j]); }
              catch (e3) { gagal.push(dest.getRange(i + 1, j + 1).getA1Notation() + ' = "' + isi[i][j] + '"'); break; }
            }
          }
        }
      } catch (e4) { /* diagnosa saja, abaikan */ }
      try {
        dest.clear();
        if (namaBackup) {
          const bk = db.getSheetByName(namaBackup);
          bk.getDataRange().copyTo(dest.getRange(1, 1));
        }
      } catch (e5) { Logger.log('Pemulihan ' + c.nama + ' gagal: ' + e5); }
      throw new Error('Gagal menulis ' + c.nama + ' ke database: ' + err.message +
        (gagal.length ? '\nSel bermasalah: ' + gagal.join(' ; ') : '') +
        '\nSheet ' + c.nama + ' di database dikembalikan' + (namaBackup ? ' dari ' + namaBackup : ' (kosong, tadinya memang kosong)') + '. Kabari IT.');
    }
    dest.getRange(1, 1, 1, c.lebar).setFontWeight('bold');
    dest.setFrozenRows(1);

    // Sel gambar: salin lewat salinan sementara sheet sumber (copyTo antar-sel menjaga gambarnya).
    // Kalau gagal, TIDAK menggagalkan penyalinan -- cuma dicatat di laporan.
    if (gambar.length) {
      let tmp = null;
      try {
        tmp = c.sheet.copyTo(db);
        gambar.forEach(function (a1) { tmp.getRange(a1).copyTo(dest.getRange(a1)); });
        laporan.push('🖼️ ' + c.nama + ': ' + gambar.length + ' sel gambar disalin (' + gambar.join(', ') + ')');
      } catch (eg) {
        laporan.push('⚠️ ' + c.nama + ': ' + gambar.length + ' sel gambar TIDAK tersalin (' + gambar.join(', ') + ') -- pasang manual di database kalau perlu. ' + eg.message);
      } finally {
        if (tmp) { try { db.deleteSheet(tmp); } catch (ed) { /* abaikan */ } }
      }
    }
    SpreadsheetApp.flush();

    // 3) Verifikasi: bandingkan nilai sel demi sel (sumber vs tujuan).
    const r = bandingkanSheet_(c.sheet, dest);
    laporan.push((r.beda === 0 && r.barisA === r.barisB ? '✅ ' : '❌ ') + c.nama + ': inventory ' + r.barisA + ' baris vs database ' + r.barisB +
      ' baris, ' + r.beda + ' sel berbeda' + (r.contoh.length ? '\n     contoh: ' + r.contoh.join('\n             ') : '') +
      (barisLama ? '\n     (database sebelumnya: ' + barisLama + ' baris, disimpan di ' + namaBackup + ')' : ''));

    if (c.nama === 'DISKON') {
      const namaBaru = {};
      dest.getRange(2, 1, c.last - 1, 1).getValues().forEach(function (x) {
        const k = String(x[0] || '').trim(); if (k) namaBaru[k] = true;
      });
      const hilang = Object.keys(namaLamaSet).filter(function (k) { return !namaBaru[k]; });
      const tambah = Object.keys(namaBaru).filter(function (k) { return !namaLamaSet[k]; });
      laporan.push('     toko baru di database: ' + tambah.length + (tambah.length ? ' (' + tambah.join(', ') + ')' : ''));
      laporan.push((hilang.length ? '⚠️ ' : '     ') + 'toko yang ada di database lama tapi TIDAK ada di inventory: ' + hilang.length +
        (hilang.length ? ' (' + hilang.join(', ') + ') -- masih ada di sheet ' + namaBackup : ''));
    }
  });

  laporkanDiskon_('Salin DISKON & DATA_BERSIH selesai',
    laporan.join('\n') + '\n\nKalau semua ✅, lanjut: pasangSalinanDiskonDataBersih() di inventory.\nKalau ada ❌, JANGAN lanjut -- kabari dulu.');
}

/* ------------------------------------------------------
   C) RAPIKAN NO TELP (DISKON kolom G) -- di database
   Logika sama dengan rapikanNomorTelpDiskon() yang dulu di inventory.
   Aman dijalankan berulang.
   ------------------------------------------------------ */
function rapikanNomorTelpDiskonDB() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DISKON');
  if (!sh) throw new Error('Sheet "DISKON" tidak ditemukan di database.');
  const last = barisTerakhirBerisiDiskon_(sh, 1);
  if (last < 2) return;

  const range = sh.getRange(2, 7, last - 1, 1); // G
  const values = range.getValues();
  let dirapikan = 0;
  const hasil = values.map(function (row) {
    const v = row[0];
    if (v === '' || v === null || v === undefined) return [''];
    let digit = String(v).replace(/\D/g, '');
    if (digit === '') return [''];
    if (digit.indexOf('62') === 0 && digit.length >= 11) digit = '0' + digit.substring(2);
    if (digit.charAt(0) !== '0') digit = '0' + digit;
    if (String(v) !== digit) dirapikan++;
    return [digit];
  });
  range.setNumberFormat('@');
  range.setValues(hasil);
  laporkanDiskon_('Rapikan no telp DISKON', 'Selesai. ' + dirapikan + ' sel dirapikan.');
}

/* ------------------------------------------------------
   D) SORT DISKON A-Z (nama toko). Hanya kolom A-G ikut disort.
   onEdit di bawah memanggilnya otomatis tiap ada ketikan di DISKON
   (simple trigger -- hanya jalan untuk ketikan manual di file database).
   ------------------------------------------------------ */
function sortDiskonDiDatabase() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DISKON');
  if (!sh) throw new Error('Sheet "DISKON" tidak ditemukan di database.');
  sortDiskonSheet_(sh);
  laporkanDiskon_('Sort DISKON', 'Sheet DISKON di-sort A-Z berdasarkan Nama Toko.');
}

function sortDiskonSheet_(sh) {
  const last = barisTerakhirBerisiDiskon_(sh, 1);
  if (last < 3) return;
  sh.getRange(2, 1, last - 1, 7).sort({ column: 1, ascending: true });
}

function onEdit(e) {
  try {
    if (!e || !e.range) return;
    const sh = e.range.getSheet();
    if (sh.getName() === 'DISKON' && e.range.getRow() > 1 && e.range.getColumn() <= 7) {
      sortDiskonSheet_(sh);
    }
  } catch (err) {
    Logger.log('onEdit DISKON error: ' + err);
  }
}