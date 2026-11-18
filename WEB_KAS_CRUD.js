/* ======================================================
   WEB_KAS_CRUD.js   (dipakai web SL & web SANFELO, isi file sama)
   ------------------------------------------------------
   Ubah & hapus baris PETTY CASH / BUKU BANK dari web (tambah = fungsi
   simpan yang sudah ada di WEB_KAS_LAPORAN.js, lihat = webDaftarKas_).

   PENGAMAN DATA:
   - Satu per satu (LockService, kunci yang sama dengan simpan kas).
   - Sebelum mengubah, isi baris di sheet dicek SAMA PERSIS dengan yang
     dilihat admin di web (tanggal, keterangan, kategori, nominal). Kalau
     sudah berubah (mis. baru diedit orang lain), dibatalkan -> muat ulang.
   - Hanya kolom kas (A..G + kolom catatan I) yang disentuh; baris di atas
     baris yang diubah tidak ditulis ulang sama sekali.
   - Isi lama + isi baru dicatat di LOG_WEB bersama nama admin.

   DUA BENTUK SHEET (dikenali otomatis dari kolom H / SALDO):
   - SALDO berupa ANGKA (PETTY CASH SANFELO, BUKU BANK SL): sama dengan
     popup Petty Cash -> tanggal sama digabung (merge), urut tanggal,
     SALDO dihitung ulang mulai dari grup tanggal yang berubah ke bawah.
   - SALDO berupa RUMUS (BUKU BANK SANFELO, =H(atas)+C+D-E-F-G): hanya
     isi A..G yang diubah / digeser, rumus SALDO tetap.
   ====================================================== */

const WEB_KAS_LEBAR_ = { 'PETTY CASH': 9, 'BUKU BANK': 8 }; // A..I (I = catatan/nama) / A..H
const WEB_KAS_FORMAT_TGL_ = 'dd/mm/yyyy';
const WEB_KAS_RUMUS_SALDO_ = /^=\s*(H\d+\s*\+\s*)?C\d+\s*\+\s*D\d+\s*-\s*E\d+\s*-\s*F\d+\s*-\s*G\d+\s*$/i;

function webKasNama_(a) {
  const nama = String(a.sheet || '').trim().toUpperCase();
  if (!WEB_KAS_KATEGORI_[nama]) throw new Error('Sheet kas "' + nama + '" tidak tersedia di web ini.');
  return nama;
}

function webKasAngka_(v) { return typeof v === 'number' && isFinite(v) ? v : 0; }

function webKasSamaTgl_(a, b) {
  if (!(a instanceof Date) || !(b instanceof Date)) return !(a instanceof Date) && !(b instanceof Date);
  return webTgl_(a) === webTgl_(b);
}

// kategori & nominal satu baris (kolom C..G, isi terakhir yang tidak kosong -- sama dengan webDaftarKas_)
function webKasKategoriBaris_(kat, r) {
  let kunci = '';
  let nominal = 0;
  for (let c = 2; c <= 6; c++) {
    if (r[c] !== '' && r[c] !== null) { kunci = kat[c - 2][0]; nominal = Number(r[c]) || 0; }
  }
  return { kunci: kunci, nominal: nominal };
}

function webKasBaca_(sh, nama) {
  const mulai = WEB_KAS_MULAI_[nama];
  const lebar = WEB_KAS_LEBAR_[nama] || 8;
  const lastRow = sh.getLastRow();
  if (lastRow < mulai) return { mulai: mulai, lebar: lebar, vals: [], tgl: [], rumus: false };
  const semua = sh.getRange(mulai, 1, lastRow - mulai + 1, lebar).getValues();
  // baris data terakhir = ada isi selain kolom H (SALDO) -- sama dengan findLastDataRow_ popup Petty Cash
  let n = 0;
  for (let i = semua.length - 1; i >= 0 && !n; i--) {
    if (semua[i].some(function (v, c) { return c !== 7 && v !== '' && v !== null; })) n = i + 1;
  }
  const vals = semua.slice(0, n);
  const tgl = [];
  let akhir = null;
  vals.forEach(function (r) { if (r[0] instanceof Date) akhir = r[0]; tgl.push(akhir); });
  const rumus = n ? sh.getRange(mulai, 8, n, 1).getFormulas().some(function (f) { return String(f[0] || '').trim() !== ''; }) : false;
  return { mulai: mulai, lebar: lebar, vals: vals, tgl: tgl, rumus: rumus };
}

// baris di sheet harus sama dengan yang dilihat admin di web
function webKasCekAsal_(data, kat, no, asal) {
  const idx = Number(no) - data.mulai;
  if (!(idx >= 0 && idx < data.vals.length) || Math.floor(idx) !== idx) throw new Error('Baris ' + no + ' tidak ditemukan. Muat ulang daftar kas.');
  const r = data.vals[idx];
  const k = webKasKategoriBaris_(kat, r);
  asal = asal || {};
  const sama = String(r[1] || '').trim() === String(asal.ket || '').trim() &&
    k.kunci === String(asal.kunci || '') &&
    Math.round(k.nominal) === Math.round(Number(asal.nominal) || 0) &&
    (data.tgl[idx] ? webTgl_(data.tgl[idx]) : '') === String(asal.tgl || '');
  if (!sama) throw new Error('Isi baris ' + no + ' sudah berubah sejak daftar dibuka (mungkin baru diedit di spreadsheet / admin lain). Tidak ada yang diubah. Klik Tampilkan lalu ulangi.');
  return idx;
}

function webKasTeks_(tgl, ket, kat, kunci, nominal) {
  const k = kat.filter(function (x) { return x[0] === kunci; })[0];
  return (tgl ? webTgl_(tgl) : '-') + ' | ' + ket + ' | ' + (k ? k[1] : kunci) + ' ' + Math.round(nominal);
}

function webKasIsianBaru_(kat, a) {
  const b = a.baru || {};
  const tgl = webTanggalDariInput_(b.tanggal);
  if (!(tgl instanceof Date) || isNaN(tgl.getTime())) throw new Error('Tanggal wajib diisi.');
  const k = kat.map(function (x) { return x[0]; }).indexOf(String(b.kunci || ''));
  if (k === -1) throw new Error('Kategori tidak dikenali.');
  const nominal = Number(b.nominal);
  if (!(nominal > 0)) throw new Error('Nominal tidak valid.');
  const ket = String(b.keterangan || '').trim();
  if (!ket) throw new Error('Keterangan wajib diisi.');
  return { tgl: tgl, kolom: k, kunci: kat[k][0], nominal: nominal, ket: ket };
}

function webUbahKas_(a, admin) { return webKasUbahAtauHapus_(a, admin, false); }
function webHapusKas_(a, admin) { return webKasUbahAtauHapus_(a, admin, true); }

function webKasUbahAtauHapus_(a, admin, hapus) {
  const nama = webKasNama_(a);
  const kat = WEB_KAS_KATEGORI_[nama];
  const baru = hapus ? null : webKasIsianBaru_(kat, a);

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new Error('Database sedang sibuk, coba lagi sebentar.');
  let hasil;
  let teksLama = '';
  try {
    const sh = webSS_().getSheetByName(nama);
    if (!sh) throw new Error('Sheet "' + nama + '" tidak ada di database.');
    const data = webKasBaca_(sh, nama);
    const idx = webKasCekAsal_(data, kat, a.no, a.asal);
    const lama = webKasKategoriBaris_(kat, data.vals[idx]);
    teksLama = webKasTeks_(data.tgl[idx], String(data.vals[idx][1] || ''), kat, lama.kunci, lama.nominal);
    hasil = data.rumus ? webKasTulisRumus_(sh, data, idx, baru) : webKasTulisStatis_(sh, data, idx, baru);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  webCatatLog_(admin, (hapus ? 'HAPUS ' : 'UBAH ') + nama, 'baris ' + a.no + ' | lama: ' + teksLama +
    (baru ? ' | baru: ' + webKasTeks_(baru.tgl, baru.ket, kat, baru.kunci, baru.nominal) : '') +
    (a.alasan ? ' | alasan: ' + String(a.alasan) : ''));
  return hasil;
}

/* ---------- SALDO berupa rumus (BUKU BANK SANFELO) ---------- */
function webKasTulisRumus_(sh, data, idx, baru) {
  const mulai = data.mulai;
  const n = data.vals.length;
  if (baru) {
    const row = [baru.tgl, baru.ket, '', '', '', '', ''];
    row[2 + baru.kolom] = baru.nominal;
    sh.getRange(mulai + idx, 1, 1, 7).setValues([row]);
    sh.getRange(mulai + idx, 1).setNumberFormat(WEB_KAS_FORMAT_TGL_);
  } else {
    // geser baris di bawahnya naik satu (A..G dan kolom catatan), rumus SALDO di H tidak disentuh
    if (idx < n - 1) {
      const bawah = data.vals.slice(idx + 1);
      sh.getRange(mulai + idx, 1, bawah.length, 7).setValues(bawah.map(function (r) { return r.slice(0, 7); }));
      if (data.lebar > 8) sh.getRange(mulai + idx, 9, bawah.length, data.lebar - 8).setValues(bawah.map(function (r) { return r.slice(8); }));
    }
    const akhir = mulai + n - 1;
    sh.getRange(akhir, 1, 1, 7).clearContent();
    if (data.lebar > 8) sh.getRange(akhir, 9, 1, data.lebar - 8).clearContent();
    const f = sh.getRange(akhir, 8).getFormula();
    if (WEB_KAS_RUMUS_SALDO_.test(f)) sh.getRange(akhir, 8).clearContent(); // saldo baris kosong ikut dihapus
  }
  SpreadsheetApp.flush();
  const nBaru = baru ? n : n - 1;
  const saldo = nBaru ? sh.getRange(mulai + nBaru - 1, 8).getValue() : 0;
  let peringatan = '';
  if (baru && idx > 0 && data.tgl[idx - 1] instanceof Date && baru.tgl.getTime() < data.tgl[idx - 1].getTime()) {
    peringatan = 'Tanggal baru lebih lama dari baris di atasnya; posisi baris tidak dipindah (sama seperti input Buku Bank).';
  }
  return { success: true, saldo: saldo, peringatan: peringatan };
}

/* ---------- SALDO berupa angka (PETTY CASH SANFELO, BUKU BANK SL) ---------- */
function webKasTulisStatis_(sh, data, idx, baru) {
  const mulai = data.mulai;
  const lebar = data.lebar;
  const lama = data.vals.map(function (r, i) { return { r: r, tgl: data.tgl[i], i: i }; });
  const waktu = function (x) { return x.tgl instanceof Date ? x.tgl.getTime() : 0; };
  let urut = true;
  for (let i = 1; i < lama.length; i++) if (waktu(lama[i]) < waktu(lama[i - 1])) { urut = false; break; }

  let daftar = lama.slice();
  let posBaru = idx;
  if (baru) {
    const r = lama[idx].r.slice();
    r[1] = baru.ket;
    for (let c = 2; c <= 6; c++) r[c] = '';
    r[2 + baru.kolom] = baru.nominal;
    const item = { r: r, tgl: baru.tgl, i: idx };
    daftar[idx] = item;
    // data yang sudah urut tanggal tetap urut (sama dengan popup Petty Cash: urut stabil per tanggal)
    if (urut && !webKasSamaTgl_(lama[idx].tgl, baru.tgl)) {
      daftar.sort(function (x, y) { return (waktu(x) - waktu(y)) || (x.i - y.i); });
      posBaru = daftar.indexOf(item);
    }
  } else {
    daftar.splice(idx, 1);
  }

  // mulai tulis dari awal grup tanggal yang terpengaruh (baris di atasnya tidak disentuh)
  const awalGrup = function (list, i) {
    if (i >= list.length) i = list.length - 1;
    while (i > 0 && webKasSamaTgl_(list[i - 1].tgl, list[i].tgl)) i--;
    return Math.max(0, i);
  };
  const m = Math.min(idx, posBaru);
  let k = awalGrup(lama, m);
  if (daftar.length) k = Math.min(k, awalGrup(daftar, m));

  // saldo sebelum baris k
  let saldo = 0;
  for (let i = k - 1; i >= 0; i--) {
    if (typeof lama[i].r[7] === 'number') {
      saldo = lama[i].r[7];
      for (let j = i + 1; j < k; j++) saldo += webKasAngka_(lama[j].r[2]) + webKasAngka_(lama[j].r[3]) - webKasAngka_(lama[j].r[4]) - webKasAngka_(lama[j].r[5]) - webKasAngka_(lama[j].r[6]);
      break;
    }
    if (i === 0) for (let j = 0; j < k; j++) saldo += webKasAngka_(lama[j].r[2]) + webKasAngka_(lama[j].r[3]) - webKasAngka_(lama[j].r[4]) - webKasAngka_(lama[j].r[5]) - webKasAngka_(lama[j].r[6]);
  }

  const tulis = [];
  for (let i = k; i < daftar.length; i++) {
    const r = daftar[i].r.slice(0, lebar);
    while (r.length < lebar) r.push('');
    r[0] = (i === k || !webKasSamaTgl_(daftar[i - 1].tgl, daftar[i].tgl)) ? (daftar[i].tgl || '') : '';
    saldo += webKasAngka_(r[2]) + webKasAngka_(r[3]) - webKasAngka_(r[4]) - webKasAngka_(r[5]) - webKasAngka_(r[6]);
    r[7] = saldo;
    tulis.push(r);
  }
  const jumlahLama = lama.length - k; // baris sheet yang ditulis ulang
  while (tulis.length < jumlahLama) { const kosong = []; for (let c = 0; c < lebar; c++) kosong.push(''); tulis.push(kosong); }

  if (jumlahLama > 0) {
    sh.getRange(mulai + k, 1, jumlahLama, 1).breakApart();
    sh.getRange(mulai + k, 1, jumlahLama, lebar).setValues(tulis);
    const isi = daftar.length - k;
    if (isi > 0) {
      sh.getRange(mulai + k, 1, isi, 1).setNumberFormat(WEB_KAS_FORMAT_TGL_);
      sh.getRange(mulai + k, 3, isi, 6).setNumberFormat('#,##0');
      // gabung lagi sel tanggal yang sama (seperti popup Petty Cash)
      let g = k;
      for (let i = k + 1; i <= daftar.length; i++) {
        if (i === daftar.length || !webKasSamaTgl_(daftar[i].tgl, daftar[g].tgl)) {
          if (i - g > 1) sh.getRange(mulai + g, 1, i - g, 1).merge().setVerticalAlignment('middle');
          g = i;
        }
      }
    }
  }
  return {
    success: true,
    saldo: saldo, // saldo baris terakhir
    peringatan: baru && posBaru !== idx ? 'Tanggal diubah, baris dipindah supaya tetap urut tanggal.' : ''
  };
}
