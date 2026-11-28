/* ======================================================
   WEB_HUTANG_TGL.js   (dipakai web SL & web SANFELO, isi file sama)
   ------------------------------------------------------
   Ubah TANGGAL pembayaran hutang yang salah ketik (termasuk PB yang sudah LUNAS).
   Riwayat bayar hutang tersimpan sebagai teks di baris PB sheet HUTANG, kolom K dst:
     "dd/MM/yyyy - Rp1.000.000 - TRANSFER"
   Yang diganti HANYA bagian tanggal di sel itu. Jumlah, metode, sisa bayar & status tidak berubah.
   Sebelum menulis: PB di baris itu dan isi sel dicek sama persis dengan yang tampil di web.
   Isi lama & baru dicatat di LOG_WEB.
   ====================================================== */

function webBarisHutang_(baris, pb) {
  const sh = webSS_().getSheetByName('HUTANG');
  if (!sh) throw new Error('Sheet HUTANG tidak ada.');
  const n = Number(baris);
  if (!(n >= 2) || n > sh.getLastRow()) throw new Error('Baris HUTANG tidak valid. Muat ulang daftar PB.');
  const lebar = Math.max(sh.getLastColumn(), WEB_HUTANG_.RIWAYAT);
  const r = sh.getRange(n, 1, 1, lebar).getValues()[0];
  if (String(r[WEB_HUTANG_.PB - 1] || '').trim() !== String(pb || '').trim()) {
    throw new Error('Baris ' + n + ' sekarang bukan ' + pb + '. Muat ulang daftar PB.');
  }
  return { sh: sh, r: r };
}

/** aksi baca: daftar pembayaran 1 PB. a = { baris, pb } */
function webRiwayatBayarHutang_(a) {
  const x = webBarisHutang_(a.baris, a.pb);
  const list = [];
  for (let i = WEB_HUTANG_.RIWAYAT - 1; i < x.r.length; i++) {
    const teks = String(x.r[i] || '').trim();
    if (!teks) continue;
    const bag = teks.split(' - ');
    list.push({ kolom: i + 1, teks: teks, tgl: bag[0] || '', jumlah: bag[1] || '', metode: bag.slice(2).join(' - ') });
  }
  return { pb: String(a.pb), supplier: String(x.r[WEB_HUTANG_.SUPPLIER - 1] || ''), list: list };
}

/** aksi tulis: a = { baris, pb, kolom, lama, tglBaru (yyyy-MM-dd), alasan } */
function webUbahTglBayarHutang_(a, admin) {
  const alasan = String(a.alasan || '').trim();
  if (alasan.length < 3) throw new Error('Alasan wajib diisi.');
  const p = String(a.tglBaru || '').split('-');
  if (p.length !== 3 || p[0].length !== 4) throw new Error('Tanggal baru belum diisi.');
  const tglBaru = p[2] + '/' + p[1] + '/' + p[0];
  const kolom = Number(a.kolom);
  if (!(kolom >= WEB_HUTANG_.RIWAYAT)) throw new Error('Kolom riwayat tidak valid.');

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new Error('Database sedang sibuk, coba lagi sebentar.');
  let lama, baru;
  try {
    const x = webBarisHutang_(a.baris, a.pb);
    lama = String(x.r[kolom - 1] || '').trim();
    if (!lama || lama !== String(a.lama || '').trim()) throw new Error('Riwayat bayar ini baru saja berubah. Buka ulang riwayatnya.');
    const bag = lama.split(' - ');
    if (bag.length < 3) throw new Error('Format riwayat tidak dikenal: ' + lama);
    if (bag[0].trim() === tglBaru) return { success: true, pesan: 'Tanggal sama, tidak ada yang diubah.' };
    baru = [tglBaru].concat(bag.slice(1)).join(' - ');
    x.sh.getRange(Number(a.baris), kolom).setValue(baru);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  webCatatLog_(admin, 'UBAH TGL BAYAR HUTANG', String(a.pb) + ' | lama: ' + lama + ' | baru: ' + baru + ' | alasan: ' + alasan);
  return { success: true, pesan: 'Tanggal bayar ' + a.pb + ' diubah jadi ' + tglBaru + '.' };
}
