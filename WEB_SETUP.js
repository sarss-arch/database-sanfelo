/* ======================================================
   WEB_SETUP.js   (PROJECT: database_SANFELO)
   ------------------------------------------------------
   PANDUAN PEMASANGAN WEB ADMIN (pemilik, sekali):
   1. Buka spreadsheet database_SANFELO > Ekstensi > Apps Script.
   2. Tambahkan file skrip: WEB_APP.js, WEB_SETUP.js, WEB_PO_MASUK.js,
      WEB_SURAT_JALAN.js, WEB_POSTING_LAIN.js, WEB_KAS_LAPORAN.js,
      WEB_CETAK.js, WEB_ALAT.js, WEB_INV_SALINAN.js dan file HTML: WebIndex, WebStyle,
      WebScript, WebLogo. File lama JANGAN diubah.
   2b. Jalankan webSiapkanFileCetak (sekali; ulangi kalau template nota
      di inventory diubah). Izinkan akses Drive saat diminta.
   3. Pilih fungsi webSiapkanAkses di atas > Jalankan. Lihat Log
      eksekusi: berisi KUNCI link dan PIN admin pertama.
   4. Terapkan > Deployment baru > Aplikasi web:
        Jalankan sebagai: Saya     Yang memiliki akses: Siapa saja
      (deployment BARU khusus web; deployment gateway lama biarkan).
   5. Link untuk admin = URL deployment + ?k=KUNCI
      contoh: https://script.google.com/macros/s/XXXX/exec?k=KUNCI
   6. Tambah/ubah admin: Setelan proyek (ikon gerigi) > Properti skrip >
      tambah properti WEB_PIN_NAMAADMIN (huruf besar) = PIN-nya.
      Hapus properti itu untuk mencabut akses.
   ====================================================== */

function webSiapkanAkses() {
  const props = PropertiesService.getScriptProperties();
  let kunci = props.getProperty(WEB_PROP_KUNCI);
  if (!kunci) {
    kunci = Utilities.getUuid().replace(/-/g, '').slice(0, 20);
    props.setProperty(WEB_PROP_KUNCI, kunci);
  }
  const adaAdmin = props.getKeys().some(function (k) { return k.indexOf(WEB_PREFIX_PIN) === 0; });
  if (!adaAdmin) {
    const pin = String(Math.floor(100000 + Math.random() * 900000));
    props.setProperty(WEB_PREFIX_PIN + 'ADMIN', pin);
    Logger.log('Admin pertama: nama ADMIN, PIN ' + pin);
  }
  Logger.log('KUNCI link: ' + kunci + '\nLink admin = URL deployment web + ?k=' + kunci);
  webDaftarAdmin();
}

function webDaftarAdmin() {
  const props = PropertiesService.getScriptProperties();
  const nama = props.getKeys()
    .filter(function (k) { return k.indexOf(WEB_PREFIX_PIN) === 0; })
    .map(function (k) { return k.slice(WEB_PREFIX_PIN.length); });
  Logger.log('Admin web: ' + (nama.join(', ') || '(belum ada)'));
}

// ganti kunci link (link lama langsung tidak bisa dipakai lagi)
function webGantiKunciAkses() {
  const kunci = Utilities.getUuid().replace(/-/g, '').slice(0, 20);
  PropertiesService.getScriptProperties().setProperty(WEB_PROP_KUNCI, kunci);
  Logger.log('KUNCI link BARU: ' + kunci);
}
