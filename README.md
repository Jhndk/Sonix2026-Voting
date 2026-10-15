# Sonix2026-Voting

Dashboard Voting Stand Pameran Sonix 2026.

## Status integrasi GitHub Pages + Apps Script

- `index.html` mempertahankan tampilan Vote yang sudah ada dan meneruskan panggilan backend melalui iframe.
- `Bridge.html` adalah kode sumber halaman bridge yang harus dibuat sebagai file HTML bernama **Bridge** di project Google Apps Script.
- Dashboard, Admin, Google Sheets, validasi kode, penguncian percobaan, dan validasi lokasi tetap menggunakan Apps Script.

## Wajib dilakukan di Google Apps Script sebelum GitHub Pages digunakan

**Jangan menghapus atau mengganti kode lama.** Lakukan penambahan berikut pada project Apps Script yang saat ini aktif.

### 1. Tambahkan cabang bridge ke fungsi `doGet(e)`

Di fungsi `doGet(e)` yang sudah ada, setelah nilai `page` ditentukan dan sebelum `let fileName = 'Vote'`, tambahkan:

```javascript
if (page === 'bridge') {
  return HtmlService
    .createHtmlOutputFromFile('Bridge')
    .setTitle('SONIX Voting Bridge')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
```

Pastikan cabang ini berada sebelum kode yang memilih file Vote/Dashboard/Admin.

### 2. Buat file HTML bernama `Bridge`

Di editor Apps Script: klik **+ → HTML**, beri nama `Bridge` (tanpa ekstensi), lalu salin isi file [Bridge.html](./Bridge.html) dari repositori ini ke file tersebut. Nama file Apps Script harus `Bridge`, bukan `Bridge.html`.

### 3. Simpan dan perbarui deployment

Klik **Deploy → Manage deployments → Edit** pada deployment Web App yang sama, pilih **New version**, lalu Deploy. Pertahankan URL Web App yang sudah digunakan aplikasi.

### 4. Aktifkan GitHub Pages

Pada repositori GitHub, buka **Settings → Pages**:
- Source: **Deploy from a branch**
- Branch: **main**
- Folder: **/ (root)**
- Save

URL halaman akan berbentuk `https://jhndk.github.io/Sonix2026-Voting/`.

## Pengujian sebelum dipakai pengunjung

1. Buka URL GitHub Pages dan pastikan halaman Vote tampil.
2. Periksa bahwa daftar stand berhasil dimuat setelah kode voting valid.
3. Uji kode salah dan pastikan aturan blokir 3 percobaan/10 menit serta 7 percobaan/24 jam tetap dijalankan server.
4. Uji izin lokasi dan vote uji yang sah; pastikan satu kode hanya dapat dipakai sekali.
5. Pastikan Dashboard dan Admin Apps Script tetap dapat dibuka dari URL lama.

**Penting:** perubahan pada repositori saja belum mengaktifkan bridge di Apps Script. Langkah 1–3 wajib dilakukan di editor Apps Script sebelum halaman GitHub Pages dapat berkomunikasi dengan backend.
