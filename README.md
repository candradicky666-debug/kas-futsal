# ⚽ Buku Kas Futsal (Web App & PWA)

Aplikasi web pencatatan rekapan kas futsal yang responsif, dapat diinstal langsung di layar utama **iPhone** (PWA tanpa App Store), dan terintegrasi otomatis dengan **Google Spreadsheet**.

---

## 🌟 Fitur Utama

1. **4 Indikator Lengkap**:
   - **Pemasukan**: Uang iuran pemain, donasi, hadiah turnamen.
   - **Pengeluaran**: Sewa lapangan, konsumsi/air mineral, beli bola/rompi, uang pendaftaran.
   - **Sisa Kas**: Saldo berjalan realtime setelah setiap transaksi.
   - **Total Keseluruhan**: Akumulasi total pemasukan, pengeluaran, dan saldo kas akhir tim.
2. **Semua Kolom Bisa Diisi Langsung**:
   - Setiap sel tabel dapat diketik dan diedit secara instan (No, Tanggal, Keterangan, Kategori, Pemasukan, Pengeluaran, dan Sisa Kas).
   - Mode Sisa Kas bisa diatur **Otomatis** (dihitung oleh rumus) atau **Manual** (bisa diisi sesuka hati).
3. **Bisa Diinstal di iPhone (PWA Standalone)**:
   - Tampilan bersih layaknya aplikasi native iOS.
   - Tetap bisa dibuka dan mencatat transaksi meskipun sedang offline (tanpa sinyal).
4. **Tersambung ke Google Spreadsheet**:
   - Terhubung dengan Google Apps Script Web App Deployment ID:
     `AKfycbz6mx3yARb6jFM1omx9sLjIC4rI8EyIZzPzbMAjNbaGHKPnIauOjJ_okx8BwcFEdd1LJA`
   - Sinkronisasi dua arah (web app ⇄ Google Sheets).
5. **Ekspor & Cadangan**:
   - Unduh laporan kas dalam format **Excel (.xlsx)** atau **CSV**.
   - Backup data dalam format **JSON**.

---

## 🚀 Cara Upload ke GitHub & Mengaktifkan GitHub Pages

Ikuti 3 langkah mudah ini untuk mempublikasikan website ke akun GitHub Anda:

### 1. Buat Repositori Baru di GitHub
1. Buka [github.com/new](https://github.com/new).
2. Beri nama repositori, misalnya: `kas-futsal`.
3. Biarkan opsi **Public** terpilih.
4. Jangan centang "Add a README file" (karena sudah ada di lokal).
5. Klik **Create repository**.

### 2. Push Kode dari Komputer Anda
Buka terminal PowerShell di folder ini (`D:\data pemasukan onok`), lalu jalankan perintah berikut (ganti `<username-anda>` dengan username GitHub Anda):

```bash
git remote add origin https://github.com/<username-anda>/kas-futsal.git
git branch -M main
git push -u origin main
```

### 3. Aktifkan GitHub Pages
1. Pada halaman repositori GitHub Anda, klik menu **Settings** (Pengaturan).
2. Pada bilah menu kiri, klik **Pages**.
3. Pada bagian **Build and deployment > Branch**:
   - Ubah dari *None* menjadi **`main`**.
   - Folder biarkan **`/(root)`**.
4. Klik tombol **Save**.
5. Tunggu sekitar 1–2 menit, website Anda akan aktif di:
   `https://<username-anda>.github.io/kas-futsal/`

---

## 📲 Cara Memasang di iPhone (Add to Home Screen)

1. Buka link GitHub Pages Anda di peramban **Safari** pada iPhone.
2. Ketuk tombol **Share / Bagikan** (ikon kotak dengan panah ke atas di bagian bawah layar).
3. Gulir ke bawah dan ketuk **"Tambah ke Layar Utama"** (*Add to Home Screen*).
4. Beri nama (misal: "Kas Futsal") lalu ketuk **Tambah** (*Add*) di pojok kanan atas.
5. Ikon **⚽ Kas Futsal** akan langsung muncul di layar utama iPhone Anda dan dapat dibuka seperti aplikasi biasa tanpa batas browser.

---

## ⚙️ Pengaturan Izin Google Apps Script (PENTING)

Agar Google Spreadsheet Anda dapat menerima data dari web app tanpa error otentikasi:
1. Buka Google Spreadsheet Anda.
2. Klik menu **Ekstensi** > **Apps Script**.
3. Pastikan kode dari file [google_apps_script.js](google_apps_script.js) sudah disalin ke dalam editor Apps Script.
4. Klik tombol biru **Terapkan (Deploy)** > **Kelola Penerapan (Manage Deployments)** > edit versi atau buat **Penerapan Baru (New Deployment)**.
5. Atur pengaturan berikut:
   - **Jalankan sebagai (Execute as)**: `Saya (Me)`
   - **Siapa yang memiliki akses (Who has access)**: **`Siapa saja (Anyone)`** *(Wajib agar iPhone bisa kirim data tanpa login akun Google)*
6. Klik **Terapkan (Deploy)** dan setujui perizinan (*Authorize Access*).

---

## 📁 Struktur Berkas

- `index.html`: Antarmuka web app dengan kartu total dan tabel interaktif.
- `style.css`: Desain responsif bertema turf hijau futsal yang ramah layar sentuh iPhone.
- `app.js`: Logika perhitungan saldo otomatis, edit sel tabel, LocalStorage, dan koneksi Google Apps Script.
- `manifest.json`: Konfigurasi Progressive Web App (PWA).
- `sw.js`: Service worker untuk dukungan mode offline di iPhone.
- `icons/`: Ikon aplikasi resolusi tinggi untuk layar iOS & Android.
- `google_apps_script.js`: Kode backend Apps Script untuk ditempelkan di Google Sheets.
- `Rekap_Uang_Kas_Futsal.xlsx`: File spreadsheet Excel awal.
- `Rekap_Uang_Kas_Futsal.csv`: File rekapan CSV awal.
