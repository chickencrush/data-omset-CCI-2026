# PETUNJUK DASHBOARD PENJUALAN 37 OUTLET

Dashboard web modern (futuristik + glass) yang membaca data langsung dari file
Excel Anda. Tidak perlu install apa pun — cukup upload ke GitHub, dashboard akan
online dan memuat ulang data secara otomatis.

---

## 1. Isi Paket

| File | Fungsi |
|---|---|
| `index.html` | Halaman dashboard (jangan diubah) |
| `style.css` | Tampilan dashboard (jangan diubah) |
| `app.js` | Logika dashboard (jangan diubah) |
| `data/penjualan.xlsx` | **File data Anda** — file ini yang diperbarui berkala |
| `PETUNJUK.md` | Dokumen yang sedang Anda baca |

---

## 2. Cara Upload ke GitHub Pages (tanpa coding)

1. Login ke [github.com](https://github.com), klik **+** di kanan atas → **New repository**.
2. Beri nama repo, contoh: `dashboard-penjualan`. Pilih **Public**, klik **Create repository**.
3. Klik **"uploading an existing file"** (atau tombol **Add file → Upload files**).
4. Drag **semua isi folder ini** ke area upload: `index.html`, `style.css`, `app.js`,
   `PETUNJUK.md`, dan **folder `data` berisi `penjualan.xlsx`**.
   Pastikan `index.html` berada di root repo (bukan di dalam subfolder).
5. Klik **Commit changes**.
6. Buka **Settings → Pages** (menu kiri). Pada **Build and deployment**:
   - Source: **Deploy from a branch**
   - Branch: **main** dan folder **/ (root)** → klik **Save**.
7. Tunggu 1–2 menit, muat ulang halaman Settings → Pages. Akan muncul alamat seperti:
   `https://USERNAME.github.io/dashboard-penjualan/`
   Itulah alamat dashboard Anda — bisa dibuka dari HP maupun komputer.

---

## 3. Cara Update Data (pilih salah satu)

### Opsi A — Mode Google Sheets (paling praktis, near real-time)

1. Upload `data/penjualan.xlsx` ke Google Drive, klik kanan → **Buka dengan → Google Spreadsheet**.
2. Jangan mengubah nama sheet (`INPUT PENJUALAN`, `MASTER OUTLET`, dst).
3. Klik tombol **Bagikan** → akses **"Anyone with the link" (Siapa saja yang memiliki link)**
   sebagai **Viewer/Pembaca** → Salin link.
4. Buka dashboard → klik ikon **gerigi (Pengaturan)** → pilih **Google Sheets** →
   tempel link → **Simpan & Muat Ulang**.
5. Selesai. Setiap kali Anda mengisi angka di Google Sheets, dashboard akan menarik
   data terbaru otomatis setiap 60 detik — tanpa refresh, tanpa upload ulang.

### Opsi B — Ganti file Excel di GitHub

1. Isi data di Excel seperti biasa (sheet `INPUT PENJUALAN`, mulai baris 5).
2. Simpan sebagai `.xlsx` dengan **nama yang sama**: `penjualan.xlsx`.
3. Di GitHub repo Anda: **Add file → Upload files** → masuk ke folder `data` →
   drag `penjualan.xlsx` (file lama otomatis tertimpa) → **Commit changes**.
4. Sekitar 1–2 menit setelah commit, dashboard memuat data baru secara otomatis
   (pembaruan terjadi tiap 60 detik, tanpa perlu refresh halaman).

### Opsi C — Drag & drop manual (untuk cek cepat)

Tarik file `.xlsx` ke halaman dashboard kapan saja. Data hanya tampil di
perangkat Anda saat itu dan tidak tersimpan untuk pengunjung lain.

---

## 4. Aturan Input Data (agar dashboard akurat)

1. Input **hanya** di sheet `INPUT PENJUALAN`, mulai **baris 5**.
2. Satu baris = satu outlet untuk satu bulan.
3. Kolom `Periode Bulan` diisi **tanggal 1** pada bulan laporan (contoh: `01-Jan-2026`).
4. Isi nominal rupiah **tanpa simbol Rp** dan tanpa titik (contoh: `125000000`).
5. Jangan menimpa kolom `TOTAL OMSET` — kolom itu berisi rumus otomatis.
6. Lengkapi **Wilayah** dan **Jenis Outlet (Pusat/Mitra)** di sheet `MASTER OUTLET`
   agar filter dan perbandingan Pusat vs Mitra berfungsi.

---

## 5. Apa yang Ditampilkan Dashboard

- **Ringkasan**: total omset tahunan, omset bulan berjalan, growth bulanan,
  jumlah laporan masuk, grafik omset 12 bulan, kontribusi 6 kategori
  (Dine in, Take away, GoFood, GrabFood, ShopeeFood, ESB Order), dan Pusat vs Mitra.
- **Outlet**: ranking 37 outlet + kontribusi + kategori terbesar, pencarian,
  grafik top 8, dan kelengkapan laporan per bulan (outlet yang belum lapor).
- **Bulanan**: perbandingan antar bulan + growth, tren tiap kategori, tabel rincian.

Tahun analisis dapat dipilih pada dropdown **Tahun** di kanan atas.

---

## 6. Troubleshooting

| Masalah | Solusi |
|---|---|
| Dashboard menampilkan "Belum Ada Data" | Pastikan data diisi di sheet `INPUT PENJUALAN` mulai baris 5, dengan tanggal dan nama outlet terisi. Cek juga pilihan tahun. |
| "Gagal memuat data ... HTTP 404" | File `data/penjualan.xlsx` tidak ada / namanya berubah. Upload ulang dengan nama yang sama, atau sesuaikan path di Pengaturan. |
| Google Sheets gagal dimuat | Pastikan sharing **"Anyone with the link"** dan nama sheet masih `INPUT PENJUALAN` / `MASTER OUTLET`. |
| Angka tidak berubah | Tunggu maksimal 60 detik (refresh otomatis), atau klik **Pengaturan → Muat Sekarang**. Tekan Ctrl+F5 bila perlu. |
| Pusat vs Mitra kosong | Isi kolom `Jenis Outlet` di `MASTER OUTLET` dengan tepat `Pusat` atau `Mitra`. |
| Halaman tampil polos tanpa gaya | Pastikan `style.css` dan `app.js` berada di folder yang sama dengan `index.html`. |
