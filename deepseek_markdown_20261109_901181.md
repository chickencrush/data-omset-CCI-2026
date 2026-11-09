# 🚀 Sales Command Center — 37 Outlet

Dashboard web app modern futuristik (glassmorphism) untuk monitoring penjualan 37 outlet. Bisa langsung di-host gratis di GitHub Pages.

## ✨ Fitur

- **Upload Excel** — Membaca file `Template_Penjualan_37_Outlet_Advanced.xlsx` langsung di browser
- **Input Data** — Form tambah baris penjualan (auto-save ke localStorage)
- **Export Excel** — Download kembali sebagai `.xlsx` (formula Total Omset tetap aktif)
- **3 Tab Minimalis**: Overview · Outlets · Data
- **Grafik interaktif** (Chart.js): Kategori & Perbandingan Bulan
- **Auto-agregasi**: KPI, ranking, growth MoM, monitoring kelengkapan, Pusat vs Mitra
- **Penyimpanan lokal** — Data tidak hilang saat refresh

## 📥 Cara Deploy ke GitHub Pages

1. **Buat repository baru** di GitHub (bisa private/public)
2. **Upload 4 file** ini ke root repository:
   - `index.html`
   - `style.css`
   - `app.js`
   - `README.md`
3. Buka **Settings → Pages**
4. Di **Source**, pilih `Deploy from a branch`
5. Branch: `main` · Folder: `/ (root)` → **Save**
6. Tunggu ~1 menit, akses di: `https://USERNAME.github.io/NAMA-REPO/`

## 🎯 Alur Penggunaan

1. **Upload Excel** → klik tombol "Upload Excel" dan pilih file template Anda
2. Dashboard otomatis menampilkan data
3. **Input baris baru** → klik "Input Data" (form modal)
4. **Export** → klik "Export" untuk download `.xlsx` terbaru (bisa dibuka di Excel/Sheets)
5. Ganti file asli dengan hasil export jika perlu

## 📝 Catatan Penting

- Karena GitHub Pages bersifat statis, **tidak bisa menulis langsung ke file Excel**. Semua perubahan disimpan di browser (localStorage). Gunakan tombol **Export** untuk menyimpan ke Excel.
- Saat upload Excel baru, data di app akan **diganti** dengan data dari file tersebut.
- Rumus `TOTAL OMSET` (kolom I) tetap ditulis sebagai formula Excel saat export.
- Sheet selain `MASTER OUTLET`, `INPUT PENJUALAN`, `DASHBOARD` dari file asli ikut dipertahankan saat export.

## 🎨 Tema

Dark futuristic glassmorphism dengan aksen neon cyan & purple, animated gradient orbs, dan grid pattern halus.

## 🛠️ Tech

- Vanilla HTML/CSS/JS (zero build)
- [SheetJS](https://sheetjs.com/) untuk baca/tulis Excel
- [Chart.js](https://www.chartjs.org/) untuk grafik
- Google Fonts: Inter + JetBrains Mono

## 📄 Lisensi

Bebas digunakan dan dimodifikasi.