# PRODUCT REQUIREMENTS DOCUMENT (PRD)
**Sistem Pelaporan dan Rekapitulasi Pembayaran Retribusi SPPG**  
**Dinas Lingkungan Hidup (DLH)**  
*Stack: Google Apps Script (GAS), Google Sheets, Google Drive, & HTML Service (Tailwind UI)*

---

## 1. Ringkasan Produk (Product Overview)
Sebuah aplikasi web modern berbasis **Google Apps Script (GAS)** dan **Google Sheets** yang berfungsi untuk mendigitalisasi pelaporan setoran retribusi sampah oleh SPPG (*Satuan Penugasan Pengelolaan Gotong Royong* di tingkat kecamatan).

**Prinsip Desain & Alur Kerja:**
1. **Konfirmasi Pembayaran (Akses Publik Tanpa Login):** Pihak SPPG dapat langsung mengakses form konfirmasi penyetoran secara terbuka tanpa perlu membuat akun atau login. Cukup memilih Kecamatan $\rightarrow$ Nama SPPG, mengisi data penyetor, nominal, dan mengunggah slip bukti transfer bank.
2. **Dashboard Publik Transparan:** Menampilkan visualisasi capaian setoran per SPPG (Apr - Des), daftar pembayaran terakhir (live), serta matriks tabular status pembayaran dengan popover detail (Tgl Setor, No. STS, No. STBP).
3. **Portal Khusus Bendahara (Dengan Login, 4 Sub-Menu & Pengaturan Sistem):** Staf Bendahara DLH mengakses portal internal dengan login khusus untuk:
   - Manajemen Pembayaran & Dokumen STS/STBP terpisah.
   - Penentuan Periode Wajib Bayar per SPPG dengan fitur *select all* per bulan.
   - Manajemen Master SPPG (Daftar & Tambah SPPG).
   - Pengaturan Konfigurasi Aplikasi dinamis (Logo, Nama App, `<title>`, Tagline, Tahun Anggaran, Footer, dan Parameter Teknis).

---

## 2. Arsitektur & Teknologi (Tech Stack)

| Komponen | Teknologi Pilihan | Peran dalam Sistem |
| :--- | :--- | :--- |
| **Backend & Logic** | Google Apps Script (GAS) | Menangani Web App routing (`doGet`), server-side actions (`google.script.run`), upload file ke Drive, dan operasi Google Sheets. |
| **Database** | Google Sheets | Spreadsheet sebagai database relasional sederhana (mudah diakses, diexport, dan diaudit langsung oleh dinas). |
| **File Storage** | Google Drive API (via GAS `DriveApp`) | Menyimpan file slip bukti transfer (`.png`, `.jpg`, `.pdf`) ke folder Drive khusus secara terpusat. |
| **Frontend Web App** | HTML Service + Tailwind CSS + Alpine.js | Antarmuka pengguna modern (Clean Light Theme), mobile-friendly dengan *Bottom Navigation Bar*. |
| **Otentikasi** | Single-Role Auth (Sheet `users_bendahara`) | Otentikasi berbasis username/email & password untuk akses Bendahara (dengan demo bypass). |

---

## 3. Fitur Utama & Modul Sistem

### A. Dashboard Utama (Public View)
1. **Hero Welcome Banner:** Menampilkan judul aplikasi dinamis, tahun anggaran, dan tombol aksi cepat.
2. **Section 1: Rasio 2:1**
   - **Kiri (2 Bagian) - Grafik Monitoring SPPG:** Grafik Batang Horizontal Gantt/Timeline monitoring rentang bulan terbayar per baris SPPG (Apr s.d. Des) format ultra-compact dengan latar baris selang-seling (zebra striping) agar nyaman dilihat, tanpa teks kecamatan pada baris grafik, dan dilengkapi *internal vertical scrollbar* yang selaras presisi dengan tinggi kontainer kanan.
   - **Kanan (1 Bagian) - Penyetor Terbaru:** Menampilkan 5–6 data setoran terkini dengan indikator dokumen (**STBP** / **STS** jika sudah terbit; dikosongkan jika belum terbit) dan tombol aksi konfirmasi langsung.
3. **Section 2 - Matriks Pembayaran Retribusi SPPG (Compact Layout)**
   - **Tampilan Compact:** Tinggi baris rapat (`py-1.5`), header `py-2`, dan icon proporsional (`w-3.5 h-3.5`) untuk efisiensi ruang pandang.
   - **Baris & Kolom:** Nama SPPG & Kecamatan x Periode 9 Bulan (`Apr` s.d. `Des`).
   - **Latar Sel Utuh (2 Status):**
     - 🟩 **Sudah Bayar:** Background hijau soft (`bg-emerald-100`) + icon centang hijau.
     - 🟥 **Belum Bayar:** Background merah soft (`bg-rose-100`) + icon silang merah.
   - **Popover Detail Setoran Interaktif:** Mengklik sel matriks menampilkan rincian:
     - **Tgl Setor**
     - **No. STS**
     - **No. STBP**
     - **ID Transaksi**
     - **Status Verifikasi:** Ditetapkan **"Terverifikasi"** jika No. STS, No. STBP, dan ID Transaksi sudah lengkap ada; berstatus **"Belum Terverifikasi"** jika ada dokumen yang belum terbit.

### B. Konfirmasi Pembayaran (Form Publik SPPG)
- **Tanpa Nilai Default:** Form dimulai dalam keadaan kosong murni.
- **Identitas Wilayah:** Dropdown Kecamatan $\rightarrow$ Nama SPPG (tanpa kelurahan).
- **Pilihan Periode Bulan Dinamis:** Dropdown bulan otomatis menyesuaikan dengan SPPG yang dipilih dan **hanya menampilkan bulan yang belum terbayar** (bulan yang sudah lunas otomatis disembunyikan).
- **Masking Otomatis:**
  - Masking Nomor Telepon/WhatsApp: Pola `08xx-xxxx-xxxx`.
  - Masking Nominal Setoran: Pola `Rp X.XXX.XXX`.
- **Upload Slip Bank:** Unggah gambar / PDF langsung ke Google Drive DLH.
- **Modal Sukses:** Tanda terima ringkas tanpa nomor registrasi.

### C. Portal Bendahara (Internal DLH)
1. **Submenu 1: Pembayaran & Dokumen (STS / STBP / Tranx)**
   - Kolom No. Transaksi, Tanggal, SPPG, Penyetor, Nominal, Bukti Slip, No. Transaksi Bank (Tranx), Nomor STBP, dan Nomor STS.
   - **Aksi Input Terpisah (3 Tombol):** Tombol modal khusus **Tranx** (No. Transaksi Bank), **STBP**, dan **STS**.
   - **Sistem Paginasi & Baris Dinamis:** Menampilkan 10 baris per halaman secara default, dilengkapi dropdown pilihan ukuran baris (`10`, `20`, `25`, `50`, `100`, `Semua`) dan navigasi halaman lengkap.
2. **Submenu 2: Periode Wajib**
   - Matriks checkbox SPPG x Bulan (`Apr` s.d. `Des`).
   - Cekboks **Centang Semua** di baris bawah setiap kolom bulan untuk aksi massal.
3. **Submenu 3: Master SPPG (Terpadu)**
   - Menggabungkan daftar 30 SPPG, filter pencarian & kecamatan, formulir inline pendaftaran SPPG baru, serta dialog edit data SPPG.
4. **Submenu 4: Pengaturan Sistem (Halaman Penuh / Bukan Modal)**
   - **Kelompok 1: Identitas & Upload Logo Instansi:** Upload berkas logo instansi (PNG/JPG/SVG) dengan *live thumbnail preview*, preset warna/gradien latar logo, Nama Aplikasi & `<title>`, Tagline, dan Teks Footer.
   - **Kelompok 2: Finansial & Penomoran Dokumen Kasda:** Tahun Anggaran Aktif, Format Prefix STBP, Format Prefix STS Kasda.
   - **Kelompok 3: Integrasi Cloud & Layanan SPPG:** Google Drive Folder ID (penyimpanan slip setoran), Batas Maks. Upload Slip (MB), dan WhatsApp Helpdesk SPPG.

---

## 4. Struktur Database Google Sheets

```
SPPG_RETRIBUSI_DB (Spreadsheet)
├── 1. master_sppg          (Data Master SPPG & Kecamatan)
├── 2. pembayaran_retribusi  (Transaksi Setoran, No STBP, No STS)
├── 3. periode_wajib        (Matriks Flag Kewajiban Setor SPPG x Bulan)
├── 4. users_bendahara       (Kredensial Login Bendahara)
└── 5. app_config           (Penyimpanan Parameter Pengaturan Aplikasi)
```

---

## 5. Daftar 30 SPPG Resmi (Kabupaten Pangkep)

| No | Kecamatan | Nama SPPG | Kode SPPG |
| :---: | :--- | :--- | :--- |
| 1 | Balocci | Kassi | `SPPG-001` |
| 2 | Balocci | Kassi 2 | `SPPG-002` |
| 3 | Bungoro | Samalewa 1 | `SPPG-003` |
| 4 | Bungoro | Samalewa 2 | `SPPG-004` |
| 5 | Bungoro | Samalewa 3 | `SPPG-005` |
| 6 | Bungoro | Samalewa 4 | `SPPG-006` |
| 7 | Bungoro | Samalewa 5 | `SPPG-007` |
| 8 | Labakkang | Labakkang | `SPPG-008` |
| 9 | Labakkang | Labakkang 2 | `SPPG-009` |
| 10 | Labakkang | Labakkang 3 | `SPPG-010` |
| 11 | Labakkang | Manakku | `SPPG-011` |
| 12 | Labakkang | Mangallekana | `SPPG-012` |
| 13 | Labakkang | Batara | `SPPG-013` |
| 14 | Mandalle | Manggalung | `SPPG-014` |
| 15 | Mandalle | Tamarupa | `SPPG-015` |
| 16 | Marang | Talaka | `SPPG-016` |
| 17 | Marang | Talaka 2 | `SPPG-017` |
| 18 | Minasa Tene | Bonto Langkasa | `SPPG-018` |
| 19 | Minasa Tene | Kabba | `SPPG-019` |
| 20 | Minasa Tene | Biraeng | `SPPG-020` |
| 21 | Pangkajene | Bonto Perak 1 | `SPPG-021` |
| 22 | Pangkajene | Bonto Perak 2 | `SPPG-022` |
| 23 | Pangkajene | Mappasaile | `SPPG-023` |
| 24 | Pangkajene | Mappasaile 2 | `SPPG-024` |
| 25 | Pangkajene | Padoang Doangan | `SPPG-025` |
| 26 | Pangkajene | Tumampua | `SPPG-026` |
| 27 | Segeri | Bone | `SPPG-027` |
| 28 | Segeri | Segeri | `SPPG-028` |
| 29 | Segeri | Segeri 2 | `SPPG-029` |
| 30 | Tondong Tallasa | Bantimurung | `SPPG-030` |