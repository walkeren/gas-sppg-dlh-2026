# SPPG DLH Retribusi Design System

## 1. Overview
Design system untuk **Portal Retribusi SPPG Dinas Lingkungan Hidup (DLH)** yang dirancang dengan tema **Clean Modern Light**, profesional, ramah diakses oleh perangkat mobile, serta mengutamakan kemudahan navigasi publik dan privasi portal admin bendahara.

---

## 2. Color Palette (Light Theme)

- **Brand Primary (DLH Blue)**:
  - `brand-600` (`#2563EB`): Tombol aksi utama, tautan aktif, aksen brand.
  - `brand-700` (`#1D4ED8`): Hover state untuk elemen primary.
  - `brand-800` / `900` (`#1E40AF` / `#1E3A8A`): Hero gradient & header accents.
  - `brand-50` (`#EFF6FF`): Background badge dan highlight elemen.

- **Secondary Accent (Emerald Leaf)**:
  - `leaf-600` (`#059669`): Simbol retribusi sampah, status terbit STS, icon keberhasilan.
  - `leaf-50` (`#ECFDF5`): Background chip status sukses.

- **Neutral Surfaces & Borders**:
  - `surface-base` (`#F8FAFC`): Background halaman utama (Slate-50).
  - `surface-card` (`#FFFFFF`): Kartu form, tabel, dan modal.
  - `surface-border` (`#E2E8F0`): Border kartu dan pembatas tabel.
  - `surface-border-hover` (`#CBD5E1`): Border input saat hover.
  - `surface-sunken` (`#F1F5F9`): Header tabel dan area drag-and-drop file.

- **Status Colors**:
  - **Success / STS Terbit**: Background `#ECFDF5`, Text `#047857`, Border `#A7F3D0`
  - **Warning / Menunggu Verifikasi**: Background `#FFFBEB`, Text `#B45309`, Border `#FDE68A`
  - **Error / Validasi Batal**: Background `#FEF2F2`, Text `#B91C1C`, Border `#FECACA`

---

## 3. Tipografi
- **Headline Font**: `Sora` (Bobot: 600 SemiBold, 700 Bold) — untuk judul hero, nama SPPG, dan header modal.
- **Body Font**: `DM Sans` (Bobot: 400 Regular, 500 Medium, 600 SemiBold) — untuk label form, teks tabel, dan instruksi.
- **Monospace Font**: `Fira Code` — untuk nomor registrasi (`TRX-...`), kode STS, tanggal, dan format nominal Rupiah.

---

## 4. Struktur Navigasi & Tata Letak

### A. Desktop & Navigasi Portal
- **Header**: Logo DLH dinamis (mendukung *custom upload logo* atau fallback icon dengan preset gradien), nama portal & tahun anggaran, pill navigation (Dashboard & Konfirmasi Pembayaran).
- **Discreet Admin Access**: Akses login petugas ke Portal Bendahara **hanya berada di footer** berupa icon gembok minimalis tanpa tulisan.
- **Portal Bendahara Submenu Layout (Dikelompokkan Berdasarkan Kemiripan Fungsi)**:
  1. `Pembayaran & Dokumen`: Tabel transaksi setoran, filter/pencarian, aksi input Tranx, STBP, STS, preview slip bukti, dan paginasi dinamis.
  2. `Periode Wajib`: Matriks keteraturan kewajiban setor per SPPG $\times$ bulan dengan cekboks centang semua per kolom.
  3. `Master SPPG`: Pengelolaan terpadu daftar 30 SPPG, filter wilayah kecamatan, form inline registrasi SPPG baru, dan modal pengeditan.
  4. `Pengaturan Sistem (Halaman Penuh)`: Tata letak kartu fungsional untuk *Identitas & Upload Logo*, *Finansial & Dokumen Kasda*, serta *Integrasi Cloud & Helpdesk*.

### B. Mobile
- **Header**: Menampilkan logo dan nama aplikasi secara ringkas.
- **Bottom Navigation Bar**: Fixed bar di bagian bawah layar seperti aplikasi native modern (Dashboard & Konfirmasi Pembayaran).

---

## 5. Sistem Notifikasi, Alert & Modal Interaktif

### A. Dynamic Toast Notifications (Floating Stack)
- Terletak mengambang di sudut kanan atas layar (`z-[9999]`) dengan animasi geser dan puding (*slide & fade*).
- Dilengkapi **countdown progress bar** otomatis (4 detik) dan tombol *close* manual.
- **Varian**:
  - `Success` (Hijau Emerald): Konfirmasi submit berhasil, input STS/STBP tersimpan, master SPPG bertambah.
  - `Error / Danger` (Merah Rose): Validasi gagal, ukuran berkas melebihi batas MB.
  - `Warning` (Kuning Amber): Peringatan data belum lengkap, lampiran dibatalkan.
  - `Info` (Biru DLH): Pemberitahuan ekspor data, reset pengaturan.

### B. Dialog Konfirmasi Tindakan (Action Modal)
- Menggantikan `window.confirm()` dan `window.alert()` browser native dengan modal kustom elegan yang memiliki ikon kontekstual, pesan jelas, serta tombol aksi berkode warna (*Amber* untuk warning, *Rose* untuk destructive action).
- Digunakan saat: Konfirmasi reset konfigurasi, konfirmasi logout bendahara, dan penerbitan STS kas daerah.

### C. Contextual Inline Alert Banners
- Banner informasi di halaman formulir konfirmasi pembayaran publik (*Petunjuk pengisian dan verifikasi transfer*).
- Banner status di tabel bendahara (*Informasi sinkronisasi STS ke matriks dashboard secara real-time*).
- Banner pedoman di tabel periode wajib (*Penjelasan dampak centang kewajiban bayar per bulan*).