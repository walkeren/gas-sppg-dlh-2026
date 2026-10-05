# 🚀 Panduan Deployment Frontend Full Vercel (Backend GAS REST API)

Folder `vercelapps/` ini berisi **Frontend Single Page Application (SPA)** murni yang siap di-deploy langsung ke Vercel tanpa menggunakan `<iframe>`.

Dengan arsitektur ini:
- ✅ **100% Bebas Error Multi-Login Google & Google Drive**
- ✅ **Bebas masalah pemblokiran third-party cookie di Chrome, Safari, dan Edge**
- ✅ **Loading super cepat dari Vercel CDN Edge Network**
- ✅ **Google Sheets & Google Drive tetap menjadi Database & Penyimpanan Bukti Transfer**

---

## 📋 Langkah 1: Dapatkan URL Web App Google Apps Script
1. Buka spreadsheet database Anda $\rightarrow$ **Ekstensi** $\rightarrow$ **Apps Script**.
2. Pastikan file `code.gs` di Apps Script Anda sudah diperbarui dengan kode terbaru dari repositori ini.
3. Klik tombol **Deploy** di kanan atas $\rightarrow$ **Manage deployments** (atau **New deployment**).
4. Pastikan pengaturan:
   - **Type:** `Web app`
   - **Execute as:** `Me (email-anda@gmail.com)`
   - **Who has access:** `Anyone`
5. Klik **Deploy**, lalu salin **Web app URL** yang berakhiran `/exec`.  
   *(Contoh: `https://script.google.com/macros/s/AKfycbyxxxxxxx/exec`)*.

---

## 📋 Langkah 2: Masukkan URL API ke `vercelapps/index.html`
1. Buka file [`vercelapps/index.html`](index.html).
2. Cari baris `const GAS_API_URL` (sekitar baris 2150):
   ```javascript
   const GAS_API_URL = "https://script.google.com/macros/s/GANTI_DENGAN_DEPLOYMENT_ID_ANDA/exec";
   ```
3. Ganti `"https://script.google.com/macros/s/GANTI_DENGAN_DEPLOYMENT_ID_ANDA/exec"` dengan URL Web App Anda dari **Langkah 1**.
4. Simpan file `index.html`.

---

## 📋 Langkah 3: Deploy ke Vercel

Pilih salah satu cara berikut yang paling mudah bagi Anda:

### Opsi A: Menggunakan Vercel CLI (Sangat Cepat via Terminal)
1. Buka terminal/PowerShell di folder project:
   ```powershell
   cd v:\GithubRepo\gas-sppg-dlh-2026
   ```
2. Jalankan perintah:
   ```bash
   npx vercel
   ```
3. Ikuti petunjuk di terminal:
   - *Set up and deploy?* Ketik **`y`**
   - *Which scope?* Pilih akun Vercel Anda
   - *Link to existing project?* Ketik **`n`**
   - *What's your project's name?* Ketik nama domain yang diinginkan, misal: `gas-sppg-dlh`
   - *In which directory is your code located?* Tekan **Enter** (`./`)
4. Untuk deploy ke Production URL utama:
   ```bash
   npx vercel --prod
   ```

---

### Opsi B: Deploy via Dashboard Vercel (Drag & Drop / GitHub)
1. Login ke [vercel.com](https://vercel.com).
2. Klik tombol **Add New...** $\rightarrow$ **Project**.
3. Hubungkan repositori GitHub Anda.
4. Pada bagian **Root Directory**, klik **Edit** dan pilih folder **`vercelapps`**.
5. Klik **Deploy**.
6. Selesai! Web app Anda akan aktif di URL seperti `https://gas-sppg-dlh.vercel.app` (atau nama project Anda) dengan 100% performa optimal.

---

## 📁 Struktur File di Folder `vercelapps/`
- `index.html` : Frontend Single Page Application (UI Tailwind, Alpine.js, Lucide Icons, modal, tabel interaktif, dan komunikasi asynchronous `fetch()` API).
- `vercel.json` : Konfigurasi routing static SPA, clean URLs, dan header keamanan Vercel.
- `README.md` : Panduan lengkap instalasi dan deployment.
