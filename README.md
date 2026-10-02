# Panduan Deploy ke Vercel (Memendekkan URL GAS)

Vercel adalah layanan hosting cloud gratis yang sangat cepat dan otomatis memberikan URL pendek yang rapi (misal: `https://sppg-pangkep.vercel.app` atau domain instansi).

---

## Langkah Deploy ke Vercel:

1. **Buka file [`index.html`](./index.html)** di folder ini.
2. Ganti teks `GANTI_DENGAN_DEPLOYMENT_ID_ANDA` dengan URL deployment Web App Google Apps Script Anda:
   ```html
   <iframe 
     id="app-frame" 
     src="https://script.google.com/macros/s/AKfycbx.../exec"
     ...
   ```
3. Buka situs [**vercel.com**](https://vercel.com) dan login (bisa login menggunakan akun GitHub atau Google).
4. Klik tombol **"Add New..."** $\rightarrow$ **"Project"**.
5. Pilih repositori GitHub ini, lalu pada bagian **Root Directory**, pilih folder `vercel`.
   *(Atau tarik / drag-and-drop langsung folder `vercel` ini ke dashboard Vercel)*.
6. Beri nama project sesuai nama URL yang diinginkan, misal: `sppg-pangkep` atau `retribusi-dlh`.
7. Klik **Deploy**.
8. Dalam hitungan detik, portal Anda langsung aktif di:
   `https://sppg-pangkep.vercel.app`

---

## Keuntungan Menggunakan Vercel:
- **100% Gratis & Otomatis HTTPS (SSL)**.
- **URL Pendek & Bersih**: Pengunjung tidak akan melihat URL panjang Google Apps Script.
- **Mendukung Custom Domain**: Anda bisa menghubungkan domain resmi pemda seperti `retribusi-sppg.pangkepkab.go.id` langsung dari menu *Settings* $\rightarrow$ *Domains* di Vercel.
