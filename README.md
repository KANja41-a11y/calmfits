# Calm Fits — Deploy Package (dengan Google Sheets + Admin)

## Isi paket
- `index.html` — frontend (GitHub Pages)
- `Code.gs` — backend Apps Script
- `images/` — foto outfits & wallpaper
- `audio/ghost.mp3` — musik
- `manifest.json` + `sw.js` — PWA

## Fitur baru
1. **Database Google Sheets** (Outfits, ShopProducts, Orders) — sheet & header digenerate otomatis dari script
2. **Tombol Admin** di header menu → minta password → panel admin
3. **Password admin `Bismillaah`** hanya disimpan di **Script Properties** (tidak ada di frontend / GitHub)
4. **CRUD** Outfits, Shop Products, Orders di panel admin
5. Saat **Checkout**: user wajib isi **Nama + Nomor WA** → pesanan masuk ke sheet Orders (status: pending / processed / done)
6. Admin bisa ubah status pesanan & klik WA langsung

---

## Langkah deploy Backend (Apps Script)

1. Buat **Google Spreadsheet baru** (kosong)
2. **Extensions → Apps Script**
3. Hapus kode default, **paste seluruh isi `Code.gs`**
4. **Project Settings** (ikon gear) → **Script properties** → Add:
   - Property: `ADMIN_PASSWORD`
   - Value: `Bismillaah`
5. Simpan → pilih fungsi **`setupSheets`** di dropdown → **Run**  
   (izinkan permission Google). Ini akan membuat sheet Outfits, ShopProducts, Orders + header + data awal.
6. **Deploy → New deployment**
   - Type: **Web app**
   - Execute as: **Me**
   - Who has access: **Anyone**
   - Deploy → **salin Web App URL**

---

## Langkah deploy Frontend (GitHub Pages)

1. Buka `index.html`, cari baris:
   ```js
   const API_URL = 'YOUR_WEB_APP_URL_HERE';
   ```
   Ganti dengan Web App URL yang kamu dapat (contoh: `https://script.google.com/macros/s/AKfycb.../exec`)

2. Upload seluruh folder ini ke repo GitHub (atau replace file di repo yang sudah ada)

3. Aktifkan **GitHub Pages** (Settings → Pages → Deploy from branch `main` / `docs`)

4. Buka situs → klik **⚙ Admin** di header → password: **Bismillaah**

---

## Catatan penting
- Password **tidak pernah** disimpan di `index.html` atau repo GitHub.
- Token admin berlaku ±2 jam (disimpan di Script Properties).
- Jika API_URL belum di-set, website tetap jalan dengan data fallback lokal; order akan tersimpan lokal saja.
- Setelah ubah data di admin, refresh halaman publik agar data terbaru ter-load.
- Untuk gambar baru: upload ke folder `images/` di repo, lalu isi path `images/nama-file.jpg` di form admin.

## Status pesanan
- `pending` — baru masuk
- `processed` — sedang dilayani / dijawab
- `done` — selesai

Selamat pakai ☁️
