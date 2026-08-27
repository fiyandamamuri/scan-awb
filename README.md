# PDF AWB Extractor

Web statis sederhana untuk:

1. Upload PDF hasil **Trace & Tracking CORESYS**.
2. OCR isi PDF langsung di browser.
3. Mengambil kode pada kolom **No. AWB** dari pola URL `/detail_awb/<kode>`.
4. Secara default menghapus huruf `Q` di awal, contoh:
   - `QCGK8163187700011` → `CGK8163187700011`
5. Menghapus duplikat.
6. Export hasil menjadi `.xlsx` dengan satu kolom **No. AWB**.

## Teknologi

- PDF.js — render PDF di browser.
- Tesseract.js — OCR client-side.
- SheetJS — membuat file Excel.

Tidak membutuhkan backend/database.

## Menjalankan lokal

Karena browser membatasi beberapa file ketika dibuka langsung dengan `file://`, paling aman jalankan web server sederhana.

Contoh dengan Python:

```bash
python -m http.server 8000
```

Lalu buka:

```text
http://localhost:8000
```

## Publish ke GitHub Pages

1. Buat repository baru di GitHub, misalnya `awb-pdf-extractor`.
2. Upload seluruh file dalam folder ini ke repository.
3. Buka **Settings → Pages**.
4. Pada **Build and deployment**, pilih **Deploy from a branch**.
5. Pilih branch `main` dan folder `/ (root)`.
6. Klik **Save**.
7. Tunggu GitHub Pages selesai deploy.

Website akan tersedia pada alamat seperti:

```text
https://USERNAME.github.io/awb-pdf-extractor/
```

## Catatan akurasi

Aplikasi ini disesuaikan untuk format PDF CORESYS yang pada kolom No. AWB menampilkan AWB sekaligus hyperlink `.../detail_awb/<AWB>`.

OCR tetap bergantung pada kualitas PDF. Untuk hasil terbaik:

- gunakan PDF export/print yang jelas;
- hindari scan miring atau blur;
- proses sekitar 20 halaman atau kurang sekali jalan pada komputer standar.

Semua proses berlangsung di browser pengguna. PDF tidak dikirim ke server aplikasi ini.
