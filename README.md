# 🚀 ARCDVNpdf - The Ultimate Offline PDF & Document Utility

![Version](https://img.shields.io/badge/version-0.2.0-blue.svg) ![Platform](https://img.shields.io/badge/platform-Windows-lightgrey.svg) ![Python](https://img.shields.io/badge/python-3.10%2B-green.svg)

**ARCDVNpdf** adalah *Command Line Interface (CLI)* interaktif super ringan yang dirancang untuk menjadi "Pisau Swiss Army" bagi segala kebutuhan pemrosesan PDF dan Dokumen Anda—langsung dari PC Anda secara **100% Offline!**

---

**ARCDVNpdf hadir sebagai solusi:**
- ✅ **100% Offline & Aman**: Dokumen Anda tidak pernah meninggalkan komputer Anda. Semua diproses langsung di *RAM* komputer lokal. Privasi terjamin mutlak!
- ✅ **Sangat Ringan (Zero-Cache)**: Begitu aplikasi ditutup, sistem operasi langsung mengembalikan 100% memori. Tidak ada proses tersembunyi yang tertinggal berjalan (*No background daemons*), dan tidak ada file sampah (temp) yang menumpuk. Performa PC tetap secepat kilat!
- ✅ **User-Friendly (Terintegrasi GUI Windows)**: Walaupun berbentuk aplikasi terminal (CLI), ARCDVNpdf **terintegrasi penuh dengan Jendela File (*GUI File Picker*)**. Anda tidak perlu mengetik *path* lokasi folder yang panjang; cukup tekan `Enter` dan jendela pemilihan file Windows (seperti saat fitur *Open/Save As* di Word) akan langsung terbuka.

---

## 💡 Fitur & Kemampuan Utama

ARCDVNpdf hadir dengan menu interaktif dan fungsionalitas super cepat:

1. **[1] Merge PDF (Gabung PDF)**  
   Gabungkan puluhan atau ratusan file PDF secara berurutan menjadi satu file instan, tanpa batasan ukuran file maksimum!
   
2. **[2] Split PDF (Pisah PDF)**  
   Pecah satu file PDF bervolume besar menjadi kepingan-kepingan file PDF kecil (per halaman) secara akurat.

3. **[3] Rotate PDF (Putar PDF)**  
   Koreksi orientasi halaman PDF Anda dari hasil *scan* yang terbalik. Putar semua halaman ke kanan atau kiri (90, 180, atau 270 derajat).

4. **[4] Compress PDF (Kompresi)**  
   Optimasi struktur file PDF Anda secara pintar (*structural compression*) untuk memperkecil *file-size* tanpa merusak tulisan aslinya.

5. **[5] Convert Document (Word/Excel/PowerPoint → PDF)**  
   Ubah dokumen `docx`, `xlsx`, `pptx`, `odt`, dan format *office* lainnya menjadi bentuk PDF sejati. *(Catatan: ARCDVNpdf meminjam mesin Microsoft Word atau LibreOffice yang terinstal di PC Anda untuk melakukan sihir ini secara otomatis tanpa harus membuka aplikasinya).*

6. **[6] Images ↔ PDF (Gambar ke PDF & Sebaliknya)**  
   - Konversi dan satukan tumpukan file foto/gambar (`JPG`, `PNG`, dll) menjadi satu dokumen PDF yang rapi (cocok untuk pengumpulan tugas foto).
   - Ekstrak seluruh halaman dari sebuah PDF menjadi kumpulan file gambar gambar (`PNG`) yang terpisah dengan resolusi/kualitas tinggi.

7. **[7] PDF Info (Cek Meta-data)**  
   Bongkar info dari PDF untuk mengetahui ukuran aslinya, total halaman, dan memeriksa apakah PDF tersebut terkunci (memiliki enkripsi/kata sandi).

8. **[8] PDF to Word (PDF → DOCX)**  
   Sulap kembali file PDF Anda menjadi format dokumen teks Microsoft Word (`.docx`) yang bisa diketik dan diedit ulang secara instan.

---

## 💻 Panduan Instalasi (Windows)

1. Pastikan Anda sudah menginstal **Python (minimal 3.10+)** di PC Anda (Saat menginstal Python, pastikan untuk mencentang opsi *"Add Python to PATH"*).
2. *Extract* (unzip) folder dari proyek ini dan taruh di mana saja.
3. **Klik ganda (*Double-click*)** pada file `install.bat` dan tunggu proses pengunduhan dependensi *library* pintar selesai secara otomatis.
4. Setelah instalasi selesai tertulis sukses, tutup jendela CMD tersebut.
5. Buka **Command Prompt (CMD) atau PowerShell** baru.
6. Ketik perintah ajaib ini:
```cmd
   arcdvnpdf
   ```
   *Selamat! Menu interaktif ARCDVNpdf siap melayani Anda.*

---

## 🛠️ Cara Penggunaan

### 1. Interactive Mode (Sangat Disarankan!)
Cukup panggil nama program tanpa embel-embel lain di CMD Anda:
```cmd
arcdvnpdf
```
Aplikasi akan membuka tampilan menu angka, memandu Anda langkah demi langkah. Kapan pun ia meminta file/folder, cukup tekan **Enter** dan Jendela GUI Windows akan muncul agar Anda tinggal mengeklik saja!

### 2. Command-Line Mode (Bagi Para Programmer/Otomatisasi)
Untuk Anda yang suka mengetik *script* otomatis dengan cepat:
```cmd
arcdvnpdf merge fileA.pdf fileB.pdf -o hasil.pdf
arcdvnpdf split laporan.pdf -o folder_pages
arcdvnpdf rotate laporan.pdf --angle 90 -o rotated.pdf
arcdvnpdf compress laporan.pdf
arcdvnpdf convert dokumen_saya.docx
arcdvnpdf images foto1.jpg foto2.png -o album.pdf
arcdvnpdf pdf-to-images majalah.pdf -o folder_gambar
arcdvnpdf pdf-to-word presentasi.pdf -o hasil.docx
arcdvnpdf info rahasia.pdf
arcdvnpdf doctor
```

---

## 🗑️ Cara Uninstall
Ingin menghapusnya? Cukup klik ganda pada `uninstall.bat`. Aplikasi akan tercabut dengan bersih dan cepat, tanpa meninggalkan jejak kotor (*registry/cache*) di PC Anda.

---
*Dibuat untuk kecepatan dan kemudahan Anda. Karena keamanan dan privasi data Anda adalah yang utama.*
