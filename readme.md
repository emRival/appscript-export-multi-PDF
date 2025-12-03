# Ekspor Multi-PDF Spreadsheet Otomatis dengan Google Apps Script 🗂️

<p align="center">
    <img style="margin-right: 8px;" src="https://img.shields.io/badge/Language-HTML-orange.svg" alt="HTML">
    <img style="margin-right: 8px;" src="https://img.shields.io/badge/Technology-Google%20Apps%20Script-green.svg" alt="Google Apps Script">
    <img style="margin-right: 8px;" src="https://img.shields.io/badge/Topic-Spreadsheet%20Automation-blue.svg" alt="Spreadsheet Automation">
    <img style="margin-right: 8px;" src="https://img.shields.io/badge/Status-Development-yellow" alt="Development">
</p>

Proyek ini menyediakan solusi untuk mengekspor beberapa *sheet* dari Google Spreadsheet ke dalam file PDF secara otomatis menggunakan Google Apps Script.  Dengan proyek ini, Anda dapat menghasilkan laporan atau dokumentasi PDF dengan cepat dan efisien, sehingga menghemat waktu dan mengurangi pekerjaan manual.

## Fitur Utama ✨

*   **Otomatisasi Ekspor PDF:** 🤖 Otomatiskan proses ekspor *sheet* ke PDF langsung dari Google Spreadsheet.
*   **Ekspor Multi-*Sheet*:** 📑 Mampu mengekspor beberapa *sheet* sekaligus menjadi file PDF terpisah.
*   **Penyimpanan Otomatis di Google Drive:** ☁️ File PDF yang dihasilkan secara otomatis disimpan di Google Drive.
*   **Konfigurasi Mudah:** ⚙️ Mudah dikonfigurasi dan disesuaikan dengan kebutuhan spesifik Anda.
*   **Integrasi Apps Script:** 🔗 Menggunakan Google Apps Script untuk integrasi yang mulus dengan Google Spreadsheet dan Drive.

## Tech Stack 🛠️

*   HTML 📃
*   Google Apps Script 📜
*   JavaScript (ES6+) ☕

## Instalasi & Menjalankan 🚀

1.  Clone repositori:
    ```bash
    git clone https://github.com/emRival/appscript-export-multi-PDF
    ```
2.  Masuk ke direktori:
    ```bash
    cd appscript-export-multi-PDF
    ```
3.  Deploy ke Google Apps Script:
    *   Buka Google Spreadsheet.
    *   Buka "Tools" > "Script editor".
    *   Salin isi file `code.js`, `runpdfgenerator.js`, `test.js`, dan `sitebar.html` ke *script editor* Google Apps Script. Pastikan untuk membuat file HTML baru untuk `sitebar.html` di dalam editor script.
    *   Simpan proyek.
4.  Konfigurasi dan Jalankan:
    *   Modifikasi kode di `code.js` sesuai kebutuhan (misalnya, ID *spreadsheet*, nama *sheet*, lokasi penyimpanan).
    *   Jalankan fungsi `run()` di *script editor* untuk menguji. Anda mungkin perlu memberikan izin akses ke Google Drive Anda.

## Cara Berkontribusi 🤝

1.  Fork repositori ini.
2.  Buat *branch* baru dengan nama yang deskriptif: `git checkout -b fitur-baru`
3.  Lakukan perubahan dan *commit*: `git commit -m "Menambahkan fitur baru"`
4.  *Push* ke repositori Anda: `git push origin fitur-baru`
5.  Buat *Pull Request*.

## Lisensi 📄

Tidak ada lisensi yang ditentukan.


---
README.md ini dihasilkan secara otomatis oleh [README.MD Generator](https://github.com/emRival) — dibuat dengan ❤️ oleh [emRival](https://github.com/emRival)