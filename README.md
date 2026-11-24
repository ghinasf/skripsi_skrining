# Sistem Skrining Tingkat Risiko Kehamilan — Puskesmas Bihbul

Aplikasi web skrining tingkat risiko kehamilan berbasis **Certainty Factor** yang dibangun menggunakan **Django** untuk mendukung pencatatan dan rekapitulasi data pasien di Poli KIA Puskesmas Bihbul.

## Fitur Utama

* **Autentikasi & Hak Akses**: Login pengguna dilengkapi verifikasi **Google reCAPTCHA v2** serta pembagian akses antara pengguna biasa dan admin/superuser.
* **Skrining Risiko Kehamilan**: Menghitung persentase risiko, status risiko, dan keterangan medis berdasarkan profil pasien dan gejala klinis menggunakan metode **Certainty Factor**.
* **Pencarian & Pengelompokan Data**: Pencarian pasien berdasarkan nama, filter tanggal kunjungan, serta pengelompokan daftar pasien per tanggal kunjungan pada dashboard.
* **Statistik & Visualisasi**: Ringkasan jumlah pasien serta grafik perbandingan status risiko menggunakan **Chart.js** pada panel admin.
* **Ekspor Laporan**: Unduh laporan rekapitulasi skrining dalam format **Excel (`.xlsx`)** dan **PDF** berdasarkan rentang tanggal tertentu.

## Teknologi yang Digunakan

* **Backend**: Python, Django
* **Frontend**: HTML, CSS, JavaScript, Chart.js
* **Database**: SQLite / database proyek sesuai konfigurasi lokal
* **Ekspor Dokumen**: `openpyxl`, `reportlab`
* **Keamanan & HTTP**: Google reCAPTCHA v2, `requests`

## Cara Menjalankan Proyek Secara Lokal

1. **Clone repository**
   ```bash
   git clone https://github.com/ghinasf/skripsi_skrining.git
cd skripsi_skrining