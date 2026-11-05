# 🧾 PPN Masukan: Barang & Jasa — Automation & Reconciliation Pipeline

![Python](https://img.shields.io/badge/Python-3.8%2B-3776AB?style=for-the-badge&logo=python&logoColor=white)
![Architecture](https://img.shields.io/badge/Pipeline-7--Stage%20ETL%20%26%20Reconciliation-2EA44F?style=for-the-badge)
![Domain](https://img.shields.io/badge/Domain-PPN%20Masukan%20(Accounting%20vs%20Coretax)-F59E0B?style=for-the-badge)
![Status](https://img.shields.io/badge/Status-Active-success?style=for-the-badge)

Sistem otomatisasi berbasis Python untuk pemrosesan, rekonsiliasi, dan pelaporan **Pajak Pertambahan Nilai (PPN) Masukan** atas transaksi **Barang Kena Pajak (BKP)** dan **Jasa Kena Pajak (JKP)**. Sistem ini menghubungkan serta mencocokkan data pencatatan akuntansi internal (**Accounting / Accurate / ERP**) dengan data Faktur Pajak Masukan DJP (**Coretax**), melakukan penandaan status rekonsiliasi, hingga memperkaya deskripsi transaksi menggunakan data **Journal Voucher (JV)**.

---

## 📑 Daftar Isi

- [Fitur Utama](#-fitur-utama)
- [Arsitektur & Alur Kerja](#-arsitektur--alur-kerja)
- [Teknologi yang Digunakan](#-teknologi-yang-digunakan)
- [Struktur Direktori](#-struktur-direktori)
- [Prasyarat Sistem](#-prasyarat-sistem)
- [Panduan Instalasi & Konfigurasi](#-panduan-instalasi--konfigurasi)
- [Cara Penggunaan](#-cara-penggunaan)
- [Referensi Berkas Pendukung](#-referensi-berkas-pendukung)
- [Penanganan Masalah (Troubleshooting)](#-penanganan-masalah-troubleshooting)
- [Kontak & Kontribusi](#-kontak--kontribusi)

---

## ✨ Fitur Utama

Sistem dirancang secara modular melalui **7 tahap pemrosesan** di dalam folder `Dapur/`:

1. **Pembersihan & Pemisahan Data Accounting (`1_AccCleaner&PshBrgJs.py`)**
   - Menghapus baris kosong, merapikan spasi, dan menormalisasi format angka DPP/PPN.
   - Memisahkan data transaksi akuntansi menjadi kategori **Barang** dan **Jasa** berdasarkan referensi `hbrg.txt`.

2. **Pembersihan & Pemisahan Data Coretax (`2_CtxPshBrgJs.py`)**
   - Mengekstrak dan menstandarkan data Faktur Pajak Masukan dari Coretax.
   - Mengelompokkan transaksi Coretax ke dalam kategori **Barang** dan **Jasa** agar sepadan dengan data akuntansi.

3. **Rekonsiliasi & Analisis Barang (`3_AnalyticsBrgAccCtx.py`)**
   - Pencocokan silang (*cross-matching*) Nomor Faktur, NPWP, dan Masa Pajak untuk kategori **Barang**.
   - Menghitung dan mendeteksi selisih nilai DPP maupun PPN antara Accounting dan Coretax.

4. **Rekonsiliasi & Analisis Jasa (`4_AnalyticsJsAccCtx.py`)**
   - Analisis komparatif khusus transaksi **Jasa**, mengakomodasi karakteristik perbedaan tanggal pencatatan akrual (*accrual*) vs tanggal penerbitan faktur.

5. **Konsolidasi Hasil Analisis (`5_MergeHasil.py`)**
   - Menggabungkan (*merge*) hasil analisis kategori Barang dan Jasa ke dalam satu struktur dataset rekonsiliasi terpadu.

6. **Penandaan Status Rekonsiliasi (`6_HelperMarkedData.py`)**
   - Memberikan *flagging* atau anotasi status (misalnya: *Matched*, *Selisih DPP/PPN*, *Unmatched in Accounting*, atau *Unmatched in Coretax*) untuk mempermudah audit tindak lanjut.

7. **Pengayaan Deskripsi Journal Voucher (`7_LookUpDescJV.py`)**
   - Melakukan pencarian silang (*lookup*) ke basis data **Journal Voucher (JV)** menggunakan skema `hjv.txt`.
   - Menambahkan rincian keterangan transaksi jurnal sehingga laporan akhir memiliki konteks operasional yang lengkap.

---

## 🔄 Arsitektur & Alur Kerja

```text
+-------------------------------------------------------------------------+
|                        SUMBER DATA & KONFIGURASI                        |
|  [Data Accounting/Accurate]   [Data Faktur Coretax]   [Data Jurnal JV]  |
|  [Dapur/config.conf]          [Dapur/hbrg.txt]        [Dapur/hjv.txt]   |
+-----------------------------------+-------------------------------------+
                                    |
                                    v
+-------------------------------------------------------------------------+
|                 ENTRY POINT: Masukan_Barang_Jasa.py                     |
+-----------------------------------+-------------------------------------+
                                    |
         +--------------------------+--------------------------+
         |                                                     |
         v                                                     v
+---------------------------------+   +-----------------------------------+
| 1_AccCleaner&PshBrgJs.py        |   | 2_CtxPshBrgJs.py                  |
| - Cleansing data Accounting     |   | - Standarisasi data Coretax       |
| - Pisah kategori Barang & Jasa  |   | - Pisah kategori Barang & Jasa    |
+--------+---------------+--------+   +---------+---------------+---------+
         |               |                      |               |
   (Acc Barang)     (Acc Jasa)            (Ctx Barang)     (Ctx Jasa)
         |               |                      |               |
         |               +-----------+----------+               |
         |                           |                          |
         +-------------+             |            +-------------+
                       |             |            |
                       v             |            v
+---------------------------------+  |  +---------------------------------+
| 3_AnalyticsBrgAccCtx.py         |  |  | 4_AnalyticsJsAccCtx.py          |
| - Rekonsiliasi BARANG           |  +-+| - Rekonsiliasi JASA             |
|   (Accounting vs Coretax)       |     |   (Accounting vs Coretax)       |
+----------------+----------------+     +----------------+----------------+
                 |                                       |
                 +-------------------+-------------------+
                                     |
                                     v
                 +---------------------------------------+
                 | 5_MergeHasil.py                       |
                 | - Penggabungan hasil analisis         |
                 |   kategori Barang & Jasa              |
                 +-------------------+-------------------+
                                     |
                                     v
                 +---------------------------------------+
                 | 6_HelperMarkedData.py                 |
                 | - Penandaan status (Marking/Flagging) |
                 |   kecocokan, selisih, & anomali data  |
                 +-------------------+-------------------+
                                     |
                                     v
                 +---------------------------------------+
                 | 7_LookUpDescJV.py                     |
                 | - Lookup & injeksi deskripsi transaksi|
                 |   dari sumber Journal Voucher (JV)    |
                 +-------------------+-------------------+
                                     |
                                     v
+-------------------------------------------------------------------------+
|                             OUTPUT AKHIR                                |
|     Laporan Rekonsiliasi PPN Masukan (Barang & Jasa) Teranotasi + JV    |
+-------------------------------------------------------------------------+
```

---

## 🛠️ Teknologi yang Digunakan

| Komponen | Pustaka / Alat | Peran Utama |
| :--- | :--- | :--- |
| **Runtime** | Python 3.8+ | Lingkungan eksekusi skrip utama dan modul pemrosesan. |
| **Manipulasi Data** | `pandas`, `numpy` | Pembersihan data tabular, pemisahan kategori, analisis selisih, dan penggabungan dataset. |
| **Manajemen Spreadsheet** | `openpyxl`, `xlrd` | Pembacaan dan penulisan berkas Excel (`.xlsx` / `.xls`). |
| **Konfigurasi** | `configparser` (`config.conf`) | Pengaturan jalur berkas sumber data dan parameter dinamis tanpa *hardcode*. |
| **Format Skema** | Teks Datar (`.txt`) | Kamus acuan header/kolom Barang (`hbrg.txt`) dan Journal Voucher (`hjv.txt`). |

---

## 📂 Struktur Direktori

```text
PPN-Masukan-Barang-Jasa-test/
├── Masukan_Barang_Jasa.py            # Orkestrator / Entry point utama seluruh pipeline
└── Dapur/                            # Paket pemrosesan data terpadu
    ├── __init__.py                   # Inisialisasi modul (Maintainer: me@ihfandica.com)
    ├── 1_AccCleaner&PshBrgJs.py      # Tahap 1: Cleansing data Accounting & pemisahan Barang/Jasa
    ├── 2_CtxPshBrgJs.py              # Tahap 2: Pemrosesan data Coretax & pemisahan Barang/Jasa
    ├── 3_AnalyticsBrgAccCtx.py       # Tahap 3: Analisis & rekonsiliasi Barang (Acc vs Ctx)
    ├── 4_AnalyticsJsAccCtx.py        # Tahap 4: Analisis & rekonsiliasi Jasa (Acc vs Ctx)
    ├── 5_MergeHasil.py               # Tahap 5: Penggabungan hasil analisis Barang & Jasa
    ├── 6_HelperMarkedData.py         # Tahap 6: Flagging & penandaan status rekonsiliasi
    ├── 7_LookUpDescJV.py             # Tahap 7: Lookup deskripsi transaksi dari Journal Voucher
    ├── config.conf                   # Parameter konfigurasi jalur berkas & folder kerja
    ├── hbrg.txt                      # Header/kata kunci pemetaan kategori Barang
    └── hjv.txt                       # Header acuan pembacaan Journal Voucher (JV)
```

---

## 📋 Prasyarat Sistem

1. **Python 3.8 ke atas** sudah terpasang di sistem (`python --version` atau `python3 --version`).
2. Tiga berkas sumber data telah disiapkan:
   - Data transaksi PPN Masukan dari sistem akuntansi (**Accounting**).
   - Data Faktur Pajak Masukan hasil unduhan dari portal **Coretax**.
   - Data rincian jurnal (**Journal Voucher / JV**).
3. Struktur kolom berkas masukan sesuai dengan acuan pada `Dapur/hbrg.txt` dan `Dapur/hjv.txt`.

---

## 🚀 Panduan Instalasi & Konfigurasi

### 1. Kloning Repositori
```bash
git clone https://github.com/ihfandica/PPN-Masukan-Barang-Jasa-test.git
cd PPN-Masukan-Barang-Jasa-test
```

### 2. Buat Lingkungan Virtual (Virtual Environment)
- **Windows (Command Prompt / PowerShell):**
  ```cmd
  python -m venv .venv
  .venv\Scripts\activate
  ```
- **macOS / Linux:**
  ```bash
  python3 -m venv .venv
  source .venv/bin/activate
  ```

### 3. Pasang Dependensi
```bash
pip install pandas openpyxl numpy
```

### 4. Sesuaikan Konfigurasi
Buka berkas `Dapur/config.conf` dan sesuaikan parameter berikut dengan jalur berkas di komputer Anda:
- Jalur file data Accounting, Coretax, dan JV.
- Direktori output untuk menyimpan laporan akhir.

---

## 💻 Cara Penggunaan

### A. Eksekusi Seluruh Alur (Direkomendasikan)
Jalankan file orkestrator dari direktori utama proyek:
```bash
python Masukan_Barang_Jasa.py
```
Perintah ini akan mengeksekusi Tahap 1 hingga Tahap 7 secara otomatis dan menghasilkan laporan konsolidasi akhir.

### B. Eksekusi Modular (Tahap per Tahap)
Jika ingin melakukan pengujian atau *debugging* pada tahap tertentu:

```bash
# Tahap 1 (Gunakan tanda kutip karena nama berkas memuat karakter '&')
python "Dapur/1_AccCleaner&PshBrgJs.py"

# Tahap 2
python Dapur/2_CtxPshBrgJs.py

# Tahap 3
python Dapur/3_AnalyticsBrgAccCtx.py

# Tahap 4
python Dapur/4_AnalyticsJsAccCtx.py

# Tahap 5
python Dapur/5_MergeHasil.py

# Tahap 6
python Dapur/6_HelperMarkedData.py

# Tahap 7
python Dapur/7_LookUpDescJV.py
```

---

## ⚙️ Referensi Berkas Pendukung

| Berkas | Jenis | Deskripsi |
| :--- | :--- | :--- |
| `Dapur/config.conf` | File Konfigurasi | Mengatur *path* berkas input/output dan parameter operasional modul. |
| `Dapur/hbrg.txt` | File Referensi | Berisi daftar nama kolom atau kata kunci untuk mendeteksi transaksi **Barang**. |
| `Dapur/hjv.txt` | File Referensi | Berisi struktur nama kolom acuan saat membaca berkas **Journal Voucher (JV)**. |

---

## 🔧 Penanganan Masalah (Troubleshooting)

| Masalah | Kemungkinan Penyebab | Solusi |
| :--- | :--- | :--- |
| Perintah `python Dapur/1_AccCleaner&PshBrgJs.py` gagal di terminal | Karakter `&` dianggap sebagai pemisah baris perintah oleh terminal shell. | Bungkus nama berkas dengan tanda kutip: `python "Dapur/1_AccCleaner&PshBrgJs.py"` atau jalankan melalui `python Masukan_Barang_Jasa.py`. |
| `FileNotFoundError` saat skrip berjalan | Jalur berkas di `config.conf` salah atau terminal tidak berada di direktori akar proyek. | Pastikan terminal berada di folder `PPN-Masukan-Barang-Jasa-test` dan periksa kembali *path* di dalam `config.conf`. |
| `KeyError` pada kolom data | Nama kolom pada berkas Excel ekspor tidak sesuai dengan daftar di `hbrg.txt` atau `hjv.txt`. | Periksa kembali header berkas Excel masukan dan samakan dengan acuan di `hbrg.txt` atau `hjv.txt`. |
| `PermissionError` saat menyimpan hasil | Berkas Excel keluaran sedang terbuka di aplikasi lain (Microsoft Excel). | Tutup aplikasi Excel yang sedang membuka berkas tersebut, lalu jalankan kembali skrip. |

---


## 📜 Lisensi

Proyek ini dikembangkan untuk keperluan internal perusahaan. Silakan sesuaikan dengan kebutuhan organisasi Anda.

---

*Dikembangkan oleh [ACC-TAX-REIGHTEEN](https://github.com/ACC-TAX-REIGHTEEN)*
