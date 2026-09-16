# 🧾 PPN Masukan Barang & Jasa — Rekonsiliasi Accurate vs Coretax DJP

> **Rekonsiliasi otomatis PPN Masukan antara Accurate dan Coretax — terpisah untuk Barang dan Jasa, dengan algoritma pencocokan bertingkat dan laporan selisih siap audit**

Pipeline Python lima langkah yang membaca ekspor buku besar PPN Masukan dari **Accurate** (`Accuratem.xls`) dan data Faktur Pajak Masukan dari **Coretax DJP** (`Coretaxm.xlsx`), memisahkan transaksi Barang dan Jasa, melakukan rekonsiliasi dengan algoritma berbeda untuk masing-masing kategori, lalu menggabungkan hasilnya ke satu laporan Excel multi-sheet **`Hasil_Analisis_Barang_Dan_Jasa.xlsx`** — berisi ringkasan, rincian selisih, dan data lengkap siap untuk keperluan audit pajak. Alat ini sudah mendukung untuk rekonsiliasi dua laporan, yaitu sebelum dan setelah di laporkan ke DJP.

---

## 📋 Daftar Isi

- [Gambaran Umum](#-gambaran-umum)
- [Fitur Utama](#-fitur-utama)
- [Prasyarat](#-prasyarat)
- [Struktur Folder & File](#-struktur-folder--file)
- [Cara Penggunaan](#-cara-penggunaan)
- [Alur Kerja Pipeline](#-alur-kerja-pipeline)
- [Detail Tiap Skrip](#-detail-tiap-skrip)
  - [Skrip 1 — Ekstraksi Accurate](#skrip-1--ekstraksi-accurate)
  - [Skrip 2 — Filter Coretax](#skrip-2--filter-coretax)
  - [Skrip 3 — Rekonsiliasi Barang](#skrip-3--rekonsiliasi-barang)
  - [Skrip 4 — Rekonsiliasi JV/Jasa](#skrip-4--rekonsiliasi-jvjasa)
  - [Skrip 5 — Gabung Hasil Akhir](#skrip-5--gabung-hasil-akhir)
- [Konfigurasi](#-konfigurasi)
  - [`config.conf` — Alias nama pemasok](#configconf--alias-nama-pemasok)
  - [`hbrg.txt` — Daftar pemasok Barang](#hbrgtxt--daftar-pemasok-barang)
  - [`hjv.txt` — Daftar pemasok JV/Jasa](#hjvtxt--daftar-pemasok-jvjasa)
- [Format File Input](#-format-file-input)
- [Output & Struktur Laporan](#-output--struktur-laporan)
- [Logika Pencocokan & Keterangan](#-logika-pencocokan--keterangan)
- [Troubleshooting](#-troubleshooting)
- [Catatan Penting](#-catatan-penting)

---

## 🗂️ Gambaran Umum

Dalam pelaporan PPN Masukan, data dari Accurate (sisi pembukuan internal) dan Coretax DJP (sisi sistem pajak pemerintah) sering kali berbeda karena perbedaan waktu input, pembulatan, atau perbedaan pengelompokan transaksi. Pipeline ini mengotomasi proses rekonsiliasi keduanya secara terpisah:

| Jalur | Sumber Accurate | Sumber Coretax | Algoritma |
|---|---|---|---|
| **Barang** | Seksi `P: PPN (11.00%)` | Filter `hbrg.txt` | Pencocokan via Nomor Faktur Pajak |
| **JV/Jasa** | Seksi `Transaksi Lainnya` | Filter `hjv.txt` | Pencocokan bertingkat (group, single, kombinasi) |

Kedua hasil digabungkan ke satu workbook dengan ringkasan total terpadu.

---

## ✨ Fitur Utama

- **Pemisahan otomatis & kondisional Barang vs Jasa** — Accurate dibaca berdasarkan penanda seksi teks; Coretax difilter berdasarkan nama pemasok dan pilihan kriteria nominal PPN spesifik (`| nominal`) dari dua file teks konfigurasi terpisah.
- **Auto-fit Formatting** — Memformat secara otomatis lebar kolom Excel output temporary agar rapi dan mudah dibaca.
- **Deteksi kolom dinamis** — Mendukung dua format ekspor Coretax (nama kolom lama dan baru) secara otomatis via `find_column()` dengan kandidat dan fallback.
- **Normalisasi nama pemasok** — Uppercase, hapus prefix PT./CV./UD./FA., terapkan alias dari `config.conf` sebelum pencocokan — mengurangi miss-match akibat variasi penulisan.
- **Bersihkan nomor faktur** — Strip semua karakter non-alphanumerik, tangani notasi saintifik (misalnya `1.23e+10`) sebelum pencocokan.
- **Daily Matching (Barang)** — Setelah join per nomor faktur, dilakukan pengecekan agregat harian per pemasok. Jika total sehari cocok (|diff| ≤ 50), seluruh transaksi hari itu dianggap match meski ada perbedaan pada level faktur individual.
- **4-level Smart Matching (JV/Jasa)** — Karena JV tidak memiliki nomor faktur, algoritma menggunakan: (1) deteksi offset pair +/-; (2) vendor group sum; (3) single 1:1; dan (4) kombinatorial 2–4 entri Accurate terhadap 1 entri Coretax.
- **Deteksi offset pair** — Menemukan pasangan entri Accurate yang saling meniadakan (nilai +X dan -X) dan melabelinya sebagai `Offset Pair (Zeroed)`, bukan unmatched.
- **Keterangan selisih otomatis** — Setiap baris yang tidak match diberi label keterangan: `Belum Input di Accurate / Beda Masa`, `Input di Accurate Tidak Dikenal Coretax`, `Selisih Pembulatan`, atau `Selisih Nominal`.
- **Laporan terintegrasi** — Skrip 5 menggabungkan hasil Barang dan JV ke satu workbook, menyalin sheet beserta seluruh formatting, dan menambahkan baris total gabungan di Summary.
- **Auto-cleanup** — Semua file sementara dihapus dari `Dapur/` setelah proses selesai.

---

## 🔧 Prasyarat

### Python
Python **3.8+** disarankan.

### Library yang dibutuhkan

```bash
pip install pandas openpyxl xlrd xlsxwriter numpy
```

| Library | Digunakan di | Kegunaan |
|---|---|---|
| `pandas` | Skrip 1–4 | Baca Excel, filter, groupby, merge, agregasi |
| `openpyxl` | Skrip 1, 2, 5 | Baca/tulis `.xlsx`, auto-fit, salin style & merge |
| `xlrd` | Skrip 1 | Baca file legacy `.xls` dari Accurate |
| `xlsxwriter` | Skrip 3, 4 | Buat `.xlsx` dengan format accounting dan wrap text |
| `numpy` | Skrip 3 | `np.nan`, `combine_first` |
| `re`, `itertools`, `configparser`, `datetime`, `copy`, `os`, `sys` | Semua | Standard library |

> **Catatan `xlrd`:** Gunakan versi yang kompatibel dengan `.xls`:
> ```bash
> pip install "xlrd>=1.0.0,<2.0.0"
> ```

---

## 📁 Struktur Folder & File

```
📦 PPN-Masukan-Barang-Jasa/
│
├── 📄 Masukan_Barang_Jasa.py              ← Orkestrator utama. Jalankan ini
│
├── 📄 Accuratem.xls                       ← [INPUT] Ekspor buku besar PPN Masukan dari Accurate
├── 📄 Coretaxm.xlsx                       ← [INPUT] Ekspor Faktur Pajak Masukan dari Coretax DJP
│
└── 📁 Dapur/                              ← Folder pipeline (jangan diubah)
    ├── 📄 __init__.py
    ├── 📄 1_AccCleaner&PshBrgJs.py        ← Ekstrak Barang & JV dari Accuratem.xls
    ├── 📄 2_CtxPshBrgJs.py                ← Filter Barang & JV dari Coretaxm.xlsx
    ├── 📄 3_AnalyticsBrgAccCtx.py         ← Rekonsiliasi Barang (invoice matching)
    ├── 📄 4_AnalyticsJsAccCtx.py          ← Rekonsiliasi JV/Jasa (smart matching)
    ├── 📄 5_MergeHasil.py                 ← Gabung hasil Barang + JV ke satu workbook
    ├── 📄 config.conf                     ← Alias normalisasi nama pemasok
    ├── 📄 hbrg.txt                        ← Daftar nama pemasok Barang (filter Coretax)
    └── 📄 hjv.txt                         ← Daftar nama pemasok JV/Jasa (filter Coretax)
```

> Output `Hasil_Analisis_Barang_Dan_Jasa.xlsx` disalin ke folder utama setelah selesai.

---

## 🚀 Cara Penggunaan

### Langkah 1 — Siapkan file input

Letakkan dua file berikut di folder utama (sejajar dengan `Masukan_Barang_Jasa.py`):

| File | Nama wajib | Sumber |
|---|---|---|
| Ekspor Accurate | `Accuratem.xls` | Laporan buku besar PPN Masukan dari Accurate |
| Ekspor Coretax | `Coretaxm.xlsx` | Unduh data Faktur Pajak Masukan dari Coretax DJP |

> Kedua nama file bersifat **eksak dan case-sensitive**.

### Langkah 2 — Perbarui file konfigurasi

Buka tiga file konfigurasi di `Dapur/` dan sesuaikan dengan kondisi periode berjalan:

- **`hbrg.txt`** — tambah/hapus nama pemasok yang masuk kategori **Barang**
- **`hjv.txt`** — tambah/hapus nama pemasok yang masuk kategori **Jasa/JV**
- **`config.conf`** — tambah alias jika ada nama pemasok yang variasinya berbeda antara Accurate dan Coretax

Lihat panduan lengkap di bagian [Konfigurasi](#-konfigurasi).

### Langkah 3 — Jalankan

```bash
python Masukan_Barang_Jasa.py
```

### Langkah 4 — Pantau progress

```
--> Sedang menjalankan 1_AccCleaner&PshBrgJs.py...
--> Sedang membaca file Accuratem.xls...
--> Memproses Tahap A: AccBarang
--> Sukses: AccBarang.xlsx (42 baris) - Kolom F sudah Angka
--> Memproses Tahap B: AccJV
--> Sukses: AccJV.xlsx (24 baris) - Kolom D sudah Angka

--> Sedang menjalankan 2_CtxPshBrgJs.py...
--> Membaca Coretaxm.xlsx...
--> Selesai! Semua file berhasil dibuat.

--> Sedang menjalankan 3_AnalyticsBrgAccCtx.py...
--> Total Transaksi: 38 baris.
--> Transaksi yang Match: 35 baris.
--> Sisa Selisih yang harus dicek: 3 baris.
--> Selesai! File 'Hasil_Analisis_Barang_temp.xlsx' berhasil dibuat.

--> Sedang menjalankan 4_AnalyticsJsAccCtx.py...
--> Mencari angka saling hapus (offset)...
--> Melakukan pencocokan data (Smart Matching)...
--> Menyusun laporan akhir...
--> Selesai!

--> Sedang menjalankan 5_MergeHasil.py...
--> Berhasil menambahkan data dan rumus ke Summary Total.
--> Berhasil menggabungkan file dan menyimpannya...

--> Proses selesai. File hasil telah disalin ke folder utama.
```

### Langkah 5 — Buka laporan

File **`Hasil_Analisis_Barang_Dan_Jasa.xlsx`** tersedia di folder utama.

---

## 🔄 Alur Kerja Pipeline

```
[Mulai: Masukan_Barang_Jasa.py]
   │
   ├─── Validasi root: Accuratem.xls & Coretaxm.xlsx ada
   ├─── Validasi Dapur/: 9 file syarat ada
   ├─── Bersihkan Dapur/ dari *.xlsx/*.xls/*temp lama
   ├─── Salin Accuratem.xls & Coretaxm.xlsx → Dapur/
   │
   ├─── [1] 1_AccCleaner&PshBrgJs.py
   │       Baca Accuratem.xls (header=None)
   │       TAHAP A — cari seksi "P : PPN (11.00%)"
   │         Extract kolom [3,4,6,9,10,12] → 6 kolom Barang
   │         Bersihkan ,00/.0 dari No.Referensi & No.Faktur Pajak
   │         Konversi Jumlah Pajak → float
   │         → AccBarang_temp.xlsx (sheet: AccBarang)
   │       TAHAP B — cari seksi setelah "Total dari P : PPN (11.00%)"
   │               sampai "Total dari Transaksi Lainnya"
   │         Extract kolom [3,4,6,12] → 4 kolom JV
   │         → AccJV_temp.xlsx (sheet: AccJV)
   │
   ├─── [2] 2_CtxPshBrgJs.py
   │       Baca Coretaxm.xlsx (sheet: "data")
   │       Parse kolom tanggal → format Indonesia "D Mon YYYY"
   │       Filter dengan hbrg.txt → CtxBarang_temp.xlsx (sheet: CoretaxBarang)
   │       Filter dengan hjv.txt → CtxJV_temp.xlsx (sheet: CoretaxJV)
   │       (filter: case-insensitive substring match pada kolom nama penjual)
   │
   ├─── [3] 3_AnalyticsBrgAccCtx.py
   │       Baca config.conf → build alias_dict
   │       Baca CtxBarang_temp + AccBarang_temp
   │       Normalisasi keduanya: clean_invoice, clean_name, parse_date
   │       Aggregate per (clean_invoice, clean_name)
   │       OUTER JOIN pada clean_invoice
   │       Hitung Selisih = PPN_Coretax - PPN_Accurate
   │       Daily Matching: cek agregat harian per pemasok (threshold ≤ 50)
   │       Klasifikasi unmatched → Keterangan otomatis
   │       → Hasil_Analisis_Barang_temp.xlsx
   │           Sheet: Summary Total, Rincian Selisih, Detail Data
   │
   ├─── [4] 4_AnalyticsJsAccCtx.py
   │       Baca CtxJV_temp + AccJV_temp
   │       LEVEL 1: Deteksi offset pair +X/-X di Accurate
   │       LEVEL 2: Vendor Group — jumlah Coretax per vendor vs satu entri Accurate
   │       LEVEL 3: Single 1:1 — satu Coretax vs satu Accurate (tolerance ≤ 50)
   │       LEVEL 4: Kombinasi — satu Coretax vs 2–4 entri Accurate dalam window 60 hari
   │       Sisa unmatched → "Unmatched Coretax" / "Unmatched Accurate"
   │       → Hasil_Analisis_JV_temp.xlsx (sheet: Analisis)
   │
   ├─── [5] 5_MergeHasil.py
   │       Salin sheet "Analisis" dari JV → tambahkan sebagai "Analisis JV" di file Barang
   │         (salin: nilai sel, font, border, fill, alignment, merge cell, kolom/baris)
   │       Baca GRAND TOTAL dari baris terakhir "Analisis JV"
   │       Tambah baris "JV" di "Summary Total" + rumus grand total gabungan
   │       → Hasil_Analisis_Barang_Dan_Jasa.xlsx
   │
   ├─── Validasi output ada
   ├─── Salin ke folder utama
   └─── Bersihkan Dapur/ dari file sementara & input
```

---

## 🔍 Detail Tiap Skrip

### Skrip 1 — Ekstraksi Accurate

Membaca `Accuratem.xls` tanpa header (header=None) karena file ekspor Accurate mengandung judul laporan, metadata, dan berbagai seksi sebelum data sesungguhnya.

Pendeteksian seksi menggunakan **penanda teks eksak** di kolom ke-3 (index 2):

| Penanda | Arti |
|---|---|
| `P : PPN (11.00%)` | Awal seksi Barang |
| `Total dari P : PPN (11.00%)` | Akhir seksi Barang + awal area JV |
| `Total dari Transaksi Lainnya` | Akhir seksi JV |

**Kolom yang diekstrak (Barang):**

| Indeks asli | Nama kolom baru |
|---|---|
| 3 | `Tanggal` |
| 4 | `Tgl. Pajak` |
| 6 | `No. Referensi` |
| 9 | `No. Faktur Pajak` |
| 10 | `Nama Pemasok` |
| 12 | `Jumlah Pajak` |

**Kolom yang diekstrak (JV):**

| Indeks asli | Nama kolom baru |
|---|---|
| 3 | `Tanggal` |
| 4 | `Tgl. Pajak` |
| 6 | `No. Referensi` |
| 12 | `Jumlah Pajak` |

> JV tidak memiliki kolom No. Faktur Pajak dan Nama Pemasok karena transaksi Accurate tipe JV tidak merekam informasi tersebut.

---

### Skrip 2 — Filter Coretax

Membaca `Coretaxm.xlsx` dari sheet bernama **`data`** (wajib ada). Mendukung dua versi format kolom Coretax:

| Data | Versi kolom lama | Versi kolom baru |
|---|---|---|
| Nama penjual | `Nama Penjual` | `Nama Penjual Barang Kena Pajak/...` |
| Nomor faktur | `Nomor Faktur Pajak` | `Faktur Pajak/... - Nomor` |
| Tanggal | `Tanggal Faktur Pajak` | `Faktur Pajak/... - Tanggal` |

Kolom tanggal dikonversi ke format teks Indonesia: `"1 Jan 2025"`, `"15 Des 2025"`.

Filtering dilakukan via substring match case-insensitive menggunakan pola gabungan `|` (OR) dari semua keyword di masing-masing file teks.

**Logika Filtering & Filter Kondisional:**
1. Membaca rule dari `hbrg.txt` dan `hjv.txt`.
2. Mendukung 2 mode penulisan di file filter:
   - **Mode Standar (`NAMA_VENDOR`)**: Memfilter berdasarkan nama penjual (substring match, case-insensitive).
   - **Mode Kondisional (`NAMA_VENDOR | NOMINAL`)**: Memfilter berdasarkan nama penjual **DAN** mencocokkan nilai PPN dengan nominal yang ditentukan (toleransi selisih < 1.0).
3. Hasil pemisahan disimpan ke `CtxBarang_temp.xlsx` (sheet: `CoretaxBarang`) dan `CtxJV_temp.xlsx` (sheet: `CoretaxJV`).
4. Menerapkan fungsi `auto_fit_columns()` pada file temporary Excel yang dihasilkan.

---

### Skrip 3 — Rekonsiliasi Barang

Rekonsiliasi berbasis **nomor faktur pajak** sebagai kunci utama. Pipeline normalisasi diterapkan sebelum pencocokan:

**`clean_invoice_number()`:**
```
"010.000-25.12345678" → "01000025124345678"
"1.23e+10"           → "12300000000"
```
Menghapus semua karakter non-alphanumerik dan menangani notasi saintifik.

**`normalize_name()`:**
```
"PT. Shell Indonesia" → "SHELL INDONESIA"
"GAJAH TUNGGAL"      → "GAJAH TUNGGAL TBK"  (via alias config.conf)
```
Uppercase → terapkan alias → hapus prefix PT./CV./UD./FA.

**Logika matching:**

Setelah OUTER JOIN pada `clean_invoice`:
1. Hitung `Selisih = PPN_Coretax - PPN_Accurate`
2. **Daily Matching** — untuk tiap kombinasi (Nama Penjual, Tanggal), jumlahkan semua baris di hari itu. Jika `|total_Coretax - total_Accurate| ≤ 50` → seluruh baris hari itu dianggap **match** meski ada perbedaan per faktur
3. Baris yang tidak match (|Selisih| > 50 DAN tidak tercover Daily Match) → masuk Rincian Selisih

---

### Skrip 4 — Rekonsiliasi JV/Jasa

JV tidak memiliki nomor faktur, sehingga pencocokan dilakukan berdasarkan **jumlah PPN dan tanggal** melalui empat level secara berurutan:

**Level 1 — Deteksi Offset Pair:**
```
Accurate: +5.500.000 (Jan) dan -5.500.000 (Jan) → saling hapus
Label: "Offset Pair (Zeroed)", Selisih = 0
```
Pasangan yang terdeteksi dikeluarkan dari pool sebelum level berikutnya.

**Level 2 — Vendor Group Match:**
```
Coretax: vendor X total = 8.000.000
Accurate: satu entri 8.000.000 (±50) → Match
Label: "Matched (Vendor Group)"
```
Menggabungkan semua entri Coretax dari vendor yang sama, bandingkan ke satu entri Accurate.

**Level 3 — Single Match (1:1):**
```
Coretax: 3.500.000 ↔ Accurate: 3.500.000 (±50) → Match
Label: "Matched (Single)"
```
Satu entri Coretax dicocokkan dengan satu entri Accurate terbaik (diff terkecil).

**Level 4 — Combination Match:**
```
Coretax: 9.000.000 ↔ Accurate: 4.500.000 + 4.500.000 = 9.000.000 (±50) → Match
Label: "Matched (Combination)"
```
Satu entri Coretax dicocokkan terhadap kombinasi 2–4 entri Accurate dalam window 60 hari (`itertools.combinations`). Dibatasi pada pool ≤ 40 entri untuk performa.

**Entri yang tidak match setelah Level 4:**
- Coretax tanpa pasangan → `"Unmatched Coretax"`
- Accurate tanpa pasangan → `"Unmatched Accurate"`

---

### Skrip 5 — Gabung Hasil Akhir

Menggabungkan dua file hasil ke satu workbook menggunakan `openpyxl` dengan penyalinan styling lengkap (font, border, fill, alignment, merge cell, lebar kolom, tinggi baris).

Setelah sheet "Analisis JV" ditambahkan, skrip memperbarui sheet "Summary Total" dengan:

```
[baris pemasok Barang 1]
[baris pemasok Barang N]
TOTAL         [=SUM Barang Coretax]   [=SUM Barang Accurate]
JV            [Grand Total JV Ctx]    [Grand Total JV Acc]
              [=B_TOTAL + B_JV]       [=C_TOTAL + C_JV]   [=B_SUM - C_SUM]
```
Formula kolom gabungan ditulis sebagai rumus Excel (`=B{row}+B{jv_row}`), bukan nilai statis.

---

## ⚙️ Konfigurasi

### `config.conf` — Alias nama pemasok

```ini
[ALIAS]
GAJAH TUNGGAL = GAJAH TUNGGAL TBK
SPEEDWORK SOLUSI UTAMA = SPEEDWORK SOLUSI UTAMA
SUMBERHARPINDO LESTARISENTOSA = SUMBERHARPINDO LESTARI SENTOSA
SHELL INDONESIA = SHELL INDONESIA
PT SAMA SAMA MITRA MAJU = SAMA SAMA MITRA MAJU
PT SAHASRABHANU CIPTA KARYA = SAHASRABHANU CIPTA KARYA
```

**Format:** `KUNCI_ALIAS = NAMA_KANONIK`

**Cara kerja:** Sebelum pencocokan, setiap nama pemasok diubah ke uppercase lalu dicek apakah mengandung `KUNCI_ALIAS` (substring match, case-sensitive setelah uppercase). Jika cocok, nama diganti dengan `NAMA_KANONIK`.

**Kapan perlu ditambah:** Jika nama pemasok yang sama ditulis berbeda antara Accurate dan Coretax. Contoh: Accurate mencatat `"GAJAH TUNGGAL"` sementara Coretax mencatat `"PT GAJAH TUNGGAL TBK"` → tambah alias `GAJAH TUNGGAL = GAJAH TUNGGAL TBK` agar keduanya menuju nama kanonik yang sama.

> Alias diterapkan **setelah** normalisasi uppercase dan **sebelum** penghapusan prefix. Urutan entri di `config.conf` tidak berpengaruh.

---

### `hbrg.txt` & `hjv.txt` — Daftar pemasok & Rule Kondisional

File ini menentukan pemisahan transaksi dari Coretax ke kelompok **Barang** atau **JV/Jasa**.

**Sintaks yang didukung:**

1. **Filter Nama Vendor (Standar)**
   ```txt
   ARVIA JAYA
   ASTRA INTERNATIONAL TBK
   SHELL INDONESIA
   ...

Satu nama per baris. Setiap nama digunakan sebagai **substring filter** (case-insensitive) terhadap kolom nama penjual di `Coretaxm.xlsx`. Baris Coretax yang nama penjualnya mengandung salah satu dari nama-nama ini akan dimasukkan ke `CtxBarang_temp.xlsx`.

> **Penting:** Nama di sini dicocokkan ke data Coretax **sebelum** normalisasi alias. Gunakan nama sebagaimana tertulis di Coretax.

---

2. **Filter Nama Vendor (Kondisional)**
   ```txt
ADI SARANA ARMADA TBK
CAKRAWALA PUTRA NUSANTARA
NASMOCO
SERASI AUTORAYA
...
PT TELEKOMUNIKASI INDONESIA | 550.000
PLN PERSERO | 1.250.000,00

Format dan mekanisme sama dengan `hbrg.txt`, namun hasilnya masuk ke `CtxJV_temp.xlsx`.

---

## 📋 Format File Input

### `Accuratem.xls`

Ekspor buku besar PPN Masukan dari Accurate. **Syarat kritis:**
- Format `.xls` (bukan `.xlsx`)
- Harus mengandung dua seksi dengan penanda teks tepat:
  - `"P : PPN (11.00%)"` — awal seksi Barang
  - `"Total dari P : PPN (11.00%)"` — akhir seksi Barang
  - `"Total dari Transaksi Lainnya"` — akhir seksi JV
- Data Barang harus berada di kolom indeks 3, 4, 6, 9, 10, 12
- Data JV harus berada di kolom indeks 3, 4, 6, 12

> Jika penanda teks tidak ditemukan, skrip menampilkan error `"Gagal Tahap A/B: Marker tidak ditemukan"` dan melewati seksi tersebut.

### `Coretaxm.xlsx`

Ekspor Faktur Pajak Masukan dari portal Coretax DJP. **Syarat kritis:**
- Format `.xlsx`
- Harus memiliki sheet bernama persis **`data`** (huruf kecil semua)
- Harus mengandung kolom nama penjual (salah satu nama yang dikenali)
- Harus mengandung kolom tanggal dan PPN (salah satu nama yang dikenali)

---

## 📤 Output & Struktur Laporan

### File output: `Hasil_Analisis_Barang_Dan_Jasa.xlsx`

Berisi 4 sheet:

#### Sheet 1 — `Summary Total`

Ringkasan PPN per pemasok untuk Barang, diikuti total JV dan grand total gabungan.

| Kolom | Format |
|---|---|
| Nama Penjual | Teks |
| PPN Coretax | `_(* #,##0_)` (accounting) |
| PPN Accurate | `_(* #,##0_)` (accounting) |
| Selisih | `=PPN Coretax - PPN Accurate` (rumus Excel) |

Baris-baris spesial di akhir:
- `TOTAL` — sum semua pemasok Barang
- `JV` — grand total dari analisis JV (diisi oleh Skrip 5)
- Baris terakhir — rumus `=TOTAL + JV` per kolom (jumlah gabungan final)

#### Sheet 2 — `Rincian Selisih`

Daftar transaksi Barang yang **tidak dapat direkonsiliasi** (|Selisih| > 50 dan tidak tercover Daily Match).

| Kolom | Keterangan |
|---|---|
| Nama Penjual | Nama pemasok (setelah normalisasi) |
| Nomor Faktur | Nomor faktur yang bersangkutan |
| Tanggal Coretax | Tanggal di data Coretax |
| Tanggal Accurate | Tanggal di data Accurate |
| PPN Coretax | Nilai PPN di Coretax |
| PPN Accurate | Nilai PPN di Accurate |
| Selisih | PPN Coretax − PPN Accurate |
| Keterangan | Klasifikasi otomatis (lihat tabel di bawah) |

#### Sheet 3 — `Detail Data`

Semua transaksi Barang (matched maupun unmatched), berguna untuk verifikasi manual.

#### Sheet 4 — `Analisis JV`

Hasil rekonsiliasi JV/Jasa dalam format berbeda karena tidak berbasis nomor faktur.

| Kolom | Keterangan |
|---|---|
| Tanggal Coretax | Tanggal dari sisi Coretax |
| Nama Pemasok | Nama vendor Coretax |
| PPN Coretax | Nilai PPN dari Coretax |
| Rincian PPN Accurate | Deskripsi multi-baris entri Accurate yang dicocokkan |
| Total PPN Accurate | Total nilai Accurate yang dipasangkan |
| Selisih | PPN Coretax − Total PPN Accurate |
| Keterangan | Status matching (lihat tabel di bawah) |

---

## 🏷️ Logika Pencocokan & Keterangan

### Keterangan Barang (Skrip 3)

| Keterangan | Kondisi | Interpretasi |
|---|---|---|
| *(tanpa keterangan / masuk Daily Match)* | `|Selisih| ≤ 50` atau tercakup Daily Match | Data cocok antara Coretax dan Accurate |
| `Belum Input di Accurate / Beda Masa` | `PPN_Accurate = 0` | Faktur ada di Coretax tapi belum diinput di Accurate, atau berbeda masa pajak |
| `Input di Accurate Tidak Dikenal Coretax` | `PPN_Coretax = 0` | Faktur diinput di Accurate tapi tidak ditemukan di Coretax |
| `Selisih Pembulatan` | `|Selisih| ≤ 100` | Selisih kecil akibat perbedaan pembulatan antar sistem |
| `Selisih Nominal (Input Tidak Sesuai)` | `|Selisih| > 100` | Nominal PPN yang dicatat berbeda signifikan antara kedua sistem |

### Keterangan JV (Skrip 4)

| Keterangan | Interpretasi |
|---|---|
| `Matched (Vendor Group)` | Total semua Coretax satu vendor cocok dengan satu entri Accurate |
| `Matched (Single)` | Satu entri Coretax cocok 1:1 dengan satu entri Accurate |
| `Matched (Combination)` | Satu entri Coretax cocok dengan gabungan 2–4 entri Accurate |
| `Offset Pair (Zeroed)` | Dua entri Accurate saling meniadakan (+X dan -X) |
| `Unmatched Coretax` | Entri Coretax tidak memiliki pasangan di Accurate |
| `Unmatched Accurate` | Entri Accurate tidak memiliki pasangan di Coretax |

### Toleransi matching

| Konteks | Threshold | Keterangan |
|---|---|---|
| Daily Match (Barang) | ≤ 50 | Perbedaan harian dianggap match |
| Keterangan "Selisih Pembulatan" | ≤ 100 | Threshold lebih longgar untuk identifikasi |
| Smart Matching JV | ≤ 50 | Berlaku di semua level (group, single, kombinasi) |

---

## 🛠️ Troubleshooting

### ❌ `File berikut tidak ditemukan di folder utama: Accuratem.xls`
Pastikan file ada di folder utama dengan nama persis `Accuratem.xls` (termasuk huruf besar `A` dan ekstensi `.xls`).

### ❌ `Gagal Tahap A: Marker tidak ditemukan`
Skrip tidak menemukan teks `"P : PPN (11.00%)"` di kolom ke-3 file Accurate. Pastikan ekspor dari Accurate menggunakan format yang sama dan mencakup seksi PPN Masukan dari pembelian barang.

### ❌ `Gagal Tahap B: Penutup 'Total dari Transaksi Lainnya' tidak ditemukan`
Tidak ada transaksi JV di file ekspor, atau penanda teks berbeda. Jika memang tidak ada transaksi JV dalam periode ini, `AccJV_temp.xlsx` tidak akan terbentuk dan Skrip 4 akan gagal dengan error.

### ❌ `Error: File Coretaxm.xlsx tidak ditemukan` atau `sheet 'data' tidak ada`
Pastikan nama file persis `Coretaxm.xlsx` dan sheet di dalamnya bernama `data` (huruf kecil semua).

### ❌ `Kolom Coretax berikut tidak ditemukan: [...]`
Format ekspor Coretax yang Anda gunakan tidak mengandung nama kolom yang dikenali. Periksa nama kolom aktual di file Coretax dan tambahkan sebagai kandidat di fungsi `find_column()` di Skrip 3.

### ❌ Hasil `CtxBarang_temp.xlsx` kosong (0 baris)
Tidak ada nama pemasok di `hbrg.txt` yang cocok dengan nama di Coretax. Buka `Coretaxm.xlsx` → sheet `data` → lihat isi kolom nama penjual → perbarui `hbrg.txt` sesuai nama yang tertulis di sana.

### ❌ Hampir semua transaksi Barang masuk "Belum Input di Accurate"
Kemungkinan: (1) normalisasi nama tidak berhasil karena variasi penulisan — tambah alias ke `config.conf`; (2) nomor faktur di Accurate direkam dalam format berbeda dari Coretax — periksa isi `AccBarang_temp.xlsx` dan `CtxBarang_temp.xlsx`.

### ❌ `Gagal membuka file: Hasil_Analisis_JV_temp.xlsx`
Skrip 4 gagal menghasilkan file ini. Jalankan Skrip 4 secara manual untuk melihat error lengkap:
```bash
cd Dapur
python 4_AnalyticsJsAccCtx.py
```

### ❌ `Sheet 'Summary Total' tidak ditemukan di file barang`
Skrip 3 gagal menghasilkan sheet tersebut. Pastikan Skrip 3 selesai tanpa error sebelum Skrip 5 dijalankan.

---

## 📌 Catatan Penting

- **Kedua file input disalin, bukan dipindahkan** — `Accuratem.xls` dan `Coretaxm.xlsx` di folder utama tetap aman setelah proses selesai.
- **Toleransi 50 bukan mutlak** — Threshold ≤ 50 digunakan sebagai definisi "match". Transaksi dengan selisih di atas 50 perlu dicek manual meski mungkin hanya selisih pembulatan dalam beberapa kasus.
- **Daily Match di Barang bersifat agregat** — Dua faktur berbeda dari pemasok yang sama di hari yang sama bisa saling mengkompensasi dalam Daily Match, meski masing-masing punya selisih individual. Ini disengaja untuk mengakomodasi perbedaan cara sistem memecah atau menggabungkan faktur.
- **Kombinasi JV dibatasi 2–4 entri dan pool ≤ 40** — Batasan ini ditetapkan untuk menjaga performa. Dataset dengan transaksi JV yang sangat banyak dan tidak terstruktur mungkin memerlukan penyesuaian parameter.
- **Alias di `config.conf` case-sensitive setelah uppercase** — Semua kunci alias dicocokkan setelah nama sudah di-uppercase. Pastikan kunci alias ditulis dalam huruf kapital semua.
- **Output ditimpa setiap run** — `Hasil_Analisis_Barang_Dan_Jasa.xlsx` di folder utama akan langsung diganti setiap kali pipeline dijalankan. Simpan salinan jika hasil sebelumnya perlu dipertahankan.
- **Jangan ubah nama sheet `data` di Coretaxm.xlsx** — Skrip 2 membaca sheet secara eksak berdasarkan nama `"data"`.

---

## 📜 Lisensi

Proyek ini dikembangkan untuk keperluan internal perusahaan. Silakan sesuaikan dengan kebutuhan organisasi Anda.

---

*Dikembangkan oleh [ACC-TAX-REIGHTEEN](https://github.com/ACC-TAX-REIGHTEEN)*
