# 📦 Ekstrak PDF Gudang Shell

Skrip otomasi Python untuk **mengunduh lampiran PDF faktur dari Gmail**, **mengekstrak data faktur Shell** dari PDF tersebut, mencocokkannya dengan data export CTX (Nomor Faktur Pajak), lalu **menyusunnya ke dalam template laporan rekapitulasi gudang** secara otomatis per bulan.

---

## 📋 Daftar Isi

- [Latar Belakang](#latar-belakang)
- [Fitur Utama](#fitur-utama)
- [Struktur Folder](#struktur-folder)
- [Persyaratan Sistem](#persyaratan-sistem)
- [Instalasi](#instalasi)
- [Konfigurasi](#konfigurasi)
  - [credentials.json & token.json (Gmail API)](#credentialsjson--tokenjson-gmail-api)
  - [gmail.conf](#gmailconf)
  - [gudang.conf](#gudangconf)
- [File Input yang Dibutuhkan](#file-input-yang-dibutuhkan)
- [Cara Penggunaan](#cara-penggunaan)
  - [Menu 1: Ambil Data dari Gmail dan Proses](#menu-1-ambil-data-dari-gmail-dan-proses)
  - [Menu 2: Hanya Proses Pengecekan Nomor Faktur Pajak](#menu-2-hanya-proses-pengecekan-nomor-faktur-pajak)
- [Alur Kerja Internal (Pipeline)](#alur-kerja-internal-pipeline)
- [Penjelasan File di Folder Dapur](#penjelasan-file-di-folder-dapur)
- [Logika Pencocokan Nomor Faktur Pajak](#logika-pencocokan-nomor-faktur-pajak)
- [File Output](#file-output)
- [Penanganan Error](#penanganan-error)
- [Catatan Penting & Keamanan](#catatan-penting--keamanan)
- [Lisensi](#lisensi)

---

## 📌 Latar Belakang

Setiap bulan, faktur penjualan dari Shell dikirim melalui email dalam format PDF ke alamat email perusahaan. Faktur-faktur ini perlu direkap ke dalam laporan gudang per cabang (SMG, PATI, PWT, dst), dicocokkan dengan Nomor Faktur Pajak dari data export sistem pajak (CTX), lalu disusun rapi per bulan dalam satu file Excel rekap. Proyek ini mengotomasi seluruh proses tersebut — dari pengambilan email, ekstraksi PDF, deduplikasi, pencocokan data, hingga penyusunan laporan akhir — tanpa perlu input manual satu per satu.

---

## ✨ Fitur Utama

- ✅ **Auto-download lampiran PDF** dari Gmail menggunakan Gmail API (dengan filter tanggal & query pencarian)
- ✅ **Ekstraksi data otomatis** dari PDF faktur Shell (No. Invoice, No. SJ, No. PO, Tanggal, DPP, Total, Jatuh Tempo, dll.) menggunakan `pdfplumber`
- ✅ **Pemetaan kode gudang otomatis** (mis. `SMG` → `SMG`, `MAKASSAR`/`MAKSSAR` → `MKS`) lewat file konfigurasi
- ✅ **Filter rentang tanggal** baik untuk pencarian email maupun ekstraksi data PDF
- ✅ **Deduplikasi data** di beberapa tahap (per baris identik & per kombinasi Tanggal+No Inv+No SJ+No PO+Tgl FP+DPP)
- ✅ **Pencocokan otomatis Nomor Faktur Pajak** dari file `data_export` (CTX) berdasarkan kombinasi Tanggal Faktur Pajak + DPP, menggunakan sistem antrean (FIFO) agar tidak salah pasang saat ada nilai duplikat
- ✅ **Penggabungan data baru dengan data lama** (incremental) tanpa duplikasi
- ✅ **Penyisipan otomatis ke template laporan bulanan** — bila bulan sudah ada blok di template, data baru disisipkan; bila belum ada, blok baru dibuat otomatis lengkap dengan style, border, dan rumus `SUM`
- ✅ **Auto-update rumus total (SUM)** dan perbaikan merge cell antar blok "GUDANG"
- ✅ **Mode alternatif**: hanya mencocokkan Nomor Faktur Pajak tanpa mengulang seluruh pipeline ekstraksi
- ✅ Pembersihan otomatis seluruh file sementara (`*temp.xlsx`, `*.pdf`) setelah proses selesai

---

## 🗂️ Struktur Folder

```
Ekstrak-PDF-Gudang-Shell-main/
│
├── Proses Faktur SHELL.py          ← Script utama (orchestrator/menu)
├── LICENSE.txt
├── README.md
│
└── Dapur/                          ← Folder engine pemrosesan (jangan diubah)
    ├── 1_AmbilLampiranGmail.py
    ├── 2_EkstrakPdfShell.py
    ├── 3_HelperDeleteDuplicate.py
    ├── 4_XlookupData.py
    ├── 5_KnifeToOperationFinalData.py
    ├── 6_CombineDeleteDuplicateAndSort.py
    ├── 7_HelperDeleteDuplicateFinalData.py
    ├── 8_XlookupData.py
    ├── 9_CopyDataToTemplate.py
    ├── 10_HelperDeleteTemplateData.py
    ├── 11_HelperMergedSUM.py
    ├── A_JustXlookupData.py
    │
    ├── TEMPLATE.xlsx                ← Template laporan rekap gudang (jangan diubah)
    ├── CTX.xlsx                     ← Template dummy data_export (fallback)
    │
    ├── credentials.json             ← Kredensial OAuth Gmail API (RAHASIA)
    ├── token.json                   ← Token akses Gmail tersimpan (RAHASIA, auto-generate)
    ├── gmail.conf                   ← Konfigurasi pencarian & filter email
    └── gudang.conf                  ← Konfigurasi pemetaan kode gudang & filter tanggal
```

---

## 💻 Persyaratan Sistem

| Komponen | Keterangan |
|---|---|
| OS | Windows / Linux / macOS (tidak bergantung pada Excel/COM, murni Python) |
| Python | 3.8 atau lebih baru |
| Akun Google | Wajib, dengan akses Gmail API diaktifkan (untuk Menu 1) |
| Koneksi Internet | Wajib untuk Menu 1 (akses Gmail API) |

---

## 📦 Instalasi

**1. Clone atau download repositori ini**

```bash
git clone https://github.com/<username>/Ekstrak-PDF-Gudang-Shell.git
cd Ekstrak-PDF-Gudang-Shell-main
```

**2. Install semua dependensi Python**

```bash
pip install pandas openpyxl pdfplumber google-auth google-auth-oauthlib google-api-python-client
```

Atau jika tersedia file `requirements.txt`:

```bash
pip install -r requirements.txt
```

**Ringkasan library yang digunakan:**

| Library | Kegunaan |
|---|---|
| `pandas` | Membaca, menggabungkan, dan mentransformasi data Excel |
| `openpyxl` | Membaca, menulis, dan memformat file `.xlsx` (termasuk merge cell, border, style) |
| `pdfplumber` | Mengekstrak teks dan tabel dari file PDF faktur |
| `google-auth`, `google-auth-oauthlib` | Autentikasi OAuth 2.0 ke akun Google |
| `google-api-python-client` | Mengakses Gmail API untuk mencari & mengunduh lampiran email |

---

## ⚙️ Konfigurasi

### credentials.json & token.json (Gmail API)

Diperlukan agar script dapat mengakses Gmail melalui API resmi Google.

1. Buka [Google Cloud Console](https://console.cloud.google.com/), buat project baru.
2. Aktifkan **Gmail API** untuk project tersebut.
3. Buat kredensial OAuth Client ID jenis **Desktop App**.
4. Download file JSON kredensial, lalu ganti isi `Dapur/credentials.json` dengan isi file tersebut.
5. Saat script `1_AmbilLampiranGmail.py` dijalankan pertama kali, browser akan terbuka meminta login & izin akses — `token.json` akan dibuat/diperbarui otomatis setelahnya.

> File `token.json` di repositori ini **masih kosong** dan akan terisi otomatis setelah proses autentikasi pertama berhasil.

### gmail.conf

Mengatur kriteria pencarian email dan folder output unduhan.

```ini
[DEFAULT]
output_folder = .

[SEARCH_CONFIG]
gmail_query = from:*@shell.com has:attachment filename:pdf after:2026/05/01 before:2026/07/31
strict_start_date = 2026-05-01
strict_end_date   = 2026-07-31
filename_must_contain =
```

| Parameter | Keterangan |
|---|---|
| `output_folder` | Folder tujuan menyimpan PDF yang diunduh (`.` = folder Dapur) |
| `gmail_query` | Query pencarian Gmail (format sama seperti search bar Gmail) |
| `strict_start_date` / `strict_end_date` | Filter tanggal tambahan di sisi Python (format `YYYY-MM-DD`), sebagai pengaman ganda selain `gmail_query` |
| `filename_must_contain` | Hanya unduh lampiran yang nama filenya mengandung teks ini (kosongkan untuk tidak memfilter) |

### gudang.conf

Mengatur pemetaan **kode gudang** (dari teks "Your Reference" di PDF) ke **nama sheet** di file hasil, serta filter tanggal untuk ekstraksi PDF.

```ini
SMG = SMG
PATI = PA
PWT = PWT
TGL = TGL
SOLO = SL
MGL = MGL
YOGYA = YY
JOGJA = YY
MKSR = MKS
MAKASSAR = MKS
MAKSSAR = MKS
PALU = PLU
KENDARI = KDR

TANGGAL_DARI = 01/02/2026
TANGGAL_SAMPAI = 28/02/2026
```

| Parameter | Keterangan |
|---|---|
| `<KODE_SUMBER> = <KODE_SHEET>` | Pemetaan banyak-ke-satu; beberapa variasi penulisan (mis. `MAKASSAR`, `MAKSSAR`, `MKSR`) bisa dipetakan ke satu kode sheet yang sama (`MKS`) |
| `TANGGAL_DARI` / `TANGGAL_SAMPAI` | Filter rentang tanggal faktur (format `DD/MM/YYYY`) yang akan disertakan dalam hasil ekstraksi PDF |

> Gudang yang kodenya **tidak ditemukan** dalam pemetaan ini akan otomatis dikelompokkan ke sheet `DATA_LAIN`.

---

## 📥 File Input yang Dibutuhkan

| Nama File | Wajib? | Keterangan |
|---|---|---|
| `data_export_*.xlsx` | Opsional | Data export dari sistem pajak (CTX), berisi sheet `data` dengan kolom `Tanggal Faktur Pajak`, `Harga Jual/Penggantian/DPP`, `Nomor Faktur Pajak`. Jika tidak tersedia, sistem otomatis membuat dummy dari `Dapur/CTX.xlsx` |
| `Laporan SHELL*.xlsx` | Opsional | File rekap lama (jika sudah ada riwayat sebelumnya) untuk digabung dengan data baru |
| Lampiran PDF Faktur Shell | Otomatis | Diunduh otomatis dari Gmail oleh script `1_AmbilLampiranGmail.py`, tidak perlu disiapkan manual |

Letakkan `data_export_*.xlsx` dan/atau `Laporan SHELL*.xlsx` di **folder utama** (sejajar dengan `Proses Faktur SHELL.py`) sebelum menjalankan script — keduanya akan otomatis disalin/dipindahkan ke folder `Dapur`.

---

## 🚀 Cara Penggunaan

1. Pastikan `credentials.json` sudah dikonfigurasi dan `gmail.conf` / `gudang.conf` sudah disesuaikan.
2. (Opsional) Letakkan `data_export_*.xlsx` dan/atau `Laporan SHELL*.xlsx` di folder utama.
3. Jalankan script utama:

```bash
python "Proses Faktur SHELL.py"
```

4. Script akan menampilkan menu pilihan:

```
--> Pilih Menu Proses:
--> 1. Ambil data dari Gmail dan proses
--> 2. Hanya proses pengecekan nomor Faktur Pajak
--> Masukkan pilihan (1/2):
```

### Menu 1: Ambil Data dari Gmail dan Proses

Alur lengkap dari nol: unduh PDF dari Gmail → ekstrak → bersihkan → cocokkan No. Faktur Pajak → gabung ke laporan lama (jika ada) → susun ke template bulanan.

Output akhir: **`Laporan SHELL BARU.xlsx`** di folder utama.

### Menu 2: Hanya Proses Pengecekan Nomor Faktur Pajak

Digunakan saat sudah memiliki file `Laporan SHELL*.xlsx` (hasil ekstraksi manual/sebelumnya) dan `data_export_*.xlsx`, namun kolom **No FP** belum terisi. Mode ini **melewati seluruh proses pengunduhan & ekstraksi PDF**, langsung mencocokkan Nomor Faktur Pajak ke file Laporan SHELL yang sudah ada.

> Menu ini mengharuskan kedua file (`data_export_*.xlsx` dan `Laporan SHELL*.xlsx`) sudah berada di folder utama sebelum dijalankan, karena keduanya disalin ke `Dapur` di awal proses.

---

## ⚙️ Alur Kerja Internal (Pipeline)

### Pipeline Menu 1 (lengkap)

```
[START]
   │
   ├─ Validasi folder Dapur & seluruh file wajib di dalamnya
   │
   ├─ Salin data_export*.xlsx (jika ada) ke Dapur
   ├─ Pindahkan Laporan SHELL*.xlsx (jika ada) ke Dapur
   ├─ Jika data_export tidak ada → buat dummy dari CTX.xlsx (nama acak)
   │
   ├── [1] 1_AmbilLampiranGmail.py
   │     └─ Login Gmail API → cari email sesuai gmail_query
   │        Unduh semua lampiran PDF yang lolos filter tanggal & nama file
   │        Simpan ke folder Dapur dengan prefix tanggal & message-id
   │
   ├── [2] 2_EkstrakPdfShell.py
   │     └─ Baca semua PDF di folder Dapur dengan pdfplumber
   │        Ekstrak: No Inv, Tanggal, No SJ, No PO, DPP, Rp, JT, dll.
   │        Petakan kode gudang via gudang.conf, filter tanggal via TANGGAL_DARI/SAMPAI
   │        Simpan per sheet gudang → Hasil_Ekstrak_temp.xlsx
   │
   ├── [3] 3_HelperDeleteDuplicate.py
   │     └─ Hapus baris yang identik persis di setiap sheet Hasil_Ekstrak_temp.xlsx
   │
   ├── [4] 4_XlookupData.py
   │     └─ Cocokkan Nomor Faktur Pajak dari data_export ke Hasil_Ekstrak_temp.xlsx
   │        berdasarkan kombinasi (Tanggal, DPP) — sistem antrean FIFO
   │
   ├── [5] 5_KnifeToOperationFinalData.py
   │     └─ Jika ada Laporan SHELL*.xlsx (riwayat lama):
   │          bersihkan & rapikan format → Hasil_Ekstrak_Shell_temp.xlsx
   │        (baris dengan tanggal dummy 01/01/2001 dibuang)
   │
   ├── [6] 6_CombineDeleteDuplicateAndSort.py
   │     └─ Gabungkan data lama (Hasil_Ekstrak_Shell_temp.xlsx) dengan data baru
   │        (Hasil_Ekstrak_temp.xlsx), hapus duplikat, urutkan berdasarkan tanggal
   │
   ├── [7] 7_HelperDeleteDuplicateFinalData.py
   │     └─ Deduplikasi lanjutan berdasarkan kombinasi
   │        (Tanggal, No Inv, No SJ, No PO, Tgl FP, DPP)
   │        Jika ada duplikat, prioritaskan baris yang sudah punya No FP
   │
   ├── [8] 8_XlookupData.py
   │     └─ Ulangi pencocokan Nomor Faktur Pajak (mengisi baris yang masih kosong)
   │
   ├── [9] 9_CopyDataToTemplate.py
   │     └─ Salin TEMPLATE.xlsx → TEMPLATE_temp.xlsx
   │        Kelompokkan data per (tahun, bulan)
   │        Jika blok bulan sudah ada di template → sisipkan baris baru sebelum baris SUM
   │        Jika blok bulan belum ada → buat blok baru lengkap (header, style, border, rumus SUM)
   │
   ├── [10] 10_HelperDeleteTemplateData.py
   │     └─ Hapus baris contoh/dummy "JANUARI 2001" bawaan template (jika masih ada)
   │
   ├── [11] 11_HelperMergedSUM.py
   │     └─ Perbaiki merge cell & rumus SUM di setiap batas blok "GUDANG"
   │
   ├─ Salin TEMPLATE_temp.xlsx → "Laporan SHELL BARU.xlsx" di folder utama
   │
   └─ Bersihkan semua file sementara (*temp.xlsx, Laporan SHELL*.xlsx, data_export*.xlsx, *.pdf) dari Dapur
[SELESAI]
```

### Pipeline Menu 2 (ringkas)

```
[START]
   │
   ├─ Validasi data_export*.xlsx & Laporan SHELL*.xlsx tersedia di Dapur
   │
   ├── A_JustXlookupData.py
   │     └─ Cocokkan Nomor Faktur Pajak langsung ke file Laporan SHELL*.xlsx
   │        (kolom Tanggal/DPP/No FP dideteksi otomatis posisinya, tidak harus di baris 1)
   │
   ├─ Pindahkan Laporan SHELL*.xlsx hasil proses kembali ke folder utama
   │
   └─ Bersihkan file sementara
[SELESAI]
```

---

## 📄 Penjelasan File di Folder Dapur

| File | Fungsi |
|---|---|
| `1_AmbilLampiranGmail.py` | Autentikasi Gmail API (OAuth), mencari email sesuai `gmail_query`, lalu mengunduh seluruh lampiran PDF yang sesuai filter tanggal & nama file |
| `2_EkstrakPdfShell.py` | Mengekstrak data terstruktur dari PDF faktur Shell (halaman 1: info faktur & nominal; halaman 2: referensi gudang, No SJ, No PO), lalu mengelompokkan ke sheet sesuai `gudang.conf` |
| `3_HelperDeleteDuplicate.py` | Menghapus baris yang identik persis (exact match seluruh kolom) di setiap sheet |
| `4_XlookupData.py` | Mencocokkan Nomor Faktur Pajak dari `data_export_*.xlsx` ke `Hasil_Ekstrak_temp.xlsx` menggunakan kunci (Tanggal, DPP) — pengganti fungsi `XLOOKUP` Excel dengan sistem antrean agar nilai duplikat tetap terpasang satu-satu secara berurutan |
| `5_KnifeToOperationFinalData.py` | Membersihkan & menstandarkan format file `Laporan SHELL*.xlsx` lama (riwayat) agar siap digabung — termasuk membuang baris dengan tanggal dummy `01/01/2001` |
| `6_CombineDeleteDuplicateAndSort.py` | Menggabungkan data lama dan data baru, menghapus duplikat antar keduanya, lalu mengurutkan seluruh data berdasarkan kolom Tanggal |
| `7_HelperDeleteDuplicateFinalData.py` | Deduplikasi final berbasis kombinasi 6 kolom kunci; bila ada duplikat, baris yang **sudah memiliki No FP** diprioritaskan untuk dipertahankan |
| `8_XlookupData.py` | Identik dengan script `4_`, dijalankan ulang setelah penggabungan & deduplikasi untuk memastikan baris baru tetap mendapat Nomor Faktur Pajak |
| `9_CopyDataToTemplate.py` | Menyalin `TEMPLATE.xlsx`, mengelompokkan data final per bulan, lalu menyisipkan/menyusun data ke template laporan rekap — termasuk membuat blok bulan baru lengkap dengan styling, border, dan rumus `SUM` otomatis jika belum ada |
| `10_HelperDeleteTemplateData.py` | Menghapus baris contoh "JANUARI 2001" yang merupakan placeholder bawaan template (jika template masih dalam kondisi awal/belum pernah dipakai) |
| `11_HelperMergedSUM.py` | Memperbaiki merge cell dan rumus `SUM` di antara setiap penanda blok "GUDANG" pada hasil akhir, memastikan border dan total per blok konsisten |
| `A_JustXlookupData.py` | Versi mandiri dari `4_`/`8_` untuk **Menu 2** — bekerja langsung pada file `Laporan SHELL*.xlsx` tanpa melalui seluruh pipeline ekstraksi PDF |
| `TEMPLATE.xlsx` | Template Excel laporan rekap gudang dengan struktur kolom: Gudang, Tanggal, No Inv, No SJ, No PO, Tgl FP, No FP, Byr, Klaim/Retur, DPP, Debet/Kredit, dll. |
| `CTX.xlsx` | Template fallback yang disalin sebagai dummy `data_export` apabila file export pajak sungguhan belum tersedia |
| `credentials.json` | Kredensial OAuth 2.0 untuk autentikasi ke Gmail API |
| `token.json` | Token akses & refresh token Gmail (dibuat/diperbarui otomatis setelah login pertama) |
| `gmail.conf` | Pengaturan query pencarian email & filter tanggal pengunduhan |
| `gudang.conf` | Pemetaan kode gudang ke nama sheet & filter tanggal ekstraksi PDF |

---

## 🧮 Logika Pencocokan Nomor Faktur Pajak

Pencocokan Nomor Faktur Pajak (kolom **No FP**) dilakukan dengan mencari kombinasi **Tanggal** dan **DPP** yang sama antara data hasil ekstraksi PDF dan data `data_export` (CTX), menggunakan pendekatan **sistem antrean (FIFO queue)**:

1. Dari `data_export`, setiap baris dengan kombinasi `(Tanggal Faktur Pajak, DPP)` yang sama dikumpulkan ke dalam antrean per kunci.
2. Untuk setiap baris di hasil ekstraksi PDF dengan kombinasi `(Tanggal, DPP)` yang sama, sistem mengambil **nilai paling depan dari antrean** tersebut dan menghapusnya dari antrean.
3. Pendekatan ini mencegah satu Nomor Faktur Pajak terpasang berulang pada beberapa baris berbeda yang kebetulan memiliki kombinasi Tanggal+DPP identik.

> **Catatan:** Akurasi pencocokan sangat bergantung pada keunikan kombinasi Tanggal+DPP. Jika dua transaksi berbeda memiliki Tanggal dan DPP yang persis sama, urutan pemasangan No FP mengikuti urutan baris pada data sumber (bukan dijamin 100% sesuai pasangan aslinya).

---

## 📤 File Output

| Skenario | Nama File Output | Lokasi |
|---|---|---|
| Menu 1 selesai | `Laporan SHELL BARU.xlsx` | Folder utama |
| Menu 2 selesai | `Laporan SHELL*.xlsx` (nama asli, diperbarui) | Folder utama (dipindah kembali dari Dapur) |

File output berisi satu sheet per kode gudang (sesuai `gudang.conf`), dengan data dikelompokkan per blok bulan (`JANUARI 2026`, `FEBRUARI 2026`, dst.), masing-masing dilengkapi baris total (`SUM`) otomatis pada kolom DPP.

---

## ⚠️ Penanganan Error

Script utama (`Proses Faktur SHELL.py`) akan menampilkan pesan dan menghentikan proses jika:

- Folder `Dapur` tidak ditemukan
- Salah satu file wajib di dalam folder `Dapur` hilang (lihat daftar `required_files` di script utama)
- Salah satu sub-script gagal dieksekusi (`subprocess.CalledProcessError`)
- Pada Menu 2: `data_export_*.xlsx` atau `Laporan SHELL*.xlsx` tidak ditemukan di folder `Dapur`

Setiap sub-script juga memiliki penanganan error masing-masing (file tidak ditemukan, kolom tidak lengkap, dsb.) dan akan mencetak pesan `-->` dengan keterangan, lalu melewati (skip) bagian yang gagal jika memungkinkan, atau menghentikan seluruh proses jika krusial.

---

## 🔒 Catatan Penting & Keamanan

1. **`credentials.json` dan `token.json` bersifat RAHASIA.** Jangan pernah mempublikasikan isi file ini ke repositori publik — keduanya memberi akses ke Gmail akun Anda. Disarankan menambahkan kedua file ini ke `.gitignore` dan hanya membagikan template/contoh kosong.
2. **Jangan rename atau pindahkan file di folder `Dapur`** — semua file dipanggil berdasarkan nama yang sudah ditentukan secara hardcoded di setiap script.
3. **`gmail_query`** menggunakan sintaks pencarian Gmail standar (sama seperti search bar Gmail), sehingga bisa disesuaikan untuk domain pengirim, kata kunci, atau rentang tanggal lain.
4. **Baris dengan tanggal `01/01/2001`** pada file `Laporan SHELL*.xlsx` lama akan otomatis dianggap data dummy/placeholder dan dibuang oleh `5_KnifeToOperationFinalData.py`.
5. Script tidak membuat backup otomatis untuk `Laporan SHELL*.xlsx` lama sebelum diproses. Disarankan menyimpan cadangan secara manual sebelum menjalankan Menu 1 jika data tersebut penting.
6. Folder `Dapur` dibersihkan otomatis dari seluruh file sementara (`*temp.xlsx`, `*.pdf`, `data_export*.xlsx`, `Laporan SHELL*.xlsx`) setiap kali proses selesai — pastikan hasil akhir sudah tersalin ke folder utama sebelum menjalankan proses berikutnya.

---

## 📜 Lisensi

Proyek ini dikembangkan untuk keperluan internal perpajakan dan internal perusahaan. Silakan sesuaikan dengan kebutuhan organisasi Anda.

---

*Dikembangkan oleh [ACC-TAX-REIGHTEEN](https://github.com/ACC-TAX-REIGHTEEN)*
