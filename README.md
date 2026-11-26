# 📊 Analisis Data Warehouse & Business Intelligence — Retail & Bank

Repository ini berisi hasil pengerjaan tugas mata kuliah **Data Warehouse** berupa implementasi *End-to-End Business Intelligence* menggunakan **Microsoft Excel (Power Query, Power Pivot, DAX, PivotChart, dan Slicer)** untuk dua studi kasus: **Penjualan Retail Indonesia** dan **Retensi & Churn Nasabah Kartu Kredit Bank**.

---

## 👤 Identitas Penyusun

| Keterangan | Detail |
| :--- | :--- |
| **Nama** | Fatimah Azzahra |
| **NIM** | E42241976 |
| **Program Studi** | Bisnis Digital |
| **Jurusan** | Bisnis |
| **Institusi** | Politeknik Negeri Jember |
| **Mata Kuliah** | Data Warehouse |
| **Dosen Pengampu** | Eka Yuniar, S.Kom., MMSI. |
| **Tahun** | 2026 |

---

## 📁 Daftar Isi Repository

| No | Nama File | Jenis | Deskripsi |
| :-: | :--- | :---: | :--- |
| 1 | [`Dashboard_Retail_Zahrunyy.xlsx`](./Dashboard_Retail_Zahrunyy.xlsx) | Excel (BI) | Dashboard interaktif Penjualan Retail Indonesia berbasis Star Schema & DAX |
| 2 | [`Dashboard_Bank_Zahrunyy.xlsx`](./Dashboard_Bank_Zahrunyy.xlsx) | Excel (BI) | Dashboard interaktif Analisis Retensi & Churn Nasabah Kartu Kredit Bank |
| 3 | [`Laporan Tugas Data Warehouse - Fatimah Azzahra (E42241976).pdf`](./Laporan%20Tugas%20Data%20Warehouse%20-%20Fatimah%20Azzahra%20(E42241976).pdf) | Laporan PDF | Laporan lengkap analisis ETL, Data Model, DAX, visualisasi, dan rekomendasi bisnis |
| 4 | [`Video Penjelasan Studi Kasus Penjualan Retail.mp4`](./Video%20Penjelasan%20Studi%20Kasus%20Penjualan%20Retail.mp4) | Video MP4 | Video presentasi dan demonstrasi dashboard interaktif |

---

## 🛒 Studi Kasus I: Analisis Penjualan Retail Indonesia

* **Sumber Data Mentah:** `data_retail2.xlsx` (541.909 baris transaksi, periode 1 Des 2020 – 9 Des 2021)
* **Data Bersih Setelah ETL:** **531.282 baris valid** (eliminasi 10.624 transaksi retur `Quantity <= 0`, pembersihan 3 transaksi penyesuaian non-penjualan, serta normalisasi *case-insensitivity* pada 3.941 kode produk unik).
* **Arsitektur Star Schema (Power Pivot):**
  * **Tabel Fakta:** `Fakta_Penjualan` (531.282 baris)
  * **Tabel Dimensi:** `Dim_Cabang` (34 cabang), `Dim_Channel` (42 saluran distribusi), `Dim_Produk` (3.941 produk), dan `Dim_Tanggal` (396 hari, ditandai sebagai *Date Table* dengan pengurutan bulan berbasis `MonthNo`).

### Ringkasan Metrik Utama (DAX Measures)
| Measure | Rumus DAX | Nilai Aktual |
| :--- | :--- | :---: |
| **Total Omzet** | `SUM(Fakta_Penjualan[Total_Omzet])` | **Rp 152.375.401.520** |
| **Total Qty** | `SUM(Fakta_Penjualan[Quantity])` | **5.660.978 unit** |
| **Total Transaksi** | `DISTINCTCOUNT(Fakta_Penjualan[InvoiceNo])` | **20.724 invoice** |
| **Avg Qty per Transaksi** | `DIVIDE([Total Qty], [Total Transaksi], 0)` | **273,16 unit/invoice** |
| **Rata-rata Harga** | `AVERAGE(Fakta_Penjualan[UnitPriceRupiah])` | **Rp 55.457,39** |

### Visualisasi & Temuan Kunci
1. **Tren Omzet & Rata-Rata Qty per Kuartal *(Combo Chart)*:** Penjualan memuncak pada **Kuartal 4 (Rp 59,02 Miliar)** dengan rekor bulanan tertinggi pada **November 2021 (Rp 21,59 Miliar)**.
2. **Grosir vs Eceran pada 8 Cabang Terbesar *(Stacked Horizontal Bar)*:** Segmen **Grosir** (≥ 10 unit) mendominasi **63,48%** omzet (Rp 96,72 Miliar), dipimpin Cabang **Kediri** (Rp 16,74 Miliar) dan **Kudus** (Rp 15,99 Miliar).
3. **Top 10 Produk Kontributor Omzet Tertinggi *(Horizontal Bar)*:** Dipimpin oleh `BASKOM WINSTON 9021 HW 175` (Rp 2,95 Miliar), `APUR CAR AQUA ST` (Rp 2,50 Miliar), dan `TOKYO 828 CF9 SET ORANGE` (Rp 2,41 Miliar).
4. **Pola Penjualan Weekday vs Weekend *(Doughnut Chart)*:** **75,66%** omzet terjadi pada hari kerja (puncak Rabu Rp 31,51 Miliar & Senin Rp 31,15 Miliar), sedangkan akhir pekan **24,34%** (Sabtu terendah Rp 11,64 Miliar).
5. **Top 10 Kategori Harga Rata-Rata Tertinggi *(Column Chart)*:** Kategori `PECAH BELAH` (Rp 94.169/unit) dan `PEMBERSIH` (Rp 89.147/unit) menempati nilai rata-rata tertinggi, sementara **Toko Kelontong** menjadi saluran distribusi utama (**66,42%** total omzet).

---

## 🏦 Studi Kasus II: Analisis Retensi & Churn Nasabah Bank

* **Sumber Data Mentah:** `bank1.xls` — sheet `bnk` (**10.127 nasabah kartu kredit**, 21 variabel asli)
* **Data Preparation & Feature Engineering:** Konversi `CLIENTNUM` ke tipe *Text*, pembuatan flag biner `Is_Churn`, serta kolom segmentasi tingkat pemakaian kartu (`Util_Tier`) dan tingkat kevakuman transaksi (`Inactive_Tier`).

### Ringkasan Metrik Utama (DAX Measures)
| Measure | Rumus DAX | Nilai Aktual |
| :--- | :--- | :---: |
| **Total Nasabah** | `COUNTROWS(Bank_Nasabah)` | **10.127 orang** |
| **Nasabah Churn** | `CALCULATE(COUNTROWS(Bank_Nasabah), Bank_Nasabah[Attrition_Flag] = "Attrited Customer")` | **1.627 orang** |
| **Nasabah Aktif** | `CALCULATE(COUNTROWS(Bank_Nasabah), Bank_Nasabah[Attrition_Flag] = "Existing Customer")` | **8.500 orang** |
| **Rasio Churn** | `DIVIDE([Nasabah Churn], [Total Nasabah], 0)` | **16,07%** |
| **Rata-rata Kontak CS** | `AVERAGE(Bank_Nasabah[Contacts_Count_12_mon])` | **2,46 kali** |
| **Rata-rata Utilisasi** | `AVERAGE(Bank_Nasabah[Avg_Utilization_Ratio])` | **27,49%** |

### Visualisasi & Temuan Kunci
1. **Intensitas Kontak CS vs Rasio Churn *(Combo Chart)*:** Risiko churn melonjak tajam seiring frekuensi komplain — dari **1,75%** (0 kontak), **20,15%** (3 kontak), **33,52%** (5 kontak), hingga mencapai **100,00% churn pada 6 kali kontak CS**.
2. **Tingkat Pemakaian Kartu vs Rasio Churn *(Column Chart)*:** Nasabah kartu tidur (**0% Utilisasi / Idle**) mencatat churn tertinggi sebesar **36,15%**, berbanding jauh dengan pengguna aktif 51–80% yang hanya **6,40%**.
3. **Pengaruh Jumlah Produk Dimiliki *(100% Stacked Bar)*:** Pemilik 1–2 produk memiliki churn tinggi (**25,60%–27,84%**), sedangkan kepemilikan 4–6 produk menekan churn menjadi **10,50%–12,00%** (*product stickiness*).
4. **Tren Risiko Berdasarkan Lama Tidak Aktif *(Line Chart with Markers)*:** Kevakuman transaksi bulan ke-3 hingga ke-4 menjadi *tipping point* utama dengan lonjakan churn mencapai **21,48%–29,89%**.
5. **Matriks Risiko Kategori Kartu vs Tingkat Pendidikan *(PivotTable Heatmap)*:** Pemegang kartu **Platinum** memiliki rasio churn tertinggi (**25,00%**), dengan titik risiko paling tinggi pada irisan *Platinum × Doctorate* (**50,00%**) dan *Platinum × Graduate* (**37,50%**).

---

## 🛠️ Teknologi & Fitur yang Digunakan

* **Microsoft Excel (English UI)**
* **Power Query** — Pembersihan data, konversi tipe data, dan pembentukan tabel fakta & dimensi
* **Power Pivot & Data Model** — Relasi Star Schema (*One-to-Many*), *Mark as Date Table*, dan *Sort by Column*
* **DAX (Data Analysis Expressions)** — *Explicit Measures* dinamis (`SUM`, `DISTINCTCOUNT`, `CALCULATE`, `DIVIDE`, `AVERAGE`)
* **Visualisasi Interaktif** — Combo Chart, Stacked Bar Chart, Horizontal Bar Chart, Doughnut Chart, Column Chart, Line Chart, PivotTable Heatmap (*Conditional Formatting Color Scales*), KPI Cards, serta 4 Slicer yang terhubung melalui *Report Connections*

---

## 💡 Catatan Penggunaan File Excel

1. Unduh file `.xlsx` terlebih dahulu dan buka menggunakan **Microsoft Excel Desktop** (versi 2016 ke atas / Microsoft 365) agar fitur **Data Model**, **Power Pivot**, dan **Slicer** berjalan optimal.
2. Jika muncul bilah kuning *Security Warning* saat pertama kali membuka file, klik **Enable Content**.
3. Gunakan tombol **Slicer** di halaman `Dashboard` untuk memfilter seluruh metrik dan grafik secara interaktif.
