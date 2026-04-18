# BCA Mutasi Rekening PDF to Excel

Script Python ini digunakan untuk mengekstrak data mutasi rekening BCA dari file PDF, membersihkan data transaksi, memisahkan nilai debit dan kredit, lalu mengekspornya ke file Excel.

## Fitur

- Membaca file PDF mutasi rekening BCA dari **semua halaman**
- Menggabungkan tabel dari beberapa halaman
- Membersihkan header berulang, nomor halaman, dan baris sampah
- Menggabungkan baris transaksi yang terpotong ke beberapa baris
- Memisahkan data **summary** seperti:
  - Saldo Awal
  - Mutasi CR
  - Mutasi DB
  - Saldo Akhir
- Memisahkan kolom **Mutasi** menjadi:
  - **DB** untuk debit
  - **CR** untuk kredit
- Mengekspor hasil ke file Excel dengan 2 sheet:
  - `Transaksi`
  - `Summary`

---

## Struktur Folder

Disarankan menggunakan struktur folder seperti berikut:

```bash
project_folder/
│
├── pdf_file/
│   └── 6000150426_SEP_2025.pdf
│
├── excel_file/
│   └── 6000150426_SEP_2025.xlsx
│
├── script.py
└── README.md
