# Modul VBA Excel: Konversi Jumlah IDR ke Teks (Terbilang) – Proyek Magang CANDI TOUR N TRAVEL

## ID (Bahasa Indonesia)
Selama magang di **CANDI TOUR N TRAVEL** (periode: Juni - Agustus 2025), saya mengembangkan modul VBA untuk Microsoft Excel yang mengonversi nilai mata uang Rupiah (IDR) dari format numerik/string menjadi representasi kata-kata (terbilang) dalam bahasa Indonesia dan Inggris. Proyek ini digunakan untuk otomatisasi dokumen keuangan seperti invoice paket tour, kwitansi pembayaran, dan laporan biaya travel, memastikan akurasi dan kepatuhan standar formal (misalnya, menghindari pemalsuan jumlah).

Teknologi: **VBA (Visual Basic for Applications)** di Microsoft Excel.  
Repo ini berisi kode sumber dan contoh penggunaan.

### Fitur Utama
- **Parsing IDR Fleksibel**: Fungsi `ParseIDR` menormalkan input seperti "Rp 1.234.567,89", "IDR1,234.56", atau dari PDF (menangani titik ribuan, koma desimal, spasi).
- **Terbilang Bahasa Indonesia (`TerbilangIDR`)**: Contoh: 1.234.567,89 → "satu juta dua ratus tiga puluh empat ribu lima ratus enam puluh tujuh rupiah dan delapan puluh sembilan sen". Dukung negatif, fraksi (sen), hingga triliun.
- **Terbilang Bahasa Inggris (`AmountInWordsIDR_EN`)**: Contoh: 1.234.567,89 → "one million two hundred thirty-four thousand five hundred sixty-seven rupiah and eighty-nine cents".
- **Fungsi Pendukung**: `Div1000` & `Mod1000` untuk handling Currency aman; rekursi untuk angka besar.
- **Error Handling**: Fallback ke 0 jika input invalid.

### Cara Menggunakan
1. Buka Excel > Alt + F11 (VBA Editor) > Insert > Module.
2. Paste kode dari `IDR_Terbilang_Module.bas`.
3. Di sheet Excel, gunakan formula:
   - `=TerbilangIDR(A1)` untuk sel A1 berisi 1234567.89.
   - `=AmountInWordsIDR_EN("Rp 1.234,56")`.
4. Tes dengan file demo jika ada.

### Contoh Output
| Input | Output (ID) | Output (EN) |
|-------|-------------|-------------|
| 1234.56 | seribu dua ratus tiga puluh empat rupiah dan lima puluh enam sen | one thousand two hundred thirty-four rupiah and fifty-six cents |
| -5000 | minus lima ribu rupiah | minus five thousand rupiah |

### Dampak di CANDI TOUR N TRAVEL
- Menghemat waktu manual penulisan terbilang di 100+ invoice bulanan.
- Mengurangi error manusia di dokumen keuangan travel (misalnya, booking tour domestik/internasional).
- Dapat diintegrasikan ke template Excel perusahaan untuk skalabilitas.

### Teknologi & Skills
- VBA Macros untuk automasi.
- Handling string parsing & rekursi numerik.
- Testing: Diverifikasi dengan angka ekstrem (hingga 1E+15).

## EN (English)
During my internship at **CANDI TOUR N TRAVEL** (period: June - August 2025), I developed a VBA module for Microsoft Excel that converts Indonesian Rupiah (IDR) currency values from numeric/string formats into worded representations (terbilang) in both Indonesian and English. This project automates financial documents such as tour package invoices, payment receipts, and travel expense reports, ensuring accuracy and compliance with formal standards (e.g., preventing amount forgery).

Technology: **VBA (Visual Basic for Applications)** in Microsoft Excel.  
This repo contains the source code and usage examples.

### Key Features
- **Flexible IDR Parsing**: Function `ParseIDR` normalizes inputs like "Rp 1.234.567,89", "IDR1,234.56", or from PDFs (handles thousand dots, decimal commas, spaces).
- **Indonesian Spelling (`TerbilangIDR`)**: Example: 1,234,567.89 → "satu juta dua ratus tiga puluh empat ribu lima ratus enam puluh tujuh rupiah dan delapan puluh sembilan sen". Supports negatives, fractions (cents), up to trillions.
- **English Spelling (`AmountInWordsIDR_EN`)**: Example: 1,234,567.89 → "one million two hundred thirty-four thousand five hundred sixty-seven rupiah and eighty-nine cents".
- **Supporting Functions**: `Div1000` & `Mod1000` for safe Currency handling; recursion for large numbers.
- **Error Handling**: Fallback to 0 for invalid inputs.

### How to Use
1. Open Excel > Alt + F11 (VBA Editor) > Insert > Module.
2. Paste code from `IDR_Terbilang_Module.bas`.
3. In Excel sheet, use formulas:
   - `=TerbilangIDR(A1)` for cell A1 containing 1234567.89.
   - `=AmountInWordsIDR_EN("Rp 1.234,56")`.
4. Test with demo file if available.

### Example Outputs
| Input | Output (ID) | Output (EN) |
|-------|-------------|-------------|
| 1234.56 | seribu dua ratus tiga puluh empat rupiah dan lima puluh enam sen | one thousand two hundred thirty-four rupiah and fifty-six cents |
| -5000 | minus lima ribu rupiah | minus five thousand rupiah |

### Impact at CANDI TOUR N TRAVEL
- Saved time on manual spelling for 100+ monthly invoices.
- Reduced human errors in travel financial documents (e.g., domestic/international tour bookings).
- Integratable into company Excel templates for scalability.

### Technologies & Skills
- VBA Macros for automation.
- String parsing & numeric recursion handling.
- Testing: Verified with extreme numbers (up to 1E+15).

## License
MIT License – bebas digunakan/modifikasi (free to use/modify).
