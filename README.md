# Sistem Otomatisasi Instrumen Penjaminan Mutu

Repository ini mengelola **Form Penilaian Mandiri Penyelenggara**, **Google Sheet Tim Penilai**, dan **Google Sheet Progress Penyelenggara** menggunakan Google Apps Script.

Sistem berjalan dengan satu Google Form utama. Setiap submit atau edit response akan memproses data ke beberapa keluaran:

- tab `Rekap` sebagai ringkasan satu baris per pelatihan;
- tab `Detail Jawaban` sebagai data long-table per indikator;
- tab khusus berdasarkan `Nama Pelatihan` untuk kerja tim penilai;
- Google Sheets khusus penyelenggara untuk memantau status penilaian indikator;
- email otomatis ke penyelenggara berisi link langsung ke tab `Progress Indikator`.

## Struktur Berkas

- `Kode.js`: kode utama Google Apps Script.
- `appsscript.json`: manifest Apps Script.
- `panduan_instrumen.md`: panduan lengkap operasional sistem.
- `scripts/google-form-browser-test.js`: runner Playwright untuk testing Google Form.
- `test-data/google-form-answers.form-data-1uHzZr.json`: konfigurasi dummy lengkap dari struktur form aktif.
- `build_drive_workbook.py`: pembantu pembuatan workbook lokal.

## Alur Operasional

1. Penyelenggara mengisi atau mengedit Google Form Penilaian Mandiri.
2. Trigger `handleMandiriSubmit` memproses response.
3. Sistem memperbarui `Rekap`, `Detail Jawaban`, dan tab nama pelatihan.
4. Sistem membuat atau memperbarui spreadsheet khusus penyelenggara.
5. Link spreadsheet penyelenggara dikirim via email dan langsung membuka tab `Progress Indikator`.
6. Tim penilai mengisi `Nilai Tim Penilai`, `Catatan Tim Penilai`, dan dropdown `Status Penilaian` di tab nama pelatihan.
7. Perubahan status penilaian tersinkron ke spreadsheet penyelenggara.

## Struktur Sheet Utama

### `Rekap`

Tab ringkasan lintas pelatihan. Satu pelatihan = satu baris.

Kolom utama:

- `ID Pelatihan`
- `Nama Pelatihan`
- `Nama Pelatihan Normalisasi`
- `Metode`
- `Tanggal Mulai`
- `Tanggal Selesai`
- `Tempat`
- `Unit Penyelenggara`
- `Tanggal Penjaminan Mutu`
- `Status Validasi`
- `Status Penilaian`
- `Jumlah Indikator`
- `Jumlah Terisi`
- `Jumlah Tidak Terisi`
- `Jumlah Tidak Berlaku`
- `Jumlah Bukti Link`
- `Jumlah Bukti File`
- `Jumlah Catatan`
- `Terakhir Diubah`
- `Nama Tab Detail`

### `Detail Jawaban`

Tab long-table untuk audit dan filter detail indikator. Satu baris = satu indikator pada satu pelatihan.

### Tab Nama Pelatihan

Tab kerja tim penilai. Kolom `Status Penilaian` menggunakan dropdown:

- `Belum Dinilai`
- `Sudah Sesuai`
- `Perlu Tindak Lanjut`

Warna status:

- kuning: `Belum Dinilai`
- hijau: `Sudah Sesuai`
- orange: `Perlu Tindak Lanjut`

### Sheet Progress Penyelenggara

Dibuat otomatis per pelatihan. Terdiri dari:

- `Ringkasan`
- `Progress Indikator`

Penyelenggara menerima link via email yang langsung membuka tab `Progress Indikator`.

## Fungsi Apps Script Penting

- `setupInstrumenPenjaminanMutu`: membuat struktur awal form dan sheet.
- `handleMandiriSubmit`: trigger utama saat form disubmit atau response diedit.
- `resyncMandiriResponsesToSummaryModel`: memproses ulang semua response lama.
- `migrateExistingRekapToSummaryModel`: membangun ulang `Rekap`, `Detail Jawaban`, dan `Dashboard`.
- `updateExistingRekapColumnsForSimpleModel`: menyederhanakan kolom `Rekap` lama.
- `installPenilaianEditTriggerForExistingSpreadsheet`: memasang trigger edit status penilaian pada spreadsheet lama.
- `applyStatusPenilaianDropdownsToExistingSheets`: menerapkan dropdown dan warna status pada tab penilaian yang sudah ada.

## Push Kode ke Cloud

```powershell
npx clasp push
```

Untuk membuka editor Apps Script:

```powershell
npx clasp open-script
```

## Testing Google Form dengan Browser Automation

Dry-run, hanya mengisi form dan screenshot:

```powershell
npm run test:form -- --config "test-data/google-form-answers.form-data-1uHzZr.json"
```

Submit response dummy:

```powershell
npm run test:form:submit -- --config "test-data/google-form-answers.form-data-1uHzZr.json" --count 10
```

Catatan: jika form masih meminta login Google, browser automation tidak bisa login otomatis secara stabil karena pembatasan keamanan Google. Untuk testing tanpa login, matikan sementara collect email terverifikasi dan hindari item file upload.
