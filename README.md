# EWSSYNC

Source code VBA untuk workbook EWSSYNC, aplikasi Excel yang membantu impor data EWS dan faktur pajak, sinkronisasi ke kertas kerja PPN, serta ekspor kembali ke template EWS.

## Struktur proyek

- `*.bas` adalah standard module VBA.
- `*.frm` adalah definisi UserForm.
- `*.frx` adalah resource binary yang dipasangkan dengan UserForm bernama sama.

Impor file melalui Visual Basic Editor (`Alt+F11`), kemudian pilih **Debug → Compile VBAProject** sebelum menyimpan workbook sebagai `.xlsm`.

## Alur kerja utama

1. Impor CSV faktur melalui `frmImportFaktur`.
2. Jalankan `SinkronkanFakturKePPN`.
3. Impor file EWS melalui `TampilkanFormImportEWS` bila diperlukan.
4. Ekspor data melalui `BukaFormExportEWS`.
5. Periksa sheet log workbook sebelum memakai hasil ekspor.

## Modul penting

Lihat [docs/MODULES.md](docs/MODULES.md) untuk fungsi setiap modul, form, serta dependensi penting.

## Aturan pengembangan

- Simpan source VBA sebagai file `.bas` dan `.frm`; jangan menjadikan binary workbook satu-satunya sumber perubahan.
- Jangan impor dua modul yang memiliki `Attribute VB_Name` sama.
- Uji perubahan pada salinan workbook dan template EWS.
- Untuk ekspor rincian PPN, uji satu masa pajak sebelum memproses seluruh masa.
- TIN/NPWP dan nomor faktur harus diperlakukan sebagai teks agar nol di depan dan angka panjang tetap utuh.

## Catatan ekspor PPN

`modEngineExportEWS_v055_DETAILFIX.bas` berisi perbaikan ekspor rinci PPN. Prosesnya menghitung baris per masa, menyisipkan kapasitas sekali, lalu menulis data per kolom secara batch. Pendekatan ini mengurangi operasi per-sel yang sebelumnya membuat ekspor lambat.
