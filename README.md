# EWSSYNC

Source code VBA untuk workbook **EWSSYNC**. Workbook ini membantu tim pemeriksaan mengolah data EWS, data faktur pajak, dan kertas kerja pemeriksaan (KKP), lalu mengekspor hasilnya kembali ke template EWS.

## Fitur

| Fitur | Keterangan |
| --- | --- |
| Impor faktur pajak CSV | Mengimpor beberapa file CSV sekaligus, mengenali masa dari nama file, memeriksa pasangan file induk–detail, dan menyimpan data mentah beserta sumber filenya. Impor ulang dapat mengganti data pada masa yang sama. |
| Normalisasi faktur | Menyusun data menjadi `FakturPK` dan `FakturPM`, membentuk kunci faktur, serta menggabungkan rincian barang/jasa per nomor faktur. |
| Sinkronisasi faktur ke PPN | Memetakan faktur keluaran dan masukan ke sheet detail PPN serta memperbarui ringkasan/induk dan total per masa. |
| Impor workbook EWS | Membaca file EWS hingga 12 masa, mengenali NPWP dan periode dari file, memetakan nilai ke KKP lokal, dan mencatat riwayat impor. |
| Dukungan jenis pajak | Mendukung PPN Dalam Negeri, PPh Pasal 21, PPh Pasal 21 Final, PPh Pasal 23, PPh Pasal 26, serta PPh Pasal 4 ayat (2). |
| Validasi impor | Memeriksa identitas file, struktur sheet, periode, duplikasi impor, dan kecocokan identitas Wajib Pajak sebelum data diterapkan. |
| Reset data impor EWS | Menghapus nilai hasil impor dan riwayat impor tanpa menghapus formula, format, hyperlink, dropdown, atau header log. |
| Ekspor ke EWS | Memilih jenis pajak, kelompok KKP, KKP, masa, dan file EWS tujuan melalui form ekspor. File tujuan divalidasi sebelum penulisan. |
| Ekspor rincian faktur PPN | Menulis faktur untuk masa yang dipilih ke N-X-2, mengosongkan data lama, memberi nomor urut mulai dari 1, memasukkan detail barang bila tersedia, dan menyesuaikan merge cell sesuai pola header template EWS. |
| Ekspor batch | Menghitung kebutuhan baris terlebih dahulu, menyisipkan baris sekali, lalu menulis data per kolom agar ekspor banyak faktur lebih cepat dan tidak menimpa total. |
| Pencadangan dan rollback | Menyimpan nilai target sebelum ekspor sehingga perubahan pada file EWS dapat dibatalkan bila proses gagal sebelum file disimpan. |
| Log proses | Menyimpan catatan impor dan ekspor agar hasil, peringatan, dan kegagalan dapat ditelusuri. |

## Alur kerja yang dianjurkan

1. Buat salinan kerja workbook EWSSYNC dan template EWS.
2. Impor CSV faktur melalui `BukaFormImporFaktur`.
3. Jalankan `SinkronkanFakturKePPN` untuk membentuk detail serta ringkasan PPN.
4. Impor file EWS melalui `TampilkanFormImportEWS` apabila KKP perlu diisi dari file EWS.
5. Periksa hasil dan log impor.
6. Buka form ekspor melalui `BukaFormExportEWS`, pilih masa, KKP, dan file EWS tujuan.
7. Periksa total, rincian faktur, dan `Log Ekspor EWS` pada file hasil sebelum digunakan lebih lanjut.

## Cara menggunakan kode VBA

- `*.bas` adalah standard module VBA.
- `*.frm` adalah definisi UserForm.
- `*.frx` adalah resource binary untuk UserForm dengan nama yang sama.

Impor file melalui Visual Basic Editor (`Alt+F11`), lalu jalankan **Debug → Compile VBAProject** sebelum menyimpan workbook sebagai `.xlsm`. Pastikan macro diaktifkan saat workbook digunakan.

## Modul dan form

| Komponen | Titik masuk | Fungsi |
| --- | --- | --- |
| `modImportFakturPajak` | `BukaFormImporFaktur` | Impor dan normalisasi CSV faktur pajak. |
| `frmImportFaktur` | — | Memilih file CSV dan opsi penggantian data masa yang sama. |
| `modSinkronFakturKePPN` | `SinkronkanFakturKePPN` | Menulis faktur ke detail dan ringkasan PPN. |
| `modImportEWS` | `TampilkanFormImportEWS` | Impor nilai dari workbook EWS. |
| `modResetKKP` | `ResetDataImporKKP` | Menghapus data hasil impor EWS secara terkendali. |
| `modKatalogExportEWS` | `BukaFormExportEWS` | Menyiapkan pilihan pajak, KKP, masa, dan file pada form ekspor. |
| `frmExportEWS` | — | Antarmuka pemilihan parameter ekspor. |
| `modEngineExportEWS` | `EksporTerpilihEWS` | Validasi dan ekspor nilai KKP serta rincian PPN ke EWS. |
| `modSinkronisasiCFMWP_v054` | `TerapkanCFMWP054` | Memperbarui ringkasan CFM-WP pada file EWS. |
| `modPetaKKP` | — | Konstanta pajak, normalisasi kode, dan pemetaan sheet. |

Lihat [docs/MODULES.md](docs/MODULES.md) untuk daftar fungsi, dependensi, dan panduan validasi yang lebih rinci.

## Catatan penting

- Nomor NPWP, TIN, dan nomor faktur diperlakukan sebagai teks agar angka panjang serta nol di depan tidak berubah.
- Jalankan ekspor pada salinan template EWS, terutama saat menguji perubahan kode.
- Uji satu masa pajak terlebih dahulu sebelum mengekspor seluruh masa.
- Saat proses besar berjalan, Excel dapat memakai kalkulasi manual agar lebih cepat. Jika muncul pemberitahuan **stale values**, tekan `F9` atau **Calculate Now** sebelum menilai hasil.
- Jangan mengimpor dua modul dengan `Attribute VB_Name` yang sama.

## Pengembangan

- Simpan setiap perubahan VBA sebagai `.bas` atau `.frm`; jangan menjadikan file `.xlsm` sebagai satu-satunya sumber kode.
- Setelah perubahan, jalankan **Debug → Compile VBAProject** dan uji pada salinan workbook.
- Untuk perubahan ekspor PPN, verifikasi jumlah baris, nomor urut, kolom faktur, detail barang, total DPP/PPN, dan merge cell pada sheet EWS tujuan.
