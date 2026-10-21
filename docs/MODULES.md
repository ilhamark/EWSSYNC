# Katalog modul EWSSYNC

## Modul VBA

| File | Nama modul VBA | Peran | Titik masuk utama |
| --- | --- | --- | --- |
| `modEngineExportEWS_v055_DETAILFIX.bas` | `modEngineExportEWS` | Mesin ekspor data KKP ke template EWS, termasuk rincian faktur PPN dan log ekspor. | `EksporTerpilihEWS` |
| `modKatalogExportEWS_v0.54.bas` | `modKatalogExportEWS` | Mengisi form ekspor: jenis pajak, kelompok KKP, file target, dan validasi pilihan. | `BukaFormExportEWS` |
| `modPetaKKP_v0.5_PPh21F.bas` | `modPetaKKP` | Konstanta jenis pajak, normalisasi kode EWS, dan pemetaan kode ke sheet lokal. | `NormalisasiKodeKKPEWS`, `DapatkanSheetLokalKKP` |
| `modImportEWS_v0.54.bas` | `modImportEWS` | Impor file EWS untuk kelompok pajak dan pencatatan log impor. | `TampilkanFormImportEWS` |
| `modImportFakturPajak_v.054.bas` | `modImportFakturPajak` | Impor CSV faktur, normalisasi data mentah, dan pembaruan `FakturPM`/`FakturPK`. | `BukaFormImporFaktur` |
| `modSinkronFakturKePPN_v0.55.bas` | `modSinkronFakturKePPN` | Memindahkan data faktur hasil normalisasi ke lembar detail dan induk PPN. | `SinkronkanFakturKePPN` |
| `modSinkronisasiCFMWP_v054.bas` | `modSinkronisasiCFMWP_v054` | Memperbarui ringkasan CFM-WP pada workbook EWS. | `TerapkanCFMWP054` |
| `modResetKKP_v0.5.bas` | `modResetKKP` | Menghapus data impor KKP secara terkontrol berdasarkan struktur workbook. | `ResetDataImporKKP` |
| `modImportPPh26_v0.54.bas` | `modImportPPh26Alternatif` | Jalur alternatif impor PPh 26; berinteraksi dengan Excel eksternal dan Windows API. | `ImportPPh26AlternatifDariDaftarForm` |

## UserForm

| File | Nama form | Peran |
| --- | --- | --- |
| `frmExportEWS_v0.54.frm` + `frmExportEWS.frx` | `frmExportEWS` | Memilih jenis pajak, KKP, masa, dan file EWS tujuan untuk ekspor. |
| `frmImportFaktur_v.054.frm` + `frmImportFaktur.frx` | `frmImportFaktur` | Memilih dan mengimpor file CSV faktur. |

## Dependensi penting

`modEngineExportEWS` bergantung pada `modPetaKKP` untuk konstanta jenis pajak dan resolusi nama sheet. Form ekspor bergantung pada `modKatalogExportEWS` dan `modEngineExportEWS`. Jangan menghapus `modPetaKKP` saat mengganti modul ekspor.

## Validasi sebelum rilis

1. Impor seluruh modul/form ke salinan workbook.
2. Jalankan **Debug → Compile VBAProject**.
3. Uji impor faktur satu masa.
4. Uji sinkronisasi PPN.
5. Uji ekspor satu masa ke salinan template EWS.
6. Periksa log dan total DPP/PPN sebelum memproses data penuh.
