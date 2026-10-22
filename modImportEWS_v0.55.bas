Attribute VB_Name = "modImportEWS"
'-------------------------------------------------------------
'THIS MODULE IS FUELED BY SPITE OF EWS AND CORETAX
'@2026
'-------------------------------------------------------------
Option Explicit

'EWSSYNC v0.5 - engine impor dan validasi transaksi.

Private Const NAMA_SHEET_LOG As String = "Log Impor EWS"

Private Const KOLOM_NAMA_EWS_1 As Long = 8
Private Const KOLOM_NAMA_EWS_2 As Long = 9
Private Const KOLOM_NAMA_EWS_3 As Long = 10
Private Const KOLOM_REF_EWS As Long = 25
Private Const KOLOM_JENIS_NILAI_EWS As Long = 27
Private Const KOLOM_NILAI_EWS As Long = 31

Private Const KOLOM_NAMA_KKP_AWAL As Long = 1
Private Const KOLOM_NAMA_KKP_AKHIR As Long = 4
Private Const KOLOM_REF_KKP As Long = 6
Private Const KOLOM_JENIS_NILAI_KKP As Long = 7

Private Const KOLOM_NAMA_PPH21_EWS_AWAL As Long = 8
Private Const KOLOM_NAMA_PPH21_EWS_AKHIR As Long = 11
Private Const KOLOM_ID_PPH21_EWS As Long = 1
Private Const KOLOM_NAMA_PPH21_KKP_AWAL As Long = 1
Private Const KOLOM_NAMA_PPH21_KKP_AKHIR As Long = 5

Private Const BATAS_BARIS_HEADER_PERIODE As Long = 250
Private Const BATAS_TUNGGU_REBUILD_DETIK As Double = 300
Private Const BATAS_TUNGGU_SHEET_DETIK As Double = 30
Private Const BATAS_ERROR_SUMBER_DETIK As Double = 60
Private Const MINIMUM_TUNGGU_REBUILD_DETIK As Double = 2
Private Const DURASI_STABIL_REBUILD_DETIK As Double = 2

Private Const LST_JENIS As Long = 0
Private Const LST_PERIODE As Long = 1
Private Const LST_FILE As Long = 2
Private Const LST_STATUS As Long = 3
Private Const LST_PATH As Long = 4
Private Const LST_KODE As Long = 5
Private Const LST_NPWP As Long = 6
Private Const LST_BULAN_AWAL As Long = 7
Private Const LST_BULAN_AKHIR As Long = 8
Private Const LST_TAHUN As Long = 9

Private Const STATUS_SIAP As String = "Siap"
Private Const STATUS_VERSI_BARU As String = "Versi baru"
Private Const STATUS_SUDAH_DIIMPOR As String = "Sudah diimpor"
Private Const STATUS_BELUM_DIDUKUNG As String = "Belum didukung"
Private Const HANDLER_PPN_DN As String = "PPN_DN"
Private Const HANDLER_PPH21 As String = "PPH21"
Private Const HANDLER_PPH21F As String = "PPH21F"
Private Const HANDLER_PPH23 As String = "PPH23"
Private Const HANDLER_PPH26 As String = "PPH26"
Private Const HANDLER_PPH42 As String = "PPH42"

Private mFolderTrustedDikonfirmasi As Object

'Bagian: antarmuka impor.
'Tujuan: Menampilkan form import ews.
Public Sub TampilkanFormImportEWS()

    frmImportEWS.Show

End Sub

'Tujuan: Menyiapkan daftar masa.
Public Sub SiapkanDaftarMasa(ByVal frm As Object)

    With frm.lstEWS

        .Clear
        .ColumnCount = 10
        .ColumnWidths = _
            "150 pt;95 pt;235 pt;85 pt;" & _
            "0 pt;0 pt;0 pt;0 pt;0 pt;0 pt"
        .MultiSelect = fmMultiSelectSingle
        .IntegralHeight = False

    End With

    frm.npwpReferensi = vbNullString
    frm.namaWPReferensi = vbNullString
    frm.TahunReferensi = 0

    frm.lblWPValue.Caption = "-"
    frm.lblTahunValue.Caption = "-"
    frm.lblJumlah.Caption = "0 file siap"
    frm.lblStatus.Caption = _
        "Pilih satu atau beberapa file EWS."

    frm.cmdImport.Enabled = False

End Sub

'Tujuan: Memilih file ews 12 masa.
Public Sub PilihFileEWS12Masa(ByVal frm As Object)

    Dim dlg As Object
    Dim i As Long
    Dim pathFile As String

    Set dlg = Application.FileDialog(msoFileDialogFilePicker)

    With dlg

        .Title = "Pilih File EWS"
        .AllowMultiSelect = True

        .Filters.Clear
        .Filters.Add _
            "File Excel", _
            "*.xls;*.xlsx;*.xlsm;*.xlsb"

        If .Show <> -1 Then Exit Sub

        For i = 1 To .SelectedItems.Count

            pathFile = CStr(.SelectedItems(i))
            TambahFileKeForm frm, pathFile

        Next i

    End With

    UpdateTampilanForm frm

End Sub

'Bagian: daftar file EWS.
'Tujuan: Menambahkan file ke form.
Private Sub TambahFileKeForm( _
    ByVal frm As Object, _
    ByVal pathFile As String)

    Dim kodeKKP As String
    Dim namaKKP As String
    Dim npwpFile As String
    Dim kodeMasa As String
    Dim identifierFile As String
    Dim handler As String
    Dim pesanError As String

    Dim bulanAwal As Long
    Dim bulanAkhir As Long
    Dim tahunFile As Long
    Dim dikenal As Boolean
    Dim didukung As Boolean
    Dim indeks As Long
    Dim statusFile As String

    If Not BacaIdentitasDariNamaFile( _
            pathFile, _
            kodeKKP, _
            namaKKP, _
            npwpFile, _
            kodeMasa, _
            bulanAwal, _
            bulanAkhir, _
            tahunFile, _
            identifierFile, _
            dikenal, _
            didukung, _
            handler, _
            pesanError) Then

        TambahBarisDaftar _
            frm, _
            "Tidak dikenali", _
            "-", _
            Dir$(pathFile), _
            "Nama tidak valid", _
            pathFile, _
            vbNullString, _
            vbNullString, _
            0, _
            0, _
            0, _
            vbNullString, _
            vbNullString, _
            pesanError

        Exit Sub
    End If

    If Len(frm.npwpReferensi) = 0 Then

        frm.npwpReferensi = npwpFile
        frm.TahunReferensi = tahunFile

    ElseIf frm.npwpReferensi <> npwpFile Then

        statusFile = "NPWP berbeda"
        pesanError = _
            "NPWP file berbeda dengan NPWP referensi batch."

    End If

    If Len(statusFile) = 0 Then

        If Not dikenal Then

            statusFile = "Kode tidak dikenal"
            pesanError = _
                "Kode KKP belum tercantum dalam katalog EWSSYNC."

        ElseIf Not didukung Then

            statusFile = STATUS_BELUM_DIDUKUNG
            pesanError = _
                "Jenis KKP dikenali, tetapi mapping impor " & _
                "belum tersedia."

        Else

            statusFile = STATUS_SIAP

            TerapkanStatusRiwayat _
                pathFile, kodeKKP, npwpFile, _
                bulanAwal, bulanAkhir, tahunFile, _
                statusFile, pesanError

        End If

    End If

    indeks = CariIdentifierPadaForm(frm, identifierFile)

    If indeks >= 0 Then

        IsiBarisDaftar _
            frm, _
            indeks, _
            namaKKP, _
            NamaPeriode(bulanAwal, bulanAkhir, tahunFile), _
            NamaFileSingkat(pathFile, namaKKP, kodeMasa), _
            statusFile, _
            pathFile, _
            kodeKKP, _
            npwpFile, _
            bulanAwal, _
            bulanAkhir, _
            tahunFile, _
            identifierFile, _
            handler, _
            pesanError

    Else

        TambahBarisDaftar _
            frm, _
            namaKKP, _
            NamaPeriode(bulanAwal, bulanAkhir, tahunFile), _
            NamaFileSingkat(pathFile, namaKKP, kodeMasa), _
            statusFile, _
            pathFile, _
            kodeKKP, _
            npwpFile, _
            bulanAwal, _
            bulanAkhir, _
            tahunFile, _
            identifierFile, _
            handler, _
            pesanError

    End If

End Sub

'Tujuan: Menambahkan baris daftar.
Private Sub TambahBarisDaftar( _
    ByVal frm As Object, _
    ByVal namaKKP As String, _
    ByVal periode As String, _
    ByVal namaFile As String, _
    ByVal statusFile As String, _
    ByVal pathFile As String, _
    ByVal kodeKKP As String, _
    ByVal npwp As String, _
    ByVal bulanAwal As Long, _
    ByVal bulanAkhir As Long, _
    ByVal tahun As Long, _
    ByVal identifier As String, _
    ByVal handler As String, _
    ByVal pesan As String)

    Dim indeks As Long

    frm.lstEWS.AddItem namaKKP
    indeks = frm.lstEWS.ListCount - 1

    IsiBarisDaftar _
        frm, indeks, namaKKP, periode, namaFile, statusFile, _
        pathFile, kodeKKP, npwp, bulanAwal, bulanAkhir, _
        tahun, identifier, handler, pesan

End Sub

'Tujuan: Menjalankan isi baris daftar.
Private Sub IsiBarisDaftar( _
    ByVal frm As Object, _
    ByVal indeks As Long, _
    ByVal namaKKP As String, _
    ByVal periode As String, _
    ByVal namaFile As String, _
    ByVal statusFile As String, _
    ByVal pathFile As String, _
    ByVal kodeKKP As String, _
    ByVal npwp As String, _
    ByVal bulanAwal As Long, _
    ByVal bulanAkhir As Long, _
    ByVal tahun As Long, _
    ByVal identifier As String, _
    ByVal handler As String, _
    ByVal pesan As String)

    With frm.lstEWS

        .List(indeks, LST_JENIS) = namaKKP
        .List(indeks, LST_PERIODE) = periode
        .List(indeks, LST_FILE) = namaFile
        .List(indeks, LST_STATUS) = statusFile
        .List(indeks, LST_PATH) = pathFile
        .List(indeks, LST_KODE) = kodeKKP
        .List(indeks, LST_NPWP) = npwp
        .List(indeks, LST_BULAN_AWAL) = bulanAwal
        .List(indeks, LST_BULAN_AKHIR) = bulanAkhir
        .List(indeks, LST_TAHUN) = tahun
    End With

End Sub

'Tujuan: Mencari identifier pada form.
Private Function CariIdentifierPadaForm( _
    ByVal frm As Object, _
    ByVal identifier As String) As Long

    Dim i As Long

    CariIdentifierPadaForm = -1

    If Len(identifier) = 0 Then Exit Function

    For i = 0 To frm.lstEWS.ListCount - 1

        If StrComp( _
            IdentifierDariBarisForm(frm, i), _
            identifier, _
            vbTextCompare) = 0 Then

            CariIdentifierPadaForm = i
            Exit Function

        End If

    Next i

End Function

'Bagian: identitas file.
'Tujuan: Membaca identitas dari nama file.
Public Function BacaIdentitasDariNamaFile( _
    ByVal pathFile As String, _
    ByRef kodeKKP As String, _
    ByRef namaKKP As String, _
    ByRef npwp As String, _
    ByRef kodeMasa As String, _
    ByRef bulanAwal As Long, _
    ByRef bulanAkhir As Long, _
    ByRef tahun As Long, _
    ByRef identifier As String, _
    ByRef dikenal As Boolean, _
    ByRef didukung As Boolean, _
    ByRef handler As String, _
    ByRef pesanError As String) As Boolean

    Dim namaFile As String
    Dim namaDasar As String
    Dim bagian() As String
    Dim kandidatKode As String
    Dim i As Long
    Dim j As Long

    BacaIdentitasDariNamaFile = False

    kodeKKP = vbNullString
    namaKKP = vbNullString
    npwp = vbNullString
    kodeMasa = vbNullString
    identifier = vbNullString
    handler = vbNullString
    pesanError = vbNullString

    bulanAwal = 0
    bulanAkhir = 0
    tahun = 0
    dikenal = False
    didukung = False

    namaFile = Dir$(pathFile)

    If Len(namaFile) = 0 Then
        pesanError = "File tidak ditemukan."
        Exit Function
    End If

    If InStrRev(namaFile, ".") > 0 Then
        namaDasar = Left$(namaFile, InStrRev(namaFile, ".") - 1)
    Else
        namaDasar = namaFile
    End If

    bagian = Split(UCase$(namaDasar), "_")

    For i = LBound(bagian) To UBound(bagian)

        If Len(bagian(i)) = 16 Then

            If HanyaAngka(bagian(i)) Then
                npwp = bagian(i)
            End If

        End If

        If Len(bagian(i)) = 8 Then

            If KodeMasaValid(bagian(i)) Then
                kodeMasa = bagian(i)
            End If

        End If

        If bagian(i) = "WP" Then

            kandidatKode = "WP"

            For j = i + 1 To UBound(bagian)

                kandidatKode = kandidatKode & "_" & bagian(j)

                If Left$(bagian(j), 3) = "TOT" Then

                    If Len(bagian(j)) = 5 _
                       And HanyaAngka(Right$(bagian(j), 2)) Then

                        kodeKKP = kandidatKode
                        Exit For

                    End If

                End If

            Next j

        End If

    Next i

    If Len(kodeKKP) = 0 Then
        pesanError = "Kode KKP EWS tidak ditemukan pada nama file."
        Exit Function
    End If

    If Len(npwp) = 0 Then
        pesanError = "NPWP 16 digit tidak ditemukan pada nama file."
        Exit Function
    End If

    If Len(kodeMasa) = 0 Then
        pesanError = _
            "Kode periode delapan digit tidak ditemukan pada nama file."
        Exit Function
    End If

    bulanAwal = CLng(Left$(kodeMasa, 2))
    bulanAkhir = CLng(Mid$(kodeMasa, 3, 2))
    tahun = CLng(Right$(kodeMasa, 4))

    DapatkanInfoKKP _
        kodeKKP, namaKKP, dikenal, didukung, handler

    If Len(namaKKP) = 0 Then
        namaKKP = kodeKKP
    End If

    identifier = _
        kodeKKP & "|" & npwp & "|" & kodeMasa

    BacaIdentitasDariNamaFile = True

End Function

'Tujuan: Membaca masa dari nama file.
Public Function BacaMasaDariNamaFile( _
    ByVal pathFile As String, _
    ByRef npwp As String, _
    ByRef kodeMasa As String, _
    ByRef bulan As Long, _
    ByRef tahun As Long, _
    ByRef identifier As String, _
    ByRef pesanError As String) As Boolean

    Dim kodeKKP As String
    Dim namaKKP As String
    Dim handler As String
    Dim bulanAkhir As Long
    Dim dikenal As Boolean
    Dim didukung As Boolean

    BacaMasaDariNamaFile = False

    If Not BacaIdentitasDariNamaFile( _
            pathFile, kodeKKP, namaKKP, npwp, kodeMasa, _
            bulan, bulanAkhir, tahun, identifier, _
            dikenal, didukung, handler, pesanError) Then

        Exit Function
    End If

    If bulan <> bulanAkhir Then

        pesanError = _
            "File EWS bukan file untuk satu masa." & vbCrLf & _
            "Kode ditemukan: " & kodeMasa

        Exit Function
    End If

    BacaMasaDariNamaFile = True

End Function

'Tujuan: Mengambil info kkp.
Private Sub DapatkanInfoKKP( _
    ByVal kodeKKP As String, _
    ByRef namaKKP As String, _
    ByRef dikenal As Boolean, _
    ByRef didukung As Boolean, _
    ByRef handler As String)

    dikenal = True
    didukung = False
    handler = vbNullString

    Select Case UCase$(kodeKKP)

        Case "WP_ICT_CI_SB_TOT01"
            namaKKP = "PPh Badan Bank Syariah"
        Case "WP_ICT_CI_SF_TOT01"
            namaKKP = "PPh Badan Sekuritas dan Keuangan"
        Case "WP_ICT_CI_CB_TOT01"
            namaKKP = "PPh Badan Bank Konvensional"
        Case "WP_ICT_CI_GTS_TOT01"
            namaKKP = "PPh Badan Umum, Perdagangan, dan Jasa"
        Case "WP_ICT_CI_MAN_TOT01"
            namaKKP = "PPh Badan Manufaktur"
        Case "WP_ICT_CI_PFI_TOT01"
            namaKKP = "PPh Badan Dana Pensiun dan Asuransi"
        Case "WP_ICT_CI_PI_TOT01"
            namaKKP = "PPh Badan Properti dan Infrastruktur"
        Case "WP_ICT_PI_TOT01"
            namaKKP = "PPh Orang Pribadi Pasal 25/29"
        Case KODE_EWS_PPH21
            namaKKP = NAMA_EWS_PPH21
            didukung = True
            handler = HANDLER_PPH21
        Case KODE_EWS_PPH21F
            namaKKP = NAMA_EWS_PPH21F
            didukung = True
            handler = HANDLER_PPH21F
        Case "WP_ICT_WH_TOT03"
            namaKKP = "PPh Pasal 15"
        Case "WP_ICT_WH_TOT04"
            namaKKP = "PPh Pasal 15 Final"
        Case "WP_ICT_WH_TOT05"
            namaKKP = "PPh Pasal 19 Final"
        Case "WP_ICT_WH_TOT06"
            namaKKP = "PPh Pasal 22"
        Case "WP_ICT_WH_TOT07"
            namaKKP = "PPh Pasal 22 Final"
        Case KODE_EWS_PPH23
            namaKKP = NAMA_EWS_PPH23
            didukung = True
            handler = HANDLER_PPH23
        Case KODE_EWS_PPH26
            namaKKP = NAMA_EWS_PPH26
            didukung = True
            handler = HANDLER_PPH26
        Case KODE_EWS_PPH42
            namaKKP = NAMA_EWS_PPH42
            didukung = True
            handler = HANDLER_PPH42
        Case "WP_ICT_WH_TOT11"
            namaKKP = "PPh Final atas Harta Bersih Tertentu"
        Case "WP_LBT_LB_TOT01"
            namaKKP = "PBB Sektor Perhutanan"
        Case "WP_LBT_LB_TOT02"
            namaKKP = "PBB Sektor Perkebunan"
        Case "WP_LBT_LB_TOT03"
            namaKKP = "PBB Sektor Pertambangan Minerba"
        Case "WP_LBT_LB_TOT04"
            namaKKP = "PBB Sektor Pertambangan Migas"
        Case "WP_LBT_LB_TOT05"
            namaKKP = "PBB Sektor Panas Bumi"
        Case "WP_LBT_LB_TOT06"
            namaKKP = "PBB Sektor Lainnya"
        Case "WP_OGT_OG_TOT01"
            namaKKP = "PPh Minyak dan Gas Bumi"
        Case "WP_SDT_SD_TOT01"
            namaKKP = "Bea Meterai"
        Case "WP_VAT_STLG_TOT01"
            namaKKP = "PPnBM Penyerahan Dalam Negeri"
        Case "WP_VAT_STLG_TOT02"
            namaKKP = "PPnBM Impor"
        Case "WP_VAT_STLG_TOT03"
            namaKKP = "PPnBM Dipungut Pemungut PPN"
        Case "WP_VAT_STLG_TOT04"
            namaKKP = "PPnBM Pembayaran Fasilitas"
        Case "WP_VAT_STLG_TOT05"
            namaKKP = "PPnBM Tanggung Jawab Renteng"
        Case "WP_VAT_STLG_TOT06"
            namaKKP = "PPnBM yang Seharusnya Tidak Dibebaskan"
        Case KODE_EWS_PPN_DN
            namaKKP = NAMA_EWS_PPN_DN
            didukung = True
            handler = HANDLER_PPN_DN
        Case "WP_VAT_VAT_TOT02"
            namaKKP = "PPN Impor"
        Case "WP_VAT_VAT_TOT03"
            namaKKP = "PPN Pemanfaatan BKP Tidak Berwujud/JKP LN"
        Case "WP_VAT_VAT_TOT04"
            namaKKP = "PPN Dipungut Pemungut PPN"
        Case "WP_VAT_VAT_TOT05"
            namaKKP = "PPN Kegiatan Membangun Sendiri"
        Case "WP_VAT_VAT_TOT06"
            namaKKP = "PPN Pembayaran Kembali Pajak Masukan"
        Case "WP_VAT_VAT_TOT07"
            namaKKP = "PPN Perolehan yang Tidak Seharusnya Dibebaskan"
        Case "WP_VAT_VAT_TOT08"
            namaKKP = "PPN Tanggung Jawab Renteng"
        Case "WP_VAT_VAT_TOT09"
            namaKKP = "PPN Pembayaran yang Semula Mendapat Fasilitas"
        Case Else
            dikenal = False
            namaKKP = kodeKKP

    End Select

End Sub

'Tujuan: Menjalankan update tampilan form.
Private Sub UpdateTampilanForm(ByVal frm As Object)

    Dim jumlahSiap As Long
    Dim jumlahVersiBaru As Long
    Dim jumlahImporUlang As Long
    Dim jumlahSudahDiimpor As Long
    Dim jumlahBelumDidukung As Long
    Dim jumlahTidakValid As Long
    Dim i As Long
    Dim statusFile As String
    Dim periodeAwal As Date
    Dim periodeAkhir As Date
    Dim periodeFile As Date
    Dim adaPeriode As Boolean
    Dim bulanFile As Long
    Dim tahunFile As Long

    For i = 0 To frm.lstEWS.ListCount - 1

        statusFile = CStr(frm.lstEWS.List(i, LST_STATUS))

        bulanFile = CLng(Val(CStr( _
            frm.lstEWS.List(i, LST_BULAN_AWAL))))
        tahunFile = CLng(Val(CStr( _
            frm.lstEWS.List(i, LST_TAHUN))))

        If bulanFile >= 1 And bulanFile <= 12 _
           And tahunFile >= 2000 And tahunFile <= 2100 Then

            periodeFile = DateSerial(tahunFile, bulanFile, 1)

            If Not adaPeriode Then
                periodeAwal = periodeFile
                periodeAkhir = periodeFile
                adaPeriode = True
            Else
                If periodeFile < periodeAwal Then
                    periodeAwal = periodeFile
                End If
                If periodeFile > periodeAkhir Then
                    periodeAkhir = periodeFile
                End If
            End If

        End If

        Select Case statusFile

            Case STATUS_SIAP
                jumlahSiap = jumlahSiap + 1
            Case STATUS_VERSI_BARU
                jumlahVersiBaru = jumlahVersiBaru + 1
            Case STATUS_SUDAH_DIIMPOR
                jumlahImporUlang = jumlahImporUlang + 1
            Case "Diimpor"
                jumlahSudahDiimpor = jumlahSudahDiimpor + 1
            Case STATUS_BELUM_DIDUKUNG
                jumlahBelumDidukung = jumlahBelumDidukung + 1
            Case Else
                jumlahTidakValid = jumlahTidakValid + 1

        End Select

    Next i

    If Len(frm.npwpReferensi) > 0 Then
        frm.lblWPValue.Caption = frm.npwpReferensi
    Else
        frm.lblWPValue.Caption = "-"
    End If

    If adaPeriode Then

        If periodeAwal = periodeAkhir Then
            frm.lblTahunValue.Caption = _
                NamaBulanIndonesia(Month(periodeAwal)) & _
                " " & CStr(Year(periodeAwal))
        Else
            frm.lblTahunValue.Caption = _
                NamaBulanIndonesia(Month(periodeAwal)) & _
                " " & CStr(Year(periodeAwal)) & " - " & _
                NamaBulanIndonesia(Month(periodeAkhir)) & _
                " " & CStr(Year(periodeAkhir))
        End If

    Else
        frm.lblTahunValue.Caption = "-"
    End If

    frm.lblJumlah.Caption = _
        CStr(jumlahSiap) & " siap; " & _
        CStr(jumlahVersiBaru) & " revisi; " & _
        CStr(jumlahImporUlang) & " impor ulang; " & _
        CStr(jumlahSudahDiimpor) & " selesai; " & _
        CStr(jumlahBelumDidukung) & " belum didukung; " & _
        CStr(jumlahTidakValid) & " tidak valid"

    If jumlahSiap + jumlahVersiBaru + jumlahImporUlang > 0 Then

        If jumlahImporUlang > 0 _
           And jumlahSiap + jumlahVersiBaru = 0 Then

            frm.lblStatus.Caption = _
                "Data pernah diimpor. Klik Impor untuk mengimpor ulang."
        Else
            frm.lblStatus.Caption = _
                "File baru, revisi, dan impor ulang siap diproses."
        End If

        frm.cmdImport.Enabled = True

    Else

        If jumlahSudahDiimpor > 0 Then
            frm.lblStatus.Caption = _
                "File yang sama sudah pernah berhasil diimpor."
        Else
            frm.lblStatus.Caption = _
                "Belum ada file dengan mapping impor yang aktif."
        End If

        frm.cmdImport.Enabled = False

    End If

End Sub

'Bagian: orkestrasi batch.
'Tujuan: Mengimpor semua ews.
Public Sub ImportSemuaEWS(ByVal frm As Object)

    Dim i As Long
    Dim jumlahSiap As Long
    Dim jumlahVersiBaru As Long
    Dim jumlahImporUlang As Long
    Dim jumlahBerhasil As Long
    Dim jumlahGagal As Long
    Dim pathFile As String
    Dim kodeKKP As String
    Dim handler As String
    Dim namaKKPHandler As String
    Dim dikenalHandler As Boolean
    Dim didukungHandler As Boolean
    Dim namaWP As String
    Dim pesanError As String
    Dim pesanKonfirmasi As String

    Dim screenUpdatingLama As Boolean
    Dim enableEventsLama As Boolean
    Dim displayAlertsLama As Boolean
    Dim calculationLama As XlCalculation
    Dim askLinksLama As Boolean

    For i = 0 To frm.lstEWS.ListCount - 1

        If StatusDapatDiimpor( _
                CStr(frm.lstEWS.List(i, LST_STATUS))) Then

            jumlahSiap = jumlahSiap + 1

            If CStr(frm.lstEWS.List(i, LST_STATUS)) = _
                    STATUS_VERSI_BARU Then

                jumlahVersiBaru = jumlahVersiBaru + 1
            ElseIf CStr(frm.lstEWS.List(i, LST_STATUS)) = _
                    STATUS_SUDAH_DIIMPOR Then

                jumlahImporUlang = jumlahImporUlang + 1
            End If

        End If

    Next i

    If jumlahSiap = 0 Then

        MsgBox _
            "Belum ada file EWS yang siap diimpor.", _
            vbExclamation, _
            "Impor EWS"

        Exit Sub
    End If

    ValidasiUlangDaftar frm

    jumlahSiap = 0
    jumlahVersiBaru = 0
    jumlahImporUlang = 0

    For i = 0 To frm.lstEWS.ListCount - 1

        If StatusDapatDiimpor( _
                CStr(frm.lstEWS.List(i, LST_STATUS))) Then

            jumlahSiap = jumlahSiap + 1

            If CStr(frm.lstEWS.List(i, LST_STATUS)) = _
                    STATUS_VERSI_BARU Then

                jumlahVersiBaru = jumlahVersiBaru + 1
            ElseIf CStr(frm.lstEWS.List(i, LST_STATUS)) = _
                    STATUS_SUDAH_DIIMPOR Then

                jumlahImporUlang = jumlahImporUlang + 1
            End If

        End If

    Next i

    UpdateTampilanForm frm

    If jumlahSiap = 0 Then

        MsgBox _
            "Tidak ada file yang lolos validasi ulang.", _
            vbExclamation, _
            "Impor EWS"

        Exit Sub
    End If

    If jumlahImporUlang > 0 Then

        pesanKonfirmasi = _
            "Data sudah pernah diimpor: " & _
            CStr(jumlahImporUlang) & " file." & _
            vbCrLf & vbCrLf & _
            "Impor ulang akan menimpa data pada periode dan KKP " & _
            "yang sama." & vbCrLf & vbCrLf & _
            "Yakin akan impor ulang?"

        If jumlahSiap > jumlahImporUlang Then
            pesanKonfirmasi = pesanKonfirmasi & _
                vbCrLf & vbCrLf & _
                CStr(jumlahSiap - jumlahImporUlang) & _
                " file baru atau revisi juga akan diproses."
        End If

    Else

        pesanKonfirmasi = _
            "Impor " & jumlahSiap & _
            " file EWS ke EWSSYNC?"

    End If

    If jumlahVersiBaru > 0 Then
        pesanKonfirmasi = pesanKonfirmasi & _
            vbCrLf & vbCrLf & _
            CStr(jumlahVersiBaru) & _
            " file adalah versi baru dan akan menimpa " & _
            "data impor sebelumnya."
    End If

    pesanKonfirmasi = pesanKonfirmasi & _
        vbCrLf & vbCrLf & _
        "File yang gagal akan dicatat dan proses file lain tetap berjalan."

    If MsgBox( _
        pesanKonfirmasi, _
        vbQuestion + vbYesNo + vbDefaultButton2, _
        "Konfirmasi Impor") <> vbYes Then

        Exit Sub
    End If

    screenUpdatingLama = Application.ScreenUpdating
    enableEventsLama = Application.EnableEvents
    displayAlertsLama = Application.DisplayAlerts
    calculationLama = Application.Calculation
    askLinksLama = Application.AskToUpdateLinks

    On Error GoTo GagalUmum

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False

    For i = 0 To frm.lstEWS.ListCount - 1

        If StatusDapatDiimpor( _
                CStr(frm.lstEWS.List(i, LST_STATUS))) Then

            pathFile = CStr(frm.lstEWS.List(i, LST_PATH))
            kodeKKP = CStr(frm.lstEWS.List(i, LST_KODE))

            DapatkanInfoKKP _
                kodeKKP, _
                namaKKPHandler, _
                dikenalHandler, _
                didukungHandler, _
                handler

            frm.lstEWS.List(i, LST_STATUS) = "Mengimpor..."
            frm.lblStatus.Caption = _
                "Mengimpor " & _
                CStr(frm.lstEWS.List(i, LST_JENIS)) & " " & _
                CStr(frm.lstEWS.List(i, LST_PERIODE)) & "..."

            frm.Repaint
            DoEvents

            pesanError = vbNullString
            namaWP = vbNullString

            If ImportSatuFileEWS( _
                    pathFile, _
                    kodeKKP, _
                    handler, _
                    CStr(frm.lstEWS.List(i, LST_NPWP)), _
                    CLng(frm.lstEWS.List(i, LST_BULAN_AWAL)), _
                    CLng(frm.lstEWS.List(i, LST_BULAN_AKHIR)), _
                    CLng(frm.lstEWS.List(i, LST_TAHUN)), _
                    namaWP, _
                    pesanError) Then

                frm.lstEWS.List(i, LST_STATUS) = "Diimpor"
                jumlahBerhasil = jumlahBerhasil + 1

                If Len(frm.namaWPReferensi) = 0 Then
                    frm.namaWPReferensi = namaWP
                End If

                TulisLogImpor _
                    pathFile, kodeKKP, _
                    CStr(frm.lstEWS.List(i, LST_PERIODE)), _
                    "Berhasil", pesanError, _
                    CStr(frm.lstEWS.List(i, LST_NPWP)), _
                    CLng(frm.lstEWS.List(i, LST_BULAN_AWAL)), _
                    CLng(frm.lstEWS.List(i, LST_BULAN_AKHIR)), _
                    CLng(frm.lstEWS.List(i, LST_TAHUN))

            Else

                frm.lstEWS.List(i, LST_STATUS) = "Gagal"
                jumlahGagal = jumlahGagal + 1

                TulisLogImpor _
                    pathFile, kodeKKP, _
                    CStr(frm.lstEWS.List(i, LST_PERIODE)), _
                    "Gagal", pesanError, _
                    CStr(frm.lstEWS.List(i, LST_NPWP)), _
                    CLng(frm.lstEWS.List(i, LST_BULAN_AWAL)), _
                    CLng(frm.lstEWS.List(i, LST_BULAN_AKHIR)), _
                    CLng(frm.lstEWS.List(i, LST_TAHUN))

            End If

        End If

    Next i

    HitungUlangKKPLokal TAXKEY_PPN_DN
    HitungUlangKKPLokal TAXKEY_PPH21
    HitungUlangKKPLokal TAXKEY_PPH21F
    HitungUlangKKPLokal TAXKEY_PPH23
    HitungUlangKKPLokal TAXKEY_PPH26

Selesai:

    On Error Resume Next

    Application.ScreenUpdating = screenUpdatingLama
    Application.EnableEvents = enableEventsLama
    Application.DisplayAlerts = displayAlertsLama
    Application.Calculation = calculationLama
    Application.AskToUpdateLinks = askLinksLama

    On Error GoTo 0

    UpdateTampilanForm frm

    If jumlahGagal = 0 Then

        frm.lblStatus.Caption = _
            CStr(jumlahBerhasil) & " file berhasil diimpor."

    Else

        frm.lblStatus.Caption = _
            CStr(jumlahBerhasil) & " berhasil, " & _
            CStr(jumlahGagal) & " gagal."

        MsgBox _
            "Proses impor selesai." & vbCrLf & vbCrLf & _
            "Berhasil : " & jumlahBerhasil & vbCrLf & _
            "Gagal    : " & jumlahGagal & vbCrLf & vbCrLf & _
            "Rincian kegagalan tersedia pada sheet '" & _
            NAMA_SHEET_LOG & "'.", _
            vbExclamation, _
            "Impor Selesai"

    End If

    Exit Sub

GagalUmum:

    pesanError = Err.Description

    On Error Resume Next

    TulisLogImpor _
        pathFile, kodeKKP, vbNullString, _
        "Gagal umum", pesanError

    On Error GoTo 0

    jumlahGagal = jumlahGagal + 1
    Resume Selesai

End Sub

'Tujuan: Memvalidasi ulang daftar.
Private Sub ValidasiUlangDaftar(ByVal frm As Object)

    Dim i As Long
    Dim pathFile As String
    Dim kodeKKP As String
    Dim namaKKP As String
    Dim npwp As String
    Dim kodeMasa As String
    Dim identifier As String
    Dim handler As String
    Dim pesanError As String
    Dim statusFile As String
    Dim bulanAwal As Long
    Dim bulanAkhir As Long
    Dim tahun As Long
    Dim dikenal As Boolean
    Dim didukung As Boolean

    For i = 0 To frm.lstEWS.ListCount - 1

        If StatusDapatDiimpor( _
                CStr(frm.lstEWS.List(i, LST_STATUS))) Then

            pathFile = CStr(frm.lstEWS.List(i, LST_PATH))

            If Not BacaIdentitasDariNamaFile( _
                    pathFile, kodeKKP, namaKKP, npwp, kodeMasa, _
                    bulanAwal, bulanAkhir, tahun, identifier, _
                    dikenal, didukung, handler, pesanError) Then

                frm.lstEWS.List(i, LST_STATUS) = "Tidak valid"

            ElseIf Not didukung Then

                frm.lstEWS.List(i, LST_STATUS) = _
                    STATUS_BELUM_DIDUKUNG

            ElseIf npwp <> frm.npwpReferensi Then

                frm.lstEWS.List(i, LST_STATUS) = "NPWP berbeda"

            ElseIf identifier <> _
                IdentifierDariBarisForm(frm, i) Then

                frm.lstEWS.List(i, LST_STATUS) = "Identitas berubah"

            Else

                statusFile = STATUS_SIAP
                pesanError = vbNullString

                TerapkanStatusRiwayat _
                    pathFile, kodeKKP, npwp, _
                    bulanAwal, bulanAkhir, tahun, _
                    statusFile, pesanError

                frm.lstEWS.List(i, LST_STATUS) = statusFile

            End If

        End If

    Next i

End Sub

'Tujuan: Menjalankan identifier dari baris form.
Private Function IdentifierDariBarisForm( _
    ByVal frm As Object, _
    ByVal indeks As Long) As String

    IdentifierDariBarisForm = _
        CStr(frm.lstEWS.List(indeks, LST_KODE)) & "|" & _
        CStr(frm.lstEWS.List(indeks, LST_NPWP)) & "|" & _
        Format$( _
            CLng(frm.lstEWS.List(indeks, LST_BULAN_AWAL)), _
            "00") & _
        Format$( _
            CLng(frm.lstEWS.List(indeks, LST_BULAN_AKHIR)), _
            "00") & _
        Format$( _
            CLng(frm.lstEWS.List(indeks, LST_TAHUN)), _
            "0000")

End Function

'Tujuan: Menentukan status dapat diimpor.
Private Function StatusDapatDiimpor( _
    ByVal statusFile As String) As Boolean

    StatusDapatDiimpor = _
        StrComp(statusFile, STATUS_SIAP, vbTextCompare) = 0 _
        Or StrComp( _
            statusFile, _
            STATUS_VERSI_BARU, _
            vbTextCompare) = 0 _
        Or StrComp( _
            statusFile, _
            STATUS_SUDAH_DIIMPOR, _
            vbTextCompare) = 0

End Function

'Tujuan: Menjalankan terapkan status riwayat.
Private Sub TerapkanStatusRiwayat( _
    ByVal pathFile As String, _
    ByVal kodeKKP As String, _
    ByVal npwp As String, _
    ByVal bulanAwal As Long, _
    ByVal bulanAkhir As Long, _
    ByVal tahun As Long, _
    ByRef statusFile As String, _
    ByRef pesan As String)

    Dim pernahBerhasil As Boolean
    Dim versiSama As Boolean

    CekRiwayatImpor _
        pathFile, kodeKKP, npwp, _
        bulanAwal, bulanAkhir, tahun, _
        pernahBerhasil, versiSama

    If Not pernahBerhasil Then Exit Sub

    If versiSama Then

        statusFile = STATUS_SUDAH_DIIMPOR
        pesan = _
            "Versi file ini sudah pernah berhasil diimpor."

    Else

        statusFile = STATUS_VERSI_BARU
        pesan = _
            "Data dengan KKP, NPWP, dan periode yang sama " & _
            "pernah diimpor, tetapi file ini telah berubah."

    End If

End Sub

'Tujuan: Menjalankan cek riwayat impor.
Private Sub CekRiwayatImpor( _
    ByVal pathFile As String, _
    ByVal kodeKKP As String, _
    ByVal npwp As String, _
    ByVal bulanAwal As Long, _
    ByVal bulanAkhir As Long, _
    ByVal tahun As Long, _
    ByRef pernahBerhasil As Boolean, _
    ByRef versiSama As Boolean)

    Dim ws As Worksheet
    Dim baris As Long
    Dim barisAkhir As Long
    Dim kunci As String
    Dim ukuranFile As Double
    Dim tanggalModifikasi As Date

    pernahBerhasil = False
    versiSama = False

    On Error GoTo Selesai

    If Len(Dir$(pathFile)) = 0 Then Exit Sub

    If Not WorksheetAda(NAMA_SHEET_LOG, ThisWorkbook) Then
        Exit Sub
    End If

    ukuranFile = CDbl(FileLen(pathFile))
    tanggalModifikasi = FileDateTime(pathFile)
    kunci = KunciImpor( _
        kodeKKP, npwp, bulanAwal, bulanAkhir, tahun)

    Set ws = ThisWorkbook.Worksheets(NAMA_SHEET_LOG)
    barisAkhir = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For baris = barisAkhir To 2 Step -1

        If StrComp( _
                Trim$(CStr(ws.Cells(baris, 5).Value2)), _
                "Berhasil", _
                vbTextCompare) = 0 Then

            If StrComp( _
                    Trim$(CStr(ws.Cells(baris, 11).Value2)), _
                    kunci, _
                    vbTextCompare) = 0 Then

                pernahBerhasil = True

                If IsNumeric(ws.Cells(baris, 12).Value2) _
                   And IsDate(ws.Cells(baris, 13).Value) Then

                    versiSama = _
                        Abs( _
                            CDbl(ws.Cells(baris, 12).Value2) - _
                            ukuranFile) < 0.5 _
                        And Abs( _
                            CDbl(ws.Cells(baris, 13).Value2) - _
                            CDbl(tanggalModifikasi)) < _
                            (1# / 86400#)

                End If

                Exit For

            ElseIf Len(Trim$(CStr( _
                        ws.Cells(baris, 11).Value2))) = 0 _
                   And StrComp( _
                        Trim$(CStr( _
                            ws.Cells(baris, 2).Value2)), _
                        kodeKKP, _
                        vbTextCompare) = 0 _
                   And StrComp( _
                        Trim$(CStr( _
                            ws.Cells(baris, 3).Value2)), _
                        NamaPeriode( _
                            bulanAwal, bulanAkhir, tahun), _
                        vbTextCompare) = 0 _
                   And InStr( _
                        1, _
                        CStr(ws.Cells(baris, 4).Value2), _
                        npwp, _
                        vbTextCompare) > 0 Then

                pernahBerhasil = True
                versiSama = False
                Exit For

            End If

        End If

    Next baris

Selesai:

End Sub

'Tujuan: Membentuk kunci impor.
Private Function KunciImpor( _
    ByVal kodeKKP As String, _
    ByVal npwp As String, _
    ByVal bulanAwal As Long, _
    ByVal bulanAkhir As Long, _
    ByVal tahun As Long) As String

    KunciImpor = _
        UCase$(Trim$(kodeKKP)) & "|" & _
        BersihkanNPWP(npwp) & "|" & _
        Format$(bulanAwal, "00") & _
        Format$(bulanAkhir, "00") & _
        Format$(tahun, "0000")

End Function

'Bagian: transaksi per file.
'Tujuan: Mengimpor satu file ews.
Private Function ImportSatuFileEWS( _
    ByVal pathAsli As String, _
    ByVal kodeKKP As String, _
    ByVal handler As String, _
    ByVal npwpDiharapkan As String, _
    ByVal bulanAwal As Long, _
    ByVal bulanAkhir As Long, _
    ByVal tahun As Long, _
    ByRef namaWP As String, _
    ByRef pesanError As String) As Boolean

    Dim wbEWS As Workbook
    Dim wsEWS As Worksheet
    Dim npwpWorkbook As String
    Dim tahap As String
    Dim workbookDibukaImporter As Boolean
    Dim workbookSudahTerbuka As Boolean

    Dim screenUpdatingSebelum As Boolean
    Dim enableEventsSebelum As Boolean
    Dim displayAlertsSebelum As Boolean
    Dim askLinksSebelum As Boolean
    Dim calculationSebelum As XlCalculation
    Dim automationSecuritySebelum As Long
    Dim nomorErrorBuka As Long
    Dim pesanErrorBuka As String

    ImportSatuFileEWS = False
    pesanError = vbNullString
    namaWP = vbNullString

    Set wbEWS = DapatkanWorkbookTerbuka(pathAsli)
    workbookSudahTerbuka = Not wbEWS Is Nothing

    If kodeKKP = KODE_EWS_PPH26 And handler = HANDLER_PPH26 Then
        If Not KonfirmasiTrustedLocationPPh26(pathAsli) Then
            pesanError = "Impor dibatalkan karena Trusted Location PPh26 belum dikonfirmasi."
            Exit Function
        End If
        ImportSatuFileEWS = ImportPPh26PascaError(pathAsli, npwpDiharapkan, _
            bulanAwal, bulanAkhir, tahun, namaWP, pesanError)
        Exit Function
    End If

    screenUpdatingSebelum = Application.ScreenUpdating
    enableEventsSebelum = Application.EnableEvents
    displayAlertsSebelum = Application.DisplayAlerts
    askLinksSebelum = Application.AskToUpdateLinks
    calculationSebelum = Application.Calculation
    automationSecuritySebelum = Application.AutomationSecurity

    On Error GoTo Gagal

    tahap = "validasi file"

    If Len(Dir$(pathAsli)) = 0 Then
        Err.Raise vbObjectError + 3100, , "File sumber tidak ditemukan."
    End If

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.AskToUpdateLinks = True
    Application.Calculation = xlCalculationAutomatic
    Application.AutomationSecurity = msoAutomationSecurityByUI
    Application.Visible = True

    If workbookSudahTerbuka Then
        Application.StatusBar = _
            "Menggunakan EWS yang sudah terbuka: " & _
            Dir$(pathAsli) & "..."
    Else
        Application.StatusBar = _
            "Membuka dan membangun ulang " & Dir$(pathAsli) & "..."
    End If

    DoEvents

    If workbookSudahTerbuka Then

        tahap = "menggunakan file EWS yang sudah terbuka"

    Else

        tahap = "membuka file EWS"

        If handler = HANDLER_PPH42 Then
            Application.SendKeys "{ENTER}", False
            DoEvents
        End If

        On Error Resume Next

        Set wbEWS = Workbooks.Open( _
            Filename:=pathAsli, _
            ReadOnly:=False, _
            IgnoreReadOnlyRecommended:=True, _
            AddToMru:=False, _
            Notify:=False)

        nomorErrorBuka = Err.Number
        pesanErrorBuka = Err.Description

        Err.Clear
        On Error GoTo Gagal

        If wbEWS Is Nothing Then

            If nomorErrorBuka <> 0 Then
                Err.Raise nomorErrorBuka, , pesanErrorBuka
            Else
                Err.Raise vbObjectError + 3108, , _
                    "File EWS tidak dapat dibuka."
            End If

        End If

        workbookDibukaImporter = True

    End If

    If wbEWS.Windows.Count > 0 Then
        wbEWS.Windows(1).Visible = True
    End If

    wbEWS.Activate
    DoEvents

    If workbookDibukaImporter _
       And wbEWS.ReadOnly Then

        Err.Raise _
            vbObjectError + 3107, , _
            "EWS hanya dapat dibuka dalam mode read-only. File mungkin " & _
            "sedang terbuka pada instance Excel lain. Tutup file pada " & _
            "instance tersebut, lalu ulangi impor."
    End If

    tahap = "menunggu rebuild EWS"

    If Not TungguRebuildEWS( _
            wbEWS, _
            handler, _
            BATAS_TUNGGU_REBUILD_DETIK, _
            pesanError, _
            workbookSudahTerbuka) Then

        pesanError = _
            "Tahap: " & tahap & vbCrLf & pesanError

        GoTo Selesai
    End If

    If workbookDibukaImporter Then

        tahap = "menyimpan hasil rebuild EWS"
        Application.StatusBar = _
            "Menyimpan hasil rebuild " & wbEWS.Name & "..."
        DoEvents

        If wbEWS.ReadOnly Then
            Err.Raise _
                vbObjectError + 3107, , _
                "EWS terbuka dalam mode read-only. File mungkin " & _
                "sedang terbuka pada instance Excel lain."
        End If

        wbEWS.Save

    End If

    tahap = "memvalidasi struktur EWS"
    Application.StatusBar = _
        "Rebuild selesai. Memvalidasi " & wbEWS.Name & "..."
    DoEvents

    If Not WorksheetAda(NAMA_SHEET_MAIN_EWS, wbEWS) Then
        Err.Raise _
            vbObjectError + 3101, , _
            "Sheet '" & NAMA_SHEET_MAIN_EWS & "' tidak ditemukan."
    End If

    Set wsEWS = wbEWS.Worksheets(NAMA_SHEET_MAIN_EWS)

    BacaMetadataDariSheet _
        wsEWS, namaWP, npwpWorkbook

    If Len(npwpWorkbook) = 16 Then

        If npwpWorkbook <> npwpDiharapkan Then

            Err.Raise _
                vbObjectError + 3102, , _
                "NPWP dalam workbook berbeda dengan nama file." & _
                vbCrLf & _
                "Nama file: " & npwpDiharapkan & vbCrLf & _
                "Workbook : " & npwpWorkbook

        End If

    End If

    Select Case handler

        Case HANDLER_PPN_DN

            tahap = "menyalin data PPN Dalam Negeri"
            Application.StatusBar = _
                "Menyalin data PPN Dalam Negeri ke EWSSYNC..."
            DoEvents

            If kodeKKP <> KODE_EWS_PPN_DN Then
                Err.Raise _
                    vbObjectError + 3103, , _
                    "Handler PPN DN menerima kode KKP yang tidak sesuai."
            End If

            If bulanAwal <> bulanAkhir Then
                Err.Raise _
                    vbObjectError + 3104, , _
                    "PPN Dalam Negeri harus menggunakan satu masa pajak."
            End If

            If Not ImportPPNDariWorkbook( _
                    wsEWS, bulanAwal, tahun, pesanError) Then

                GoTo Selesai
            End If

        Case HANDLER_PPH21

            tahap = "menyalin data PPh Pasal 21"
            Application.StatusBar = _
                "Menyalin data PPh Pasal 21 ke EWSSYNC..."
            DoEvents

            If kodeKKP <> KODE_EWS_PPH21 Then
                Err.Raise _
                    vbObjectError + 3110, , _
                    "Handler PPh 21 menerima kode KKP yang tidak sesuai."
            End If

            If bulanAwal <> bulanAkhir Then
                Err.Raise _
                    vbObjectError + 3111, , _
                    "PPh Pasal 21 harus menggunakan satu masa pajak."
            End If

            If Not ImportPPh21DariWorkbook( _
                    wsEWS, bulanAwal, tahun, pesanError) Then

                GoTo Selesai
            End If

        Case HANDLER_PPH21F

            tahap = "menyalin data PPh Pasal 21 Final"
            Application.StatusBar = _
                "Menyalin data PPh Pasal 21 Final ke EWSSYNC..."
            DoEvents

            If kodeKKP <> KODE_EWS_PPH21F Then
                Err.Raise _
                    vbObjectError + 3116, , _
                    "Handler PPh 21 Final menerima kode KKP yang tidak sesuai."
            End If

            If bulanAwal <> bulanAkhir Then
                Err.Raise _
                    vbObjectError + 3117, , _
                    "PPh Pasal 21 Final harus menggunakan satu masa pajak."
            End If

            If Not ImportPPh21FinalDariWorkbook( _
                    wbEWS, wsEWS, bulanAwal, tahun, pesanError) Then

                GoTo Selesai
            End If

        Case HANDLER_PPH23

            tahap = "menyalin data PPh Pasal 23"
            Application.StatusBar = "Menyalin data PPh Pasal 23 ke EWSSYNC..."
            DoEvents

            If kodeKKP <> KODE_EWS_PPH23 Then
                Err.Raise vbObjectError + 3112, , _
                    "Handler PPh 23 menerima kode KKP yang tidak sesuai."
            End If
            If bulanAwal <> bulanAkhir Then
                Err.Raise vbObjectError + 3113, , _
                    "PPh Pasal 23 harus menggunakan satu masa pajak."
            End If
            If Not ImportPPh23DariWorkbook( _
                    wsEWS, bulanAwal, tahun, pesanError) Then
                GoTo Selesai
            End If

        Case HANDLER_PPH26

            tahap = "menyalin data PPh Pasal 26"
            Application.StatusBar = "Menyalin data PPh Pasal 26 ke EWSSYNC..."
            DoEvents

            If kodeKKP <> KODE_EWS_PPH26 Then
                Err.Raise vbObjectError + 3112, , _
                    "Handler PPh 26 menerima kode KKP yang tidak sesuai."
            End If
            If bulanAwal <> bulanAkhir Then
                Err.Raise vbObjectError + 3113, , _
                    "PPh Pasal 26 harus menggunakan satu masa pajak."
            End If
            If Not ImportPPh26DariWorkbook( _
                    wsEWS, bulanAwal, tahun, pesanError) Then
                GoTo Selesai
            End If

        Case HANDLER_PPH42

            tahap = "menyalin data PPh Pasal 4 ayat (2)"
            Application.StatusBar = "Menyalin data PPh Pasal 4 ayat (2) ke EWSSYNC..."
            DoEvents
            If kodeKKP <> KODE_EWS_PPH42 Then
                Err.Raise vbObjectError + 3114, , _
                    "Handler PPh 4(2) menerima kode KKP yang tidak sesuai."
            End If
            If bulanAwal <> bulanAkhir Then
                Err.Raise vbObjectError + 3115, , _
                    "PPh Pasal 4 ayat (2) harus menggunakan satu masa pajak."
            End If
            If Not ImportPPh42DariWorkbook( _
                    wsEWS, bulanAwal, tahun, pesanError) Then
                GoTo Selesai
            End If

        Case Else

            Err.Raise _
                vbObjectError + 3105, , _
                "Handler impor belum tersedia untuk " & kodeKKP & "."

    End Select

    ImportSatuFileEWS = True

Selesai:

    On Error Resume Next

    If workbookDibukaImporter _
       And Not wbEWS Is Nothing Then
        wbEWS.Close SaveChanges:=False
    End If

    ThisWorkbook.Activate

    Application.StatusBar = False
    Application.ScreenUpdating = screenUpdatingSebelum
    Application.EnableEvents = enableEventsSebelum
    Application.DisplayAlerts = displayAlertsSebelum
    Application.AskToUpdateLinks = askLinksSebelum
    Application.Calculation = calculationSebelum
    Application.AutomationSecurity = automationSecuritySebelum

    On Error GoTo 0

    Exit Function

Gagal:

    pesanError = _
        "Gagal mengimpor " & Dir$(pathAsli) & "." & _
        vbCrLf & "Tahap: " & tahap & vbCrLf & _
        Err.Description

    Resume Selesai

End Function

'Bagian: helper bersama PPh pemotongan.
'Tujuan: Membaca nama baris PPh dari EWS.
Private Function AmbilNamaEksplisitPPhPemotonganEWS( _
    ByVal ws As Worksheet, _
    ByVal baris As Long) As String

    Dim c As Long
    Dim teks As String

    For c = KOLOM_NAMA_PPH21_EWS_AKHIR _
        To KOLOM_NAMA_PPH21_EWS_AWAL Step -1

        teks = NilaiTeksAman(ws.Cells(baris, c).Value2)

        If NamaBarisDapatDipakai(teks) Then
            AmbilNamaEksplisitPPhPemotonganEWS = teks
            Exit Function
        End If

    Next c

End Function

'Tujuan: Membaca nama baris PPh dari KKP lokal.
Private Function AmbilNamaEksplisitPPhPemotonganKKP( _
    ByVal ws As Worksheet, _
    ByVal baris As Long) As String

    Dim c As Long
    Dim teks As String

    For c = KOLOM_NAMA_PPH21_KKP_AKHIR _
        To KOLOM_NAMA_PPH21_KKP_AWAL Step -1

        teks = NilaiTeksAman(ws.Cells(baris, c).Value2)

        If NamaBarisDapatDipakai(teks) Then
            AmbilNamaEksplisitPPhPemotonganKKP = teks
            Exit Function
        End If

    Next c

End Function

'Tujuan: Memperbarui konteks kelompok DPP, PPh, O, atau P.
Private Sub PerbaruiKonteksPPhPemotongan( _
    ByVal namaBaris As String, _
    ByRef konteks As String)

    Dim namaNormal As String

    namaNormal = NormalisasiNamaBaris(namaBaris)

    If InStr(1, namaNormal, _
            "PENGHASILANKENAPAJAKDASARPENGENAANPAJAK", _
            vbTextCompare) > 0 Then

        konteks = "DPP"

    ElseIf namaNormal = "PPHTERUTANG" Then

        konteks = "PPH"

    ElseIf namaNormal = "DIPERHITUNGKAN" Then

        konteks = "O"

    ElseIf InStr(1, namaNormal, _
            "PPHYANGSEHARUSNYATIDAKTERUTANG", _
            vbTextCompare) > 0 Then

        konteks = "P"

    End If

End Sub

'Tujuan: Mencari kolom nilai Wajib Pajak pada MAIN LEVEL.
Private Function CariKolomNilaiWPPPhPemotongan( _
    ByVal ws As Worksheet, ByRef pesanError As String, _
    ByVal namaPajak As String) As Long
    Dim sel As Range, nama As String, kolom As Long
    For Each sel In ws.UsedRange.Cells
        nama = NormalisasiNamaBaris(NilaiTeksAman(sel.Value2))
        If InStr(1, nama, "MENURUTWAJIBPAJAK", vbTextCompare) = 1 Then
            If kolom <> 0 And kolom <> sel.Column Then
                pesanError = "Kolom Menurut Wajib Pajak " & namaPajak & " ambigu."
                Exit Function
            End If
            kolom = sel.Column
        End If
    Next sel
    If kolom = 0 Then pesanError = "Header Menurut Wajib Pajak " & namaPajak & " tidak ditemukan."
    CariKolomNilaiWPPPhPemotongan = kolom
End Function

'Bagian: handler PPh 21.
'Tujuan: Mengimpor pph 21 dari workbook.
Private Function ImportPPh21DariWorkbook( _
    ByVal wsEWS As Worksheet, _
    ByVal bulanPajak As Long, _
    ByVal tahunPajak As Long, _
    ByRef pesanError As String) As Boolean

    Dim wsLokal As Worksheet
    Dim petaSumber As Object
    Dim petaTujuan As Object

    Dim namaSheetRingkasan As String
    Dim kolomTujuan As Long
    Dim barisAwal As Long
    Dim barisAkhir As Long
    Dim baris As Long
    Dim jumlahMapping As Long

    Dim namaBaris As String
    Dim konteks As String
    Dim kunci As String
    Dim dataSumber As Variant

    Dim daftarBaris() As Long
    Dim daftarNilaiBaru() As Variant
    Dim daftarNilaiLama() As Variant
    Dim i As Long
    Dim penulisanDimulai As Boolean

    ImportPPh21DariWorkbook = False

    On Error GoTo Gagal

    If bulanPajak < 1 Or bulanPajak > 12 Then
        Err.Raise vbObjectError + 3300, , "Nomor bulan tidak valid."
    End If

    If tahunPajak < 2000 Or tahunPajak > 2100 Then
        Err.Raise vbObjectError + 3301, , "Tahun pajak tidak valid."
    End If

    namaSheetRingkasan = _
        NamaSheetRingkasanPajak(TAXKEY_PPH21)

    If Not WorksheetAda(namaSheetRingkasan, ThisWorkbook) Then
        Err.Raise _
            vbObjectError + 3302, , _
            "Sheet lokal '" & namaSheetRingkasan & _
            "' tidak ditemukan."
    End If

    Set wsLokal = _
        ThisWorkbook.Worksheets(namaSheetRingkasan)

    kolomTujuan = CariKolomPeriodePPN( _
        wsLokal, bulanPajak, tahunPajak, pesanError)

    If kolomTujuan = 0 Then GoTo GagalTanpaError

    If Not BangunPetaSumberPPh21( _
            wsEWS, petaSumber, pesanError) Then
        GoTo GagalTanpaError
    End If

    If Not CariBlokMenurutWP( _
            wsLokal, barisAwal, barisAkhir, pesanError) Then
        GoTo GagalTanpaError
    End If

    Set petaTujuan = CreateObject("Scripting.Dictionary")
    petaTujuan.CompareMode = vbTextCompare

    ReDim daftarBaris(1 To barisAkhir - barisAwal + 1)
    ReDim daftarNilaiBaru(1 To barisAkhir - barisAwal + 1)
    ReDim daftarNilaiLama(1 To barisAkhir - barisAwal + 1)

    konteks = vbNullString

    For baris = barisAwal To barisAkhir

        namaBaris = AmbilNamaEksplisitPPh21KKP(wsLokal, baris)

        If Len(namaBaris) > 0 Then

            PerbaruiKonteksPPh21 namaBaris, konteks
            kunci = KunciNamaPPh21( _
                namaBaris, vbNullString, konteks)

            If Len(kunci) > 0 _
               And petaSumber.Exists(kunci) Then

                If petaTujuan.Exists(kunci) Then
                    pesanError = _
                        "Nama baris tujuan PPh 21 ditemukan lebih dari " & _
                        "satu kali: " & NamaKunciUntukPesan(kunci)
                    GoTo GagalTanpaError
                End If

                dataSumber = petaSumber(kunci)

                If IsError(dataSumber(0)) Then
                    pesanError = _
                        "Nilai sumber EWS mengalami error pada baris '" & _
                        NamaKunciUntukPesan(kunci) & "'."
                    GoTo GagalTanpaError
                End If

                jumlahMapping = jumlahMapping + 1
                daftarBaris(jumlahMapping) = baris
                daftarNilaiBaru(jumlahMapping) = dataSumber(0)
                daftarNilaiLama(jumlahMapping) = _
                    wsLokal.Cells(baris, kolomTujuan).Formula

                petaTujuan.Add kunci, baris

            End If

        End If

    Next baris

    If jumlahMapping = 0 Then
        pesanError = _
            "Tidak ada nama baris EWS yang cocok dengan sheet PPh 21. " & _
            "Periksa nama baris dan blok Menurut Wajib Pajak."
        GoTo GagalTanpaError
    End If

    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    penulisanDimulai = True

    For i = 1 To jumlahMapping
        wsLokal.Cells( _
            daftarBaris(i), kolomTujuan).Value2 = _
            daftarNilaiBaru(i)
    Next i

    ImportPPh21DariWorkbook = True
    Exit Function

Gagal:

    pesanError = Err.Description

GagalTanpaError:

    If penulisanDimulai _
       And jumlahMapping > 0 _
       And Not ImportPPh21DariWorkbook Then

        On Error Resume Next

        For i = 1 To jumlahMapping
            If daftarBaris(i) > 0 Then
                wsLokal.Cells( _
                    daftarBaris(i), kolomTujuan).Formula = _
                    daftarNilaiLama(i)
            End If
        Next i

        On Error GoTo 0

    End If

End Function

'Tujuan: Membangun peta sumber pph 21.
Private Function BangunPetaSumberPPh21( _
    ByVal ws As Worksheet, _
    ByRef peta As Object, _
    ByRef pesanError As String) As Boolean

    Dim baris As Long
    Dim barisAkhir As Long
    Dim kolomNilai As Long
    Dim namaBaris As String
    Dim identifierBaris As String
    Dim konteks As String
    Dim kunci As String
    Dim nilai As Variant

    BangunPetaSumberPPh21 = False
    Set peta = CreateObject("Scripting.Dictionary")
    peta.CompareMode = vbTextCompare

    kolomNilai = CariKolomNilaiWPEWS(ws, pesanError)
    If kolomNilai = 0 Then Exit Function

    barisAkhir = BarisTerakhirTerpakai(ws)
    konteks = vbNullString

    For baris = 1 To barisAkhir

        namaBaris = AmbilNamaEksplisitPPh21EWS(ws, baris)

        If Len(namaBaris) > 0 Then
            PerbaruiKonteksPPh21 namaBaris, konteks
        End If

        identifierBaris = UCase$(Trim$(NilaiTeksAman( _
            ws.Cells(baris, KOLOM_ID_PPH21_EWS).Value2)))

        If Left$(identifierBaris, 4) = "NIL_" _
           And Len(namaBaris) > 0 Then

            kunci = KunciNamaPPh21( _
                namaBaris, identifierBaris, konteks)

            If Len(kunci) > 0 Then

                If peta.Exists(kunci) Then
                    pesanError = _
                        "Nama baris sumber PPh 21 ditemukan lebih dari " & _
                        "satu kali: " & NamaKunciUntukPesan(kunci)
                    Exit Function
                End If

                nilai = ws.Cells(baris, kolomNilai).Value
                peta.Add kunci, Array(nilai, baris, identifierBaris)

            End If

        End If

    Next baris

    If peta.Count = 0 Then
        pesanError = _
            "Tidak ditemukan baris data PPh 21 pada MAIN LEVEL EWS."
        Exit Function
    End If

    BangunPetaSumberPPh21 = True

End Function

'Tujuan: Mencari kolom nilai wpews.
Private Function CariKolomNilaiWPEWS( _
    ByVal ws As Worksheet, _
    ByRef pesanError As String) As Long

    Dim ditemukan As Range

    Set ditemukan = ws.Cells.Find( _
        What:="Menurut Wajib Pajak", _
        After:=ws.Cells(ws.Rows.Count, ws.Columns.Count), _
        LookIn:=xlValues, _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, _
        MatchCase:=False)

    If ditemukan Is Nothing Then
        pesanError = _
            "Header 'Menurut Wajib Pajak' tidak ditemukan pada MAIN LEVEL."
        Exit Function
    End If

    CariKolomNilaiWPEWS = ditemukan.Column

End Function

'Tujuan: Mengambil nama eksplisit pph 21 ews.
Private Function AmbilNamaEksplisitPPh21EWS( _
    ByVal ws As Worksheet, _
    ByVal baris As Long) As String

    Dim c As Long
    Dim teks As String

    For c = KOLOM_NAMA_PPH21_EWS_AKHIR _
        To KOLOM_NAMA_PPH21_EWS_AWAL Step -1

        teks = NilaiTeksAman(ws.Cells(baris, c).Value2)

        If NamaBarisDapatDipakai(teks) Then
            AmbilNamaEksplisitPPh21EWS = teks
            Exit Function
        End If

    Next c

End Function

'Tujuan: Mengambil nama eksplisit pph 21 kkp.
Private Function AmbilNamaEksplisitPPh21KKP( _
    ByVal ws As Worksheet, _
    ByVal baris As Long) As String

    Dim c As Long
    Dim teks As String

    For c = KOLOM_NAMA_PPH21_KKP_AKHIR _
        To KOLOM_NAMA_PPH21_KKP_AWAL Step -1

        teks = NilaiTeksAman(ws.Cells(baris, c).Value2)

        If NamaBarisDapatDipakai(teks) Then
            AmbilNamaEksplisitPPh21KKP = teks
            Exit Function
        End If

    Next c

End Function

'Tujuan: Memperbarui konteks pph 21.
Private Sub PerbaruiKonteksPPh21( _
    ByVal namaBaris As String, _
    ByRef konteks As String)

    Dim namaNormal As String

    namaNormal = NormalisasiNamaBaris(namaBaris)

    If InStr(1, namaNormal, _
            "PENGHASILANKENAPAJAKDASARPENGENAANPAJAK", _
            vbTextCompare) > 0 Then

        konteks = "DPP"

    ElseIf namaNormal = "PPHTERUTANG" Then

        konteks = "PPH"

    ElseIf namaNormal = "DIPERHITUNGKAN" Then

        konteks = "O"

    ElseIf InStr(1, namaNormal, _
            "PPHYANGSEHARUSNYATIDAKTERUTANG", _
            vbTextCompare) > 0 Then

        konteks = "P"

    End If

End Sub

'Tujuan: Membentuk kunci nama pph 21.
Private Function KunciNamaPPh21( _
    ByVal namaBaris As String, _
    ByVal identifierEWS As String, _
    ByVal konteks As String) As String

    Dim namaNormal As String
    Dim pembeda As String

    namaNormal = NormalisasiNamaBaris(namaBaris)
    If Len(namaNormal) = 0 Then Exit Function

    If InStr(1, namaNormal, _
            "JUMLAHPAJAKYANGDAPATDIKREDITKAN", _
            vbTextCompare) = 1 Then
        namaNormal = "JUMLAHPAJAKYANGDAPATDIKREDITKAN"
    End If

    Select Case namaNormal

        Case "KJS"

            If InStr(1, identifierEWS, _
                    "NIL_DPP_KJS", vbTextCompare) > 0 _
               Or UCase$(konteks) = "DPP" Then

                pembeda = "DPP"
            Else
                pembeda = "PPH"
            End If

        Case "JUMLAHA+B+C"

            If InStr(1, identifierEWS, _
                    "NIL_DIPERHITUNGKAN", vbTextCompare) > 0 _
               Or UCase$(konteks) = "O" Then

                pembeda = "O"
            Else
                pembeda = "P"
            End If

    End Select

    KunciNamaPPh21 = _
        "PPH21|" & namaNormal & "|" & pembeda

End Function

'Bagian: handler PPh 21 Final.
'Tujuan: Mengimpor MAIN LEVEL dan rincian B-1-3 secara transaksional.
Private Function ImportPPh21FinalDariWorkbook( _
    ByVal wbEWS As Workbook, _
    ByVal wsEWS As Worksheet, _
    ByVal bulanPajak As Long, _
    ByVal tahunPajak As Long, _
    ByRef pesanError As String) As Boolean

    Dim wsRingkasan As Worksheet
    Dim wsRincian As Worksheet
    Dim petaSumber As Object
    Dim petaRincian As Object
    Dim petaTujuan As Object
    Dim daftarSel As Collection
    Dim daftarNilaiBaru As Collection
    Dim daftarNilaiLama As Collection
    Dim selTujuan As Range
    Dim dataSumber As Variant
    Dim kodeObjek As Variant
    Dim kode As Variant
    Dim namaBaris As String
    Dim konteks As String
    Dim kunci As String
    Dim barisAwal As Long
    Dim barisAkhir As Long
    Dim baris As Long
    Dim barisObjek As Long
    Dim kolomRingkasan As Long
    Dim kolomRincian As Long
    Dim i As Long
    Dim penulisanDimulai As Boolean

    ImportPPh21FinalDariWorkbook = False
    On Error GoTo Gagal

    If bulanPajak < 1 Or bulanPajak > 12 Then
        Err.Raise vbObjectError + 3350, , "Nomor bulan tidak valid."
    End If

    If tahunPajak < 2000 Or tahunPajak > 2100 Then
        Err.Raise vbObjectError + 3351, , "Tahun pajak tidak valid."
    End If

    If Not WorksheetAda( _
            NamaSheetRingkasanPajak(TAXKEY_PPH21F), _
            ThisWorkbook) Then

        Err.Raise vbObjectError + 3352, , _
            "Sheet lokal '" & _
            NamaSheetRingkasanPajak(TAXKEY_PPH21F) & _
            "' tidak ditemukan."
    End If

    If Not SheetLokalKKPAda( _
            TAXKEY_PPH21F, "B-1-3", False) Then

        Err.Raise vbObjectError + 3353, , _
            "Sheet lokal 'P21F-B-1-3' tidak ditemukan."
    End If

    If Not WorksheetAda( _
            NamaSheetIndukPajak(TAXKEY_PPH21F), _
            ThisWorkbook) Then

        Err.Raise vbObjectError + 3354, , _
            "Sheet lokal '" & _
            NamaSheetIndukPajak(TAXKEY_PPH21F) & _
            "' tidak ditemukan."
    End If

    Set wsRingkasan = ThisWorkbook.Worksheets( _
        NamaSheetRingkasanPajak(TAXKEY_PPH21F))
    Set wsRincian = DapatkanSheetLokalKKP( _
        TAXKEY_PPH21F, "B-1-3", False)

    kolomRingkasan = CariKolomPeriodePPN( _
        wsRingkasan, bulanPajak, tahunPajak, pesanError)
    If kolomRingkasan = 0 Then GoTo GagalTanpaError

    kolomRincian = bulanPajak + 3
    If Val(wsRincian.Cells(21, kolomRincian).Value2) <> bulanPajak Then
        pesanError = _
            "Kolom masa " & CStr(bulanPajak) & _
            " tidak ditemukan pada P21F-B-1-3."
        GoTo GagalTanpaError
    End If

    If Not BangunPetaSumberPPh21Final( _
            wsEWS, petaSumber, pesanError) Then
        GoTo GagalTanpaError
    End If

    If Not BacaRincianB13PPh21Final( _
            wbEWS, petaRincian, pesanError) Then
        GoTo GagalTanpaError
    End If

    If Not CariBlokMenurutWP( _
            wsRingkasan, barisAwal, barisAkhir, pesanError) Then
        GoTo GagalTanpaError
    End If

    Set petaTujuan = CreateObject("Scripting.Dictionary")
    petaTujuan.CompareMode = vbTextCompare
    Set daftarSel = New Collection
    Set daftarNilaiBaru = New Collection
    Set daftarNilaiLama = New Collection

    konteks = vbNullString

    For baris = barisAwal To barisAkhir

        namaBaris = AmbilNamaEksplisitPPh21KKP( _
            wsRingkasan, baris)

        If Len(namaBaris) > 0 Then

            PerbaruiKonteksPPh21 namaBaris, konteks
            kunci = KunciNamaPPh21( _
                namaBaris, vbNullString, konteks)

            If Len(kunci) > 0 _
               And petaSumber.Exists(kunci) Then

                If petaTujuan.Exists(kunci) Then
                    pesanError = _
                        "Baris tujuan PPh 21 Final ditemukan lebih " & _
                        "dari satu kali: " & NamaKunciUntukPesan(kunci)
                    GoTo GagalTanpaError
                End If

                dataSumber = petaSumber(kunci)

                If Not NilaiNumerikSiapPPh21Final( _
                        dataSumber(0), kunci, pesanError) Then
                    GoTo GagalTanpaError
                End If

                Set selTujuan = wsRingkasan.Cells( _
                    baris, kolomRingkasan)

                If Not TambahRencanaPPh21Final( _
                        selTujuan, CDbl(dataSumber(0)), _
                        daftarSel, daftarNilaiBaru, _
                        daftarNilaiLama, pesanError) Then
                    GoTo GagalTanpaError
                End If

                petaTujuan.Add kunci, baris

            End If

        End If

    Next baris

    If petaTujuan.Count = 0 Then
        pesanError = _
            "Tidak ada baris MAIN LEVEL PPh 21 Final yang cocok " & _
            "dengan sheet lokal."
        GoTo GagalTanpaError
    End If

    For Each kode In petaSumber.Keys
        If Not petaTujuan.Exists(CStr(kode)) Then
            pesanError = _
                "Baris lokal PPh 21 Final belum terpetakan: " & _
                NamaKunciUntukPesan(CStr(kode))
            GoTo GagalTanpaError
        End If
    Next kode

    kodeObjek = Array( _
        "21-100-01", "21-100-02", _
        "21-100-03", "21-100-07")

    For Each kode In kodeObjek

        barisObjek = CariBarisObjekPPh21Final( _
            wsRincian, CStr(kode), pesanError)
        If barisObjek = 0 Then GoTo GagalTanpaError

        kunci = CStr(kode) & "|DPP"
        dataSumber = petaRincian(kunci)

        If Not TambahRencanaPPh21Final( _
                wsRincian.Cells(barisObjek, kolomRincian), _
                CDbl(dataSumber(0)), daftarSel, _
                daftarNilaiBaru, daftarNilaiLama, pesanError) Then
            GoTo GagalTanpaError
        End If

        kunci = CStr(kode) & "|PPH"
        dataSumber = petaRincian(kunci)

        If Not TambahRencanaPPh21Final( _
                wsRincian.Cells(barisObjek + 1, kolomRincian), _
                CDbl(dataSumber(0)), daftarSel, _
                daftarNilaiBaru, daftarNilaiLama, pesanError) Then
            GoTo GagalTanpaError
        End If

    Next kode

    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    penulisanDimulai = True

    For i = 1 To daftarSel.Count
        Set selTujuan = daftarSel(i)
        selTujuan.Value2 = daftarNilaiBaru(i)
    Next i

    HitungUlangKKPLokal TAXKEY_PPH21F
    ImportPPh21FinalDariWorkbook = True
    Exit Function

Gagal:

    pesanError = Err.Description

GagalTanpaError:

    If penulisanDimulai _
       And Not ImportPPh21FinalDariWorkbook Then

        If Not daftarSel Is Nothing Then

            On Error Resume Next

            For i = 1 To daftarSel.Count
                Set selTujuan = daftarSel(i)
                selTujuan.Formula = daftarNilaiLama(i)
            Next i

            On Error GoTo 0

        End If

    End If

End Function

'Tujuan: Membangun peta lengkap MAIN LEVEL PPh 21 Final.
Private Function BangunPetaSumberPPh21Final( _
    ByVal ws As Worksheet, _
    ByRef peta As Object, _
    ByRef pesanError As String) As Boolean

    Dim daftarID As Object
    Dim baris As Long
    Dim barisAkhir As Long
    Dim kolomNilai As Long
    Dim namaBaris As String
    Dim identifierBaris As String
    Dim konteks As String
    Dim kunci As String
    Dim nilai As Variant

    BangunPetaSumberPPh21Final = False
    Set peta = CreateObject("Scripting.Dictionary")
    peta.CompareMode = vbTextCompare
    Set daftarID = CreateObject("Scripting.Dictionary")
    daftarID.CompareMode = vbTextCompare

    kolomNilai = CariKolomNilaiWPEWS(ws, pesanError)
    If kolomNilai = 0 Then Exit Function

    barisAkhir = BarisTerakhirTerpakai(ws)
    konteks = vbNullString

    For baris = 1 To barisAkhir

        namaBaris = AmbilNamaEksplisitPPh21EWS(ws, baris)

        If Len(namaBaris) > 0 Then
            PerbaruiKonteksPPh21 namaBaris, konteks
        End If

        identifierBaris = UCase$(Trim$(NilaiTeksAman( _
            ws.Cells(baris, KOLOM_ID_PPH21_EWS).Value2)))

        If (Left$(identifierBaris, 4) = "NIL_" _
            Or Left$(identifierBaris, 11) = "KOMPENSASI_") _
           And Len(namaBaris) > 0 Then

            kunci = KunciNamaPPh21( _
                namaBaris, identifierBaris, konteks)

            If Len(kunci) > 0 Then

                If peta.Exists(kunci) Then
                    pesanError = _
                        "Nama baris sumber PPh 21 Final ditemukan " & _
                        "lebih dari satu kali: " & _
                        NamaKunciUntukPesan(kunci)
                    Exit Function
                End If

                If daftarID.Exists(identifierBaris) Then
                    pesanError = _
                        "Identifier sumber PPh 21 Final ganda: " & _
                        identifierBaris
                    Exit Function
                End If

                nilai = ws.Cells(baris, kolomNilai).Value2

                If Not NilaiNumerikSiapPPh21Final( _
                        nilai, identifierBaris, pesanError) Then
                    Exit Function
                End If

                daftarID.Add identifierBaris, baris
                peta.Add _
                    kunci, Array(CDbl(nilai), baris, identifierBaris)

            End If

        End If

    Next baris

    If peta.Count = 0 Then
        pesanError = _
            "Tidak ditemukan data PPh 21 Final pada MAIN LEVEL EWS."
        Exit Function
    End If

    If Not ValidasiIdentifierPPh21Final( _
            daftarID, pesanError) Then
        Exit Function
    End If

    BangunPetaSumberPPh21Final = True

End Function

'Tujuan: Memastikan struktur utama EWS PPh 21 Final lengkap.
Private Function ValidasiIdentifierPPh21Final( _
    ByVal daftarID As Object, _
    ByRef pesanError As String) As Boolean

    Dim wajib As Variant
    Dim item As Variant

    wajib = Array( _
        "NIL_DPP_KJS_411121_100_FINAL", _
        "NIL_TOTAL_DPP_PASAL_21_FINAL", _
        "NIL_PPH_KJS_411121_100_FINAL", _
        "NIL_TOTAL_PPH_PASAL21_FINAL", _
        "NIL_PH_TDK_TERHUTANG_PPH_PASAL_17_AYAT_2", _
        "NIL_PPH_HARUS_DIPUNGUT_DISETOR", _
        "NIL_KREDIT_PAJAK_DITANGGUNG_PEMERINTAH", _
        "NIL_KREDIT_PAJAK_SETORAN_MASA", _
        "NIL_KREDIT_PAJAK_STP", _
        "KOMPENSASI_DARI_MASA_PAJAK", _
        "NIL_KREDIT_PAJAK_LAIN", _
        "NIL_KREDIT_PAJAK_UANG_TEBUSAN", _
        "NIL_KREDIT_PAJAK_UNGKAP_HARTA_BERSIH", _
        "KOMPENSASI_KE_MASA_PAJAK", _
        "NIL_DIPERHITUNGKAN_SKPKB", _
        "NIL_DIPERHITUNGKAN_SKPKBT", _
        "NIL_DIPERHITUNGKAN_SKPLB", _
        "NIL_DIPERHITUNGKAN", _
        "NIL_PPH_PASAL17_DIBAYAR_NPWP_SENDIRI", _
        "NIL_PPH_PASAL17_DIBAYAR_NPWP_LAIN", _
        "NIL_PPH_PASAL17_TELAH_POTPUT", _
        "NIL_PPH_TDK_TERUTANG_PASAL17", _
        "NIL_KREDIT_PAJAK", _
        "NIL_KLB")

    For Each item In wajib

        If Not daftarID.Exists(CStr(item)) Then
            pesanError = _
                "Struktur PPh 21 Final belum lengkap: " & _
                CStr(item) & " tidak ditemukan pada MAIN LEVEL."
            Exit Function
        End If

    Next item

    ValidasiIdentifierPPh21Final = True

End Function

'Tujuan: Menjumlahkan empat objek pajak langsung dari detail EWS B-1-3-1.
Private Function BacaRincianB13PPh21Final( _
    ByVal wbEWS As Workbook, _
    ByRef peta As Object, _
    ByRef pesanError As String) As Boolean

    Dim ws As Worksheet
    Dim kodeObjek As Variant
    Dim kode As Variant
    Dim baris As Long
    Dim barisAkhir As Long
    Dim barisSumber As Long
    Dim kodeBaris As String
    Dim nilaiDPP As Double
    Dim nilaiPPh As Double
    Dim nilai As Variant

    BacaRincianB13PPh21Final = False
    Set peta = CreateObject("Scripting.Dictionary")
    peta.CompareMode = vbTextCompare

    If Not WorksheetAda("B-1-3-1", wbEWS) Then
        pesanError = "Sheet 'B-1-3-1' tidak ditemukan pada EWS PPh 21 Final."
        Exit Function
    End If

    Set ws = wbEWS.Worksheets("B-1-3-1")

    On Error Resume Next
    ws.Calculate
    DoEvents
    On Error GoTo 0

    barisAkhir = ws.Cells(ws.Rows.Count, 47).End(xlUp).Row
    kodeObjek = Array( _
        "21-100-01", "21-100-02", _
        "21-100-03", "21-100-07")

    For Each kode In kodeObjek

        nilaiDPP = 0
        nilaiPPh = 0
        barisSumber = 0

        For baris = 1 To barisAkhir

            kodeBaris = Trim$(Replace$( _
                CStr(ws.Cells(baris, 47).Value2), ChrW$(160), " "))

            If StrComp(kodeBaris, CStr(kode), vbTextCompare) = 0 Then

                If barisSumber = 0 Then barisSumber = baris

                nilai = ws.Cells(baris, 64).Value2
                If Not NilaiNumerikSiapPPh21Final( _
                        nilai, CStr(kode) & " / DPP pada B-1-3-1", _
                        pesanError) Then
                    Exit Function
                End If
                nilaiDPP = nilaiDPP + CDbl(nilai)

                nilai = ws.Cells(baris, 76).Value2
                If Not NilaiNumerikSiapPPh21Final( _
                        nilai, CStr(kode) & " / PPh pada B-1-3-1", _
                        pesanError) Then
                    Exit Function
                End If
                nilaiPPh = nilaiPPh + CDbl(nilai)

            End If

        Next baris

        peta.Add _
            CStr(kode) & "|DPP", _
            Array(nilaiDPP, barisSumber, "B-1-3-1")
        peta.Add _
            CStr(kode) & "|PPH", _
            Array(nilaiPPh, barisSumber, "B-1-3-1")

    Next kode

    BacaRincianB13PPh21Final = True

End Function

'Tujuan: Memastikan nilai sumber siap ditulis.
Private Function NilaiNumerikSiapPPh21Final( _
    ByVal nilai As Variant, _
    ByVal namaNilai As String, _
    ByRef pesanError As String) As Boolean

    If IsError(nilai) Or IsNull(nilai) Or IsEmpty(nilai) Then
        pesanError = _
            "Nilai PPh 21 Final belum siap: " & namaNilai & "."
        Exit Function
    End If

    If Not IsNumeric(nilai) Then
        pesanError = _
            "Nilai PPh 21 Final bukan angka: " & namaNilai & "."
        Exit Function
    End If

    NilaiNumerikSiapPPh21Final = True

End Function

'Tujuan: Menemukan pasangan baris DPP dan PPh suatu objek lokal.
Private Function CariBarisObjekPPh21Final( _
    ByVal ws As Worksheet, _
    ByVal kodeObjek As String, _
    ByRef pesanError As String) As Long

    Dim baris As Long
    Dim barisAkhir As Long
    Dim teks As String

    barisAkhir = BarisTerakhirTerpakai(ws)

    For baris = 1 To barisAkhir

        teks = UCase$(NilaiTeksAman(ws.Cells(baris, 2).Value2))

        If InStr(1, teks, UCase$(kodeObjek), vbTextCompare) > 0 Then

            If NormalisasiNamaBaris( _
                    NilaiTeksAman(ws.Cells(baris, 3).Value2)) <> "DPP" _
               Or NormalisasiNamaBaris( _
                    NilaiTeksAman(ws.Cells(baris + 1, 3).Value2)) <> "PPH" Then

                pesanError = _
                    "Pasangan baris DPP/PPh tidak sesuai untuk " & _
                    kodeObjek & " pada P21F-B-1-3."
                Exit Function
            End If

            CariBarisObjekPPh21Final = baris
            Exit Function

        End If

    Next baris

    pesanError = _
        "Kode objek " & kodeObjek & _
        " tidak ditemukan pada P21F-B-1-3."

End Function

'Tujuan: Menyimpan target, nilai baru, dan nilai pemulihan.
Private Function TambahRencanaPPh21Final( _
    ByVal selTujuan As Range, _
    ByVal nilaiBaru As Double, _
    ByVal daftarSel As Collection, _
    ByVal daftarNilaiBaru As Collection, _
    ByVal daftarNilaiLama As Collection, _
    ByRef pesanError As String) As Boolean

    If selTujuan.MergeCells Then
        pesanError = _
            "Sel tujuan PPh 21 Final tergabung: " & _
            selTujuan.Parent.Name & "!" & _
            selTujuan.Address(False, False)
        Exit Function
    End If

    If selTujuan.Parent.ProtectContents Then
        If selTujuan.Locked Then
            pesanError = _
                "Sel tujuan PPh 21 Final terkunci: " & _
                selTujuan.Parent.Name & "!" & _
                selTujuan.Address(False, False)
            Exit Function
        End If
    End If

    daftarSel.Add selTujuan
    daftarNilaiBaru.Add nilaiBaru
    daftarNilaiLama.Add selTujuan.Formula
    TambahRencanaPPh21Final = True

End Function

'Bagian: handler PPh 23.
'Tujuan: Mengimpor pph 23 dari workbook.
Private Function ImportPPh23DariWorkbook( _
    ByVal wsEWS As Worksheet, _
    ByVal bulanPajak As Long, _
    ByVal tahunPajak As Long, _
    ByRef pesanError As String) As Boolean

    Dim wsLokal As Worksheet
    Dim petaSumber As Object
    Dim petaTujuan As Object

    Dim namaSheetRingkasan As String
    Dim kolomTujuan As Long
    Dim barisAwal As Long
    Dim barisAkhir As Long
    Dim baris As Long
    Dim jumlahMapping As Long

    Dim namaBaris As String
    Dim konteks As String
    Dim kunci As String
    Dim dataSumber As Variant

    Dim daftarBaris() As Long
    Dim daftarNilaiBaru() As Variant
    Dim daftarNilaiLama() As Variant
    Dim i As Long
    Dim penulisanDimulai As Boolean
    Dim kunciSumber As Variant
    Dim daftarKKP As Variant
    Dim kodeDetail As Variant
    Dim selTujuan As Range

    ImportPPh23DariWorkbook = False

    On Error GoTo Gagal

    If bulanPajak < 1 Or bulanPajak > 12 Then
        Err.Raise vbObjectError + 3300, , "Nomor bulan tidak valid."
    End If

    If tahunPajak < 2000 Or tahunPajak > 2100 Then
        Err.Raise vbObjectError + 3301, , "Tahun pajak tidak valid."
    End If

    namaSheetRingkasan = _
        NamaSheetRingkasanPajak(TAXKEY_PPH23)

    If Not WorksheetAda(namaSheetRingkasan, ThisWorkbook) Then
        Err.Raise _
            vbObjectError + 3302, , _
            "Sheet lokal '" & namaSheetRingkasan & _
            "' tidak ditemukan."
    End If

    Set wsLokal = _
        ThisWorkbook.Worksheets(namaSheetRingkasan)

    daftarKKP = DaftarKodeEWSPPh23()
    For Each kodeDetail In daftarKKP
        If Not SheetLokalKKPAda(TAXKEY_PPH23, CStr(kodeDetail), False) Then
            Err.Raise vbObjectError + 3320, , _
                "Sheet detail PPh 23 belum tersedia: " & _
                NamaSheetLokalKKP(TAXKEY_PPH23, CStr(kodeDetail))
        End If
    Next kodeDetail
    If Not WorksheetAda(NamaSheetIndukPajak(TAXKEY_PPH23), ThisWorkbook) Then
        Err.Raise vbObjectError + 3321, , _
            "Sheet '" & NamaSheetIndukPajak(TAXKEY_PPH23) & "' belum tersedia."
    End If

    kolomTujuan = CariKolomPeriodePPN( _
        wsLokal, bulanPajak, tahunPajak, pesanError)

    If kolomTujuan = 0 Then GoTo GagalTanpaError

    If Not BangunPetaSumberPPh23( _
            wsEWS, petaSumber, pesanError) Then
        GoTo GagalTanpaError
    End If

    If Not CariBlokMenurutWP( _
            wsLokal, barisAwal, barisAkhir, pesanError) Then
        GoTo GagalTanpaError
    End If

    Set petaTujuan = CreateObject("Scripting.Dictionary")
    petaTujuan.CompareMode = vbTextCompare

    ReDim daftarBaris(1 To barisAkhir - barisAwal + 1)
    ReDim daftarNilaiBaru(1 To barisAkhir - barisAwal + 1)
    ReDim daftarNilaiLama(1 To barisAkhir - barisAwal + 1)

    konteks = vbNullString

    For baris = barisAwal To barisAkhir

        namaBaris = AmbilNamaEksplisitPPhPemotonganKKP(wsLokal, baris)

        If Len(namaBaris) > 0 Then

            PerbaruiKonteksPPhPemotongan namaBaris, konteks
            kunci = KunciNamaPPh23( _
                namaBaris, vbNullString, konteks)

            If Len(kunci) > 0 _
               And petaSumber.Exists(kunci) Then

                If petaTujuan.Exists(kunci) Then
                    pesanError = _
                        "Nama baris tujuan PPh 23 ditemukan lebih dari " & _
                        "satu kali: " & NamaKunciUntukPesan(kunci)
                    GoTo GagalTanpaError
                End If

                dataSumber = petaSumber(kunci)

                If IsError(dataSumber(0)) Then
                    pesanError = _
                        "Nilai sumber EWS mengalami error pada baris '" & _
                        NamaKunciUntukPesan(kunci) & "'."
                    GoTo GagalTanpaError
                End If

                If IsEmpty(dataSumber(0)) Or IsNull(dataSumber(0)) Then
                    pesanError = "Nilai sumber PPh 23 kosong: " & _
                        NamaKunciUntukPesan(kunci)
                    GoTo GagalTanpaError
                End If
                If Not IsNumeric(dataSumber(0)) Then
                    pesanError = "Nilai sumber PPh 23 bukan angka: " & _
                        NamaKunciUntukPesan(kunci)
                    GoTo GagalTanpaError
                End If
                Set selTujuan = wsLokal.Cells(baris, kolomTujuan)
                If selTujuan.MergeCells Then
                    pesanError = "Sel nilai WP PPh 23 tergabung: " & _
                        selTujuan.Address(False, False)
                    GoTo GagalTanpaError
                End If
                If wsLokal.ProtectContents Then
                    If selTujuan.Locked Then
                        pesanError = "Sel nilai WP PPh 23 terkunci: " & _
                            selTujuan.Address(False, False)
                        GoTo GagalTanpaError
                    End If
                End If

                jumlahMapping = jumlahMapping + 1
                daftarBaris(jumlahMapping) = baris
                daftarNilaiBaru(jumlahMapping) = CDbl(dataSumber(0))
                daftarNilaiLama(jumlahMapping) = _
                    wsLokal.Cells(baris, kolomTujuan).Formula

                petaTujuan.Add kunci, baris

            End If

        End If

    Next baris

    If jumlahMapping = 0 Then
        pesanError = _
            "Tidak ada nama baris EWS yang cocok dengan sheet PPh 23. " & _
            "Periksa nama baris dan blok Menurut Wajib Pajak."
        GoTo GagalTanpaError
    End If

    For Each kunciSumber In petaSumber.Keys
        If Not petaTujuan.Exists(CStr(kunciSumber)) Then
            pesanError = "Baris WP lokal PPh 23 belum terpetakan: " & _
                NamaKunciUntukPesan(CStr(kunciSumber))
            GoTo GagalTanpaError
        End If
    Next kunciSumber

    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    penulisanDimulai = True

    For i = 1 To jumlahMapping
        wsLokal.Cells( _
            daftarBaris(i), kolomTujuan).Value2 = _
            daftarNilaiBaru(i)
    Next i

    HitungUlangKKPLokal TAXKEY_PPH23
    ImportPPh23DariWorkbook = True
    Exit Function

Gagal:

    pesanError = Err.Description

GagalTanpaError:

    If penulisanDimulai _
       And jumlahMapping > 0 _
       And Not ImportPPh23DariWorkbook Then

        On Error Resume Next

        For i = 1 To jumlahMapping
            If daftarBaris(i) > 0 Then
                wsLokal.Cells( _
                    daftarBaris(i), kolomTujuan).Formula = _
                    daftarNilaiLama(i)
            End If
        Next i

        On Error GoTo 0

    End If

End Function

'Tujuan: Membangun peta sumber pph 23.
Private Function BangunPetaSumberPPh23( _
    ByVal ws As Worksheet, _
    ByRef peta As Object, _
    ByRef pesanError As String) As Boolean

    Dim baris As Long
    Dim barisAkhir As Long
    Dim kolomNilai As Long
    Dim namaBaris As String
    Dim identifierBaris As String
    Dim konteks As String
    Dim kunci As String
    Dim nilai As Variant
    Dim daftarID As Object

    BangunPetaSumberPPh23 = False
    Set peta = CreateObject("Scripting.Dictionary")
    peta.CompareMode = vbTextCompare
    Set daftarID = CreateObject("Scripting.Dictionary")
    daftarID.CompareMode = vbTextCompare

    kolomNilai = CariKolomNilaiWPPPh23(ws, pesanError)
    If kolomNilai = 0 Then Exit Function

    barisAkhir = BarisTerakhirTerpakai(ws)
    konteks = vbNullString

    For baris = 1 To barisAkhir

        namaBaris = AmbilNamaEksplisitPPhPemotonganEWS(ws, baris)

        If Len(namaBaris) > 0 Then
            PerbaruiKonteksPPhPemotongan namaBaris, konteks
        End If

        identifierBaris = UCase$(Trim$(NilaiTeksAman( _
            ws.Cells(baris, KOLOM_ID_PPH21_EWS).Value2)))

        If Left$(identifierBaris, 4) = "NIL_" _
           And Len(namaBaris) > 0 Then

            kunci = KunciNamaPPh23( _
                namaBaris, identifierBaris, konteks)

            If Len(kunci) > 0 Then

                If peta.Exists(kunci) Then
                    pesanError = _
                        "Nama baris sumber PPh 23 ditemukan lebih dari " & _
                        "satu kali: " & NamaKunciUntukPesan(kunci)
                    Exit Function
                End If

                If daftarID.Exists(identifierBaris) Then
                    pesanError = "Identifier sumber PPh 23 ganda: " & identifierBaris
                    Exit Function
                End If
                daftarID.Add identifierBaris, baris
                nilai = ws.Cells(baris, kolomNilai).Value2
                peta.Add kunci, Array(nilai, baris, identifierBaris)

            End If

        End If

    Next baris

    If peta.Count = 0 Then
        pesanError = _
            "Tidak ditemukan baris data PPh 23 pada MAIN LEVEL EWS."
        Exit Function
    End If

    If Not ValidasiIdentifierPPh23(daftarID, pesanError) Then Exit Function
    BangunPetaSumberPPh23 = True

End Function

'Tujuan: Mencari kolom nilai wppph 23.
Private Function CariKolomNilaiWPPPh23( _
    ByVal ws As Worksheet, _
    ByRef pesanError As String) As Long

    Dim ditemukan As Range
    Dim alamatAwal As String
    Dim kolomWP As Long
    Dim namaHeader As String

    Set ditemukan = ws.Cells.Find( _
        What:="Menurut Wajib Pajak", _
        After:=ws.Cells(ws.Rows.Count, ws.Columns.Count), _
        LookIn:=xlValues, LookAt:=xlPart, _
        SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False)

    If ditemukan Is Nothing Then
        pesanError = "Header Menurut Wajib Pajak PPh 23 tidak ditemukan."
        Exit Function
    End If
    alamatAwal = ditemukan.Address
    Do
        namaHeader = NormalisasiNamaBaris(NilaiTeksAman(ditemukan.Value2))
        If InStr(1, namaHeader, "MENURUTWAJIBPAJAK", vbTextCompare) = 1 Then
            If kolomWP <> 0 And kolomWP <> ditemukan.Column Then
                pesanError = "Kolom Menurut Wajib Pajak PPh 23 ambigu."
                Exit Function
            End If
            kolomWP = ditemukan.Column
        End If
        Set ditemukan = ws.Cells.FindNext(After:=ditemukan)
        If ditemukan Is Nothing Then Exit Do
    Loop While ditemukan.Address <> alamatAwal

    If kolomWP = 0 Then
        pesanError = "Header nilai WP PPh 23 tidak dikenali."
    Else
        CariKolomNilaiWPPPh23 = kolomWP
    End If
End Function

'Tujuan: Membentuk kunci nama pph 23.
Private Function KunciNamaPPh23( _
    ByVal namaBaris As String, _
    ByVal identifierEWS As String, _
    ByVal konteks As String) As String

    Dim namaNormal As String
    Dim pembeda As String

    namaNormal = NormalisasiNamaBaris(namaBaris)
    If Len(namaNormal) = 0 Then Exit Function

    If InStr(1, namaNormal, "JUMLAHPAJAKYANGDAPATDIKREDITKAN", _
            vbTextCompare) = 1 Then
        namaNormal = "JUMLAHPAJAKYANGDAPATDIKREDITKAN"
    End If

    Select Case namaNormal

        Case "KJS"

            If InStr(1, identifierEWS, _
                    "NIL_DPP_KJS", vbTextCompare) > 0 _
               Or UCase$(konteks) = "DPP" Then

                pembeda = "DPP"
            Else
                pembeda = "PPH"
            End If

        Case "JUMLAHA+B+C"

            If InStr(1, identifierEWS, _
                    "NIL_DIPERHITUNGKAN", vbTextCompare) > 0 _
               Or UCase$(konteks) = "O" Then

                pembeda = "O"
            Else
                pembeda = "P"
            End If

    End Select

    KunciNamaPPh23 = _
        "PPH23|" & namaNormal & "|" & pembeda

End Function

'Tujuan: Memvalidasi identifier pph 23.
Private Function ValidasiIdentifierPPh23( _
    ByVal daftarID As Object, _
    ByRef pesanError As String) As Boolean

    Dim wajib As Variant
    Dim item As Variant
    wajib = Array( _
        "NIL_DPP_KJS_411124_100", "NIL_TOTAL_DPP_PASAL_23", _
        "NIL_PPH_KJS_411124_100", "NIL_TOTAL_PPH_PASAL23", _
        "NIL_PH_TDK_TERHUTANG_PPH_PASAL_17_AYAT_2", "NIL_PPH_HARUS_DIPUNGUT_DISETOR", _
        "NIL_KREDIT_PAJAK_DITANGGUNG_PEMERINTAH", "NIL_KREDIT_PAJAK_SETORAN_MASA", _
        "NIL_KREDIT_PAJAK_STP", "NIL_KOMPENSASI_DARI_MASA_PAJAK", _
        "NIL_KREDIT_PAJAK_LAIN", "NIL_KREDIT_PAJAK_UANG_TEBUSAN", _
        "NIL_KREDIT_PAJAK_UNGKAP_HARTA_BERSIH", "NIL_KOMPENSASI_KE_MASA_PAJAK", _
        "NIL_DIPERHITUNGKAN_SKPKB", "NIL_DIPERHITUNGKAN_SKPKBT", _
        "NIL_DIPERHITUNGKAN_SKPLB", "NIL_DIPERHITUNGKAN", _
        "NIL_PPH_PASAL17_DIBAYAR_NPWP_SENDIRI", "NIL_PPH_PASAL17_DIBAYAR_NPWP_LAIN", _
        "NIL_PPH_PASAL17_TELAH_POTPUT", "NIL_PPH_TDK_TERUTANG_PASAL17", _
        "NIL_KREDIT_PAJAK", "NIL_KLB")
    For Each item In wajib
        If Not daftarID.Exists(CStr(item)) Then
            pesanError = "Struktur PPh 23 belum lengkap/tidak sesuai: " & _
                CStr(item) & " tidak ditemukan pada MAIN LEVEL."
            Exit Function
        End If
    Next item
    ValidasiIdentifierPPh23 = True
End Function

'Bagian: handler PPh 26.
'Tujuan: Mengimpor pph 26 dari workbook.
Private Function ImportPPh26DariWorkbook( _
    ByVal wsEWS As Worksheet, _
    ByVal bulanPajak As Long, _
    ByVal tahunPajak As Long, _
    ByRef pesanError As String) As Boolean

    Dim wsLokal As Worksheet
    Dim petaSumber As Object
    Dim petaTujuan As Object

    Dim namaSheetRingkasan As String
    Dim kolomTujuan As Long
    Dim barisAwal As Long
    Dim barisAkhir As Long
    Dim baris As Long
    Dim jumlahMapping As Long

    Dim namaBaris As String
    Dim konteks As String
    Dim kunci As String
    Dim dataSumber As Variant

    Dim daftarBaris() As Long
    Dim daftarNilaiBaru() As Variant
    Dim daftarNilaiLama() As Variant
    Dim i As Long
    Dim penulisanDimulai As Boolean
    Dim kunciSumber As Variant
    Dim daftarKKP As Variant
    Dim kodeDetail As Variant
    Dim selTujuan As Range

    ImportPPh26DariWorkbook = False

    On Error GoTo Gagal

    If bulanPajak < 1 Or bulanPajak > 12 Then
        Err.Raise vbObjectError + 3300, , "Nomor bulan tidak valid."
    End If

    If tahunPajak < 2000 Or tahunPajak > 2100 Then
        Err.Raise vbObjectError + 3301, , "Tahun pajak tidak valid."
    End If

    namaSheetRingkasan = _
        NamaSheetRingkasanPajak(TAXKEY_PPH26)

    If Not WorksheetAda(namaSheetRingkasan, ThisWorkbook) Then
        Err.Raise _
            vbObjectError + 3302, , _
            "Sheet lokal '" & namaSheetRingkasan & _
            "' tidak ditemukan."
    End If

    Set wsLokal = _
        ThisWorkbook.Worksheets(namaSheetRingkasan)

    daftarKKP = DaftarKodeEWSPPh26()
    For Each kodeDetail In daftarKKP
        If Not SheetLokalKKPAda(TAXKEY_PPH26, CStr(kodeDetail), False) Then
            Err.Raise vbObjectError + 3320, , _
                "Sheet detail PPh 26 belum tersedia: " & _
                NamaSheetLokalKKP(TAXKEY_PPH26, CStr(kodeDetail))
        End If
    Next kodeDetail
    If Not WorksheetAda(NamaSheetIndukPajak(TAXKEY_PPH26), ThisWorkbook) Then
        Err.Raise vbObjectError + 3321, , _
            "Sheet '" & NamaSheetIndukPajak(TAXKEY_PPH26) & "' belum tersedia."
    End If

    kolomTujuan = CariKolomPeriodePPN( _
        wsLokal, bulanPajak, tahunPajak, pesanError)

    If kolomTujuan = 0 Then GoTo GagalTanpaError

    If Not BangunPetaSumberPPh26( _
            wsEWS, petaSumber, pesanError) Then
        GoTo GagalTanpaError
    End If

    If Not CariBlokMenurutWP( _
            wsLokal, barisAwal, barisAkhir, pesanError) Then
        GoTo GagalTanpaError
    End If

    Set petaTujuan = CreateObject("Scripting.Dictionary")
    petaTujuan.CompareMode = vbTextCompare

    ReDim daftarBaris(1 To barisAkhir - barisAwal + 1)
    ReDim daftarNilaiBaru(1 To barisAkhir - barisAwal + 1)
    ReDim daftarNilaiLama(1 To barisAkhir - barisAwal + 1)

    konteks = vbNullString

    For baris = barisAwal To barisAkhir

        namaBaris = AmbilNamaEksplisitPPhPemotonganKKP(wsLokal, baris)

        If Len(namaBaris) > 0 Then

            PerbaruiKonteksPPhPemotongan namaBaris, konteks
            kunci = KunciNamaPPh26( _
                namaBaris, vbNullString, konteks, AmbilKJSLokalPPh26(wsLokal, baris))

            If Len(kunci) > 0 _
               And petaSumber.Exists(kunci) Then

                If petaTujuan.Exists(kunci) Then
                    pesanError = _
                        "Nama baris tujuan PPh 26 ditemukan lebih dari " & _
                        "satu kali: " & NamaKunciUntukPesan(kunci)
                    GoTo GagalTanpaError
                End If

                dataSumber = petaSumber(kunci)

                If IsError(dataSumber(0)) Then
                    pesanError = _
                        "Nilai sumber EWS mengalami error pada baris '" & _
                        NamaKunciUntukPesan(kunci) & "'."
                    GoTo GagalTanpaError
                End If

                If IsEmpty(dataSumber(0)) Or IsNull(dataSumber(0)) Then
                    pesanError = "Nilai sumber PPh 26 kosong: " & _
                        NamaKunciUntukPesan(kunci)
                    GoTo GagalTanpaError
                End If
                If Not IsNumeric(dataSumber(0)) Then
                    pesanError = "Nilai sumber PPh 26 bukan angka: " & _
                        NamaKunciUntukPesan(kunci)
                    GoTo GagalTanpaError
                End If
                Set selTujuan = wsLokal.Cells(baris, kolomTujuan)
                If selTujuan.MergeCells Then
                    pesanError = "Sel nilai WP PPh 26 tergabung: " & _
                        selTujuan.Address(False, False)
                    GoTo GagalTanpaError
                End If
                If wsLokal.ProtectContents Then
                    If selTujuan.Locked Then
                        pesanError = "Sel nilai WP PPh 26 terkunci: " & _
                            selTujuan.Address(False, False)
                        GoTo GagalTanpaError
                    End If
                End If

                jumlahMapping = jumlahMapping + 1
                daftarBaris(jumlahMapping) = baris
                daftarNilaiBaru(jumlahMapping) = CDbl(dataSumber(0))
                daftarNilaiLama(jumlahMapping) = _
                    wsLokal.Cells(baris, kolomTujuan).Formula

                petaTujuan.Add kunci, baris

            End If

        End If

    Next baris

    If jumlahMapping = 0 Then
        pesanError = _
            "Tidak ada nama baris EWS yang cocok dengan sheet PPh 26. " & _
            "Periksa nama baris dan blok Menurut Wajib Pajak."
        GoTo GagalTanpaError
    End If

    For Each kunciSumber In petaSumber.Keys
        If Not petaTujuan.Exists(CStr(kunciSumber)) Then
            pesanError = "Baris WP lokal PPh 26 belum terpetakan: " & _
                NamaKunciUntukPesan(CStr(kunciSumber))
            GoTo GagalTanpaError
        End If
    Next kunciSumber

    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    penulisanDimulai = True

    For i = 1 To jumlahMapping
        wsLokal.Cells( _
            daftarBaris(i), kolomTujuan).Value2 = _
            daftarNilaiBaru(i)
    Next i

    HitungUlangKKPLokal TAXKEY_PPH26
    ImportPPh26DariWorkbook = True
    Exit Function

Gagal:

    pesanError = Err.Description

GagalTanpaError:

    If penulisanDimulai _
       And jumlahMapping > 0 _
       And Not ImportPPh26DariWorkbook Then

        On Error Resume Next

        For i = 1 To jumlahMapping
            If daftarBaris(i) > 0 Then
                wsLokal.Cells( _
                    daftarBaris(i), kolomTujuan).Formula = _
                    daftarNilaiLama(i)
            End If
        Next i

        On Error GoTo 0

    End If

End Function

'Tujuan: Membangun peta sumber pph 26.
Private Function BangunPetaSumberPPh26( _
    ByVal ws As Worksheet, _
    ByRef peta As Object, _
    ByRef pesanError As String) As Boolean

    Dim baris As Long
    Dim barisAkhir As Long
    Dim kolomNilai As Long
    Dim namaBaris As String
    Dim identifierBaris As String
    Dim konteks As String
    Dim kunci As String
    Dim nilai As Variant
    Dim daftarID As Object

    BangunPetaSumberPPh26 = False
    Set peta = CreateObject("Scripting.Dictionary")
    peta.CompareMode = vbTextCompare
    Set daftarID = CreateObject("Scripting.Dictionary")
    daftarID.CompareMode = vbTextCompare

    kolomNilai = CariKolomNilaiWPPPhPemotongan(ws, pesanError, "PPh26")
    If kolomNilai = 0 Then Exit Function

    barisAkhir = BarisTerakhirTerpakai(ws)
    konteks = vbNullString

    For baris = 1 To barisAkhir

        namaBaris = AmbilNamaEksplisitPPhPemotonganEWS(ws, baris)

        If Len(namaBaris) > 0 Then
            PerbaruiKonteksPPhPemotongan namaBaris, konteks
        End If

        identifierBaris = UCase$(Trim$(NilaiTeksAman( _
            ws.Cells(baris, KOLOM_ID_PPH21_EWS).Value2)))

        If Left$(identifierBaris, 4) = "NIL_" _
           And Len(namaBaris) > 0 Then

            kunci = KunciNamaPPh26( _
                namaBaris, identifierBaris, konteks)

            If Len(kunci) > 0 Then

                If peta.Exists(kunci) Then
                    pesanError = _
                        "Nama baris sumber PPh 26 ditemukan lebih dari " & _
                        "satu kali: " & NamaKunciUntukPesan(kunci)
                    Exit Function
                End If

                If daftarID.Exists(identifierBaris) Then
                    pesanError = "Identifier sumber PPh 26 ganda: " & identifierBaris
                    Exit Function
                End If
                daftarID.Add identifierBaris, baris
                nilai = ws.Cells(baris, kolomNilai).Value2
                If IsError(nilai) Or IsNull(nilai) Or IsEmpty(nilai) Then
                    pesanError = "Nilai WP PPh26 belum siap pada " & identifierBaris
                    Exit Function
                End If
                If Not IsNumeric(nilai) Then
                    pesanError = "Nilai WP PPh26 belum numerik pada " & identifierBaris
                    Exit Function
                End If
                peta.Add kunci, Array(nilai, baris, identifierBaris)

            End If

        End If

    Next baris

    If peta.Count = 0 Then
        pesanError = _
            "Tidak ditemukan baris data PPh 26 pada MAIN LEVEL EWS."
        Exit Function
    End If

    If Not ValidasiIdentifierPPh26(daftarID, pesanError) Then Exit Function
    BangunPetaSumberPPh26 = True

End Function

'Tujuan: Membentuk kunci nama pph 26.
Private Function KunciNamaPPh26( _
    ByVal namaBaris As String, _
    ByVal identifierEWS As String, _
    ByVal konteks As String, _
    Optional ByVal kodeKJS As String = "") As String

    Dim namaNormal As String
    Dim pembeda As String
    Dim potongan As Variant

    namaNormal = NormalisasiNamaBaris(namaBaris)
    If Len(namaNormal) = 0 Then Exit Function

    If InStr(1, namaNormal, "JUMLAHPAJAKYANGDAPATDIKREDITKAN", _
            vbTextCompare) = 1 Then
        namaNormal = "JUMLAHPAJAKYANGDAPATDIKREDITKAN"
    End If

    Select Case namaNormal

        Case "KJS"
            If Len(identifierEWS) > 0 Then
                potongan = Split(identifierEWS, "_")
                If UBound(potongan) >= 1 Then kodeKJS = _
                    CStr(potongan(UBound(potongan) - 1)) & "-" & CStr(potongan(UBound(potongan)))
            End If
            If kodeKJS <> "411127-100" And kodeKJS <> "411127-110" Then
                Err.Raise vbObjectError + 3420, , "KJS PPh26 kosong/tidak didukung: " & kodeKJS
            End If

            If InStr(1, identifierEWS, _
                    "NIL_DPP_KJS", vbTextCompare) > 0 _
               Or UCase$(konteks) = "DPP" Then

                pembeda = "DPP"
            Else
                pembeda = "PPH"
            End If

            pembeda = pembeda & "|" & kodeKJS

        Case "JUMLAHA+B+C"

            If InStr(1, identifierEWS, _
                    "NIL_DIPERHITUNGKAN", vbTextCompare) > 0 _
               Or UCase$(konteks) = "O" Then

                pembeda = "O"
            Else
                pembeda = "P"
            End If

    End Select

    KunciNamaPPh26 = _
        "PPH26|" & namaNormal & "|" & pembeda

End Function

'Tujuan: Memvalidasi identifier pph 26.
Private Function ValidasiIdentifierPPh26( _
    ByVal daftarID As Object, _
    ByRef pesanError As String) As Boolean

    Dim wajib As Variant
    Dim item As Variant
    wajib = Array( _
        "NIL_DPP_KJS_411127_100", "NIL_DPP_KJS_411127_110", _
        "NIL_TOTAL_DPP_PASAL_26", "NIL_PPH_KJS_411127_100", _
        "NIL_PPH_KJS_411127_110", "NIL_TOTAL_PPH_PASAL26", _
        "NIL_PH_TDK_TERHUTANG_PPH_PASAL_17_AYAT_2", "NIL_PPH_HARUS_DIPUNGUT_DISETOR", _
        "NIL_KREDIT_PAJAK_DITANGGUNG_PEMERINTAH", "NIL_KREDIT_PAJAK_SETORAN_MASA", _
        "NIL_KREDIT_PAJAK_STP", "NIL_KOMPENSASI_DARI_MASA_PAJAK", _
        "NIL_KREDIT_PAJAK_LAIN", "NIL_KREDIT_PAJAK_UANG_TEBUSAN", _
        "NIL_KREDIT_PAJAK_UNGKAP_HARTA_BERSIH", "NIL_KOMPENSASI_KE_MASA_PAJAK", _
        "NIL_DIPERHITUNGKAN_SKPKB", "NIL_DIPERHITUNGKAN_SKPKBT", _
        "NIL_DIPERHITUNGKAN_SKPLB", "NIL_DIPERHITUNGKAN", _
        "NIL_PPH_PASAL17_DIBAYAR_NPWP_SENDIRI", "NIL_PPH_PASAL17_DIBAYAR_NPWP_LAIN", _
        "NIL_PPH_PASAL17_TELAH_POTPUT", "NIL_PPH_TDK_TERUTANG_PASAL17", _
        "NIL_KREDIT_PAJAK", "NIL_KLB")
    For Each item In wajib
        If Not daftarID.Exists(CStr(item)) Then
            pesanError = "Struktur PPh 26 belum lengkap/tidak sesuai: " & _
                CStr(item) & " tidak ditemukan pada MAIN LEVEL."
            Exit Function
        End If
    Next item
    ValidasiIdentifierPPh26 = True
End Function

'Tujuan: Mengambil kjslokal pph 26.
Private Function AmbilKJSLokalPPh26(ByVal ws As Worksheet, ByVal baris As Long) As String
    Dim c As Long, t As String, hasil As String
    For c = KOLOM_NAMA_PPH21_KKP_AWAL To KOLOM_NAMA_PPH21_KKP_AKHIR
        t = Replace(NilaiTeksAman(ws.Cells(baris, c).Value2), " ", "")
        If t Like "######-###" Then
            If Len(hasil) > 0 And hasil <> t Then Err.Raise vbObjectError + 3421, , _
                "KJS lokal ganda pada baris " & CStr(baris)
            hasil = t
        End If
    Next c
    AmbilKJSLokalPPh26 = hasil
End Function

'Bagian: handler PPh 4 ayat (2).
'Tujuan: Mengimpor pph 42 dari workbook.
Private Function ImportPPh42DariWorkbook( _
    ByVal wsEWS As Worksheet, _
    ByVal bulanPajak As Long, _
    ByVal tahunPajak As Long, _
    ByRef pesanError As String) As Boolean

    Dim wsLokal As Worksheet
    Dim petaSumber As Object
    Dim petaTujuan As Object

    Dim namaSheetRingkasan As String
    Dim kolomTujuan As Long
    Dim barisAwal As Long
    Dim barisAkhir As Long
    Dim baris As Long
    Dim jumlahMapping As Long

    Dim namaBaris As String
    Dim konteks As String
    Dim kunci As String
    Dim dataSumber As Variant

    Dim daftarBaris() As Long
    Dim daftarNilaiBaru() As Variant
    Dim daftarNilaiLama() As Variant
    Dim i As Long
    Dim penulisanDimulai As Boolean
    Dim kunciSumber As Variant
    Dim daftarKKP As Variant
    Dim kodeDetail As Variant
    Dim selTujuan As Range

    ImportPPh42DariWorkbook = False

    On Error GoTo Gagal

    If bulanPajak < 1 Or bulanPajak > 12 Then
        Err.Raise vbObjectError + 3700, , "Nomor bulan tidak valid."
    End If

    If tahunPajak < 2000 Or tahunPajak > 2100 Then
        Err.Raise vbObjectError + 3701, , "Tahun pajak tidak valid."
    End If

    namaSheetRingkasan = _
        NamaSheetRingkasanPajak(TAXKEY_PPH42)

    If Not WorksheetAda(namaSheetRingkasan, ThisWorkbook) Then
        Err.Raise _
            vbObjectError + 3702, , _
            "Sheet lokal '" & namaSheetRingkasan & _
            "' tidak ditemukan."
    End If

    Set wsLokal = _
        ThisWorkbook.Worksheets(namaSheetRingkasan)

    daftarKKP = DaftarKodeEWSPPh42()
    For Each kodeDetail In daftarKKP
        If Not SheetLokalKKPAda(TAXKEY_PPH42, CStr(kodeDetail), False) Then
            Err.Raise vbObjectError + 3720, , _
                "Sheet detail PPh 4(2) belum tersedia: " & _
                NamaSheetLokalKKP(TAXKEY_PPH42, CStr(kodeDetail))
        End If
    Next kodeDetail
    If Not WorksheetAda(NamaSheetIndukPajak(TAXKEY_PPH42), ThisWorkbook) Then
        Err.Raise vbObjectError + 3721, , _
            "Sheet '" & NamaSheetIndukPajak(TAXKEY_PPH42) & "' belum tersedia."
    End If

    kolomTujuan = CariKolomPeriodePPN( _
        wsLokal, bulanPajak, tahunPajak, pesanError)

    If kolomTujuan = 0 Then GoTo GagalTanpaError

    If Not BangunPetaSumberPPh42( _
            wsEWS, petaSumber, pesanError) Then
        GoTo GagalTanpaError
    End If

    If Not CariBlokMenurutWP( _
            wsLokal, barisAwal, barisAkhir, pesanError) Then
        GoTo GagalTanpaError
    End If

    Set petaTujuan = CreateObject("Scripting.Dictionary")
    petaTujuan.CompareMode = vbTextCompare

    ReDim daftarBaris(1 To barisAkhir - barisAwal + 1)
    ReDim daftarNilaiBaru(1 To barisAkhir - barisAwal + 1)
    ReDim daftarNilaiLama(1 To barisAkhir - barisAwal + 1)

    konteks = vbNullString

    For baris = barisAwal To barisAkhir

        namaBaris = AmbilNamaEksplisitPPhPemotonganKKP(wsLokal, baris)

        If Len(namaBaris) > 0 Then

            PerbaruiKonteksPPhPemotongan namaBaris, konteks
            kunci = KunciNamaPPh42( _
                namaBaris, vbNullString, konteks, AmbilKJSLokalPPh42(wsLokal, baris))

            If Len(kunci) > 0 _
               And petaSumber.Exists(kunci) Then

                If petaTujuan.Exists(kunci) Then
                    pesanError = _
                        "Nama baris tujuan PPh 4(2) ditemukan lebih dari " & _
                        "satu kali: " & NamaKunciUntukPesan(kunci)
                    GoTo GagalTanpaError
                End If

                dataSumber = petaSumber(kunci)

                If IsError(dataSumber(0)) Then
                    pesanError = _
                        "Nilai sumber EWS mengalami error pada baris '" & _
                        NamaKunciUntukPesan(kunci) & "'."
                    GoTo GagalTanpaError
                End If

                If IsEmpty(dataSumber(0)) Or IsNull(dataSumber(0)) Then
                    pesanError = "Nilai sumber PPh 4(2) kosong: " & _
                        NamaKunciUntukPesan(kunci)
                    GoTo GagalTanpaError
                End If
                If Not IsNumeric(dataSumber(0)) Then
                    pesanError = "Nilai sumber PPh 4(2) bukan angka: " & _
                        NamaKunciUntukPesan(kunci)
                    GoTo GagalTanpaError
                End If
                Set selTujuan = wsLokal.Cells(baris, kolomTujuan)
                If selTujuan.MergeCells Then
                    pesanError = "Sel nilai WP PPh 4(2) tergabung: " & _
                        selTujuan.Address(False, False)
                    GoTo GagalTanpaError
                End If
                If wsLokal.ProtectContents Then
                    If selTujuan.Locked Then
                        pesanError = "Sel nilai WP PPh 4(2) terkunci: " & _
                            selTujuan.Address(False, False)
                        GoTo GagalTanpaError
                    End If
                End If

                jumlahMapping = jumlahMapping + 1
                daftarBaris(jumlahMapping) = baris
                daftarNilaiBaru(jumlahMapping) = CDbl(dataSumber(0))
                daftarNilaiLama(jumlahMapping) = _
                    wsLokal.Cells(baris, kolomTujuan).Formula

                petaTujuan.Add kunci, baris

            End If

        End If

    Next baris

    If jumlahMapping = 0 Then
        pesanError = _
            "Tidak ada nama baris EWS yang cocok dengan sheet PPh 4(2). " & _
            "Periksa nama baris dan blok Menurut Wajib Pajak."
        GoTo GagalTanpaError
    End If

    For Each kunciSumber In petaSumber.Keys
        If Not petaTujuan.Exists(CStr(kunciSumber)) Then
            pesanError = "Baris WP lokal PPh 4(2) belum terpetakan: " & _
                NamaKunciUntukPesan(CStr(kunciSumber))
            GoTo GagalTanpaError
        End If
    Next kunciSumber

    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    penulisanDimulai = True

    For i = 1 To jumlahMapping
        wsLokal.Cells( _
            daftarBaris(i), kolomTujuan).Value2 = _
            daftarNilaiBaru(i)
    Next i

    HitungUlangKKPLokal TAXKEY_PPH42
    ImportPPh42DariWorkbook = True
    Exit Function

Gagal:

    pesanError = Err.Description

GagalTanpaError:

    If penulisanDimulai _
       And jumlahMapping > 0 _
       And Not ImportPPh42DariWorkbook Then

        On Error Resume Next

        For i = 1 To jumlahMapping
            If daftarBaris(i) > 0 Then
                wsLokal.Cells( _
                    daftarBaris(i), kolomTujuan).Formula = _
                    daftarNilaiLama(i)
            End If
        Next i

        On Error GoTo 0

    End If

End Function

'Tujuan: Membangun peta sumber pph 42.
Private Function BangunPetaSumberPPh42( _
    ByVal ws As Worksheet, _
    ByRef peta As Object, _
    ByRef pesanError As String) As Boolean

    Dim baris As Long
    Dim barisAkhir As Long
    Dim kolomNilai As Long
    Dim namaBaris As String
    Dim identifierBaris As String
    Dim konteks As String
    Dim kunci As String
    Dim nilai As Variant
    Dim daftarID As Object

    BangunPetaSumberPPh42 = False
    Set peta = CreateObject("Scripting.Dictionary")
    peta.CompareMode = vbTextCompare
    Set daftarID = CreateObject("Scripting.Dictionary")
    daftarID.CompareMode = vbTextCompare

    kolomNilai = CariKolomNilaiWPPPhPemotongan(ws, pesanError, "PPh4(2)")
    If kolomNilai = 0 Then Exit Function

    barisAkhir = BarisTerakhirTerpakai(ws)
    konteks = vbNullString

    For baris = 1 To barisAkhir

        namaBaris = AmbilNamaEksplisitPPhPemotonganEWS(ws, baris)

        If Len(namaBaris) > 0 Then
            PerbaruiKonteksPPhPemotongan namaBaris, konteks
        End If

        identifierBaris = UCase$(Trim$(NilaiTeksAman( _
            ws.Cells(baris, KOLOM_ID_PPH21_EWS).Value2)))

        If Left$(identifierBaris, 4) = "NIL_" _
           And Len(namaBaris) > 0 Then

            kunci = KunciNamaPPh42( _
                namaBaris, identifierBaris, konteks)

            If Len(kunci) > 0 Then

                If peta.Exists(kunci) Then
                    pesanError = _
                        "Nama baris sumber PPh 4(2) ditemukan lebih dari " & _
                        "satu kali: " & NamaKunciUntukPesan(kunci)
                    Exit Function
                End If

                If daftarID.Exists(identifierBaris) Then
                    pesanError = "Identifier sumber PPh 4(2) ganda: " & identifierBaris
                    Exit Function
                End If
                daftarID.Add identifierBaris, baris
                nilai = ws.Cells(baris, kolomNilai).Value2
                If IsError(nilai) Or IsNull(nilai) Or IsEmpty(nilai) Then
                    pesanError = "Nilai WP PPh42 belum siap pada " & identifierBaris
                    Exit Function
                End If
                If Not IsNumeric(nilai) Then
                    pesanError = "Nilai WP PPh42 belum numerik pada " & identifierBaris
                    Exit Function
                End If
                peta.Add kunci, Array(nilai, baris, identifierBaris)

            End If

        End If

    Next baris

    If peta.Count = 0 Then
        pesanError = _
            "Tidak ditemukan baris data PPh 4(2) pada MAIN LEVEL EWS."
        Exit Function
    End If

    If Not ValidasiIdentifierPPh42(daftarID, pesanError) Then Exit Function
    BangunPetaSumberPPh42 = True

End Function

'Tujuan: Membentuk kunci nama pph 42.
Private Function KunciNamaPPh42( _
    ByVal namaBaris As String, _
    ByVal identifierEWS As String, _
    ByVal konteks As String, _
    Optional ByVal kodeKJS As String = "") As String

    Dim namaNormal As String
    Dim pembeda As String
    Dim potongan As Variant

    namaNormal = NormalisasiNamaBaris(namaBaris)
    If Len(namaNormal) = 0 Then Exit Function

    If InStr(1, namaNormal, "JUMLAHPAJAKYANGDAPATDIKREDITKAN", _
            vbTextCompare) = 1 Then
        namaNormal = "JUMLAHPAJAKYANGDAPATDIKREDITKAN"
    End If

    Select Case namaNormal

        Case "KJS"
            If Len(identifierEWS) > 0 Then
                potongan = Split(identifierEWS, "_")
                If UBound(potongan) >= 1 Then kodeKJS = _
                    CStr(potongan(UBound(potongan) - 1)) & "-" & CStr(potongan(UBound(potongan)))
            End If
            If kodeKJS <> "411128-100" And kodeKJS <> "411128-402" And _
               kodeKJS <> "411128-403" And kodeKJS <> "411128-420" Then
                Err.Raise vbObjectError + 3820, , "KJS PPh42 kosong/tidak didukung: " & kodeKJS
            End If

            If InStr(1, identifierEWS, _
                    "NIL_DPP_KJS", vbTextCompare) > 0 _
               Or UCase$(konteks) = "DPP" Then

                pembeda = "DPP"
            Else
                pembeda = "PPH"
            End If

            pembeda = pembeda & "|" & kodeKJS

        Case "JUMLAHA+B+C"

            If InStr(1, identifierEWS, _
                    "NIL_DIPERHITUNGKAN", vbTextCompare) > 0 _
               Or UCase$(konteks) = "O" Then

                pembeda = "O"
            Else
                pembeda = "P"
            End If

    End Select

    KunciNamaPPh42 = _
        "PPH42|" & namaNormal & "|" & pembeda

End Function

'Tujuan: Memvalidasi identifier pph 42.
Private Function ValidasiIdentifierPPh42( _
    ByVal daftarID As Object, _
    ByRef pesanError As String) As Boolean

    Dim wajib As Variant
    Dim item As Variant
    wajib = Array( _
        "NIL_DPP_KJS_411128_100", "NIL_DPP_KJS_411128_402", _
        "NIL_DPP_KJS_411128_403", "NIL_DPP_KJS_411128_420", _
        "NIL_TOTAL_DPP_PASAL4_AYAT2", _
        "NIL_PPH_KJS_411128_100", "NIL_PPH_KJS_411128_402", _
        "NIL_PPH_KJS_411128_403", "NIL_PPH_KJS_411128_420", _
        "NIL_TOTAL_PPH_PASAL4_AYAT2", _
        "NIL_PH_TDK_TERHUTANG_PPH_PASAL_17_AYAT_2", "NIL_PPH_HARUS_DIPUNGUT_DISETOR", _
        "NIL_KREDIT_PAJAK_DITANGGUNG_PEMERINTAH", "NIL_KREDIT_PAJAK_SETORAN_MASA", _
        "NIL_KREDIT_PAJAK_STP", "NIL_KOMPENSASI_DARI_MASA_PAJAK", _
        "NIL_KREDIT_PAJAK_LAIN", "NIL_KREDIT_PAJAK_UANG_TEBUSAN", _
        "NIL_KREDIT_PAJAK_UNGKAP_HARTA_BERSIH", "NIL_KOMPENSASI_KE_MASA_PAJAK", _
        "NIL_DIPERHITUNGKAN_SKPKB", "NIL_DIPERHITUNGKAN_SKPKBT", _
        "NIL_DIPERHITUNGKAN_SKPLB", "NIL_DIPERHITUNGKAN", _
        "NIL_PPH_PASAL17_DIBAYAR_NPWP_SENDIRI", "NIL_PPH_PASAL17_DIBAYAR_NPWP_LAIN", _
        "NIL_PPH_PASAL17_TELAH_POTPUT", "NIL_PPH_TDK_TERUTANG_PASAL17", _
        "NIL_KREDIT_PAJAK", "NIL_KLB")
    For Each item In wajib
        If Not daftarID.Exists(CStr(item)) Then
            pesanError = "Struktur PPh 4(2) belum lengkap/tidak sesuai: " & _
                CStr(item) & " tidak ditemukan pada MAIN LEVEL."
            Exit Function
        End If
    Next item
    ValidasiIdentifierPPh42 = True
End Function

'Tujuan: Mengambil kjslokal pph 42.
Private Function AmbilKJSLokalPPh42(ByVal ws As Worksheet, ByVal baris As Long) As String
    Dim c As Long, t As String, hasil As String
    For c = KOLOM_NAMA_PPH21_KKP_AWAL To KOLOM_NAMA_PPH21_KKP_AKHIR
        t = Replace(NilaiTeksAman(ws.Cells(baris, c).Value2), " ", "")
        If t Like "######-###" Then
            If Len(hasil) > 0 And hasil <> t Then Err.Raise vbObjectError + 3821, , _
                "KJS lokal ganda pada baris " & CStr(baris)
            hasil = t
        End If
    Next c
    AmbilKJSLokalPPh42 = hasil
End Function

'Bagian: handler PPN.
'Tujuan: Mengimpor ppndari workbook.
Private Function ImportPPNDariWorkbook( _
    ByVal wsEWS As Worksheet, _
    ByVal bulanPajak As Long, _
    ByVal tahunPajak As Long, _
    ByRef pesanError As String) As Boolean

    Dim wsLokal As Worksheet
    Dim petaSumber As Object
    Dim petaTujuan As Object

    Dim kolomTujuan As Long
    Dim barisAwal As Long
    Dim barisAkhir As Long
    Dim baris As Long
    Dim jumlahMapping As Long

    Dim namaAktif As String
    Dim refAktif As String
    Dim namaEksplisit As String
    Dim refBaris As String
    Dim jenisNilai As String
    Dim kunci As String
    Dim kunciPencarian As String
    Dim dataSumber As Variant

    Dim daftarBaris() As Long
    Dim daftarNilaiBaru() As Variant
    Dim daftarNilaiLama() As Variant
    Dim i As Long
    Dim penulisanDimulai As Boolean
    Dim namaSheetRingkasan As String

    ImportPPNDariWorkbook = False

    On Error GoTo Gagal

    If bulanPajak < 1 Or bulanPajak > 12 Then
        Err.Raise vbObjectError + 3200, , "Nomor bulan tidak valid."
    End If

    If tahunPajak < 2000 Or tahunPajak > 2100 Then
        Err.Raise vbObjectError + 3204, , "Tahun pajak tidak valid."
    End If

    namaSheetRingkasan = _
        NamaSheetRingkasanPajak(TAXKEY_PPN_DN)

    If Not WorksheetAda(namaSheetRingkasan, ThisWorkbook) Then
        Err.Raise _
            vbObjectError + 3201, , _
            "Sheet lokal '" & namaSheetRingkasan & _
            "' tidak ditemukan."
    End If

    Set wsLokal = _
        ThisWorkbook.Worksheets(namaSheetRingkasan)

    kolomTujuan = CariKolomPeriodePPN( _
        wsLokal, bulanPajak, tahunPajak, pesanError)

    If kolomTujuan = 0 Then GoTo GagalTanpaError

    If Not BangunPetaSumberPPN( _
            wsEWS, petaSumber, pesanError) Then
        GoTo GagalTanpaError
    End If

    If Not CariBlokMenurutWP( _
            wsLokal, barisAwal, barisAkhir, pesanError) Then
        GoTo GagalTanpaError
    End If

    Set petaTujuan = CreateObject("Scripting.Dictionary")
    petaTujuan.CompareMode = vbTextCompare

    ReDim daftarBaris(1 To barisAkhir - barisAwal + 1)
    ReDim daftarNilaiBaru(1 To barisAkhir - barisAwal + 1)
    ReDim daftarNilaiLama(1 To barisAkhir - barisAwal + 1)

    namaAktif = vbNullString
    refAktif = vbNullString

    For baris = barisAwal To barisAkhir

        namaEksplisit = AmbilNamaEksplisitKKP(wsLokal, baris)
        refBaris = NilaiTeksAman( _
            wsLokal.Cells(baris, KOLOM_REF_KKP).Value2)
        refBaris = BersihkanReferensiBaris(refBaris)

        PerbaruiKonteksNamaBaris _
            namaEksplisit, refBaris, namaAktif, refAktif

        jenisNilai = NilaiTeksAman( _
            wsLokal.Cells( _
                baris, KOLOM_JENIS_NILAI_KKP).Value2)

        If Len(namaEksplisit) > 0 _
           Or Len(refBaris) > 0 _
           Or Len(jenisNilai) > 0 Then

            kunci = BuatKunciNamaBaris( _
                namaAktif, jenisNilai, refAktif)

            If Len(kunci) > 0 Then

                kunciPencarian = kunci

                If Not petaSumber.Exists(kunciPencarian) Then
                    kunciPencarian = KunciDenganPembedaDuplikat( _
                        kunci, refAktif)
                End If

                If Len(kunciPencarian) > 0 _
                   And petaSumber.Exists(kunciPencarian) Then

                    If petaTujuan.Exists(kunciPencarian) Then
                        pesanError = _
                            "Nama baris tujuan ditemukan lebih dari satu " & _
                            "kali: " & _
                            NamaKunciUntukPesan(kunciPencarian)
                        GoTo GagalTanpaError
                    End If

                    dataSumber = petaSumber(kunciPencarian)

                    If IsError(dataSumber(0)) Then
                        pesanError = _
                            "Nilai sumber EWS mengalami error pada baris '" & _
                            NamaKunciUntukPesan(kunciPencarian) & "'."
                        GoTo GagalTanpaError
                    End If

                    jumlahMapping = jumlahMapping + 1
                    daftarBaris(jumlahMapping) = baris
                    daftarNilaiBaru(jumlahMapping) = dataSumber(0)
                    daftarNilaiLama(jumlahMapping) = _
                        wsLokal.Cells(baris, kolomTujuan).formula

                    petaTujuan.Add kunciPencarian, baris

                End If

            End If

        End If

    Next baris

    If jumlahMapping = 0 Then
        pesanError = _
            "Tidak ada nama baris EWS yang cocok dengan sheet PPN. " & _
            "Periksa nama baris dan struktur bagian Menurut Wajib Pajak."
        GoTo GagalTanpaError
    End If

    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    penulisanDimulai = True

    For i = 1 To jumlahMapping
        wsLokal.Cells( _
            daftarBaris(i), kolomTujuan).Value2 = _
            daftarNilaiBaru(i)
    Next i

    ImportPPNDariWorkbook = True
    Exit Function

Gagal:

    pesanError = Err.Description

GagalTanpaError:

    If penulisanDimulai _
       And jumlahMapping > 0 _
       And Not ImportPPNDariWorkbook Then

        On Error Resume Next

        For i = 1 To jumlahMapping
            If daftarBaris(i) > 0 Then
                wsLokal.Cells( _
                    daftarBaris(i), kolomTujuan).formula = _
                    daftarNilaiLama(i)
            End If
        Next i

        On Error GoTo 0

    End If

End Function

'Tujuan: Membangun peta sumber ppn.
Private Function BangunPetaSumberPPN( _
    ByVal ws As Worksheet, _
    ByRef peta As Object, _
    ByRef pesanError As String) As Boolean

    Dim baris As Long
    Dim barisAkhir As Long
    Dim namaAktif As String
    Dim refAktif As String
    Dim namaEksplisit As String
    Dim refBaris As String
    Dim jenisNilai As String
    Dim kunci As String
    Dim nilai As Variant
    Dim dataLama As Variant
    Dim kunciLama As String
    Dim kunciBaru As String

    BangunPetaSumberPPN = False
    Set peta = CreateObject("Scripting.Dictionary")
    peta.CompareMode = vbTextCompare

    barisAkhir = BarisTerakhirTerpakai(ws)

    For baris = 1 To barisAkhir

        namaEksplisit = AmbilNamaEksplisitEWS(ws, baris)
        refBaris = NilaiTeksAman( _
            ws.Cells(baris, KOLOM_REF_EWS).Value2)
        refBaris = BersihkanReferensiBaris(refBaris)

        PerbaruiKonteksNamaBaris _
            namaEksplisit, refBaris, namaAktif, refAktif

        jenisNilai = NilaiTeksAman( _
            ws.Cells( _
                baris, KOLOM_JENIS_NILAI_EWS).Value2)

        kunci = BuatKunciNamaBaris( _
            namaAktif, jenisNilai, refAktif)

        If Len(kunci) > 0 Then

            nilai = ws.Cells( _
                baris, KOLOM_NILAI_EWS).Value

            If Not IsEmpty(nilai) Then

                If peta.Exists(kunci) Then

                    dataLama = peta(kunci)
                    kunciLama = KunciDenganPembedaDuplikat( _
                        kunci, CStr(dataLama(2)))
                    kunciBaru = KunciDenganPembedaDuplikat( _
                        kunci, refAktif)

                    If Len(kunciLama) = 0 Or Len(kunciBaru) = 0 _
                       Or StrComp( _
                            kunciLama, kunciBaru, _
                            vbTextCompare) = 0 Then

                        pesanError = _
                            "Nama baris sumber ditemukan lebih dari satu " & _
                            "kali dan tidak memiliki pembeda Ref yang " & _
                            "unik: " & NamaKunciUntukPesan(kunci)
                        Exit Function
                    End If

                    peta.Remove kunci
                    peta.Add kunciLama, dataLama
                    peta.Add kunciBaru, Array(nilai, baris, refAktif)

                ElseIf KamusMemilikiAwalan( _
                        peta, _
                        kunci & "|PEMBEDA_REF|") Then

                    kunciBaru = KunciDenganPembedaDuplikat( _
                        kunci, refAktif)

                    If Len(kunciBaru) = 0 _
                       Or peta.Exists(kunciBaru) Then

                        pesanError = _
                            "Nama baris sumber ditemukan lebih dari satu " & _
                            "kali dan tidak memiliki pembeda Ref yang " & _
                            "unik: " & NamaKunciUntukPesan(kunci)
                        Exit Function
                    End If

                    peta.Add kunciBaru, Array(nilai, baris, refAktif)

                Else

                    peta.Add kunci, Array(nilai, baris, refAktif)

                End If

            End If

        End If

    Next baris

    If peta.Count = 0 Then
        pesanError = _
            "Tidak ditemukan baris data pada MAIN LEVEL EWS."
        Exit Function
    End If

    BangunPetaSumberPPN = True

End Function

'Tujuan: Mencari blok menurut wp.
Private Function CariBlokMenurutWP( _
    ByVal ws As Worksheet, _
    ByRef barisAwal As Long, _
    ByRef barisAkhir As Long, _
    ByRef pesanError As String) As Boolean

    Dim ditemukan As Range
    Dim r As Long
    Dim c As Long
    Dim teks As String
    Dim barisTerakhir As Long

    CariBlokMenurutWP = False

    Set ditemukan = ws.Cells.Find( _
        What:="Menurut Wajib Pajak", _
        After:=ws.Cells(ws.Rows.Count, ws.Columns.Count), _
        LookIn:=xlValues, _
        LookAt:=xlWhole, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, _
        MatchCase:=False)

    If ditemukan Is Nothing Then
        pesanError = _
            "Marker 'Menurut Wajib Pajak' tidak ditemukan pada sheet '" & _
            ws.Name & "'."
        Exit Function
    End If

    barisAwal = ditemukan.Row + 1
    barisTerakhir = BarisTerakhirTerpakai(ws)
    barisAkhir = barisTerakhir

    For r = barisAwal To barisTerakhir
        For c = KOLOM_NAMA_KKP_AWAL To KOLOM_NAMA_KKP_AKHIR

            teks = UCase$(Trim$(NilaiTeksAman( _
                ws.Cells(r, c).Value2)))

            If Left$(teks, 8) = "MENURUT " Then
                barisAkhir = r - 1
                Exit For
            End If

        Next c

        If barisAkhir < barisTerakhir Then Exit For
    Next r

    If barisAwal > barisAkhir Then
        pesanError = "Blok Menurut Wajib Pajak tidak memiliki data."
        Exit Function
    End If

    CariBlokMenurutWP = True

End Function

'Tujuan: Mengambil nama eksplisit ews.
Private Function AmbilNamaEksplisitEWS( _
    ByVal ws As Worksheet, _
    ByVal baris As Long) As String

    Dim c As Long
    Dim teks As String

    For c = KOLOM_NAMA_EWS_3 To KOLOM_NAMA_EWS_1 Step -1
        teks = NilaiTeksAman(ws.Cells(baris, c).Value2)

        If NamaBarisDapatDipakai(teks) Then
            AmbilNamaEksplisitEWS = teks
            Exit Function
        End If
    Next c

End Function

'Tujuan: Mengambil nama eksplisit kkp.
Private Function AmbilNamaEksplisitKKP( _
    ByVal ws As Worksheet, _
    ByVal baris As Long) As String

    Dim c As Long
    Dim teks As String

    For c = KOLOM_NAMA_KKP_AKHIR _
        To KOLOM_NAMA_KKP_AWAL Step -1

        teks = NilaiTeksAman(ws.Cells(baris, c).Value2)

        If NamaBarisDapatDipakai(teks) Then
            AmbilNamaEksplisitKKP = teks
            Exit Function
        End If
    Next c

End Function

'Tujuan: Memperbarui konteks nama baris.
Private Sub PerbaruiKonteksNamaBaris( _
    ByVal namaEksplisit As String, _
    ByVal refBaris As String, _
    ByRef namaAktif As String, _
    ByRef refAktif As String)

    If Len(Trim$(namaEksplisit)) > 0 Then
        namaAktif = namaEksplisit
        refAktif = vbNullString
    End If

    If Len(Trim$(refBaris)) > 0 Then
        refAktif = refBaris
    End If

End Sub

'Tujuan: Membentuk kunci nama baris.
Private Function BuatKunciNamaBaris( _
    ByVal namaBaris As String, _
    ByVal jenisNilai As String, _
    ByVal referensi As String) As String

    Dim namaNormal As String
    Dim jenisNormal As String
    namaNormal = NormalisasiNamaBaris(namaBaris)
    jenisNormal = NormalisasiNamaBaris(jenisNilai)

    If Len(namaNormal) > 0 Then
        BuatKunciNamaBaris = _
            "NAMA|" & namaNormal & "|" & jenisNormal
    End If

End Function

'Tujuan: Menjalankan kamus memiliki awalan.
Private Function KamusMemilikiAwalan( _
    ByVal kamus As Object, _
    ByVal awalan As String) As Boolean

    Dim kunci As Variant

    For Each kunci In kamus.Keys
        If Left$(CStr(kunci), Len(awalan)) = awalan Then
            KamusMemilikiAwalan = True
            Exit Function
        End If
    Next kunci

End Function

'Tujuan: Membentuk kunci dengan pembeda duplikat.
Private Function KunciDenganPembedaDuplikat( _
    ByVal kunciNama As String, _
    ByVal referensi As String) As String

    Dim refNormal As String

    refNormal = NormalisasiNamaBaris(referensi)
    If Len(refNormal) = 0 Then Exit Function

    KunciDenganPembedaDuplikat = _
        kunciNama & "|PEMBEDA_REF|" & refNormal

End Function

'Tujuan: Mengembalikan nama kunci untuk pesan.
Private Function NamaKunciUntukPesan( _
    ByVal kunci As String) As String

    NamaKunciUntukPesan = Replace(kunci, "|", " / ")

End Function

'Tujuan: Menormalkan nama baris.
Private Function NormalisasiNamaBaris( _
    ByVal teks As String) As String

    Dim hasil As String
    Dim daftarTanda As Variant
    Dim i As Long

    hasil = UCase$(Trim$(teks))
    hasil = Replace(hasil, Chr$(160), " ")
    hasil = Replace(hasil, vbCr, " ")
    hasil = Replace(hasil, vbLf, " ")
    hasil = Replace(hasil, "SEARUSNYA", "SEHARUSNYA")
    hasil = Replace(hasil, "PENGHASIHAN", "PENGHASILAN")

    hasil = Replace( _
        hasil, _
        "PENYERAHAN BARANG DAN/ATAU JASA YANG TIDAK TERUTANG PPN", _
        "TIDAK TERUTANG PPN")

    daftarTanda = Array( _
        " ", ".", ",", ";", ":", "-", "_", _
        "/", "\", "(", ")", "[", "]", "&")

    For i = LBound(daftarTanda) To UBound(daftarTanda)
        hasil = Replace(hasil, CStr(daftarTanda(i)), vbNullString)
    Next i

    NormalisasiNamaBaris = hasil

End Function

'Tujuan: Mengembalikan nama baris dapat dipakai.
Private Function NamaBarisDapatDipakai( _
    ByVal teks As String) As Boolean

    Dim s As String
    Dim sisa As String
    Dim i As Long
    Dim karakter As String

    s = UCase$(Trim$(teks))

    If Len(s) = 0 Then Exit Function

    Select Case s
        Case "URAIAN", "REF", "REF.", _
             "DPP", "DPP LAIN", "PPN"
            Exit Function
    End Select

    sisa = vbNullString

    For i = 1 To Len(s)
        karakter = Mid$(s, i, 1)

        If InStr(1, "0123456789.() ", karakter, _
                vbBinaryCompare) = 0 Then
            sisa = sisa & karakter
        End If
    Next i

    If Len(sisa) <= 1 And Len(s) <= 10 Then Exit Function

    NamaBarisDapatDipakai = True

End Function

'Tujuan: Mengolah nilai teks aman.
Private Function NilaiTeksAman(ByVal nilai As Variant) As String

    If IsError(nilai) Or IsNull(nilai) Or IsEmpty(nilai) Then
        NilaiTeksAman = vbNullString
    Else
        NilaiTeksAman = Trim$(CStr(nilai))
    End If

End Function

'Tujuan: Menjalankan bersihkan referensi baris.
Private Function BersihkanReferensiBaris( _
    ByVal referensi As String) As String

    Dim s As String

    s = Trim$(referensi)

    If StrComp(s, "REF", vbTextCompare) = 0 _
       Or StrComp(s, "REF.", vbTextCompare) = 0 Then
        BersihkanReferensiBaris = vbNullString
    Else
        BersihkanReferensiBaris = s
    End If

End Function

'Tujuan: Mencari kolom periode pada header KKP berformat YYYY-MM.
Private Function CariKolomPeriodePPN( _
    ByVal ws As Worksheet, _
    ByVal bulanPajak As Long, _
    ByVal tahunPajak As Long, _
    ByRef pesanError As String) As Long

    Dim kandidat As Object
    Dim kunciTarget As String
    Dim barisMaks As Long
    Dim kolomMaks As Long
    Dim r As Long
    Dim c As Long
    Dim daftarKandidat As Variant

    Set kandidat = CreateObject("Scripting.Dictionary")
    kandidat.CompareMode = vbTextCompare

    kunciTarget = KunciPeriode(tahunPajak, bulanPajak)

    barisMaks = BarisTerakhirTerpakai(ws)
    If barisMaks > BATAS_BARIS_HEADER_PERIODE Then
        barisMaks = BATAS_BARIS_HEADER_PERIODE
    End If

    kolomMaks = KolomTerakhirTerpakai(ws)

    For r = 1 To barisMaks
        For c = 1 To kolomMaks
            If KunciPeriodeDariSel(ws.Cells(r, c)) = kunciTarget Then
                kandidat(CStr(c)) = c
            End If
        Next c
    Next r

    If kandidat.Count = 1 Then
        daftarKandidat = kandidat.Items
        CariKolomPeriodePPN = CLng(daftarKandidat(0))
    ElseIf kandidat.Count = 0 Then
        pesanError = _
            "Kolom periode " & kunciTarget & _
            " tidak ditemukan pada header sheet '" & ws.Name & _
            "'. Header harus menggunakan format YYYY-MM."
    Else
        pesanError = _
            "Kolom periode " & kunciTarget & _
            " ditemukan pada lebih dari satu posisi."
    End If

End Function

'Tujuan: Membentuk kunci periode yang konsisten untuk mesin impor.
Private Function KunciPeriode( _
    ByVal tahun As Long, _
    ByVal bulan As Long) As String

    If tahun < 2000 Or tahun > 2100 Then Exit Function
    If bulan < 1 Or bulan > 12 Then Exit Function

    KunciPeriode = _
        Format$(tahun, "0000") & "-" & Format$(bulan, "00")

End Function

'Tujuan: Membaca kunci periode hanya dari header yang tampil sebagai YYYY-MM.
Private Function KunciPeriodeDariSel( _
    ByVal sel As Range) As String

    Dim nilai As Variant
    Dim teks As String
    Dim tahun As Long
    Dim bulan As Long

    nilai = sel.Value2
    If IsError(nilai) Or IsEmpty(nilai) Then Exit Function

    'Header dapat berisi tanggal Excel asli dengan format angka yyyy-mm,
    'atau teks literal YYYY-MM. Keduanya dibaca dari tampilan yang sama.
    teks = Trim$(sel.Text)
    If Len(teks) = 0 Or teks = "###" Then
        teks = Trim$(CStr(nilai))
    End If

    If teks Like "####-##" Then
        tahun = CLng(Left$(teks, 4))
        bulan = CLng(Right$(teks, 2))
        KunciPeriodeDariSel = KunciPeriode(tahun, bulan)
    End If

End Function

'Tujuan: Menjalankan nomor bulan dari teks.
Private Function NomorBulanDariTeks( _
    ByVal teks As String) As Long

    Dim s As String
    Dim daftar As Variant
    Dim i As Long

    s = UCase$(Trim$(teks))
    s = Replace(s, ".", vbNullString)

    daftar = Array( _
        "", _
        "JAN|JANUARI|JANUARY", _
        "FEB|FEBRUARI|FEBRUARY", _
        "MAR|MARET|MARCH", _
        "APR|APRIL", _
        "MEI|MAY", _
        "JUN|JUNI|JUNE", _
        "JUL|JULI|JULY", _
        "AGU|AGS|AGUSTUS|AUG|AUGUST", _
        "SEP|SEPTEMBER", _
        "OKT|OKTOBER|OCT|OCTOBER", _
        "NOV|NOVEMBER", _
        "DES|DESEMBER|DEC|DECEMBER")

    For i = 1 To 12
        If TeksMemuatToken(s, CStr(daftar(i))) Then
            NomorBulanDariTeks = i
            Exit Function
        End If
    Next i

End Function

'Tujuan: Menjalankan teks memuat token.
Private Function TeksMemuatToken( _
    ByVal teks As String, _
    ByVal daftarToken As String) As Boolean

    Dim token As Variant
    Dim s As String

    s = " " & UCase$(teks) & " "
    s = Replace(s, "-", " ")
    s = Replace(s, "/", " ")
    s = Replace(s, "_", " ")

    For Each token In Split(daftarToken, "|")
        If InStr(1, s, " " & CStr(token) & " ", _
                vbTextCompare) > 0 Then
            TeksMemuatToken = True
            Exit Function
        End If
    Next token

End Function

'Tujuan: Menjalankan tahun dari nilai.
Private Function TahunDariNilai(ByVal nilai As Variant) As Long

    Dim angka As Double

    If IsError(nilai) Or IsEmpty(nilai) Then Exit Function

    If IsNumeric(nilai) Then
        angka = CDbl(nilai)

        If angka >= 2000 And angka <= 2100 Then
            TahunDariNilai = CLng(angka)
            Exit Function
        End If
    End If

    TahunDariNilai = TahunDariTeks(NilaiTeksAman(nilai))

End Function

'Tujuan: Menjalankan tahun dari teks.
Private Function TahunDariTeks(ByVal teks As String) As Long

    Dim s As String
    Dim bagian As Variant
    Dim token As Variant
    Dim angka As Long

    s = UCase$(Trim$(teks))
    s = Replace(s, "-", " ")
    s = Replace(s, "/", " ")
    s = Replace(s, ".", " ")
    s = Replace(s, "_", " ")

    bagian = Split(s, " ")

    For Each token In bagian
        If Len(CStr(token)) = 4 _
           And HanyaAngka(CStr(token)) Then

            angka = CLng(token)
            If angka >= 2000 And angka <= 2100 Then
                TahunDariTeks = angka
                Exit Function
            End If
        End If
    Next token

End Function

'Tujuan: Menjalankan baris terakhir terpakai.
Private Function BarisTerakhirTerpakai( _
    ByVal ws As Worksheet) As Long

    Dim ditemukan As Range

    Set ditemukan = ws.Cells.Find( _
        What:="*", _
        After:=ws.Cells(1, 1), _
        LookIn:=xlFormulas, _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious, _
        MatchCase:=False)

    If ditemukan Is Nothing Then
        BarisTerakhirTerpakai = 1
    Else
        BarisTerakhirTerpakai = ditemukan.Row
    End If

End Function

'Tujuan: Menjalankan kolom terakhir terpakai.
Private Function KolomTerakhirTerpakai( _
    ByVal ws As Worksheet) As Long

    Dim ditemukan As Range

    Set ditemukan = ws.Cells.Find( _
        What:="*", _
        After:=ws.Cells(1, 1), _
        LookIn:=xlFormulas, _
        LookAt:=xlPart, _
        SearchOrder:=xlByColumns, _
        SearchDirection:=xlPrevious, _
        MatchCase:=False)

    If ditemukan Is Nothing Then
        KolomTerakhirTerpakai = 1
    Else
        KolomTerakhirTerpakai = ditemukan.Column
    End If

End Function

'Bagian: rebuild dan stabilisasi EWS.
'Tujuan: Menunggu rebuild ews.
Private Function TungguRebuildEWS( _
    ByVal wb As Workbook, _
    ByVal handler As String, _
    ByVal batasDetik As Double, _
    ByRef pesanError As String, _
    Optional ByVal workbookSudahTerbuka As Boolean = False) As Boolean

    Dim mulai As Double
    Dim berlalu As Double
    Dim mulaiStabil As Double
    Dim mulaiError As Double
    Dim mulaiTidakSiap As Double
    Dim durasiStabil As Double
    Dim sidikLama As String
    Dim sidikBaru As String
    Dim alasanTidakSiap As String
    Dim batasTidakSiap As Double

    TungguRebuildEWS = False

    On Error GoTo Gagal

    mulai = Timer

    Do

        DoEvents

        berlalu = DetikBerlalu(mulai)

        Application.StatusBar = _
            "Menunggu rebuild " & wb.Name & " - " & _
            Format$(berlalu, "0") & " detik"

        If Not WorksheetAda(NAMA_SHEET_MAIN_EWS, wb) Then

            If berlalu >= BATAS_TUNGGU_SHEET_DETIK Then

                pesanError = _
                    "Sheet '" & NAMA_SHEET_MAIN_EWS & _
                    "' tidak tersedia setelah menunggu " & _
                    Format$(BATAS_TUNGGU_SHEET_DETIK, "0") & _
                    " detik."

                Exit Function
            End If

            GoTo LanjutMenunggu
        End If

        alasanTidakSiap = vbNullString
        sidikBaru = SidikJariEWS( _
            wb, handler, alasanTidakSiap)

        If Len(sidikBaru) = 0 Then

            sidikLama = vbNullString
            mulaiStabil = 0
            mulaiError = 0

            If mulaiTidakSiap = 0 Then
                mulaiTidakSiap = Timer
            End If

            If workbookSudahTerbuka Then
                batasTidakSiap = 10
            Else
                batasTidakSiap = BATAS_ERROR_SUMBER_DETIK
            End If

            Application.StatusBar = _
                "Menunggu data EWS siap - " & _
                Format$(berlalu, "0") & " detik: " & _
                alasanTidakSiap

            If DetikBerlalu(mulaiTidakSiap) >= _
                    batasTidakSiap Then

                pesanError = _
                    "Data EWS belum siap setelah " & _
                    Format$(batasTidakSiap, "0") & _
                    " detik." & vbCrLf & alasanTidakSiap

                Exit Function
            End If

            GoTo LanjutMenunggu

        End If

        mulaiTidakSiap = 0

        If InStr( _
                1, _
                sidikBaru, _
                "#ERROR", _
                vbTextCompare) > 0 Then

            sidikLama = vbNullString
            mulaiStabil = 0

            If mulaiError = 0 Then
                mulaiError = Timer
            End If

            If DetikBerlalu(mulaiError) >= _
                    BATAS_ERROR_SUMBER_DETIK Then

                pesanError = _
                    "Nilai sumber EWS masih mengalami error " & _
                    "setelah menunggu " & _
                    Format$(BATAS_ERROR_SUMBER_DETIK, "0") & _
                    " detik."

                Exit Function
            End If

        ElseIf Len(sidikBaru) > 0 _
           And sidikBaru = sidikLama Then

            mulaiError = 0

            If mulaiStabil = 0 Then
                mulaiStabil = Timer
            End If

            durasiStabil = DetikBerlalu(mulaiStabil)

            If berlalu >= MINIMUM_TUNGGU_REBUILD_DETIK _
               And durasiStabil >= _
                   DURASI_STABIL_REBUILD_DETIK Then

                TungguRebuildEWS = True
                Exit Function

            End If

        Else

            sidikLama = sidikBaru
            mulaiStabil = 0
            mulaiError = 0

        End If

LanjutMenunggu:

        If berlalu >= batasDetik Then

            pesanError = _
                "Rebuild atau kalkulasi EWS melewati batas " & _
                Format$(batasDetik, "0") & " detik."

            Exit Function
        End If

    Loop

Gagal:

    pesanError = _
        "Proses rebuild EWS gagal." & vbCrLf & Err.Description

End Function

'Tujuan: Menjalankan sidik jari ews.
Private Function SidikJariEWS( _
    ByVal wb As Workbook, _
    ByVal handler As String, _
    ByRef alasanTidakSiap As String) As String

    Dim ws As Worksheet
    Dim peta As Object
    Dim daftarKunci As Variant
    Dim dataBaris As Variant
    Dim pesanError As String
    Dim hasil As String
    Dim i As Long

    On Error GoTo TidakSiap

    alasanTidakSiap = vbNullString

    If Not WorksheetAda(NAMA_SHEET_MAIN_EWS, wb) Then Exit Function

    Set ws = wb.Worksheets(NAMA_SHEET_MAIN_EWS)

    Select Case handler

        Case HANDLER_PPN_DN

            If Not BangunPetaSumberPPN( _
                    ws, peta, pesanError) Then
                GoTo TidakSiap
            End If

        Case HANDLER_PPH21, HANDLER_PPH21F

            If Not BangunPetaSumberPPh21( _
                    ws, peta, pesanError) Then
                GoTo TidakSiap
            End If

        Case HANDLER_PPH23

            If Not BangunPetaSumberPPh23(ws, peta, pesanError) Then
                GoTo TidakSiap
            End If

        Case HANDLER_PPH26

            If Not BangunPetaSumberPPh26(ws, peta, pesanError) Then
                GoTo TidakSiap
            End If

        Case HANDLER_PPH42

            If Not BangunPetaSumberPPh42(ws, peta, pesanError) Then
                GoTo TidakSiap
            End If

        Case Else
            pesanError = "Handler sidik jari tidak tersedia."
            GoTo TidakSiap

    End Select

    daftarKunci = peta.Keys

    For i = LBound(daftarKunci) To UBound(daftarKunci)
        dataBaris = peta(CStr(daftarKunci(i)))

        hasil = hasil & _
            CStr(daftarKunci(i)) & "=" & _
            NilaiSidikJari(dataBaris(0)) & "|"
    Next i

    SidikJariEWS = hasil

    Exit Function

TidakSiap:

    If Len(pesanError) = 0 Then
        If Err.Number <> 0 Then
            pesanError = Err.Description
        Else
            pesanError = "Nilai sumber belum dapat dibaca."
        End If
    End If

    alasanTidakSiap = pesanError
    SidikJariEWS = vbNullString

End Function

'Tujuan: Mengolah nilai sidik jari.
Private Function NilaiSidikJari( _
    ByVal nilai As Variant) As String

    If IsError(nilai) Then
        NilaiSidikJari = "#ERROR"
    ElseIf IsEmpty(nilai) Then
        NilaiSidikJari = "#EMPTY"
    Else
        NilaiSidikJari = CStr(nilai)
    End If

End Function

'Tujuan: Menjalankan detik berlalu.
Private Function DetikBerlalu( _
    ByVal waktuMulai As Double) As Double

    DetikBerlalu = Timer - waktuMulai

    If DetikBerlalu < 0 Then
        DetikBerlalu = DetikBerlalu + 86400
    End If

End Function

'Tujuan: Membaca metadata dari sheet.
Private Sub BacaMetadataDariSheet( _
    ByVal ws As Worksheet, _
    ByRef namaWP As String, _
    ByRef npwpWorkbook As String)

    namaWP = vbNullString
    npwpWorkbook = vbNullString

    If Not IsError(ws.Range("AG7").Value) Then
        namaWP = Trim$(CStr(ws.Range("AG7").Value))
    End If

    If Not IsError(ws.Range("AG8").Value) Then

        npwpWorkbook = BersihkanNPWP(CStr(ws.Range("AG8").Text))

        If Len(npwpWorkbook) <> 16 Then
            npwpWorkbook = _
                BersihkanNPWP(CStr(ws.Range("AG8").Value))
        End If

    End If

End Sub

'Bagian: pencatatan hasil.
'Tujuan: Menulis log impor.
Private Sub TulisLogImpor( _
    ByVal pathFile As String, _
    ByVal kodeKKP As String, _
    ByVal periode As String, _
    ByVal statusImpor As String, _
    ByVal pesan As String, _
    Optional ByVal npwp As String = "", _
    Optional ByVal bulanAwal As Long = 0, _
    Optional ByVal bulanAkhir As Long = 0, _
    Optional ByVal tahun As Long = 0)

    Dim ws As Worksheet
    Dim baris As Long
    Dim kunci As String
    Dim ukuranFile As Double
    Dim tanggalModifikasi As Date

    On Error GoTo Selesai

    Set ws = DapatkanSheetLog()
    baris = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

    If baris < 2 Then baris = 2

    ws.Cells(baris, 1).Value2 = Now
    ws.Cells(baris, 2).Value2 = kodeKKP
    ws.Cells(baris, 3).Value2 = periode
    If Len(pathFile) > 0 Then
        ws.Cells(baris, 4).Value2 = Dir$(pathFile)
    End If
    ws.Cells(baris, 5).Value2 = statusImpor
    ws.Cells(baris, 6).Value2 = pesan

    If Len(npwp) > 0 _
       And bulanAwal > 0 _
       And bulanAkhir > 0 _
       And tahun > 0 Then

        kunci = KunciImpor( _
            kodeKKP, npwp, bulanAwal, bulanAkhir, tahun)

        ws.Cells(baris, 7).Value2 = npwp
        ws.Cells(baris, 8).Value2 = bulanAwal
        ws.Cells(baris, 9).Value2 = bulanAkhir
        ws.Cells(baris, 10).Value2 = tahun
        ws.Cells(baris, 11).Value2 = kunci

    End If

    If Len(pathFile) > 0 Then

        If Len(Dir$(pathFile)) = 0 Then GoTo Selesai

        ukuranFile = CDbl(FileLen(pathFile))
        tanggalModifikasi = FileDateTime(pathFile)

        ws.Cells(baris, 12).Value2 = ukuranFile
        ws.Cells(baris, 13).Value = tanggalModifikasi

    End If

Selesai:

End Sub

'Tujuan: Mengambil sheet log.
Private Function DapatkanSheetLog() As Worksheet

    Dim ws As Worksheet
    Dim sheetBaru As Boolean

    If WorksheetAda(NAMA_SHEET_LOG, ThisWorkbook) Then

        Set ws = ThisWorkbook.Worksheets(NAMA_SHEET_LOG)

    Else

        Set ws = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets( _
                ThisWorkbook.Worksheets.Count))

        ws.Name = NAMA_SHEET_LOG
        sheetBaru = True

    End If

    ws.Range("A1:M1").Value = Array( _
        "Waktu", "Kode KKP", "Periode", _
        "Nama File", "Status", "Pesan", _
        "NPWP", "Bulan Awal", "Bulan Akhir", "Tahun", _
        "Kunci Impor", "Ukuran File", "Tanggal Modifikasi")

    ws.Rows(1).Font.Bold = True
    ws.Columns("A").NumberFormat = "dd/mm/yyyy hh:mm:ss"
    ws.Columns("M").NumberFormat = "dd/mm/yyyy hh:mm:ss"

    If sheetBaru Then
        ws.Columns("A:M").EntireColumn.AutoFit
    End If

    Set DapatkanSheetLog = ws

End Function

'Tujuan: Menjalankan array mengandung error.
Private Function ArrayMengandungError( _
    ByVal data As Variant) As Boolean

    Dim r As Long
    Dim c As Long

    If Not IsArray(data) Then
        ArrayMengandungError = IsError(data)
        Exit Function
    End If

    For r = LBound(data, 1) To UBound(data, 1)

        For c = LBound(data, 2) To UBound(data, 2)

            If IsError(data(r, c)) Then
                ArrayMengandungError = True
                Exit Function
            End If

        Next c

    Next r

End Function

'Tujuan: Mengembalikan nama file singkat.
Private Function NamaFileSingkat( _
    ByVal pathFile As String, _
    ByVal namaKKP As String, _
    ByVal kodeMasa As String) As String

    Dim namaFile As String
    Dim ekstensi As String

    namaFile = Dir$(pathFile)

    If InStrRev(namaFile, ".") > 0 Then
        ekstensi = Mid$(namaFile, InStrRev(namaFile, "."))
    End If

    NamaFileSingkat = _
        namaKKP & " - " & kodeMasa & ekstensi

End Function

'Tujuan: Mengembalikan nama periode.
Private Function NamaPeriode( _
    ByVal bulanAwal As Long, _
    ByVal bulanAkhir As Long, _
    ByVal tahun As Long) As String

    If bulanAwal = bulanAkhir Then

        NamaPeriode = _
            NamaBulanIndonesia(bulanAwal) & " " & CStr(tahun)

    Else

        NamaPeriode = _
            NamaBulanIndonesia(bulanAwal) & "-" & _
            NamaBulanIndonesia(bulanAkhir) & " " & CStr(tahun)

    End If

End Function

'Tujuan: Menjalankan kode masa valid.
Private Function KodeMasaValid( _
    ByVal kodeMasa As String) As Boolean

    Dim bulanAwal As Long
    Dim bulanAkhir As Long
    Dim tahun As Long

    If Len(kodeMasa) <> 8 Then Exit Function
    If Not HanyaAngka(kodeMasa) Then Exit Function

    bulanAwal = CLng(Left$(kodeMasa, 2))
    bulanAkhir = CLng(Mid$(kodeMasa, 3, 2))
    tahun = CLng(Right$(kodeMasa, 4))

    If bulanAwal < 1 Or bulanAwal > 12 Then Exit Function
    If bulanAkhir < 1 Or bulanAkhir > 12 Then Exit Function
    If tahun < 2000 Or tahun > 2100 Then Exit Function

    KodeMasaValid = True

End Function

'Tujuan: Menjalankan bersihkan npwp.
Private Function BersihkanNPWP( _
    ByVal nilai As String) As String

    Dim hasil As String
    Dim karakter As String
    Dim i As Long

    For i = 1 To Len(nilai)

        karakter = Mid$(nilai, i, 1)

        If karakter >= "0" And karakter <= "9" Then
            hasil = hasil & karakter
        End If

    Next i

    BersihkanNPWP = hasil

End Function

'Tujuan: Menjalankan hanya angka.
Private Function HanyaAngka( _
    ByVal teks As String) As Boolean

    Dim karakter As String
    Dim i As Long

    If Len(teks) = 0 Then Exit Function

    For i = 1 To Len(teks)

        karakter = Mid$(teks, i, 1)

        If karakter < "0" Or karakter > "9" Then
            Exit Function
        End If

    Next i

    HanyaAngka = True

End Function

'Tujuan: Mengambil workbook yang sudah terbuka pada instance aktif.
Private Function DapatkanWorkbookTerbuka( _
    ByVal pathFile As String) As Workbook

    Dim wb As Workbook
    Dim cocok As Boolean

    For Each wb In Application.Workbooks

        cocok = False

        On Error Resume Next
        cocok = (StrComp( _
            wb.FullName, pathFile, vbTextCompare) = 0)
        On Error GoTo 0

        If cocok Then

            Set DapatkanWorkbookTerbuka = wb
            Exit Function

        End If

    Next wb

End Function

'Tujuan: Memeriksa workbook sedang terbuka.
Private Function WorkbookSedangTerbuka( _
    ByVal pathFile As String) As Boolean

    Dim wb As Workbook

    Set wb = DapatkanWorkbookTerbuka(pathFile)
    WorkbookSedangTerbuka = Not wb Is Nothing

End Function

'Tujuan: Menjalankan worksheet ada.
Private Function WorksheetAda( _
    ByVal namaSheet As String, _
    ByVal wb As Workbook) As Boolean

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = wb.Worksheets(namaSheet)
    On Error GoTo 0

    WorksheetAda = Not ws Is Nothing

End Function

'Tujuan: Mengembalikan nama bulan indonesia.
Private Function NamaBulanIndonesia( _
    ByVal nomorBulan As Long) As String

    Dim daftarBulan As Variant

    daftarBulan = Array( _
        "", "Januari", "Februari", "Maret", _
        "April", "Mei", "Juni", "Juli", _
        "Agustus", "September", "Oktober", _
        "November", "Desember")

    If nomorBulan >= 1 And nomorBulan <= 12 Then
        NamaBulanIndonesia = daftarBulan(nomorBulan)
    Else
        NamaBulanIndonesia = "-"
    End If

End Function

'Bagian: pemulihan PPh 26.
'Tujuan: Mengimpor pph 26 pasca error.
Private Function ImportPPh26PascaError( _
    ByVal pathAsli As String, ByVal npwpDiharapkan As String, _
    ByVal bulanAwal As Long, ByVal bulanAkhir As Long, ByVal tahun As Long, _
    ByRef namaWP As String, ByRef pesanError As String) As Boolean
    Dim wb As Workbook, candidate As Workbook, ws As Worksheet
    Dim ev As Boolean, alerts As Boolean, calc As XlCalculation, sec As Long
    Dim openNumber As Long, openText As String, npwp As String
    Dim tahap As String, sudahTerbuka As Boolean, catatan As String
    ev = Application.EnableEvents: alerts = Application.DisplayAlerts
    calc = Application.Calculation: sec = Application.AutomationSecurity
    On Error GoTo Gagal
    tahap = "mencari EWS PPh26"
    For Each candidate In Application.Workbooks
        If StrComp(candidate.FullName, pathAsli, vbTextCompare) = 0 Then
            Set wb = candidate: sudahTerbuka = True: Exit For
        End If
    Next candidate
    If wb Is Nothing Then
        tahap = "membuka EWS dengan startup/rebuild aktif"
        Application.EnableEvents = True
        Application.DisplayAlerts = True
        Application.Calculation = xlCalculationAutomatic
        Application.AutomationSecurity = msoAutomationSecurityByUI
        On Error Resume Next
        Set wb = Workbooks.Open(Filename:=pathAsli, ReadOnly:=False, _
            IgnoreReadOnlyRecommended:=True, AddToMru:=False, Notify:=False)
        openNumber = Err.Number: openText = Err.Description
        Err.Clear
        On Error GoTo Gagal
        If wb Is Nothing Then
            For Each candidate In Application.Workbooks
                If StrComp(candidate.FullName, pathAsli, vbTextCompare) = 0 Then
                    Set wb = candidate: Exit For
                End If
            Next candidate
        End If
        If openNumber <> 0 Then
            If openNumber <> 1004 Or _
                (InStr(1, openText, "Locked", vbTextCompare) = 0 And _
                 InStr(1, openText, "HorizontalAlignment", vbTextCompare) = 0) Then
                Err.Raise openNumber, , openText
            End If
        End If
    End If
    If wb Is Nothing Then Err.Raise vbObjectError + 3470, , _
        "EWS belum terbuka. Buka normal, klik End pada error format, lalu jalankan LanjutImporPPh26Terbuka."
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    tahap = "validasi nilai dan identitas saat ini"
    If Not WorksheetAda(NAMA_SHEET_MAIN_EWS, wb) Then Err.Raise vbObjectError + 3471, , "MAIN LEVEL tidak tersedia."
    If bulanAwal <> bulanAkhir Then Err.Raise vbObjectError + 3472, , "PPh26 harus satu masa pajak."
    Set ws = wb.Worksheets(NAMA_SHEET_MAIN_EWS)
    BacaMetadataPemulihanPPh26 wb, namaWP, npwp
    If Len(npwp) <> 16 Or npwp <> npwpDiharapkan Then Err.Raise vbObjectError + 3473, , _
        "NPWP EWS=[" & npwp & "]; nama file=[" & npwpDiharapkan & "]. Identitas berbeda/belum tersedia."
    tahap = "impor angka WP yang tersedia setelah startup EWS"
    If Not ImportPPh26DariWorkbook(ws, bulanAwal, tahun, pesanError) Then GoTo Selesai
    catatan = "PENGECUALIAN PPh26: memakai angka WP yang tersedia setelah startup/rebuild " & _
        "dijalankan atau dihentikan karena error EWS. Kelengkapan rebuild tidak diverifikasi. " & _
        "Tidak dilakukan kalkulasi ulang atau save EWS oleh handler pemulihan."
    If openNumber <> 0 Then catatan = catatan & " Error startup " & CStr(openNumber) & ": " & openText
    If sudahTerbuka Then catatan = catatan & " Sumber: workbook yang sudah terbuka."
    pesanError = catatan
    ImportPPh26PascaError = True
Selesai:
    On Error Resume Next
    If ImportPPh26PascaError And Not sudahTerbuka Then
        If Not wb Is Nothing Then wb.Close SaveChanges:=False
    End If
    Application.AutomationSecurity = sec
    Application.DisplayAlerts = alerts
    Application.Calculation = calc
    Application.EnableEvents = ev
    On Error GoTo 0
    Exit Function
Gagal:
    pesanError = "Pemulihan PPh26; tahap " & tahap & "; Error " & CStr(Err.Number) & ": " & Err.Description
    Resume Selesai
End Function

'Tujuan: Menjalankan lanjut impor pph 26 terbuka.
Public Sub LanjutImporPPh26Terbuka()
    Dim wb As Workbook, sumber As Workbook, jumlah As Long
    Dim kode As String, jenis As String, npwp As String, masa As String
    Dim awal As Long, akhir As Long, tahun As Long, ident As String
    Dim dikenal As Boolean, didukung As Boolean, handler As String
    Dim pesan As String, namaWP As String, path As String, ok As Boolean
    On Error GoTo Gagal
    For Each wb In Application.Workbooks
        If Not wb Is ThisWorkbook Then
        If InStr(1, wb.Name, KODE_EWS_PPH26 & "_", vbTextCompare) > 0 Then
                jumlah = jumlah + 1: Set sumber = wb
                If wb Is ActiveWorkbook Then Exit For
            End If
        End If
    Next wb
    If sumber Is Nothing Then Err.Raise vbObjectError + 3474, , "Buka EWS PPh26 normal dahulu dan klik End pada error EWS."
    If jumlah > 1 Then
        If Not sumber Is ActiveWorkbook Then Err.Raise vbObjectError + 3475, , "Aktifkan EWS PPh26 yang ingin diimpor."
    End If
    path = sumber.FullName
    If Not BacaIdentitasDariNamaFile(path, kode, jenis, npwp, masa, awal, akhir, tahun, _
        ident, dikenal, didukung, handler, pesan) Then Err.Raise vbObjectError + 3476, , pesan
    If kode <> KODE_EWS_PPH26 Or handler <> HANDLER_PPH26 Then Err.Raise vbObjectError + 3477, , "Hanya untuk EWS PPh26."
    ok = ImportPPh26PascaError(path, npwp, awal, akhir, tahun, namaWP, pesan)
    TulisLogImpor path, kode, NamaPeriode(awal, akhir, tahun), _
        IIf(ok, "Berhasil", "Gagal"), pesan, npwp, awal, akhir, tahun
    If ok Then
        MsgBox "Impor PPh26 berhasil dengan pengecualian pasca-error EWS. Rincian di Log Impor EWS.", vbInformation
    Else
        MsgBox pesan, vbExclamation, "Impor PPh26"
    End If
    Exit Sub
Gagal:
    MsgBox Err.Description, vbExclamation, "Lanjut impor PPh26"
End Sub

'Tujuan: Membaca metadata pemulihan pph 26.
Private Sub BacaMetadataPemulihanPPh26(ByVal wb As Workbook, _
    ByRef namaWP As String, ByRef npwp As String)
    Dim ws As Worksheet, raw As Worksheet, r As Long, k As String
    Dim npwpHeader As String, npwpRaw As String, kandidat As String
    Set ws = wb.Worksheets(NAMA_SHEET_MAIN_EWS)
    namaWP = NilaiTeksAman(ws.Range("AI8").Value2)
    npwpHeader = NormalisasiNPWPPemulihanPPh26(ws.Range("AI9").Value2)
    If WorksheetAda("Main Level Data Source", wb) Then
        Set raw = wb.Worksheets("Main Level Data Source")
        For r = 1 To raw.UsedRange.Row + raw.UsedRange.Rows.Count - 1
            k = UCase$(NilaiTeksAman(raw.Cells(r, 1).Value2))
            If k = "NPWP" Then
                kandidat = NormalisasiNPWPPemulihanPPh26(raw.Cells(r, 2).Value2)
                If Len(npwpRaw) > 0 And Len(kandidat) > 0 And kandidat <> npwpRaw Then _
                    Err.Raise vbObjectError + 3480, , "Field NPWP sumber saling berbeda."
                If Len(kandidat) > 0 Then npwpRaw = kandidat
            ElseIf k = "NAMA_WP" And Len(namaWP) = 0 Then
                namaWP = NilaiTeksAman(raw.Cells(r, 2).Value2)
            End If
        Next r
    End If
    If Len(npwpHeader) > 0 And Len(npwpRaw) > 0 Then
        If npwpHeader <> npwpRaw Then Err.Raise vbObjectError + 3481, , _
            "NPWP header AI9 berbeda dengan field NPWP data sumber."
    End If
    npwp = npwpHeader
    If Len(npwp) = 0 Then npwp = npwpRaw
End Sub

'Tujuan: Menormalkan npwppemulihan pph 26.
Private Function NormalisasiNPWPPemulihanPPh26(ByVal v As Variant) As String
    Dim t As String
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then Exit Function
    If VarType(v) <> vbString And IsNumeric(v) Then
        If CDbl(v) <= 0 Or CDbl(v) <> Fix(CDbl(v)) Then Exit Function
        t = Format$(CDbl(v), "0000000000000000")
    Else
        t = Trim$(CStr(v))
        t = Replace(t, ".", "")
        t = Replace(t, "-", "")
        t = Replace(t, " ", "")
    End If
    If Len(t) = 16 And HanyaAngka(t) Then NormalisasiNPWPPemulihanPPh26 = t
End Function

'Tujuan: Mengimpor pph 26 otomatis.
Public Sub ImportPPh26Otomatis()
    Dim dlg As Object, path As String, helper As String, cmd As String
    Dim kode As String, jenis As String, npwp As String, masa As String
    Dim awal As Long, akhir As Long, tahun As Long, ident As String
    Dim dikenal As Boolean, didukung As Boolean, handler As String
    Dim pesan As String, namaWP As String, ok As Boolean
    Dim ticket As String, nomor As Integer
    On Error GoTo Gagal
    helper = ThisWorkbook.Path & Application.PathSeparator & "EWSSYNC_P26_AutoEnd.ps1"
    If Len(Dir$(helper)) = 0 Then Err.Raise vbObjectError + 3490, , _
        "Letakkan EWSSYNC_P26_AutoEnd.ps1 di folder yang sama dengan workbook EWSSYNC."
    Set dlg = Application.FileDialog(msoFileDialogFilePicker)
    With dlg
        .Title = "Pilih satu EWS PPh26 untuk impor otomatis pasca-error"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "EWS Excel", "*.xlsm"
        If .Show <> -1 Then Exit Sub
        path = .SelectedItems(1)
    End With
    If Not BacaIdentitasDariNamaFile(path, kode, jenis, npwp, masa, awal, akhir, tahun, _
        ident, dikenal, didukung, handler, pesan) Then Err.Raise vbObjectError + 3491, , pesan
    If kode <> KODE_EWS_PPH26 Or handler <> HANDLER_PPH26 Then Err.Raise vbObjectError + 3492, , "Pilih EWS PPh26 TOT09."
    If WorkbookSedangTerbuka(path) Then
        LanjutImporPPh26Otomatis path
        Exit Sub
    End If
    ticket = Environ$("TEMP") & Application.PathSeparator & CreateObject("Scripting.FileSystemObject").GetTempName
    nomor = FreeFile
    Open ticket For Output As #nomor
    Print #nomor, path
    Close #nomor
    cmd = "powershell.exe -NoProfile -NonInteractive -WindowStyle Hidden -File " & _
        KutipArgPPh26(helper) & " -ExcelHwnd " & CStr(Application.Hwnd) & _
        " -HostWorkbook " & KutipArgPPh26(ThisWorkbook.FullName) & _
        " -EwsPath " & KutipArgPPh26(path) & " -TicketPath " & KutipArgPPh26(ticket)
    CreateObject("WScript.Shell").Run cmd, 0, False
    ok = ImportPPh26PascaError(path, npwp, awal, akhir, tahun, namaWP, pesan)
    On Error Resume Next
    Kill ticket
    On Error GoTo Gagal
    TulisLogImpor path, kode, NamaPeriode(awal, akhir, tahun), _
        IIf(ok, "Berhasil", "Gagal"), pesan, npwp, awal, akhir, tahun
    If ok Then
        MsgBox "Impor PPh26 selesai. Rincian pengecualian ada di Log Impor EWS.", vbInformation
    Else
        MsgBox pesan, vbExclamation, "Impor PPh26"
    End If
    Exit Sub
Gagal:
    pesan = Err.Description
    On Error Resume Next
    If Len(ticket) > 0 Then Kill ticket
    On Error GoTo 0
    MsgBox pesan, vbExclamation, "Impor PPh26 otomatis"
End Sub

'Tujuan: Mengutip arg pph 26.
Private Function KutipArgPPh26(ByVal teks As String) As String
    If InStr(teks, Chr$(34)) > 0 Then Err.Raise vbObjectError + 3493, , "Path mengandung tanda kutip tidak valid."
    KutipArgPPh26 = Chr$(34) & teks & Chr$(34)
End Function

'Tujuan: Menjalankan lanjut impor pph 26 otomatis.
Public Sub LanjutImporPPh26Otomatis(ByVal path As String)
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, path, vbTextCompare) = 0 Then
            wb.Activate
            LanjutImporPPh26Terbuka
            Exit Sub
        End If
    Next wb
    MsgBox "EWS tujuan pemulihan tidak lagi terbuka.", vbExclamation, "PPh26"
End Sub

'Tujuan: Mengimpor pph 26 dari daftar form.
Public Sub ImportPPh26DariDaftarForm(ByVal frm As Object)
    Dim helper As String, manifest As String, cmd As String
    Dim nomor As Integer, i As Long, jumlahP26 As Long, jumlahSiap As Long
    Dim jumlahLewat As Long, jumlahImporUlang As Long
    Dim pathFile As String, pathPertama As String
    Dim statusFile As String, kodeKKP As String
    Dim pesan As String, konfirmasi As String, powershellExe As String
    Dim idProses As Double
    Dim unik As Object

    On Error GoTo Gagal

    If frm Is Nothing Then Err.Raise vbObjectError + 3510, , _
        "Form importer tidak tersedia."

    If frm.lstEWS.ListCount = 0 Then
        MsgBox "Muat file EWS terlebih dahulu melalui tombol Pilih File EWS.", _
            vbExclamation, "Impor batch PPh26"
        Exit Sub
    End If

    ValidasiUlangDaftar frm
    UpdateTampilanForm frm
    Set unik = CreateObject("Scripting.Dictionary")
    unik.CompareMode = vbTextCompare

    For i = 0 To frm.lstEWS.ListCount - 1
        kodeKKP = Trim$(CStr(frm.lstEWS.List(i, LST_KODE)))
        If StrComp(kodeKKP, KODE_EWS_PPH26, vbTextCompare) = 0 Then
            jumlahP26 = jumlahP26 + 1
            statusFile = CStr(frm.lstEWS.List(i, LST_STATUS))
            If StatusDapatDiimpor(statusFile) Then
                pathFile = Trim$(CStr(frm.lstEWS.List(i, LST_PATH)))
                If Len(pathFile) = 0 Or Len(Dir$(pathFile)) = 0 Then
                    jumlahLewat = jumlahLewat + 1
                ElseIf Not unik.Exists(pathFile) Then
                    unik.Add pathFile, True
                    jumlahSiap = jumlahSiap + 1
                    If Len(pathPertama) = 0 Then pathPertama = pathFile
                    If StrComp(statusFile, STATUS_SUDAH_DIIMPOR, _
                            vbTextCompare) = 0 Then
                        jumlahImporUlang = jumlahImporUlang + 1
                    End If
                End If
            Else
                jumlahLewat = jumlahLewat + 1
            End If
        End If
    Next i

    If jumlahP26 = 0 Then
        MsgBox "Daftar belum memuat file EWS PPh26 (" & _
            KODE_EWS_PPH26 & ").", _
            vbExclamation, "Impor batch PPh26"
        Exit Sub
    End If

    If jumlahSiap = 0 Then
        MsgBox "Tidak ada file EWS PPh26 yang lolos validasi ulang.", _
            vbExclamation, "Impor batch PPh26"
        Exit Sub
    End If

    If Not KonfirmasiTrustedLocationPPh26(pathPertama) Then Exit Sub

    If jumlahLewat > 0 Or jumlahImporUlang > 0 Then
        konfirmasi = CStr(jumlahSiap) & " file PPh26 siap diproses."
        If jumlahLewat > 0 Then konfirmasi = konfirmasi & vbCrLf & _
            CStr(jumlahLewat) & " file PPh26 dilewati karena status/path tidak valid."
        If jumlahImporUlang > 0 Then konfirmasi = konfirmasi & vbCrLf & vbCrLf & _
            CStr(jumlahImporUlang) & " file pernah diimpor dan akan menimpa periode yang sama."
        konfirmasi = konfirmasi & vbCrLf & vbCrLf & "Lanjutkan batch PPh26?"
        If MsgBox(konfirmasi, vbOKCancel + vbQuestion, _
                "Impor batch PPh26") <> vbOK Then Exit Sub
    End If

    helper = ThisWorkbook.Path & Application.PathSeparator & _
        "EWSSYNC_P26_Batch.ps1"
    If Len(Dir$(helper)) = 0 Then Err.Raise vbObjectError + 3511, , _
        "Letakkan EWSSYNC_P26_Batch.ps1 di folder yang sama dengan workbook EWSSYNC."

    manifest = Environ$("TEMP") & Application.PathSeparator & _
        "EWSSYNC_P26_FORM_" & Format$(Now, "yyyymmdd_hhnnss") & ".txt"
    nomor = FreeFile
    Open manifest For Output As #nomor
    For i = 0 To frm.lstEWS.ListCount - 1
        kodeKKP = Trim$(CStr(frm.lstEWS.List(i, LST_KODE)))
        statusFile = CStr(frm.lstEWS.List(i, LST_STATUS))
        If StrComp(kodeKKP, KODE_EWS_PPH26, vbTextCompare) = 0 _
           And StatusDapatDiimpor(statusFile) Then
            pathFile = Trim$(CStr(frm.lstEWS.List(i, LST_PATH)))
            If Len(pathFile) > 0 And Len(Dir$(pathFile)) > 0 Then
                If unik.Exists(pathFile) Then
                    Print #nomor, pathFile
                    unik.Remove pathFile
                End If
            End If
        End If
    Next i
    Close #nomor
    nomor = 0

    powershellExe = Environ$("SystemRoot") & _
        "\System32\WindowsPowerShell\v1.0\powershell.exe"
    If Len(Dir$(powershellExe)) = 0 Then powershellExe = "powershell.exe"
    cmd = KutipArgPPh26(powershellExe) & _
        " -NoProfile -NonInteractive -ExecutionPolicy Bypass -WindowStyle Hidden -File " & _
        KutipArgPPh26(helper) & " -Mode Batch -ExcelHwnd " & CStr(Application.Hwnd) & _
        " -HostWorkbook " & KutipArgPPh26(ThisWorkbook.FullName) & _
        " -ManifestPath " & KutipArgPPh26(manifest)

    idProses = Shell(cmd, vbHide)
    If idProses = 0 Then Err.Raise vbObjectError + 3512, , _
        "Helper PPh26 gagal dijalankan."

    frm.Hide
    Exit Sub

Gagal:
    pesan = Err.Description
    On Error Resume Next
    If nomor > 0 Then Close #nomor
    If Len(manifest) > 0 Then Kill manifest
    On Error GoTo 0
    MsgBox pesan, vbExclamation, "Impor batch PPh26"
End Sub

'Tujuan: Mengimpor pph 26 batch otomatis.
Public Sub ImportPPh26BatchOtomatis()
    Dim dlg As Object, item As Variant, helper As String, manifest As String
    Dim cmd As String, nomor As Integer, jumlah As Long
    Dim kode As String, jenis As String, npwp As String, masa As String
    Dim awal As Long, akhir As Long, tahun As Long, ident As String
    Dim dikenal As Boolean, didukung As Boolean, handler As String, pesan As String
    On Error GoTo Gagal
    helper = ThisWorkbook.Path & Application.PathSeparator & "EWSSYNC_P26_Batch.ps1"
    If Len(Dir$(helper)) = 0 Then Err.Raise vbObjectError + 3500, , _
        "Letakkan EWSSYNC_P26_Batch.ps1 di folder yang sama dengan workbook EWSSYNC."
    Set dlg = Application.FileDialog(msoFileDialogFilePicker)
    With dlg
        .Title = "Pilih file EWS PPh26 (dapat memilih 12 masa sekaligus)"
        .AllowMultiSelect = True
        .Filters.Clear
        .Filters.Add "EWS Excel", "*.xlsm"
        If .Show <> -1 Then Exit Sub
    End With
    If Not KonfirmasiTrustedLocationPPh26( _
            CStr(dlg.SelectedItems(1))) Then Exit Sub
    manifest = Environ$("TEMP") & Application.PathSeparator & _
        "EWSSYNC_P26_" & Format$(Now, "yyyymmdd_hhnnss") & ".txt"
    nomor = FreeFile
    Open manifest For Output As #nomor
    For Each item In dlg.SelectedItems
        pesan = vbNullString
        If Not BacaIdentitasDariNamaFile(CStr(item), kode, jenis, npwp, masa, awal, akhir, tahun, _
            ident, dikenal, didukung, handler, pesan) Then Err.Raise vbObjectError + 3501, , pesan
        If kode <> KODE_EWS_PPH26 Or handler <> HANDLER_PPH26 Then _
            Err.Raise vbObjectError + 3502, , "Daftar hanya boleh berisi EWS PPh26 TOT09."
        Print #nomor, CStr(item)
        jumlah = jumlah + 1
    Next item
    Close #nomor
    If jumlah = 0 Then Exit Sub
    cmd = "powershell.exe -NoProfile -NonInteractive -WindowStyle Hidden -File " & _
        KutipArgPPh26(helper) & " -Mode Batch -ExcelHwnd " & CStr(Application.Hwnd) & _
        " -HostWorkbook " & KutipArgPPh26(ThisWorkbook.FullName) & _
        " -ManifestPath " & KutipArgPPh26(manifest)
    CreateObject("WScript.Shell").Run cmd, 0, False
    MsgBox CStr(jumlah) & " file PPh26 mulai diproses berurutan." & vbCrLf & _
        "Excel dapat berkedip saat error EWS ditutup otomatis." & vbCrLf & _
        "Hasil tersedia di Log Impor EWS setelah proses selesai.", _
        vbInformation, "Impor batch PPh26"
    Exit Sub
Gagal:
    pesan = Err.Description
    On Error Resume Next
    If nomor > 0 Then Close #nomor
    If Len(manifest) > 0 Then Kill manifest
    On Error GoTo 0
    MsgBox pesan, vbExclamation, "Impor batch PPh26"
End Sub

'Bagian: pemeriksaan keamanan folder EWS.
'Tujuan: Mempertahankan nama prosedur yang dipakai alur PPh 26.
Public Function KonfirmasiTrustedLocationPPh26( _
    ByVal pathFile As String) As Boolean

    KonfirmasiTrustedLocationPPh26 = _
        KonfirmasiTrustedLocationEWS(pathFile, NAMA_EWS_PPH26)

End Function

'Tujuan: Memastikan folder EWS sudah menjadi Trusted Location.
Private Function KonfirmasiTrustedLocationEWS( _
    ByVal pathFile As String, _
    ByVal namaPajak As String) As Boolean

    Dim folderEWS As String
    Dim posisi As Long
    Dim jawaban As VbMsgBoxResult

    posisi = InStrRev(pathFile, Application.PathSeparator)
    If posisi = 0 Then
        folderEWS = pathFile
    Else
        folderEWS = Left$(pathFile, posisi - 1)
    End If

    If mFolderTrustedDikonfirmasi Is Nothing Then
        Set mFolderTrustedDikonfirmasi = _
            CreateObject("Scripting.Dictionary")
        mFolderTrustedDikonfirmasi.CompareMode = vbTextCompare
    End If

    If mFolderTrustedDikonfirmasi.Exists(folderEWS) Then
        KonfirmasiTrustedLocationEWS = True
        Exit Function
    End If

    jawaban = MsgBox( _
        "Impor " & namaPajak & _
        " memerlukan macro bawaan EWS untuk menjalankan rebuild." & _
        vbCrLf & vbCrLf & _
        "Pastikan folder berikut sudah menjadi Trusted Location di Excel:" & _
        vbCrLf & folderEWS & vbCrLf & vbCrLf & _
        "Excel: File > Options > Trust Center > Trust Center Settings > " & _
        "Trusted Locations > Add new location." & vbCrLf & _
        "Aktifkan juga Subfolders of this location are trusted bila diperlukan." & _
        vbCrLf & vbCrLf & _
        "Klik OK hanya jika folder sudah dipercaya. " & _
        "Klik Cancel untuk menghentikan impor dan mengatur Trusted Location.", _
        vbOKCancel + vbExclamation + vbDefaultButton2, _
        "Trusted Location EWS " & namaPajak)

    If jawaban = vbOK Then
        mFolderTrustedDikonfirmasi.Add folderEWS, True
        KonfirmasiTrustedLocationEWS = True
    End If

End Function

'Tujuan: Membuka ews pph 26 batch.
Public Sub BukaEwsPPh26Batch(ByVal path As String)
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.AskToUpdateLinks = True
    Application.Calculation = xlCalculationAutomatic
    Application.AutomationSecurity = msoAutomationSecurityByUI
    Workbooks.Open Filename:=path, ReadOnly:=False, _
        IgnoreReadOnlyRecommended:=True, AddToMru:=False, Notify:=False
End Sub

'Tujuan: Mengimpor pph 26 terbuka batch.
Public Function ImporPPh26TerbukaBatch(ByVal path As String) As String
    Dim kode As String, jenis As String, npwp As String, masa As String
    Dim awal As Long, akhir As Long, tahun As Long, ident As String
    Dim dikenal As Boolean, didukung As Boolean, handler As String
    Dim pesan As String, namaWP As String, ok As Boolean
    On Error GoTo Gagal
    If Not BacaIdentitasDariNamaFile(path, kode, jenis, npwp, masa, awal, akhir, tahun, _
        ident, dikenal, didukung, handler, pesan) Then GoTo Catat
    If kode <> KODE_EWS_PPH26 Or handler <> HANDLER_PPH26 Then
        pesan = "File bukan EWS PPh26 TOT09.": GoTo Catat
    End If
    ok = ImportPPh26PascaError(path, npwp, awal, akhir, tahun, namaWP, pesan)
Catat:
    TulisLogImpor path, kode, NamaPeriode(awal, akhir, tahun), _
        IIf(ok, "Berhasil", "Gagal"), pesan, npwp, awal, akhir, tahun
    If ok Then
        ImporPPh26TerbukaBatch = "OK"
    Else
        ImporPPh26TerbukaBatch = "GAGAL|" & Replace(Replace(pesan, vbCr, " "), vbLf, " ")
    End If
    Exit Function
Gagal:
    pesan = "Error " & CStr(Err.Number) & ": " & Err.Description
    Resume Catat
End Function
