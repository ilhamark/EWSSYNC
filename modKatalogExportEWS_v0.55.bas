Attribute VB_Name = "modKatalogExportEWS"
'-------------------------------------------------------------
'THIS MODULE IS FUELED BY SPITE OF EWS AND CORETAX
'@2026
'-------------------------------------------------------------
Option Explicit

'EWSSYNC v0.5 - katalog dan tampilan ekspor EWS.

'Tujuan: Menyediakan pilihan jenis pajak, file EWS, dan KKP pada form.

Private Const KODE_PPN_DN As String = "WP_VAT_VAT_TOT01"
Private Const KODE_PPH21 As String = "WP_ICT_WH_TOT01"
Private Const KODE_PPH21F As String = "WP_ICT_WH_TOT02"
Private Const KODE_PPH23 As String = "WP_ICT_WH_TOT08"
Private Const KODE_PPH26 As String = "WP_ICT_WH_TOT09"
Private Const KODE_PPH42 As String = "WP_ICT_WH_TOT10"

'---------------------------------------------------------------------
' MACRO PEMBUKA FORM
'---------------------------------------------------------------------
' Public Sub tanpa parameter pada Standard Module akan tampil pada
' daftar Macro Excel (Alt+F8).
Public Sub BukaFormExportEWS()

    frmExportEWS.Show

End Sub

'---------------------------------------------------------------------
' NAMA BULAN KHUSUS UNTUK FORM
'---------------------------------------------------------------------
' Fungsi ini sengaja ditempatkan pada modul katalog sebagai Public.
' Dengan demikian frmExportEWS tidak bergantung pada NamaBulanEkspor
' di modExportEWS, yang pada beberapa versi masih bersifat Private.
Public Function NamaBulanUntukFormExportEWS( _
    ByVal nomorBulan As Long) As String

    Dim daftarBulan As Variant

    daftarBulan = Array( _
        "", _
        "Januari", _
        "Februari", _
        "Maret", _
        "April", _
        "Mei", _
        "Juni", _
        "Juli", _
        "Agustus", _
        "September", _
        "Oktober", _
        "November", _
        "Desember")

    If nomorBulan >= 1 And nomorBulan <= 12 Then
        NamaBulanUntukFormExportEWS = daftarBulan(nomorBulan)
    Else
        NamaBulanUntukFormExportEWS = "Tidak diketahui"
    End If

End Function

'---------------------------------------------------------------------
' MENGISI PILIHAN JENIS PAJAK / JENIS EWS
'---------------------------------------------------------------------
Public Sub MuatJenisPajakExportEWS(ByVal frm As Object)

    With frm.cboJenisPajak

        .Clear
        .ColumnCount = 3
        .ColumnWidths = "180 pt;0 pt;0 pt"
        .BoundColumn = 2

        TambahJenisPajakExportEWS _
            frm.cboJenisPajak, _
            "PPN Dalam Negeri", _
            KODE_PPN_DN, _
            True

        TambahJenisPajakExportEWS _
            frm.cboJenisPajak, _
            "PPh Badan (belum aktif)", _
            "WP_ICT_CI_GTS_TOT01", _
            False

        TambahJenisPajakExportEWS _
            frm.cboJenisPajak, _
            "PPh Pasal 21", _
            KODE_PPH21, _
            True

        TambahJenisPajakExportEWS _
            frm.cboJenisPajak, _
            "PPh Pasal 21 Final", _
            KODE_PPH21F, _
            True

        TambahJenisPajakExportEWS _
            frm.cboJenisPajak, _
            "PPh Pasal 23", _
            KODE_PPH23, _
            True

        TambahJenisPajakExportEWS _
            frm.cboJenisPajak, _
            "PPh Pasal 26", _
            KODE_PPH26, _
            True

        TambahJenisPajakExportEWS _
            frm.cboJenisPajak, _
            "PPh Pasal 4 ayat (2)", _
            KODE_PPH42, _
            True

        If .ListCount > 0 Then .ListIndex = 0

    End With

End Sub

Private Sub TambahJenisPajakExportEWS( _
    ByVal cbo As Object, _
    ByVal namaJenis As String, _
    ByVal kodeJenis As String, _
    ByVal sudahAktif As Boolean)

    Dim indeks As Long

    cbo.AddItem namaJenis
    indeks = cbo.ListCount - 1

    cbo.List(indeks, 1) = kodeJenis
    cbo.List(indeks, 2) = sudahAktif

End Sub

'---------------------------------------------------------------------
' MEMBACA JENIS PAJAK YANG DIPILIH
'---------------------------------------------------------------------
Public Function KodeJenisPajakExportEWS( _
    ByVal frm As Object) As String

    If frm.cboJenisPajak.ListIndex < 0 Then Exit Function

    KodeJenisPajakExportEWS = CStr( _
        frm.cboJenisPajak.List( _
            frm.cboJenisPajak.ListIndex, 1))

End Function

Public Function NamaJenisPajakExportEWS( _
    ByVal frm As Object) As String

    If frm.cboJenisPajak.ListIndex < 0 Then Exit Function

    NamaJenisPajakExportEWS = CStr( _
        frm.cboJenisPajak.List( _
            frm.cboJenisPajak.ListIndex, 0))

End Function

Public Function JenisPajakAktifExportEWS( _
    ByVal frm As Object) As Boolean

    If frm.cboJenisPajak.ListIndex < 0 Then Exit Function

    JenisPajakAktifExportEWS = CBool( _
        frm.cboJenisPajak.List( _
            frm.cboJenisPajak.ListIndex, 2))

End Function

'---------------------------------------------------------------------
' MENGISI FILTER KELOMPOK KKP
'---------------------------------------------------------------------
Public Sub MuatKelompokKKPExportEWS(ByVal frm As Object)

    Dim kodeJenis As String

    kodeJenis = KodeJenisPajakExportEWS(frm)

    With frm.cboKelompokKKP

        .Clear
        .ColumnCount = 2
        .ColumnWidths = "115 pt;0 pt"
        .BoundColumn = 2

        If kodeJenis = KODE_PPN_DN Then

            TambahKelompokExportEWS _
                frm.cboKelompokKKP, "Semua kelompok", "*"

            TambahKelompokExportEWS _
                frm.cboKelompokKKP, "N - Penghitungan PPN", "N"

            TambahKelompokExportEWS _
                frm.cboKelompokKKP, "O - Pembayaran", "O"

            TambahKelompokExportEWS _
                frm.cboKelompokKKP, "P - Ketetapan", "P"

            TambahKelompokExportEWS _
                frm.cboKelompokKKP, "Q - Pemungutan", "Q"

        ElseIf kodeJenis = KODE_PPH21 Or _
               kodeJenis = KODE_PPH21F Or _
               kodeJenis = KODE_PPH23 Or _
               kodeJenis = KODE_PPH26 Or _
               kodeJenis = KODE_PPH42 Then

            TambahKelompokExportEWS _
                frm.cboKelompokKKP, "Semua kelompok", "*"

            TambahKelompokExportEWS _
                frm.cboKelompokKKP, "B - Dasar Pengenaan/PPh", "B"

            TambahKelompokExportEWS _
                frm.cboKelompokKKP, "C - PPh Tidak Terutang", "C"

            TambahKelompokExportEWS _
                frm.cboKelompokKKP, "N - Kredit/Pembayaran", "N"

            TambahKelompokExportEWS _
                frm.cboKelompokKKP, "O - Diperhitungkan", "O"

            TambahKelompokExportEWS _
                frm.cboKelompokKKP, "P - PPh Tidak Terutang", "P"

        End If

        If .ListCount > 0 Then .ListIndex = 0

    End With

End Sub

Private Sub TambahKelompokExportEWS( _
    ByVal cbo As Object, _
    ByVal namaKelompok As String, _
    ByVal kodeKelompok As String)

    Dim indeks As Long

    cbo.AddItem namaKelompok
    indeks = cbo.ListCount - 1
    cbo.List(indeks, 1) = kodeKelompok

End Sub

Private Function KodeKelompokKKPExportEWS( _
    ByVal frm As Object) As String

    If frm.cboKelompokKKP.ListIndex < 0 Then
        KodeKelompokKKPExportEWS = "*"
        Exit Function
    End If

    KodeKelompokKKPExportEWS = CStr( _
        frm.cboKelompokKKP.List( _
            frm.cboKelompokKKP.ListIndex, 1))

End Function

'---------------------------------------------------------------------
' MENGISI DAFTAR KKP SESUAI JENIS PAJAK DAN KELOMPOK
'---------------------------------------------------------------------
Public Sub MuatDaftarKKPExportEWS(ByVal frm As Object)

    Dim daftarKode As Variant
    Dim kodeJenis As String
    Dim kodeKelompok As String
    Dim kodeKKP As Variant
    Dim judulKKP As String
    Dim indeks As Long
    Dim wsKKP As Worksheet
    Dim taxKey As String

    kodeJenis = KodeJenisPajakExportEWS(frm)
    kodeKelompok = KodeKelompokKKPExportEWS(frm)
    taxKey = TaxKeyDariKodeJenisExportEWS(kodeJenis)

    With frm.lstKKP

        .Clear
        .ColumnCount = 4
        .ColumnWidths = "42 pt;310 pt;0 pt;0 pt"

    End With

    If Not JenisPajakAktifExportEWS(frm) Then
        PerbaruiRingkasanFormExportEWS frm
        Exit Sub
    End If

    daftarKode = DaftarKodeKKPExportEWS(kodeJenis)

    If IsEmpty(daftarKode) Then
        PerbaruiRingkasanFormExportEWS frm
        Exit Sub
    End If

    For Each kodeKKP In daftarKode

        If kodeKelompok = "*" Or _
           Left$(CStr(kodeKKP), 1) = kodeKelompok Then

            Set wsKKP = DapatkanSheetLokalKKP( _
                taxKey, _
                CStr(kodeKKP), _
                True)

            If Not wsKKP Is Nothing Then

                judulKKP = Trim$(CStr( _
                    wsKKP.Range("A7").Value2))

                If Len(judulKKP) = 0 Then
                    judulKKP = "Judul KKP tidak ditemukan"
                End If

                frm.lstKKP.AddItem CStr(kodeKKP)
                indeks = frm.lstKKP.ListCount - 1

                frm.lstKKP.List(indeks, 1) = judulKKP
                frm.lstKKP.List(indeks, 2) = CStr(kodeKKP)
                frm.lstKKP.List(indeks, 3) = kodeJenis

            End If

        End If

    Next kodeKKP

    PerbaruiRingkasanFormExportEWS frm

End Sub

Private Function DaftarKodeKKPExportEWS( _
    ByVal kodeJenis As String) As Variant

    If kodeJenis = KODE_PPN_DN Then

        DaftarKodeKKPExportEWS = DaftarKodeEWSPPN()

    ElseIf kodeJenis = KODE_PPH21 Then

        DaftarKodeKKPExportEWS = DaftarKodeEWSPPh21()

    ElseIf kodeJenis = KODE_PPH21F Then

        DaftarKodeKKPExportEWS = DaftarKodeEWSPPh21F()

    ElseIf kodeJenis = KODE_PPH23 Then

        DaftarKodeKKPExportEWS = DaftarKodeEWSPPh23()

    ElseIf kodeJenis = KODE_PPH26 Then

        DaftarKodeKKPExportEWS = DaftarKodeEWSPPh26()

    ElseIf kodeJenis = KODE_PPH42 Then

        DaftarKodeKKPExportEWS = DaftarKodeEWSPPh42()

    Else

        DaftarKodeKKPExportEWS = Empty

    End If

End Function

' Mengubah identifier pada nama file EWS menjadi namespace sheet lokal.
' Fungsi ini sengaja terpusat agar katalog tidak membentuk nama sheet
' sendiri dan kode EWS yang sama tidak bertabrakan antarjenis pajak.
Private Function TaxKeyDariKodeJenisExportEWS( _
    ByVal kodeJenis As String) As String

    Select Case UCase$(Trim$(kodeJenis))

        Case KODE_PPN_DN
            TaxKeyDariKodeJenisExportEWS = TAXKEY_PPN_DN

        Case KODE_PPH21
            TaxKeyDariKodeJenisExportEWS = TAXKEY_PPH21

        Case KODE_PPH21F
            TaxKeyDariKodeJenisExportEWS = TAXKEY_PPH21F

        Case KODE_PPH23
            TaxKeyDariKodeJenisExportEWS = TAXKEY_PPH23

        Case KODE_PPH26
            TaxKeyDariKodeJenisExportEWS = TAXKEY_PPH26

        Case KODE_PPH42
            TaxKeyDariKodeJenisExportEWS = TAXKEY_PPH42

    End Select

End Function

Private Function SheetAdaUntukExportEWS( _
    ByVal namaSheet As String) As Boolean

    'Dipertahankan untuk kompatibilitas kode lama. Katalog baru memakai
    'DapatkanSheetLokalKKP dengan taxKey yang sesuai jenis EWS.
    SheetAdaUntukExportEWS = SheetLokalKKPAda( _
        TAXKEY_PPN_DN, _
        namaSheet, _
        True)

End Function

'---------------------------------------------------------------------
' MEMILIH FILE EWS TANPA BERGANTUNG PADA modExportEWS
'---------------------------------------------------------------------
' Tahap form hanya membaca identifier dan masa dari nama file.
' Workbook EWS tidak dibuka. Pemeriksaan struktur dan identitas WP akan
' dilakukan oleh engine ekspor pada tahap validasi sebelum penulisan.
Public Sub PilihFileUntukFormExportEWS(ByVal frm As Object)

    Dim filePath As Variant
    Dim namaFile As String
    Dim kodeJenis As String
    Dim bulanFile As Long
    Dim tahunFile As Long
    Dim pesanDitolak As String

    If Not JenisPajakAktifExportEWS(frm) Then

        MsgBox _
            "Mapping jenis pajak ini belum diaktifkan.", _
            vbInformation, _
            "Pilih File EWS"

        Exit Sub

    End If

    kodeJenis = KodeJenisPajakExportEWS(frm)

    With Application.FileDialog(msoFileDialogFilePicker)

        .Title = "Pilih File EWS"
        .AllowMultiSelect = True

        .Filters.Clear
        .Filters.Add "File Excel Macro", "*.xlsm"
        .Filters.Add _
            "Semua File Excel", _
            "*.xlsx;*.xlsm;*.xlsb"

        If .Show <> -1 Then Exit Sub

        For Each filePath In .SelectedItems

            namaFile = Dir$(CStr(filePath))

            If InStr( _
                1, _
                namaFile, _
                kodeJenis, _
                vbTextCompare) = 0 Then

                pesanDitolak = pesanDitolak & _
                    vbCrLf & "- " & namaFile & _
                    " (jenis EWS tidak sesuai)"

                GoTo FileBerikut
            End If

            bulanFile = 0
            tahunFile = 0

            If Not BacaPeriodeNamaFileUntukFormExportEWS( _
                    namaFile, _
                    bulanFile, _
                    tahunFile) Then

                pesanDitolak = pesanDitolak & _
                    vbCrLf & "- " & namaFile & _
                    " (masa tidak ditemukan)"

                GoTo FileBerikut
            End If

            frm.lstEWS.List(bulanFile - 1, 1) = _
                NamaFilePendekUntukFormExportEWS( _
                    kodeJenis, bulanFile, tahunFile, namaFile)
            frm.lstEWS.List(bulanFile - 1, 2) = "Siap"
            frm.lstEWS.List(bulanFile - 1, 3) = CStr(filePath)
            frm.lstEWS.List(bulanFile - 1, 4) = bulanFile
            frm.lstEWS.List(bulanFile - 1, 5) = tahunFile

FileBerikut:

        Next filePath

    End With

    PerbaruiIdentitasFileFormExportEWS frm
    PerbaruiRingkasanFormExportEWS frm

    If Len(pesanDitolak) > 0 Then

        MsgBox _
            "File berikut tidak dimasukkan:" & _
            pesanDitolak, _
            vbExclamation, _
            "Sebagian File Dilewati"

    End If

End Sub

' Mencari token delapan angka pada nama file.
' Untuk PPN Dalam Negeri format aktifnya adalah MMDDYYYY dengan MM = DD,
' misalnya 01012025 untuk Januari 2025.
Private Function BacaPeriodeNamaFileUntukFormExportEWS( _
    ByVal namaFile As String, _
    ByRef bulan As Long, _
    ByRef tahun As Long) As Boolean

    Dim bagian As Variant
    Dim token As Variant
    Dim teksToken As String
    Dim bulanAwal As Long
    Dim bulanAkhir As Long
    Dim tahunToken As Long

    bagian = Split(namaFile, "_")

    For Each token In bagian

        teksToken = CStr(token)

        If InStrRev(teksToken, ".") > 0 Then
            teksToken = Left$( _
                teksToken, _
                InStrRev(teksToken, ".") - 1)
        End If

        If Len(teksToken) = 8 And IsNumeric(teksToken) Then

            bulanAwal = CLng(Left$(teksToken, 2))
            bulanAkhir = CLng(Mid$(teksToken, 3, 2))
            tahunToken = CLng(Right$(teksToken, 4))

            If bulanAwal >= 1 And bulanAwal <= 12 And _
               bulanAkhir >= 1 And bulanAkhir <= 12 And _
               bulanAwal = bulanAkhir And _
               tahunToken >= 1900 Then

                bulan = bulanAwal
                tahun = tahunToken

                BacaPeriodeNamaFileUntukFormExportEWS = True
                Exit Function

            End If

        End If

    Next token

End Function



'Nama pendek hanya untuk tampilan. Path lengkap tetap berada pada kolom
'tersembunyi lstEWS(,3), sehingga proses ekspor tidak berubah.
Private Function NamaFilePendekUntukFormExportEWS( _
    ByVal kodeJenis As String, _
    ByVal bulan As Long, _
    ByVal tahun As Long, _
    ByVal namaAsli As String) As String

    Dim ekstensi As String
    Dim posisiTitik As Long

    posisiTitik = InStrRev(namaAsli, ".")
    If posisiTitik > 0 Then
        ekstensi = Mid$(namaAsli, posisiTitik)
    Else
        ekstensi = ".xlsm"
    End If

    NamaFilePendekUntukFormExportEWS = _
        kodeJenis & "_" & _
        Format$(bulan, "00") & Format$(bulan, "00") & _
        Format$(tahun, "0000") & ekstensi

End Function

'Ambil nama WP dari INFO tanpa membuka seluruh file EWS pada tahap pemilihan.
'D16 adalah kontrak workbook lokal saat ini; pencarian label menjadi fallback.
Private Function NamaWPLokalUntukFormExportEWS() As String

    Dim ws As Worksheet
    Dim sel As Range
    Dim c As Long
    Dim teks As String
    Dim kandidat As String

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("INFO")
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    kandidat = Trim$(CStr(ws.Range("D16").Value2))
    If NamaWPTampilanValidExportEWS(kandidat) Then
        NamaWPLokalUntukFormExportEWS = kandidat
        Exit Function
    End If

    For Each sel In ws.UsedRange.Cells
        teks = UCase$(Trim$(CStr(sel.Value2)))
        teks = Replace(teks, " ", vbNullString)
        teks = Replace(teks, ".", vbNullString)
        teks = Replace(teks, ":", vbNullString)

        If teks = "NAMAWP" Or teks = "NAMAWAJIBPAJAK" Then
            For c = sel.Column + 1 To sel.Column + 3
                kandidat = Trim$(CStr(ws.Cells(sel.Row, c).Value2))
                If NamaWPTampilanValidExportEWS(kandidat) Then
                    NamaWPLokalUntukFormExportEWS = kandidat
                    Exit Function
                End If
            Next c
        End If
    Next sel

End Function

Private Function NamaWPTampilanValidExportEWS(ByVal teks As String) As Boolean
    Dim i As Long, ch As String, punyaHuruf As Boolean
    teks = Trim$(teks)
    If Len(teks) = 0 Then Exit Function
    For i = 1 To Len(teks)
        ch = Mid$(teks, i, 1)
        If UCase$(ch) Like "[A-Z]" Then punyaHuruf = True: Exit For
    Next i
    NamaWPTampilanValidExportEWS = punyaHuruf
End Function

'---------------------------------------------------------------------
' MENGOSONGKAN DAFTAR FILE PADA FORM
'---------------------------------------------------------------------
Public Sub KosongkanFilePadaFormExportEWS(ByVal frm As Object)

    Dim i As Long

    For i = 0 To frm.lstEWS.ListCount - 1

        frm.lstEWS.List(i, 1) = vbNullString
        frm.lstEWS.List(i, 2) = "Belum dipilih"
        frm.lstEWS.List(i, 3) = vbNullString
        frm.lstEWS.List(i, 4) = i + 1
        frm.lstEWS.List(i, 5) = 0

    Next i

    frm.NamaWPReferensi = vbNullString
    frm.NPWPReferensi = vbNullString
    frm.TahunReferensi = 0

    frm.lblWPValue.Caption = "-"
    frm.lblTahunValue.Caption = "-"
    frm.lblJumlah.Caption = "0 dari 12 masa siap"

    PerbaruiRingkasanFormExportEWS frm

End Sub

' Menampilkan jumlah file dan rentang tahun dari nama file.
'Nama WP ditampilkan dari INFO workbook lokal; identitas file tetap divalidasi engine.
Private Sub PerbaruiIdentitasFileFormExportEWS(ByVal frm As Object)

    Dim i As Long
    Dim jumlahFile As Long
    Dim tahunFile As Long
    Dim tahunAwal As Long
    Dim tahunAkhir As Long

    For i = 0 To frm.lstEWS.ListCount - 1

        If Len(Trim$(CStr(frm.lstEWS.List(i, 3)))) > 0 Then

            jumlahFile = jumlahFile + 1
            tahunFile = CLng(frm.lstEWS.List(i, 5))

            If tahunAwal = 0 Or tahunFile < tahunAwal Then
                tahunAwal = tahunFile
            End If

            If tahunFile > tahunAkhir Then
                tahunAkhir = tahunFile
            End If

        End If

    Next i

    frm.lblJumlah.Caption = _
        CStr(jumlahFile) & " dari 12 masa siap"

    If jumlahFile > 0 Then
        frm.NamaWPReferensi = NamaWPLokalUntukFormExportEWS()
        If Len(frm.NamaWPReferensi) > 0 Then
            frm.lblWPValue.Caption = frm.NamaWPReferensi
        Else
            frm.lblWPValue.Caption = "-"
        End If
    Else
        frm.NamaWPReferensi = vbNullString
        frm.lblWPValue.Caption = "-"
    End If

    If tahunAwal = 0 Then

        frm.lblTahunValue.Caption = "-"
        frm.TahunReferensi = 0

    ElseIf tahunAwal = tahunAkhir Then

        frm.lblTahunValue.Caption = CStr(tahunAwal)
        frm.TahunReferensi = tahunAwal

    Else

        frm.lblTahunValue.Caption = _
            CStr(tahunAwal) & " - " & CStr(tahunAkhir)

        frm.TahunReferensi = 0

    End If

End Sub

'---------------------------------------------------------------------
' PILIHAN KKP
'---------------------------------------------------------------------
Public Sub PilihSemuaKKPExportEWS(ByVal frm As Object)

    Dim i As Long

    For i = 0 To frm.lstKKP.ListCount - 1
        frm.lstKKP.Selected(i) = True
    Next i

    PerbaruiRingkasanFormExportEWS frm

End Sub

Public Sub KosongkanPilihanKKPExportEWS(ByVal frm As Object)

    Dim i As Long

    For i = 0 To frm.lstKKP.ListCount - 1
        frm.lstKKP.Selected(i) = False
    Next i

    PerbaruiRingkasanFormExportEWS frm

End Sub

Public Function JumlahKKPTerpilihExportEWS( _
    ByVal frm As Object) As Long

    Dim i As Long

    For i = 0 To frm.lstKKP.ListCount - 1

        If frm.lstKKP.Selected(i) Then
            JumlahKKPTerpilihExportEWS = _
                JumlahKKPTerpilihExportEWS + 1
        End If

    Next i

End Function

' Fungsi ini menjadi antarmuka bagi engine ekspor tahap berikutnya.
' Hasilnya berupa daftar kode KKP yang dipisahkan karakter |.
Public Function DaftarKodeKKPTerpilihExportEWS( _
    ByVal frm As Object) As String

    Dim i As Long
    Dim hasil As String

    For i = 0 To frm.lstKKP.ListCount - 1

        If frm.lstKKP.Selected(i) Then

            If Len(hasil) > 0 Then hasil = hasil & "|"
            hasil = hasil & CStr(frm.lstKKP.List(i, 2))

        End If

    Next i

    DaftarKodeKKPTerpilihExportEWS = hasil

End Function

'---------------------------------------------------------------------
' STATUS DAN VALIDASI FORM
'---------------------------------------------------------------------
Public Sub PerbaruiRingkasanFormExportEWS(ByVal frm As Object)

    Dim jumlahFile As Long
    Dim jumlahKKP As Long
    Dim i As Long

    For i = 0 To frm.lstEWS.ListCount - 1

        If Len(Trim$(CStr(frm.lstEWS.List(i, 3)))) > 0 Then
            jumlahFile = jumlahFile + 1
        End If

    Next i

    jumlahKKP = JumlahKKPTerpilihExportEWS(frm)

    frm.lblJumlahKKP.Caption = _
        CStr(jumlahKKP) & " dari " & _
        CStr(frm.lstKKP.ListCount) & " KKP dipilih"

    frm.cmdExport.Enabled = False

    If frm.cboJenisPajak.ListIndex < 0 Then

        frm.lblStatus.Caption = "Pilih jenis pajak/EWS."

    ElseIf Not JenisPajakAktifExportEWS(frm) Then

        frm.lblStatus.Caption = _
            "Jenis pajak ini sudah disiapkan, tetapi mapping belum aktif."

    ElseIf jumlahFile = 0 Then

        frm.lblStatus.Caption = _
            "Pilih satu atau beberapa file EWS."

    ElseIf jumlahKKP = 0 Then

        frm.lblStatus.Caption = _
            "Pilih sekurang-kurangnya satu KKP yang akan diekspor."

    Else

        frm.lblStatus.Caption = _
            "Pilihan lengkap. Siap untuk tahap validasi engine ekspor."

        frm.cmdExport.Enabled = True

    End If

End Sub

Public Function ValidasiPilihanFormExportEWS( _
    ByVal frm As Object) As Boolean

    If frm.cboJenisPajak.ListIndex < 0 Then

        MsgBox _
            "Pilih jenis pajak/EWS terlebih dahulu.", _
            vbExclamation, _
            "Ekspor EWS"

        Exit Function

    End If

    If Not JenisPajakAktifExportEWS(frm) Then

        MsgBox _
            "Mapping untuk jenis pajak ini belum diaktifkan.", _
            vbInformation, _
            "Ekspor EWS"

        Exit Function

    End If

    If JumlahKKPTerpilihExportEWS(frm) = 0 Then

        MsgBox _
            "Pilih sekurang-kurangnya satu KKP.", _
            vbExclamation, _
            "Ekspor EWS"

        Exit Function

    End If

    ValidasiPilihanFormExportEWS = True

End Function

'---------------------------------------------------------------------
' PENGAMAN TAHAP PENGEMBANGAN
'---------------------------------------------------------------------
' Menghubungkan form ke engine ekspor selektif.
Public Sub JalankanExportTerpilihDariForm(ByVal frm As Object)

    If Not ValidasiPilihanFormExportEWS(frm) Then Exit Sub

    EksporTerpilihEWS frm

End Sub
