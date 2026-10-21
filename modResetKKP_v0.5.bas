Attribute VB_Name = "modResetKKP"
'-------------------------------------------------------------
'THIS MODULE IS FUELED BY SPITE OF EWS AND CORETAX
'@2026
'-------------------------------------------------------------
Option Explicit


'EWSSYNC v0.5 - reset nilai KKP hasil impor EWS.

Private Const NAMA_LOG_IMPOR_RESET As String = "Log Impor EWS"
Private Const BARIS_HEADER_MAKS_RESET As Long = 40

'Bagian: proses utama.
'Tujuan: Menghapus seluruh nilai KKP hasil impor dan riwayat impor.
Public Sub ResetDataImporKKP()

    Dim daftarSel As Collection
    Dim daftarNilai As Collection
    Dim daftarTaxKey As Variant
    Dim taxKey As Variant
    Dim jumlahNilai As Long
    Dim jumlahLog As Long
    Dim jawaban As VbMsgBoxResult
    Dim calcLama As XlCalculation
    Dim eventsLama As Boolean
    Dim screenLama As Boolean
    Dim statusLama As Variant
    Dim logLama As Variant
    Dim areaLog As Range
    Dim i As Long
    Dim tahap As String
    Dim pesanError As String
    Dim statusAplikasiDisimpan As Boolean

    On Error GoTo Gagal

    Set daftarSel = New Collection
    Set daftarNilai = New Collection

    daftarTaxKey = Array( _
        TAXKEY_PPN_DN, TAXKEY_PPH21, TAXKEY_PPH21F, TAXKEY_PPH23, _
        TAXKEY_PPH26, TAXKEY_PPH42)

    tahap = "memetakan nilai hasil impor"

    For Each taxKey In daftarTaxKey
        KumpulkanTargetReset _
            CStr(taxKey), daftarSel, daftarNilai
    Next taxKey

    KumpulkanTargetResetPPh21FDetail daftarSel, daftarNilai

    jumlahNilai = daftarSel.Count
    Set areaLog = AreaDataLogImporReset()
    If Not areaLog Is Nothing Then jumlahLog = areaLog.Rows.Count

    If jumlahNilai = 0 And jumlahLog = 0 Then
        MsgBox "Tidak ada data hasil impor yang perlu direset.", _
            vbInformation, "Reset Data KKP"
        Exit Sub
    End If

    jawaban = MsgBox( _
        "Reset akan menghapus:" & vbCrLf & _
        "- " & jumlahNilai & " nilai pada blok Menurut Wajib Pajak" & vbCrLf & _
        "- " & jumlahLog & " baris riwayat pada Log Impor EWS" & vbCrLf & vbCrLf & _
        "Formula, format, hyperlink, dropdown, dan header log tetap dipertahankan." & _
        vbCrLf & vbCrLf & "Lanjutkan reset?", _
        vbQuestion + vbYesNo + vbDefaultButton2, _
        "Reset Data KKP")

    If jawaban <> vbYes Then Exit Sub

    calcLama = Application.Calculation
    eventsLama = Application.EnableEvents
    screenLama = Application.ScreenUpdating
    statusLama = Application.StatusBar
    statusAplikasiDisimpan = True

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Mereset data KKP hasil impor..."

    tahap = "menghapus nilai KKP"

    For i = 1 To daftarSel.Count
        daftarSel(i).ClearContents
    Next i

    tahap = "menghapus riwayat impor"

    If Not areaLog Is Nothing Then
        logLama = areaLog.Formula
        areaLog.ClearContents
    End If

    tahap = "menghitung ulang KKP"

    For Each taxKey In daftarTaxKey
        HitungUlangKKPLokal CStr(taxKey)
    Next taxKey

    PulihkanAplikasiReset _
        calcLama, eventsLama, screenLama, statusLama

    MsgBox "Reset selesai." & vbCrLf & _
        jumlahNilai & " nilai KKP dan " & jumlahLog & _
        " baris riwayat impor telah dihapus.", _
        vbInformation, "Reset Data KKP"

    Exit Sub

Gagal:

    pesanError = Err.Description
    On Error Resume Next

    For i = 1 To daftarSel.Count
        daftarSel(i).Value2 = daftarNilai(i)
    Next i

    If Not areaLog Is Nothing Then
        If Not IsEmpty(logLama) Then areaLog.Formula = logLama
    End If

    If statusAplikasiDisimpan Then
        PulihkanAplikasiReset _
            calcLama, eventsLama, screenLama, statusLama
    End If

    On Error GoTo 0

    MsgBox "Reset dibatalkan pada tahap " & tahap & "." & vbCrLf & _
        pesanError, vbCritical, "Reset Data KKP"

End Sub

'Bagian: pemetaan target.
'Tujuan: Mengumpulkan sel konstanta pada blok impor satu kelompok pajak.
Private Sub KumpulkanTargetReset( _
    ByVal taxKey As String, _
    ByVal daftarSel As Collection, _
    ByVal daftarNilai As Collection)

    Dim ws As Worksheet
    Dim namaSheet As String
    Dim kolomPeriode As Object
    Dim barisAwal As Long
    Dim barisAkhir As Long
    Dim kolom As Variant
    Dim baris As Long
    Dim sel As Range

    namaSheet = NamaSheetRingkasanPajak(taxKey)

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(namaSheet)
    On Error GoTo 0

    If ws Is Nothing Then Exit Sub

    Set kolomPeriode = KolomPeriodeReset(ws)
    If kolomPeriode.Count = 0 Then
        Err.Raise vbObjectError + 5901, , _
            "Kolom periode tidak ditemukan pada sheet '" & ws.Name & "'."
    End If

    If Not BlokMenurutWPReset(ws, barisAwal, barisAkhir) Then
        Err.Raise vbObjectError + 5902, , _
            "Blok Menurut Wajib Pajak tidak ditemukan pada sheet '" & _
            ws.Name & "'."
    End If

    For Each kolom In kolomPeriode.Items
        For baris = barisAwal To barisAkhir

            Set sel = ws.Cells(baris, CLng(kolom))

            If Not sel.HasFormula And _
               Not IsEmpty(sel.Value2) Then

                ValidasiSelReset sel
                daftarSel.Add sel
                daftarNilai.Add sel.Value2

            End If

        Next baris
    Next kolom

End Sub

'Tujuan: Mengumpulkan nilai impor pada matriks horizontal PPh 21 Final.
Private Sub KumpulkanTargetResetPPh21FDetail( _
    ByVal daftarSel As Collection, _
    ByVal daftarNilai As Collection)

    Dim ws As Worksheet
    Dim kolomMasa As Object
    Dim daftarKode As Variant
    Dim kode As Variant
    Dim kolom As Variant
    Dim sel As Range
    Dim nilai As Variant
    Dim teks As String
    Dim barisKode As Long
    Dim jumlahKode As Long
    Dim r As Long
    Dim c As Long
    Dim bulan As Long

    Set ws = DapatkanSheetLokalKKP( _
        TAXKEY_PPH21F, "B-1-3", False)
    If ws Is Nothing Then Exit Sub

    Set kolomMasa = CreateObject("Scripting.Dictionary")

    For c = 1 To KolomTerakhirReset(ws)
        nilai = ws.Cells(21, c).Value2
        If Not IsError(nilai) And IsNumeric(nilai) Then
            bulan = CLng(nilai)
            If bulan >= 1 And bulan <= 12 Then
                If kolomMasa.Exists(CStr(bulan)) Then
                    Err.Raise vbObjectError + 5905, , _
                        "Kolom masa ganda pada sheet '" & ws.Name & "'."
                End If
                kolomMasa.Add CStr(bulan), c
            End If
        End If
    Next c

    If kolomMasa.Count <> 12 Then
        Err.Raise vbObjectError + 5906, , _
            "Dua belas kolom masa tidak ditemukan pada sheet '" & _
            ws.Name & "'."
    End If

    daftarKode = Array( _
        "21-100-01", "21-100-02", _
        "21-100-03", "21-100-07")

    For Each kode In daftarKode
        barisKode = 0
        jumlahKode = 0

        For r = 1 To BarisTerakhirReset(ws)
            teks = UCase$(Trim$(TeksReset(ws.Cells(r, 2).Value2)))
            If InStr(1, teks, CStr(kode), vbTextCompare) > 0 Then
                barisKode = r
                jumlahKode = jumlahKode + 1
            End If
        Next r

        If jumlahKode <> 1 Then
            Err.Raise vbObjectError + 5907, , _
                "Kode objek " & CStr(kode) & _
                " tidak ditemukan atau ganda pada sheet '" & ws.Name & "'."
        End If

        If UCase$(Trim$(TeksReset(ws.Cells(barisKode, 3).Value2))) <> "DPP" Or _
           UCase$(Trim$(TeksReset(ws.Cells(barisKode + 1, 3).Value2))) <> "PPH" Then
            Err.Raise vbObjectError + 5908, , _
                "Pasangan DPP/PPh tidak sesuai untuk " & CStr(kode) & "."
        End If

        For Each kolom In kolomMasa.Items
            For r = barisKode To barisKode + 1
                Set sel = ws.Cells(r, CLng(kolom))

                If Not sel.HasFormula And Not IsEmpty(sel.Value2) Then
                    ValidasiSelReset sel
                    daftarSel.Add sel
                    daftarNilai.Add sel.Value2
                End If
            Next r
        Next kolom
    Next kode

End Sub

'Tujuan: Memvalidasi bahwa sel hasil impor dapat dikosongkan.
Private Sub ValidasiSelReset(ByVal sel As Range)

    If sel.MergeCells Then
        Err.Raise vbObjectError + 5903, , _
            "Sel hasil impor tergabung: " & _
            sel.Worksheet.Name & "!" & sel.Address(False, False)
    End If

    If sel.Worksheet.ProtectContents And sel.Locked Then
        Err.Raise vbObjectError + 5904, , _
            "Sel hasil impor terkunci: " & _
            sel.Worksheet.Name & "!" & sel.Address(False, False)
    End If

End Sub

'Tujuan: Menemukan blok Menurut Wajib Pajak pada sheet ringkasan.
Private Function BlokMenurutWPReset( _
    ByVal ws As Worksheet, _
    ByRef barisAwal As Long, _
    ByRef barisAkhir As Long) As Boolean

    Dim marker As Range
    Dim barisTerakhir As Long
    Dim r As Long
    Dim c As Long
    Dim teks As String

    Set marker = ws.Cells.Find( _
        What:="Menurut Wajib Pajak", _
        After:=ws.Cells(ws.Rows.Count, ws.Columns.Count), _
        LookIn:=xlValues, LookAt:=xlWhole, _
        SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False)

    If marker Is Nothing Then Exit Function

    barisAwal = marker.Row + 1
    barisTerakhir = BarisTerakhirReset(ws)
    barisAkhir = barisTerakhir

    For r = barisAwal To barisTerakhir
        For c = 1 To 5
            teks = UCase$(Trim$(TeksReset(ws.Cells(r, c).Value2)))
            If Left$(teks, 8) = "MENURUT " Then
                barisAkhir = r - 1
                Exit For
            End If
        Next c
        If barisAkhir < barisTerakhir Then Exit For
    Next r

    BlokMenurutWPReset = (barisAwal <= barisAkhir)

End Function

'Tujuan: Menemukan seluruh kolom periode bulanan pada sheet ringkasan.
Private Function KolomPeriodeReset(ByVal ws As Worksheet) As Object

    Dim hasil As Object
    Dim kandidat As Object
    Dim terbaik As Object
    Dim barisMaks As Long
    Dim kolomMaks As Long
    Dim r As Long
    Dim c As Long
    Dim bulan As Long

    Set hasil = CreateObject("Scripting.Dictionary")
    Set terbaik = CreateObject("Scripting.Dictionary")

    barisMaks = BarisTerakhirReset(ws)
    If barisMaks > BARIS_HEADER_MAKS_RESET Then
        barisMaks = BARIS_HEADER_MAKS_RESET
    End If

    kolomMaks = KolomTerakhirReset(ws)

    For r = 1 To barisMaks

        Set kandidat = CreateObject("Scripting.Dictionary")

        For c = 1 To kolomMaks
            bulan = NomorBulanReset(ws.Cells(r, c).Value)
            If bulan > 0 Then kandidat(CStr(bulan)) = c
        Next c

        If kandidat.Count > terbaik.Count Then
            Set terbaik = kandidat
        End If

    Next r

    If terbaik.Count = 12 Then Set hasil = terbaik
    Set KolomPeriodeReset = hasil

End Function

'Tujuan: Membaca nomor bulan dari nilai header periode.
Private Function NomorBulanReset(ByVal nilai As Variant) As Long

    Dim teks As String
    Dim daftar As Variant
    Dim i As Long

    If IsError(nilai) Or IsEmpty(nilai) Then Exit Function

    If VarType(nilai) = vbDate Then
        NomorBulanReset = Month(CDate(nilai))
        Exit Function
    End If

    teks = UCase$(Trim$(TeksReset(nilai)))
    teks = Replace(teks, ".", vbNullString)

    daftar = Array( _
        "JANUARI", "FEBRUARI", "MARET", "APRIL", _
        "MEI", "JUNI", "JULI", "AGUSTUS", _
        "SEPTEMBER", "OKTOBER", "NOVEMBER", "DESEMBER")

    For i = LBound(daftar) To UBound(daftar)
        If teks = CStr(daftar(i)) Or _
           Left$(teks, 3) = Left$(CStr(daftar(i)), 3) Then
            NomorBulanReset = i + 1
            Exit Function
        End If
    Next i

End Function

'Bagian: log dan utilitas.
'Tujuan: Mengembalikan area data log tanpa menyertakan header.
Private Function AreaDataLogImporReset() As Range

    Dim ws As Worksheet
    Dim barisAkhir As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(NAMA_LOG_IMPOR_RESET)
    On Error GoTo 0

    If ws Is Nothing Then Exit Function

    barisAkhir = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If barisAkhir < 2 Then Exit Function

    Set AreaDataLogImporReset = ws.Range("A2:M" & barisAkhir)

End Function

'Tujuan: Mengembalikan baris terakhir yang digunakan.
Private Function BarisTerakhirReset(ByVal ws As Worksheet) As Long

    Dim sel As Range

    Set sel = ws.Cells.Find( _
        What:="*", After:=ws.Cells(1, 1), _
        LookIn:=xlFormulas, LookAt:=xlPart, _
        SearchOrder:=xlByRows, SearchDirection:=xlPrevious)

    If sel Is Nothing Then
        BarisTerakhirReset = 1
    Else
        BarisTerakhirReset = sel.Row
    End If

End Function

'Tujuan: Mengembalikan kolom terakhir yang digunakan.
Private Function KolomTerakhirReset(ByVal ws As Worksheet) As Long

    Dim sel As Range

    Set sel = ws.Cells.Find( _
        What:="*", After:=ws.Cells(1, 1), _
        LookIn:=xlFormulas, LookAt:=xlPart, _
        SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)

    If sel Is Nothing Then
        KolomTerakhirReset = 1
    Else
        KolomTerakhirReset = sel.Column
    End If

End Function

'Tujuan: Mengubah nilai menjadi teks tanpa error tipe.
Private Function TeksReset(ByVal nilai As Variant) As String

    If IsError(nilai) Or IsNull(nilai) Or IsEmpty(nilai) Then
        TeksReset = vbNullString
    Else
        TeksReset = CStr(nilai)
    End If

End Function

'Tujuan: Mengembalikan kondisi aplikasi setelah reset.
Private Sub PulihkanAplikasiReset( _
    ByVal calcLama As XlCalculation, _
    ByVal eventsLama As Boolean, _
    ByVal screenLama As Boolean, _
    ByVal statusLama As Variant)

    Application.Calculation = calcLama
    Application.EnableEvents = eventsLama
    Application.ScreenUpdating = screenLama
    Application.StatusBar = statusLama

End Sub
