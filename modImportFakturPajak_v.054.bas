Attribute VB_Name = "modImportFakturPajak"
Option Explicit

' ================================================================
'  ENGINE IMPOR FAKTUR PAJAK v0.11
'  - Tahap 1: salin setiap record CSV utuh ke satu sel tanpa format
'  - Tahap 2: verifikasi data raw sudah tertulis di workbook
'  - Tahap 3: normalisasi ke sheet FakturPK dan FakturPM
'  - Gabungkan NAMA_BARANG_JASA per nomor faktur dengan pemisah ", "
'  - Bentuk KUNCI_FAKTUR = nomor faktur ternormalisasi | YYYYMM
'  - Pilih banyak file sekaligus: maksimal 12 masa / 48 file
'  - Validasi pasangan induk-detail untuk setiap arah dan masa
'  - Impor ulang dapat mengganti seluruh data pada masa yang sama
' ================================================================

Private Const SHEET_PENJUAL As String = "FakturPK"
Private Const SHEET_PEMBELI As String = "FakturPM"
Private Const TABLE_PENJUAL As String = "tblFakturPK"
Private Const TABLE_PEMBELI As String = "tblFakturPM"
Private Const RAW_PENJUAL As String = "Raw_FakturPK"
Private Const RAW_DETAIL_PENJUAL As String = "Raw_DetailPK"
Private Const RAW_PEMBELI As String = "Raw_FakturPM"
Private Const RAW_DETAIL_PEMBELI As String = "Raw_DetailPM"
Private Const SHEET_LOG As String = "Log Impor Faktur"
Private Const COL_DETAIL As String = "DETAIL_BARANG_JASA"
Private Const COL_KEY As String = "KUNCI_FAKTUR"
Private Const COL_RAW_RECORD As String = "RAW_RECORD"
Private Const COL_RAW_MASA As String = "__MASA_FILE"
Private Const COL_RAW_FILE As String = "__SUMBER_FILE"
Private Const COL_RAW_TYPE As String = "__TIPE_RECORD"
Private Const TEMP_ROW_MARKER As String = "__TEMP_IMPORT_FAKTUR__"
Private mTahap As String

Public Sub BukaFormImporFaktur()
    frmImportFaktur.Show
End Sub

Public Sub PilihFileFaktur(ByVal frm As Object)
    Dim fd As Object
    Dim item As Variant
    Dim nama As String
    Dim jenis As String
    Dim masa As String
    Dim key As String
    Dim fileMap As Object

    Set fileMap = CreateObject("Scripting.Dictionary")
    fileMap.CompareMode = vbTextCompare

    Set fd = Application.FileDialog(3) ' msoFileDialogFilePicker
    With fd
        .Title = "Pilih file CSV faktur Pajak (maksimal 12 masa)"
        .AllowMultiSelect = True
        .Filters.Clear
        .Filters.Add "File CSV", "*.csv"

        If .Show <> -1 Then Exit Sub

        frm.lstFile.Clear

        For Each item In .SelectedItems
            nama = LCase$(NamaFile(CStr(item)))
            jenis = KlasifikasiFile(nama)
            masa = EkstrakMasaDariNamaFile(nama)

            If Len(jenis) = 0 Or Len(masa) = 0 Then
                TulisLog "PERINGATAN", "PILIH FILE", _
                          "Nama atau masa file tidak dikenali dan tidak digunakan.", CStr(item)
            Else
                key = jenis & "|" & masa
                If fileMap.Exists(key) Then
                    TulisLog "PERINGATAN", "PILIH FILE", _
                              "Ditemukan lebih dari satu file untuk jenis dan masa yang sama.", CStr(item)
                Else
                    fileMap.Add key, CStr(item)
                    TambahBarisDaftar frm, jenis, masa, CStr(item)
                End If
            End If
        Next item
    End With

    PerbaruiStatusForm frm
End Sub

Public Sub BersihkanDaftarFile(ByVal frm As Object)
    frm.lstFile.Clear
    PerbaruiStatusForm frm
End Sub

Public Sub JalankanImporDariForm(ByVal frm As Object)
    Dim pesan As String

    If Not ValidasiDaftarFile(frm, pesan) Then
        MsgBox pesan, vbExclamation, "Impor Faktur Pajak"
        Exit Sub
    End If

    frm.cmdImpor.Enabled = False
    frm.cmdPilihCSV.Enabled = False
    frm.cmdBersihkanDaftar.Enabled = False
    frm.cmdBatal.Enabled = False
    frm.MousePointer = 11 ' fmMousePointerHourGlass
    DoEvents

    On Error GoTo Gagal

    ImportDaftarFileFaktur BangunPetaFileDariForm(frm), CBool(frm.chkGantiData.Value)

    frm.MousePointer = 0
    frm.cmdImpor.Enabled = True
    frm.cmdPilihCSV.Enabled = True
    frm.cmdBersihkanDaftar.Enabled = True
    frm.cmdBatal.Enabled = True
    Exit Sub

Gagal:
    frm.MousePointer = 0
    frm.cmdImpor.Enabled = True
    frm.cmdPilihCSV.Enabled = True
    frm.cmdBersihkanDaftar.Enabled = True
    frm.cmdBatal.Enabled = True
    MsgBox "Impor gagal. Periksa sheet '" & SHEET_LOG & "'.", _
           vbExclamation, "Impor Faktur Pajak"
End Sub

Public Sub ImportDaftarFileFaktur( _
    ByVal fileMap As Object, _
    Optional ByVal gantiMasaYangSama As Boolean = True)

    Dim oldScreen As Boolean
    Dim oldEvents As Boolean
    Dim oldAlerts As Boolean
    Dim oldCalc As XlCalculation
    Dim jumlahPK As Long
    Dim jumlahPM As Long
    Dim rincianPK As Long
    Dim rincianPM As Long
    Dim jumlahPasangan As Long
    Dim jmlInduk As Long
    Dim jmlDetail As Long
    Dim key As Variant
    Dim bagian As Variant
    Dim jenis As String
    Dim masa As String
    Dim keyDetail As String
    Dim arah As String
    Dim adaPenjual As Boolean
    Dim adaPembeli As Boolean
    Dim errNum As Long
    Dim errDesc As String

    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    oldAlerts = Application.DisplayAlerts
    oldCalc = Application.Calculation

    On Error GoTo Gagal

    mTahap = "Menyiapkan Excel untuk proses impor"
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    For Each key In fileMap.Keys
        bagian = Split(CStr(key), "|")
        jenis = CStr(bagian(0))
        masa = CStr(bagian(1))

        If Right$(jenis, 6) = "_INDUK" Then
            keyDetail = Replace(jenis, "_INDUK", "_DETAIL") & "|" & masa
            arah = Replace(jenis, "_INDUK", vbNullString)
            mTahap = "Mengimpor data mentah " & arah & " masa " & masa
            Application.StatusBar = "Mengimpor data mentah " & LCase$(arah) & _
                                    " masa " & masa & "..."

            jmlInduk = 0
            jmlDetail = 0

            If arah = "PENJUAL" Then
                ImporPasanganKeRaw CStr(fileMap(key)), CStr(fileMap(keyDetail)), _
                                   RAW_PENJUAL, RAW_DETAIL_PENJUAL, masa, _
                                   gantiMasaYangSama, jmlInduk, jmlDetail
                adaPenjual = True
            Else
                ImporPasanganKeRaw CStr(fileMap(key)), CStr(fileMap(keyDetail)), _
                                   RAW_PEMBELI, RAW_DETAIL_PEMBELI, masa, _
                                   gantiMasaYangSama, jmlInduk, jmlDetail
                adaPembeli = True
            End If

            TulisLog "RAW TERSIMPAN", arah & " " & masa, _
                      Format$(jmlInduk, "#,##0") & " baris induk dan " & _
                      Format$(jmlDetail, "#,##0") & _
                      " baris detail sudah diverifikasi di sheet raw.", CStr(fileMap(key))

            jumlahPasangan = jumlahPasangan + 1
        End If
    Next key

    If adaPenjual Then
        Application.StatusBar = "Menormalisasi faktur penjual dari sheet raw..."
        NormalisasiDariRaw RAW_PENJUAL, RAW_DETAIL_PENJUAL, _
                           SHEET_PENJUAL, TABLE_PENJUAL, jumlahPK, rincianPK
    End If

    If adaPembeli Then
        Application.StatusBar = "Menormalisasi faktur pembeli dari sheet raw..."
        NormalisasiDariRaw RAW_PEMBELI, RAW_DETAIL_PEMBELI, _
                           SHEET_PEMBELI, TABLE_PEMBELI, jumlahPM, rincianPM
    End If

    SembunyikanSheetRaw

    TulisLog "SELESAI", "RINGKASAN", _
              "Faktur penjual: " & Format$(jumlahPK, "#,##0") & _
              " baris; detail: " & Format$(rincianPK, "#,##0") & _
              " baris. Faktur pembeli: " & Format$(jumlahPM, "#,##0") & _
              " baris; detail: " & Format$(rincianPM, "#,##0") & " baris.", ""

    PulihkanExcel oldScreen, oldEvents, oldAlerts, oldCalc

    MsgBox "Impor faktur Pajak selesai." & vbCrLf & vbCrLf & _
           "Pasangan file: " & jumlahPasangan & vbCrLf & _
           "Penjual: " & Format$(jumlahPK, "#,##0") & " baris" & vbCrLf & _
           "Pembeli: " & Format$(jumlahPM, "#,##0") & " baris" & vbCrLf & vbCrLf & _
           "Rincian proses tersedia pada sheet '" & SHEET_LOG & "'.", _
           vbInformation, "Impor Faktur Pajak"
    Exit Sub

Gagal:
    errNum = Err.Number
    errDesc = Err.Description
    TulisLog "ERROR", "IMPORT", _
              "Tahap: " & mTahap & " | " & errNum & " - " & errDesc, ""
    TampilkanSheetRawYangAda
    PulihkanExcel oldScreen, oldEvents, oldAlerts, oldCalc
    Err.Raise errNum, "ImportDaftarFileFaktur", errDesc
End Sub

' Kompatibilitas untuk pemanggilan versi awal empat file.
Public Sub ImportEmpatFileFaktur( _
    ByVal pathFakturPenjual As String, _
    ByVal pathDetailPenjual As String, _
    ByVal pathFakturPembeli As String, _
    ByVal pathDetailPembeli As String, _
    Optional ByVal gantiMasaYangSama As Boolean = True)

    Dim fileMap As Object
    Dim p As Variant
    Dim jenis As String
    Dim masa As String

    Set fileMap = CreateObject("Scripting.Dictionary")
    fileMap.CompareMode = vbTextCompare

    For Each p In Array(pathFakturPenjual, pathDetailPenjual, _
                        pathFakturPembeli, pathDetailPembeli)
        jenis = KlasifikasiFile(LCase$(NamaFile(CStr(p))))
        masa = EkstrakMasaDariNamaFile(LCase$(NamaFile(CStr(p))))
        If Len(jenis) = 0 Or Len(masa) = 0 Then
            Err.Raise vbObjectError + 107, , "Nama file tidak dapat diklasifikasikan: " & NamaFile(CStr(p))
        End If
        fileMap(jenis & "|" & masa) = CStr(p)
    Next p

    ImportDaftarFileFaktur fileMap, gantiMasaYangSama
End Sub

Private Sub ImporPasanganKeRaw( _
    ByVal pathInduk As String, _
    ByVal pathDetail As String, _
    ByVal sheetInduk As String, _
    ByVal sheetDetail As String, _
    ByVal masaFile As String, _
    ByVal gantiMasaYangSama As Boolean, _
    ByRef jumlahInduk As Long, _
    ByRef jumlahDetail As Long)

    mTahap = "Menulis file mentah " & NamaFile(pathInduk) & " ke " & sheetInduk
    ImporCSVKeSheetRaw pathInduk, sheetInduk, masaFile, _
                        gantiMasaYangSama, jumlahInduk

    mTahap = "Menulis file mentah " & NamaFile(pathDetail) & " ke " & sheetDetail
    ImporCSVKeSheetRaw pathDetail, sheetDetail, masaFile, _
                        gantiMasaYangSama, jumlahDetail
End Sub

Private Sub ImporCSVKeSheetRaw( _
    ByVal filePath As String, _
    ByVal namaSheet As String, _
    ByVal masaFile As String, _
    ByVal gantiMasaYangSama As Boolean, _
    ByRef jumlahBaris As Long)

    Dim ws As Worksheet
    Dim recordsCSV As Collection
    Dim rows As Collection
    Dim rowData As Variant
    Dim data() As Variant
    Dim recordText As String
    Dim mulaiBaris As Long
    Dim r As Long
    Dim jumlahMasaSetelahTulis As Long

    Set ws = AmbilAtauBuatSheet(namaSheet)
    PastikanSheetDapatDitulis ws
    Set rows = New Collection

    If WorksheetKosong(ws) Or Not StrukturPenampungRawValid(ws) Then
        mTahap = "Menyiapkan penampung raw tanpa format pada " & namaSheet
        If Not WorksheetKosong(ws) Then
            ws.Range(ws.Cells(1, 1), _
                     ws.Cells(BarisTerakhirIsi(ws), KolomTerakhirIsi(ws))).ClearContents
        End If
        ws.Cells(1, 1).Value2 = COL_RAW_RECORD
        ws.Cells(1, 2).Value2 = COL_RAW_MASA
        ws.Cells(1, 3).Value2 = COL_RAW_FILE
        ws.Cells(1, 4).Value2 = COL_RAW_TYPE
    End If

    If gantiMasaYangSama Then
        mTahap = "Menghapus data raw masa " & masaFile & " dari " & namaSheet
        HapusMasaDariSheetRaw ws, masaFile
    End If

    Set recordsCSV = BacaSemuaRecordCSV(filePath)
    If recordsCSV.Count = 0 Then
        Err.Raise vbObjectError + 201, , "File CSV kosong: " & filePath
    End If
    rows.Add Array(CStr(recordsCSV(1)), masaFile, NamaFile(filePath), "HEADER")

    For r = 2 To recordsCSV.Count
        recordText = CStr(recordsCSV(r))
        If Len(recordText) > 0 Then
            rows.Add Array(recordText, masaFile, NamaFile(filePath), "DATA")
        End If
    Next r

    jumlahBaris = rows.Count - 1
    If jumlahBaris = 0 Then
        Err.Raise vbObjectError + 202, , _
                  "Tidak ada baris data pada " & NamaFile(filePath)
    End If

    ReDim data(1 To rows.Count, 1 To 4)
    For r = 1 To rows.Count
        rowData = rows(r)
        data(r, 1) = CStr(rowData(0))
        data(r, 2) = CStr(rowData(1))
        data(r, 3) = CStr(rowData(2))
        data(r, 4) = CStr(rowData(3))
    Next r

    mulaiBaris = BarisTerakhirIsi(ws) + 1
    mTahap = "Menyalin record mentah apa adanya ke " & namaSheet
    ws.Cells(mulaiBaris, 1).Resize(rows.Count, 4).Value2 = data

    jumlahMasaSetelahTulis = HitungBarisMasaRaw(ws, masaFile)
    If gantiMasaYangSama And jumlahMasaSetelahTulis <> jumlahBaris Then
        Err.Raise vbObjectError + 203, , _
                  "Verifikasi sheet raw gagal pada " & namaSheet & _
                  ". Seharusnya " & jumlahBaris & " baris, ditemukan " & _
                  jumlahMasaSetelahTulis & " baris."
    ElseIf jumlahMasaSetelahTulis < jumlahBaris Then
        Err.Raise vbObjectError + 204, , _
                  "Sebagian data mentah tidak tertulis pada " & namaSheet & "."
    End If
End Sub

Private Function BacaSemuaRecordCSV(ByVal filePath As String) As Collection
    Dim stm As Object
    Dim hasil As Collection
    Dim seluruhTeks As String
    Dim recordText As String
    Dim ch As String
    Dim dalamKutip As Boolean
    Dim i As Long

    Set hasil = New Collection
    Set stm = BukaTextStream(filePath)
    seluruhTeks = stm.ReadText(-1) ' adReadAll
    stm.Close

    i = 1
    Do While i <= Len(seluruhTeks)
        ch = Mid$(seluruhTeks, i, 1)

        If ch = """" Then
            recordText = recordText & ch
            If dalamKutip And i < Len(seluruhTeks) And _
               Mid$(seluruhTeks, i + 1, 1) = """" Then
                recordText = recordText & """"
                i = i + 1
            Else
                dalamKutip = Not dalamKutip
            End If
        ElseIf (ch = vbCr Or ch = vbLf) And Not dalamKutip Then
            If Len(recordText) > 0 Then hasil.Add recordText
            recordText = vbNullString
            If ch = vbCr And i < Len(seluruhTeks) Then
                If Mid$(seluruhTeks, i + 1, 1) = vbLf Then i = i + 1
            End If
        ElseIf (ch = vbCr Or ch = vbLf) And dalamKutip Then
            recordText = recordText & vbLf
            If ch = vbCr And i < Len(seluruhTeks) Then
                If Mid$(seluruhTeks, i + 1, 1) = vbLf Then i = i + 1
            End If
        Else
            recordText = recordText & ch
        End If

        i = i + 1
    Loop

    If dalamKutip Then
        Err.Raise vbObjectError + 215, , _
                  "Tanda kutip CSV tidak seimbang pada " & NamaFile(filePath)
    End If
    If Len(recordText) > 0 Then hasil.Add recordText

    Set BacaSemuaRecordCSV = hasil
End Function

Private Function StrukturPenampungRawValid(ByVal ws As Worksheet) As Boolean
    StrukturPenampungRawValid = _
        (CStr(ws.Cells(1, 1).Value2) = COL_RAW_RECORD) And _
        (CStr(ws.Cells(1, 2).Value2) = COL_RAW_MASA) And _
        (CStr(ws.Cells(1, 3).Value2) = COL_RAW_FILE) And _
        (CStr(ws.Cells(1, 4).Value2) = COL_RAW_TYPE)
End Function

Private Sub HapusMasaDariSheetRaw(ByVal ws As Worksheet, ByVal masaFile As String)
    Dim sumber As Variant
    Dim hasil() As Variant
    Dim lastRow As Long
    Dim jumlahSimpan As Long
    Dim r As Long
    Dim c As Long

    lastRow = BarisTerakhirIsi(ws)
    If lastRow < 2 Then Exit Sub
    sumber = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, 4)).Value2

    For r = 1 To UBound(sumber, 1)
        If StrComp(CStr(sumber(r, 2)), masaFile, vbTextCompare) <> 0 Then
            jumlahSimpan = jumlahSimpan + 1
        End If
    Next r

    ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, 4)).ClearContents
    If jumlahSimpan = 0 Then Exit Sub

    ReDim hasil(1 To jumlahSimpan, 1 To 4)
    jumlahSimpan = 0
    For r = 1 To UBound(sumber, 1)
        If StrComp(CStr(sumber(r, 2)), masaFile, vbTextCompare) <> 0 Then
            jumlahSimpan = jumlahSimpan + 1
            For c = 1 To 4
                hasil(jumlahSimpan, c) = sumber(r, c)
            Next c
        End If
    Next r
    ws.Cells(2, 1).Resize(jumlahSimpan, 4).Value2 = hasil
End Sub

Private Function HitungBarisMasaRaw( _
    ByVal ws As Worksheet, _
    ByVal masaFile As String) As Long

    Dim data As Variant
    Dim lastRow As Long
    Dim r As Long

    lastRow = BarisTerakhirIsi(ws)
    If lastRow < 2 Then Exit Function
    data = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, 4)).Value2

    For r = 1 To UBound(data, 1)
        If StrComp(CStr(data(r, 2)), masaFile, vbTextCompare) = 0 And _
           StrComp(CStr(data(r, 4)), "DATA", vbTextCompare) = 0 Then
            HitungBarisMasaRaw = HitungBarisMasaRaw + 1
        End If
    Next r
End Function

Private Sub NormalisasiDariRaw( _
    ByVal sheetIndukRaw As String, _
    ByVal sheetDetailRaw As String, _
    ByVal sheetHasil As String, _
    ByVal namaTabel As String, _
    ByRef jumlahInduk As Long, _
    ByRef jumlahDetail As Long)

    Dim mapDetail As Object
    Dim headers As Variant
    Dim data As Variant

    mTahap = "Membaca detail dari sheet raw " & sheetDetailRaw
    Set mapDetail = BangunKamusDetailDariRaw(sheetDetailRaw, jumlahDetail)
    mTahap = "Menormalisasi induk dari sheet raw " & sheetIndukRaw
    data = BacaIndukDariRaw(sheetIndukRaw, mapDetail, headers, jumlahInduk)
    mTahap = "Menulis hasil normalisasi ke " & sheetHasil
    TulisHasilNormalisasi sheetHasil, namaTabel, headers, data, jumlahInduk

    TulisLog "BERHASIL", "NORMALISASI " & sheetHasil, _
              Format$(jumlahInduk, "#,##0") & " baris faktur dan " & _
              Format$(jumlahDetail, "#,##0") & " baris detail diproses dari sheet raw.", ""
End Sub

Private Function BangunKamusDetailDariRaw( _
    ByVal namaSheet As String, _
    ByRef jumlahBaris As Long) As Object

    Dim ws As Worksheet
    Dim hasil As Object
    Dim idx As Object
    Dim data As Variant
    Dim header As Variant
    Dim fields As Variant
    Dim lastRow As Long
    Dim noFaktur As String
    Dim namaBarang As String
    Dim masaFile As String
    Dim key As String
    Dim r As Long

    Set ws = ThisWorkbook.Worksheets(namaSheet)
    lastRow = BarisTerakhirIsi(ws)
    If lastRow < 2 Then
        Err.Raise vbObjectError + 208, , "Sheet raw '" & namaSheet & "' tidak berisi data."
    End If

    data = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, 4)).Value2
    header = HeaderPertamaDariPenampungRaw(data, namaSheet)
    Set idx = BuatIndeksHeader(header)
    WajibAdaKolom idx, "NO_FAKTUR", namaSheet
    WajibAdaKolom idx, "NAMA_BARANG_JASA", namaSheet

    Set hasil = CreateObject("Scripting.Dictionary")
    hasil.CompareMode = vbTextCompare

    For r = 1 To UBound(data, 1)
        If StrComp(CStr(data(r, 4)), "DATA", vbTextCompare) = 0 Then
            fields = ParsePipeLine(CStr(data(r, 1)))
            If BarisMemilikiIsi(fields) Then
                jumlahBaris = jumlahBaris + 1
                noFaktur = AmbilField(fields, CLng(idx("NO_FAKTUR")))
                namaBarang = BersihkanNamaBarang( _
                                AmbilField(fields, CLng(idx("NAMA_BARANG_JASA"))))
                masaFile = CStr(data(r, 2))
                key = NormalisasiNomorFaktur(noFaktur) & "|" & masaFile

                If Len(NormalisasiNomorFaktur(noFaktur)) > 0 And Len(namaBarang) > 0 Then
                    If hasil.Exists(key) Then
                        hasil(key) = GabungkanDetail(CStr(hasil(key)), namaBarang)
                    Else
                        hasil.Add key, namaBarang
                    End If
                End If
            End If
        End If
    Next r

    Set BangunKamusDetailDariRaw = hasil
End Function

Private Function BacaIndukDariRaw( _
    ByVal namaSheet As String, _
    ByVal mapDetail As Object, _
    ByRef outputHeaders As Variant, _
    ByRef jumlahBaris As Long) As Variant

    Dim ws As Worksheet
    Dim idx As Object
    Dim rawData As Variant
    Dim header As Variant
    Dim fields As Variant
    Dim outputData() As Variant
    Dim lastRow As Long
    Dim sourceCols As Long
    Dim outCols As Long
    Dim namaKolom As String
    Dim noFaktur As String
    Dim masa As String
    Dim tahun As String
    Dim masaFile As String
    Dim masaNorm As String
    Dim keyDetail As String
    Dim penandaDokumen As String
    Dim i As Long
    Dim r As Long
    Dim outRow As Long

    Set ws = ThisWorkbook.Worksheets(namaSheet)
    lastRow = BarisTerakhirIsi(ws)
    If lastRow < 2 Then
        Err.Raise vbObjectError + 209, , "Sheet raw '" & namaSheet & "' tidak berisi data."
    End If

    rawData = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, 4)).Value2
    header = HeaderPertamaDariPenampungRaw(rawData, namaSheet)
    Set idx = BuatIndeksHeader(header)
    WajibAdaKolom idx, "TAX_INVOICE_NUMBER", namaSheet
    WajibAdaKolom idx, "TAX_INVOICE_PERIOD", namaSheet
    WajibAdaKolom idx, "TAX_INVOICE_YEAR", namaSheet

    sourceCols = UBound(header) - LBound(header) + 1
    outCols = sourceCols + 2
    ReDim outputHeaders(0 To outCols - 1)
    For i = 1 To sourceCols
        outputHeaders(i - 1) = CStr(header(LBound(header) + i - 1))
        If i = 1 Then outputHeaders(i - 1) = HapusBOM(CStr(outputHeaders(i - 1)))
    Next i
    outputHeaders(sourceCols) = COL_DETAIL
    outputHeaders(sourceCols + 1) = COL_KEY

    For r = 1 To UBound(rawData, 1)
        If StrComp(CStr(rawData(r, 4)), "DATA", vbTextCompare) = 0 Then
            fields = ParsePipeLine(CStr(rawData(r, 1)))
            If BarisMemilikiIsi(fields) Then jumlahBaris = jumlahBaris + 1
        End If
    Next r
    If jumlahBaris = 0 Then
        Err.Raise vbObjectError + 214, , _
                  "Tidak ada record data pada sheet raw '" & namaSheet & "'."
    End If
    ReDim outputData(1 To jumlahBaris, 1 To outCols)

    For r = 1 To UBound(rawData, 1)
        If StrComp(CStr(rawData(r, 4)), "DATA", vbTextCompare) = 0 Then
            fields = ParsePipeLine(CStr(rawData(r, 1)))
            If BarisMemilikiIsi(fields) Then
                outRow = outRow + 1

                penandaDokumen = AmbilFieldOpsional(fields, idx, "TABEL")

                For i = 1 To sourceCols
                    namaKolom = CStr(outputHeaders(i - 1))
                    outputData(outRow, i) = KonversiNilai( _
                                                namaKolom, _
                                                AmbilField(fields, i - 1), _
                                                penandaDokumen)
                Next i

                noFaktur = AmbilField(fields, CLng(idx("TAX_INVOICE_NUMBER")))
                masa = AmbilField(fields, CLng(idx("TAX_INVOICE_PERIOD")))
                tahun = AmbilField(fields, CLng(idx("TAX_INVOICE_YEAR")))
                masaFile = CStr(rawData(r, 2))
                masaNorm = NormalisasiMasa(tahun, masa)
                If Len(masaNorm) = 0 Then masaNorm = masaFile
                keyDetail = NormalisasiNomorFaktur(noFaktur) & "|" & masaNorm

                If mapDetail.Exists(keyDetail) Then
                    outputData(outRow, sourceCols + 1) = CStr(mapDetail(keyDetail))
                Else
                    outputData(outRow, sourceCols + 1) = vbNullString
                End If
                If Len(NormalisasiNomorFaktur(noFaktur)) > 0 And Len(masaNorm) > 0 Then
                    outputData(outRow, sourceCols + 2) = _
                        NormalisasiNomorFaktur(noFaktur) & "|" & masaNorm
                Else
                    outputData(outRow, sourceCols + 2) = vbNullString
                End If
            End If
        End If
    Next r

    BacaIndukDariRaw = outputData
End Function

Private Function HeaderPertamaDariPenampungRaw( _
    ByVal data As Variant, _
    ByVal namaSheet As String) As Variant

    Dim r As Long

    For r = 1 To UBound(data, 1)
        If StrComp(CStr(data(r, 4)), "HEADER", vbTextCompare) = 0 Then
            HeaderPertamaDariPenampungRaw = ParsePipeLine(CStr(data(r, 1)))
            Exit Function
        End If
    Next r

    Err.Raise vbObjectError + 213, , _
              "Record header tidak ditemukan pada sheet raw '" & namaSheet & "'."
End Function

Private Sub TulisHasilNormalisasi( _
    ByVal namaSheet As String, _
    ByVal namaTabel As String, _
    ByVal headers As Variant, _
    ByVal data As Variant, _
    ByVal jumlahBaris As Long)

    Dim ws As Worksheet
    Dim lo As ListObject
    Dim loKonflik As ListObject
    Dim rngHasil As Range
    Dim jumlahKolom As Long
    Dim barisLama As Long
    Dim kolomLama As Long
    Dim i As Long

    Set ws = AmbilAtauBuatSheet(namaSheet)
    PastikanSheetDapatDitulis ws

    mTahap = "Membersihkan hasil lama pada " & namaSheet
    For i = ws.ListObjects.Count To 1 Step -1
        ws.ListObjects(i).Unlist
    Next i
    barisLama = BarisTerakhirIsi(ws)
    kolomLama = KolomTerakhirIsi(ws)
    If barisLama > 0 And kolomLama > 0 Then
        ws.Range(ws.Cells(1, 1), ws.Cells(barisLama, kolomLama)).Clear
    End If

    jumlahKolom = UBound(headers) - LBound(headers) + 1
    mTahap = "Menulis header hasil ke " & namaSheet
    For i = 1 To jumlahKolom
        ws.Cells(1, i).Value2 = CStr(headers(LBound(headers) + i - 1))
    Next i
    mTahap = "Memformat header hasil pada " & namaSheet
    ws.Cells(1, 1).Resize(1, jumlahKolom).Font.Bold = True

    If jumlahBaris > 0 Then
        mTahap = "Menetapkan tipe kolom sebelum menulis " & namaSheet
        TerapkanFormatSebelumTulis ws, headers, jumlahBaris
        mTahap = "Menulis data hasil ke " & namaSheet
        ws.Cells(2, 1).Resize(jumlahBaris, jumlahKolom).Value2 = data
    End If

    If BarisTerakhirIsi(ws) <> jumlahBaris + 1 Then
        Err.Raise vbObjectError + 210, , _
                  "Verifikasi hasil normalisasi gagal pada sheet '" & namaSheet & "'."
    End If

    Set loKonflik = CariTabelBernama(namaTabel)
    If Not loKonflik Is Nothing Then
        Err.Raise vbObjectError + 211, , _
                  "Nama tabel '" & namaTabel & "' masih digunakan pada sheet '" & _
                  loKonflik.Parent.Name & "'."
    End If

    mTahap = "Membentuk tabel hasil " & namaTabel
    Set rngHasil = ws.Range(ws.Cells(1, 1), _
                            ws.Cells(jumlahBaris + 1, jumlahKolom))
    Set lo = ws.ListObjects.Add(xlSrcRange, rngHasil, , xlYes)
    lo.Name = namaTabel
    lo.TableStyle = "TableStyleMedium2"

    mTahap = "Memformat tabel hasil " & namaTabel
    TerapkanFormatTabel lo
    lo.Range.Columns.AutoFit
    BatasiLebarKolom lo, "SELLER_TAXPAYER_NAME", 28
    BatasiLebarKolom lo, "BUYER_TAXPAYER_NAME", 28
    BatasiLebarKolom lo, COL_DETAIL, 60
    If Not lo.DataBodyRange Is Nothing Then
        lo.ListColumns(COL_DETAIL).DataBodyRange.WrapText = True
    End If
End Sub

Private Sub PastikanSheetDapatDitulis(ByVal ws As Worksheet)
    If ws.ProtectContents Then
        On Error Resume Next
        ws.Unprotect
        On Error GoTo 0
    End If

    If ws.ProtectContents Then
        Err.Raise vbObjectError + 212, , _
                  "Sheet '" & ws.Name & "' diproteksi. Buka proteksi sheet sebelum impor."
    End If
End Sub

Private Function BuatIndeksHeaderDariSheet( _
    ByVal ws As Worksheet, _
    ByVal jumlahKolom As Long) As Object

    Dim hasil As Object
    Dim i As Long
    Dim nama As String

    Set hasil = CreateObject("Scripting.Dictionary")
    hasil.CompareMode = vbTextCompare
    For i = 1 To jumlahKolom
        nama = UCase$(Trim$(CStr(ws.Cells(1, i).Value2)))
        If i = 1 Then nama = HapusBOM(nama)
        If Len(nama) > 0 Then hasil(nama) = i - 1
    Next i
    Set BuatIndeksHeaderDariSheet = hasil
End Function

Private Function BarisTerakhirIsi(ByVal ws As Worksheet) As Long
    Dim selTerakhir As Range

    On Error Resume Next
    Set selTerakhir = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), _
        LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious, MatchCase:=False, SearchFormat:=False)
    On Error GoTo 0
    If Not selTerakhir Is Nothing Then BarisTerakhirIsi = selTerakhir.Row
End Function

Private Function KolomTerakhirIsi(ByVal ws As Worksheet) As Long
    Dim selTerakhir As Range

    On Error Resume Next
    Set selTerakhir = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), _
        LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, _
        SearchDirection:=xlPrevious, MatchCase:=False, SearchFormat:=False)
    On Error GoTo 0
    If Not selTerakhir Is Nothing Then KolomTerakhirIsi = selTerakhir.Column
End Function

Private Sub SembunyikanSheetRaw()
    Dim nama As Variant
    Dim ws As Worksheet
    Dim wsTujuan As Worksheet

    On Error Resume Next
    Set wsTujuan = ThisWorkbook.Worksheets(SHEET_PENJUAL)
    If wsTujuan Is Nothing Then Set wsTujuan = ThisWorkbook.Worksheets(SHEET_PEMBELI)
    If wsTujuan Is Nothing Then Set wsTujuan = ThisWorkbook.Worksheets(SHEET_LOG)
    If Not wsTujuan Is Nothing Then
        ThisWorkbook.Activate
        wsTujuan.Activate
    End If
    On Error GoTo 0

    For Each nama In Array(RAW_PENJUAL, RAW_DETAIL_PENJUAL, _
                           RAW_PEMBELI, RAW_DETAIL_PEMBELI)
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(CStr(nama))
        On Error GoTo 0
        If Not ws Is Nothing Then ws.Visible = xlSheetHidden
    Next nama
End Sub

Private Sub TampilkanSheetRawYangAda()
    Dim nama As Variant
    Dim ws As Worksheet

    For Each nama In Array(RAW_PENJUAL, RAW_DETAIL_PENJUAL, _
                           RAW_PEMBELI, RAW_DETAIL_PEMBELI)
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(CStr(nama))
        If Not ws Is Nothing Then ws.Visible = xlSheetVisible
        On Error GoTo 0
    Next nama
End Sub

Private Sub ImportSatuArah( _
    ByVal pathInduk As String, _
    ByVal pathDetail As String, _
    ByVal namaSheet As String, _
    ByVal namaTabel As String, _
    ByVal arah As String, _
    ByVal gantiMasaYangSama As Boolean, _
    ByRef jumlahInduk As Long, _
    ByRef jumlahDetail As Long)

    Dim mapDetail As Object
    Dim periods As Object
    Dim headers As Variant
    Dim data As Variant

    mTahap = "Membaca file detail " & NamaFile(pathDetail)
    Set mapDetail = BangunKamusDetail(pathDetail, jumlahDetail)
    mTahap = "Membaca file induk " & NamaFile(pathInduk)
    data = BacaFakturInduk(pathInduk, mapDetail, headers, periods, jumlahInduk)

    mTahap = "Menulis hasil ke sheet " & namaSheet
    TulisHasilKeSheet namaSheet, namaTabel, headers, data, jumlahInduk, _
                       periods, gantiMasaYangSama

    TulisLog "BERHASIL", arah, _
              Format$(jumlahInduk, "#,##0") & " baris faktur diimpor; " & _
              Format$(jumlahDetail, "#,##0") & " baris detail dibaca.", pathInduk
End Sub

Private Function BangunKamusDetail( _
    ByVal filePath As String, _
    ByRef jumlahBaris As Long) As Object

    Dim hasil As Object
    Dim idx As Object
    Dim stm As Object
    Dim header As Variant
    Dim fields As Variant
    Dim recordText As String
    Dim noFaktur As String
    Dim namaBarang As String
    Dim key As String

    Set hasil = CreateObject("Scripting.Dictionary")
    hasil.CompareMode = vbTextCompare

    Set stm = BukaTextStream(filePath)
    If Not BacaRecordBerikutnya(stm, recordText) Then
        Err.Raise vbObjectError + 101, , "File detail kosong: " & filePath
    End If

    header = ParsePipeLine(recordText)
    Set idx = BuatIndeksHeader(header)
    WajibAdaKolom idx, "NO_FAKTUR", filePath
    WajibAdaKolom idx, "NAMA_BARANG_JASA", filePath

    Do While BacaRecordBerikutnya(stm, recordText)
        If Len(recordText) > 0 Then
            fields = ParsePipeLine(recordText)
            If BarisMemilikiIsi(fields) Then
                jumlahBaris = jumlahBaris + 1
                noFaktur = AmbilField(fields, CLng(idx("NO_FAKTUR")))
                namaBarang = BersihkanNamaBarang( _
                                AmbilField(fields, CLng(idx("NAMA_BARANG_JASA"))))
                key = NormalisasiNomorFaktur(noFaktur)

                If Len(key) > 0 And Len(namaBarang) > 0 Then
                    If hasil.Exists(key) Then
                        hasil(key) = GabungkanDetail(CStr(hasil(key)), namaBarang)
                    Else
                        hasil.Add key, namaBarang
                    End If
                End If
            End If
        End If
    Loop

    stm.Close
    Set BangunKamusDetail = hasil
End Function

Private Function BacaFakturInduk( _
    ByVal filePath As String, _
    ByVal mapDetail As Object, _
    ByRef outputHeaders As Variant, _
    ByRef periods As Object, _
    ByRef jumlahBaris As Long) As Variant

    Dim stm As Object
    Dim idx As Object
    Dim rows As Collection
    Dim header As Variant
    Dim fields As Variant
    Dim rowData As Variant
    Dim outputData() As Variant
    Dim recordText As String
    Dim sourceCols As Long
    Dim outCols As Long
    Dim noFaktur As String
    Dim keyNomor As String
    Dim masa As String
    Dim tahun As String
    Dim kunci As String
    Dim penandaDokumen As String
    Dim i As Long
    Dim r As Long

    Set periods = CreateObject("Scripting.Dictionary")
    periods.CompareMode = vbTextCompare
    Set rows = New Collection

    Set stm = BukaTextStream(filePath)
    If Not BacaRecordBerikutnya(stm, recordText) Then
        Err.Raise vbObjectError + 102, , "File induk kosong: " & filePath
    End If

    header = ParsePipeLine(recordText)
    Set idx = BuatIndeksHeader(header)
    WajibAdaKolom idx, "TAX_INVOICE_NUMBER", filePath
    WajibAdaKolom idx, "TAX_INVOICE_PERIOD", filePath
    WajibAdaKolom idx, "TAX_INVOICE_YEAR", filePath

    sourceCols = UBound(header) - LBound(header) + 1
    outCols = sourceCols + 2
    ReDim outputHeaders(0 To outCols - 1)

    For i = 0 To sourceCols - 1
        outputHeaders(i) = CStr(header(LBound(header) + i))
    Next i
    outputHeaders(sourceCols) = COL_DETAIL
    outputHeaders(sourceCols + 1) = COL_KEY

    Do While BacaRecordBerikutnya(stm, recordText)
        If Len(recordText) > 0 Then
            fields = ParsePipeLine(recordText)
            If BarisMemilikiIsi(fields) Then
                ReDim rowData(0 To outCols - 1)

                penandaDokumen = AmbilFieldOpsional(fields, idx, "TABEL")

                For i = 0 To sourceCols - 1
                    rowData(i) = KonversiNilai( _
                                    CStr(header(LBound(header) + i)), _
                                    AmbilField(fields, i), _
                                    penandaDokumen)
                Next i

                noFaktur = AmbilField(fields, CLng(idx("TAX_INVOICE_NUMBER")))
                masa = AmbilField(fields, CLng(idx("TAX_INVOICE_PERIOD")))
                tahun = AmbilField(fields, CLng(idx("TAX_INVOICE_YEAR")))
                keyNomor = NormalisasiNomorFaktur(noFaktur)

                If mapDetail.Exists(keyNomor) Then
                    rowData(sourceCols) = CStr(mapDetail(keyNomor))
                Else
                    rowData(sourceCols) = vbNullString
                End If

                kunci = BentukKunciFaktur(noFaktur, tahun, masa)
                rowData(sourceCols + 1) = kunci

                If Len(NormalisasiMasa(tahun, masa)) > 0 Then
                    periods(NormalisasiMasa(tahun, masa)) = True
                End If

                rows.Add rowData
            End If
        End If
    Loop

    stm.Close
    jumlahBaris = rows.Count

    If jumlahBaris = 0 Then
        BacaFakturInduk = Empty
        Exit Function
    End If

    ReDim outputData(1 To jumlahBaris, 1 To outCols)
    For r = 1 To jumlahBaris
        rowData = rows(r)
        For i = 1 To outCols
            outputData(r, i) = rowData(i - 1)
        Next i
    Next r

    BacaFakturInduk = outputData
End Function

Private Sub TulisHasilKeSheet( _
    ByVal namaSheet As String, _
    ByVal namaTabel As String, _
    ByVal headers As Variant, _
    ByVal data As Variant, _
    ByVal jumlahBaris As Long, _
    ByVal periods As Object, _
    ByVal gantiMasaYangSama As Boolean)

    Dim ws As Worksheet
    Dim lo As ListObject
    Dim jumlahKolom As Long
    Dim oldRows As Long
    Dim target As Range
    Dim tabelBaru As Boolean

    jumlahKolom = UBound(headers) - LBound(headers) + 1
    mTahap = "Membuka atau membuat sheet " & namaSheet
    Set ws = AmbilAtauBuatSheet(namaSheet)
    mTahap = "Membuka atau membuat tabel " & namaTabel
    Set lo = AmbilAtauBuatTabel(ws, namaTabel, headers, tabelBaru)
    mTahap = "Memeriksa struktur tabel " & namaTabel
    ValidasiHeaderTabel lo, headers

    If gantiMasaYangSama And periods.Count > 0 Then
        mTahap = "Menghapus data lama untuk masa yang sama pada " & namaTabel
        HapusMasaYangSama lo, periods
    End If

    If jumlahBaris = 0 Then
        If tabelBaru Then lo.DataBodyRange.ClearContents
        Exit Sub
    End If

    oldRows = lo.ListRows.Count
    mTahap = "Memperbesar tabel " & namaTabel
    lo.Resize ws.Range( _
        lo.HeaderRowRange.Cells(1, 1), _
        lo.HeaderRowRange.Cells(1, 1).Offset(oldRows + jumlahBaris, jumlahKolom - 1))

    Set target = lo.DataBodyRange.Rows(oldRows + 1).Resize(jumlahBaris, jumlahKolom)
    mTahap = "Memformat dan menulis data ke " & namaTabel
    TerapkanFormatTabel lo
    target.Value2 = data

    If tabelBaru Then
        mTahap = "Menghapus baris sementara pada " & namaTabel
        If lo.ListRows.Count > 1 Then
            lo.ListRows(1).Delete
        ElseIf Not lo.DataBodyRange Is Nothing Then
            lo.DataBodyRange.ClearContents
        End If
    End If

    lo.Range.Columns.AutoFit
    BatasiLebarKolom lo, "SELLER_TAXPAYER_NAME", 28
    BatasiLebarKolom lo, "BUYER_TAXPAYER_NAME", 28
    BatasiLebarKolom lo, COL_DETAIL, 60
    lo.ListColumns(COL_DETAIL).DataBodyRange.WrapText = True
End Sub

Private Function AmbilAtauBuatSheet(ByVal namaSheet As String) As Worksheet
    On Error Resume Next
    Set AmbilAtauBuatSheet = ThisWorkbook.Worksheets(namaSheet)
    On Error GoTo 0

    If AmbilAtauBuatSheet Is Nothing Then
        Set AmbilAtauBuatSheet = ThisWorkbook.Worksheets.Add( _
                                        After:=ThisWorkbook.Worksheets( _
                                        ThisWorkbook.Worksheets.Count))
        AmbilAtauBuatSheet.Name = namaSheet
    End If
End Function

Private Function AmbilAtauBuatTabel( _
    ByVal ws As Worksheet, _
    ByVal namaTabel As String, _
    ByVal headers As Variant, _
    ByRef tabelBaru As Boolean) As ListObject

    Dim i As Long
    Dim jumlahKolom As Long
    Dim rngTabelAwal As Range
    Dim tabelNamaSama As ListObject

    tabelBaru = False

    mTahap = "Mencari tabel " & namaTabel & " pada sheet " & ws.Name
    On Error Resume Next
    Set AmbilAtauBuatTabel = ws.ListObjects(namaTabel)
    On Error GoTo 0

    If AmbilAtauBuatTabel Is Nothing And ws.ListObjects.Count > 0 Then
        Set AmbilAtauBuatTabel = ws.ListObjects(1)
    End If

    If Not AmbilAtauBuatTabel Is Nothing Then
        If StrComp(AmbilAtauBuatTabel.Name, namaTabel, vbTextCompare) <> 0 Then
            Set tabelNamaSama = CariTabelBernama(namaTabel)
            If Not tabelNamaSama Is Nothing Then
                Err.Raise vbObjectError + 108, , _
                          "Nama tabel '" & namaTabel & "' sudah digunakan pada sheet '" & _
                          tabelNamaSama.Parent.Name & "'."
            End If
            mTahap = "Memulihkan nama tabel menjadi " & namaTabel
            AmbilAtauBuatTabel.Name = namaTabel
        End If
        If TabelHanyaBarisSementara(AmbilAtauBuatTabel) Then tabelBaru = True
        Exit Function
    End If

    Set tabelNamaSama = CariTabelBernama(namaTabel)
    If Not tabelNamaSama Is Nothing Then
        Err.Raise vbObjectError + 108, , _
                  "Nama tabel '" & namaTabel & "' sudah digunakan pada sheet '" & _
                  tabelNamaSama.Parent.Name & "'."
    End If

    mTahap = "Memeriksa isi sheet " & ws.Name
    If Not WorksheetKosong(ws) And Not HanyaHeaderSesuai(ws, headers) Then
        Err.Raise vbObjectError + 103, , _
                  "Sheet '" & ws.Name & "' sudah berisi data, tetapi tabel impor tidak ditemukan."
    End If

    jumlahKolom = UBound(headers) - LBound(headers) + 1
    mTahap = "Menulis header tabel " & namaTabel
    For i = 1 To jumlahKolom
        ws.Cells(1, i).Value = CStr(headers(LBound(headers) + i - 1))
    Next i

    ' Gunakan baris sementara berisi nilai. Pada beberapa versi Excel,
    ' membuat tabel lalu langsung menghapus satu-satunya baris memicu error 1004.
    ws.Cells(2, 1).Value2 = TEMP_ROW_MARKER
    Set rngTabelAwal = ws.Range(ws.Cells(1, 1), ws.Cells(2, jumlahKolom))
    mTahap = "Membentuk objek tabel " & namaTabel
    Set AmbilAtauBuatTabel = ws.ListObjects.Add(xlSrcRange, rngTabelAwal, , xlYes)
    mTahap = "Memberi nama tabel " & namaTabel
    AmbilAtauBuatTabel.Name = namaTabel
    mTahap = "Menerapkan gaya tabel " & namaTabel
    AmbilAtauBuatTabel.TableStyle = "TableStyleMedium2"
    ws.Cells(1, 1).Resize(1, jumlahKolom).Font.Bold = True
    tabelBaru = True
End Function

Private Function CariTabelBernama(ByVal namaTabel As String) As ListObject
    Dim ws As Worksheet
    Dim lo As ListObject

    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.Name, namaTabel, vbTextCompare) = 0 Then
                Set CariTabelBernama = lo
                Exit Function
            End If
        Next lo
    Next ws
End Function

Private Function TabelHanyaBarisSementara(ByVal lo As ListObject) As Boolean
    Dim sel As Range

    If lo.ListRows.Count <> 1 Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function

    If CStr(lo.DataBodyRange.Cells(1, 1).Value2) = TEMP_ROW_MARKER Then
        TabelHanyaBarisSementara = True
        Exit Function
    End If

    For Each sel In lo.DataBodyRange.Cells
        If Len(CStr(sel.Value2)) > 0 Then Exit Function
    Next sel

    TabelHanyaBarisSementara = True
End Function

Private Function WorksheetKosong(ByVal ws As Worksheet) As Boolean
    Dim selTerakhir As Range

    On Error Resume Next
    Set selTerakhir = ws.Cells.Find( _
        What:="*", _
        After:=ws.Cells(1, 1), _
        LookIn:=xlFormulas, _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, _
        MatchCase:=False, _
        SearchFormat:=False)
    On Error GoTo 0

    WorksheetKosong = (selTerakhir Is Nothing)
End Function

Private Function HanyaHeaderSesuai( _
    ByVal ws As Worksheet, _
    ByVal headers As Variant) As Boolean

    Dim selTerakhir As Range
    Dim jumlahKolom As Long
    Dim i As Long

    On Error Resume Next
    Set selTerakhir = ws.Cells.Find( _
        What:="*", _
        After:=ws.Cells(1, 1), _
        LookIn:=xlFormulas, _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious, _
        MatchCase:=False, _
        SearchFormat:=False)
    On Error GoTo 0

    If selTerakhir Is Nothing Then Exit Function
    If selTerakhir.Row <> 1 Then Exit Function

    jumlahKolom = UBound(headers) - LBound(headers) + 1
    If ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column <> jumlahKolom Then Exit Function

    For i = 1 To jumlahKolom
        If StrComp(Trim$(CStr(ws.Cells(1, i).Value2)), _
                   Trim$(CStr(headers(LBound(headers) + i - 1))), _
                   vbTextCompare) <> 0 Then Exit Function
    Next i

    HanyaHeaderSesuai = True
End Function

Private Sub ValidasiHeaderTabel(ByVal lo As ListObject, ByVal headers As Variant)
    Dim jumlahKolom As Long
    Dim i As Long

    jumlahKolom = UBound(headers) - LBound(headers) + 1
    If lo.ListColumns.Count <> jumlahKolom Then
        Err.Raise vbObjectError + 104, , _
                  "Struktur kolom tabel '" & lo.Name & "' berbeda dari file sumber."
    End If

    For i = 1 To jumlahKolom
        If StrComp(Trim$(CStr(lo.HeaderRowRange.Cells(1, i).Value2)), _
                   Trim$(CStr(headers(LBound(headers) + i - 1))), _
                   vbTextCompare) <> 0 Then
            Err.Raise vbObjectError + 105, , _
                      "Header kolom ke-" & i & " pada tabel '" & lo.Name & _
                      "' tidak sesuai dengan file sumber."
        End If
    Next i
End Sub

Private Sub HapusMasaYangSama(ByVal lo As ListObject, ByVal periods As Object)
    Dim idxMasa As Long
    Dim idxTahun As Long
    Dim i As Long
    Dim key As String

    If lo.DataBodyRange Is Nothing Then Exit Sub

    idxMasa = lo.ListColumns("TAX_INVOICE_PERIOD").Index
    idxTahun = lo.ListColumns("TAX_INVOICE_YEAR").Index

    For i = lo.ListRows.Count To 1 Step -1
        key = NormalisasiMasa( _
                CStr(lo.DataBodyRange.Cells(i, idxTahun).Value2), _
                CStr(lo.DataBodyRange.Cells(i, idxMasa).Value2))
        If periods.Exists(key) Then lo.ListRows(i).Delete
    Next i
End Sub

Private Function BukaTextStream(ByVal filePath As String) As Object
    Dim stm As Object

    Set stm = CreateObject("ADODB.Stream")
    With stm
        .Type = 2 ' adTypeText
        .Charset = DeteksiCharset(filePath)
        .Open
        .LoadFromFile filePath
        .Position = 0
    End With

    Set BukaTextStream = stm
End Function

Private Function DeteksiCharset(ByVal filePath As String) As String
    Dim f As Integer
    Dim b(0 To 2) As Byte
    Dim panjang As Long

    On Error GoTo GunakanANSI
    f = FreeFile
    Open filePath For Binary Access Read As #f
    panjang = LOF(f)
    If panjang >= 3 Then Get #f, 1, b
    Close #f

    If panjang >= 3 Then
        If b(0) = 239 And b(1) = 187 And b(2) = 191 Then
            DeteksiCharset = "utf-8"
            Exit Function
        End If
    End If

    If panjang >= 2 Then
        If b(0) = 255 And b(1) = 254 Then
            DeteksiCharset = "unicode"
            Exit Function
        ElseIf b(0) = 254 And b(1) = 255 Then
            DeteksiCharset = "unicodeFFFE"
            Exit Function
        End If
    End If

GunakanANSI:
    On Error Resume Next
    If f > 0 Then Close #f
    On Error GoTo 0
    DeteksiCharset = "windows-1252"
End Function

Private Function BacaRecordBerikutnya( _
    ByVal stm As Object, _
    ByRef recordText As String) As Boolean

    Dim lineText As String

    recordText = vbNullString
    Do While Not stm.EOS
        lineText = stm.ReadText(-2) ' adReadLine
        If Len(recordText) > 0 Then
            recordText = recordText & vbLf & lineText
        Else
            recordText = lineText
        End If

        If KutipSeimbang(recordText) Then
            BacaRecordBerikutnya = True
            Exit Function
        End If
    Loop

    BacaRecordBerikutnya = (Len(recordText) > 0)
End Function

Private Function KutipSeimbang(ByVal teks As String) As Boolean
    Dim i As Long
    Dim dalamKutip As Boolean

    i = 1
    Do While i <= Len(teks)
        If Mid$(teks, i, 1) = """" Then
            If dalamKutip And i < Len(teks) And Mid$(teks, i + 1, 1) = """" Then
                i = i + 1
            Else
                dalamKutip = Not dalamKutip
            End If
        End If
        i = i + 1
    Loop

    KutipSeimbang = Not dalamKutip
End Function

Private Function ParsePipeLine(ByVal teks As String) As Variant
    Dim hasil() As String
    Dim token As String
    Dim ch As String
    Dim dalamKutip As Boolean
    Dim i As Long
    Dim n As Long

    ReDim hasil(0 To 0)
    n = 0
    i = 1

    Do While i <= Len(teks)
        ch = Mid$(teks, i, 1)

        If ch = """" Then
            If dalamKutip And i < Len(teks) And Mid$(teks, i + 1, 1) = """" Then
                token = token & """"
                i = i + 1
            Else
                dalamKutip = Not dalamKutip
            End If
        ElseIf ch = "|" And Not dalamKutip Then
            hasil(n) = token
            token = vbNullString
            n = n + 1
            ReDim Preserve hasil(0 To n)
        Else
            token = token & ch
        End If

        i = i + 1
    Loop

    hasil(n) = token
    ParsePipeLine = hasil
End Function

Private Function BuatIndeksHeader(ByVal header As Variant) As Object
    Dim hasil As Object
    Dim i As Long
    Dim nama As String

    Set hasil = CreateObject("Scripting.Dictionary")
    hasil.CompareMode = vbTextCompare

    For i = LBound(header) To UBound(header)
        nama = UCase$(Trim$(CStr(header(i))))
        If i = LBound(header) Then nama = HapusBOM(nama)
        If Len(nama) > 0 Then hasil(nama) = i - LBound(header)
    Next i

    Set BuatIndeksHeader = hasil
End Function

Private Sub WajibAdaKolom(ByVal idx As Object, ByVal nama As String, ByVal filePath As String)
    If Not idx.Exists(nama) Then
        Err.Raise vbObjectError + 106, , _
                  "Kolom wajib '" & nama & "' tidak ditemukan pada " & NamaFile(filePath)
    End If
End Sub

Private Function AmbilField(ByVal fields As Variant, ByVal indexZeroBased As Long) As String
    Dim actualIndex As Long

    actualIndex = LBound(fields) + indexZeroBased
    If actualIndex >= LBound(fields) And actualIndex <= UBound(fields) Then
        AmbilField = CStr(fields(actualIndex))
    Else
        AmbilField = vbNullString
    End If
End Function

Private Function AmbilFieldOpsional( _
    ByVal fields As Variant, _
    ByVal idx As Object, _
    ByVal namaKolom As String) As String

    If idx.Exists(namaKolom) Then
        AmbilFieldOpsional = AmbilField(fields, CLng(idx(namaKolom)))
    Else
        AmbilFieldOpsional = vbNullString
    End If
End Function

Private Function BarisMemilikiIsi(ByVal fields As Variant) As Boolean
    Dim i As Long

    For i = LBound(fields) To UBound(fields)
        If Len(Trim$(CStr(fields(i)))) > 0 Then
            BarisMemilikiIsi = True
            Exit Function
        End If
    Next i
End Function

Private Function BersihkanNamaBarang(ByVal nilai As String) As String
    nilai = Replace(nilai, vbCr, " ")
    nilai = Replace(nilai, vbLf, " ")
    nilai = Replace(nilai, vbTab, " ")

    Do While InStr(nilai, "  ") > 0
        nilai = Replace(nilai, "  ", " ")
    Loop

    BersihkanNamaBarang = Trim$(nilai)
End Function

Private Function GabungkanDetail(ByVal nilaiLama As String, ByVal nilaiBaru As String) As String
    Const MAX_CELL_TEXT As Long = 32767
    Const SUFFIX As String = " ... [dipotong]"
    Dim gabungan As String

    If Len(nilaiLama) = 0 Then
        gabungan = nilaiBaru
    Else
        gabungan = nilaiLama & ", " & nilaiBaru
    End If

    If Len(gabungan) > MAX_CELL_TEXT Then
        gabungan = Left$(gabungan, MAX_CELL_TEXT - Len(SUFFIX)) & SUFFIX
    End If

    GabungkanDetail = gabungan
End Function

Private Function BentukKunciFaktur( _
    ByVal nomorFaktur As String, _
    ByVal tahun As String, _
    ByVal masa As String) As String

    Dim nomorNorm As String
    Dim masaNorm As String

    nomorNorm = NormalisasiNomorFaktur(nomorFaktur)
    masaNorm = NormalisasiMasa(tahun, masa)

    If Len(nomorNorm) = 0 Or Len(masaNorm) = 0 Then
        BentukKunciFaktur = vbNullString
    Else
        BentukKunciFaktur = nomorNorm & "|" & masaNorm
    End If
End Function

Private Function NormalisasiNomorFaktur(ByVal nilai As String) As String
    Dim hasil As String
    Dim ch As String
    Dim i As Long

    nilai = UCase$(Trim$(nilai))
    For i = 1 To Len(nilai)
        ch = Mid$(nilai, i, 1)
        If (ch >= "0" And ch <= "9") Or (ch >= "A" And ch <= "Z") Then
            hasil = hasil & ch
        End If
    Next i

    NormalisasiNomorFaktur = hasil
End Function

Private Function NormalisasiMasa(ByVal tahun As String, ByVal masa As String) As String
    Dim y As Long
    Dim m As Long

    If Not IsNumeric(Trim$(tahun)) Or Not IsNumeric(Trim$(masa)) Then Exit Function
    y = CLng(Val(tahun))
    m = CLng(Val(masa))
    If y < 1900 Or y > 9999 Or m < 1 Or m > 12 Then Exit Function

    NormalisasiMasa = Format$(y, "0000") & Format$(m, "00")
End Function

Private Function HapusBOM(ByVal nilai As String) As String
    If Len(nilai) > 0 Then
        If AscW(Left$(nilai, 1)) = -257 Or AscW(Left$(nilai, 1)) = &HFEFF Then
            nilai = Mid$(nilai, 2)
        End If
    End If
    HapusBOM = nilai
End Function

Private Function KonversiNilai( _
    ByVal namaKolom As String, _
    ByVal nilai As String, _
    Optional ByVal penandaDokumen As String = vbNullString) As Variant
    Dim nama As String
    Dim nilaiBersih As String

    nama = UCase$(Trim$(namaKolom))
    nilaiBersih = Trim$(nilai)

    If Len(nilaiBersih) = 0 Then
        KonversiNilai = vbNullString
        Exit Function
    End If

    Select Case nama
        Case "SELLING_PRICE", "OTHER_TAX_BASE", "VAT"
            KonversiNilai = ParseAngkaAman(nilaiBersih)
        Case "TAX_INVOICE_PERIOD", "TAX_INVOICE_YEAR"
            If IsNumeric(nilaiBersih) Then
                KonversiNilai = CLng(Val(nilaiBersih))
            Else
                KonversiNilai = nilaiBersih
            End If
        Case "TAX_INVOICE_NUMBER"
            'Hanya EFAKTUR yang dinormalisasi ke format NSFP.
            'DOKUMEN LAIN harus mempertahankan nomor aslinya.
            If InStr(1, UCase$(Trim$(penandaDokumen)), "EFAKTUR", _
                     vbTextCompare) > 0 Then
                KonversiNilai = FormatNomorFakturTeks(nilai)
            Else
                KonversiNilai = nilai
            End If
        Case "TAX_INVOICE_DATE"
            KonversiNilai = ParseTanggalFaktur(nilaiBersih)
        Case Else
            ' Kolom teks dipertahankan apa adanya, termasuk nol di depan.
            KonversiNilai = nilai
    End Select
End Function

Private Function ParseAngkaAman(ByVal nilai As String) As Variant
    Dim s As String
    Dim kandidat As String
    Dim decSep As String
    Dim posDot As Long
    Dim posComma As Long
    Dim posDec As Long
    Dim jumlahDigitBelakang As Long

    s = Trim$(nilai)
    s = Replace$(s, " ", vbNullString)
    s = Replace$(s, "'", vbNullString)
    If Len(s) = 0 Then
        ParseAngkaAman = vbNullString
        Exit Function
    End If

    decSep = Application.International(xlDecimalSeparator)
    posDot = InStrRev(s, ".")
    posComma = InStrRev(s, ",")

    If posDot > 0 And posComma > 0 Then
        'Pemisah yang paling akhir diperlakukan sebagai desimal.
        If posDot > posComma Then
            posDec = posDot
        Else
            posDec = posComma
        End If
        kandidat = Left$(s, posDec - 1)
        kandidat = Replace$(kandidat, ".", vbNullString)
        kandidat = Replace$(kandidat, ",", vbNullString)
        kandidat = kandidat & decSep & Mid$(s, posDec + 1)
    ElseIf posDot > 0 Or posComma > 0 Then
        If posDot > 0 Then
            posDec = posDot
        Else
            posDec = posComma
        End If
        jumlahDigitBelakang = Len(s) - posDec

        If jumlahDigitBelakang = 3 Then
            'Pola tunggal xxx.xxx / xxx,xxx diperlakukan sebagai ribuan.
            kandidat = Replace$(s, ".", vbNullString)
            kandidat = Replace$(kandidat, ",", vbNullString)
        Else
            If posDot > 0 Then
                kandidat = Replace$(s, ".", decSep)
            Else
                kandidat = Replace$(s, ",", decSep)
            End If
        End If
    Else
        kandidat = s
    End If

    If IsNumeric(kandidat) Then
        ParseAngkaAman = CDbl(kandidat)
    Else
        ParseAngkaAman = nilai
    End If
End Function

Private Function FormatNomorFakturTeks(ByVal nilai As String) As String
    Dim s As String
    Dim i As Long
    Dim ch As String
    Dim semuaAngka As Boolean

    s = Trim$(nilai)
    If Len(s) = 0 Then
        FormatNomorFakturTeks = vbNullString
        Exit Function
    End If

    'Hilangkan pemisah hanya untuk menguji format NSFP.
    s = Replace$(s, ".", vbNullString)
    s = Replace$(s, "-", vbNullString)
    s = Replace$(s, "_", vbNullString)
    s = Replace$(s, " ", vbNullString)

    semuaAngka = (Len(s) = 16)
    If semuaAngka Then
        For i = 1 To Len(s)
            ch = Mid$(s, i, 1)
            If ch < "0" Or ch > "9" Then
                semuaAngka = False
                Exit For
            End If
        Next i
    End If

    If semuaAngka Then
        FormatNomorFakturTeks = _
            Left$(s, 3) & "." & _
            Mid$(s, 4, 3) & "-" & _
            Mid$(s, 7, 2) & "." & _
            Right$(s, 8)
    Else
        FormatNomorFakturTeks = nilai
    End If
End Function

Private Function ParseTanggalFaktur(ByVal nilai As String) As Variant
    Dim bagian As Variant
    Dim bulan As Object
    Dim d As Long
    Dim m As Long
    Dim y As Long

    bagian = Split(UCase$(Trim$(nilai)), "-")
    If UBound(bagian) <> 2 Then
        ParseTanggalFaktur = nilai
        Exit Function
    End If

    Set bulan = CreateObject("Scripting.Dictionary")
    bulan.CompareMode = vbTextCompare
    bulan.Add "JAN", 1: bulan.Add "FEB", 2: bulan.Add "MAR", 3
    bulan.Add "APR", 4: bulan.Add "MAY", 5: bulan.Add "JUN", 6
    bulan.Add "JUL", 7: bulan.Add "AUG", 8: bulan.Add "SEP", 9
    bulan.Add "OCT", 10: bulan.Add "NOV", 11: bulan.Add "DEC", 12

    If Not IsNumeric(bagian(0)) Or Not bulan.Exists(bagian(1)) Or _
       Not IsNumeric(bagian(2)) Then
        ParseTanggalFaktur = nilai
        Exit Function
    End If

    d = CLng(bagian(0))
    m = CLng(bulan(bagian(1)))
    y = CLng(bagian(2))
    If y < 100 Then y = 2000 + y

    On Error GoTo TanggalTidakValid
    ParseTanggalFaktur = DateSerial(y, m, d)
    Exit Function

TanggalTidakValid:
    ParseTanggalFaktur = nilai
End Function

Private Sub TerapkanFormatTabel(ByVal lo As ListObject)
    Dim lc As ListColumn

    For Each lc In lo.ListColumns
        lc.Range.NumberFormat = FormatKolomHasil(lc.Name)
    Next lc
End Sub

Private Sub TerapkanFormatSebelumTulis( _
    ByVal ws As Worksheet, _
    ByVal headers As Variant, _
    ByVal jumlahBaris As Long)

    Dim i As Long
    Dim namaKolom As String

    For i = LBound(headers) To UBound(headers)
        namaKolom = CStr(headers(i))
        ws.Cells(2, i - LBound(headers) + 1).Resize(jumlahBaris, 1) _
          .NumberFormat = FormatKolomHasil(namaKolom)
    Next i
End Sub

Private Function FormatKolomHasil(ByVal namaKolom As String) As String
    Select Case UCase$(Trim$(namaKolom))
        Case "TAX_INVOICE_DATE"
            FormatKolomHasil = "dd-mmm-yyyy"
        Case "TAX_INVOICE_PERIOD"
            FormatKolomHasil = "00"
        Case "TAX_INVOICE_YEAR"
            FormatKolomHasil = "0"
        Case "SELLING_PRICE", "OTHER_TAX_BASE", "VAT"
            FormatKolomHasil = "#,##0.############"
        Case Else
            FormatKolomHasil = "@"
    End Select
End Function

Private Function NamaFile(ByVal fullPath As String) As String
    NamaFile = Mid$(fullPath, InStrRev(fullPath, Application.PathSeparator) + 1)
End Function

Private Function KlasifikasiFile(ByVal namaLower As String) As String
    Dim isDetail As Boolean
    Dim isPenjual As Boolean
    Dim isPembeli As Boolean

    isDetail = (InStr(1, namaLower, "objekfakturdetil_", vbTextCompare) > 0)
    isPenjual = (InStr(1, namaLower, "_penjual_", vbTextCompare) > 0)
    isPembeli = (InStr(1, namaLower, "_pembeli_", vbTextCompare) > 0)

    If isDetail And isPenjual Then
        KlasifikasiFile = "PENJUAL_DETAIL"
    ElseIf isDetail And isPembeli Then
        KlasifikasiFile = "PEMBELI_DETAIL"
    ElseIf Not isDetail And InStr(1, namaLower, "objekfaktur_", vbTextCompare) > 0 And isPenjual Then
        KlasifikasiFile = "PENJUAL_INDUK"
    ElseIf Not isDetail And InStr(1, namaLower, "objekfaktur_", vbTextCompare) > 0 And isPembeli Then
        KlasifikasiFile = "PEMBELI_INDUK"
    End If
End Function

Private Function EkstrakMasaDariNamaFile(ByVal namaLower As String) As String
    Dim posisi As Long
    Dim kandidat As String
    Dim bulan As Long

    posisi = InStr(1, namaLower, "_tahun_", vbTextCompare)
    If posisi = 0 Then Exit Function

    kandidat = Mid$(namaLower, posisi + Len("_tahun_"), 6)
    If Len(kandidat) <> 6 Or Not IsNumeric(kandidat) Then Exit Function

    bulan = CLng(Right$(kandidat, 2))
    If bulan < 1 Or bulan > 12 Then Exit Function

    EkstrakMasaDariNamaFile = kandidat
End Function

Private Sub TambahBarisDaftar( _
    ByVal frm As Object, _
    ByVal jenis As String, _
    ByVal masa As String, _
    ByVal filePath As String)

    Dim arah As String
    Dim tipe As String
    Dim baris As Long

    If Left$(jenis, 7) = "PENJUAL" Then
        arah = "PK"
    Else
        arah = "PM"
    End If

    If Right$(jenis, 6) = "_INDUK" Then
        tipe = "Induk"
    Else
        tipe = "Detail"
    End If

    frm.lstFile.AddItem
    baris = frm.lstFile.ListCount - 1
    frm.lstFile.List(baris, 0) = arah
    frm.lstFile.List(baris, 1) = tipe
    frm.lstFile.List(baris, 2) = masa
    frm.lstFile.List(baris, 3) = NamaFile(filePath)
    frm.lstFile.List(baris, 4) = filePath
End Sub

Private Function BangunPetaFileDariForm(ByVal frm As Object) As Object
    Dim hasil As Object
    Dim i As Long
    Dim arah As String
    Dim tipe As String
    Dim jenis As String
    Dim masa As String

    Set hasil = CreateObject("Scripting.Dictionary")
    hasil.CompareMode = vbTextCompare

    For i = 0 To frm.lstFile.ListCount - 1
        arah = UCase$(CStr(frm.lstFile.List(i, 0)))
        tipe = UCase$(CStr(frm.lstFile.List(i, 1)))
        masa = CStr(frm.lstFile.List(i, 2))

        If arah = "PK" Then
            jenis = "PENJUAL_" & tipe
        Else
            jenis = "PEMBELI_" & tipe
        End If

        hasil(jenis & "|" & masa) = CStr(frm.lstFile.List(i, 4))
    Next i

    Set BangunPetaFileDariForm = hasil
End Function

Private Function ValidasiDaftarFile(ByVal frm As Object, ByRef pesan As String) As Boolean
    Dim fileMap As Object
    Dim periods As Object
    Dim counts As Object
    Dim i As Long
    Dim key As Variant
    Dim bagian As Variant
    Dim jenis As String
    Dim masa As String
    Dim pasangan As String
    Dim path As String

    If frm.lstFile.ListCount = 0 Then
        pesan = "Belum ada file CSV yang dipilih."
        Exit Function
    End If

    Set fileMap = BangunPetaFileDariForm(frm)
    Set periods = CreateObject("Scripting.Dictionary")
    Set counts = CreateObject("Scripting.Dictionary")
    periods.CompareMode = vbTextCompare
    counts.CompareMode = vbTextCompare

    For i = 0 To frm.lstFile.ListCount - 1
        path = CStr(frm.lstFile.List(i, 4))
        If Len(Dir$(path)) = 0 Then
            pesan = "File tidak ditemukan: " & NamaFile(path)
            Exit Function
        End If
    Next i

    For Each key In fileMap.Keys
        bagian = Split(CStr(key), "|")
        jenis = CStr(bagian(0))
        masa = CStr(bagian(1))
        periods(masa) = True
        If counts.Exists(jenis) Then
            counts(jenis) = CLng(counts(jenis)) + 1
        Else
            counts.Add jenis, 1
        End If

        If Right$(jenis, 6) = "_INDUK" Then
            pasangan = Replace(jenis, "_INDUK", "_DETAIL") & "|" & masa
        Else
            pasangan = Replace(jenis, "_DETAIL", "_INDUK") & "|" & masa
        End If

        If Not fileMap.Exists(pasangan) Then
            pesan = "Pasangan file belum lengkap untuk " & _
                    LabelJenisFile(jenis) & " masa " & masa & "."
            Exit Function
        End If
    Next key

    If periods.Count > 12 Then
        pesan = "Masa pajak yang dipilih melebihi batas maksimal 12 masa."
        Exit Function
    End If

    For Each key In counts.Keys
        If CLng(counts(key)) > 12 Then
            pesan = "Jumlah " & LabelJenisFile(CStr(key)) & " melebihi 12 file."
            Exit Function
        End If
    Next key

    ValidasiDaftarFile = True
End Function

Private Sub PerbaruiStatusForm(ByVal frm As Object)
    Dim pesan As String

    If ValidasiDaftarFile(frm, pesan) Then
        frm.cmdImpor.Enabled = True
    Else
        frm.cmdImpor.Enabled = False
    End If
End Sub

Private Function LabelJenisFile(ByVal jenis As String) As String
    Select Case UCase$(jenis)
        Case "PENJUAL_INDUK": LabelJenisFile = "induk PK"
        Case "PENJUAL_DETAIL": LabelJenisFile = "detail PK"
        Case "PEMBELI_INDUK": LabelJenisFile = "induk PM"
        Case "PEMBELI_DETAIL": LabelJenisFile = "detail PM"
        Case Else: LabelJenisFile = jenis
    End Select
End Function

Private Sub TulisLog( _
    ByVal level As String, _
    ByVal proses As String, _
    ByVal pesan As String, _
    ByVal filePath As String)

    Dim ws As Worksheet
    Dim nextRow As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_LOG)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add( _
                    After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = SHEET_LOG
        ws.Range("A1:E1").Value = Array("WAKTU", "LEVEL", "PROSES", "PESAN", "FILE")
        ws.Range("A1:E1").Font.Bold = True
    End If

    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(nextRow, 1).Value = Now
    ws.Cells(nextRow, 1).NumberFormat = "dd-mmm-yyyy hh:mm:ss"
    ws.Cells(nextRow, 2).Value = level
    ws.Cells(nextRow, 3).Value = proses
    ws.Cells(nextRow, 4).Value = pesan
    ws.Cells(nextRow, 5).Value = filePath
End Sub

Private Sub BatasiLebarKolom(ByVal lo As ListObject, ByVal namaKolom As String, ByVal maxWidth As Double)
    On Error Resume Next
    If lo.ListColumns(namaKolom).Range.ColumnWidth > maxWidth Then
        lo.ListColumns(namaKolom).Range.ColumnWidth = maxWidth
    End If
    On Error GoTo 0
End Sub

Private Sub PulihkanExcel( _
    ByVal oldScreen As Boolean, _
    ByVal oldEvents As Boolean, _
    ByVal oldAlerts As Boolean, _
    ByVal oldCalc As XlCalculation)

    Application.StatusBar = False
    Application.ScreenUpdating = oldScreen
    Application.EnableEvents = oldEvents
    Application.DisplayAlerts = oldAlerts
    Application.Calculation = oldCalc
End Sub

