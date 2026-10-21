Attribute VB_Name = "modSinkronFakturKePPN"
Option Explicit

' ================================================================
'  ENGINE SINKRONISASI FAKTUR KE SHEET PPN v0.54
'  Sumber  : tblFakturPK pada sheet FakturPK
'            tblFakturPM pada sheet FakturPM
'  Target  : PPN-N-1-2 s.d. PPN-N-24-2 sesuai rule yang ditetapkan
'  Catatan : N-5-2, N-9-2, N-10-2, N-23-2, dan N-25-2 bersifat opsional.
'            N-23-2 dilewati; target lain menunggu sumber data KKP Lokal.
' ================================================================

Private Const SHEET_SUMBER_PK As String = "FakturPK"
Private Const SHEET_SUMBER_PM As String = "FakturPM"
Private Const TABEL_SUMBER_PK As String = "tblFakturPK"
Private Const TABEL_SUMBER_PM As String = "tblFakturPM"
Private Const SHEET_LOG As String = "Log Sinkron PPN"

Private Const TARGET_N1 As String = "PPN-N-1-2"
Private Const TARGET_N2 As String = "PPN-N-2-2"
Private Const TARGET_N3 As String = "PPN-N-3-2"
Private Const TARGET_N4 As String = "PPN-N-4-2"
Private Const TARGET_N5 As String = "PPN-N-5-2"
Private Const TARGET_N6 As String = "PPN-N-6-2"
Private Const TARGET_N7 As String = "PPN-N-7-2"
Private Const TARGET_N8 As String = "PPN-N-8-2"
Private Const TARGET_N9 As String = "PPN-N-9-2"
Private Const TARGET_N10 As String = "PPN-N-10-2"
Private Const TARGET_N19 As String = "PPN-N-19-2"
Private Const TARGET_N20 As String = "PPN-N-20-2"
Private Const TARGET_N21 As String = "PPN-N-21-2"
Private Const TARGET_N22 As String = "PPN-N-22-2"
Private Const TARGET_N23 As String = "PPN-N-23-2"
Private Const TARGET_N24 As String = "PPN-N-24-2"
Private Const TARGET_N25 As String = "PPN-N-25-2"

Private mTahap As String
Private mRingkasanMasa As Object

Public Sub SinkronkanFakturKePPN()
    Dim oldScreen As Boolean
    Dim oldEvents As Boolean
    Dim oldAlerts As Boolean
    Dim oldCalc As XlCalculation
    Dim dataTarget As Object
    Dim ringkasan As String
    Dim jumlahTidakTerpetakanPK As Long
    Dim jumlahTidakTerpetakanPM As Long
    Dim jumlahBukanApprovedPK As Long
    Dim jumlahBukanApprovedPM As Long
    Dim namaTarget As Variant
    Dim errNum As Long
    Dim errDesc As String

    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    oldAlerts = Application.DisplayAlerts
    oldCalc = Application.Calculation

    On Error GoTo Gagal

    mTahap = "Memeriksa sumber dan seluruh sheet target"
    ValidasiSeluruhStruktur

    mTahap = "Membaca dan memetakan FakturPK serta FakturPM"
    Set dataTarget = BangunSeluruhDataTarget( _
                        jumlahTidakTerpetakanPK, jumlahTidakTerpetakanPM, _
                        jumlahBukanApprovedPK, jumlahBukanApprovedPM)

    mTahap = "Memeriksa data lama pada sheet target"
    ringkasan = RingkasanDataTargetLama()
    If Len(ringkasan) > 0 Then
        If MsgBox( _
            "Ditemukan Data pada Sheet tujuan, proses ini akan menimpa data tersebut.", _
            vbYesNo + vbQuestion + vbDefaultButton2, _
            "Konfirmasi Sinkronisasi Faktur") <> vbYes Then

            TulisLogSinkron "DIBATALKAN", "Pengguna membatalkan karena target sudah berisi data."
            Exit Sub
        End If
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Menyinkronkan faktur ke sheet PPN..."

    For Each namaTarget In DaftarTarget()
        mTahap = "Menulis data ke " & CStr(namaTarget)
        TulisDataTarget CStr(namaTarget), dataTarget(CStr(namaTarget))
    Next namaTarget

    For Each namaTarget In DaftarTarget()
        If CStr(namaTarget) = TARGET_N1 Then
            mTahap = "Memperbarui CFM Pemeriksa pada PPN-N-1"
            PerbaruiCFMPemeriksaN1
        ElseIf TargetMemilikiSheetInduk(CStr(namaTarget)) Then
            mTahap = "Memperbarui referensi dan nilai sheet induk untuk " & _
                     CStr(namaTarget)
            PerbaruiSheetInduk CStr(namaTarget)
        End If
    Next namaTarget

    mTahap = "Menandai faktur yang digunakan dan tidak digunakan"
    TandaiStatusSinkronisasi

    ringkasan = RingkasanHasil(dataTarget)
    TulisLogSinkron "BERHASIL", _
        Replace(ringkasan, vbCrLf, "; ") & _
        "; PK tidak terpetakan: " & Format$(jumlahTidakTerpetakanPK, "#,##0") & _
        "; PM tidak terpetakan: " & Format$(jumlahTidakTerpetakanPM, "#,##0") & _
        "; PK bukan Approved: " & Format$(jumlahBukanApprovedPK, "#,##0") & _
        "; PM bukan Approved: " & Format$(jumlahBukanApprovedPM, "#,##0")

    PulihkanExcel oldScreen, oldEvents, oldAlerts, oldCalc

    MsgBox "Proses mapping selesai.", vbInformation, "Sinkronisasi Faktur ke PPN"
    Exit Sub

Gagal:
    errNum = Err.Number
    errDesc = Err.Description
    On Error Resume Next
    TulisLogSinkron "GAGAL", "Tahap: " & mTahap & ". " & errDesc
    PulihkanExcel oldScreen, oldEvents, oldAlerts, oldCalc
    On Error GoTo 0
    MsgBox "Sinkronisasi gagal pada tahap:" & vbCrLf & mTahap & vbCrLf & vbCrLf & _
           errDesc & vbCrLf & vbCrLf & _
           "Tidak ada proses lanjutan yang dijalankan.", _
           vbExclamation, "Sinkronisasi Faktur ke PPN"
    Err.Clear
End Sub

Private Sub TandaiStatusSinkronisasi()
    TandaiStatusSumber SHEET_SUMBER_PK, TABEL_SUMBER_PK, True
    TandaiStatusSumber SHEET_SUMBER_PM, TABEL_SUMBER_PM, False
End Sub

Private Sub TandaiStatusSumber( _
    ByVal namaSheet As String, _
    ByVal namaTabel As String, _
    ByVal sumberPK As Boolean)

    Const NAMA_KOLOM_STATUS As String = "STATUS_SINKRONISASI"
    Dim lo As ListObject
    Dim idx As Object
    Dim data As Variant
    Dim hasil() As Variant
    Dim kolomStatus As ListColumn
    Dim r As Long
    Dim statusFaktur As String
    Dim target As String

    Set lo = AmbilTabelWajib(namaSheet, namaTabel)

    On Error Resume Next
    Set kolomStatus = lo.ListColumns(NAMA_KOLOM_STATUS)
    On Error GoTo 0

    If kolomStatus Is Nothing Then
        Set kolomStatus = lo.ListColumns.Add
        kolomStatus.Name = NAMA_KOLOM_STATUS
    End If

    If lo.DataBodyRange Is Nothing Then Exit Sub

    Set idx = IndeksKolom(lo)
    data = lo.DataBodyRange.Value2
    ReDim hasil(1 To UBound(data, 1), 1 To 1)

    For r = 1 To UBound(data, 1)
        statusFaktur = NormalisasiTeks(data(r, idx("TAX_INVOICE_STATUS")))
        target = vbNullString

        If statusFaktur = "APPROVED" Then
            If sumberPK Then
                target = TentukanTargetPK( _
                            data(r, idx("BUYER_TIN")), _
                            data(r, idx("TAX_INVOICE_CODE")), _
                            NilaiKolomOpsional(data, r, idx, "TABEL"))
            Else
                target = TentukanTargetPM( _
                            data(r, idx("SELLER_TIN")), _
                            data(r, idx("BUYER_STATUS")), _
                            data(r, idx("TAX_INVOICE_CODE")))
            End If
        End If

        If Len(target) > 0 Then
            hasil(r, 1) = "DIGUNAKAN"
        Else
            hasil(r, 1) = "TIDAK DIGUNAKAN"
        End If
    Next r

    kolomStatus.DataBodyRange.Value2 = hasil
    With kolomStatus.DataBodyRange
        .NumberFormat = "@"
        .HorizontalAlignment = xlLeft
        .WrapText = False
    End With
End Sub

Private Function BangunSeluruhDataTarget( _
    ByRef tidakTerpetakanPK As Long, _
    ByRef tidakTerpetakanPM As Long, _
    ByRef bukanApprovedPK As Long, _
    ByRef bukanApprovedPM As Long) As Object

    Dim kumpulan As Object
    Dim hasil As Object
    Dim namaTarget As Variant
    Dim col As Collection
    Dim ringkasanMasa As Object

    Set kumpulan = CreateObject("Scripting.Dictionary")
    kumpulan.CompareMode = vbTextCompare

    For Each namaTarget In DaftarTarget()
        Set col = New Collection
        kumpulan.Add CStr(namaTarget), col
    Next namaTarget

    BacaSumberPK kumpulan, tidakTerpetakanPK, bukanApprovedPK
    BacaSumberPM kumpulan, tidakTerpetakanPM, bukanApprovedPM

    Set hasil = CreateObject("Scripting.Dictionary")
    hasil.CompareMode = vbTextCompare
    Set mRingkasanMasa = CreateObject("Scripting.Dictionary")
    mRingkasanMasa.CompareMode = vbTextCompare

    For Each namaTarget In DaftarTarget()
        Set col = kumpulan(CStr(namaTarget))
        Set ringkasanMasa = BangunRingkasanMasa(CStr(namaTarget), col)
        mRingkasanMasa.Add CStr(namaTarget), ringkasanMasa
        hasil.Add CStr(namaTarget), SusunDataDenganPemisah( _
                                      CStr(namaTarget), col)
    Next namaTarget

    Set BangunSeluruhDataTarget = hasil
End Function

Private Sub BacaSumberPK( _
    ByVal kumpulan As Object, _
    ByRef tidakTerpetakan As Long, _
    ByRef bukanApproved As Long)

    Dim lo As ListObject
    Dim idx As Object
    Dim data As Variant
    Dim r As Long
    Dim target As String
    Dim status As String
    Dim col As Collection

    Set lo = AmbilTabelWajib(SHEET_SUMBER_PK, TABEL_SUMBER_PK)
    Set idx = IndeksKolom(lo)
    ValidasiKolomSumber idx, True

    If lo.DataBodyRange Is Nothing Then Exit Sub
    data = lo.DataBodyRange.Value2

    For r = 1 To UBound(data, 1)
        status = NormalisasiTeks(data(r, idx("TAX_INVOICE_STATUS")))
        If status <> "APPROVED" Then
            bukanApproved = bukanApproved + 1
        Else
            target = TentukanTargetPK( _
                        data(r, idx("BUYER_TIN")), _
                        data(r, idx("TAX_INVOICE_CODE")), _
                        NilaiKolomOpsional(data, r, idx, "TABEL"))

            If Len(target) = 0 Then
                tidakTerpetakan = tidakTerpetakan + 1
            Else
                Set col = kumpulan(target)
                col.Add BuatRecordPK(target, data, r, idx)
            End If
        End If
    Next r
End Sub

Private Sub BacaSumberPM( _
    ByVal kumpulan As Object, _
    ByRef tidakTerpetakan As Long, _
    ByRef bukanApproved As Long)

    Dim lo As ListObject
    Dim idx As Object
    Dim data As Variant
    Dim r As Long
    Dim target As String
    Dim status As String
    Dim col As Collection

    Set lo = AmbilTabelWajib(SHEET_SUMBER_PM, TABEL_SUMBER_PM)
    Set idx = IndeksKolom(lo)
    ValidasiKolomSumber idx, False

    If lo.DataBodyRange Is Nothing Then Exit Sub
    data = lo.DataBodyRange.Value2

    For r = 1 To UBound(data, 1)
        status = NormalisasiTeks(data(r, idx("TAX_INVOICE_STATUS")))
        If status <> "APPROVED" Then
            bukanApproved = bukanApproved + 1
        Else
            target = TentukanTargetPM( _
                        data(r, idx("SELLER_TIN")), _
                        data(r, idx("BUYER_STATUS")), _
                        data(r, idx("TAX_INVOICE_CODE")))

            If Len(target) = 0 Then
                tidakTerpetakan = tidakTerpetakan + 1
            Else
                Set col = kumpulan(target)
                col.Add BuatRecordPM(target, data, r, idx)
            End If
        End If
    Next r
End Sub

Private Function TentukanTargetPK( _
    ByVal buyerTIN As Variant, _
    ByVal kodeFaktur As Variant, _
    ByVal tabelDokumen As Variant) As String

    Dim kode As String

    'Kode faktur adalah kunci utama mapping.
    kode = NormalisasiKodeFaktur(kodeFaktur)
    Select Case kode
        Case "04", "05", "TD00304", "TD00305": TentukanTargetPK = TARGET_N2
        Case "TD00306": TentukanTargetPK = TARGET_N3
        Case "TD00301", "TD00309", "TD00310": TentukanTargetPK = TARGET_N4
        Case "TD00302", "TD00303": TentukanTargetPK = TARGET_N6
        Case "TD00307": TentukanTargetPK = TARGET_N7
        Case "TD00308": TentukanTargetPK = TARGET_N8
        Case Else
            'Lanjut ke rule fallback di bawah.
    End Select

    If Len(TentukanTargetPK) > 0 Then Exit Function

    'Jika kode kosong, DOKUMEN LAIN tetap diarahkan ke N-2-2.
    If AdalahDokumenLain(tabelDokumen) Then
        TentukanTargetPK = TARGET_N2
        Exit Function
    End If

    If TINSeluruhnyaNol(buyerTIN) Then
        TentukanTargetPK = TARGET_N1
        Exit Function
    End If
End Function

Private Function AdalahDokumenLain(ByVal tabelDokumen As Variant) As Boolean
    Dim s As String

    s = NormalisasiTeks(tabelDokumen)
    AdalahDokumenLain = (InStr(1, s, "DOKUMEN LAIN", vbTextCompare) > 0)
End Function

Private Function NilaiKolomOpsional( _
    ByRef data As Variant, _
    ByVal baris As Long, _
    ByVal idx As Object, _
    ByVal namaKolom As String) As Variant

    If idx.Exists(namaKolom) Then
        NilaiKolomOpsional = data(baris, idx(namaKolom))
    Else
        NilaiKolomOpsional = vbNullString
    End If
End Function

Private Function BangunRingkasanMasa( _
    ByVal target As String, _
    ByVal records As Collection) As Object

    Dim hasil As Object
    Dim record As Variant
    Dim nilai As Variant
    Dim jumlah As Variant
    Dim groupKey As String
    Dim kolomNilai As Variant
    Dim i As Long
    Dim j As Long

    Set hasil = CreateObject("Scripting.Dictionary")
    hasil.CompareMode = vbTextCompare
    kolomNilai = KolomNilaiTarget(target)

    For i = 1 To records.Count
        record = records(i)
        groupKey = CStr(record(1))
        nilai = record(2)

        If hasil.Exists(groupKey) Then
            jumlah = hasil(groupKey)
        Else
            ReDim jumlah(1 To UBound(kolomNilai) - LBound(kolomNilai) + 1)
        End If

        For j = LBound(kolomNilai) To UBound(kolomNilai)
            If IsNumeric(nilai(CLng(kolomNilai(j)))) Then
                jumlah(j - LBound(kolomNilai) + 1) = _
                    jumlah(j - LBound(kolomNilai) + 1) + _
                    CDbl(nilai(CLng(kolomNilai(j))))
            End If
        Next j

        hasil(groupKey) = jumlah
    Next i

    Set BangunRingkasanMasa = hasil
End Function

Private Function TentukanTargetPM( _
    ByVal sellerTIN As Variant, _
    ByVal buyerStatus As Variant, _
    ByVal kodeFaktur As Variant) As String

    Dim status As String
    Dim kode As String

    status = NormalisasiTeks(buyerStatus)
    kode = NormalisasiKodeFaktur(kodeFaktur)

    'Kode faktur adalah kunci utama. Status dan TIN menjadi fallback.
    If status = "CREDITED" Then
        Select Case kode
            Case "04", "05", "TD00304", "TD00305": TentukanTargetPM = TARGET_N20
            Case "TD00301", "TD00309", "TD00310": TentukanTargetPM = TARGET_N21
            Case "TD00302", "TD00303": TentukanTargetPM = TARGET_N22
            ' TD00306, TD00307, dan TD00308 belum mempunyai target PM.
        End Select

        If Len(TentukanTargetPM) > 0 Then Exit Function

        If TINSeluruhnyaNol(sellerTIN) Then
            TentukanTargetPM = TARGET_N19
        End If
    ElseIf status = "UNCREDITED" Then
        TentukanTargetPM = TARGET_N24
    End If
End Function

Private Function BuatRecordPK( _
    ByVal target As String, _
    ByRef data As Variant, _
    ByVal r As Long, _
    ByVal idx As Object) As Variant

    Dim nilai() As Variant
    Dim nomor As String
    Dim masa As Long
    Dim tahun As Long
    Dim groupKey As String
    Dim sortKey As String

    nomor = CStr(data(r, idx("TAX_INVOICE_NUMBER")))
    masa = AngkaBulat(data(r, idx("TAX_INVOICE_PERIOD")))
    tahun = AngkaBulat(data(r, idx("TAX_INVOICE_YEAR")))
    groupKey = Format$(tahun, "0000") & Format$(masa, "00")
    sortKey = groupKey & "|" & KunciTanggal(data(r, idx("TAX_INVOICE_DATE"))) & _
              "|" & nomor

    If target = TARGET_N1 Then
        ReDim nilai(1 To 8)
        nilai(2) = data(r, idx("BUYER_TAXPAYER_NAME"))
        nilai(3) = masa
        nilai(4) = nomor
        nilai(5) = NilaiTanggal(data(r, idx("TAX_INVOICE_DATE")))
        nilai(6) = data(r, idx("SELLING_PRICE"))
        nilai(7) = data(r, IndeksReferencePK(idx))
        nilai(8) = data(r, idx("DETAIL_BARANG_JASA"))
    Else
        ReDim nilai(1 To TargetJumlahKolom(target))
        nilai(2) = data(r, idx("BUYER_TAXPAYER_NAME"))
        nilai(3) = CStr(data(r, idx("BUYER_TIN")))
        nilai(4) = masa
        If AdalahDokumenLain( _
                NilaiKolomOpsional(data, r, idx, "TABEL")) Then
            'Nomor dokumen lain tidak boleh dipaksa menjadi
            'Transaction Code/Type hanya berdasarkan karakter awal nomor.
            nilai(5) = vbNullString
            nilai(6) = vbNullString
        Else
            nilai(5) = KodeTransaksi(nomor)
            nilai(6) = TipePembetulan(nomor)
        End If
        nilai(7) = nomor
        nilai(8) = NilaiTanggal(data(r, idx("TAX_INVOICE_DATE")))
        nilai(9) = data(r, idx("SELLING_PRICE"))

        If TargetMemilikiDPPNilaiLain(target) Then
            nilai(10) = data(r, idx("OTHER_TAX_BASE"))
            nilai(11) = data(r, idx("VAT"))
            nilai(12) = vbNullString
            nilai(13) = data(r, idx("DETAIL_BARANG_JASA"))
        Else
            nilai(10) = data(r, idx("VAT"))
            nilai(11) = vbNullString
            nilai(12) = data(r, idx("DETAIL_BARANG_JASA"))
        End If
    End If

    BuatRecordPK = Array(sortKey, groupKey, nilai)
End Function

Private Function IndeksReferencePK(ByVal idx As Object) As Long
    ' Header pada file sumber saat ini memakai typo REFERNCE.
    ' REFERENCE tetap diterima untuk kompatibilitas jika header diperbaiki.
    If idx.Exists("REFERNCE") Then
        IndeksReferencePK = CLng(idx("REFERNCE"))
    ElseIf idx.Exists("REFERENCE") Then
        IndeksReferencePK = CLng(idx("REFERENCE"))
    Else
        Err.Raise vbObjectError + 510, , _
                  "Kolom sumber 'REFERNCE' tidak ditemukan pada FakturPK."
    End If
End Function

Private Function BuatRecordPM( _
    ByVal target As String, _
    ByRef data As Variant, _
    ByVal r As Long, _
    ByVal idx As Object) As Variant

    Dim nilai() As Variant
    Dim nomor As String
    Dim masa As Long
    Dim tahun As Long
    Dim groupKey As String
    Dim sortKey As String

    nomor = CStr(data(r, idx("TAX_INVOICE_NUMBER")))
    masa = AngkaBulat(data(r, idx("TAX_INVOICE_PERIOD")))
    tahun = AngkaBulat(data(r, idx("TAX_INVOICE_YEAR")))
    groupKey = Format$(tahun, "0000") & Format$(masa, "00")
    sortKey = groupKey & "|" & KunciTanggal(data(r, idx("TAX_INVOICE_DATE"))) & _
              "|" & nomor

    If target = TARGET_N19 Then
        ReDim nilai(1 To 9)
        nilai(2) = data(r, idx("SELLER_TAXPAYER_NAME"))
        nilai(3) = vbNullString
        nilai(4) = masa
        nilai(5) = nomor
        nilai(6) = NilaiTanggal(data(r, idx("TAX_INVOICE_DATE")))
        nilai(7) = data(r, idx("SELLING_PRICE"))
        nilai(8) = data(r, idx("VAT"))
        nilai(9) = data(r, idx("DETAIL_BARANG_JASA"))
    Else
        ReDim nilai(1 To TargetJumlahKolom(target))
        nilai(2) = data(r, idx("SELLER_TAXPAYER_NAME"))
        nilai(3) = CStr(data(r, idx("SELLER_TIN")))
        nilai(4) = masa
        nilai(5) = KodeTransaksi(nomor)
        nilai(6) = TipePembetulan(nomor)
        nilai(7) = nomor
        nilai(8) = NilaiTanggal(data(r, idx("TAX_INVOICE_DATE")))
        nilai(9) = data(r, idx("SELLING_PRICE"))

        If TargetMemilikiDPPNilaiLain(target) Then
            nilai(10) = data(r, idx("OTHER_TAX_BASE"))
            nilai(11) = data(r, idx("VAT"))
            nilai(12) = vbNullString
            nilai(13) = data(r, idx("DETAIL_BARANG_JASA"))
        Else
            nilai(10) = data(r, idx("VAT"))
            nilai(11) = vbNullString
            nilai(12) = data(r, idx("DETAIL_BARANG_JASA"))
        End If
    End If

    BuatRecordPM = Array(sortKey, groupKey, nilai)
End Function

Private Function SusunDataDenganPemisah( _
    ByVal target As String, _
    ByVal records As Collection) As Variant

    Dim arr() As Variant
    Dim hasil() As Variant
    Dim i As Long
    Dim c As Long
    Dim baris As Long
    Dim jumlahPemisah As Long
    Dim jumlahKolom As Long
    Dim record As Variant
    Dim nilai As Variant
    Dim groupSebelumnya As String

    jumlahKolom = TargetJumlahKolom(target)

    If records.Count = 0 Then
        SusunDataDenganPemisah = Empty
        Exit Function
    End If

    ReDim arr(1 To records.Count)
    For i = 1 To records.Count
        arr(i) = records(i)
    Next i
    If records.Count > 1 Then QuickSortRecord arr, 1, records.Count

    For i = 2 To UBound(arr)
        If CStr(arr(i)(1)) <> CStr(arr(i - 1)(1)) Then
            jumlahPemisah = jumlahPemisah + 1
        End If
    Next i

    ReDim hasil(1 To records.Count + jumlahPemisah, 1 To jumlahKolom)

    For i = 1 To UBound(arr)
        record = arr(i)
        If i > 1 And CStr(record(1)) <> groupSebelumnya Then
            baris = baris + 1
        End If

        baris = baris + 1
        nilai = record(2)
        nilai(1) = i
        For c = 1 To jumlahKolom
            hasil(baris, c) = nilai(c)
        Next c
        groupSebelumnya = CStr(record(1))
    Next i

    SusunDataDenganPemisah = hasil
End Function

Private Sub QuickSortRecord( _
    ByRef arr() As Variant, _
    ByVal first As Long, _
    ByVal last As Long)

    Dim i As Long
    Dim j As Long
    Dim pivot As String
    Dim temp As Variant

    i = first
    j = last
    pivot = CStr(arr((first + last) \ 2)(0))

    Do While i <= j
        Do While CStr(arr(i)(0)) < pivot
            i = i + 1
        Loop
        Do While CStr(arr(j)(0)) > pivot
            j = j - 1
        Loop

        If i <= j Then
            temp = arr(i)
            arr(i) = arr(j)
            arr(j) = temp
            i = i + 1
            j = j - 1
        End If
    Loop

    If first < j Then QuickSortRecord arr, first, j
    If i < last Then QuickSortRecord arr, i, last
End Sub

Private Sub TulisDataTarget(ByVal target As String, ByVal data As Variant)
    Dim ws As Worksheet
    Dim dataStart As Long
    Dim totalRow As Long
    Dim requiredRows As Long
    Dim dataEnd As Long
    Dim jumlahKolom As Long
    Dim oldLastRow As Long
    Dim clearEndRow As Long
    Dim templateTotalRow As Long
    Dim areaContoh As Range
    Dim areaTujuan As Range

    Set ws = ThisWorkbook.Worksheets(target)
    dataStart = TargetBarisAwal(target)
    jumlahKolom = TargetJumlahKolom(target)

    If IsEmpty(data) Then
        requiredRows = 0
    Else
        requiredRows = UBound(data, 1)
    End If

    ' Posisi selalu dibangun dari jumlah data baru:
    ' data, satu baris kosong, lalu baris total.
    totalRow = dataStart + requiredRows + 1
    oldLastRow = BarisTerakhirIsi(ws)
    If oldLastRow < dataStart Then oldLastRow = dataStart
    clearEndRow = oldLastRow
    If totalRow > clearEndRow Then clearEndRow = totalRow

    ' Pertahankan format baris data dari template. Bila tersedia, gunakan
    ' format baris total lama; jika tidak ada, format dasar akan diperkuat.
    Set areaContoh = ws.Range(ws.Cells(dataStart, 1), _
                     ws.Cells(dataStart, jumlahKolom))
    Set areaTujuan = ws.Range(ws.Cells(dataStart, 1), _
                     ws.Cells(totalRow, jumlahKolom))
    areaContoh.Copy
    areaTujuan.PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False

    templateTotalRow = CariBarisTotal(ws, dataStart)
    If templateTotalRow = 0 And oldLastRow >= dataStart Then
        templateTotalRow = oldLastRow
    End If

    If templateTotalRow >= dataStart Then
        ws.Range(ws.Cells(templateTotalRow, 1), _
                 ws.Cells(templateTotalRow, jumlahKolom)).Copy
        ws.Range(ws.Cells(totalRow, 1), _
                 ws.Cells(totalRow, jumlahKolom)).PasteSpecial _
                 Paste:=xlPasteFormats
        Application.CutCopyMode = False
    End If

    ' Hapus seluruh isi lama pada area detail sebelum menulis hasil baru.
    ws.Range(ws.Cells(dataStart, 1), _
             ws.Cells(clearEndRow, jumlahKolom)).ClearContents

    If requiredRows > 0 Then
        dataEnd = dataStart + requiredRows - 1
        TerapkanFormatTarget ws, target, dataStart, dataEnd
        ws.Cells(dataStart, 1).Resize(requiredRows, jumlahKolom).Value2 = data
    End If

    PerbaruiFormulaTotal ws, target, dataStart, totalRow
    FormatBarisTotal ws, totalRow, jumlahKolom
End Sub

Private Sub FormatBarisTotal( _
    ByVal ws As Worksheet, _
    ByVal totalRow As Long, _
    ByVal jumlahKolom As Long)

    Dim areaTotal As Range

    Set areaTotal = ws.Range(ws.Cells(totalRow, 1), _
                    ws.Cells(totalRow, jumlahKolom))
    With areaTotal
        .Font.Bold = True
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlDouble
        End With
    End With
End Sub

Private Sub PerbaruiSheetInduk(ByVal targetDetail As String)
    Dim ws As Worksheet
    Dim namaInduk As String
    Dim headerRow As Long
    Dim kolomUraian As Long
    Dim kolomRef As Long
    Dim kolomPemeriksa As Long
    Dim jumlahMetrik As Long
    Dim lastRow As Long
    Dim r As Long
    Dim i As Long
    Dim masa As Long
    Dim masaDiharapkan As Long
    Dim alamatDetail As String
    Dim teksRef As String
    Dim areaRef As Range
    Dim areaNilai As Range

    namaInduk = NamaSheetInduk(targetDetail)
    Set ws = ThisWorkbook.Worksheets(namaInduk)

    kolomUraian = CariKolomBerisi(ws, "URAIAN", 1, 30)
    kolomRef = CariKolomBerisi(ws, "REF", 1, 30)
    kolomPemeriksa = CariKolomBerisi(ws, "MENURUT PEMERIKSA", 1, 30)
    headerRow = CariBarisBerisi(ws, "URAIAN", 1, 30)

    If headerRow = 0 Or kolomUraian = 0 Or kolomRef = 0 Or _
       kolomPemeriksa = 0 Then
        Err.Raise vbObjectError + 511, , _
                  "Struktur kolom pada sheet induk '" & namaInduk & _
                  "' tidak dapat dikenali."
    End If

    jumlahMetrik = JumlahMetrikTarget(targetDetail)
    lastRow = BarisTerakhirIsi(ws)
    alamatDetail = "'" & Replace$(targetDetail, "'", "''") & "'!A1"
    teksRef = TeksReferensiTarget(targetDetail)
    masaDiharapkan = 1

    BersihkanSpillFormulaInduk ws, targetDetail, headerRow, lastRow, _
                               kolomUraian, kolomRef, kolomPemeriksa, _
                               jumlahMetrik, teksRef

    ' Hanya blok masa 1 sampai 12 yang boleh diperbarui. Angka pada tabel
    ' Penjelasan/Dasar Hukum di bagian bawah tidak boleh dianggap sebagai masa.
    For r = headerRow + 1 To lastRow
        masa = NomorMasaDariBarisInduk(ws, r, kolomUraian, kolomRef)
        If masa = masaDiharapkan Then
            Set areaRef = AreaSelTulis(ws.Cells(r, kolomRef))
            On Error Resume Next
            areaRef.Hyperlinks.Delete
            On Error GoTo 0
            areaRef.ClearContents
            ws.Hyperlinks.Add Anchor:=areaRef, _
                              Address:=vbNullString, _
                              SubAddress:=alamatDetail, _
                              TextToDisplay:=teksRef
            TerapkanFormatRef areaRef

            For i = 1 To jumlahMetrik
                Set areaNilai = AreaSelTulis( _
                    ws.Cells(r, kolomPemeriksa + i - 1))
                areaNilai.Formula = FormulaSubtotalMasa( _
                                        targetDetail, masa, i)
            Next i

            masaDiharapkan = masaDiharapkan + 1
            If masaDiharapkan > 12 Then Exit For
        End If
    Next r
End Sub

Private Sub BersihkanSpillFormulaInduk( _
    ByVal ws As Worksheet, _
    ByVal targetDetail As String, _
    ByVal headerRow As Long, _
    ByVal lastRow As Long, _
    ByVal kolomUraian As Long, _
    ByVal kolomRef As Long, _
    ByVal kolomPemeriksa As Long, _
    ByVal jumlahMetrik As Long, _
    ByVal teksRef As String)

    Dim r As Long
    Dim i As Long
    Dim areaRef As Range
    Dim areaNilai As Range
    Dim nilaiRef As Variant
    Dim rumus As String

    For r = headerRow + 1 To lastRow
        If NomorMasaDariBarisInduk( _
            ws, r, kolomUraian, kolomRef) = 0 Then

            Set areaRef = AreaSelTulis(ws.Cells(r, kolomRef))
            nilaiRef = areaRef.Cells(1, 1).Value2
            If Not IsError(nilaiRef) Then
                If NormalisasiTeks(CStr(nilaiRef)) = _
                   NormalisasiTeks(teksRef) Then
                    On Error Resume Next
                    areaRef.Hyperlinks.Delete
                    On Error GoTo 0
                    areaRef.ClearContents
                End If
            End If

            For i = 1 To jumlahMetrik
                Set areaNilai = AreaSelTulis( _
                    ws.Cells(r, kolomPemeriksa + i - 1))
                If areaNilai.Cells(1, 1).HasFormula Then
                    rumus = CStr(areaNilai.Cells(1, 1).Formula)
                    If InStr(1, rumus, targetDetail, vbTextCompare) > 0 Then
                        areaNilai.ClearContents
                    End If
                End If
            Next i
        End If
    Next r
End Sub

Private Sub TerapkanFormatRef(ByVal areaRef As Range)
    ' Hyperlink bawaan Excel menerapkan underline. Setelah tautan dibuat,
    ' kembalikan tampilan Ref. sesuai format formulir induk.
    With areaRef
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Underline = xlUnderlineStyleNone

        With .Borders(xlEdgeLeft)
            If .LineStyle = xlLineStyleNone Then
                .LineStyle = xlContinuous
                .Weight = xlThin
            End If
        End With

        With .Borders(xlEdgeRight)
            If .LineStyle = xlLineStyleNone Then
                .LineStyle = xlContinuous
                .Weight = xlThin
            End If
        End With
    End With
End Sub

Private Function FormulaSubtotalMasa( _
    ByVal targetDetail As String, _
    ByVal masa As Long, _
    ByVal indeksMetrik As Long) As String

    Dim wsDetail As Worksheet
    Dim dataStart As Long
    Dim dataEnd As Long
    Dim totalRow As Long
    Dim kolomMasa As String
    Dim kolomNilai As String
    Dim daftarKolom As Variant
    Dim namaSheet As String

    Set wsDetail = ThisWorkbook.Worksheets(targetDetail)
    dataStart = TargetBarisAwal(targetDetail)
    totalRow = CariBarisTotal(wsDetail, dataStart)
    dataEnd = totalRow - 2
    If dataEnd < dataStart Then dataEnd = dataStart

    If targetDetail = TARGET_N1 Then
        kolomMasa = "C"
    Else
        kolomMasa = "D"
    End If

    daftarKolom = KolomTotal(targetDetail)
    kolomNilai = CStr(daftarKolom( _
                    LBound(daftarKolom) + indeksMetrik - 1))
    namaSheet = Replace$(targetDetail, "'", "''")

    FormulaSubtotalMasa = "=SUMIF('" & namaSheet & "'!$" & _
        kolomMasa & "$" & CStr(dataStart) & ":$" & kolomMasa & "$" & _
        CStr(dataEnd) & "," & CStr(masa) & ",'" & namaSheet & "'!$" & _
        kolomNilai & "$" & CStr(dataStart) & ":$" & kolomNilai & "$" & _
        CStr(dataEnd) & ")"
End Function

Private Sub PerbaruiCFMPemeriksaN1()
    Dim wsDetail As Worksheet
    Dim wsInduk As Worksheet
    Dim namaInduk As String
    Dim dataStart As Long
    Dim dataEnd As Long
    Dim totalDetail As Long
    Dim headerRow As Long
    Dim kolomUraian As Long
    Dim kolomRef As Long
    Dim kolomPemeriksa As Long
    Dim lastRow As Long
    Dim r As Long
    Dim masa As Long
    Dim uraian As String
    Dim namaDetail As String
    Dim areaNilai As Range

    Set wsDetail = ThisWorkbook.Worksheets(TARGET_N1)
    namaInduk = NamaSheetInduk(TARGET_N1)
    Set wsInduk = ThisWorkbook.Worksheets(namaInduk)
    dataStart = TargetBarisAwal(TARGET_N1)
    totalDetail = CariBarisTotal(wsDetail, dataStart)
    dataEnd = totalDetail - 2
    If dataEnd < dataStart Then dataEnd = dataStart

    CariHeaderN1 wsInduk, headerRow, kolomUraian, kolomRef, kolomPemeriksa
    lastRow = BarisTerakhirIsi(wsInduk)
    namaDetail = Replace$(TARGET_N1, "'", "''")

    For r = headerRow + 1 To lastRow
        uraian = Trim$(CStr(wsInduk.Cells(r, kolomUraian).Value2))
        If Len(uraian) > 0 Then
            If InStr(1, UCase$(uraian), "TOTAL", vbTextCompare) > 0 Or _
               InStr(1, UCase$(uraian), "JUMLAH", vbTextCompare) > 0 Then
                Set areaNilai = AreaSelTulis( _
                                    wsInduk.Cells(r, kolomPemeriksa))
                areaNilai.Formula = "='" & namaDetail & "'!F" & _
                                    CStr(totalDetail)
                areaNilai.NumberFormat = "#,##0"
                Exit For
            Else
                masa = NomorMasaDariBarisInduk( _
                            wsInduk, r, kolomUraian, kolomRef)
                If masa > 0 Then
                    Set areaNilai = AreaSelTulis( _
                                        wsInduk.Cells(r, kolomPemeriksa))
                    areaNilai.Formula = "=SUMIF('" & namaDetail & "'!$C$" & _
                        CStr(dataStart) & ":$C$" & CStr(dataEnd) & "," & _
                        CStr(masa) & ",'" & namaDetail & "'!$F$" & _
                        CStr(dataStart) & ":$F$" & CStr(dataEnd) & ")"
                    areaNilai.NumberFormat = "#,##0"
                End If
            End If
        End If
    Next r
End Sub

Private Sub CariHeaderN1( _
    ByVal ws As Worksheet, _
    ByRef headerRow As Long, _
    ByRef kolomUraian As Long, _
    ByRef kolomRef As Long, _
    ByRef kolomPemeriksa As Long)

    Dim selUraian As Range
    Dim selRef As Range
    Dim selPemeriksa As Range
    Dim lastCol As Long

    lastCol = KolomTerakhirIsi(ws)
    Set selUraian = ws.Range(ws.Cells(1, 1), ws.Cells(40, lastCol)).Find( _
        What:="URAIAN", LookIn:=xlValues, LookAt:=xlPart, _
        SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)

    If selUraian Is Nothing Then
        Err.Raise vbObjectError + 519, , _
                  "Header Uraian pada sheet '" & ws.Name & "' tidak ditemukan."
    End If

    headerRow = selUraian.Row
    kolomUraian = selUraian.Column
    Set selRef = ws.Range(ws.Cells(headerRow, 1), _
                 ws.Cells(headerRow + 2, lastCol)).Find( _
                 What:="REF", LookIn:=xlValues, LookAt:=xlPart, _
                 SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                 MatchCase:=False)

    If selRef Is Nothing Then
        Err.Raise vbObjectError + 521, , _
                  "Header Ref. pada sheet '" & ws.Name & "' tidak ditemukan."
    End If

    kolomRef = selRef.Column
    Set selPemeriksa = ws.Range(ws.Cells(headerRow, 1), _
                        ws.Cells(headerRow + 2, lastCol)).Find( _
                        What:="PEMERIKSA", LookIn:=xlValues, LookAt:=xlPart, _
                        SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                        MatchCase:=False)

    If selPemeriksa Is Nothing Then
        Err.Raise vbObjectError + 520, , _
                  "Header Pemeriksa pada sheet '" & ws.Name & "' tidak ditemukan."
    End If

    kolomPemeriksa = selPemeriksa.Column
End Sub

Private Function AreaSelTulis(ByVal sel As Range) As Range
    If sel.MergeCells Then
        Set AreaSelTulis = sel.MergeArea
    Else
        Set AreaSelTulis = sel
    End If
End Function

Private Function NamaSheetInduk(ByVal targetDetail As String) As String
    Dim dasar As String
    Dim kandidat As Variant
    Dim nama As Variant

    dasar = Left$(targetDetail, Len(targetDetail) - 2)
    kandidat = Array( _
        dasar, _
        Replace$(dasar, "PPN-", vbNullString, 1, 1, vbTextCompare), _
        Replace$(dasar, "PPN-N-", "PPN-N.", 1, 1, vbTextCompare), _
        Replace$(dasar, "PPN-N-", "N.", 1, 1, vbTextCompare))

    For Each nama In kandidat
        If SheetAda(CStr(nama)) Then
            NamaSheetInduk = CStr(nama)
            Exit Function
        End If
    Next nama

    Err.Raise vbObjectError + 512, , _
              "Sheet induk untuk '" & targetDetail & "' tidak ditemukan. " & _
              "Nama yang diharapkan antara lain '" & dasar & "'."
End Function

Private Function CariKolomBerisi( _
    ByVal ws As Worksheet, _
    ByVal teks As String, _
    ByVal barisAwal As Long, _
    ByVal barisAkhir As Long) As Long

    Dim ditemukan As Range

    Set ditemukan = ws.Range(ws.Cells(barisAwal, 1), _
                      ws.Cells(barisAkhir, KolomTerakhirIsi(ws))).Find( _
                      What:=teks, LookIn:=xlValues, LookAt:=xlPart, _
                      SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                      MatchCase:=False)
    If Not ditemukan Is Nothing Then CariKolomBerisi = ditemukan.Column
End Function

Private Function CariBarisBerisi( _
    ByVal ws As Worksheet, _
    ByVal teks As String, _
    ByVal barisAwal As Long, _
    ByVal barisAkhir As Long) As Long

    Dim ditemukan As Range

    Set ditemukan = ws.Range(ws.Cells(barisAwal, 1), _
                      ws.Cells(barisAkhir, KolomTerakhirIsi(ws))).Find( _
                      What:=teks, LookIn:=xlValues, LookAt:=xlPart, _
                      SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                      MatchCase:=False)
    If Not ditemukan Is Nothing Then CariBarisBerisi = ditemukan.Row
End Function

Private Function NomorBulanDariTeks(ByVal uraian As String) As Long
    Dim namaBulan As Variant
    Dim teks As String
    Dim i As Long

    namaBulan = Array( _
        "JANUARI", "FEBRUARI", "MARET", "APRIL", "MEI", "JUNI", _
        "JULI", "AGUSTUS", "SEPTEMBER", "OKTOBER", "NOVEMBER", "DESEMBER")
    teks = NormalisasiTeks(uraian)

    For i = LBound(namaBulan) To UBound(namaBulan)
        If InStr(1, teks, CStr(namaBulan(i)), vbTextCompare) > 0 Then
            NomorBulanDariTeks = i + 1
            Exit Function
        End If
    Next i
End Function

Private Function TahunDariTeks(ByVal nilai As String) As Long
    Dim teks As String
    Dim i As Long
    Dim potongan As String

    teks = NormalisasiTeks(nilai)
    For i = 1 To Len(teks) - 3
        potongan = Mid$(teks, i, 4)
        If IsNumeric(potongan) Then
            TahunDariTeks = CLng(potongan)
            If TahunDariTeks >= 1900 And TahunDariTeks <= 2200 Then _
                Exit Function
            TahunDariTeks = 0
        End If
    Next i
End Function

Private Function NomorMasaDariBarisInduk( _
    ByVal ws As Worksheet, _
    ByVal baris As Long, _
    ByVal kolomUraian As Long, _
    ByVal kolomRef As Long) As Long

    Dim nilai As Variant
    Dim teks As String
    Dim masa As Long
    Dim tahun As Long

    ' Kunci baris pada sheet N-X harus berupa Bulan Tahun di kolom B,
    ' misalnya Januari 2026. Kolom A tidak pernah dipakai sebagai kunci.
    nilai = ws.Cells(baris, 2).Value

    If Not IsError(nilai) Then
        If Not IsNull(nilai) And Not IsEmpty(nilai) Then
            If VarType(nilai) = vbDate Then
                masa = Month(CDate(nilai))
                tahun = Year(CDate(nilai))
                If masa >= 1 And masa <= 12 And _
                   tahun >= 1900 And tahun <= 2200 Then
                    NomorMasaDariBarisInduk = masa
                    Exit Function
                End If
            End If

            teks = CStr(nilai)
            masa = NomorBulanDariTeks(teks)
            tahun = TahunDariTeks(teks)
            If masa > 0 And tahun > 0 Then
                NomorMasaDariBarisInduk = masa
                Exit Function
            End If
        End If
    End If
End Function

Private Function TahunDariBarisInduk( _
    ByVal ws As Worksheet, _
    ByVal baris As Long, _
    ByVal kolomUraian As Long, _
    ByVal kolomRef As Long) As Long

    Dim c As Long
    Dim teks As String
    Dim i As Long
    Dim potongan As String

    For c = kolomUraian To kolomRef - 1
        If Not IsError(ws.Cells(baris, c).Value2) Then
            If Not IsNull(ws.Cells(baris, c).Value2) And _
               Not IsEmpty(ws.Cells(baris, c).Value2) Then
                teks = teks & " " & CStr(ws.Cells(baris, c).Value2)
            End If
        End If
    Next c
    teks = NormalisasiTeks(teks)

    For i = 1 To Len(teks) - 3
        potongan = Mid$(teks, i, 4)
        If IsNumeric(potongan) Then
            TahunDariBarisInduk = CLng(potongan)
            If TahunDariBarisInduk >= 1900 And _
               TahunDariBarisInduk <= 2200 Then Exit Function
            TahunDariBarisInduk = 0
        End If
    Next i
End Function

Private Function KunciRingkasanUntukBarisInduk( _
    ByVal ws As Worksheet, _
    ByVal baris As Long, _
    ByVal kolomUraian As Long, _
    ByVal kolomRef As Long, _
    ByVal ringkasan As Object) As String

    Dim masa As Long
    Dim tahun As Long
    Dim key As Variant
    Dim keyDitemukan As String
    Dim jumlahCocok As Long

    masa = NomorMasaDariBarisInduk(ws, baris, kolomUraian, kolomRef)
    If masa = 0 Then Exit Function

    tahun = TahunDariBarisInduk(ws, baris, kolomUraian, kolomRef)
    If tahun > 0 Then
        KunciRingkasanUntukBarisInduk = _
            Format$(tahun, "0000") & Format$(masa, "00")
        Exit Function
    End If

    ' Jika label induk hanya memuat nomor masa, ambil tahun tunggal dari
    ' data sumber. Lebih dari satu tahun untuk masa yang sama bersifat ambigu.
    For Each key In ringkasan.Keys
        If Right$(CStr(key), 2) = Format$(masa, "00") Then
            jumlahCocok = jumlahCocok + 1
            keyDitemukan = CStr(key)
        End If
    Next key

    If jumlahCocok > 1 Then
        Err.Raise vbObjectError + 516, , _
                  "Masa " & CStr(masa) & " pada sheet induk '" & ws.Name & _
                  "' memuat lebih dari satu tahun pada data sumber."
    ElseIf jumlahCocok = 1 Then
        KunciRingkasanUntukBarisInduk = keyDitemukan
    End If
End Function

Private Function TeksReferensiTarget(ByVal targetDetail As String) As String
    Dim s As String

    s = Replace$(targetDetail, "PPN-", vbNullString, 1, 1, vbTextCompare)
    TeksReferensiTarget = Replace$(s, "-", ".")
End Function

Private Sub TerapkanFormatTarget( _
    ByVal ws As Worksheet, _
    ByVal target As String, _
    ByVal barisAwal As Long, _
    ByVal barisAkhir As Long)

    Dim rng As Range
    Dim area As Range

    If barisAkhir < barisAwal Then Exit Sub

    Set area = ws.Range(ws.Cells(barisAwal, 1), _
                        ws.Cells(barisAkhir, TargetJumlahKolom(target)))
    With area
        .NumberFormat = "@"
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = False
    End With

    If target = TARGET_N1 Then
        Set rng = ws.Range("F" & barisAwal & ":F" & barisAkhir)
        rng.NumberFormat = "#,##0;(#,##0)"
        rng.HorizontalAlignment = xlRight

        Set rng = ws.Range("E" & barisAwal & ":E" & barisAkhir)
        rng.NumberFormat = "dd/mm/yyyy"
        rng.HorizontalAlignment = xlCenter
    ElseIf target = TARGET_N19 Then
        Set rng = ws.Range("G" & barisAwal & ":H" & barisAkhir)
        rng.NumberFormat = "#,##0;(#,##0)"
        rng.HorizontalAlignment = xlRight

        Set rng = ws.Range("F" & barisAwal & ":F" & barisAkhir)
        rng.NumberFormat = "dd/mm/yyyy"
        rng.HorizontalAlignment = xlCenter
    Else
        If TargetMemilikiDPPNilaiLain(target) Then
            Set rng = ws.Range("I" & barisAwal & ":K" & barisAkhir)
        Else
            Set rng = ws.Range("I" & barisAwal & ":J" & barisAkhir)
        End If
        rng.NumberFormat = "#,##0;(#,##0)"
        rng.HorizontalAlignment = xlRight

        Set rng = ws.Range("H" & barisAwal & ":H" & barisAkhir)
        rng.NumberFormat = "dd/mm/yyyy"
        rng.HorizontalAlignment = xlCenter
    End If
End Sub

Private Sub PerbaruiFormulaTotal( _
    ByVal ws As Worksheet, _
    ByVal target As String, _
    ByVal dataStart As Long, _
    ByVal totalRow As Long)

    Dim dataEnd As Long
    Dim fungsi As String
    Dim kolom As Variant
    Dim daftarKolom As Variant

    dataEnd = totalRow - 2
    If dataEnd < dataStart Then dataEnd = dataStart

    If TargetAdalahPM(target) Then
        fungsi = "SUBTOTAL(9,"
    Else
        fungsi = "SUM("
    End If

    daftarKolom = KolomTotal(target)
    For Each kolom In daftarKolom
        ws.Range(CStr(kolom) & totalRow).Formula = _
            "=" & fungsi & CStr(kolom) & dataStart & ":" & _
            CStr(kolom) & dataEnd & ")"
        ws.Range(CStr(kolom) & totalRow).NumberFormat = "#,##0;(#,##0)"
    Next kolom
End Sub

Private Function RingkasanDataTargetLama() As String
    Dim namaTarget As Variant
    Dim ws As Worksheet
    Dim jumlah As Long
    Dim jumlahMasaInduk As Long
    Dim namaInduk As String
    Dim hasil As String

    For Each namaTarget In DaftarTarget()
        Set ws = ThisWorkbook.Worksheets(CStr(namaTarget))
        jumlah = HitungBarisDataTarget(ws, CStr(namaTarget))
        If jumlah > 0 Then
            hasil = hasil & "- " & CStr(namaTarget) & ": " & _
                    Format$(jumlah, "#,##0") & " baris" & vbCrLf
        End If

        If CStr(namaTarget) = TARGET_N1 Then
            namaInduk = NamaSheetInduk(CStr(namaTarget))
            Set ws = ThisWorkbook.Worksheets(namaInduk)
            jumlah = HitungBarisDataTarget(ws, TARGET_N1)
            If jumlah > 0 Then
                hasil = hasil & "- " & namaInduk & ": " & _
                        Format$(jumlah, "#,##0") & " baris" & vbCrLf
            End If
        ElseIf TargetMemilikiSheetInduk(CStr(namaTarget)) Then
            namaInduk = NamaSheetInduk(CStr(namaTarget))
            Set ws = ThisWorkbook.Worksheets(namaInduk)
            jumlahMasaInduk = HitungMasaIndukBerisi(ws, CStr(namaTarget))
            If jumlahMasaInduk > 0 Then
                hasil = hasil & "- " & namaInduk & ": " & _
                        Format$(jumlahMasaInduk, "#,##0") & _
                        " masa berisi nilai pemeriksa" & vbCrLf
            End If
        End If
    Next namaTarget

    RingkasanDataTargetLama = hasil
End Function

Private Function HitungMasaIndukBerisi( _
    ByVal ws As Worksheet, _
    ByVal targetDetail As String) As Long

    Dim kolomUraian As Long
    Dim kolomRef As Long
    Dim kolomPemeriksa As Long
    Dim jumlahMetrik As Long
    Dim lastRow As Long
    Dim r As Long
    Dim i As Long
    Dim nilai As Variant
    Dim adaNilai As Boolean

    kolomUraian = CariKolomBerisi(ws, "URAIAN", 1, 30)
    kolomRef = CariKolomBerisi(ws, "REF", 1, 30)
    kolomPemeriksa = CariKolomBerisi(ws, "MENURUT PEMERIKSA", 1, 30)
    jumlahMetrik = JumlahMetrikTarget(targetDetail)
    lastRow = BarisTerakhirIsi(ws)

    For r = 1 To lastRow
        If NomorMasaDariBarisInduk(ws, r, kolomUraian, kolomRef) > 0 Then
            adaNilai = False
            For i = 1 To jumlahMetrik
                nilai = ws.Cells(r, kolomPemeriksa + i - 1).Value2
                If IsError(nilai) Then
                    ' Error formula tetap dianggap sebagai isi lama agar
                    ' pengguna memperoleh konfirmasi sebelum ditimpa.
                    adaNilai = True
                    Exit For
                ElseIf IsNull(nilai) Or IsEmpty(nilai) Then
                    ' Sel kosong tidak dihitung sebagai data lama.
                ElseIf Len(CStr(nilai)) > 0 Then
                    If IsNumeric(nilai) Then
                        If CDbl(nilai) <> 0 Then
                            adaNilai = True
                            Exit For
                        End If
                    Else
                        adaNilai = True
                        Exit For
                    End If
                End If
            Next i
            If adaNilai Then HitungMasaIndukBerisi = HitungMasaIndukBerisi + 1
        End If
    Next r
End Function

Private Function HitungBarisDataTarget( _
    ByVal ws As Worksheet, _
    ByVal target As String) As Long

    Dim r As Long
    Dim totalRow As Long
    Dim dataStart As Long
    Dim lastCol As Long

    dataStart = TargetBarisAwal(target)
    lastCol = TargetJumlahKolom(target)
    totalRow = CariBarisTotal(ws, dataStart)

    If totalRow = 0 Then
        totalRow = BarisTerakhirIsi(ws) + 2
    End If

    For r = dataStart To totalRow - 2
        If Application.WorksheetFunction.CountA( _
            ws.Range(ws.Cells(r, 1), ws.Cells(r, lastCol))) > 0 Then
            HitungBarisDataTarget = HitungBarisDataTarget + 1
        End If
    Next r
End Function

Private Function RingkasanHasil(ByVal dataTarget As Object) As String
    Dim namaTarget As Variant
    Dim data As Variant
    Dim jumlah As Long
    Dim hasil As String

    For Each namaTarget In DaftarTarget()
        data = dataTarget(CStr(namaTarget))
        jumlah = HitungRecordDalamArray(data)
        hasil = hasil & CStr(namaTarget) & ": " & _
                Format$(jumlah, "#,##0") & " faktur" & vbCrLf
    Next namaTarget

    RingkasanHasil = hasil
End Function

Private Function HitungRecordDalamArray(ByVal data As Variant) As Long
    Dim r As Long

    If IsEmpty(data) Then Exit Function
    For r = 1 To UBound(data, 1)
        If Len(CStr(data(r, 1))) > 0 Then
            HitungRecordDalamArray = HitungRecordDalamArray + 1
        End If
    Next r
End Function

Private Sub ValidasiSeluruhStruktur()
    Dim namaTarget As Variant
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim idx As Object
    Dim namaInduk As String

    Set lo = AmbilTabelWajib(SHEET_SUMBER_PK, TABEL_SUMBER_PK)
    Set idx = IndeksKolom(lo)
    ValidasiKolomSumber idx, True

    Set lo = AmbilTabelWajib(SHEET_SUMBER_PM, TABEL_SUMBER_PM)
    Set idx = IndeksKolom(lo)
    ValidasiKolomSumber idx, False

    For Each namaTarget In DaftarTarget()
        If Not SheetAda(CStr(namaTarget)) Then
            Err.Raise vbObjectError + 501, , _
                      "Sheet target '" & CStr(namaTarget) & "' tidak ditemukan."
        End If

        Set ws = ThisWorkbook.Worksheets(CStr(namaTarget))
        If ws.ProtectContents Then
            Err.Raise vbObjectError + 502, , _
                      "Sheet target '" & CStr(namaTarget) & "' sedang diproteksi."
        End If

        ValidasiHeaderTarget ws, CStr(namaTarget)

        ' N-1 hanya menerima pembaruan CFM Pemeriksa. Target lain memakai
        ' sheet induk ringkasan dengan Ref. dan kolom Menurut Pemeriksa.
        If CStr(namaTarget) = TARGET_N1 Then
            namaInduk = NamaSheetInduk(CStr(namaTarget))
            Set ws = ThisWorkbook.Worksheets(namaInduk)
            If ws.ProtectContents Then
                Err.Raise vbObjectError + 517, , _
                          "Sheet induk '" & namaInduk & "' sedang diproteksi."
            End If
            ValidasiStrukturN1 ws
        ElseIf TargetMemilikiSheetInduk(CStr(namaTarget)) Then
            namaInduk = NamaSheetInduk(CStr(namaTarget))
            ValidasiStrukturInduk ThisWorkbook.Worksheets(namaInduk), _
                                  CStr(namaTarget)
        End If
    Next namaTarget
End Sub

Private Sub ValidasiStrukturN1(ByVal ws As Worksheet)
    Dim headerRow As Long
    Dim kolomUraian As Long
    Dim kolomRef As Long
    Dim kolomPemeriksa As Long
    Dim lastRow As Long
    Dim r As Long
    Dim jumlahBarisMasa As Long

    CariHeaderN1 ws, headerRow, kolomUraian, kolomRef, kolomPemeriksa
    lastRow = BarisTerakhirIsi(ws)

    For r = headerRow + 1 To lastRow
        If NomorMasaDariBarisInduk( _
            ws, r, kolomUraian, kolomRef) > 0 Then
            jumlahBarisMasa = jumlahBarisMasa + 1
        End If
    Next r

    If jumlahBarisMasa = 0 Then
        Err.Raise vbObjectError + 522, , _
                  "Baris masa pajak 1 sampai 12 pada sheet '" & ws.Name & _
                  "' tidak ditemukan. Pulihkan sheet tersebut dari template " & _
                  "sebelum menjalankan mapping."
    End If
End Sub

Private Function TargetMemilikiSheetInduk(ByVal target As String) As Boolean
    ' PPN-N-1 memakai mekanisme khusus: hanya CFM Pemeriksa yang diperbarui.
    TargetMemilikiSheetInduk = (target <> TARGET_N1)
End Function

Private Sub ValidasiStrukturInduk( _
    ByVal ws As Worksheet, _
    ByVal targetDetail As String)

    Dim kolomUraian As Long
    Dim kolomRef As Long
    Dim kolomPemeriksa As Long
    Dim headerRow As Long
    Dim lastRow As Long
    Dim r As Long
    Dim masaDiharapkan As Long

    If ws.ProtectContents Then
        Err.Raise vbObjectError + 515, , _
                  "Sheet induk '" & ws.Name & "' sedang diproteksi."
    End If

    kolomUraian = CariKolomBerisi(ws, "URAIAN", 1, 30)
    kolomRef = CariKolomBerisi(ws, "REF", 1, 30)
    kolomPemeriksa = CariKolomBerisi(ws, "MENURUT PEMERIKSA", 1, 30)
    headerRow = CariBarisBerisi(ws, "URAIAN", 1, 30)

    If headerRow = 0 Or kolomUraian = 0 Or kolomRef = 0 Or _
       kolomPemeriksa = 0 Then
        Err.Raise vbObjectError + 513, , _
                  "Struktur sheet induk '" & ws.Name & _
                  "' tidak memiliki kolom Uraian, Ref., atau Menurut Pemeriksa."
    End If

    lastRow = BarisTerakhirIsi(ws)
    masaDiharapkan = 1
    For r = headerRow + 1 To lastRow
        If NomorMasaDariBarisInduk( _
            ws, r, kolomUraian, kolomRef) = masaDiharapkan Then
            masaDiharapkan = masaDiharapkan + 1
            If masaDiharapkan > 12 Then Exit For
        End If
    Next r

    If masaDiharapkan <= 12 Then
        Err.Raise vbObjectError + 514, , _
                  "Blok masa 1 sampai 12 pada kolom B sheet induk '" & _
                  ws.Name & "' tidak ditemukan secara lengkap dan berurutan."
    End If
End Sub

Private Sub ValidasiHeaderTarget(ByVal ws As Worksheet, ByVal target As String)
    Dim dataStart As Long
    Dim headerRow As Long

    dataStart = TargetBarisAwal(target)

    ' Header N-1-2 mempunyai beberapa variasi teks dan merge antar-template.
    ' Posisi kolom serta baris datanya tetap baku, sehingga validasi teks header
    ' dilewati agar variasi tampilan tidak menghentikan proses sinkronisasi.
    If target = TARGET_N1 Then Exit Sub

    If target = TARGET_N1 Or target = TARGET_N19 Then
        headerRow = dataStart - 2
    Else
        headerRow = dataStart - 3
    End If

    If NormalisasiLabelHeader(ws.Cells(headerRow, 1).Value2) <> "NO" Then
        Err.Raise vbObjectError + 504, , _
                  "Struktur header sheet '" & target & "' tidak sesuai."
    End If

    If target = TARGET_N19 Then
        If NormalisasiTeks(ws.Cells(dataStart - 1, 6).Value2) <> "TANGGAL" Then
            Err.Raise vbObjectError + 505, , _
                      "Subheader tanggal pada '" & target & "' tidak ditemukan."
        End If
    Else
        If NormalisasiTeks(ws.Cells(dataStart - 1, 5).Value2) <> "TRANSACTION CODE" Or _
           NormalisasiTeks(ws.Cells(dataStart - 1, 7).Value2) <> "SERIAL NUMBER" Then
            Err.Raise vbObjectError + 506, , _
                      "Subheader faktur pada '" & target & "' tidak sesuai."
        End If
    End If
End Sub

Private Function NormalisasiLabelHeader(ByVal nilai As Variant) As String
    Dim s As String

    s = NormalisasiTeks(nilai)
    s = Replace$(s, ".", vbNullString)
    s = Replace$(s, ":", vbNullString)
    s = Replace$(s, vbCr, " ")
    s = Replace$(s, vbLf, " ")
    NormalisasiLabelHeader = Trim$(s)
End Function

Private Function CariBarisTotal( _
    ByVal ws As Worksheet, _
    ByVal dataStart As Long) As Long

    Dim lastRow As Long
    Dim lastCol As Long
    Dim r As Long
    Dim c As Long
    Dim formulaText As String

    lastRow = BarisTerakhirIsi(ws)
    lastCol = KolomTerakhirIsi(ws)

    For r = lastRow To dataStart Step -1
        For c = 1 To lastCol
            If ws.Cells(r, c).HasFormula Then
                formulaText = UCase$(CStr(ws.Cells(r, c).Formula))
                If InStr(1, formulaText, "SUM(", vbTextCompare) > 0 Or _
                   InStr(1, formulaText, "SUBTOTAL(", vbTextCompare) > 0 Then
                    CariBarisTotal = r
                    Exit Function
                End If
            End If
        Next c
    Next r
End Function

Private Function AmbilTabelWajib( _
    ByVal namaSheet As String, _
    ByVal namaTabel As String) As ListObject

    Dim ws As Worksheet

    If Not SheetAda(namaSheet) Then
        Err.Raise vbObjectError + 507, , _
                  "Sheet sumber '" & namaSheet & "' tidak ditemukan."
    End If

    Set ws = ThisWorkbook.Worksheets(namaSheet)
    On Error Resume Next
    Set AmbilTabelWajib = ws.ListObjects(namaTabel)
    On Error GoTo 0

    If AmbilTabelWajib Is Nothing Then
        Err.Raise vbObjectError + 508, , _
                  "Tabel sumber '" & namaTabel & "' tidak ditemukan pada sheet '" & _
                  namaSheet & "'."
    End If
End Function

Private Function IndeksKolom(ByVal lo As ListObject) As Object
    Dim hasil As Object
    Dim lc As ListColumn

    Set hasil = CreateObject("Scripting.Dictionary")
    hasil.CompareMode = vbTextCompare

    For Each lc In lo.ListColumns
        hasil(NormalisasiTeks(lc.Name)) = lc.Index
    Next lc

    Set IndeksKolom = hasil
End Function

Private Sub ValidasiKolomSumber(ByVal idx As Object, ByVal sumberPK As Boolean)
    Dim wajib As Variant
    Dim nama As Variant
    Dim idxReference As Long

    wajib = Array( _
        "SELLER_TIN", "SELLER_TAXPAYER_NAME", _
        "BUYER_TIN", "BUYER_TAXPAYER_NAME", _
        "TAX_INVOICE_CODE", "TAX_INVOICE_NUMBER", _
        "TAX_INVOICE_DATE", "TAX_INVOICE_PERIOD", "TAX_INVOICE_YEAR", _
        "TAX_INVOICE_STATUS", "BUYER_STATUS", _
        "SELLING_PRICE", "OTHER_TAX_BASE", "VAT")

    For Each nama In wajib
        If Not idx.Exists(CStr(nama)) Then
            Err.Raise vbObjectError + 509, , _
                      "Kolom sumber '" & CStr(nama) & "' tidak ditemukan."
        End If
    Next nama

    If sumberPK Then
        idxReference = IndeksReferencePK(idx)
    End If

    If Not idx.Exists("DETAIL_BARANG_JASA") Then
        Err.Raise vbObjectError + 510, , _
                  "Kolom sumber 'DETAIL_BARANG_JASA' tidak ditemukan."
    End If
End Sub

Private Function DaftarTarget() As Variant
    DaftarTarget = Array( _
        TARGET_N1, TARGET_N2, TARGET_N3, TARGET_N4, _
        TARGET_N6, TARGET_N7, TARGET_N8, _
        TARGET_N19, TARGET_N20, TARGET_N21, TARGET_N22, TARGET_N24)
End Function

Private Function TargetBarisAwal(ByVal target As String) As Long
    If target = TARGET_N1 Or target = TARGET_N19 Then
        TargetBarisAwal = 21
    Else
        TargetBarisAwal = 22
    End If
End Function

Private Function TargetJumlahKolom(ByVal target As String) As Long
    Select Case target
        Case TARGET_N1: TargetJumlahKolom = 8
        Case TARGET_N19: TargetJumlahKolom = 9
        Case TARGET_N4: TargetJumlahKolom = 12
        Case TARGET_N21: TargetJumlahKolom = 12
        Case TARGET_N20, TARGET_N22, TARGET_N24: TargetJumlahKolom = 13
        Case Else: TargetJumlahKolom = 13
    End Select
End Function

Private Function TargetMemilikiDPPNilaiLain(ByVal target As String) As Boolean
    Select Case target
        Case TARGET_N2, TARGET_N3, TARGET_N6, TARGET_N7, TARGET_N8, _
             TARGET_N20, TARGET_N22, TARGET_N24
            TargetMemilikiDPPNilaiLain = True
    End Select
End Function

Private Function TargetMemilikiDetail(ByVal target As String) As Boolean
    Select Case target
        Case TARGET_N1, TARGET_N2, TARGET_N3, TARGET_N4, TARGET_N6, _
             TARGET_N7, TARGET_N8, TARGET_N19, TARGET_N20, TARGET_N21, _
             TARGET_N22, TARGET_N24
            TargetMemilikiDetail = True
    End Select
End Function

Private Function KolomDetail(ByVal target As String) As String
    If target = TARGET_N1 Then
        KolomDetail = "H"
    ElseIf target = TARGET_N19 Then
        KolomDetail = "I"
    ElseIf target = TARGET_N4 Or target = TARGET_N21 Then
        KolomDetail = "L"
    Else
        KolomDetail = "M"
    End If
End Function

Private Function TargetAdalahPM(ByVal target As String) As Boolean
    TargetAdalahPM = (target = TARGET_N19 Or target = TARGET_N20 Or _
                      target = TARGET_N21 Or target = TARGET_N22 Or _
                      target = TARGET_N24)
End Function

Private Function KolomTotal(ByVal target As String) As Variant
    Select Case target
        Case TARGET_N1
            KolomTotal = Array("F")
        Case TARGET_N19
            KolomTotal = Array("G", "H")
        Case TARGET_N4, TARGET_N21
            KolomTotal = Array("I", "J")
        Case Else
            KolomTotal = Array("I", "J", "K")
    End Select
End Function

Private Function KolomNilaiTarget(ByVal target As String) As Variant
    Select Case target
        Case TARGET_N1
            KolomNilaiTarget = Array(6)
        Case TARGET_N19
            KolomNilaiTarget = Array(7, 8)
        Case TARGET_N4, TARGET_N21
            KolomNilaiTarget = Array(9, 10)
        Case Else
            KolomNilaiTarget = Array(9, 10, 11)
    End Select
End Function

Private Function JumlahMetrikTarget(ByVal target As String) As Long
    Select Case target
        Case TARGET_N1
            JumlahMetrikTarget = 1
        Case TARGET_N19, TARGET_N4, TARGET_N21
            JumlahMetrikTarget = 2
        Case Else
            JumlahMetrikTarget = 3
    End Select
End Function

Private Function KodeTransaksi(ByVal nomorFaktur As String) As String
    Dim nomorKompak As String

    nomorKompak = NomorFakturKompak(nomorFaktur)
    If Len(nomorKompak) >= 3 Then
        KodeTransaksi = Left$(nomorKompak, 3)
    Else
        KodeTransaksi = nomorKompak
    End If
End Function

Private Function TipePembetulan(ByVal nomorFaktur As String) As String
    Dim nomorKompak As String

    nomorKompak = NomorFakturKompak(nomorFaktur)
    If Len(nomorKompak) >= 4 Then
        TipePembetulan = Mid$(nomorKompak, 4, 1)
    End If
End Function

Private Function NomorFakturKompak(ByVal nomorFaktur As String) As String
    Dim s As String

    s = Trim$(nomorFaktur)
    s = Replace$(s, ".", vbNullString)
    s = Replace$(s, "-", vbNullString)
    s = Replace$(s, "_", vbNullString)
    s = Replace$(s, " ", vbNullString)
    NomorFakturKompak = s
End Function

Private Function TINSeluruhnyaNol(ByVal nilai As Variant) As Boolean
    Dim s As String
    Dim digit As String
    Dim ch As String
    Dim i As Long

    s = Trim$(CStr(nilai))
    If Len(s) = 0 Then Exit Function

    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then
            digit = digit & ch
        ElseIf InStr(1, " .-/'", ch, vbBinaryCompare) = 0 Then
            Exit Function
        End If
    Next i

    If Len(digit) > 0 Then
        TINSeluruhnyaNol = (Len(Replace$(digit, "0", vbNullString)) = 0)
    End If
End Function

Private Function NormalisasiTeks(ByVal nilai As Variant) As String
    NormalisasiTeks = UCase$(Trim$(CStr(nilai)))
End Function

Private Function NormalisasiKodeFaktur(ByVal nilai As Variant) As String
    Dim s As String

    s = NormalisasiTeks(nilai)
    s = Replace$(s, ".", vbNullString)
    s = Replace$(s, "-", vbNullString)
    s = Replace$(s, "_", vbNullString)
    s = Replace$(s, " ", vbNullString)
    NormalisasiKodeFaktur = s
End Function

Private Function AngkaBulat(ByVal nilai As Variant) As Long
    If IsNumeric(nilai) Then AngkaBulat = CLng(Val(CStr(nilai)))
End Function

Private Function NilaiTanggal(ByVal nilai As Variant) As Variant
    If IsDate(nilai) Then
        NilaiTanggal = CDate(nilai)
    Else
        NilaiTanggal = nilai
    End If
End Function

Private Function KunciTanggal(ByVal nilai As Variant) As String
    If IsDate(nilai) Then
        KunciTanggal = Format$(CDate(nilai), "yyyymmdd")
    Else
        KunciTanggal = "99999999" & CStr(nilai)
    End If
End Function

Private Function SheetAda(ByVal namaSheet As String) As Boolean
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(namaSheet)
    On Error GoTo 0
    SheetAda = Not ws Is Nothing
End Function

Private Function BarisTerakhirIsi(ByVal ws As Worksheet) As Long
    Dim sel As Range

    On Error Resume Next
    Set sel = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), _
              LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, _
              SearchDirection:=xlPrevious, MatchCase:=False)
    On Error GoTo 0
    If Not sel Is Nothing Then BarisTerakhirIsi = sel.Row
End Function

Private Function KolomTerakhirIsi(ByVal ws As Worksheet) As Long
    Dim sel As Range

    On Error Resume Next
    Set sel = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), _
              LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, _
              SearchDirection:=xlPrevious, MatchCase:=False)
    On Error GoTo 0
    If Not sel Is Nothing Then KolomTerakhirIsi = sel.Column
End Function

Private Sub TulisLogSinkron(ByVal level As String, ByVal pesan As String)
    Dim ws As Worksheet
    Dim nextRow As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_LOG)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add( _
                    After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = SHEET_LOG
        ws.Range("A1:C1").Value = Array("WAKTU", "LEVEL", "PESAN")
        ws.Range("A1:C1").Font.Bold = True
    End If

    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(nextRow, 1).Value = Now
    ws.Cells(nextRow, 1).NumberFormat = "dd-mm-yyyy hh:mm:ss"
    ws.Cells(nextRow, 2).Value = level
    ws.Cells(nextRow, 3).Value = pesan
    ws.Columns("A:C").AutoFit
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
