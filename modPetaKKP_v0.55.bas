Attribute VB_Name = "modPetaKKP"
'-------------------------------------------------------------
'THIS MODULE IS FUELED BY SPITE OF EWS AND CORETAX
'@2026
'-------------------------------------------------------------
Option Explicit

'EWSSYNC v0.5 - namespace serta katalog sheet KKP.

Public Const TAXKEY_PPN_DN As String = "PPN"
Public Const TAXKEY_PPH21 As String = "P21"
Public Const TAXKEY_PPH21F As String = "P21F"
Public Const TAXKEY_PPH23 As String = "P23"
Public Const TAXKEY_PPH26 As String = "P26"
Public Const TAXKEY_PPH42 As String = "P42"

Public Const KODE_EWS_PPN_DN As String = "WP_VAT_VAT_TOT01"
Public Const KODE_EWS_PPH21 As String = "WP_ICT_WH_TOT01"
Public Const KODE_EWS_PPH21F As String = "WP_ICT_WH_TOT02"
Public Const KODE_EWS_PPH23 As String = "WP_ICT_WH_TOT08"
Public Const KODE_EWS_PPH26 As String = "WP_ICT_WH_TOT09"
Public Const KODE_EWS_PPH42 As String = "WP_ICT_WH_TOT10"

Public Const NAMA_EWS_PPN_DN As String = "PPN Dalam Negeri"
Public Const NAMA_EWS_PPH21 As String = "PPh Pasal 21"
Public Const NAMA_EWS_PPH21F As String = "PPh Pasal 21 Final"
Public Const NAMA_EWS_PPH23 As String = "PPh Pasal 23"
Public Const NAMA_EWS_PPH26 As String = "PPh Pasal 26"
Public Const NAMA_EWS_PPH42 As String = "PPh Pasal 4 ayat (2)"

Public Const NAMA_SHEET_MAIN_EWS As String = "MAIN LEVEL"

'Bagian: normalisasi kode.
'Tujuan: Menormalkan kode kkpews.
Public Function NormalisasiKodeKKPEWS( _
    ByVal kode As String) As String

    Dim hasil As String
    Dim posisi As Long

    hasil = UCase$(Trim$(kode))
    hasil = Replace(hasil, " ", vbNullString)
    hasil = Replace(hasil, ".", "-")

    If Left$(hasil, 4) = "PPN-" Then
        hasil = Mid$(hasil, 5)
    ElseIf Left$(hasil, 5) = "P21F-" Then
        hasil = Mid$(hasil, 6)
    ElseIf Left$(hasil, 4) = "P21-" Then
        hasil = Mid$(hasil, 5)
    ElseIf Left$(hasil, 4) = "P23-" Then
        hasil = Mid$(hasil, 5)
    ElseIf Left$(hasil, 4) = "P26-" Then
        hasil = Mid$(hasil, 5)
    ElseIf Left$(hasil, 4) = "P42-" Then
        hasil = Mid$(hasil, 5)
    End If

    Do While InStr(1, hasil, "--", vbBinaryCompare) > 0
        hasil = Replace(hasil, "--", "-")
    Loop

    If Len(hasil) >= 2 Then

        posisi = InStr(1, hasil, "-", vbBinaryCompare)

        If posisi = 0 _
           And Left$(hasil, 1) Like "[A-Z]" _
           And IsNumeric(Mid$(hasil, 2)) Then

            hasil = Left$(hasil, 1) & "-" & Mid$(hasil, 2)

        End If

    End If

    NormalisasiKodeKKPEWS = hasil

End Function

'Tujuan: Membentuk kunci kkpkomposit.
Public Function KunciKKPKomposit( _
    ByVal taxKey As String, _
    ByVal kodeEWS As String) As String

    KunciKKPKomposit = _
        UCase$(Trim$(taxKey)) & "|" & _
        NormalisasiKodeKKPEWS(kodeEWS)

End Function

'Bagian: resolusi nama sheet.
'Tujuan: Mengembalikan nama sheet lokal kkp.
Public Function NamaSheetLokalKKP( _
    ByVal taxKey As String, _
    ByVal kodeEWS As String) As String

    Dim kodeNormal As String

    kodeNormal = NormalisasiKodeKKPEWS(kodeEWS)

    Select Case UCase$(Trim$(taxKey))

        Case TAXKEY_PPN_DN
            NamaSheetLokalKKP = "PPN-" & kodeNormal

        Case TAXKEY_PPH21
            NamaSheetLokalKKP = "P21-" & kodeNormal

        Case TAXKEY_PPH21F
            NamaSheetLokalKKP = "P21F-" & kodeNormal

        Case TAXKEY_PPH23
            NamaSheetLokalKKP = "P23-" & kodeNormal
        Case TAXKEY_PPH26
            NamaSheetLokalKKP = "P26-" & kodeNormal
        Case TAXKEY_PPH42
            NamaSheetLokalKKP = "P42-" & kodeNormal

        Case Else
            NamaSheetLokalKKP = _
                UCase$(Trim$(taxKey)) & "-" & kodeNormal

    End Select

End Function

'Tujuan: Mengembalikan nama sheet ringkasan pajak.
Public Function NamaSheetRingkasanPajak( _
    ByVal taxKey As String) As String

    Select Case UCase$(Trim$(taxKey))

        Case TAXKEY_PPN_DN
            NamaSheetRingkasanPajak = "PPN"

        Case TAXKEY_PPH21
            NamaSheetRingkasanPajak = "PPh 21"

        Case TAXKEY_PPH21F
            NamaSheetRingkasanPajak = "PPh 21F"

        Case TAXKEY_PPH23
            NamaSheetRingkasanPajak = "PPh 23"
        Case TAXKEY_PPH26
            NamaSheetRingkasanPajak = "PPh 26"
        Case TAXKEY_PPH42
            NamaSheetRingkasanPajak = "PPh 4(2)"

        Case Else
            NamaSheetRingkasanPajak = UCase$(Trim$(taxKey))

    End Select

End Function

'Tujuan: Mengembalikan nama sheet kelompok pajak.
Public Function NamaSheetKelompokPajak( _
    ByVal taxKey As String, _
    ByVal kelompok As String) As String

    NamaSheetKelompokPajak = _
        UCase$(Trim$(taxKey)) & "-" & _
        UCase$(Trim$(kelompok))

End Function

'Tujuan: Mengambil sheet lokal kkp.
Public Function DapatkanSheetLokalKKP( _
    ByVal taxKey As String, _
    ByVal kodeEWS As String, _
    Optional ByVal izinkanLegacy As Boolean = True) As Worksheet

    Dim ws As Worksheet
    Dim namaBaru As String
    Dim namaLegacy As String

    namaBaru = NamaSheetLokalKKP(taxKey, kodeEWS)

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(namaBaru)
    On Error GoTo 0

    If Not ws Is Nothing Then
        Set DapatkanSheetLokalKKP = ws
        Exit Function
    End If

    If Not izinkanLegacy Then Exit Function

    If UCase$(Trim$(taxKey)) = TAXKEY_PPN_DN Then

        namaLegacy = Replace( _
            NormalisasiKodeKKPEWS(kodeEWS), _
            "-", _
            ".")

        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(namaLegacy)
        On Error GoTo 0

        If Not ws Is Nothing Then
            Set DapatkanSheetLokalKKP = ws
        End If

    End If

End Function

'Tujuan: Memeriksa sheet lokal kkpada.
Public Function SheetLokalKKPAda( _
    ByVal taxKey As String, _
    ByVal kodeEWS As String, _
    Optional ByVal izinkanLegacy As Boolean = True) As Boolean

    Dim ws As Worksheet

    Set ws = DapatkanSheetLokalKKP( _
        taxKey, kodeEWS, izinkanLegacy)

    SheetLokalKKPAda = Not ws Is Nothing

End Function

'Bagian: katalog KKP.
'Tujuan: Menyediakan kode ewsppn.
Public Function DaftarKodeEWSPPN() As Variant

    DaftarKodeEWSPPN = Array( _
        "N-1", "N-2", "N-3", "N-4", "N-5", _
        "N-6", "N-7", "N-8", "N-9", "N-10", _
        "N-19", "N-20", "N-21", "N-22", "N-23", _
        "N-24", "N-25", _
        "O-1", "O-2", "O-3", "O-4", _
        "P-1", "P-2", "P-3", "P-4", _
        "Q-1", "Q-2")

End Function

'Tujuan: Menyediakan kode ewspph 21.
Public Function DaftarKodeEWSPPh21() As Variant

    DaftarKodeEWSPPh21 = Array( _
        "B-1", "C-1", _
        "N-1", "N-2", "N-3", "N-4", _
        "N-5", "N-6", "N-7", "N-8", _
        "O-1", "O-2", "O-3", _
        "P-1", "P-2", "P-3")

End Function

'Tujuan: Menyediakan kode EWS PPh 21 Final.
Public Function DaftarKodeEWSPPh21F() As Variant

    DaftarKodeEWSPPh21F = Array( _
        "B-1", "B-1-3", "C-1", _
        "N-1", "N-2", "N-3", "N-4", _
        "N-5", "N-6", "N-7", "N-8", _
        "O-1", "O-2", "O-3", _
        "P-1", "P-2", "P-3")

End Function

'Bagian: urutan kalkulasi.
'Tujuan: Menghitung ulang kkplokal.
Public Sub HitungUlangKKPLokal(ByVal taxKey As String)

    Dim ws As Worksheet
    Dim namaRingkasan As String
    Dim prefix As String

    If UCase$(Trim$(taxKey)) = TAXKEY_PPH23 Then
        HitungUlangAlurPPh23
        Exit Sub
    End If

    If UCase$(Trim$(taxKey)) = TAXKEY_PPH21F Then
        HitungUlangAlurPPh21F
        Exit Sub
    End If

    If UCase$(Trim$(taxKey)) = TAXKEY_PPH26 Then
        HitungUlangAlurPPh26
        Exit Sub
    End If

    If UCase$(Trim$(taxKey)) = TAXKEY_PPH42 Then
        HitungUlangAlurPPh42
        Exit Sub
    End If

    namaRingkasan = NamaSheetRingkasanPajak(taxKey)
    prefix = UCase$(Trim$(taxKey)) & "-"

    On Error Resume Next

    Set ws = Nothing
    Set ws = ThisWorkbook.Worksheets(namaRingkasan)
    If Not ws Is Nothing Then ws.Calculate

    For Each ws In ThisWorkbook.Worksheets

        If UCase$(Left$(ws.Name, Len(prefix))) = prefix Then
            ws.Calculate
        End If

    Next ws

    On Error GoTo 0

End Sub

'Tujuan: Menghitung detail, ringkasan, dan induk PPh 21 Final.
Private Sub HitungUlangAlurPPh21F()

    Dim ws As Worksheet
    Dim ringkasan As Worksheet
    Dim induk As Worksheet
    Dim daftar As Variant
    Dim kode As Variant

    On Error Resume Next
    Set ringkasan = ThisWorkbook.Worksheets( _
        NamaSheetRingkasanPajak(TAXKEY_PPH21F))
    Set induk = ThisWorkbook.Worksheets( _
        NamaSheetIndukPajak(TAXKEY_PPH21F))
    On Error GoTo 0

    If ringkasan Is Nothing Then Exit Sub

    daftar = DaftarKodeEWSPPh21F()

    For Each kode In daftar
        Set ws = DapatkanSheetLokalKKP( _
            TAXKEY_PPH21F, CStr(kode), False)
        If Not ws Is Nothing Then ws.Calculate
    Next kode

    ringkasan.Calculate
    If Not induk Is Nothing Then induk.Calculate

End Sub

'Tujuan: Mengembalikan nama sheet induk pajak.
Public Function NamaSheetIndukPajak(ByVal taxKey As String) As String
    NamaSheetIndukPajak = NamaSheetRingkasanPajak(taxKey) & "-INDUK"
End Function

'Tujuan: Menyediakan kode ewspph 23.
Public Function DaftarKodeEWSPPh23() As Variant
    DaftarKodeEWSPPh23 = Array( _
        "B-1", "C-1", _
        "N-1", "N-2", "N-3", "N-4", _
        "N-5", "N-6", "N-7", "N-8", _
        "O-1", "O-2", "O-3", _
        "P-1", "P-2", "P-3")
End Function

'Tujuan: Menghitung ulang alur pph 23.
Private Sub HitungUlangAlurPPh23()
    Dim ws As Worksheet
    Dim ringkasan As Worksheet
    Dim induk As Worksheet
    Dim daftar As Variant
    Dim kode As Variant

    On Error Resume Next
    Set ringkasan = ThisWorkbook.Worksheets(NamaSheetRingkasanPajak(TAXKEY_PPH23))
    Set induk = ThisWorkbook.Worksheets(NamaSheetIndukPajak(TAXKEY_PPH23))
    On Error GoTo 0
    If ringkasan Is Nothing Then Exit Sub

    ringkasan.Calculate
    daftar = DaftarKodeEWSPPh23()
    For Each kode In daftar
        Set ws = DapatkanSheetLokalKKP(TAXKEY_PPH23, CStr(kode), False)
        If Not ws Is Nothing Then ws.Calculate
    Next kode
    ringkasan.Calculate
    If Not induk Is Nothing Then induk.Calculate
End Sub

'Tujuan: Menyediakan kode ewspph 26.
Public Function DaftarKodeEWSPPh26() As Variant
    DaftarKodeEWSPPh26 = Array( _
        "B-1", "B-2", "C-1", _
        "N-1", "N-2", "N-3", "N-4", _
        "N-5", "N-6", "N-7", "N-8", _
        "O-1", "O-2", "O-3", _
        "P-1", "P-2", "P-3")
End Function

'Tujuan: Menghitung ulang alur pph 26.
Private Sub HitungUlangAlurPPh26()
    Dim ws As Worksheet
    Dim ringkasan As Worksheet
    Dim induk As Worksheet
    Dim daftar As Variant
    Dim kode As Variant

    On Error Resume Next
    Set ringkasan = ThisWorkbook.Worksheets(NamaSheetRingkasanPajak(TAXKEY_PPH26))
    Set induk = ThisWorkbook.Worksheets(NamaSheetIndukPajak(TAXKEY_PPH26))
    On Error GoTo 0
    If ringkasan Is Nothing Then Exit Sub

    ringkasan.Calculate
    daftar = DaftarKodeEWSPPh26()
    For Each kode In daftar
        Set ws = DapatkanSheetLokalKKP(TAXKEY_PPH26, CStr(kode), False)
        If Not ws Is Nothing Then ws.Calculate
    Next kode
    ringkasan.Calculate
    If Not induk Is Nothing Then induk.Calculate
End Sub

'Tujuan: Menyediakan kode ewspph 42.
Public Function DaftarKodeEWSPPh42() As Variant
    DaftarKodeEWSPPh42 = Array( _
        "B-1", "B-2", "B-3", "B-4", "C-1", _
        "N-1", "N-2", "N-3", "N-4", _
        "N-5", "N-6", "N-7", "N-8", _
        "O-1", "O-2", "O-3", _
        "P-1", "P-2", "P-3")
End Function

'Tujuan: Menghitung ulang alur pph 42.
Private Sub HitungUlangAlurPPh42()
    Dim ws As Worksheet, ringkasan As Worksheet, induk As Worksheet
    Dim daftar As Variant, kode As Variant
    On Error Resume Next
    Set ringkasan = ThisWorkbook.Worksheets(NamaSheetRingkasanPajak(TAXKEY_PPH42))
    Set induk = ThisWorkbook.Worksheets(NamaSheetIndukPajak(TAXKEY_PPH42))
    On Error GoTo 0
    If ringkasan Is Nothing Then Exit Sub
    ringkasan.Calculate
    daftar = DaftarKodeEWSPPh42()
    For Each kode In daftar
        Set ws = DapatkanSheetLokalKKP(TAXKEY_PPH42, CStr(kode), False)
        If Not ws Is Nothing Then ws.Calculate
    Next kode
    ringkasan.Calculate
    If Not induk Is Nothing Then induk.Calculate
End Sub
