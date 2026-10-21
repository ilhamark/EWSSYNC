Attribute VB_Name = "modEngineExportEWS"
'-------------------------------------------------------------
'THIS MODULE IS FUELED BY SPITE OF EWS AND CORETAX
'@2026
'-------------------------------------------------------------
Option Explicit

'EWSSYNC v0.5 - engine ekspor dan validasi transaksi.

'Tujuan: Mengekspor KKP terpilih ke EWS dengan validasi dan rollback.

Private Const KODE_PPN_DN_ENGINE As String = _
    "WP_VAT_VAT_TOT01"

Private Const KODE_PPH21_ENGINE As String = _
    "WP_ICT_WH_TOT01"

Private Const KODE_PPH21F_ENGINE As String = _
    "WP_ICT_WH_TOT02"

Private Const KODE_PPH23_ENGINE As String = "WP_ICT_WH_TOT08"

Private Const KODE_PPH26_ENGINE As String = "WP_ICT_WH_TOT09"
Private Const KODE_PPH42_ENGINE As String = "WP_ICT_WH_TOT10"

Private Const NAMA_SHEET_MAIN_EWS_ENGINE As String = _
    "MAIN LEVEL"

Private Const NAMA_SHEET_LOG_EXPORT As String = _
    "Log Ekspor EWS"

Private Const TEKS_TEKNIK_LANGSUNG As String = _
    "PEMERIKSAANLANGSUNGTERHADAPBUKTI"

Private Const BATAS_TUNGGU_DETIK As Double = 3

'---------------------------------------------------------------------
' TITIK MASUK DARI frmExportEWS
'---------------------------------------------------------------------
Public Sub EksporTerpilihEWS(ByVal frm As Object)

    Dim daftarKKP As Variant
    Dim daftarKode As String
    Dim kodeJenis As String
    Dim taxKey As String
    Dim namaJenisPajak As String

    Dim jumlahFile As Long
    Dim jumlahKKP As Long
    Dim jumlahBerhasil As Long
    Dim jumlahGagal As Long
    Dim jumlahNilai As Long

    Dim i As Long
    Dim filePath As String
    Dim bulanFile As Long
    Dim tahunFile As Long
    Dim npwpFile As String
    Dim npwpReferensi As String
    Dim namaWPFile As String
    Dim namaWPReferensi As String
    Dim pesanError As String

    Dim oldScreenUpdating As Boolean
    Dim oldDisplayAlerts As Boolean
    Dim oldEnableEvents As Boolean
    Dim oldStatusBar As Variant
    Dim stateTersimpan As Boolean
    Dim errorAwal As String
    Dim tahapBatch As String
    Dim hasilFile As Boolean
    Dim perbaruiCFMWP As Boolean
    Dim gayaRingkasan As VbMsgBoxStyle

    On Error GoTo GagalUmum

    tahapBatch = "Membaca pilihan form"
    kodeJenis = KodeJenisPajakExportEWS(frm)
    taxKey = TaxKeyDariKodeJenisEngineExport(kodeJenis)
    namaJenisPajak = NamaJenisPajakExportEWS(frm)
    On Error Resume Next
    perbaruiCFMWP = CBool(frm.chkPerbaruiCFMWP.Value)
    On Error GoTo GagalUmum

    If Len(taxKey) = 0 Then

        MsgBox _
            "Engine ekspor untuk jenis pajak ini belum tersedia.", _
            vbInformation, _
            "Ekspor EWS"

        Exit Sub

    End If

    daftarKode = DaftarKodeKKPTerpilihExportEWS(frm)

    If Len(daftarKode) = 0 Then

        MsgBox _
            "Pilih sekurang-kurangnya satu KKP.", _
            vbExclamation, _
            "Ekspor EWS"

        Exit Sub

    End If

    daftarKKP = Split(daftarKode, "|")
    jumlahKKP = UBound(daftarKKP) - LBound(daftarKKP) + 1

    'Validasi ringan seluruh daftar sebelum file dibuka.
    For i = 0 To frm.lstEWS.ListCount - 1

        filePath = Trim$(CStr(frm.lstEWS.List(i, 3)))

        If Len(filePath) > 0 Then

            jumlahFile = jumlahFile + 1

            If Len(Dir$(filePath)) = 0 Then

                MsgBox _
                    "File EWS tidak ditemukan:" & vbCrLf & _
                    filePath, _
                    vbExclamation, _
                    "Ekspor Dibatalkan"

                Exit Sub

            End If

            If InStr( _
                1, _
                Dir$(filePath), _
                kodeJenis, _
                vbTextCompare) = 0 Then

                MsgBox _
                    "Jenis file tidak sesuai dengan pilihan form:" & _
                    vbCrLf & Dir$(filePath), _
                    vbExclamation, _
                    "Ekspor Dibatalkan"

                Exit Sub

            End If

            npwpFile = AmbilNPWPDariNamaFileExport(filePath)

            If Len(npwpFile) = 0 Then

                MsgBox _
                    "NPWP tidak ditemukan pada nama file:" & _
                    vbCrLf & Dir$(filePath), _
                    vbExclamation, _
                    "Ekspor Dibatalkan"

                Exit Sub

            End If

            If Len(npwpReferensi) = 0 Then
                npwpReferensi = npwpFile
            ElseIf StrComp( _
                npwpReferensi, _
                npwpFile, _
                vbTextCompare) <> 0 Then

                MsgBox _
                    "Daftar memuat file dari NPWP berbeda.", _
                    vbExclamation, _
                    "Ekspor Dibatalkan"

                Exit Sub

            End If

        End If

    Next i

    If jumlahFile = 0 Then

        MsgBox _
            "Belum ada file EWS yang dipilih.", _
            vbExclamation, _
            "Ekspor EWS"

        Exit Sub

    End If

    If MsgBox( _
        "Ekspor " & CStr(jumlahKKP) & _
        " KKP ke " & CStr(jumlahFile) & _
        " file EWS?" & vbCrLf & vbCrLf & _
        "Data ekspor sebelumnya akan ditimpa." & _
        IIf(perbaruiCFMWP, _
            vbCrLf & "CFM WP akan diperbarui dari KKP Lokal.", _
            vbNullString), _
        vbQuestion + vbYesNo + vbDefaultButton2, _
        "Konfirmasi Ekspor") <> vbYes Then

        Exit Sub

    End If

    oldScreenUpdating = Application.ScreenUpdating
    oldDisplayAlerts = Application.DisplayAlerts
    oldEnableEvents = Application.EnableEvents
    oldStatusBar = Application.StatusBar
    stateTersimpan = True

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    tahapBatch = "Menonaktifkan tombol form"
    frm.cmdExport.Enabled = False
    frm.cmdBrowse.Enabled = False

    For i = 0 To frm.lstEWS.ListCount - 1

        filePath = Trim$(CStr(frm.lstEWS.List(i, 3)))

        If Len(filePath) > 0 Then

            bulanFile = CLng(frm.lstEWS.List(i, 4))
            tahunFile = CLng(frm.lstEWS.List(i, 5))
            namaWPFile = vbNullString
            pesanError = vbNullString
            jumlahNilai = 0

    tahapBatch = "Memperbarui status sebelum proses file"
            frm.lstEWS.List(i, 2) = "Memvalidasi..."
            frm.lblStatus.caption = _
                "Memproses " & Dir$(filePath) & "..."

            Application.StatusBar = _
                "Ekspor EWS: " & Dir$(filePath)

            frm.Repaint
            DoEvents

            tahapBatch = "Memproses file EWS"
            hasilFile = ProsesSatuFileExportSelektif( _
                    filePath, _
                    bulanFile, _
                    tahunFile, _
                    daftarKKP, _
                    taxKey, _
                    perbaruiCFMWP, _
                    namaWPFile, _
                    jumlahNilai, _
                    pesanError)

            'Catat hasil sebelum menyentuh objek form yang mungkin terputus.
            CatatDiagnostikExport filePath, "Hasil transaksi", _
                IIf(hasilFile, "BERHASIL DISIMPAN; ", "GAGAL; ") & pesanError
            tahapBatch = "Memperbarui form/log setelah transaksi file"
            If hasilFile Then

                If Len(namaWPReferensi) = 0 Then
                    namaWPReferensi = namaWPFile
                    frm.namaWPReferensi = namaWPFile
                    frm.npwpReferensi = npwpReferensi
                    frm.lblWPValue.caption = namaWPFile
                End If

                frm.lstEWS.List(i, 2) = "Berhasil"
                jumlahBerhasil = jumlahBerhasil + 1

                CatatLogExportSelektif _
                    filePath, _
                    bulanFile, _
                    tahunFile, _
                    namaJenisPajak, _
                    daftarKode, _
                    "Berhasil", _
                    CStr(jumlahNilai) & " nilai ditulis" & _
                    IIf(Len(pesanError) > 0, "; " & pesanError, vbNullString)

            Else

                frm.lstEWS.List(i, 2) = "Gagal"
                jumlahGagal = jumlahGagal + 1

                CatatLogExportSelektif _
                    filePath, _
                    bulanFile, _
                    tahunFile, _
                    namaJenisPajak, _
                    daftarKode, _
                    "Gagal", _
                    pesanError

            End If

            frm.Repaint
            DoEvents

        End If

    Next i

Selesai:

    On Error Resume Next

    Application.ScreenUpdating = oldScreenUpdating
    Application.DisplayAlerts = oldDisplayAlerts
    Application.EnableEvents = oldEnableEvents
    Application.StatusBar = oldStatusBar

    frm.cmdBrowse.Enabled = True
    PerbaruiRingkasanFormExportEWS frm

    If jumlahBerhasil > 0 Or jumlahGagal > 0 Then

        frm.lblStatus.caption = _
            CStr(jumlahBerhasil) & " berhasil, " & _
            CStr(jumlahGagal) & " gagal."

        gayaRingkasan = vbInformation
        If jumlahGagal > 0 Then gayaRingkasan = vbExclamation

        MsgBox _
            "Data EWS selesai diekspor." & vbCrLf & vbCrLf & _
            "Berhasil: " & CStr(jumlahBerhasil) & vbCrLf & _
            "Gagal: " & CStr(jumlahGagal), _
            gayaRingkasan, _
            "Ekspor Selesai"

    End If

    On Error GoTo 0
    Exit Sub

GagalUmum:

    errorAwal = "Error " & CStr(Err.Number) & " [" & Err.Source & "]: " & Err.Description
    If Len(pesanError) > 0 Then errorAwal = errorAwal & vbCrLf & "Hasil file: " & pesanError
    CatatDiagnostikExport filePath, tahapBatch, errorAwal

    'Jangan akses frm lagi dan jangan Resume Selesai: referensi form dapat invalid.
    On Error Resume Next
    If stateTersimpan Then
        Application.ScreenUpdating = oldScreenUpdating
        Application.DisplayAlerts = oldDisplayAlerts
        Application.EnableEvents = oldEnableEvents
        Application.StatusBar = oldStatusBar
    End If
    MsgBox "Ekspor dihentikan." & vbCrLf & "Tahap: " & tahapBatch & _
        vbCrLf & "File: " & filePath & vbCrLf & errorAwal & vbCrLf & _
        "Log diagnostik: " & Environ$("TEMP") & "\EWSSYNC_Export_diagnostic.log", _
        vbCritical, "Ekspor EWS"
    On Error GoTo 0
    Exit Sub

End Sub

'---------------------------------------------------------------------
' TRANSAKSI SATU FILE
'---------------------------------------------------------------------
Private Function ProsesSatuFileExportSelektif( _
    ByVal filePath As String, _
    ByVal bulan As Long, _
    ByVal tahun As Long, _
    ByVal daftarKKP As Variant, _
    ByVal taxKey As String, _
    ByVal perbaruiCFMWP As Boolean, _
    ByRef namaWP As String, _
    ByRef jumlahNilai As Long, _
    ByRef pesanError As String) As Boolean

    Dim wbEWS As Workbook
    Dim wsMain As Worksheet
    Dim rencana As Collection
    Dim cadangan As Collection

    Dim kodeKKP As Variant
    Dim item As Variant
    Dim dataLama As Variant
    Dim target As Range

    Dim oldSecurity As Long
    Dim securityDiubah As Boolean
    Dim berhasil As Boolean
    Dim sudahDisimpanKeDisk As Boolean
    Dim tahapFile As String

    ProsesSatuFileExportSelektif = False
    pesanError = vbNullString
    jumlahNilai = 0

    Set rencana = New Collection
    Set cadangan = New Collection

    On Error GoTo Gagal

    tahapFile = "Memeriksa file yang sudah terbuka"
    If WorkbookTerbukaUntukExport(filePath) Then

        Err.Raise _
            vbObjectError + 5100, , _
            "File sedang terbuka. Tutup file EWS terlebih dahulu."

    End If

    oldSecurity = Application.AutomationSecurity
    securityDiubah = True

    Application.AutomationSecurity = _
        msoAutomationSecurityForceDisable

    tahapFile = "Membuka workbook EWS"
    If taxKey = TAXKEY_PPH42 Then
        'TOT10 menampilkan peringatan circular reference walau macro dimatikan.
        Application.SendKeys "{ENTER}", False
        DoEvents
    End If
    Set wbEWS = Workbooks.Open( _
        Filename:=filePath, _
        UpdateLinks:=0, _
        ReadOnly:=False, _
        IgnoreReadOnlyRecommended:=True, _
        AddToMru:=False)

    Application.AutomationSecurity = oldSecurity
    securityDiubah = False

    tahapFile = "Memeriksa akses workbook"
    If wbEWS.ReadOnly Then

        Err.Raise _
            vbObjectError + 5101, , _
            "File EWS terbuka sebagai read-only."

    End If

    If Not SheetAdaPadaWorkbookExport( _
            wbEWS, _
            NAMA_SHEET_MAIN_EWS_ENGINE) Then

        Err.Raise _
            vbObjectError + 5102, , _
            "Sheet MAIN LEVEL tidak ditemukan."

    End If

    Set wsMain = wbEWS.Worksheets( _
        NAMA_SHEET_MAIN_EWS_ENGINE)

    tahapFile = "Membaca identitas WP"
    namaWP = Trim$(CStr(wsMain.Range("AG7").Value2))

    tahapFile = "Kalkulasi workbook EWS"
    TungguPerhitunganEWS wbEWS

    If taxKey = TAXKEY_PPH21F Then
        tahapFile = "Kalkulasi KKP lokal PPh21 Final"
        HitungUlangKKPLokal TAXKEY_PPH21F
        TungguKalkulasiPPh23Export filePath, tahapFile
    ElseIf taxKey = TAXKEY_PPH23 Then
        tahapFile = "Kalkulasi KKP lokal PPh23"
        HitungUlangKKPLokal TAXKEY_PPH23
        TungguKalkulasiPPh23Export filePath, tahapFile
    ElseIf taxKey = TAXKEY_PPH26 Then
        tahapFile = "Kalkulasi KKP lokal PPh26"
        HitungUlangKKPLokal TAXKEY_PPH26
        TungguKalkulasiPPh23Export filePath, tahapFile
    ElseIf taxKey = TAXKEY_PPH42 Then
        tahapFile = "Kalkulasi KKP lokal PPh4(2)"
        HitungUlangKKPLokal TAXKEY_PPH42
        TungguKalkulasiPPh23Export filePath, tahapFile
    End If

    tahapFile = "Menyusun mapping"
    'Seluruh rencana dibuat sebelum ada nilai yang ditulis.
    For Each kodeKKP In daftarKKP

        tahapFile = "Mapping KKP " & CStr(kodeKKP)
        TambahRencanaKKPExport _
            wbEWS, _
            CStr(kodeKKP), _
            taxKey, _
            bulan, _
            tahun, _
            rencana

    Next kodeKKP

    If rencana.Count = 0 Then

        Err.Raise _
            vbObjectError + 5103, , _
            "Tidak ada nilai yang dapat diekspor."

    End If

    tahapFile = "Validasi target sebelum penulisan"
    If taxKey = TAXKEY_PPH21F Then ValidasiRencanaPPh21FExport wbEWS, rencana
    If taxKey = TAXKEY_PPH23 Then ValidasiRencanaPPh23Export wbEWS, rencana
    If taxKey = TAXKEY_PPH26 Then ValidasiRencanaPPh26Export wbEWS, rencana
    If taxKey = TAXKEY_PPH42 Then ValidasiRencanaPPh42Export wbEWS, rencana

    tahapFile = "Mencadangkan target"
    'Simpan formula/nilai lama untuk rollback selama workbook masih terbuka.
    For Each item In rencana

        Set target = wbEWS.Worksheets( _
            CStr(item(0))).Range(CStr(item(1)))

        If target.MergeCells Then
            Set target = target.MergeArea.Cells(1, 1)
        End If

        If target.HasFormula Then

            dataLama = Array( _
                target.Worksheet.Name, _
                target.Address(False, False), _
                True, _
                target.Formula, _
                Empty)

        Else

            dataLama = Array( _
                target.Worksheet.Name, _
                target.Address(False, False), _
                False, _
                vbNullString, _
                target.Value2)

        End If

        cadangan.Add dataLama

    Next item

    For Each item In rencana

        Set target = wbEWS.Worksheets( _
            CStr(item(0))).Range(CStr(item(1)))

        If target.MergeCells Then
            Set target = target.MergeArea.Cells(1, 1)
        End If

        tahapFile = "Menulis " & CStr(item(0)) & "!" & CStr(item(1))
        If Left$(UCase$(CStr(item(5))), 4) = "TEKS" Then
            target.Value2 = CStr(item(2))
        Else
            target.Value2 = CDbl(item(2))
            If UCase$(CStr(item(5))) = "ANGKA" Then
                jumlahNilai = jumlahNilai + 1
            End If
        End If

    Next item

    If UCase$(Trim$(taxKey)) = TAXKEY_PPN_DN Then
        tahapFile = "Mengekspor daftar faktur PPN"
        EksporDaftarFakturPPNKeEWS wbEWS, daftarKKP, bulan, tahun
    End If

    If perbaruiCFMWP Then
        tahapFile = "Memperbarui CFM WP dari KKP Lokal"
        TerapkanCFMWP054 wbEWS, taxKey, bulan, daftarKKP, jumlahNilai, pesanError
    End If

    tahapFile = "Kalkulasi workbook EWS"
    TungguPerhitunganEWS wbEWS

    If taxKey = TAXKEY_PPH21F Then
        TungguKalkulasiPPh23Export filePath, tahapFile
        tahapFile = "Verifikasi hasil PPh21 Final sebelum save"
        VerifikasiHasilPPh23Export wbEWS, rencana
        VerifikasiDetailPPh21FExport wbEWS, bulan, tahun, rencana
    ElseIf taxKey = TAXKEY_PPH23 Then
        TungguKalkulasiPPh23Export filePath, tahapFile
        tahapFile = "Verifikasi hasil sebelum save"
        VerifikasiHasilPPh23Export wbEWS, rencana
        pesanError = CatatanPPh23Export(wbEWS, rencana)
    ElseIf taxKey = TAXKEY_PPH26 Then
        TungguKalkulasiPPh23Export filePath, tahapFile
        tahapFile = "Verifikasi hasil sebelum save"
        VerifikasiHasilPPh23Export wbEWS, rencana
    ElseIf taxKey = TAXKEY_PPH42 Then
        TungguKalkulasiPPh23Export filePath, tahapFile
        tahapFile = "Verifikasi hasil sebelum save"
        VerifikasiHasilPPh23Export wbEWS, rencana
    End If
    tahapFile = "Menyimpan workbook EWS"
    wbEWS.Save
    sudahDisimpanKeDisk = True

    If taxKey = TAXKEY_PPH21F Or _
       taxKey = TAXKEY_PPH26 Or taxKey = TAXKEY_PPH42 Then
        tahapFile = "Menutup EWS untuk verifikasi disk " & taxKey
        wbEWS.Close SaveChanges:=False
        Set wbEWS = Nothing

        oldSecurity = Application.AutomationSecurity
        securityDiubah = True
        Application.AutomationSecurity = _
            msoAutomationSecurityForceDisable

        tahapFile = "Membuka ulang EWS untuk verifikasi disk " & taxKey
        If taxKey = TAXKEY_PPH42 Then
            Application.SendKeys "{ENTER}", False
            DoEvents
        End If
        Set wbEWS = Workbooks.Open( _
            Filename:=filePath, _
            UpdateLinks:=0, _
            ReadOnly:=False, _
            IgnoreReadOnlyRecommended:=True, _
            AddToMru:=False)

        Application.AutomationSecurity = oldSecurity
        securityDiubah = False

        tahapFile = "Verifikasi nilai bukti dan dropdown dari disk " & taxKey
        If taxKey = TAXKEY_PPH21F Then TungguPerhitunganEWS wbEWS
        VerifikasiHasilPPh23Export wbEWS, rencana
        If taxKey = TAXKEY_PPH21F Then _
            VerifikasiDetailPPh21FExport wbEWS, bulan, tahun, rencana
        CatatDiagnostikExport filePath, tahapFile, _
            "OK: nilai child dan selector teknik induk sesuai file tersimpan."
    End If

    berhasil = True

Keluar:

    On Error Resume Next

    If securityDiubah Then
        Application.AutomationSecurity = oldSecurity
    End If

    If Not wbEWS Is Nothing Then

        If berhasil Then
            wbEWS.Close SaveChanges:=False
        Else
            wbEWS.Close SaveChanges:=False
        End If

    End If

    On Error GoTo 0

    ProsesSatuFileExportSelektif = berhasil
    Exit Function

Gagal:

    pesanError = "Tahap: " & tahapFile & "; Error " & CStr(Err.Number) & _
        " [" & Err.Source & "]: " & Err.Description
    If sudahDisimpanKeDisk Then
        pesanError = "File sudah disimpan, tetapi verifikasi pascasimpan gagal. " & _
            pesanError
    End If
    CatatDiagnostikExport filePath, tahapFile, pesanError

    'Rollback memori workbook sebelum ditutup tanpa menyimpan.
    On Error Resume Next

    If Not wbEWS Is Nothing Then

        For Each dataLama In cadangan

            Set target = wbEWS.Worksheets( _
                CStr(dataLama(0))).Range(CStr(dataLama(1)))

            If CBool(dataLama(2)) Then
                target.Formula = CStr(dataLama(3))
            Else
                target.Value2 = dataLama(4)
            End If

        Next dataLama

    End If

    On Error GoTo 0

    berhasil = False
    Resume Keluar

End Function

'---------------------------------------------------------------------
' MEMBANGUN RENCANA SATU KKP
'---------------------------------------------------------------------
Private Sub TambahRencanaKKPExport( _
    ByVal wbEWS As Workbook, _
    ByVal kodeKKP As String, _
    ByVal taxKey As String, _
    ByVal bulan As Long, _
    ByVal tahun As Long, _
    ByVal rencana As Collection)

    Dim wsLokal As Worksheet
    Dim wsTarget As Worksheet
    Dim daftarKunci As Variant
    Dim daftarNilai As Variant
    Dim daftarAlamat As Variant
    Dim kodeEWS As String
    Dim teksTeknik As String
    Dim namaSheetLokal As String
    Dim i As Long

    kodeEWS = NormalisasiKodeKKPEWS(kodeKKP)
    If TargetPPNExportDilewati(kodeEWS) Then Exit Sub

    If UCase$(Trim$(taxKey)) = TAXKEY_PPH21F Then
        TambahRencanaPPh21FExport wbEWS, kodeEWS, bulan, tahun, rencana
        Exit Sub
    ElseIf UCase$(Trim$(taxKey)) = TAXKEY_PPH23 Then
        TambahRencanaPPh23Export wbEWS, kodeEWS, bulan, tahun, rencana
        Exit Sub
    ElseIf UCase$(Trim$(taxKey)) = TAXKEY_PPH26 Then
        TambahRencanaPPh26Export wbEWS, kodeEWS, bulan, tahun, rencana
        Exit Sub
    ElseIf UCase$(Trim$(taxKey)) = TAXKEY_PPH42 Then
        TambahRencanaPPh42Export wbEWS, kodeEWS, bulan, tahun, rencana
        Exit Sub
    End If

    Set wsLokal = DapatkanSheetLokalKKP( _
        taxKey, _
        kodeEWS, _
        Not TargetPPNExportOpsional(kodeEWS))

    If wsLokal Is Nothing Then

        If TargetPPNExportOpsional(kodeEWS) Then Exit Sub

        namaSheetLokal = NamaSheetLokalKKP( _
            taxKey, kodeEWS)

        Err.Raise _
            vbObjectError + 5200, , _
            "Sheet KKP lokal " & namaSheetLokal & _
            " tidak ditemukan."

    End If

    daftarKunci = KunciNilaiUntukKKPExport( _
        taxKey, _
        kodeEWS)

    daftarNilai = BacaNilaiKKPLokalExport( _
        wsLokal, _
        bulan, _
        tahun, _
        daftarKunci)

    Set wsTarget = CariSheetTargetKKPExport( _
        wbEWS, _
        kodeEWS, _
        teksTeknik)

    If wsTarget Is Nothing Then

        Err.Raise _
            vbObjectError + 5201, , _
            "Target EWS untuk KKP " & kodeKKP & _
            " tidak ditemukan."

    End If

    daftarAlamat = CariAlamatTargetKKPExport( _
        wsTarget, _
        daftarKunci)

    'Satu jenis nilai tidak boleh menunjuk sel yang sama dengan jenis
    'nilai lain. Pengaman ini mencegah PPN menimpa DPP Nilai Lain.
    ValidasiAlamatTargetUnikExport _
        daftarAlamat, _
        kodeKKP

    If UBound(daftarNilai) <> UBound(daftarAlamat) Then

        Err.Raise _
            vbObjectError + 5202, , _
            "Jumlah nilai sumber dan target berbeda untuk " & _
            kodeKKP & "."

    End If

    For i = LBound(daftarNilai) To UBound(daftarNilai)

        rencana.Add Array( _
            wsTarget.Name, _
            CStr(daftarAlamat(i)), _
            CDbl(daftarNilai(i)), _
            kodeKKP, _
            CStr(daftarKunci(i)), _
            "ANGKA")

    Next i

    'Jika hasil berasal dari child sheet teknik pemeriksaan langsung,
    'atur dropdown pada sheet induk agar formula Menurut Pemeriksa
    'mengambil hasil dari teknik yang baru diisi.
    If Len(Trim$(teksTeknik)) > 0 Then

        TambahRencanaDropdownTeknikExport _
            wbEWS.Worksheets(kodeEWS), _
            teksTeknik, _
            kodeKKP, _
            rencana

    End If

End Sub

Private Sub TambahRencanaDropdownTeknikExport( _
    ByVal wsInduk As Worksheet, _
    ByVal teksTeknik As String, _
    ByVal kodeKKP As String, _
    ByVal rencana As Collection)

    Dim area As Range
    Dim sel As Range
    Dim selLabel As Range
    Dim selDropdown As Range
    Dim teks As String

    Set area = wsInduk.UsedRange

    For Each sel In area.Cells

        teks = NormalisasiTeksExport( _
            NilaiTeksAmanExport(sel.Value2))

        If InStr( _
            1, _
            teks, _
            "TEKNIKPENGUJIANYANGDIGUNAKAN", _
            vbTextCompare) > 0 Then

            Set selLabel = SelKiriAtasMergeExport(sel)

            Set selDropdown = wsInduk.Cells( _
                selLabel.MergeArea.Row + _
                    selLabel.MergeArea.rows.Count, _
                selLabel.MergeArea.Column)

            Set selDropdown = _
                SelKiriAtasMergeExport(selDropdown)

            rencana.Add Array( _
                wsInduk.Name, _
                selDropdown.Address(False, False), _
                teksTeknik, _
                kodeKKP, _
                "TEKNIK_PENGUJIAN", _
                "TEKS")

            Exit Sub

        End If

    Next sel

    Err.Raise _
        vbObjectError + 5220, , _
        "Dropdown di bawah 'Teknik Pengujian yang digunakan' " & _
        "tidak ditemukan pada " & wsInduk.Name & "."

End Sub

'---------------------------------------------------------------------
' JENIS NILAI PER KKP
'---------------------------------------------------------------------
Private Function TargetPPNExportOpsional(ByVal kodeEWS As String) As Boolean
    Dim kodeRingkas As String

    kodeRingkas = Replace$( _
        NormalisasiKodeKKPEWS(kodeEWS), _
        "-", _
        vbNullString)

    Select Case UCase$(kodeRingkas)
        Case "N5", "N9", "N10", "N23", "N25"
            TargetPPNExportOpsional = True
    End Select
End Function

Private Function TargetPPNExportDilewati(ByVal kodeEWS As String) As Boolean
    Dim kodeRingkas As String

    kodeRingkas = Replace$( _
        NormalisasiKodeKKPEWS(kodeEWS), _
        "-", _
        vbNullString)

    TargetPPNExportDilewati = (UCase$(kodeRingkas) = "N23")
End Function

Private Function KunciNilaiUntukKKPExport( _
    ByVal taxKey As String, _
    ByVal kodeKKP As String) As Variant

    Dim kodeRingkas As String

    kodeRingkas = Replace( _
        NormalisasiKodeKKPEWS(kodeKKP), _
        "-", _
        vbNullString)

    If UCase$(Trim$(taxKey)) = TAXKEY_PPH21 Then

        Select Case UCase$(kodeRingkas)

            Case "B1"

                'B-1 memuat dua nilai: DPP/PKP dan PPh terutang.
                KunciNilaiUntukKKPExport = Array( _
                    "DPP", _
                    "PPH")

            Case "P3"

                'P-3 EWS memiliki dua total: Tax Base dan Tax Income.
                'KKP lokal menyimpan PPh yang telah dipotong/dipungut,
                'sehingga target harus kolom Tax Income, bukan Tax Base.
                KunciNilaiUntukKKPExport = Array("PPH")

            Case Else

                'C, N, O, P-1 dan P-2 masing-masing satu nilai jumlah.
                KunciNilaiUntukKKPExport = Array("JUMLAH")

        End Select

        Exit Function

    End If

    Select Case UCase$(kodeRingkas)

        Case "N2", "N3", "N5", "N6", "N7", _
             "N8", "N9", "N20", "N22", "N24"

            KunciNilaiUntukKKPExport = Array( _
                "DPP", _
                "DPP_LAIN", _
                "PPN")

        Case "N4", "N19", "N21"

            KunciNilaiUntukKKPExport = Array( _
                "DPP", _
                "PPN")

        Case "N1"

            KunciNilaiUntukKKPExport = Array("DPP")

        Case "N10", "N25"

            KunciNilaiUntukKKPExport = Array("HARGA_JUAL")

        Case "N23", "O2", "Q2"

            KunciNilaiUntukKKPExport = Array("PPN")

        Case "O1", "O3", "O4", _
             "P1", "P2", "P3", "P4", "Q1"

            KunciNilaiUntukKKPExport = Array("JUMLAH")

        Case Else

            KunciNilaiUntukKKPExport = Array("NILAI")

    End Select

End Function

'---------------------------------------------------------------------
' MEMBACA SUMBER DARI KKP LOKAL
'---------------------------------------------------------------------
Private Function BacaNilaiKKPLokalExport( _
    ByVal ws As Worksheet, _
    ByVal bulan As Long, _
    ByVal tahun As Long, _
    ByVal daftarKunci As Variant) As Variant

    Dim barisPeriode As Long
    Dim daftarKolom As Variant
    Dim hasil() As Double
    Dim nilai As Variant
    Dim i As Long

    ws.Calculate

    barisPeriode = CariBarisPeriodeLokalExport( _
        ws, _
        bulan, _
        tahun)

    If barisPeriode = 0 Then

        Err.Raise _
            vbObjectError + 5300, , _
            "Periode " & Format$(bulan, "00") & "/" & _
            CStr(tahun) & " tidak ditemukan pada " & ws.Name & "."

    End If

    daftarKolom = CariKolomPemeriksaLokalExport( _
        ws, _
        barisPeriode, _
        daftarKunci)

    'Pastikan DPP, DPP Nilai Lain, dan PPN tidak terbaca dari kolom
    'yang sama akibat kemiripan teks header.
    ValidasiKolomSumberUnikExport _
        daftarKolom, _
        ws.Name

    ReDim hasil( _
        LBound(daftarKunci) To UBound(daftarKunci))

    For i = LBound(daftarKunci) To UBound(daftarKunci)

        nilai = ws.Cells( _
            barisPeriode, _
            CLng(daftarKolom(i))).Value2

        If IsError(nilai) Then

            Err.Raise _
                vbObjectError + 5301, , _
                "Nilai " & CStr(daftarKunci(i)) & _
                " pada " & ws.Name & " mengalami error."

        End If

        If Len(Trim$(CStr(nilai))) = 0 Or _
           Not IsNumeric(nilai) Then

            Err.Raise _
                vbObjectError + 5302, , _
                "Nilai " & CStr(daftarKunci(i)) & _
                " pada " & ws.Name & " bukan angka."

        End If

        hasil(i) = CDbl(nilai)

    Next i

    BacaNilaiKKPLokalExport = hasil

End Function

Private Function CariBarisPeriodeLokalExport( _
    ByVal ws As Worksheet, _
    ByVal bulan As Long, _
    ByVal tahun As Long) As Long

    Dim area As Range
    Dim sel As Range
    Dim barisMulai As Long
    Dim barisAkhir As Long
    Dim kolomAwal As Long
    Dim kolomAkhir As Long
    Dim r As Long
    Dim c As Long

    Set area = ws.UsedRange

    barisMulai = area.Row
    barisAkhir = area.Row + area.rows.Count - 1
    kolomAwal = area.Column
    kolomAkhir = area.Column + area.Columns.Count - 1

    'Lewati metadata di bagian atas dengan mencari header Uraian.
    For Each sel In area.Cells

        If NormalisasiTeksExport( _
            NilaiTeksAmanExport(sel.Value2)) = "URAIAN" Then

            barisMulai = sel.Row + 1
            Exit For

        End If

    Next sel

    'Periode biasanya berada dekat sisi kiri tabel.
    If kolomAkhir > kolomAwal + 5 Then
        kolomAkhir = kolomAwal + 5
    End If

    For r = barisMulai To barisAkhir

        For c = kolomAwal To kolomAkhir

            If SelSesuaiPeriodeExport( _
                    ws.Cells(r, c), _
                    bulan, _
                    tahun) Then

                CariBarisPeriodeLokalExport = r
                Exit Function

            End If

        Next c

    Next r

End Function

Private Function SelSesuaiPeriodeExport( _
    ByVal sel As Range, _
    ByVal bulan As Long, _
    ByVal tahun As Long) As Boolean

    Dim nilai As Variant
    Dim teks As String

    nilai = sel.Value

    If IsError(nilai) Then Exit Function

    If IsDate(nilai) Then

        SelSesuaiPeriodeExport = _
            Month(CDate(nilai)) = bulan And _
            Year(CDate(nilai)) = tahun

        Exit Function

    End If

    teks = UCase$(Trim$(CStr(nilai)))

    If Len(teks) = 0 Then Exit Function

    SelSesuaiPeriodeExport = _
        InStr(1, teks, NamaBulanIndonesiaExport(bulan), _
            vbTextCompare) > 0 And _
        InStr(1, teks, CStr(tahun), vbTextCompare) > 0

End Function

Private Function CariKolomPemeriksaLokalExport( _
    ByVal ws As Worksheet, _
    ByVal barisPeriode As Long, _
    ByVal daftarKunci As Variant) As Variant

    Dim area As Range
    Dim sel As Range
    Dim hasil() As Long
    Dim jumlahKunci As Long
    Dim kolomMulai As Long
    Dim barisHeader As Long
    Dim kandidatKolom As Long
    Dim i As Long
    Dim r As Long
    Dim c As Long
    Dim teks As String

    Set area = ws.UsedRange
    jumlahKunci = UBound(daftarKunci) - _
        LBound(daftarKunci) + 1

    ReDim hasil( _
        LBound(daftarKunci) To UBound(daftarKunci))

    If jumlahKunci = 1 Then

        'Ambil label PEMERIKSA yang paling dekat dengan tabel data.
        For Each sel In area.Cells

            If sel.Row < barisPeriode Then

                teks = NormalisasiTeksExport( _
                    NilaiTeksAmanExport(sel.Value2))

                If InStr(1, teks, "PEMERIKSA", _
                    vbTextCompare) > 0 Then

                    If sel.Row >= barisHeader Then
                        barisHeader = sel.Row
                        kandidatKolom = sel.Column
                    End If

                End If

            End If

        Next sel

        If kandidatKolom = 0 Then

            Err.Raise _
                vbObjectError + 5310, , _
                "Kolom Menurut Pemeriksa tidak ditemukan pada " & _
                ws.Name & "."

        End If

        hasil(LBound(hasil)) = kandidatKolom
        CariKolomPemeriksaLokalExport = hasil
        Exit Function

    End If

    'Untuk beberapa nilai, cari header gabungan Menurut Pemeriksa.
    For Each sel In area.Cells

        If sel.Row < barisPeriode Then

            teks = NormalisasiTeksExport( _
                NilaiTeksAmanExport(sel.Value2))

            If InStr(1, teks, "MENURUTPEMERIKSA", _
                vbTextCompare) > 0 Then

                If sel.MergeCells Then
                    kolomMulai = sel.MergeArea.Column
                Else
                    kolomMulai = sel.Column
                End If

                barisHeader = sel.Row
                Exit For

            End If

        End If

    Next sel

    If kolomMulai = 0 Then

        Err.Raise _
            vbObjectError + 5311, , _
            "Header Menurut Pemeriksa tidak ditemukan pada " & _
            ws.Name & "."

    End If

    For i = LBound(daftarKunci) To UBound(daftarKunci)

        kandidatKolom = 0

        For r = barisHeader To barisHeader + 3

            For c = kolomMulai To _
                kolomMulai + jumlahKunci + 3

                If HeaderSesuaiKunciExport( _
                        ws.Cells(r, c).Value2, _
                        CStr(daftarKunci(i))) Then

                    kandidatKolom = c
                    Exit For

                End If

            Next c

            If kandidatKolom > 0 Then Exit For

        Next r

        If kandidatKolom = 0 Then
            kandidatKolom = kolomMulai + _
                (i - LBound(daftarKunci))
        End If

        hasil(i) = kandidatKolom

    Next i

    CariKolomPemeriksaLokalExport = hasil

End Function

'---------------------------------------------------------------------
' EKSPOR RINCIAN FAKTUR PPN
'---------------------------------------------------------------------
Private Sub EksporDaftarFakturPPNKeEWS( _
    ByVal wbEWS As Workbook, ByVal daftarKKP As Variant, _
    ByVal bulan As Long, ByVal tahun As Long)

    Dim kodeKKP As Variant, kodeEWS As String
    Dim wsSumber As Worksheet, wsTarget As Worksheet
    Dim teksTeknik As String

    For Each kodeKKP In daftarKKP
        kodeEWS = NormalisasiKodeKKPEWS(CStr(kodeKKP))
        If Not TargetPPNExportDilewati(kodeEWS) Then
            Set wsSumber = SheetSumberDetailPPNExport(kodeEWS)
            If Not wsSumber Is Nothing Then
                Set wsTarget = CariSheetTargetKKPExport(wbEWS, kodeEWS, teksTeknik)
                If Not wsTarget Is Nothing Then
                    TulisDetailFakturPPNKeEWSV2 wsSumber, wsTarget, kodeEWS, bulan, tahun
                End If
            End If
        End If
    Next kodeKKP
End Sub

Private Sub TulisDetailFakturPPNKeEWS( _
    ByVal wsSumber As Worksheet, ByVal wsTarget As Worksheet, _
    ByVal kodeEWS As String, ByVal bulan As Long, ByVal tahun As Long)

    Dim hs As Long, ht As Long, mulai As Long, total As Long
    Dim kapasitas As Long, jumlah As Long, rs As Long, rt As Long
    Dim cs As Object, ct As Object

    hs = CariHeaderSumberDetailPPNV2(wsSumber)
    ht = CariHeaderDetailPPN(wsTarget, False)
    If hs = 0 Or ht = 0 Then Exit Sub
    Set cs = PetaKolomDetailPPN(wsSumber, hs)
    Set ct = PetaKolomDetailPPN(wsTarget, ht)
    If Not cs.Exists("MASA") Then Exit Sub

    'Template EWS: judul kolom, subjudul kolom, lalu baris nomor (1)-(12).
    'Data selalu dimulai pada baris setelah tiga lapis header tersebut.
    mulai = ht + 3
    total = CariBarisTotalDetailPPN(wsTarget, mulai)
    If total = 0 Then total = BarisTerakhirIsiDetailPPN(wsTarget) + 1
    kapasitas = total - mulai

    For rs = hs + 1 To BarisTerakhirIsiDetailPPN(wsSumber)
        If BarisFakturPPNSesuaiMasa(wsSumber, rs, cs, bulan, tahun) Then jumlah = jumlah + 1
    Next rs

    If jumlah > kapasitas Then
        wsTarget.rows(total).Resize(jumlah - kapasitas).Insert _
            Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        total = total + jumlah - kapasitas
    End If

    If total > mulai Then
        wsTarget.Range(wsTarget.Cells(mulai, 1), _
            wsTarget.Cells(total - 1, wsTarget.UsedRange.Column + wsTarget.UsedRange.Columns.Count - 1)).ClearContents
    End If

    rt = mulai
    For rs = hs + 1 To BarisTerakhirIsiDetailPPN(wsSumber)
        If BarisFakturPPNSesuaiMasa(wsSumber, rs, cs, bulan, tahun) Then
            SalinSatuBarisDetailPPN wsSumber, wsTarget, rs, rt, cs, ct
            rt = rt + 1
        End If
    Next rs
    If rt <= total - 1 Then
        wsTarget.Range(wsTarget.Cells(rt, 1), _
            wsTarget.Cells(total - 1, wsTarget.UsedRange.Column + wsTarget.UsedRange.Columns.Count - 1)).ClearContents
    End If
End Sub

Private Function CariHeaderSumberDetailPPNV2(ByVal ws As Worksheet) As Long
    Dim r As Long, c As Long, akhirR As Long, akhirC As Long
    Dim adaMasa As Boolean, adaNomor As Boolean, t As String

    akhirR = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    akhirC = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1

    For r = ws.UsedRange.Row To akhirR
        adaMasa = False
        adaNomor = False
        For c = ws.UsedRange.Column To akhirC
            t = NormalisasiTeksExport(NilaiTeksAmanExport(ws.Cells(r, c).Value2))
            If InStr(1, t, "MASA", vbTextCompare) > 0 Then adaMasa = True
            If InStr(1, t, "NOMOR", vbTextCompare) > 0 Or InStr(1, t, "NUMBER", vbTextCompare) > 0 Then adaNomor = True
        Next c

        If adaMasa Then
            For c = 1 To 2
                Dim r2 As Long
                r2 = r + c
                If r2 > akhirR Then Exit For
                If Application.CountIf(ws.Rows(r2), "*Nomor*") > 0 Or _
                   Application.CountIf(ws.Rows(r2), "*Number*") > 0 Then adaNomor = True
            Next c
        End If

        If adaMasa And adaNomor Then
            CariHeaderSumberDetailPPNV2 = r
            Exit Function
        End If
    Next r
End Function

Private Function CariHeaderDetailPPN(ByVal ws As Worksheet, ByVal sumber As Boolean) As Long
    Dim r As Long, c As Long, akhirR As Long, akhirC As Long
    Dim teks As String, adaMasa As Boolean, adaNomor As Boolean

    akhirR = ws.UsedRange.Row + ws.UsedRange.rows.Count - 1
    akhirC = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
    For r = ws.UsedRange.Row To akhirR
        adaMasa = False
        adaNomor = False
        For c = ws.UsedRange.Column To akhirC
            teks = NormalisasiTeksExport( _
                NilaiTeksAmanExport(ws.Cells(r, c).Value2))
            If sumber Then
                If InStr(1, teks, "MASA", vbTextCompare) > 0 Then adaMasa = True
                If InStr(1, teks, "NOMOR", vbTextCompare) > 0 Or _
                   InStr(1, teks, "NUMBER", vbTextCompare) > 0 Then adaNomor = True
            ElseIf InStr(1, teks, "TAXINVOICE", vbTextCompare) > 0 Or _
                   InStr(1, teks, "SELLINGPRICE", vbTextCompare) > 0 Or _
                   InStr(1, teks, "HARGAJUAL", vbTextCompare) > 0 Then
                CariHeaderDetailPPN = r
                Exit Function
            End If
        Next c
        If sumber And adaMasa And adaNomor Then
            CariHeaderDetailPPN = r
            Exit Function
        End If
    Next r
End Function

Private Function PetaKolomDetailPPN(ByVal ws As Worksheet, ByVal baris As Long) As Object
    Dim d As Object, c As Long, r As Long, akhir As Long, akhirBaris As Long, t As String
    Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = vbTextCompare
    akhir = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
    akhirBaris = baris + 2
    For r = baris To akhirBaris
    For c = ws.UsedRange.Column To akhir
        t = NormalisasiTeksExport(NilaiTeksAmanExport(ws.Cells(r, c).Value2))
        If t = "NO" Or t = "NO." Then d("NO") = c
        If InStr(1, t, "MASA", vbTextCompare) > 0 Then d("MASA") = c
        If InStr(1, t, "TAHUN", vbTextCompare) > 0 Then d("TAHUN") = c
        If InStr(1, t, "BUYER", vbTextCompare) > 0 Or InStr(1, t, "SELLER", vbTextCompare) > 0 Or InStr(1, t, "NAME", vbTextCompare) > 0 Or InStr(1, t, "NAMA", vbTextCompare) > 0 Then d("NAMA") = c
        If InStr(1, t, "TIN", vbTextCompare) > 0 Or InStr(1, t, "NPWP", vbTextCompare) > 0 Or InStr(1, t, "NATIONALID", vbTextCompare) > 0 Then d("TIN") = c
        If InStr(1, t, "TRANSACTIONCODE", vbTextCompare) > 0 Then d("KODE") = c
        If t = "TYPE" Or InStr(1, t, "TYPE", vbTextCompare) > 0 Then d("TIPE") = c
        If (InStr(1, t, "SERIALNUMBER", vbTextCompare) > 0 Or _
            InStr(1, t, "TAXINVOICENUMBER", vbTextCompare) > 0 Or _
            InStr(1, t, "NOMOR", vbTextCompare) > 0) And _
           InStr(1, t, "REPLACED", vbTextCompare) = 0 And _
           InStr(1, t, "RETURNED", vbTextCompare) = 0 Then d("NOMOR") = c
        If InStr(1, t, "DATE", vbTextCompare) > 0 Or InStr(1, t, "TANGGAL", vbTextCompare) > 0 Then d("TANGGAL") = c
        If InStr(1, t, "SELLINGPRICE", vbTextCompare) > 0 Or InStr(1, t, "HARGAJUAL", vbTextCompare) > 0 Then d("DPP") = c
        If InStr(1, t, "OTHERTAXBASE", vbTextCompare) > 0 Or InStr(1, t, "DPPNILAILAIN", vbTextCompare) > 0 Then d("DPP_LAIN") = c
        If t = "VAT" Or InStr(1, t, "VAT", vbTextCompare) > 0 Or _
           InStr(1, t, "PPN", vbTextCompare) > 0 Then d("PPN") = c
        If InStr(1, t, "BARANG", vbTextCompare) > 0 Or InStr(1, t, "JASA", vbTextCompare) > 0 Or InStr(1, t, "DETAIL", vbTextCompare) > 0 Then d("DETAIL") = c
        If InStr(1, t, "KETERANGAN", vbTextCompare) > 0 Or InStr(1, t, "REFERENCE", vbTextCompare) > 0 Then d("KET") = c
    Next c
    Next r
    Set PetaKolomDetailPPN = d
End Function

Private Function BarisFakturPPNSesuaiMasa(ByVal ws As Worksheet, ByVal r As Long, ByVal kol As Object, ByVal bulan As Long, ByVal tahun As Long) As Boolean
    Dim m As Variant, y As Variant
    m = ws.Cells(r, CLng(kol("MASA"))).Value2
    If Not IsNumeric(m) Or CLng(m) <> bulan Then Exit Function
    If kol.Exists("TAHUN") Then
        y = ws.Cells(r, CLng(kol("TAHUN"))).Value2
        If IsNumeric(y) And CLng(y) <> tahun Then Exit Function
    End If
    If Len(Trim$(CStr(ws.Cells(r, 1).Value2))) = 0 Then Exit Function
    BarisFakturPPNSesuaiMasa = True
End Function

Private Sub SalinSatuBarisDetailPPN(ByVal wsSumber As Worksheet, ByVal wsTarget As Worksheet, ByVal rs As Long, ByVal rt As Long, ByVal sumber As Object, ByVal target As Object)
    Dim k As Variant, v As Variant
    For Each k In Array("NAMA", "TIN", "MASA", "KODE", "TIPE", "NOMOR", "TANGGAL", "DPP", "DPP_LAIN", "PPN", "DETAIL", "KET")
        If sumber.Exists(CStr(k)) And target.Exists(CStr(k)) Then
            v = wsSumber.Cells(rs, CLng(sumber(CStr(k)))).Value2
            wsTarget.Cells(rt, CLng(target(CStr(k)))).Value2 = v
        End If
    Next k
    If target.Exists("NO") Then wsTarget.Cells(rt, CLng(target("NO"))).Value2 = rt
    If target.Exists("DPP") Then wsTarget.Cells(rt, CLng(target("DPP"))).NumberFormat = "#,##0;(#,##0);-"
    If target.Exists("DPP_LAIN") Then wsTarget.Cells(rt, CLng(target("DPP_LAIN"))).NumberFormat = "#,##0;(#,##0);-"
    If target.Exists("PPN") Then wsTarget.Cells(rt, CLng(target("PPN"))).NumberFormat = "#,##0;(#,##0);-"
End Sub

Private Function CariBarisTotalDetailPPN(ByVal ws As Worksheet, ByVal mulai As Long) As Long
    Dim r As Long, c As Long, akhirR As Long, akhirC As Long, t As String
    akhirR = ws.UsedRange.Row + ws.UsedRange.rows.Count - 1
    akhirC = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
    For r = mulai To akhirR
        For c = ws.UsedRange.Column To akhirC
            t = NormalisasiTeksExport(NilaiTeksAmanExport(ws.Cells(r, c).Value2))
            If InStr(1, t, "MENURUTPEMERIKSA", vbTextCompare) > 0 Then CariBarisTotalDetailPPN = r: Exit Function
        Next c
    Next r
End Function

Private Function BarisTerakhirIsiDetailPPN(ByVal ws As Worksheet) As Long
    Dim sel As Range
    On Error Resume Next
    Set sel = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    On Error GoTo 0
    If sel Is Nothing Then BarisTerakhirIsiDetailPPN = 1 Else BarisTerakhirIsiDetailPPN = sel.Row
End Function

'---------------------------------------------------------------------
' MENCARI SHEET TARGET EWS
'---------------------------------------------------------------------
Private Function CariSheetTargetKKPExport( _
    ByVal wb As Workbook, _
    ByVal kodeEWS As String, _
    ByRef teksTeknikTerpilih As String) As Worksheet

    Dim wsInduk As Worksheet
    Dim ws As Worksheet
    Dim area As Range
    Dim sel As Range
    Dim kandidat As String
    Dim c As Long
    Dim akhirKolom As Long
    Dim jumlahKandidat As Long

    If Not SheetAdaPadaWorkbookExport(wb, kodeEWS) Then
        Exit Function
    End If

    Set wsInduk = wb.Worksheets(kodeEWS)
    Set area = wsInduk.UsedRange

    'Prioritas pertama: kode child sheet yang berada pada baris yang
    'sama dengan teknik pemeriksaan langsung. Baris tersebut juga harus
    'memiliki nomor pengujian di sebelah kiri. Syarat ini membedakan
    'baris pilihan aktif (misalnya N18/U18 atau N17/U17) dari daftar
    'template pada kolom A:B.
    For Each sel In area.Cells

        If InStr( _
            1, _
            NormalisasiTeksExport( _
                NilaiTeksAmanExport(sel.Value2)), _
            TEKS_TEKNIK_LANGSUNG, _
            vbTextCompare) > 0 Then

            If BarisMemilikiNomorPengujianExport( _
                    wsInduk, _
                    sel.Row, _
                    sel.Column) Then

                akhirKolom = area.Column + _
                    area.Columns.Count - 1

                'Referensi child sheet berada di sebelah kanan nama
                'teknik pada baris yang sama.
                For c = sel.Column + 1 To akhirKolom

                    kandidat = Trim$( _
                        NilaiTeksAmanExport( _
                            wsInduk.Cells(sel.Row, c).Value2))

                    If Len(kandidat) > Len(kodeEWS) And _
                       StrComp( _
                           Left$(kandidat, Len(kodeEWS) + 1), _
                           kodeEWS & "-", _
                           vbTextCompare) = 0 Then

                        If SheetAdaPadaWorkbookExport( _
                                wb, _
                                kandidat) Then

                            Set CariSheetTargetKKPExport = _
                                wb.Worksheets(kandidat)

                            teksTeknikTerpilih = Trim$( _
                                NilaiTeksAmanExport( _
                                    sel.Value2))

                            Exit Function

                        End If

                    End If

                Next c

            End If

        End If

    Next sel

    'Fallback aman: child sheet harus memiliki baris target sekaligus
    'judul yang menyebut teknik pemeriksaan langsung. Jika lebih dari
    'satu kandidat ditemukan, jangan memilih secara sembarang.
    For Each ws In wb.Worksheets

        If Len(ws.Name) > Len(kodeEWS) And _
           Left$(ws.Name, Len(kodeEWS) + 1) = kodeEWS & "-" Then

            If SheetMemilikiBarisTargetExport(ws) And _
               InStr( _
                   1, _
                   NormalisasiTeksExport( _
                       NilaiTeksAmanExport( _
                           ws.Range("G7").Value2)), _
                   TEKS_TEKNIK_LANGSUNG, _
                   vbTextCompare) > 0 Then

                jumlahKandidat = jumlahKandidat + 1
                Set CariSheetTargetKKPExport = ws

            End If

        End If

    Next ws

    If jumlahKandidat = 1 Then

        teksTeknikTerpilih = _
            CariTeksTeknikChildExport( _
                wsInduk, _
                CariSheetTargetKKPExport.Name)

        Exit Function

    End If

    If jumlahKandidat > 1 Then

        Set CariSheetTargetKKPExport = Nothing

        Err.Raise _
            vbObjectError + 5390, , _
            "Target pemeriksaan langsung untuk " & kodeEWS & _
            " ambigu: ditemukan " & CStr(jumlahKandidat) & _
            " child sheet."

    End If

    'Beberapa KKP O-Q menempatkan hasil langsung pada sheet induk.
    If SheetMemilikiBarisTargetExport(wsInduk) Then
        Set CariSheetTargetKKPExport = wsInduk
    End If

End Function

Private Function CariTeksTeknikChildExport( _
    ByVal wsInduk As Worksheet, _
    ByVal namaChild As String) As String

    Dim area As Range
    Dim sel As Range
    Dim c As Long
    Dim teks As String

    Set area = wsInduk.UsedRange

    For Each sel In area.Cells

        If StrComp( _
            Trim$(NilaiTeksAmanExport(sel.Value2)), _
            namaChild, _
            vbTextCompare) = 0 Then

            For c = sel.Column - 1 To area.Column Step -1

                teks = Trim$(NilaiTeksAmanExport( _
                    wsInduk.Cells(sel.Row, c).Value2))

                If InStr( _
                    1, _
                    NormalisasiTeksExport(teks), _
                    TEKS_TEKNIK_LANGSUNG, _
                    vbTextCompare) > 0 Then

                    CariTeksTeknikChildExport = teks
                    Exit Function

                End If

            Next c

        End If

    Next sel

End Function

Private Function BarisMemilikiNomorPengujianExport( _
    ByVal ws As Worksheet, _
    ByVal nomorBaris As Long, _
    ByVal kolomTeknik As Long) As Boolean

    Dim kolomAwal As Long
    Dim c As Long
    Dim nilai As Variant

    kolomAwal = ws.UsedRange.Column

    For c = kolomTeknik - 1 To kolomAwal Step -1

        nilai = ws.Cells(nomorBaris, c).Value2

        If Not IsError(nilai) Then

            If IsNumeric(nilai) Then

                If CDbl(nilai) >= 1 And CDbl(nilai) <= 99 Then
                    BarisMemilikiNomorPengujianExport = True
                    Exit Function
                End If

            End If

        End If

    Next c

End Function

Private Function SheetMemilikiBarisTargetExport( _
    ByVal ws As Worksheet) As Boolean

    Dim kolomLabel As Long
    Dim jenisBaris As String

    SheetMemilikiBarisTargetExport = _
        CariBarisTargetEWSExport( _
            ws, _
            kolomLabel, _
            jenisBaris) > 0

End Function

'---------------------------------------------------------------------
' MENCARI ALAMAT TARGET PADA SHEET EWS
'---------------------------------------------------------------------
Private Function CariAlamatTargetKKPExport( _
    ByVal ws As Worksheet, _
    ByVal daftarKunci As Variant) As Variant

    Dim barisTarget As Long
    Dim kolomLabel As Long
    Dim jenisBaris As String
    Dim jumlahKunci As Long
    Dim hasil() As String
    Dim kolomTarget As Long
    Dim fallback As Variant
    Dim i As Long

    barisTarget = CariBarisTargetEWSExport( _
        ws, _
        kolomLabel, _
        jenisBaris)

    If barisTarget = 0 Then

        Err.Raise _
            vbObjectError + 5400, , _
            "Baris Menurut Pemeriksa/Total tidak ditemukan pada " & _
            ws.Name & "."

    End If

    jumlahKunci = UBound(daftarKunci) - _
        LBound(daftarKunci) + 1

    ReDim hasil( _
        LBound(daftarKunci) To UBound(daftarKunci))

    If jumlahKunci = 1 Then

        'O-1-2 dan O-4-2 memiliki empat blok merge setelah label
        'Total/Jumlah. Nilai harus masuk ke blok keempat:
        'X:AB, AC:AH, AI:AN, lalu AO:AY.
        If StrComp( _
                ws.Name, _
                "O-1-2", _
                vbTextCompare) = 0 Or _
           StrComp( _
                ws.Name, _
                "O-4-2", _
                vbTextCompare) = 0 Then

            kolomTarget = _
                CariKolomBlokMergeSetelahLabelExport( _
                    ws, _
                    barisTarget, _
                    kolomLabel, _
                    4)

        Else

            'Nilai tunggal lainnya dicari berdasarkan judul kolom.
            kolomTarget = CariKolomTargetHeaderExport( _
                ws, _
                barisTarget, _
                kolomLabel, _
                CStr(daftarKunci(LBound(daftarKunci))))

        End If

        'Fallback hanya dipakai untuk template khusus yang tidak
        'menempatkan judul nilai dekat baris hasil, misalnya N-23-1.
        If kolomTarget = 0 Then

            kolomTarget = CariKolomNilaiKananExport( _
                ws, _
                barisTarget, _
                kolomLabel)

        End If

        If kolomTarget = 0 Then

            Err.Raise _
                vbObjectError + 5401, , _
                "Sel nilai target tidak ditemukan pada " & _
                ws.Name & "."

        End If

        hasil(LBound(hasil)) = _
            SelKiriAtasMergeExport( _
                ws.Cells(barisTarget, kolomTarget)) _
                .Address(False, False)

        CariAlamatTargetKKPExport = hasil
        Exit Function

    End If

    For i = LBound(daftarKunci) To UBound(daftarKunci)

        kolomTarget = CariKolomTargetHeaderExport( _
            ws, _
            barisTarget, _
            kolomLabel, _
            CStr(daftarKunci(i)))

        If kolomTarget = 0 Then Exit For

        hasil(i) = SelKiriAtasMergeExport( _
            ws.Cells(barisTarget, kolomTarget)) _
            .Address(False, False)

    Next i

    If i <= UBound(daftarKunci) Then

        fallback = CariAlamatTargetBerurutanExport( _
            ws, _
            barisTarget, _
            kolomLabel, _
            jumlahKunci)

        For i = LBound(hasil) To UBound(hasil)
            hasil(i) = CStr(fallback(i))
        Next i

    End If

    CariAlamatTargetKKPExport = hasil

End Function

Private Function CariKolomBlokMergeSetelahLabelExport( _
    ByVal ws As Worksheet, _
    ByVal nomorBaris As Long, _
    ByVal kolomLabel As Long, _
    ByVal urutanBlok As Long) As Long

    Dim selLabel As Range
    Dim selBlok As Range
    Dim kolomAkhir As Long
    Dim c As Long
    Dim nomorBlok As Long

    Set selLabel = SelKiriAtasMergeExport( _
        ws.Cells(nomorBaris, kolomLabel))

    If selLabel.MergeCells Then
        c = selLabel.MergeArea.Column + _
            selLabel.MergeArea.Columns.Count
    Else
        c = selLabel.Column + 1
    End If

    kolomAkhir = ws.UsedRange.Column + _
        ws.UsedRange.Columns.Count - 1

    Do While c <= kolomAkhir

        Set selBlok = SelKiriAtasMergeExport( _
            ws.Cells(nomorBaris, c))

        nomorBlok = nomorBlok + 1

        If nomorBlok = urutanBlok Then

            CariKolomBlokMergeSetelahLabelExport = _
                selBlok.Column

            Exit Function

        End If

        If selBlok.MergeCells Then
            c = selBlok.MergeArea.Column + _
                selBlok.MergeArea.Columns.Count
        Else
            c = selBlok.Column + 1
        End If

    Loop

    Err.Raise _
        vbObjectError + 5411, , _
        "Blok merge target ke-" & CStr(urutanBlok) & _
        " tidak ditemukan pada " & ws.Name & "."

End Function

Private Function CariBarisTargetEWSExport( _
    ByVal ws As Worksheet, _
    ByRef kolomLabel As Long, _
    ByRef jenisBaris As String) As Long

    Dim area As Range
    Dim sel As Range
    Dim teks As String

    Set area = ws.UsedRange

    'Prioritaskan Menurut Pemeriksa dibanding Total.
    For Each sel In area.Cells

        teks = NormalisasiTeksExport( _
            NilaiTeksAmanExport(sel.Value2))

        If teks = "MENURUTPEMERIKSA" Or _
           InStr(1, teks, "JUMLAHMENURUTPEMERIKSA", _
                vbTextCompare) > 0 Then

            If BarisAdaNilaiDiKananExport( _
                    ws, _
                    sel.Row, _
                    sel.Column) Then

                CariBarisTargetEWSExport = sel.Row
                kolomLabel = sel.Column
                jenisBaris = "PEMERIKSA"
                Exit Function

            End If

        End If

    Next sel

    For Each sel In area.Cells

        teks = NormalisasiTeksExport( _
            NilaiTeksAmanExport(sel.Value2))

        If teks = "TOTAL" Then

            If BarisAdaNilaiDiKananExport( _
                    ws, _
                    sel.Row, _
                    sel.Column) Then

                CariBarisTargetEWSExport = sel.Row
                kolomLabel = sel.Column
                jenisBaris = "TOTAL"
                Exit Function

            End If

        End If

    Next sel

End Function

Private Function BarisAdaNilaiDiKananExport( _
    ByVal ws As Worksheet, _
    ByVal nomorBaris As Long, _
    ByVal kolomLabel As Long) As Boolean

    Dim kolomAkhir As Long
    Dim c As Long
    Dim sel As Range

    kolomAkhir = ws.UsedRange.Column + _
        ws.UsedRange.Columns.Count - 1

    For c = kolomLabel + 1 To kolomAkhir

        Set sel = SelKiriAtasMergeExport( _
            ws.Cells(nomorBaris, c))

        If sel.Column = c Then

            If sel.HasFormula Or _
               (Not IsError(sel.Value2) And _
                IsNumeric(sel.Value2)) Then

                BarisAdaNilaiDiKananExport = True
                Exit Function

            End If

        End If

    Next c

End Function

Private Function CariKolomNilaiKananExport( _
    ByVal ws As Worksheet, _
    ByVal nomorBaris As Long, _
    ByVal kolomLabel As Long) As Long

    Dim kolomAkhir As Long
    Dim c As Long
    Dim sel As Range

    kolomAkhir = ws.UsedRange.Column + _
        ws.UsedRange.Columns.Count - 1

    For c = kolomLabel + 1 To kolomAkhir

        Set sel = SelKiriAtasMergeExport( _
            ws.Cells(nomorBaris, c))

        If sel.Column = c Then

            If sel.HasFormula Or _
               (Not IsError(sel.Value2) And _
                IsNumeric(sel.Value2)) Then

                CariKolomNilaiKananExport = c
                Exit Function

            End If

        End If

    Next c

End Function

Private Function CariKolomTargetHeaderExport( _
    ByVal ws As Worksheet, _
    ByVal barisTarget As Long, _
    ByVal kolomLabel As Long, _
    ByVal kunci As String) As Long

    Dim kolomAkhir As Long
    Dim barisAwal As Long
    Dim r As Long
    Dim c As Long
    Dim selTarget As Range

    kolomAkhir = ws.UsedRange.Column + _
        ws.UsedRange.Columns.Count - 1

    'P-3 PPh 21 menempatkan judul Tax Income jauh di atas baris Total.
    'Perluasan pencarian hanya berlaku untuk kunci PPH agar mapping PPN
    'yang telah stabil tetap menggunakan jarak header semula.
    If UCase$(Trim$(kunci)) = "PPH" Then
        barisAwal = barisTarget - 25
    Else
        barisAwal = barisTarget - 6
    End If
    If barisAwal < 1 Then barisAwal = 1

    For r = barisTarget - 1 To barisAwal Step -1

        For c = kolomLabel + 1 To kolomAkhir

            If HeaderSesuaiKunciExport( _
                    ws.Cells(r, c).Value2, _
                    kunci) Then

                Set selTarget = SelKiriAtasMergeExport( _
                    ws.Cells(barisTarget, c))

                If selTarget.HasFormula Or _
                   (Not IsError(selTarget.Value2) And _
                    IsNumeric(selTarget.Value2)) Then

                    CariKolomTargetHeaderExport = _
                        selTarget.Column

                    Exit Function

                End If

            End If

        Next c

    Next r

End Function

Private Function CariAlamatTargetBerurutanExport( _
    ByVal ws As Worksheet, _
    ByVal barisTarget As Long, _
    ByVal kolomLabel As Long, _
    ByVal jumlahDibutuhkan As Long) As Variant

    Dim hasil() As String
    Dim kolomAkhir As Long
    Dim jumlah As Long
    Dim c As Long
    Dim sel As Range

    ReDim hasil(0 To jumlahDibutuhkan - 1)

    kolomAkhir = ws.UsedRange.Column + _
        ws.UsedRange.Columns.Count - 1

    For c = kolomLabel + 1 To kolomAkhir

        Set sel = SelKiriAtasMergeExport( _
            ws.Cells(barisTarget, c))

        If sel.Column = c Then

            If sel.HasFormula Or _
               (Not IsError(sel.Value2) And _
                IsNumeric(sel.Value2)) Then

                hasil(jumlah) = sel.Address(False, False)
                jumlah = jumlah + 1

                If jumlah = jumlahDibutuhkan Then Exit For

            End If

        End If

    Next c

    If jumlah <> jumlahDibutuhkan Then

        Err.Raise _
            vbObjectError + 5410, , _
            "Jumlah sel target pada " & ws.Name & _
            " tidak sesuai. Ditemukan " & CStr(jumlah) & _
            ", diperlukan " & CStr(jumlahDibutuhkan) & "."

    End If

    CariAlamatTargetBerurutanExport = hasil

End Function

Private Function HeaderSesuaiKunciExport( _
    ByVal nilaiHeader As Variant, _
    ByVal kunci As String) As Boolean

    Dim teks As String

    If IsError(nilaiHeader) Then Exit Function

    teks = NormalisasiTeksExport(CStr(nilaiHeader))

    Select Case kunci

        Case "DPP_LAIN"

            HeaderSesuaiKunciExport = _
                InStr(1, teks, "DPPNILAILAIN", _
                    vbTextCompare) > 0 Or _
                InStr(1, teks, "DPPLAIN", _
                    vbTextCompare) > 0

        Case "DPP"

            HeaderSesuaiKunciExport = _
                InStr(1, teks, "DPP", _
                    vbTextCompare) > 0 And _
                InStr(1, teks, "NILAILAIN", _
                    vbTextCompare) = 0 And _
                InStr(1, teks, "DPPLAIN", _
                    vbTextCompare) = 0

        Case "PPN"

            'PPN tidak boleh cocok dengan DPP Nilai Lain. Setelah
            'normalisasi, DPP Nilai Lain menjadi DPPNILAILAIN dan
            'mengandung rangkaian huruf PPN pada batas DPP + Nilai.
            HeaderSesuaiKunciExport = _
                InStr(1, teks, "PPN", _
                    vbTextCompare) > 0 And _
                InStr(1, teks, "DPP", _
                    vbTextCompare) = 0

        Case "PPH"

            'Mendukung judul Indonesia dan Inggris pada EWS PPh 21.
            'Tax Base sengaja tidak dianggap PPh agar P-3 tidak masuk
            'ke total pertama (BL) ketika target yang benar adalah
            'Tax Income (BX).
            HeaderSesuaiKunciExport = _
                (InStr(1, teks, "PPH", _
                    vbTextCompare) > 0 Or _
                 InStr(1, teks, "TAXINCOME", _
                    vbTextCompare) > 0 Or _
                 InStr(1, teks, "PAJAKPENGHASILAN", _
                    vbTextCompare) > 0) And _
                InStr(1, teks, "DPP", _
                    vbTextCompare) = 0 And _
                InStr(1, teks, "TAXBASE", _
                    vbTextCompare) = 0

        Case "JUMLAH"

            HeaderSesuaiKunciExport = _
                InStr(1, teks, "JUMLAH", _
                    vbTextCompare) > 0

        Case "HARGA_JUAL"

            HeaderSesuaiKunciExport = _
                InStr(1, teks, "HARGAJUAL", _
                    vbTextCompare) > 0 Or _
                InStr(1, teks, "PENGGANTIAN", _
                    vbTextCompare) > 0

    End Select

End Function

'---------------------------------------------------------------------
' VALIDASI KEUNIKAN SUMBER DAN TARGET
'---------------------------------------------------------------------
Private Sub ValidasiKolomSumberUnikExport( _
    ByVal daftarKolom As Variant, _
    ByVal namaSheet As String)

    Dim i As Long
    Dim j As Long

    For i = LBound(daftarKolom) To UBound(daftarKolom) - 1

        For j = i + 1 To UBound(daftarKolom)

            If CLng(daftarKolom(i)) = _
               CLng(daftarKolom(j)) Then

                Err.Raise _
                    vbObjectError + 5420, , _
                    "Mapping kolom sumber pada " & namaSheet & _
                    " tidak unik. Ekspor dibatalkan agar nilai " & _
                    "tidak tertukar."

            End If

        Next j

    Next i

End Sub

Private Sub ValidasiAlamatTargetUnikExport( _
    ByVal daftarAlamat As Variant, _
    ByVal kodeKKP As String)

    Dim i As Long
    Dim j As Long

    For i = LBound(daftarAlamat) To UBound(daftarAlamat) - 1

        For j = i + 1 To UBound(daftarAlamat)

            If StrComp( _
                CStr(daftarAlamat(i)), _
                CStr(daftarAlamat(j)), _
                vbTextCompare) = 0 Then

                Err.Raise _
                    vbObjectError + 5421, , _
                    "Mapping sel target untuk " & kodeKKP & _
                    " tidak unik. Ekspor dibatalkan agar nilai " & _
                    "tidak saling menimpa."

            End If

        Next j

    Next i

End Sub

Private Function SelKiriAtasMergeExport( _
    ByVal sel As Range) As Range

    If sel.MergeCells Then
        Set SelKiriAtasMergeExport = sel.MergeArea.Cells(1, 1)
    Else
        Set SelKiriAtasMergeExport = sel
    End If

End Function

'---------------------------------------------------------------------
' UTILITAS WORKBOOK DAN NAMA FILE
'---------------------------------------------------------------------
Private Function SheetAdaPadaWorkbookExport( _
    ByVal wb As Workbook, _
    ByVal namaSheet As String) As Boolean

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = wb.Worksheets(namaSheet)
    On Error GoTo 0

    SheetAdaPadaWorkbookExport = Not ws Is Nothing

End Function

Private Function WorkbookTerbukaUntukExport( _
    ByVal filePath As String) As Boolean

    Dim wb As Workbook

    For Each wb In Application.Workbooks

        If StrComp( _
            wb.FullName, _
            filePath, _
            vbTextCompare) = 0 Then

            WorkbookTerbukaUntukExport = True
            Exit Function

        End If

    Next wb

End Function

Private Function AmbilNPWPDariNamaFileExport( _
    ByVal filePath As String) As String

    Dim bagian As Variant
    Dim token As Variant
    Dim teks As String

    bagian = Split(Dir$(filePath), "_")

    For Each token In bagian

        teks = CStr(token)

        If Len(teks) = 16 And IsNumeric(teks) Then
            AmbilNPWPDariNamaFileExport = teks
            Exit Function
        End If

    Next token

End Function

' Menghasilkan namespace sheet lokal dari identifier jenis file EWS.
' Nilai kosong berarti jenis file belum didukung oleh engine ini.
Private Function TaxKeyDariKodeJenisEngineExport( _
    ByVal kodeJenis As String) As String

    Select Case UCase$(Trim$(kodeJenis))

        Case KODE_PPN_DN_ENGINE
            TaxKeyDariKodeJenisEngineExport = TAXKEY_PPN_DN

        Case KODE_PPH21_ENGINE
            TaxKeyDariKodeJenisEngineExport = TAXKEY_PPH21

        Case KODE_PPH21F_ENGINE
            TaxKeyDariKodeJenisEngineExport = TAXKEY_PPH21F

        Case KODE_PPH23_ENGINE
            TaxKeyDariKodeJenisEngineExport = TAXKEY_PPH23

        Case KODE_PPH26_ENGINE
            TaxKeyDariKodeJenisEngineExport = TAXKEY_PPH26

        Case KODE_PPH42_ENGINE
            TaxKeyDariKodeJenisEngineExport = TAXKEY_PPH42

    End Select

End Function

Private Sub TungguPerhitunganEWS(ByVal wb As Workbook)

    Dim mulai As Date

    mulai = Now

    'Workbook tidak memiliki metode Calculate yang konsisten pada
    'seluruh versi Excel. Gunakan Application.Calculate agar tidak
    'menimbulkan runtime error 438.
    Application.Calculate

    Do

        DoEvents

        If Application.CalculationState = xlDone Then Exit Do

        If DateDiff("s", mulai, Now) >= _
            BATAS_TUNGGU_DETIK Then Exit Do

    Loop

End Sub

Private Function NilaiTeksAmanExport( _
    ByVal nilai As Variant) As String

    If IsError(nilai) Then
        NilaiTeksAmanExport = vbNullString
    Else
        NilaiTeksAmanExport = CStr(nilai)
    End If

End Function

Private Function NormalisasiTeksExport( _
    ByVal teks As String) As String

    teks = UCase$(Trim$(teks))
    teks = Replace(teks, " ", vbNullString)
    teks = Replace(teks, vbCr, vbNullString)
    teks = Replace(teks, vbLf, vbNullString)
    teks = Replace(teks, ".", vbNullString)
    teks = Replace(teks, ",", vbNullString)
    teks = Replace(teks, "/", vbNullString)
    teks = Replace(teks, "-", vbNullString)
    teks = Replace(teks, "(", vbNullString)
    teks = Replace(teks, ")", vbNullString)
    teks = Replace(teks, ":", vbNullString)

    NormalisasiTeksExport = teks

End Function

Private Function NamaBulanIndonesiaExport( _
    ByVal bulan As Long) As String

    Dim daftar As Variant

    daftar = Array( _
        "", "JANUARI", "FEBRUARI", "MARET", _
        "APRIL", "MEI", "JUNI", "JULI", _
        "AGUSTUS", "SEPTEMBER", "OKTOBER", _
        "NOVEMBER", "DESEMBER")

    If bulan >= 1 And bulan <= 12 Then
        NamaBulanIndonesiaExport = daftar(bulan)
    End If

End Function

'---------------------------------------------------------------------
' LOG EKSPOR PADA WORKBOOK EWSSYNC
'---------------------------------------------------------------------
Private Sub CatatLogExportSelektif( _
    ByVal filePath As String, _
    ByVal bulan As Long, _
    ByVal tahun As Long, _
    ByVal namaJenisPajak As String, _
    ByVal daftarKKP As String, _
    ByVal status As String, _
    ByVal keterangan As String)

    Dim ws As Worksheet
    Dim baris As Long

    On Error GoTo AbaikanLog

    If SheetAdaPadaWorkbookExport( _
            ThisWorkbook, _
            NAMA_SHEET_LOG_EXPORT) Then

        Set ws = ThisWorkbook.Worksheets( _
            NAMA_SHEET_LOG_EXPORT)

    Else

        Set ws = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets( _
                ThisWorkbook.Worksheets.Count))

        ws.Name = NAMA_SHEET_LOG_EXPORT

        ws.Range("A1:H1").Value = Array( _
            "Waktu", _
            "File EWS", _
            "Bulan", _
            "Tahun", _
            "Jenis Pajak", _
            "KKP", _
            "Status", _
            "Keterangan")

        ws.Range("A1:H1").Font.Bold = True

    End If

    baris = ws.Cells(ws.rows.Count, 1).End(xlUp).Row + 1

    ws.Cells(baris, 1).Value = Now
    ws.Cells(baris, 2).Value = filePath
    ws.Cells(baris, 3).Value = bulan
    ws.Cells(baris, 4).Value = tahun
    ws.Cells(baris, 5).Value = namaJenisPajak
    ws.Cells(baris, 6).Value = Replace(daftarKKP, "|", ", ")
    ws.Cells(baris, 7).Value = status
    ws.Cells(baris, 8).Value = keterangan

    ws.Columns("A:H").EntireColumn.AutoFit

AbaikanLog:

End Sub

'=====================================================================
' PPh26: target Total + header semantik; dua KJS divalidasi terpisah.
' Macro startup EWS tidak diperlukan untuk ekspor dan sudah dinonaktifkan
' sebelum Workbooks.Open melalui msoAutomationSecurityForceDisable.
'=====================================================================
Private Sub TambahRencanaPPh26Export( _
    ByVal wb As Workbook, ByVal kode As String, _
    ByVal bulan As Long, ByVal tahun As Long, ByVal rencana As Collection)

    Dim lokal As Worksheet, tujuan As Worksheet
    Dim kunci As Variant, nilai As Variant, alamat As Variant
    Dim teknik As String, daftar As Variant, item As Variant
    Dim dikenal As Boolean, i As Long

    daftar = DaftarKodeEWSPPh26()
    For Each item In daftar
        If CStr(item) = kode Then dikenal = True
    Next item
    If Not dikenal Then Err.Raise vbObjectError + 5700, , _
        "Kode KKP PPh 26 tidak didukung: " & kode

    Set lokal = DapatkanSheetLokalKKP(TAXKEY_PPH26, kode, False)
    If lokal Is Nothing Then Err.Raise vbObjectError + 5701, , _
        "Sheet lokal tidak ditemukan: " & NamaSheetLokalKKP(TAXKEY_PPH26, kode)

    Select Case kode
        Case "B-1", "B-2": kunci = Array("DPP", "PPH")
        Case "C-1": kunci = Array("JUMLAH")
        Case "P-3": kunci = Array("PPH")
        Case Else: kunci = Array("JUMLAH")
    End Select
    nilai = BacaSumberPPh23Export(lokal, bulan, tahun, kunci)

    If kode = "B-1" Or kode = "B-2" Or kode = "C-1" Then
        Set tujuan = CariSheetTargetKKPExport(wb, kode, teknik)
        If tujuan Is Nothing Then Err.Raise vbObjectError + 5702, , _
            "Child teknik langsung tidak ditemukan: " & kode
        'Sheet -3-1 adalah rincian transaksi; ekspor memakai rekap -3.
        If tujuan.Name <> kode & "-3" Or Len(teknik) = 0 Then
            Err.Raise vbObjectError + 5703, , _
                "Referensi teknik langsung tidak sesuai rekap " & kode & "-3."
        End If
    Else
        If Not SheetAdaPadaWorkbookExport(wb, kode) Then _
            Err.Raise vbObjectError + 5704, , "Sheet EWS tidak ditemukan: " & kode
        Set tujuan = wb.Worksheets(kode)
    End If

    alamat = AlamatTotalPPh23Export(tujuan, kunci)
    ValidasiAlamatTargetUnikExport alamat, kode
    For i = LBound(kunci) To UBound(kunci)
        rencana.Add Array(tujuan.Name, CStr(alamat(i)), CDbl(nilai(i)), _
            kode, CStr(kunci(i)), "ANGKA")
    Next i
    If Len(teknik) > 0 Then
        TambahRencanaDropdownTeknikExport wb.Worksheets(kode), teknik, kode, rencana
    End If
End Sub

'Pra-write: pastikan file benar-benar template TOT09 dengan kedua KJS,
'seluruh target unik/writable, serta dropdown teknik berupa validation list.
Private Sub ValidasiRencanaPPh26Export(ByVal wb As Workbook, ByVal rencana As Collection)
    Dim seen As Object, item As Variant, sel As Range, k As String, id As Variant
    Dim wsMain As Worksheet, r As Long, akhir As Long
    Dim ids As Object, teksID As String, wajib As Variant

    Set wsMain = wb.Worksheets("MAIN LEVEL")
    Set ids = CreateObject("Scripting.Dictionary")
    ids.CompareMode = vbTextCompare
    akhir = wsMain.UsedRange.Row + wsMain.UsedRange.rows.Count - 1
    For r = 1 To akhir
        teksID = NormalisasiIdentifierPPh23Export(wsMain.Cells(r, 1).Value2)
        If Len(teksID) > 0 Then ids(teksID) = r
    Next r

    wajib = Array( _
        "NIL_DPP_KJS_411127_100", "NIL_DPP_KJS_411127_110", _
        "NIL_TOTAL_DPP_PASAL_26", "NIL_PPH_KJS_411127_100", _
        "NIL_PPH_KJS_411127_110", "NIL_TOTAL_PPH_PASAL26", _
        "NIL_PH_TDK_TERHUTANG_PPH_PASAL_17_AYAT_2", _
        "NIL_PPH_HARUS_DIPUNGUT_DISETOR", _
        "NIL_KREDIT_PAJAK_DITANGGUNG_PEMERINTAH", _
        "NIL_KREDIT_PAJAK_SETORAN_MASA", "NIL_KREDIT_PAJAK_STP", _
        "NIL_KOMPENSASI_DARI_MASA_PAJAK", "NIL_KREDIT_PAJAK_LAIN", _
        "NIL_KREDIT_PAJAK_UANG_TEBUSAN", _
        "NIL_KREDIT_PAJAK_UNGKAP_HARTA_BERSIH", _
        "NIL_KOMPENSASI_KE_MASA_PAJAK", "NIL_DIPERHITUNGKAN_SKPKB", _
        "NIL_DIPERHITUNGKAN_SKPKBT", "NIL_DIPERHITUNGKAN_SKPLB", _
        "NIL_DIPERHITUNGKAN", "NIL_PPH_PASAL17_DIBAYAR_NPWP_SENDIRI", _
        "NIL_PPH_PASAL17_DIBAYAR_NPWP_LAIN", _
        "NIL_PPH_PASAL17_TELAH_POTPUT", "NIL_PPH_TDK_TERUTANG_PASAL17", _
        "NIL_KREDIT_PAJAK", "NIL_KLB")

    For Each id In wajib
        If Not ids.Exists(CStr(id)) Then Err.Raise vbObjectError + 5730, , _
            "Identifier wajib PPh26 belum ditemukan pada MAIN LEVEL kolom A: " & CStr(id)
    Next id

    Set seen = CreateObject("Scripting.Dictionary")
    For Each item In rencana
        Set sel = SelKiriAtasMergeExport( _
            wb.Worksheets(CStr(item(0))).Range(CStr(item(1))))
        k = sel.Worksheet.Name & "!" & sel.Address
        If seen.Exists(k) Then Err.Raise vbObjectError + 5731, , "Target ganda: " & k
        seen.Add k, True
        If sel.HasArray Then Err.Raise vbObjectError + 5732, , _
            "Target merupakan formula array: " & k
        If sel.Worksheet.ProtectContents Then
            If sel.Locked Then Err.Raise vbObjectError + 5733, , "Target terkunci: " & k
        End If
        If CStr(item(5)) = "TEKS" Then
            If Not DropdownListPPh23Export(sel) Then Err.Raise vbObjectError + 5734, , _
                "Validasi list dropdown teknik tidak ditemukan: " & k
        End If
    Next item
End Sub

'=====================================================================
' PPh21 Final: KKP biasa memakai total; B-1-3 memakai tabel horizontal.
'=====================================================================
Private Sub TambahRencanaPPh21FExport( _
    ByVal wb As Workbook, ByVal kode As String, _
    ByVal bulan As Long, ByVal tahun As Long, ByVal rencana As Collection)

    Dim lokal As Worksheet
    Dim tujuan As Worksheet
    Dim kunci As Variant
    Dim nilai As Variant
    Dim alamat As Variant
    Dim teknik As String
    Dim daftar As Variant
    Dim item As Variant
    Dim dikenal As Boolean
    Dim i As Long

    daftar = DaftarKodeEWSPPh21F()
    For Each item In daftar
        If CStr(item) = kode Then dikenal = True
    Next item

    If Not dikenal Then Err.Raise vbObjectError + 5750, , _
        "Kode KKP PPh 21 Final tidak didukung: " & kode

    Set lokal = DapatkanSheetLokalKKP(TAXKEY_PPH21F, kode, False)
    If lokal Is Nothing Then Err.Raise vbObjectError + 5751, , _
        "Sheet lokal tidak ditemukan: " & _
        NamaSheetLokalKKP(TAXKEY_PPH21F, kode)

    If kode = "B-1" Then
        If Not SheetAdaPadaWorkbookExport(wb, "B-1-3") Then _
            Err.Raise vbObjectError + 5752, , _
                "Sheet EWS B-1-3 tidak ditemukan."

        kunci = Array("DPP", "PPH")
        nilai = BacaSumberPPh23Export(lokal, bulan, tahun, kunci)

        rencana.Add Array( _
            "B-1-3", "BA21", CDbl(nilai(0)), _
            "B-1", "DPP", "ANGKA")
        rencana.Add Array( _
            "B-1-3", "BF21", CDbl(nilai(1)), _
            "B-1", "PPH", "ANGKA")

        TambahDropdownB1PPh21FSekali wb, rencana
        Exit Sub
    End If

    If kode = "B-1-3" Then
        TambahRencanaDetailPPh21FExport lokal, bulan, tahun, wb, rencana
        TambahDropdownB1PPh21FSekali wb, rencana
        Exit Sub
    End If

    If kode = "P-3" Then
        kunci = Array("PPH")
    Else
        kunci = Array("JUMLAH")
    End If

    nilai = BacaSumberPPh23Export(lokal, bulan, tahun, kunci)

    If kode = "C-1" Then
        Set tujuan = CariSheetTargetKKPExport(wb, kode, teknik)
        If tujuan Is Nothing Then Err.Raise vbObjectError + 5752, , _
            "Child teknik langsung tidak ditemukan: " & kode
        If tujuan.Name <> "C-1-3" Or Len(teknik) = 0 Then
            Err.Raise vbObjectError + 5753, , _
                "Referensi teknik langsung C-1 harus menuju C-1-3."
        End If
    Else
        If Not SheetAdaPadaWorkbookExport(wb, kode) Then _
            Err.Raise vbObjectError + 5754, , _
                "Sheet EWS tidak ditemukan: " & kode
        Set tujuan = wb.Worksheets(kode)
    End If

    alamat = AlamatTotalPPh23Export(tujuan, kunci)
    ValidasiAlamatTargetUnikExport alamat, kode

    For i = LBound(kunci) To UBound(kunci)
        rencana.Add Array( _
            tujuan.Name, CStr(alamat(i)), CDbl(nilai(i)), _
            kode, CStr(kunci(i)), "ANGKA")
    Next i

    If Len(teknik) > 0 Then
        TambahRencanaDropdownTeknikExport _
            wb.Worksheets(kode), teknik, kode, rencana
    End If

End Sub

'Tujuan: Menulis delapan nilai B-1-3 ke sepuluh baris detail EWS.
Private Sub TambahRencanaDetailPPh21FExport( _
    ByVal lokal As Worksheet, ByVal bulan As Long, ByVal tahun As Long, _
    ByVal wb As Workbook, ByVal rencana As Collection)

    Dim daftarKode As Variant
    Dim daftarDPP As Variant
    Dim daftarPPh As Variant
    Dim i As Long
    Dim barisTarget As Long
    Dim punyaNilai As Boolean

    If Not SheetAdaPadaWorkbookExport(wb, "B-1-3-1") Then
        Err.Raise vbObjectError + 5755, , _
            "Sheet EWS B-1-3-1 tidak ditemukan."
    End If

    BacaDetailLokalPPh21FExport _
        lokal, bulan, tahun, daftarKode, daftarDPP, daftarPPh

    For i = 0 To 3
        If CDbl(daftarDPP(i)) <> 0 Or CDbl(daftarPPh(i)) <> 0 Then
            punyaNilai = True
            Exit For
        End If
    Next i

    'Tabel detail kosong tidak boleh menghapus rincian yang sudah ada.
    If Not punyaNilai Then Exit Sub

    For i = 0 To 9
        barisTarget = 20 + i

        If i <= 3 Then
            rencana.Add Array( _
                "B-1-3-1", "AU" & CStr(barisTarget), _
                CStr(daftarKode(i)), "B-1-3", "KODE_OBJEK", _
                "TEKS_BANTU")
            rencana.Add Array( _
                "B-1-3-1", "BL" & CStr(barisTarget), _
                CDbl(daftarDPP(i)), "B-1-3", "DPP", "ANGKA")
            rencana.Add Array( _
                "B-1-3-1", "BX" & CStr(barisTarget), _
                CDbl(daftarPPh(i)), "B-1-3", "PPH", "ANGKA")
        Else
            rencana.Add Array( _
                "B-1-3-1", "AU" & CStr(barisTarget), _
                vbNullString, "B-1-3", "BERSIHKAN_KODE", _
                "TEKS_BANTU")
            rencana.Add Array( _
                "B-1-3-1", "BL" & CStr(barisTarget), _
                0#, "B-1-3", "BERSIHKAN_DPP", "BANTU")
            rencana.Add Array( _
                "B-1-3-1", "BX" & CStr(barisTarget), _
                0#, "B-1-3", "BERSIHKAN_PPH", "BANTU")
        End If
    Next i

End Sub

'Tujuan: Membaca empat pasangan DPP/PPh pada kolom masa terpilih.
Private Sub BacaDetailLokalPPh21FExport( _
    ByVal ws As Worksheet, ByVal bulan As Long, ByVal tahun As Long, _
    ByRef daftarKode As Variant, ByRef daftarDPP As Variant, _
    ByRef daftarPPh As Variant)

    Dim wsInfo As Worksheet
    Dim i As Long
    Dim r As Long
    Dim kolomMasa As Long
    Dim barisKode As Long
    Dim jumlahKode As Long
    Dim nilai As Variant
    Dim teks As String

    daftarKode = Array( _
        "21-100-01", "21-100-02", _
        "21-100-03", "21-100-07")
    daftarDPP = Array(0#, 0#, 0#, 0#)
    daftarPPh = Array(0#, 0#, 0#, 0#)

    If Not SheetAdaPadaWorkbookExport(ThisWorkbook, "INFO") Then _
        Err.Raise vbObjectError + 5756, , "Sheet INFO tidak ditemukan."
    Set wsInfo = ThisWorkbook.Worksheets("INFO")

    For r = 27 To 38
        If SelSesuaiPeriodeExport(wsInfo.Cells(r, 2), bulan, tahun) Then
            If kolomMasa <> 0 Then Err.Raise vbObjectError + 5756, , _
                "Periode ganda pada tabel Data Bulan di INFO."
            kolomMasa = 4 + (r - 27)
        End If
    Next r

    If kolomMasa = 0 Then Err.Raise vbObjectError + 5757, , _
        "Periode " & Format$(bulan, "00") & "/" & CStr(tahun) & _
        " tidak ditemukan pada tabel Data Bulan di INFO."

    For i = LBound(daftarKode) To UBound(daftarKode)
        barisKode = 0
        jumlahKode = 0

        For r = 1 To ws.UsedRange.Row + ws.UsedRange.rows.Count - 1
            teks = UCase$(Trim$(NilaiTeksAmanExport(ws.Cells(r, 2).Value2)))
            If InStr(1, teks, CStr(daftarKode(i)), vbTextCompare) > 0 Then
                barisKode = r
                jumlahKode = jumlahKode + 1
            End If
        Next r

        If jumlahKode <> 1 Then Err.Raise vbObjectError + 5758, , _
            "Kode " & CStr(daftarKode(i)) & _
            " tidak ditemukan atau ganda pada " & ws.Name & "."

        If UCase$(Trim$(NilaiTeksAmanExport( _
                ws.Cells(barisKode, 3).Value2))) <> "DPP" Or _
           UCase$(Trim$(NilaiTeksAmanExport( _
                ws.Cells(barisKode + 1, 3).Value2))) <> "PPH" Then
            Err.Raise vbObjectError + 5759, , _
                "Pasangan DPP/PPh tidak sesuai untuk " & _
                CStr(daftarKode(i)) & "."
        End If

        nilai = ws.Cells(barisKode, kolomMasa).Value2
        If IsError(nilai) Or Not IsNumeric(nilai) Then _
            Err.Raise vbObjectError + 5760, , _
                "Nilai DPP bukan angka: " & CStr(daftarKode(i))
        daftarDPP(i) = CDbl(nilai)

        nilai = ws.Cells(barisKode + 1, kolomMasa).Value2
        If IsError(nilai) Or Not IsNumeric(nilai) Then _
            Err.Raise vbObjectError + 5761, , _
                "Nilai PPh bukan angka: " & CStr(daftarKode(i))
        daftarPPh(i) = CDbl(nilai)
    Next i

End Sub

'Tujuan: Memasang pilihan pemeriksaan langsung B-1 satu kali.
Private Sub TambahDropdownB1PPh21FSekali( _
    ByVal wb As Workbook, ByVal rencana As Collection)

    Dim item As Variant
    Dim ws As Worksheet
    Dim teknik As String
    Dim r As Long

    For Each item In rencana
        If CStr(item(0)) = "B-1" And _
           CStr(item(4)) = "TEKNIK_PENGUJIAN" Then Exit Sub
    Next item

    If Not SheetAdaPadaWorkbookExport(wb, "B-1") Then _
        Err.Raise vbObjectError + 5762, , "Sheet EWS B-1 tidak ditemukan."

    Set ws = wb.Worksheets("B-1")

    For r = 14 To 19
        If InStr(1, NormalisasiTeksExport( _
                NilaiTeksAmanExport(ws.Cells(r, 1).Value2)), _
                TEKS_TEKNIK_LANGSUNG, vbTextCompare) > 0 Then
            teknik = Trim$(NilaiTeksAmanExport(ws.Cells(r, 1).Value2))
            Exit For
        End If
    Next r

    If Len(teknik) = 0 Then
        Err.Raise vbObjectError + 5763, , _
            "Pilihan pemeriksaan langsung tidak ditemukan pada daftar B-1."
    End If

    rencana.Add Array( _
        "B-1", "I39", teknik, "B-1", _
        "TEKNIK_PENGUJIAN", "TEKS")

End Sub

'Tujuan: Memastikan template dan seluruh target PPh 21 Final valid.
Private Sub ValidasiRencanaPPh21FExport( _
    ByVal wb As Workbook, ByVal rencana As Collection)

    Dim seen As Object
    Dim ids As Object
    Dim wsMain As Worksheet
    Dim item As Variant
    Dim id As Variant
    Dim wajib As Variant
    Dim sel As Range
    Dim kunci As String
    Dim teksID As String
    Dim r As Long
    Dim akhir As Long

    Set wsMain = wb.Worksheets(NAMA_SHEET_MAIN_EWS_ENGINE)
    Set ids = CreateObject("Scripting.Dictionary")
    ids.CompareMode = vbTextCompare
    akhir = wsMain.UsedRange.Row + wsMain.UsedRange.rows.Count - 1

    For r = 1 To akhir
        teksID = NormalisasiIdentifierPPh23Export(wsMain.Cells(r, 1).Value2)
        If Len(teksID) > 0 Then ids(teksID) = r
    Next r

    wajib = Array( _
        "NIL_DPP_KJS_411121_100_FINAL", _
        "NIL_TOTAL_DPP_PASAL_21_FINAL", _
        "NIL_PPH_KJS_411121_100_FINAL", _
        "NIL_TOTAL_PPH_PASAL21_FINAL", _
        "NIL_PH_TDK_TERHUTANG_PPH_PASAL_17_AYAT_2", _
        "NIL_PPH_HARUS_DIPUNGUT_DISETOR", _
        "NIL_KREDIT_PAJAK_DITANGGUNG_PEMERINTAH", _
        "NIL_KREDIT_PAJAK_SETORAN_MASA", _
        "NIL_KREDIT_PAJAK_STP", "KOMPENSASI_DARI_MASA_PAJAK", _
        "NIL_KREDIT_PAJAK_LAIN", "NIL_KREDIT_PAJAK_UANG_TEBUSAN", _
        "NIL_KREDIT_PAJAK_UNGKAP_HARTA_BERSIH", _
        "KOMPENSASI_KE_MASA_PAJAK", "NIL_DIPERHITUNGKAN_SKPKB", _
        "NIL_DIPERHITUNGKAN_SKPKBT", "NIL_DIPERHITUNGKAN_SKPLB", _
        "NIL_DIPERHITUNGKAN", _
        "NIL_PPH_PASAL17_DIBAYAR_NPWP_SENDIRI", _
        "NIL_PPH_PASAL17_DIBAYAR_NPWP_LAIN", _
        "NIL_PPH_PASAL17_TELAH_POTPUT", _
        "NIL_PPH_TDK_TERUTANG_PASAL17", "NIL_KREDIT_PAJAK", "NIL_KLB")

    For Each id In wajib
        If Not ids.Exists(CStr(id)) Then Err.Raise vbObjectError + 5764, , _
            "Identifier wajib PPh 21 Final belum ditemukan: " & CStr(id)
    Next id

    If SheetAdaPadaWorkbookExport(wb, "B-1-3-1") Then
        If NormalisasiTeksExport(wb.Worksheets("B-1-3-1").Range("AU15").Value2) _
                <> "TAXOBJECTCODE" Or _
           NormalisasiTeksExport(wb.Worksheets("B-1-3-1").Range("BL15").Value2) _
                <> "TAXBASERP" Or _
           NormalisasiTeksExport(wb.Worksheets("B-1-3-1").Range("BX15").Value2) _
                <> "TAXINCOMERP" Then
            Err.Raise vbObjectError + 5765, , _
                "Struktur kolom B-1-3-1 tidak sesuai template PPh 21 Final."
        End If
    End If

    Set seen = CreateObject("Scripting.Dictionary")
    For Each item In rencana
        Set sel = SelKiriAtasMergeExport( _
            wb.Worksheets(CStr(item(0))).Range(CStr(item(1))))
        kunci = sel.Worksheet.Name & "!" & sel.Address
        If seen.Exists(kunci) Then Err.Raise vbObjectError + 5766, , _
            "Target ganda: " & kunci
        seen.Add kunci, True
        If sel.HasArray Then Err.Raise vbObjectError + 5767, , _
            "Target merupakan formula array: " & kunci
        If sel.Worksheet.ProtectContents Then
            If sel.Locked Then Err.Raise vbObjectError + 5768, , _
                "Target terkunci: " & kunci
        End If
        If CStr(item(4)) = "TEKNIK_PENGUJIAN" Then
            If Not DropdownListPPh23Export(sel) Then _
                Err.Raise vbObjectError + 5769, , _
                    "Validasi list dropdown teknik tidak ditemukan: " & kunci
        End If
    Next item

End Sub

'Tujuan: Memastikan nilai detail lokal sama dengan file EWS.
Private Sub VerifikasiDetailPPh21FExport( _
    ByVal wb As Workbook, ByVal bulan As Long, ByVal tahun As Long, _
    ByVal rencana As Collection)

    Dim item As Variant
    Dim perlu As Boolean
    Dim lokal As Worksheet
    Dim daftarKode As Variant
    Dim daftarDPP As Variant
    Dim daftarPPh As Variant
    Dim nilai As Variant
    Dim i As Long
    Dim r As Long

    For Each item In rencana
        If CStr(item(3)) = "B-1-3" Then perlu = True: Exit For
    Next item
    If Not perlu Then Exit Sub

    Set lokal = DapatkanSheetLokalKKP(TAXKEY_PPH21F, "B-1-3", False)
    If lokal Is Nothing Then Err.Raise vbObjectError + 5770, , _
        "Sheet lokal P21F-B-1-3 tidak ditemukan saat verifikasi."

    BacaDetailLokalPPh21FExport _
        lokal, bulan, tahun, daftarKode, daftarDPP, daftarPPh

    For i = 0 To 3
        r = 20 + i
        If StrComp(Trim$(CStr(wb.Worksheets("B-1-3-1").Cells(r, 47).Value2)), _
                CStr(daftarKode(i)), vbTextCompare) <> 0 Then
            Err.Raise vbObjectError + 5771, , _
                "Kode objek hasil ekspor tidak sesuai pada B-1-3-1."
        End If

        nilai = wb.Worksheets("B-1-3-1").Cells(r, 64).Value2
        If IsError(nilai) Or Not IsNumeric(nilai) Or _
           Abs(CDbl(nilai) - CDbl(daftarDPP(i))) > 0.01 Then
            Err.Raise vbObjectError + 5772, , _
                "DPP hasil ekspor tidak sesuai: " & CStr(daftarKode(i))
        End If

        nilai = wb.Worksheets("B-1-3-1").Cells(r, 76).Value2
        If IsError(nilai) Or Not IsNumeric(nilai) Or _
           Abs(CDbl(nilai) - CDbl(daftarPPh(i))) > 0.01 Then
            Err.Raise vbObjectError + 5773, , _
                "PPh hasil ekspor tidak sesuai: " & CStr(daftarKode(i))
        End If
    Next i

End Sub

'=====================================================================
' PPh42: target Total + header semantik; empat KJS divalidasi terpisah.
' Macro startup EWS tidak diperlukan untuk ekspor dan sudah dinonaktifkan
' sebelum Workbooks.Open melalui msoAutomationSecurityForceDisable.
'=====================================================================
Private Sub TambahRencanaPPh42Export( _
    ByVal wb As Workbook, ByVal kode As String, _
    ByVal bulan As Long, ByVal tahun As Long, ByVal rencana As Collection)

    Dim lokal As Worksheet, tujuan As Worksheet
    Dim kunci As Variant, nilai As Variant, alamat As Variant
    Dim teknik As String, daftar As Variant, item As Variant
    Dim dikenal As Boolean, i As Long

    daftar = DaftarKodeEWSPPh42()
    For Each item In daftar
        If CStr(item) = kode Then dikenal = True
    Next item
    If Not dikenal Then Err.Raise vbObjectError + 5800, , _
        "Kode KKP PPh 4(2) tidak didukung: " & kode

    Set lokal = DapatkanSheetLokalKKP(TAXKEY_PPH42, kode, False)
    If lokal Is Nothing Then Err.Raise vbObjectError + 5801, , _
        "Sheet lokal tidak ditemukan: " & NamaSheetLokalKKP(TAXKEY_PPH42, kode)

    Select Case kode
        Case "B-1", "B-2", "B-3", "B-4": kunci = Array("DPP", "PPH")
        Case "C-1": kunci = Array("JUMLAH")
        Case "P-3": kunci = Array("PPH")
        Case Else: kunci = Array("JUMLAH")
    End Select
    nilai = BacaSumberPPh23Export(lokal, bulan, tahun, kunci)

    If kode = "B-1" Or kode = "B-2" Or kode = "B-3" Or kode = "B-4" Or kode = "C-1" Then
        Set tujuan = CariSheetTargetKKPExport(wb, kode, teknik)
        If tujuan Is Nothing Then Err.Raise vbObjectError + 5802, , _
            "Child teknik langsung tidak ditemukan: " & kode
        'B-1/B-2/B-3 dan C-1 memakai rekap -3. Pada B-4,
        'pemeriksaan langsung berada pada B-4-4 sesuai parent X24.
        If Len(teknik) = 0 Then Err.Raise vbObjectError + 5803, , _
            "Teks teknik pemeriksaan langsung tidak ditemukan: " & kode
        If kode = "B-4" Then
            If tujuan.Name <> "B-4-4" Then Err.Raise vbObjectError + 5804, , _
                "Target teknik langsung B-4 harus B-4-4, ditemukan " & tujuan.Name
        ElseIf tujuan.Name <> kode & "-3" Then
            Err.Raise vbObjectError + 5805, , _
                "Referensi teknik langsung tidak sesuai rekap " & kode & "-3."
        End If
    Else
        If Not SheetAdaPadaWorkbookExport(wb, kode) Then _
            Err.Raise vbObjectError + 5804, , "Sheet EWS tidak ditemukan: " & kode
        Set tujuan = wb.Worksheets(kode)
    End If

    alamat = AlamatTotalPPh23Export(tujuan, kunci)
    ValidasiAlamatTargetUnikExport alamat, kode
    For i = LBound(kunci) To UBound(kunci)
        rencana.Add Array(tujuan.Name, CStr(alamat(i)), CDbl(nilai(i)), _
            kode, CStr(kunci(i)), "ANGKA")
    Next i
    If Len(teknik) > 0 Then
        TambahRencanaDropdownTeknikExport wb.Worksheets(kode), teknik, kode, rencana
    End If
End Sub

'Pra-write: pastikan file benar-benar template TOT10 dengan keempat KJS,
'seluruh target unik/writable, serta dropdown teknik berupa validation list.
Private Sub ValidasiRencanaPPh42Export(ByVal wb As Workbook, ByVal rencana As Collection)
    Dim seen As Object, item As Variant, sel As Range, k As String, id As Variant
    Dim wsMain As Worksheet, r As Long, akhir As Long
    Dim ids As Object, teksID As String, wajib As Variant

    Set wsMain = wb.Worksheets("MAIN LEVEL")
    Set ids = CreateObject("Scripting.Dictionary")
    ids.CompareMode = vbTextCompare
    akhir = wsMain.UsedRange.Row + wsMain.UsedRange.rows.Count - 1
    For r = 1 To akhir
        teksID = NormalisasiIdentifierPPh23Export(wsMain.Cells(r, 1).Value2)
        If Len(teksID) > 0 Then ids(teksID) = r
    Next r

    wajib = Array( _
        "NIL_DPP_KJS_411128_100", "NIL_DPP_KJS_411128_402", _
        "NIL_DPP_KJS_411128_403", "NIL_DPP_KJS_411128_420", _
        "NIL_TOTAL_DPP_PASAL4_AYAT2", _
        "NIL_PPH_KJS_411128_100", "NIL_PPH_KJS_411128_402", _
        "NIL_PPH_KJS_411128_403", "NIL_PPH_KJS_411128_420", _
        "NIL_TOTAL_PPH_PASAL4_AYAT2", _
        "NIL_PH_TDK_TERHUTANG_PPH_PASAL_17_AYAT_2", _
        "NIL_PPH_HARUS_DIPUNGUT_DISETOR", _
        "NIL_KREDIT_PAJAK_DITANGGUNG_PEMERINTAH", _
        "NIL_KREDIT_PAJAK_SETORAN_MASA", "NIL_KREDIT_PAJAK_STP", _
        "NIL_KOMPENSASI_DARI_MASA_PAJAK", "NIL_KREDIT_PAJAK_LAIN", _
        "NIL_KREDIT_PAJAK_UANG_TEBUSAN", _
        "NIL_KREDIT_PAJAK_UNGKAP_HARTA_BERSIH", _
        "NIL_KOMPENSASI_KE_MASA_PAJAK", "NIL_DIPERHITUNGKAN_SKPKB", _
        "NIL_DIPERHITUNGKAN_SKPKBT", "NIL_DIPERHITUNGKAN_SKPLB", _
        "NIL_DIPERHITUNGKAN", "NIL_PPH_PASAL17_DIBAYAR_NPWP_SENDIRI", _
        "NIL_PPH_PASAL17_DIBAYAR_NPWP_LAIN", _
        "NIL_PPH_PASAL17_TELAH_POTPUT", "NIL_PPH_TDK_TERUTANG_PASAL17", _
        "NIL_KREDIT_PAJAK", "NIL_KLB")

    For Each id In wajib
        If Not ids.Exists(CStr(id)) Then Err.Raise vbObjectError + 5830, , _
            "Identifier wajib PPh42 belum ditemukan pada MAIN LEVEL kolom A: " & CStr(id)
    Next id

    Set seen = CreateObject("Scripting.Dictionary")
    For Each item In rencana
        Set sel = SelKiriAtasMergeExport( _
            wb.Worksheets(CStr(item(0))).Range(CStr(item(1))))
        k = sel.Worksheet.Name & "!" & sel.Address
        If seen.Exists(k) Then Err.Raise vbObjectError + 5831, , "Target ganda: " & k
        seen.Add k, True
        If sel.HasArray Then Err.Raise vbObjectError + 5832, , _
            "Target merupakan formula array: " & k
        If sel.Worksheet.ProtectContents Then
            If sel.Locked Then Err.Raise vbObjectError + 5833, , "Target terkunci: " & k
        End If
        If CStr(item(5)) = "TEKS" Then
            If Not DropdownListPPh23Export(sel) Then Err.Raise vbObjectError + 5834, , _
                "Validasi list dropdown teknik tidak ditemukan: " & k
        End If
    Next item
End Sub

'=====================================================================
' PPh23: target Total + header semantik; tanpa fallback numerik pertama.
' C-1 sengaja menulis Total DPP child, tidak memperbaiki formula induk.
'=====================================================================
Private Sub TambahRencanaPPh23Export( _
    ByVal wb As Workbook, ByVal kode As String, _
    ByVal bulan As Long, ByVal tahun As Long, ByVal rencana As Collection)

    Dim lokal As Worksheet, tujuan As Worksheet
    Dim kunci As Variant, nilai As Variant, alamat As Variant
    Dim teknik As String, daftar As Variant, item As Variant
    Dim dikenal As Boolean, i As Long

    daftar = DaftarKodeEWSPPh23()
    For Each item In daftar
        If CStr(item) = kode Then dikenal = True
    Next item
    If Not dikenal Then Err.Raise vbObjectError + 5600, , _
        "Kode KKP PPh 23 tidak didukung: " & kode

    Set lokal = DapatkanSheetLokalKKP(TAXKEY_PPH23, kode, False)
    If lokal Is Nothing Then Err.Raise vbObjectError + 5601, , _
        "Sheet lokal tidak ditemukan: " & NamaSheetLokalKKP(TAXKEY_PPH23, kode)

    Select Case kode
        Case "B-1": kunci = Array("DPP", "PPH")
        Case "C-1": kunci = Array("DPP")
        Case "P-3": kunci = Array("PPH")
        Case Else: kunci = Array("JUMLAH")
    End Select
    nilai = BacaSumberPPh23Export(lokal, bulan, tahun, kunci)

    If kode = "B-1" Or kode = "C-1" Then
        Set tujuan = CariSheetTargetKKPExport(wb, kode, teknik)
        If tujuan Is Nothing Then Err.Raise vbObjectError + 5602, , _
            "Child teknik langsung tidak ditemukan: " & kode
        'Sheet -3-1 adalah rincian bukti potong, bukan rekap objek.
        If tujuan.Name <> kode & "-3" Or Len(teknik) = 0 Then
            Err.Raise vbObjectError + 5603, , _
                "Referensi teknik langsung tidak sesuai rekap " & kode & "-3."
        End If
    Else
        If Not SheetAdaPadaWorkbookExport(wb, kode) Then _
            Err.Raise vbObjectError + 5604, , "Sheet EWS tidak ditemukan: " & kode
        Set tujuan = wb.Worksheets(kode)
    End If

    alamat = AlamatTotalPPh23Export(tujuan, kunci)
    ValidasiAlamatTargetUnikExport alamat, kode
    For i = LBound(kunci) To UBound(kunci)
        rencana.Add Array(tujuan.Name, CStr(alamat(i)), CDbl(nilai(i)), _
            kode, CStr(kunci(i)), "ANGKA")
    Next i
    If Len(teknik) > 0 Then
        TambahRencanaDropdownTeknikExport wb.Worksheets(kode), teknik, kode, rencana
    End If
End Sub

'Menemukan satu periode dan header Pemeriksa. Tidak boleh jatuh ke WP
'atau koreksi ketika header nilai tidak ditemukan.
Private Function BacaSumberPPh23Export( _
    ByVal ws As Worksheet, ByVal bulan As Long, ByVal tahun As Long, _
    ByVal kunci As Variant) As Variant

    Dim sel As Range, uraian As Range, header As Range, kandidat As Range
    Dim periode As Object, kolom As Object, rows As Variant, cols As Variant
    Dim r As Long, c As Long, baris As Long, i As Long
    Dim hasil() As Double, v As Variant, teks As String
    Set periode = CreateObject("Scripting.Dictionary")
    For Each sel In ws.UsedRange.Cells
        If NormalisasiTeksExport(NilaiTeksAmanExport(sel.Value2)) = "URAIAN" Then
            If Not uraian Is Nothing Then Err.Raise vbObjectError + 5610, , _
                "Header Uraian ganda: " & ws.Name
            Set uraian = SelKiriAtasMergeExport(sel)
        End If
    Next sel
    If uraian Is Nothing Then Err.Raise vbObjectError + 5611, , _
        "Header Uraian tidak ditemukan: " & ws.Name
    For r = uraian.Row + uraian.MergeArea.rows.Count To _
            ws.UsedRange.Row + ws.UsedRange.rows.Count - 1
        For c = uraian.Column To uraian.Column + uraian.MergeArea.Columns.Count - 1
            If SelSesuaiPeriodeExport(ws.Cells(r, c), bulan, tahun) Then periode(CStr(r)) = r
        Next c
    Next r
    If periode.Count <> 1 Then Err.Raise vbObjectError + 5612, , _
        "Periode tidak ditemukan atau ganda pada " & ws.Name & ": " & bulan & "/" & tahun
    rows = periode.Items: baris = CLng(rows(0))
    For r = uraian.Row To baris - 1
        For c = ws.UsedRange.Column To ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
            teks = NormalisasiTeksExport(NilaiTeksAmanExport(ws.Cells(r, c).Value2))
            If teks = "MENURUTPEMERIKSARP" Or teks = "MENURUTPEMERIKSA" Or _
               teks = "PEMERIKSARP" Or teks = "PEMERIKSA" Then
                If Not header Is Nothing Then Err.Raise vbObjectError + 5613, , _
                    "Header Pemeriksa ganda: " & ws.Name
                Set header = SelKiriAtasMergeExport(ws.Cells(r, c))
            End If
        Next c
    Next r
    If header Is Nothing Then Err.Raise vbObjectError + 5614, , _
        "Header Pemeriksa tidak ditemukan: " & ws.Name
    ReDim hasil(LBound(kunci) To UBound(kunci))
    For i = LBound(kunci) To UBound(kunci)
        Set kolom = CreateObject("Scripting.Dictionary")
        If LBound(kunci) = UBound(kunci) Then
            If header.MergeArea.Columns.Count <> 1 Then Err.Raise vbObjectError + 5615, , _
                "Kolom Pemeriksa tunggal ambigu: " & ws.Name
            kolom(CStr(header.Column)) = header.Column
        Else
            For r = header.Row + 1 To baris - 1
                For c = header.Column To header.Column + header.MergeArea.Columns.Count - 1
                    If HeaderPPh23Export(ws.Cells(r, c).Value2, CStr(kunci(i))) Then kolom(CStr(c)) = c
                Next c
            Next r
        End If
        If kolom.Count <> 1 Then Err.Raise vbObjectError + 5616, , _
            "Header nilai Pemeriksa tidak unik: " & ws.Name & " / " & CStr(kunci(i))
        cols = kolom.Items
        v = ws.Cells(baris, CLng(cols(0))).Value2
        If IsError(v) Or IsEmpty(v) Or IsNull(v) Then Err.Raise vbObjectError + 5617, , _
            "Nilai Pemeriksa kosong/error: " & ws.Name
        If Not IsNumeric(v) Then Err.Raise vbObjectError + 5618, , _
            "Nilai Pemeriksa bukan angka: " & ws.Name
        hasil(i) = CDbl(v)
    Next i
    BacaSumberPPh23Export = hasil
End Function

'Header dapat jauh di atas Total (baris 15 -> 91 pada B/C).
'Hanya satu baris Total dan satu kolom per jenis nilai boleh cocok.
Private Function AlamatTotalPPh23Export( _
    ByVal ws As Worksheet, ByVal kunci As Variant) As Variant
    Dim sel As Range, label As Range, target As Range
    Dim kolom As Object, daftar As Variant, hasil() As String
    Dim r As Long, c As Long, i As Long
    For Each sel In ws.UsedRange.Cells
        If NormalisasiTeksExport(NilaiTeksAmanExport(sel.Value2)) = "TOTAL" Then
            If Not label Is Nothing Then Err.Raise vbObjectError + 5620, , _
                "Baris Total ganda: " & ws.Name
            Set label = SelKiriAtasMergeExport(sel)
        End If
    Next sel
    If label Is Nothing Then Err.Raise vbObjectError + 5621, , _
        "Baris Total tidak ditemukan: " & ws.Name
    ReDim hasil(LBound(kunci) To UBound(kunci))
    For i = LBound(kunci) To UBound(kunci)
        Set kolom = CreateObject("Scripting.Dictionary")
        For r = ws.UsedRange.Row To label.Row - 1
            For c = label.Column + label.MergeArea.Columns.Count To _
                    ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
                If HeaderPPh23Export(ws.Cells(r, c).Value2, CStr(kunci(i))) Then
                    Set target = SelKiriAtasMergeExport(ws.Cells(label.Row, c))
                    If target.Column = c Then kolom(target.Address(False, False)) = True
                End If
            Next c
        Next r
        If kolom.Count <> 1 Then Err.Raise vbObjectError + 5622, , _
            "Header target Total tidak unik: " & ws.Name & " / " & CStr(kunci(i))
        daftar = kolom.Keys: hasil(i) = CStr(daftar(0))
    Next i
    AlamatTotalPPh23Export = hasil
End Function

'Exact semantic matching: uraian panjang tidak dianggap header nilai.
Private Function HeaderPPh23Export(ByVal v As Variant, ByVal kunci As String) As Boolean
    Dim t As String
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then Exit Function
    t = NormalisasiTeksExport(CStr(v))
    t = Replace(t, "$", "")
    If Right$(t, 2) = "RP" Then t = Left$(t, Len(t) - 2)
    Select Case kunci
        Case "DPP": HeaderPPh23Export = (t = "DPP" Or t = "PKP" Or t = "DPPPKP" Or t = "PKPDPP" Or t = "TAXBASE")
        Case "PPH": HeaderPPh23Export = (t = "PPH" Or t = "PPHTERUTANG" Or t = "PAJAKPENGHASILAN" Or t = "TAXINCOME")
        Case "JUMLAH": HeaderPPh23Export = (t = "JUMLAH")
    End Select
End Function

'Pra-write: identitas template, target unik, writable, dan dropdown list.
Private Sub ValidasiRencanaPPh23Export(ByVal wb As Workbook, ByVal rencana As Collection)
    Dim seen As Object, item As Variant, sel As Range, k As String, id As Variant
    Dim wsMain As Worksheet, r As Long, akhir As Long
    Dim ids As Object, teksID As String
    Set wsMain = wb.Worksheets("MAIN LEVEL")
    Set ids = CreateObject("Scripting.Dictionary")
    ids.CompareMode = vbTextCompare
    akhir = wsMain.UsedRange.Row + wsMain.UsedRange.rows.Count - 1
    'Value2 membaca isi sel, termasuk baris/kolom tersembunyi atau terfilter.
    'Tidak memakai Find yang gagal menemukan identifier pada sesi pengguna.
    For r = 1 To akhir
        teksID = NormalisasiIdentifierPPh23Export(wsMain.Cells(r, 1).Value2)
        If Len(teksID) > 0 Then ids(teksID) = r
    Next r
    For Each id In Array("NIL_DPP_KJS_411124_100", "NIL_PPH_KJS_411124_100")
        If Not ids.Exists(CStr(id)) Then Err.Raise vbObjectError + 5630, , _
            "Identifier wajib belum ditemukan pada MAIN LEVEL kolom A: " & CStr(id) & _
            "; A16=[" & NilaiTeksAmanExport(wsMain.Range("A16").Value2) & _
            "]; A19=[" & NilaiTeksAmanExport(wsMain.Range("A19").Value2) & "]"
    Next id
    Set seen = CreateObject("Scripting.Dictionary")
    For Each item In rencana
        Set sel = SelKiriAtasMergeExport(wb.Worksheets(CStr(item(0))).Range(CStr(item(1))))
        k = sel.Worksheet.Name & "!" & sel.Address
        If seen.Exists(k) Then Err.Raise vbObjectError + 5631, , "Target ganda: " & k
        seen.Add k, True
        If sel.HasArray Then Err.Raise vbObjectError + 5632, , "Target merupakan formula array: " & k
        If sel.Worksheet.ProtectContents Then
            If sel.Locked Then Err.Raise vbObjectError + 5633, , "Target terkunci: " & k
        End If
        If CStr(item(5)) = "TEKS" Then
            If Not DropdownListPPh23Export(sel) Then Err.Raise vbObjectError + 5634, , _
                "Validasi list dropdown teknik tidak ditemukan: " & k
        End If
    Next item
End Sub

Private Function DropdownListPPh23Export(ByVal sel As Range) As Boolean
    On Error GoTo TidakAda
    DropdownListPPh23Export = (sel.Validation.Type = xlValidateList)
TidakAda:
End Function

'Verifikasi nilai yang ditulis sebelum Save, tidak memaksa roll-up C-1.
Private Sub VerifikasiHasilPPh23Export(ByVal wb As Workbook, ByVal rencana As Collection)
    Dim item As Variant, sel As Range, v As Variant, jenisTeks As Boolean
    For Each item In rencana
        Set sel = wb.Worksheets(CStr(item(0))).Range(CStr(item(1)))
        v = sel.Value2
        jenisTeks = (Left$(UCase$(CStr(item(5))), 4) = "TEKS")
        If IsError(v) Or IsNull(v) Then Err.Raise vbObjectError + 5640, , _
            "Hasil tulis kosong/error: " & sel.Worksheet.Name & "!" & sel.Address
        If IsEmpty(v) And Not (jenisTeks And Len(CStr(item(2))) = 0) Then _
            Err.Raise vbObjectError + 5640, , _
                "Hasil tulis kosong/error: " & sel.Worksheet.Name & "!" & sel.Address
        If sel.HasFormula Then Err.Raise vbObjectError + 5641, , "Target kembali menjadi formula."
        If jenisTeks Then
            If CStr(v) <> CStr(item(2)) Then Err.Raise vbObjectError + 5642, , "Teks hasil tulis tidak sesuai."
        Else
            If Not IsNumeric(v) Then Err.Raise vbObjectError + 5643, , "Hasil tulis bukan angka."
            If CDbl(v) <> CDbl(item(2)) Then Err.Raise vbObjectError + 5644, , "Hasil tulis tidak sesuai."
        End If
    Next item
End Sub

'Catatan sukses C-1: tidak ada penulisan tambahan pada AB21 atau MAIN LEVEL.
Private Function CatatanPPh23Export(ByVal wb As Workbook, ByVal rencana As Collection) As String
    Dim item As Variant, sel As Range, f As String, sesuai As Boolean, expected As String
    For Each item In rencana
        If CStr(item(3)) = "C-1" And CStr(item(5)) = "ANGKA" Then
            expected = "='" & CStr(item(0)) & "'!" & CStr(item(1))
            For Each sel In wb.Worksheets("C-1").UsedRange.Cells
                If sel.HasFormula Then
                    f = Replace(CStr(sel.Formula), "$", "")
                    If StrComp(f, expected, vbTextCompare) = 0 Then sesuai = True
                End If
            Next sel
            If Not sesuai Then CatatanPPh23Export = _
                "CATATAN C-1: nilai ditulis ke " & CStr(item(0)) & "!" & CStr(item(1)) & _
                "; referensi langsung induk ke total ini tidak ditemukan. " & _
                "Formula bawaan EWS tidak diubah; nilai induk dapat belum mengikuti total child."
            Exit Function
        End If
    Next item
End Function

 'Status global xlPending bukan berarti Excel sedang menghitung.
'Sumber telah dihitung oleh HitungUlangKKPLokal; EWS oleh TungguPerhitunganEWS.
'Jangan memicu Application.Calculate kedua kali hanya untuk menunggu status.
Private Sub TungguKalkulasiPPh23Export(ByVal filePath As String, ByVal tahap As String)
    Dim mulai As Double, durasi As Double, status As XlCalculationState
    mulai = Timer
    Do
        status = Application.CalculationState
        Select Case status
            Case xlDone
                Exit Sub
            Case xlPending
                CatatDiagnostikExport filePath, tahap, _
                    "INFO: status global xlPending; kalkulasi eksplisit sudah dipanggil. " & _
                    "Lanjut dengan validasi sumber/target dan verifikasi hasil tulis."
                Exit Sub
            Case xlCalculating
                durasi = Timer - mulai
                If durasi < 0 Then durasi = durasi + 86400#
                If durasi >= 30# Then Err.Raise vbObjectError + 5650, , _
                    "Excel masih aktif menghitung (xlCalculating) setelah 30 detik; " & _
                    "tahap: " & tahap
            Case Else
                Err.Raise vbObjectError + 5651, , _
                    "Status kalkulasi Excel tidak dikenal: " & CStr(status)
        End Select
        DoEvents
    Loop
End Sub

'Log teks tidak membutuhkan form atau objek worksheet Excel.
Private Sub CatatDiagnostikExport(ByVal filePath As String, _
    ByVal tahap As String, ByVal pesan As String)
    Dim nomor As Integer, terbuka As Boolean
    On Error GoTo SelesaiLog
    nomor = FreeFile
    Open Environ$("TEMP") & "\EWSSYNC_Export_diagnostic.log" For Append As #nomor
    terbuka = True
    Print #nomor, Format$(Now, "yyyy-mm-dd hh:nn:ss") & " | " & filePath & _
        " | " & tahap & " | " & Replace(Replace(pesan, vbCr, " "), vbLf, " ")
SelesaiLog:
    On Error Resume Next
    If terbuka Then Close #nomor
End Sub

'Hanya menormalkan kapitalisasi dan whitespace, bukan angka KJS/pemisahnya.
Private Function NormalisasiIdentifierPPh23Export(ByVal nilai As Variant) As String
    Dim teks As String
    If IsError(nilai) Or IsNull(nilai) Or IsEmpty(nilai) Then Exit Function
    teks = CStr(nilai)
    teks = Replace(teks, ChrW(160), " ")
    teks = Replace(teks, vbCr, "")
    teks = Replace(teks, vbLf, "")
    teks = Replace(teks, vbTab, "")
    NormalisasiIdentifierPPh23Export = UCase$(Trim$(teks))
End Function

' PATCH EWSSYNC v0.55: sumber daftar faktur PPN berada pada sheet -2.

'V2: ekspor daftar faktur dengan posisi kolom template EWS yang tetap.
Private Sub TulisDetailFakturPPNKeEWSV2(ByVal wsSumber As Worksheet, ByVal wsTarget As Worksheet, ByVal kodeEWS As String, ByVal bulan As Long, ByVal tahun As Long)
    Dim hs As Long, akhir As Long, r As Long, rt As Long, jumlah As Long
    Dim batasTotal As Long, kapasitas As Long, nomor As Long, masa As Variant
    Dim sumber As Object, target As Object
    Dim calcLama As XlCalculation, eventsLama As Boolean, screenLama As Boolean
    On Error GoTo Gagal
    calcLama = Application.Calculation: eventsLama = Application.EnableEvents: screenLama = Application.ScreenUpdating
    Application.Calculation = xlCalculationManual: Application.EnableEvents = False: Application.ScreenUpdating = False
    hs = CariHeaderSumberDetailPPNV2(wsSumber)
    If hs = 0 Then Err.Raise vbObjectError + 6101, , "Header sumber tidak ditemukan: " & wsSumber.Name
    Set sumber = PetaKolomDetailPPN(wsSumber, hs)
    If Not sumber.Exists("MASA") Then Err.Raise vbObjectError + 6102, , "Kolom Masa tidak ditemukan: " & wsSumber.Name
    Set target = KolomTargetDetailPPNV2(wsTarget)
    If Not target.Exists("NO") Then Err.Raise vbObjectError + 6103, , "Template target tidak dikenali: " & wsTarget.Name
    If wsTarget.Name = "N-19-2" Then PastikanKolomDetailBarangN19 wsTarget
    akhir = BarisTerakhirIsiDetailPPN(wsSumber)
    For r = hs + 1 To akhir
        masa = wsSumber.Cells(r, CLng(sumber("MASA"))).Value2
        If IsNumeric(masa) Then
            If CLng(masa) = bulan And Len(Trim$(CStr(wsSumber.Cells(r, 1).Value2))) > 0 Then
                jumlah = jumlah + 1
            End If
        End If
    Next r
    batasTotal = CariBarisTotalDetailPPN(wsTarget, 23)
    If batasTotal = 0 Then Err.Raise vbObjectError + 6104, , "Baris total target tidak ditemukan: " & wsTarget.Name
    kapasitas = batasTotal - 23
    If jumlah > kapasitas Then
        wsTarget.Rows(batasTotal).Resize(jumlah - kapasitas).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        batasTotal = batasTotal + jumlah - kapasitas
    End If
    wsTarget.Range(wsTarget.Cells(23, 1), wsTarget.Cells(batasTotal - 1, wsTarget.UsedRange.Column + wsTarget.UsedRange.Columns.Count - 1)).ClearContents
    rt = 23: nomor = 1
    For r = hs + 1 To akhir
        masa = wsSumber.Cells(r, CLng(sumber("MASA"))).Value2
        If IsNumeric(masa) Then
            If CLng(masa) = bulan And Len(Trim$(CStr(wsSumber.Cells(r, 1).Value2))) > 0 Then
                TulisFieldDetailV2 wsSumber, wsTarget, r, rt, sumber, target, "NAMA"
                TulisFieldDetailV2 wsSumber, wsTarget, r, rt, sumber, target, "TIN"
                TulisFieldDetailV2 wsSumber, wsTarget, r, rt, sumber, target, "KODE"
                TulisFieldDetailV2 wsSumber, wsTarget, r, rt, sumber, target, "TIPE"
                TulisFieldDetailV2 wsSumber, wsTarget, r, rt, sumber, target, "NOMOR"
                TulisFieldDetailV2 wsSumber, wsTarget, r, rt, sumber, target, "TANGGAL"
                TulisFieldDetailV2 wsSumber, wsTarget, r, rt, sumber, target, "DPP"
                TulisFieldDetailV2 wsSumber, wsTarget, r, rt, sumber, target, "DPP_LAIN"
                TulisFieldDetailV2 wsSumber, wsTarget, r, rt, sumber, target, "PPN"
                If wsTarget.Name = "N-19-2" Then TulisFieldDetailV2 wsSumber, wsTarget, r, rt, sumber, target, "DETAIL"
                SelKiriAtasMergeV2(wsTarget.Cells(rt, CLng(target("NO")))).Value2 = nomor
                nomor = nomor + 1: rt = rt + 1
            End If
        End If
    Next r
Selesai:
    Application.CutCopyMode = False: Application.Calculation = calcLama: Application.EnableEvents = eventsLama: Application.ScreenUpdating = screenLama
    Exit Sub
Gagal:
    MsgBox "Ekspor detail " & kodeEWS & " dihentikan: " & Err.Description, vbExclamation, "EWSSYNC"
    Resume Selesai
End Sub

Private Function KolomTargetDetailPPNV2(ByVal ws As Worksheet) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d("NO") = 7: d("NAMA") = 9: d("TIN") = 17
    If ws.Name = "N-19-2" Then
        d("NOMOR") = 24: d("TANGGAL") = 31: d("DPP") = 38: d("PPN") = 45: d("DETAIL") = 59
    Else
        d("KODE") = 24: d("TIPE") = 28: d("NOMOR") = 30: d("TANGGAL") = 35: d("DPP") = 40: d("DPP_LAIN") = 46: d("PPN") = 52
        If ws.Name = "N-20-2" Then d("DETAIL") = 9
    End If
    Set KolomTargetDetailPPNV2 = d
End Function

Private Sub TulisFieldDetailV2(ByVal wsSumber As Worksheet, ByVal wsTarget As Worksheet, ByVal rs As Long, ByVal rt As Long, ByVal sumber As Object, ByVal target As Object, ByVal nama As String)
    Dim tujuan As Range

    If sumber.Exists(nama) And target.Exists(nama) Then
        Set tujuan = SelKiriAtasMergeV2(wsTarget.Cells(rt, CLng(target(nama))))
        With tujuan
            If nama = "TIN" Or nama = "NOMOR" Then
                .NumberFormat = "@"
                .Value2 = CStr(wsSumber.Cells(rs, CLng(sumber(nama))).Value2)
            Else
                .Value2 = wsSumber.Cells(rs, CLng(sumber(nama))).Value2
            End If
            .NumberFormat = wsSumber.Cells(rs, CLng(sumber(nama))).NumberFormat
            If nama = "TIN" Or nama = "NOMOR" Then .NumberFormat = "@"
            .Font.Name = wsSumber.Cells(rs, CLng(sumber(nama))).Font.Name
            .Font.Size = wsSumber.Cells(rs, CLng(sumber(nama))).Font.Size
            .Font.Bold = wsSumber.Cells(rs, CLng(sumber(nama))).Font.Bold
            .HorizontalAlignment = wsSumber.Cells(rs, CLng(sumber(nama))).HorizontalAlignment
            .VerticalAlignment = wsSumber.Cells(rs, CLng(sumber(nama))).VerticalAlignment
            .WrapText = True
        End With
    End If
End Sub

Private Function SelKiriAtasMergeV2(ByVal sel As Range) As Range
    If sel.MergeCells Then
        Set SelKiriAtasMergeV2 = sel.MergeArea.Cells(1, 1)
    Else
        Set SelKiriAtasMergeV2 = sel
    End If
End Function

Private Sub PastikanKolomDetailBarangN19(ByVal ws As Worksheet)
    If LCase$(Trim$(CStr(ws.Cells(20, 59).Value2))) = "detail barang" Then Exit Sub
    ws.Columns(59).Resize(, 6).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ws.Range(ws.Cells(20, 59), ws.Cells(22, 64)).Merge
    ws.Cells(20, 59).Value2 = "Detail Barang"
End Sub

'Mengambil sheet rincian faktur PPN; contoh N-2 menggunakan PPN-N-2-2.
Private Function SheetSumberDetailPPNExport(ByVal kodeEWS As String) As Worksheet
    Dim ws As Worksheet
    Dim namaSheet As String

    namaSheet = "PPN-" & NormalisasiKodeKKPEWS(kodeEWS) & "-2"

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(namaSheet)
    On Error GoTo 0

    If Not ws Is Nothing Then
        Set SheetSumberDetailPPNExport = ws
    Else
        Set SheetSumberDetailPPNExport = _
            DapatkanSheetLokalKKP(TAXKEY_PPN_DN, kodeEWS, False)
    End If
End Function
