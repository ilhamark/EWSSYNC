Attribute VB_Name = "modImportPPh26Alternatif"
Option Explicit

'EWSSYNC v0.5 - pemulihan alternatif impor PPh 26.

Private Type GUID_P26_ALT
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

#If VBA7 Then
    Private Declare PtrSafe Function EnumWindows Lib "user32" ( _
        ByVal lpEnumFunc As LongPtr, ByVal lParam As LongPtr) As Long
    Private Declare PtrSafe Function EnumChildWindows Lib "user32" ( _
        ByVal hWndParent As LongPtr, ByVal lpEnumFunc As LongPtr, _
        ByVal lParam As LongPtr) As Long
    Private Declare PtrSafe Function GetWindowThreadProcessId Lib "user32" ( _
        ByVal hwnd As LongPtr, ByRef lpdwProcessId As Long) As Long
    Private Declare PtrSafe Function GetWindowTextLengthW Lib "user32" ( _
        ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function GetWindowTextW Lib "user32" ( _
        ByVal hwnd As LongPtr, ByVal lpString As LongPtr, _
        ByVal cch As Long) As Long
    Private Declare PtrSafe Function GetClassNameW Lib "user32" ( _
        ByVal hwnd As LongPtr, ByVal lpClassName As LongPtr, _
        ByVal nMaxCount As Long) As Long
    Private Declare PtrSafe Function IsWindowVisible Lib "user32" ( _
        ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function PostMessageW Lib "user32" ( _
        ByVal hwnd As LongPtr, ByVal wMsg As Long, _
        ByVal wParam As LongPtr, ByVal lParam As LongPtr) As Long
    Private Declare PtrSafe Function AccessibleObjectFromWindow Lib "oleacc" ( _
        ByVal hwnd As LongPtr, ByVal dwId As Long, _
        ByRef riid As GUID_P26_ALT, ByRef ppvObject As Object) As Long
    Private Declare PtrSafe Function OpenProcess Lib "kernel32" ( _
        ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, _
        ByVal dwProcessId As Long) As LongPtr
    Private Declare PtrSafe Function TerminateProcess Lib "kernel32" ( _
        ByVal hProcess As LongPtr, ByVal uExitCode As Long) As Long
    Private Declare PtrSafe Function CloseHandle Lib "kernel32" ( _
        ByVal hObject As LongPtr) As Long
#Else
    Private Declare Function EnumWindows Lib "user32" ( _
        ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
    Private Declare Function EnumChildWindows Lib "user32" ( _
        ByVal hWndParent As Long, ByVal lpEnumFunc As Long, _
        ByVal lParam As Long) As Long
    Private Declare Function GetWindowThreadProcessId Lib "user32" ( _
        ByVal hwnd As Long, ByRef lpdwProcessId As Long) As Long
    Private Declare Function GetWindowTextLengthW Lib "user32" ( _
        ByVal hwnd As Long) As Long
    Private Declare Function GetWindowTextW Lib "user32" ( _
        ByVal hwnd As Long, ByVal lpString As Long, _
        ByVal cch As Long) As Long
    Private Declare Function GetClassNameW Lib "user32" ( _
        ByVal hwnd As Long, ByVal lpClassName As Long, _
        ByVal nMaxCount As Long) As Long
    Private Declare Function IsWindowVisible Lib "user32" ( _
        ByVal hwnd As Long) As Long
    Private Declare Function PostMessageW Lib "user32" ( _
        ByVal hwnd As Long, ByVal wMsg As Long, _
        ByVal wParam As Long, ByVal lParam As Long) As Long
    Private Declare Function AccessibleObjectFromWindow Lib "oleacc" ( _
        ByVal hwnd As Long, ByVal dwId As Long, _
        ByRef riid As GUID_P26_ALT, ByRef ppvObject As Object) As Long
    Private Declare Function OpenProcess Lib "kernel32" ( _
        ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, _
        ByVal dwProcessId As Long) As Long
    Private Declare Function TerminateProcess Lib "kernel32" ( _
        ByVal hProcess As Long, ByVal uExitCode As Long) As Long
    Private Declare Function CloseHandle Lib "kernel32" ( _
        ByVal hObject As Long) As Long
#End If

Private Const BM_CLICK As Long = &HF5
Private Const OBJID_NATIVEOM As Long = &HFFFFFFF0
Private Const PROCESS_TERMINATE As Long = &H1
Private Const INTERVAL_CEK_DETIK As Long = 1
Private Const BATAS_TUNGGU_ERROR_DETIK As Long = 120
Private Const BATAS_TUNGGU_SETELAH_END_DETIK As Long = 30

Private Const IDX_STATUS As Long = 3
Private Const IDX_PATH As Long = 4
Private Const IDX_KODE As Long = 5

Private mAntrian As Collection
Private mPathAktif As String
Private mPidAktif As Long
Private mMulaiTahap As Date
Private mWaktuBerikutnya As Date
Private mEndSudahDitekan As Boolean
Private mSedangBerjalan As Boolean
Private mBerhasil As Long
Private mGagal As Long
Private mDilewati As Long
Private mForm As Object

#If VBA7 Then
    Private mDialogHwnd As LongPtr
    Private mTombolEndHwnd As LongPtr
    Private mExcelHwnd As LongPtr
    Private mExcelChildHwnd As LongPtr
#Else
    Private mDialogHwnd As Long
    Private mTombolEndHwnd As Long
    Private mExcelHwnd As Long
    Private mExcelChildHwnd As Long
#End If
Private mTeksDialog As String
Private mFolderTrustedP26Alt As Object

'Tujuan: Mengimpor pph 26 alternatif dari daftar form.
Public Sub ImportPPh26AlternatifDariDaftarForm(ByVal frm As Object)
    Dim i As Long
    Dim pathFile As String, kode As String, status As String
    Dim kodeDariNama As String, namaKKP As String, npwp As String
    Dim kodeMasa As String, identifier As String, handler As String
    Dim pesanValidasi As String
    Dim bulanFile As Long, bulanAkhir As Long, tahunFile As Long
    Dim dikenal As Boolean, didukung As Boolean
    Dim jumlahPeriodeBerbeda As Long
    Dim unik As Object

    On Error GoTo Gagal

    If mSedangBerjalan Then
        MsgBox "Batch alternatif PPh26 masih berjalan.", _
            vbExclamation, "Impor alternatif PPh26"
        Exit Sub
    End If
    If frm Is Nothing Then Err.Raise vbObjectError + 3600, , _
        "Form importer tidak tersedia."
    If frm.lstEWS.ListCount = 0 Then
        MsgBox "Muat file EWS terlebih dahulu melalui tombol Pilih File EWS.", _
            vbExclamation, "Impor alternatif PPh26"
        Exit Sub
    End If

    Set mAntrian = New Collection
    Set unik = CreateObject("Scripting.Dictionary")
    unik.CompareMode = vbTextCompare
    mDilewati = 0
    If Not TabelDataBulanInfoTersediaP26Alt() Then _
        Err.Raise vbObjectError + 3606, , _
        "Tabel Data Bulan belum terbaca dari bagian INFO KKP lokal."

    For i = 0 To frm.lstEWS.ListCount - 1
        kode = Trim$(CStr(frm.lstEWS.List(i, IDX_KODE)))
        If StrComp(kode, KODE_EWS_PPH26, vbTextCompare) = 0 Then
            status = Trim$(CStr(frm.lstEWS.List(i, IDX_STATUS)))
            pathFile = Trim$(CStr(frm.lstEWS.List(i, IDX_PATH)))
            pesanValidasi = vbNullString

            If Len(pathFile) = 0 Then
                'Slot masa kosong tidak dihitung sebagai file dilewati.
            ElseIf Not BacaIdentitasDariNamaFile( _
                    pathFile, kodeDariNama, namaKKP, npwp, kodeMasa, _
                    bulanFile, bulanAkhir, tahunFile, identifier, _
                    dikenal, didukung, handler, pesanValidasi) Then
                frm.lstEWS.List(i, IDX_STATUS) = _
                    "Dilewati (nama file tidak valid)"
                mDilewati = mDilewati + 1
            ElseIf Not PeriodeAdaInfoKKPP26Alt(bulanFile, tahunFile) Or _
                    Not PeriodeAdaInfoKKPP26Alt(bulanAkhir, tahunFile) Then
                frm.lstEWS.List(i, IDX_STATUS) = _
                    "Dilewati (periode berbeda)"
                mDilewati = mDilewati + 1
                jumlahPeriodeBerbeda = jumlahPeriodeBerbeda + 1
            ElseIf StatusBolehDiprosesP26Alt(status) Then
                If Len(Dir$(pathFile)) > 0 Then
                    If Not unik.Exists(pathFile) Then
                        unik.Add pathFile, True
                        mAntrian.Add pathFile
                    End If
                Else
                    mDilewati = mDilewati + 1
                End If
            Else
                mDilewati = mDilewati + 1
            End If
        End If
    Next i

    If jumlahPeriodeBerbeda > 0 Then
        TampilkanPesanPeriodeEWSBerbedaP26Alt jumlahPeriodeBerbeda
    End If

    If mAntrian.Count = 0 Then
        MsgBox "Tidak ada file EWS PPh26 yang siap diproses.", _
            vbExclamation, "Impor alternatif PPh26"
        Set mAntrian = Nothing
        Exit Sub
    End If

    If Not KonfirmasiTrustedLocationP26Alt( _
            CStr(mAntrian(1))) Then
        Set mAntrian = Nothing
        Exit Sub
    End If

    Set mForm = frm
    mBerhasil = 0
    mGagal = 0
    mSedangBerjalan = True
    frm.Hide
    MulaiFileBerikutnyaP26Alt
    Exit Sub

Gagal:
    BersihkanStatusP26Alt
    MsgBox Err.Description, vbExclamation, "Impor alternatif PPh26"
End Sub

'Tujuan: Menjalankan pantau impor pph 26 alternatif.
Public Sub PantauImporPPh26Alternatif()
    Dim usia As Long, batasCoba As Long
    Dim sumber As Object

    If Not mSedangBerjalan Then Exit Sub
    mWaktuBerikutnya = 0
    On Error GoTo CobaLagi

    If Len(mPathAktif) = 0 Then
        MulaiFileBerikutnyaP26Alt
        Exit Sub
    End If

    usia = DateDiff("s", mMulaiTahap, Now)
    If Not mEndSudahDitekan Then
        If TemukanDialogErrorP26Alt() Then
            CatatJejakP26Alt "Dialog error cocok; mengirim klik End."
            If PostMessageW(mTombolEndHwnd, BM_CLICK, 0, 0) = 0 Then
                Err.Raise vbObjectError + 3601, , _
                    "Tombol End ditemukan tetapi tidak dapat ditekan."
            End If
            mEndSudahDitekan = True
            mMulaiTahap = Now
            JadwalkanPantauanP26Alt 2
            Exit Sub
        End If

        Set sumber = AmbilWorkbookEksternalP26Alt(mPathAktif)
        If Not sumber Is Nothing Then
            If AplikasiSumberSiapP26Alt(sumber) Then
                CatatJejakP26Alt _
                    "Startup selesai tanpa dialog error yang dikenal."
                SnapshotTutupDanImporP26Alt sumber
                Exit Sub
            End If
        End If

        If usia >= BATAS_TUNGGU_ERROR_DETIK Then
            GagalkanFileAktifP26Alt _
                "Dialog error EWS tidak ditemukan dalam " & _
                CStr(BATAS_TUNGGU_ERROR_DETIK) & " detik."
            Exit Sub
        End If
    Else
        Set sumber = AmbilWorkbookEksternalP26Alt(mPathAktif)
        If Not sumber Is Nothing Then
            CatatJejakP26Alt "Workbook eksternal ditemukan setelah End."
            SnapshotTutupDanImporP26Alt sumber
            Exit Sub
        End If
        If usia >= BATAS_TUNGGU_SETELAH_END_DETIK Then
            GagalkanFileAktifP26Alt _
                "EWS belum dapat dikendalikan setelah tombol End ditekan."
            Exit Sub
        End If
    End If

    JadwalkanPantauanP26Alt INTERVAL_CEK_DETIK
    Exit Sub

CobaLagi:
    If mEndSudahDitekan Then
        batasCoba = BATAS_TUNGGU_SETELAH_END_DETIK
    Else
        batasCoba = BATAS_TUNGGU_ERROR_DETIK
    End If
    If usia >= batasCoba Then
        GagalkanFileAktifP26Alt "Pemantauan gagal; Error " & _
            CStr(Err.Number) & ": " & Err.Description
    Else
        Err.Clear
        JadwalkanPantauanP26Alt INTERVAL_CEK_DETIK
    End If
End Sub

'Tujuan: Menjalankan batalkan impor pph 26 alternatif.
Public Sub BatalkanImporPPh26Alternatif()
    If Not mSedangBerjalan Then Exit Sub
    On Error Resume Next
    BatalkanJadwalP26Alt
    TutupProsesAktifP26Alt False
    On Error GoTo 0
    BersihkanStatusP26Alt
    MsgBox "Batch alternatif PPh26 dibatalkan.", _
        vbInformation, "Impor alternatif PPh26"
End Sub

'Tujuan: Menjalankan mulai file berikutnya p 26 alt.
Private Sub MulaiFileBerikutnyaP26Alt()
    Dim excelExe As String, perintah As String
    Dim idProses As Variant

    On Error GoTo Gagal
    If mAntrian Is Nothing Then GoTo Selesai
    If mAntrian.Count = 0 Then GoTo Selesai

    mPathAktif = CStr(mAntrian(1))
    mAntrian.Remove 1
    If WorkbookTerbukaDiHostP26Alt(mPathAktif) Then
        Err.Raise vbObjectError + 3602, , _
            "File sudah terbuka di Excel EWSSYNC. Tutup file EWS sebelum batch."
    End If

    excelExe = Application.Path & Application.PathSeparator & "EXCEL.EXE"
    If Len(Dir$(excelExe)) = 0 Then Err.Raise vbObjectError + 3603, , _
        "EXCEL.EXE tidak ditemukan pada folder instalasi Excel."
    perintah = KutipP26Alt(excelExe) & " /x /e " & KutipP26Alt(mPathAktif)
    idProses = Shell(perintah, vbNormalFocus)
    If CLng(idProses) = 0 Then Err.Raise vbObjectError + 3604, , _
        "Excel kedua gagal dijalankan."

    mPidAktif = CLng(idProses)
    mEndSudahDitekan = False
    mMulaiTahap = Now
    CatatJejakP26Alt "Excel kedua diluncurkan."
    JadwalkanPantauanP26Alt INTERVAL_CEK_DETIK
    Exit Sub

Selesai:
    SelesaikanBatchP26Alt
    Exit Sub
Gagal:
    GagalkanFileAktifP26Alt Err.Description
End Sub

'Tujuan: Menjalankan snapshot tutup dan impor p 26 alt.
Private Sub SnapshotTutupDanImporP26Alt(ByVal wbEksternal As Object)
    Dim appEksternal As Object, wbLokal As Workbook
    Dim evEksternal As Boolean, alertsEksternal As Boolean
    Dim calcEksternal As Long
    Dim evLama As Boolean, alertsLama As Boolean
    Dim calcLama As XlCalculation, securityLama As MsoAutomationSecurity
    Dim hasil As Variant, pesan As String
    Dim simpanBerhasil As Boolean, statusHostTersimpan As Boolean

    On Error GoTo Gagal
    Set appEksternal = wbEksternal.Application
    evEksternal = appEksternal.EnableEvents
    alertsEksternal = appEksternal.DisplayAlerts
    calcEksternal = appEksternal.Calculation

    appEksternal.EnableEvents = False
    appEksternal.DisplayAlerts = False
    appEksternal.Calculation = -4135
    CatatJejakP26Alt "Menyimpan snapshot EWS."
    wbEksternal.Save
    simpanBerhasil = True
    wbEksternal.Close False
    appEksternal.Quit
    Set wbEksternal = Nothing
    Set appEksternal = Nothing
    HentikanProsesExcelP26Alt mPidAktif
    CatatJejakP26Alt "Proses Excel kedua ditutup."

    evLama = Application.EnableEvents
    alertsLama = Application.DisplayAlerts
    calcLama = Application.Calculation
    securityLama = Application.AutomationSecurity
    statusHostTersimpan = True

    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.AutomationSecurity = msoAutomationSecurityForceDisable
    Set wbLokal = Workbooks.Open(Filename:=mPathAktif, UpdateLinks:=0, _
        ReadOnly:=False, IgnoreReadOnlyRecommended:=True, _
        AddToMru:=False, Notify:=False)
    CatatJejakP26Alt "EWS dibuka ulang dengan macro dan event mati."

    hasil = Application.Run("'" & Replace(ThisWorkbook.Name, "'", "''") & _
        "'!ImporPPh26TerbukaBatch", mPathAktif)
    pesan = CStr(hasil)
    If StrComp(pesan, "OK", vbTextCompare) <> 0 Then
        Err.Raise vbObjectError + 3605, , pesan
    End If
    CatatJejakP26Alt "Impor selesai dengan hasil OK."

    wbLokal.Close False
    Set wbLokal = Nothing
    If statusHostTersimpan Then
        Application.AutomationSecurity = securityLama
        Application.Calculation = calcLama
        Application.DisplayAlerts = alertsLama
        Application.EnableEvents = evLama
    End If

    mBerhasil = mBerhasil + 1
    TandaiStatusFormP26Alt mPathAktif, "Sudah diimpor"
    mPathAktif = vbNullString
    mPidAktif = 0
    MulaiFileBerikutnyaP26Alt
    Exit Sub

Gagal:
    pesan = Err.Description
    On Error Resume Next
    If Not wbLokal Is Nothing Then wbLokal.Close False
    If statusHostTersimpan Then
        Application.AutomationSecurity = securityLama
        Application.Calculation = calcLama
        Application.DisplayAlerts = alertsLama
        Application.EnableEvents = evLama
    End If
    If Not appEksternal Is Nothing Then
        appEksternal.EnableEvents = False
        appEksternal.DisplayAlerts = False
        If Not wbEksternal Is Nothing Then wbEksternal.Close False
        appEksternal.Quit
    End If
    On Error GoTo 0
    If simpanBerhasil Then pesan = "Snapshot tersimpan; " & pesan
    GagalkanFileAktifP26Alt pesan
End Sub

'Tujuan: Menjalankan aplikasi sumber siap p 26 alt.
Private Function AplikasiSumberSiapP26Alt(ByVal wb As Object) As Boolean
    On Error GoTo BelumSiap
    AplikasiSumberSiapP26Alt = CBool(wb.Application.Ready)
    Exit Function
BelumSiap:
    Err.Clear
End Function

'Tujuan: Mengambil workbook eksternal p 26 alt.
Private Function AmbilWorkbookEksternalP26Alt( _
    ByVal pathFile As String) As Object
    Dim wb As Object, appSumber As Object, objekNative As Object
    Dim iid As GUID_P26_ALT
    Dim hasil As Long

    On Error Resume Next
    mExcelHwnd = 0
    mExcelChildHwnd = 0
    EnumWindows AddressOf EnumExcelProcP26Alt, 0
    If mExcelHwnd = 0 Or mExcelChildHwnd = 0 Then Exit Function

    iid.Data1 = &H20400
    iid.Data2 = 0
    iid.Data3 = 0
    iid.Data4(0) = &HC0
    iid.Data4(7) = &H46
    hasil = AccessibleObjectFromWindow( _
        mExcelChildHwnd, OBJID_NATIVEOM, iid, objekNative)
    If hasil <> 0 Or objekNative Is Nothing Then Exit Function
    Set appSumber = objekNative.Application
    If appSumber Is Nothing Then Set appSumber = objekNative
    If appSumber Is Nothing Then Exit Function
    For Each wb In appSumber.Workbooks
        If StrComp(wb.FullName, pathFile, vbTextCompare) = 0 Then
            Set AmbilWorkbookEksternalP26Alt = wb
            Exit Function
        End If
    Next wb
    On Error GoTo 0
End Function

#If VBA7 Then

'Tujuan: Menjalankan enum excel proc p 26 alt.
Public Function EnumExcelProcP26Alt( _
    ByVal hwnd As LongPtr, ByVal lParam As LongPtr) As Long
#Else

'Tujuan: Menjalankan enum excel proc p 26 alt.
Public Function EnumExcelProcP26Alt( _
    ByVal hwnd As Long, ByVal lParam As Long) As Long
#End If
    Dim pid As Long
    EnumExcelProcP26Alt = 1
    If IsWindowVisible(hwnd) = 0 Then Exit Function
    GetWindowThreadProcessId hwnd, pid
    If pid <> mPidAktif Then Exit Function
    If StrComp(TeksKelasJendelaP26Alt(hwnd), _
            "XLMAIN", vbTextCompare) <> 0 Then Exit Function
    mExcelHwnd = hwnd
    EnumChildWindows hwnd, AddressOf EnumExcelChildProcP26Alt, 0
    If mExcelChildHwnd <> 0 Then EnumExcelProcP26Alt = 0
End Function

#If VBA7 Then

'Tujuan: Menjalankan enum excel child proc p 26 alt.
Public Function EnumExcelChildProcP26Alt( _
    ByVal hwnd As LongPtr, ByVal lParam As LongPtr) As Long
#Else

'Tujuan: Menjalankan enum excel child proc p 26 alt.
Public Function EnumExcelChildProcP26Alt( _
    ByVal hwnd As Long, ByVal lParam As Long) As Long
#End If
    Dim pid As Long
    EnumExcelChildProcP26Alt = 1
    GetWindowThreadProcessId hwnd, pid
    If pid <> mPidAktif Then Exit Function
    If StrComp(TeksKelasJendelaP26Alt(hwnd), _
            "EXCEL7", vbTextCompare) <> 0 Then Exit Function
    mExcelChildHwnd = hwnd
    EnumExcelChildProcP26Alt = 0
End Function

'Tujuan: Menjalankan temukan dialog error p 26 alt.
Private Function TemukanDialogErrorP26Alt() As Boolean
    mDialogHwnd = 0
    mTombolEndHwnd = 0
    mTeksDialog = vbNullString
    EnumWindows AddressOf EnumWindowsProcP26Alt, 0
    If mDialogHwnd = 0 Or mTombolEndHwnd = 0 Then Exit Function
    If InStr(1, mTeksDialog, _
            "Unable to set the HorizontalAlignment property", _
            vbTextCompare) = 0 And _
       InStr(1, mTeksDialog, _
            "Unable to set the Locked property", _
            vbTextCompare) = 0 Then Exit Function
    TemukanDialogErrorP26Alt = True
End Function

#If VBA7 Then

'Tujuan: Menjalankan enum windows proc p 26 alt.
Public Function EnumWindowsProcP26Alt( _
    ByVal hwnd As LongPtr, ByVal lParam As LongPtr) As Long
#Else

'Tujuan: Menjalankan enum windows proc p 26 alt.
Public Function EnumWindowsProcP26Alt( _
    ByVal hwnd As Long, ByVal lParam As Long) As Long
#End If
    Dim pid As Long, judul As String, kelas As String
    EnumWindowsProcP26Alt = 1
    If IsWindowVisible(hwnd) = 0 Then Exit Function
    GetWindowThreadProcessId hwnd, pid
    If pid <> mPidAktif Then Exit Function
    kelas = TeksKelasJendelaP26Alt(hwnd)
    If StrComp(kelas, "#32770", vbTextCompare) <> 0 Then Exit Function
    judul = TeksJendelaP26Alt(hwnd)
    If StrComp(judul, "Microsoft Visual Basic", vbTextCompare) <> 0 Then _
        Exit Function
    mDialogHwnd = hwnd
    EnumChildWindows hwnd, AddressOf EnumChildProcP26Alt, 0
    If mTombolEndHwnd <> 0 Then EnumWindowsProcP26Alt = 0
End Function

#If VBA7 Then

'Tujuan: Menjalankan enum child proc p 26 alt.
Public Function EnumChildProcP26Alt( _
    ByVal hwnd As LongPtr, ByVal lParam As LongPtr) As Long
#Else

'Tujuan: Menjalankan enum child proc p 26 alt.
Public Function EnumChildProcP26Alt( _
    ByVal hwnd As Long, ByVal lParam As Long) As Long
#End If
    Dim teks As String, kelas As String, caption As String
    EnumChildProcP26Alt = 1
    teks = TeksJendelaP26Alt(hwnd)
    If Len(teks) > 0 Then mTeksDialog = mTeksDialog & " " & teks
    kelas = TeksKelasJendelaP26Alt(hwnd)
    If StrComp(kelas, "Button", vbTextCompare) = 0 Then
        caption = Replace(Trim$(teks), "&", vbNullString)
        If StrComp(caption, "End", vbTextCompare) = 0 Then
            mTombolEndHwnd = hwnd
        End If
    End If
End Function

#If VBA7 Then

'Tujuan: Menjalankan teks jendela p 26 alt.
Private Function TeksJendelaP26Alt(ByVal hwnd As LongPtr) As String
#Else

'Tujuan: Menjalankan teks jendela p 26 alt.
Private Function TeksJendelaP26Alt(ByVal hwnd As Long) As String
#End If
    Dim panjang As Long, hasil As Long, buffer As String
    panjang = GetWindowTextLengthW(hwnd)
    If panjang <= 0 Then Exit Function
    buffer = String$(panjang + 1, vbNullChar)
    hasil = GetWindowTextW(hwnd, StrPtr(buffer), panjang + 1)
    If hasil > 0 Then TeksJendelaP26Alt = Left$(buffer, hasil)
End Function

#If VBA7 Then

'Tujuan: Menjalankan teks kelas jendela p 26 alt.
Private Function TeksKelasJendelaP26Alt(ByVal hwnd As LongPtr) As String
#Else

'Tujuan: Menjalankan teks kelas jendela p 26 alt.
Private Function TeksKelasJendelaP26Alt(ByVal hwnd As Long) As String
#End If
    Dim hasil As Long, buffer As String
    buffer = String$(256, vbNullChar)
    hasil = GetClassNameW(hwnd, StrPtr(buffer), 256)
    If hasil > 0 Then TeksKelasJendelaP26Alt = Left$(buffer, hasil)
End Function

'Tujuan: Menjalankan gagalkan file aktif p 26 alt.
Private Sub GagalkanFileAktifP26Alt(ByVal pesan As String)
    Dim sisa As Long
    CatatJejakP26Alt "GAGAL: " & pesan
    On Error Resume Next
    TutupProsesAktifP26Alt False
    On Error GoTo 0
    mGagal = mGagal + 1
    TandaiStatusFormP26Alt mPathAktif, "Gagal"
    If Not mAntrian Is Nothing Then sisa = mAntrian.Count
    mDilewati = mDilewati + sisa
    Set mAntrian = New Collection
    mPathAktif = vbNullString
    mPidAktif = 0
    SelesaikanBatchP26Alt
End Sub

'Tujuan: Menjalankan catat jejak p 26 alt.
Private Sub CatatJejakP26Alt(ByVal pesan As String)
    Dim nomor As Integer, logPath As String
    On Error Resume Next
    logPath = ThisWorkbook.Path & Application.PathSeparator & _
        "EWSSYNC_P26_Alternatif_Debug.log"
    nomor = FreeFile
    Open logPath For Append As #nomor
    Print #nomor, Format$(Now, "yyyy-mm-dd hh:nn:ss") & _
        " | PID=" & CStr(mPidAktif) & _
        " | " & NamaFileP26Alt(mPathAktif) & " | " & pesan
    Close #nomor
    On Error GoTo 0
End Sub

'Tujuan: Menutup proses aktif p 26 alt.
Private Sub TutupProsesAktifP26Alt(ByVal simpan As Boolean)
    Dim wb As Object, appSumber As Object
    On Error Resume Next
    Set wb = AmbilWorkbookEksternalP26Alt(mPathAktif)
    If Not wb Is Nothing Then
        Set appSumber = wb.Application
        appSumber.EnableEvents = False
        appSumber.DisplayAlerts = False
        wb.Close simpan
        appSumber.Quit
    End If
    Set wb = Nothing
    Set appSumber = Nothing
    HentikanProsesExcelP26Alt mPidAktif
    On Error GoTo 0
End Sub

'Tujuan: Mengakhiri hanya proses Excel kedua milik file aktif.
Private Sub HentikanProsesExcelP26Alt(ByVal pid As Long)
    Dim pidHost As Long
#If VBA7 Then
    Dim proses As LongPtr
#Else
    Dim proses As Long
#End If

    If pid <= 0 Then Exit Sub
    GetWindowThreadProcessId Application.hwnd, pidHost
    If pid = pidHost Then Exit Sub

    proses = OpenProcess(PROCESS_TERMINATE, 0, pid)
    If proses = 0 Then Exit Sub
    TerminateProcess proses, 0
    CloseHandle proses
End Sub

'Tujuan: Menjalankan jadwalkan pantauan p 26 alt.
Private Sub JadwalkanPantauanP26Alt(ByVal detik As Long)
    BatalkanJadwalP26Alt
    mWaktuBerikutnya = Now + TimeSerial(0, 0, detik)
    Application.OnTime EarliestTime:=mWaktuBerikutnya, _
        Procedure:="'" & Replace(ThisWorkbook.Name, "'", "''") & _
        "'!PantauImporPPh26Alternatif", Schedule:=True
End Sub

'Tujuan: Menjalankan batalkan jadwal p 26 alt.
Private Sub BatalkanJadwalP26Alt()
    If mWaktuBerikutnya = 0 Then Exit Sub
    On Error Resume Next
    Application.OnTime EarliestTime:=mWaktuBerikutnya, _
        Procedure:="'" & Replace(ThisWorkbook.Name, "'", "''") & _
        "'!PantauImporPPh26Alternatif", Schedule:=False
    On Error GoTo 0
    mWaktuBerikutnya = 0
End Sub

'Tujuan: Menjalankan selesaikan batch p 26 alt.
Private Sub SelesaikanBatchP26Alt()
    Dim ringkasan As String
    ringkasan = "Batch alternatif PPh26 selesai." & vbCrLf & _
        "Berhasil: " & CStr(mBerhasil) & vbCrLf & _
        "Gagal: " & CStr(mGagal) & vbCrLf & _
        "Dilewati: " & CStr(mDilewati)
    BersihkanStatusP26Alt
    MsgBox ringkasan & vbCrLf & vbCrLf & _
        "Rincian tersedia di Log Impor EWS.", _
        vbInformation, "Impor alternatif PPh26"
End Sub

'Tujuan: Menjalankan bersihkan status p 26 alt.
Private Sub BersihkanStatusP26Alt()
    On Error Resume Next
    BatalkanJadwalP26Alt
    Set mAntrian = Nothing
    Set mForm = Nothing
    mPathAktif = vbNullString
    mPidAktif = 0
    mEndSudahDitekan = False
    mSedangBerjalan = False
    On Error GoTo 0
End Sub

'Tujuan: Menjalankan tandai status form p 26 alt.
Private Sub TandaiStatusFormP26Alt( _
    ByVal pathFile As String, ByVal statusBaru As String)
    Dim i As Long
    On Error Resume Next
    If mForm Is Nothing Then Exit Sub
    For i = 0 To mForm.lstEWS.ListCount - 1
        If StrComp(Trim$(CStr(mForm.lstEWS.List(i, IDX_PATH))), _
                pathFile, vbTextCompare) = 0 Then
            mForm.lstEWS.List(i, IDX_STATUS) = statusBaru
        End If
    Next i
    On Error GoTo 0
End Sub

'Tujuan: Menentukan status boleh diproses p 26 alt.
Private Function StatusBolehDiprosesP26Alt(ByVal status As String) As Boolean
    StatusBolehDiprosesP26Alt = _
        StrComp(status, "Siap", vbTextCompare) = 0 Or _
        StrComp(status, "Sudah diimpor", vbTextCompare) = 0 Or _
        StrComp(status, "Versi baru", vbTextCompare) = 0
End Function

'Tujuan: Konfirmasi satu kali untuk folder EWS PPh 26 pada engine alternatif.
Private Function KonfirmasiTrustedLocationP26Alt( _
    ByVal pathFile As String) As Boolean

    Dim folderEWS As String
    Dim posisi As Long
    Dim jawaban As VbMsgBoxResult

    posisi = InStrRev(pathFile, Application.PathSeparator)
    If posisi > 0 Then
        folderEWS = Left$(pathFile, posisi - 1)
    Else
        folderEWS = CurDir$
    End If

    If mFolderTrustedP26Alt Is Nothing Then
        Set mFolderTrustedP26Alt = CreateObject("Scripting.Dictionary")
        mFolderTrustedP26Alt.CompareMode = vbTextCompare
    End If

    If mFolderTrustedP26Alt.Exists(folderEWS) Then
        KonfirmasiTrustedLocationP26Alt = True
        Exit Function
    End If

    jawaban = MsgBox( _
        "Impor PPh Pasal 26 memerlukan prosedur tambahan menggunakan PowerShell." & _
        vbCrLf & vbCrLf & _
        "Pastikan folder berikut sudah menjadi Trusted Location:" & _
        vbCrLf & folderEWS & vbCrLf & vbCrLf & _
        "Klik OK untuk melanjutkan." & vbCrLf & _
        "Klik Cancel untuk membatalkan.", _
        vbOKCancel + vbExclamation + vbDefaultButton2, _
        "Impor PPh 26")

    If jawaban = vbOK Then
        mFolderTrustedP26Alt.Add folderEWS, True
        KonfirmasiTrustedLocationP26Alt = True
    End If

End Function

'Bagian: validasi periode PPh 26.
'Tujuan: Memastikan tabel Data Bulan tersedia tanpa ketergantungan modul lain.
Private Function TabelDataBulanInfoTersediaP26Alt() As Boolean
    Dim ws As Worksheet, penanda As Range
    Dim i As Long, bulan As Long, tahun As Long

    If Not AmbilPenandaDataBulanP26Alt(ws, penanda) Then Exit Function
    For i = 1 To 12
        If BacaBulanTahunP26Alt( _
                ws.Cells(penanda.Row + i, penanda.Column), _
                bulan, tahun) Then
            TabelDataBulanInfoTersediaP26Alt = True
            Exit Function
        End If
    Next i
End Function

'Tujuan: Memeriksa pasangan bulan dan tahun pada Data Bulan.
Private Function PeriodeAdaInfoKKPP26Alt( _
    ByVal bulanDicari As Long, ByVal tahunDicari As Long) As Boolean
    Dim ws As Worksheet, penanda As Range
    Dim i As Long, bulan As Long, tahun As Long

    If bulanDicari < 1 Or bulanDicari > 12 Or tahunDicari = 0 Then _
        Exit Function
    If Not AmbilPenandaDataBulanP26Alt(ws, penanda) Then Exit Function

    For i = 1 To 12
        If BacaBulanTahunP26Alt( _
                ws.Cells(penanda.Row + i, penanda.Column), _
                bulan, tahun) Then
            If bulan = bulanDicari And tahun = tahunDicari Then
                PeriodeAdaInfoKKPP26Alt = True
                Exit Function
            End If
        End If
    Next i
End Function

'Tujuan: Menemukan judul Data Bulan pada INFO.
Private Function AmbilPenandaDataBulanP26Alt( _
    ByRef ws As Worksheet, ByRef penanda As Range) As Boolean
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("INFO")
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    Set penanda = ws.Cells.Find(What:="Data Bulan", _
        After:=ws.Cells(1, 1), LookIn:=xlValues, LookAt:=xlWhole, _
        SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
    AmbilPenandaDataBulanP26Alt = Not penanda Is Nothing
End Function

'Tujuan: Membaca identifier masa YYYY-MM dari Data Bulan pada INFO.
Private Function BacaBulanTahunP26Alt( _
    ByVal sel As Range, ByRef bulan As Long, ByRef tahun As Long) As Boolean
    Dim nilai As Variant, teks As String

    bulan = 0
    tahun = 0
    nilai = sel.Value2
    If IsError(nilai) Or IsEmpty(nilai) Then Exit Function

    If IsNumeric(nilai) Then
        If CDbl(nilai) >= CDbl(DateSerial(2000, 1, 1)) And _
                CDbl(nilai) <= CDbl(DateSerial(2100, 12, 31)) Then
            bulan = Month(CDate(CDbl(nilai)))
            tahun = Year(CDate(CDbl(nilai)))
        End If
    Else
        teks = Trim$(CStr(nilai))
        If teks Like "####-##" Then
            tahun = CLng(Left$(teks, 4))
            bulan = CLng(Right$(teks, 2))
        End If
    End If

    BacaBulanTahunP26Alt = _
        (bulan >= 1 And bulan <= 12 And tahun >= 2000 And tahun <= 2100)
End Function

'Tujuan: Menjelaskan file PPh 26 yang dilewati karena periode berbeda.
Private Sub TampilkanPesanPeriodeEWSBerbedaP26Alt(ByVal jumlahFile As Long)
    MsgBox CStr(jumlahFile) & _
        " file EWS PPh26 dilewati karena periodenya tidak terdapat pada " & _
        "tabel Data Bulan di INFO." & vbCrLf & vbCrLf & _
        "File PPh26 lain yang periodenya sesuai tetap diproses.", _
        vbExclamation, "Periode EWS PPh26 Tidak Sesuai"
End Sub

'Tujuan: Memeriksa workbook terbuka di host p 26 alt.
Private Function WorkbookTerbukaDiHostP26Alt( _
    ByVal pathFile As String) As Boolean
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, pathFile, vbTextCompare) = 0 Then
            WorkbookTerbukaDiHostP26Alt = True
            Exit Function
        End If
    Next wb
End Function

'Tujuan: Mengutip p 26 alt.
Private Function KutipP26Alt(ByVal teks As String) As String
    KutipP26Alt = Chr$(34) & Replace(teks, Chr$(34), _
        Chr$(34) & Chr$(34)) & Chr$(34)
End Function

'Tujuan: Mengembalikan nama file p 26 alt.
Private Function NamaFileP26Alt(ByVal pathFile As String) As String
    Dim posisi As Long
    posisi = InStrRev(pathFile, Application.PathSeparator)
    If posisi > 0 Then
        NamaFileP26Alt = Mid$(pathFile, posisi + 1)
    Else
        NamaFileP26Alt = pathFile
    End If
End Function
