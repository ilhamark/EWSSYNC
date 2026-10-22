VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmExportEWS 
   Caption         =   "UserForm1"
   ClientHeight    =   7635
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15300
   OleObjectBlob   =   "frmExportEWS_v0.55.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmExportEWS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Identitas kumpulan file EWS yang dipilih.
Public namaWPReferensi As String
Public npwpReferensi As String
Public TahunReferensi As Long

'Mencegah event ComboBox berjalan ketika daftar sedang dibangun.
Private sedangInisialisasi As Boolean
Private chkCFMWP As Object

Private Sub UserForm_Initialize()

    Dim i As Long
    Dim tahap As String

    On Error GoTo GagalInisialisasi

    sedangInisialisasi = True

    tahap = "mengatur judul form"
    Me.caption = "Ekspor Data KKP ke EWS"

    Set chkCFMWP = Me.Controls.Add("Forms.CheckBox.1", "chkPerbaruiCFMWP", True)
    With chkCFMWP
        .Caption = "Perbarui CFM WP dari KKP Lokal"
        .Left = 9600
        .Top = 7000
        .Width = 250
        .Height = 18
        .Value = False
    End With

    namaWPReferensi = vbNullString
    npwpReferensi = vbNullString
    TahunReferensi = 0

    'Daftar file tetap memakai struktur engine multifile saat ini.
    tahap = "menyiapkan lstEWS"
    With Me.lstEWS

        .Clear
        .ColumnCount = 6
        .ColumnWidths = _
            "85 pt;200 pt;50 pt;0 pt;0 pt;0 pt"
        .MultiSelect = fmMultiSelectSingle

        For i = 1 To 12

            .AddItem NamaBulanUntukFormExportEWS(i)
            .List(i - 1, 1) = vbNullString
            .List(i - 1, 2) = "Belum dipilih"
            .List(i - 1, 3) = vbNullString
            .List(i - 1, 4) = i
            .List(i - 1, 5) = 0

        Next i

    End With

    Me.lblWPValue.caption = "-"
    Me.lblTahunValue.caption = "-"
    Me.lblJumlah.caption = "0 dari 12 masa siap"
    Me.lblJumlahKKP.caption = "0 KKP dipilih"

    tahap = "memuat jenis pajak"
    MuatJenisPajakExportEWS Me

    tahap = "memuat kelompok KKP"
    MuatKelompokKKPExportEWS Me

    tahap = "memuat daftar KKP"
    MuatDaftarKKPExportEWS Me

    sedangInisialisasi = False

    PerbaruiRingkasanFormExportEWS Me

    Exit Sub

GagalInisialisasi:

    sedangInisialisasi = False

    MsgBox _
        "Form gagal saat " & tahap & "." & _
        vbCrLf & vbCrLf & _
        "Error " & CStr(Err.Number) & ": " & _
        Err.Description, _
        vbCritical, _
        "Inisialisasi Form Ekspor"

End Sub

Private Sub cboJenisPajak_Change()

    If sedangInisialisasi Then Exit Sub

    sedangInisialisasi = True

    'File yang sudah dipilih harus dikosongkan ketika jenis EWS berubah.
    'Ini mencegah file PPN diproses memakai pilihan KKP jenis pajak lain.
    KosongkanFilePadaFormExportEWS Me

    MuatKelompokKKPExportEWS Me
    MuatDaftarKKPExportEWS Me

    sedangInisialisasi = False

    PerbaruiRingkasanFormExportEWS Me

End Sub

Private Sub cboKelompokKKP_Change()

    If sedangInisialisasi Then Exit Sub

    MuatDaftarKKPExportEWS Me

End Sub

Private Sub lstKKP_Change()

    If sedangInisialisasi Then Exit Sub

    PerbaruiRingkasanFormExportEWS Me

End Sub

Private Sub cmdPilihSemuaKKP_Click()

    PilihSemuaKKPExportEWS Me

End Sub

Private Sub cmdKosongkanKKP_Click()

    KosongkanPilihanKKPExportEWS Me

End Sub

Private Sub cmdBrowse_Click()

    PilihFileUntukFormExportEWS Me
    PerbaruiRingkasanFormExportEWS Me

End Sub

Private Sub cmdKosongkan_Click()

    KosongkanFilePadaFormExportEWS Me
    PerbaruiRingkasanFormExportEWS Me

End Sub

Private Sub cmdExport_Click()

    JalankanExportTerpilihDariForm Me

End Sub

Private Sub cmdBatal_Click()

    Unload Me

End Sub


