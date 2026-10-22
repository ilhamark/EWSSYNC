VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmImportFaktur 
   Caption         =   "Impor Faktur Pajak"
   ClientHeight    =   7230
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7560
   OleObjectBlob   =   "frmImportFaktur_v0.55.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmImportFaktur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()

    With Me.lstFile
        .Clear
        .ColumnCount = 5
        .ColumnWidths = "55 pt;70 pt;60 pt;250 pt;0 pt"
    End With

    Me.chkGantiData.Value = True
    Me.cmdImpor.Enabled = False

End Sub

Private Sub cmdPilihCSV_Click()

    modImportFakturPajak.PilihFileFaktur Me

End Sub

Private Sub cmdBersihkanDaftar_Click()

    modImportFakturPajak.BersihkanDaftarFile Me

End Sub

Private Sub cmdBatal_Click()

    Unload Me

End Sub

Private Sub cmdImpor_Click()

    modImportFakturPajak.JalankanImporDariForm Me

End Sub
