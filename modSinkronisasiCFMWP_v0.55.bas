Attribute VB_Name = "modSinkronisasiCFMWP_v054"
Option Explicit

' EWSSYNC v0.54
' Mengambil CFM WP dari tabel Menurut Wajib Pajak pada KKP Lokal
' (kolom H:S = masa 1:12) dan menulisnya ke MAIN LEVEL EWS.

Public Sub TerapkanCFMWP054(ByVal wbEWS As Workbook, ByVal taxKey As String, _
    ByVal bulan As Long, ByVal daftarKKP As Variant, ByRef jumlahDitulis As Long, _
    ByRef catatan As String)

    Dim ws As Worksheet, kode As Variant, komponen As Variant
    Dim s As String, tujuan As String, nilai As Variant
    Dim i As Long, kolomMasa As Long

    kolomMasa = 7 + bulan ' H:S
    If bulan < 1 Or bulan > 12 Then Err.Raise vbObjectError + 5400, , "Nomor masa tidak valid."

    For Each kode In daftarKKP
        Set ws = SheetSumberCFMWP054(taxKey)
        If ws Is Nothing Then Err.Raise vbObjectError + 5401, , _
            "Sheet sumber CFM WP tidak ditemukan untuk " & taxKey & "."

        For Each komponen In KomponenCFMWP054(taxKey, CStr(kode))
            s = Split(CStr(komponen), "|")(0)
            tujuan = Split(CStr(komponen), "|")(1)
            i = CLng(Split(CStr(komponen), "|")(2))
            nilai = ws.Cells(i, kolomMasa).Value2

            If Not IsError(nilai) Then
                If Len(Trim$(CStr(nilai))) > 0 And IsNumeric(nilai) Then
                    wbEWS.Worksheets("MAIN LEVEL").Range(tujuan).Value2 = CDbl(nilai)
                    jumlahDitulis = jumlahDitulis + 1
                End If
            End If
        Next komponen
    Next kode

    HitungRingkasanCFMWP054 wbEWS, taxKey, bulan
    catatan = "CFM WP diperbarui dari KKP Lokal masa " & CStr(bulan) & "."
End Sub

Private Function SheetSumberCFMWP054(ByVal taxKey As String) As Worksheet
    Dim nama As String
    Select Case UCase$(Trim$(taxKey))
        Case "P21", "PPH21": nama = "PPh 21"
        Case "P21F", "PPH21F": nama = "PPh 21F"
        Case "P23", "PPH23": nama = "PPh 23"
        Case "P26", "PPH26": nama = "PPh 26"
        Case "P42", "PPH42": nama = "PPh 4(2)"
        Case "PPNDN": nama = "PPN"
    End Select
    On Error Resume Next
    Set SheetSumberCFMWP054 = ThisWorkbook.Worksheets(nama)
    On Error GoTo 0
End Function

Private Function KomponenCFMWP054(ByVal taxKey As String, ByVal kode As String) As Collection
    Dim c As New Collection, k As String
    k = Replace(UCase$(kode), "-", vbNullString)

    If UCase$(taxKey) = "PPNDN" Then
        Select Case k
            Case "N1": c.Add "DPP|AE15|143"
            Case "N2": c.Add "DPP|AE16|144": c.Add "DPP_LAIN|AE17|145": c.Add "PPN|AE18|146"
            Case "N3": c.Add "DPP|AE19|147": c.Add "DPP_LAIN|AE20|148": c.Add "PPN|AE21|149"
            Case "N4": c.Add "DPP|AE22|150": c.Add "PPN|AE24|151"
            Case "N5": c.Add "DPP|AE25|152": c.Add "DPP_LAIN|AE26|153": c.Add "PPN|AE27|154"
            Case "N6": c.Add "DPP|AE28|155": c.Add "DPP_LAIN|AE29|156": c.Add "PPN|AE30|157"
            Case "N7": c.Add "DPP|AE31|158": c.Add "DPP_LAIN|AE32|159": c.Add "PPN|AE33|160"
            Case "N8": c.Add "DPP|AE34|161": c.Add "DPP_LAIN|AE35|162": c.Add "PPN|AE36|163"
            Case "N9": c.Add "DPP|AE37|164": c.Add "DPP_LAIN|AE38|165": c.Add "PPN|AE39|166"
            Case "N10": c.Add "DPP|AE42|169"
            Case "N19": c.Add "DPP|AE58|173": c.Add "PPN|AE59|174"
            Case "N20": c.Add "DPP|AE60|175": c.Add "DPP_LAIN|AE61|176": c.Add "PPN|AE62|177"
            Case "N21": c.Add "DPP|AE63|178": c.Add "PPN|AE64|179"
            Case "N22": c.Add "DPP|AE65|180": c.Add "DPP_LAIN|AE66|181": c.Add "PPN|AE67|182"
            Case "N23": c.Add "PPN|AE69|184"
            Case "N24": c.Add "DPP|AE72|187": c.Add "DPP_LAIN|AE73|188": c.Add "PPN|AE74|189"
            Case "N25": c.Add "DPP|AE75|190"
            Case "O1": c.Add "JUMLAH|AE81|198"
            Case "O2": c.Add "PPN|AE83|200"
            Case "O3": c.Add "JUMLAH|AE84|201"
            Case "O4": c.Add "JUMLAH|AE85|202"
            Case "P1": c.Add "JUMLAH|AE88|205"
            Case "P2": c.Add "JUMLAH|AE89|206"
            Case "P3": c.Add "JUMLAH|AE90|207"
            Case "P4": c.Add "JUMLAH|AE91|208"
            Case "Q1": c.Add "JUMLAH|AE94|211"
            Case "Q2": c.Add "PPN|AE95|212"
        End Select
    ElseIf UCase$(taxKey) = "PPH26" Then
        Select Case k
            Case "B1": c.Add "DPP|AD16|89": c.Add "PPH|AD20|93"
            Case "B2": c.Add "DPP|AD17|90": c.Add "PPH|AD21|94"
            Case "C1": c.Add "JUMLAH|AD23|96"
            Case "N1" To "N8": c.Add "JUMLAH|AD" & CStr(26 + CLng(Mid$(k, 2))) & "|" & CStr(99 + CLng(Mid$(k, 2)))
            Case "O1" To "O3": c.Add "JUMLAH|AD" & CStr(35 + CLng(Mid$(k, 2))) & "|" & CStr(108 + CLng(Mid$(k, 2)))
            Case "P1" To "P3": c.Add "JUMLAH|AD" & CStr(41 + CLng(Mid$(k, 2))) & "|" & CStr(113 + CLng(Mid$(k, 2)))
        End Select
    Else
        Select Case k
            Case "B1": c.Add "DPP|AD16|87": c.Add "PPH|AD20|90"
            Case "C1": c.Add "JUMLAH|AD21|92"
            Case "N1" To "N8": c.Add "JUMLAH|AD" & CStr(24 + CLng(Mid$(k, 2))) & "|" & CStr(95 + CLng(Mid$(k, 2)))
            Case "O1" To "O3": c.Add "JUMLAH|AD" & CStr(33 + CLng(Mid$(k, 2))) & "|" & CStr(104 + CLng(Mid$(k, 2)))
            Case "P1" To "P3": c.Add "JUMLAH|AD" & CStr(38 + CLng(Mid$(k, 2))) & "|" & CStr(109 + CLng(Mid$(k, 2)))
        End Select
    End If
    Set KomponenCFMWP054 = c
End Function

Private Sub HitungRingkasanCFMWP054(ByVal wb As Workbook, ByVal taxKey As String, ByVal bulan As Long)
    Dim ws As Worksheet: Set ws = wb.Worksheets("MAIN LEVEL")
    Dim wsSumber As Worksheet
    Dim kolomMasa As Long
    kolomMasa = 7 + bulan
    If UCase$(taxKey) = "PPNDN" Then
        ws.Range("AE40").Formula = "=SUM(AE15:AE39)"
        ws.Range("AE41").Formula = "=SUM(AE18,AE21,AE24,AE27,AE30,AE33,AE36,AE39)"
        ws.Range("AE43").Formula = "=SUM(AE40,AE42)"
        ws.Range("AE70").Formula = "=SUM(AE58,AE60,AE63,AE65,AE72,AE75)"
        ws.Range("AE71").Formula = "=SUM(AE59,AE62,AE64,AE67,AE69,AE74)"
        ws.Range("AE76").Formula = "=SUM(AE70,AE71)"
        ws.Range("AE86").Formula = "=SUM(AE81:AE85)"
        ws.Range("AE92").Formula = "=SUM(AE88:AE91)"
        ws.Range("AE97").Formula = "=SUM(AE86,AE92,AE94:AE95)"
        ws.Range("AE98").Formula = "=AE79-AE97"
    ElseIf UCase$(taxKey) = "PPH26" Then
        ws.Range("AD18").Formula = "=SUM(AD16:AD17)"
        ws.Range("AD22").Formula = "=SUM(AD20:AD21)"
        ws.Range("AD39").Formula = "=SUM(AD36:AD38)"
        ws.Range("AD45").Formula = "=SUM(AD42:AD44)"
        ws.Range("AD46").Formula = "=AD22-AD25-AD39-AD45"
    ElseIf UCase$(taxKey) = "P21" Or UCase$(taxKey) = "PPH21" _
        Or UCase$(taxKey) = "P21F" Or UCase$(taxKey) = "PPH21F" _
        Or UCase$(taxKey) = "P23" Or UCase$(taxKey) = "PPH23" Then

        Set wsSumber = SheetSumberCFMWP054(taxKey)
        If IsNumeric(wsSumber.Cells(113, kolomMasa).Value2) Then
            ws.Range("AD42").Value2 = CDbl(wsSumber.Cells(113, kolomMasa).Value2)
        End If
        If IsNumeric(wsSumber.Cells(114, kolomMasa).Value2) Then
            ws.Range("AD43").Value2 = CDbl(wsSumber.Cells(114, kolomMasa).Value2)
        End If
        If IsNumeric(wsSumber.Cells(115, kolomMasa).Value2) Then
            ws.Range("AD44").Value2 = CDbl(wsSumber.Cells(115, kolomMasa).Value2)
        End If
    Else
        ws.Range("AD17").Formula = "=AD16"
        ws.Range("AD42").Formula = "=SUM(AD39:AD41)"
        ws.Range("AD44").Formula = "=AD23-AD42"
    End If
End Sub
