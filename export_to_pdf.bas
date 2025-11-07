Attribute VB_Name = "ExportToPDF"
Sub ExportToPDF()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Dane")

    Dim savePath As String
    savePath = ThisWorkbook.Path & "\Raport_" & Format(Now(), "yyyymmdd_hhmmss") & ".pdf"

    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=savePath, Quality:=xlQualityStandard
    MsgBox "Raport zapisany jako PDF: " & savePath
End Sub