Option Explicit

' === Merge all .xlsx files from a folder into current workbook (sheet "Merged") ===
' Wklej do: Module
' Użycie: Alt+F8 -> MergeFilesFromFolder
' Wymaga: zaufany folder, pliki bez haseł
Sub MergeFilesFromFolder()
    Dim FSO As Object, folderPath As String, f As Object
    Dim wbSrc As Workbook, wsSrc As Worksheet, wsDst As Worksheet
    Dim lastRow As Long, nextRow As Long

    On Error GoTo CleanFail
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    folderPath = GetFolderPath()
    If Len(folderPath) = 0 Then GoTo CleanExit

    Set wsDst = PrepareDestinationSheet(ThisWorkbook, "Merged")
    Set FSO = CreateObject("Scripting.FileSystemObject")

    nextRow = 2
    wsDst.Range("A1:E1").Value = Array("Plik", "Arkusz", "A", "B", "C") ' <-- dostosuj nagłówki

    Dim file As Object
    For Each file In FSO.GetFolder(folderPath).Files
        If LCase$(FSO.GetExtensionName(file.Name)) = "xlsx" Then
            Set wbSrc = Workbooks.Open(file.Path, ReadOnly:=True)
            For Each wsSrc In wbSrc.Worksheets
                lastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row
                If lastRow >= 2 Then
                    With wsDst
                        .Range(.Cells(nextRow, 1), .Cells(nextRow + (lastRow - 1) - 1, 1)).Value = file.Name
                        .Range(.Cells(nextRow, 2), .Cells(nextRow + (lastRow - 1) - 1, 2)).Value = wsSrc.Name
                        .Range(.Cells(nextRow, 3), .Cells(nextRow + (lastRow - 1) - 1, 3)).Value = wsSrc.Range("A2:A" & lastRow).Value
                        .Range(.Cells(nextRow, 4), .Cells(nextRow + (lastRow - 1) - 1, 4)).Value = wsSrc.Range("B2:B" & lastRow).Value
                        .Range(.Cells(nextRow, 5), .Cells(nextRow + (lastRow - 1) - 1, 5)).Value = wsSrc.Range("C2:C" & lastRow).Value
                    End With
                    nextRow = nextRow + (lastRow - 1)
                End If
            Next wsSrc
            wbSrc.Close SaveChanges:=False
        End If
    Next file

CleanExit:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub
CleanFail:
    MsgBox "Błąd: " & Err.Description, vbExclamation
    Resume CleanExit
End Sub

Private Function GetFolderPath() As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Wybierz folder z plikami XLSX"
        If .Show = -1 Then GetFolderPath = .SelectedItems(1)
    End With
End Function

Private Function PrepareDestinationSheet(wb As Workbook, name As String) As Worksheet
    On Error Resume Next
    Application.DisplayAlerts = False
    wb.Worksheets(name).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    Set PrepareDestinationSheet = wb.Worksheets.Add
    PrepareDestinationSheet.Name = name
End Function