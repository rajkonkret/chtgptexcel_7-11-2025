' Wklej do: konkretny Sheet (np. Arkusz1)
Private Sub Worksheet_Change(ByVal Target As Range)
    If Intersect(Target, Me.Range("C2:C1000")) Is Nothing Then Exit Sub
    On Error GoTo SafeExit
    Application.EnableEvents = False
    Dim c As Range
    For Each c In Intersect(Target, Me.Range("C2:C1000")).Cells
        If Not IsNumeric(c.Value) Or c.Value < 0 Then
            c.Interior.Color = RGB(255, 200, 200)
        Else
            c.Interior.ColorIndex = xlColorIndexNone
        End If
    Next c
SafeExit:
    Application.EnableEvents = True
End Sub