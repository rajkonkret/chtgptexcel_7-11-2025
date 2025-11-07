Option Explicit
' Wklej do: Module
Public Function MedianNoZeros(rng As Range) As Double
    Dim arr() As Double, c As Range, n As Long
    ReDim arr(1 To rng.Count)
    For Each c In rng.Cells
        If IsNumeric(c.Value) Then
            If CDbl(c.Value) <> 0 Then
                n = n + 1: arr(n) = CDbl(c.Value)
            End If
        End If
    Next c
    If n = 0 Then MedianNoZeros = 0: Exit Function
    ReDim Preserve arr(1 To n)
    MedianNoZeros = Application.WorksheetFunction.Median(arr)
End Function