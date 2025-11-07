Attribute VB_Name = "Module1"
Option Explicit

Public Sub CleanDates_WithSlashes()
    Dim ws As Worksheet
    Dim firstRow As Long, lastRow As Long
    Dim srcCol As Long, dstCol As Long
    Dim r As Long, v, d

    Set ws = ActiveSheet           ' ustaw na w³aœciwy arkusz, jeœli chcesz na sztywno
    srcCol = 5                     ' E
    dstCol = 7                     ' G
    firstRow = 6

    ' Ustal ostatni wiersz po kolumnie E
    lastRow = ws.Cells(ws.Rows.Count, srcCol).End(xlUp).Row
    If lastRow < firstRow Then
        MsgBox "Brak danych w kolumnie E od wiersza 6.", vbExclamation
        Exit Sub
    End If

    ' Nag³ówek i format docelowy z ukoœnikami (nie kropkami)
    ws.Cells(firstRow - 1, dstCol).Value = "Data (dd/mm/yyyy)"
    ' Dwie równowa¿ne opcje formatu – wybierz jedn¹:
    ws.Columns(dstCol).NumberFormat = "dd\/mm\/yyyy"        ' neutralne
    'ws.Columns(dstCol).NumberFormatLocal = "dd\/mm\/rrrr"  ' wariant PL

    ' Przelicz daty
    For r = firstRow To lastRow
        v = ws.Cells(r, srcCol).Value
        d = ParsePolishDate(v)
        If IsDate(d) Then
            ws.Cells(r, dstCol).Value = CDate(d)            ' prawdziwa data
        Else
            ws.Cells(r, dstCol).ClearContents               ' nie rozpoznano -> pusto
        End If
    Next r

    ws.Columns(dstCol).AutoFit
    MsgBox "Gotowe. Nowe daty w kolumnie G, format dd/mm/yyyy.", vbInformation
End Sub

'---------------------- PARSER DAT ----------------------

Private Function ParsePolishDate(ByVal v As Variant) As Variant
    ' Zwraca Date, jeœli siê uda³o, albo Empty, jeœli nie.
    Dim s As String
    Dim parts() As String
    Dim y As Long, m As Long, d As Long

    If IsDate(v) Then
        ParsePolishDate = CDate(v)
        Exit Function
    End If

    s = Trim$(CStr(v))
    If Len(s) = 0 Then Exit Function

    ' Uproœæ: ma³e litery, zamieñ polskie znaki, zamieñ miesi¹ce s³owne na MM
    s = LCase$(s)
    s = ReplacePolChars(s)
    s = NormalizeMonthTokensToMM(s)

    ' Ujednolicenie separatorów do "/"
    s = Replace(s, ".", "/")
    s = Replace(s, "-", "/")
    s = Replace(s, " ", "/")
    Do While InStr(s, "//") > 0
        s = Replace(s, "//", "/")
    Loop
    s = Trim$(s)
    If Left$(s, 1) = "/" Then s = Mid$(s, 2)
    If Right$(s, 1) = "/" Then s = Left$(s, Len(s) - 1)

    ' Teraz oczekujemy jednego z: dd/mm/rrrr, dd/mm/rr, rrrr/mm/dd
    parts = Split(s, "/")
    If UBound(parts) <> 2 Then Exit Function
    If parts(0) = "" Or parts(1) = "" Or parts(2) = "" Then Exit Function
    If Not (IsNumeric(parts(0)) And IsNumeric(parts(1)) And IsNumeric(parts(2))) Then Exit Function

    If Len(parts(0)) = 4 Then
        ' rrrr/mm/dd
        y = CLng(parts(0)): m = CLng(parts(1)): d = CLng(parts(2))
    Else
        ' dd/mm/rrrr lub dd/mm/rr
        d = CLng(parts(0)): m = CLng(parts(1)): y = CLng(parts(2))
        If y < 100 Then
            y = IIf(y < 30, 2000 + y, 1900 + y)
        End If
    End If

    If Not IsValidYMD(y, m, d) Then Exit Function

    On Error Resume Next
    ParsePolishDate = DateSerial(y, m, d)
    If Err.Number <> 0 Then
        ParsePolishDate = Empty
        Err.Clear
    End If
    On Error GoTo 0
End Function

Private Function ReplacePolChars(ByVal s As String) As String
    s = Replace(s, "¹", "a")
    s = Replace(s, "æ", "c")
    s = Replace(s, "ê", "e")
    s = Replace(s, "³", "l")
    s = Replace(s, "ñ", "n")
    s = Replace(s, "ó", "o")
    s = Replace(s, "œ", "s")
    s = Replace(s, "Ÿ", "z")
    s = Replace(s, "¿", "z")
    ReplacePolChars = s
End Function

Private Function NormalizeMonthTokensToMM(ByVal s As String) As String
    ' Zast¹p polskie skróty/nazwy miesiêcy numerami "01"..."12".
    ' Obs³uga kropek po skrótach.
    Dim i As Long
    Dim key As Variant, full As Variant
    Dim mm As String

    ' usuñ kropki po skrótach (np. "kwi.")
    s = Replace(s, "sty.", "sty")
    s = Replace(s, "lut.", "lut")
    s = Replace(s, "mar.", "mar")
    s = Replace(s, "kwi.", "kwi")
    s = Replace(s, "maj.", "maj")
    s = Replace(s, "cze.", "cze")
    s = Replace(s, "lip.", "lip")
    s = Replace(s, "sie.", "sie")
    s = Replace(s, "wrz.", "wrz")
    s = Replace(s, "paz.", "paz")
    s = Replace(s, "lis.", "lis")
    s = Replace(s, "gru.", "gru")

    Dim shortMon As Variant, longMon As Variant
    shortMon = Array("sty", "lut", "mar", "kwi", "maj", "cze", "lip", "sie", "wrz", "paz", "lis", "gru")
    longMon = Array("styczen", "luty", "marzec", "kwiecien", "maj", "czerwiec", "lipiec", "sierpien", "wrzesien", "pazdziernik", "listopad", "grudzien")

    For i = 0 To 11
        mm = Format$(i + 1, "00")
        s = ReplaceToken(s, shortMon(i), mm)
        s = ReplaceToken(s, longMon(i), mm)
    Next i

    NormalizeMonthTokensToMM = s
End Function

Private Function ReplaceToken(ByVal text As String, ByVal token As String, ByVal repl As String) As String
    ' Zamiana tokenu jako samodzielnego cz³onu (prosto, bez RegExp).
    Dim t As String: t = " " & text & " "
    t = Replace(t, " " & token & " ", " " & repl & " ")
    t = Replace(t, "/" & token & "/", "/" & repl & "/")
    t = Replace(t, "-" & token & "-", "-" & repl & "-")
    t = Replace(t, "." & token & ".", "." & repl & ".")
    t = Replace(t, " " & token & "/", " " & repl & "/")
    t = Replace(t, "/" & token & " ", "/" & repl & " ")
    t = Replace(t, " " & token & "-", " " & repl & "-")
    t = Replace(t, "-" & token & " ", "-" & repl & " ")
    t = Replace(t, "." & token & " ", "." & repl & " ")
    t = Replace(t, " " & token & ".", " " & repl & ".")
    ReplaceToken = Mid$(t, 2, Len(t) - 2)
End Function

Private Function IsValidYMD(ByVal y As Long, ByVal m As Long, ByVal d As Long) As Boolean
    IsValidYMD = (y >= 1900 And y <= 9999 And m >= 1 And m <= 12 And d >= 1 And d <= 31)
End Function

