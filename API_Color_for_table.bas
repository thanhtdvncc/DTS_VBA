Attribute VB_Name = "API_Color_for_table"
Option Explicit
' Sub to color rows by the value in a user-specified key column.
' Auto-detect: If selection includes a contiguous range starting at row 2 (>=4 columns wide),
' use that as color range. If a full column is also selected, use it as key column.
Public Sub ColorRowsByPattern()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim keyValue As String
    Dim dict As Object ' Scripting.Dictionary
    Dim idx As Long
    Dim colorValue As Long
    Dim keyCount As Long
    Dim keyColInput As String
    Dim keyCol As Long
    Dim rangeInput As String
    Dim startCol As Long, endCol As Long
    Dim retMsg As String
    Dim sel As Range
    Dim autoRangeDetected As Boolean
    Dim colorRange As Range, keyRange As Range
    Dim area As Range
    Set ws = ActiveSheet
    Set sel = Selection
    autoRangeDetected = False
    Set colorRange = Nothing
    Set keyRange = Nothing
    ' --- Auto detect ---
    If Not sel Is Nothing Then
        For Each area In sel.Areas
            ' Detect color range: must start at row 2, >=4 columns, any number of rows (but we'll use only row 2 headers for range)
            If area.row = 2 And area.Columns.count >= 4 Then
                If colorRange Is Nothing Then
                    Set colorRange = area
                Else
                    ' If multiple qualifying areas, merge if contiguous horizontally
                    If area.Column = colorRange.Column + colorRange.Columns.count Then
                        Set colorRange = ws.Range(colorRange, area)
                    ElseIf area.Column + area.Columns.count - 1 = colorRange.Column - 1 Then
                        Set colorRange = ws.Range(area, colorRange)
                    End If
                End If
            ElseIf area.Columns.count = 1 And area.rows.count > 1 Then
                ' Detect key column: full column selection
                If keyRange Is Nothing Then
                    Set keyRange = area
                End If
            End If
        Next area
        
        If Not colorRange Is Nothing Then
            autoRangeDetected = True
            startCol = colorRange.Column
            endCol = startCol + colorRange.Columns.count - 1
            If Not keyRange Is Nothing Then
                keyCol = keyRange.Column
            Else
                keyCol = startCol ' Default to first column
            End If
        End If
    End If
    ' --- Manual input mode if not auto detected ---
    If Not autoRangeDetected Then
        keyColInput = InputBox("Enter key column letter or number to group by (e.g. B or 2):" & vbCrLf & _
                               "You can select range from A2 and Key column first to process immediately!", _
                               "Key Column", "B")
        If Len(Trim(keyColInput)) = 0 Then Exit Sub
        keyColInput = Trim(keyColInput)
        If IsNumeric(keyColInput) Then
            keyCol = CLng(keyColInput)
            If keyCol < 1 Or keyCol > 16384 Then
                MsgBox "Invalid column number.", vbCritical, "ColorRowsByPattern"
                Exit Sub
            End If
        Else
            keyCol = ColumnLetterToNumber(keyColInput)
            If keyCol = 0 Then
                MsgBox "Invalid column letter.", vbCritical, "ColorRowsByPattern"
                Exit Sub
            End If
        End If
        rangeInput = InputBox("Enter column range to color (e.g. A-G or 1-7). You can also use A:G or 1:7. Default A-G:", "Fill Column Range", "A-G")
        If Len(Trim(rangeInput)) = 0 Then Exit Sub
        rangeInput = Trim(rangeInput)
        If Not ParseColumnRange(rangeInput, startCol, endCol) Then
            MsgBox "Invalid column range format. Use examples like A-G or 1-7 or A:G or 1:7.", vbCritical, "ColorRowsByPattern"
            Exit Sub
        End If
        If startCol < 1 Or endCol > 16384 Or startCol > endCol Then
            MsgBox "Invalid column range values.", vbCritical, "ColorRowsByPattern"
            Exit Sub
        End If
    End If
    ' --- Determine last row based on key column ---
    lastRow = ws.Cells(ws.rows.count, keyCol).End(xlUp).row
    If lastRow < 2 Then
        MsgBox "No data found in the selected column starting at row 2.", vbInformation, "ColorRowsByPattern"
        Exit Sub
    End If
    ' --- Process coloring ---
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    For r = 2 To lastRow
        keyValue = Trim(CStr(ws.Cells(r, keyCol).Value))
        If Len(keyValue) = 0 Then
            ws.Range(ws.Cells(r, startCol), ws.Cells(r, endCol)).Interior.color = RGB(242, 242, 242)
        Else
            If Not dict.exists(keyValue) Then
                keyCount = dict.count + 1
                dict.Add keyValue, keyCount
            End If
            idx = dict(keyValue)
            colorValue = GetPastelColorByIndex(idx)
            ws.Range(ws.Cells(r, startCol), ws.Cells(r, endCol)).Interior.color = colorValue
        End If
    Next r
    If autoRangeDetected Then
        MsgBox "Auto range detected: " & ws.Cells(2, startCol).Address(False, False) & ":" & ws.Cells(2, endCol).Address(False, False) & _
                vbCrLf & "Key column: " & ColumnNumberToLetter(keyCol), vbInformation, "ColorRowsByPattern"
    End If
End Sub
' Parse a column range input like "A-G", "A:G", "1-7", "1:7", or single "B" or "2".
Private Function ParseColumnRange(ByVal inputStr As String, ByRef startCol As Long, ByRef endCol As Long) As Boolean
    Dim parts() As String
    Dim a As String, b As String
    inputStr = Trim(inputStr)
    ParseColumnRange = False
    If InStr(inputStr, ":") > 0 Then
        parts = Split(inputStr, ":")
    ElseIf InStr(inputStr, "-") > 0 Then
        parts = Split(inputStr, "-")
    Else
        a = inputStr
        If IsNumeric(a) Then
            startCol = CLng(a)
            endCol = startCol
            ParseColumnRange = True
            Exit Function
        Else
            startCol = ColumnLetterToNumber(a)
            If startCol > 0 Then
                endCol = startCol
                ParseColumnRange = True
                Exit Function
            Else
                Exit Function
            End If
        End If
    End If
    If UBound(parts) <> 1 Then Exit Function
    a = Trim(parts(0))
    b = Trim(parts(1))
    If Len(a) = 0 Or Len(b) = 0 Then Exit Function
    If IsNumeric(a) Then
        startCol = CLng(a)
    Else
        startCol = ColumnLetterToNumber(a)
    End If
    If IsNumeric(b) Then
        endCol = CLng(b)
    Else
        endCol = ColumnLetterToNumber(b)
    End If
    If startCol = 0 Or endCol = 0 Then Exit Function
    If startCol > endCol Then
        Dim tmp As Long
        tmp = startCol
        startCol = endCol
        endCol = tmp
    End If
    ParseColumnRange = True
End Function
' Convert column letter (e.g. "A", "B", "AA") to numeric index. Returns 0 if invalid.
Private Function ColumnLetterToNumber(ByVal colLetter As String) As Long
    Dim i As Long, ch As String, result As Long, pos As Long
    colLetter = Trim(UCase(colLetter))
    If Len(colLetter) = 0 Then Exit Function
    For i = 1 To Len(colLetter)
        ch = mid$(colLetter, i, 1)
        If ch < "A" Or ch > "Z" Then Exit Function
        pos = Asc(ch) - Asc("A") + 1
        result = result * 26 + pos
    Next i
    ColumnLetterToNumber = result
End Function
' Convert column number to letter (e.g. 1 -> "A")
Private Function ColumnNumberToLetter(ByVal colNum As Long) As String
    Dim n As Long, s As String, r As Long
    If colNum < 1 Then Exit Function
    n = colNum
    Do While n > 0
        r = (n - 1) Mod 26
        s = Chr(65 + r) & s
        n = Int((n - 1) / 26)
    Loop
    ColumnNumberToLetter = s
End Function
' Generate a pastel color (RGB Long) for a given index.
Private Function GetPastelColorByIndex(ByVal Index As Long) As Long
    Dim hue As Double, saturation As Double, lightness As Double
    hue = (Index * 137.5) Mod 360
    saturation = 0.45
    lightness = 0.85
    GetPastelColorByIndex = HslToRgbLong(hue, saturation, lightness)
End Function
' Convert HSL (Hue 0-360, S 0-1, L 0-1) to VBA Long color (RGB).
Private Function HslToRgbLong(ByVal h As Double, ByVal s As Double, ByVal l As Double) As Long
    Dim c As Double, hPrime As Double, X As Double, m As Double
    Dim r1 As Double, G1 As Double, B1 As Double
    Dim seg As Long, mod2 As Double
    Dim r As Long, g As Long, b As Long
    h = h Mod 360
    If h < 0 Then h = h + 360
    c = (1 - Abs(2 * l - 1)) * s
    hPrime = h / 60#
    mod2 = hPrime - 2 * Int(hPrime / 2)
    X = c * (1 - Abs(mod2 - 1))
    seg = Int(hPrime)
    Select Case seg
        Case 0: r1 = c: G1 = X: B1 = 0
        Case 1: r1 = X: G1 = c: B1 = 0
        Case 2: r1 = 0: G1 = c: B1 = X
        Case 3: r1 = 0: G1 = X: B1 = c
        Case 4: r1 = X: G1 = 0: B1 = c
        Case Else: r1 = c: G1 = 0: B1 = X
    End Select
    m = l - c / 2
    r = CLng(Round((r1 + m) * 255, 0))
    g = CLng(Round((G1 + m) * 255, 0))
    b = CLng(Round((B1 + m) * 255, 0))
    ' Clamp values to valid RGB range - Corrected syntax
    If r < 0 Then r = 0
    If r > 255 Then r = 255
    If g < 0 Then g = 0
    If g > 255 Then g = 255
    If b < 0 Then b = 0
    If b > 255 Then b = 255
    HslToRgbLong = RGB(r, g, b)
End Function

