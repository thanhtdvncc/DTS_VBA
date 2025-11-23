Attribute VB_Name = "m06_SAP2000_PatternValidation"
Option Explicit
'===============================================================
' Module: modSAP2000_PatternValidation
' Purpose: Extract load patterns & build validation list (range-based)
' Notes  :
'   - Builds a workbook-level named range "PatternList" containing
'     all load pattern names from the model.
'   - Avoids Excel 255-char validation limit by using a named range.
'===============================================================

Public Sub WritePatterns()
    ' Optional helper: write raw pattern list to a dedicated sheet "Patterns"
    If SapModel Is Nothing Then
        LogMsg "WritePatterns: SapModel is not connected."
        Exit Sub
    End If
    
    Dim ws As Worksheet
    Set ws = SheetOrCreate("Patterns", True)
    If ws Is Nothing Then
        LogMsg "WritePatterns: Cannot access Patterns sheet."
        Exit Sub
    End If
    ws.Cells.Clear
    
    Dim NumberPatterns As Long
    Dim patternNames() As String
    Dim ret As Long
    
    ret = SapModel.loadPatterns.GetNameList(NumberPatterns, patternNames)
    If ret <> 0 Or NumberPatterns = 0 Then
        ws.Cells(1, 1).Value = "(No load patterns found)"
        LogMsg "WritePatterns: No load patterns found or error."
        Exit Sub
    End If
    
    ws.Range("A1").Resize(NumberPatterns, 1).Value = ToVerticalVariant(patternNames)
    ws.Columns("A").AutoFit
    LogMsg "WritePatterns: Wrote " & NumberPatterns & " patterns to sheet 'Patterns'."
End Sub

Public Sub BuildPatternValidation()
    ' Build a helper sheet "PatternList" containing one pattern per row
    ' and create a workbook-level named range "PatternList" pointing to that range.
    If SapModel Is Nothing Then
        LogMsg "BuildPatternValidation: SapModel is not connected."
        Exit Sub
    End If
    
    Dim NumberPatterns As Long
    Dim patternNames() As String
    Dim ret As Long
    
    On Error Resume Next
    ret = SapModel.loadPatterns.GetNameList(NumberPatterns, patternNames)
    If err.number <> 0 Or ret <> 0 Then
        LogMsg "BuildPatternValidation: Failed to get load pattern list. ret=" & ret & " err=" & err.number
        err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0
    
    If NumberPatterns = 0 Or IsArrayEmpty(patternNames) Then
        LogMsg "BuildPatternValidation: No patterns to build."
        Exit Sub
    End If
    
    ' Remove duplicates and blank entries
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Dim i As Long, nm As String
    For i = LBound(patternNames) To UBound(patternNames)
        nm = CStr(patternNames(i))
        If Len(Trim(nm)) > 0 Then
            If Not dict.exists(nm) Then dict.Add nm, 1
        End If
    Next i
    
    If dict.count = 0 Then
        LogMsg "BuildPatternValidation: No valid pattern names."
        Exit Sub
    End If
    
    Dim arrPatterns() As Variant
    arrPatterns = dict.keys ' 0-based array
    
    ' Write to helper sheet PatternList
    Dim wsList As Worksheet
    Set wsList = SheetOrCreate("PatternList", True)
    If wsList Is Nothing Then
        LogMsg "BuildPatternValidation: Cannot create/access PatternList sheet."
        Exit Sub
    End If
    
    wsList.Cells.Clear
    wsList.Range("A1").Resize(UBound(arrPatterns) - LBound(arrPatterns) + 1, 1).Value = ToVerticalVariant(arrPatterns)
    
    Dim lastRow As Long
    lastRow = UBound(arrPatterns) - LBound(arrPatterns) + 1
    If lastRow <= 0 Then
        LogMsg "BuildPatternValidation: Nothing written to PatternList sheet."
        Exit Sub
    End If
    
    ' Create or update workbook-level named range "PatternList"
    Dim refStr As String
    refStr = "=" & wsList.Name & "!$A$1:$A$" & lastRow
    
    Dim nmObj As Name
    Dim found As Boolean
    found = False
    For Each nmObj In ThisWorkbook.names
        If StrComp(nmObj.Name, "PatternList", vbTextCompare) = 0 Then
            nmObj.RefersTo = refStr
            found = True
            Exit For
        End If
    Next nmObj
    If Not found Then
        ThisWorkbook.names.Add Name:="PatternList", RefersTo:=refStr
    End If
    
    ' Also support legacy misspelling "PartternList" if desired (create/update)
    On Error Resume Next
    Dim legacyName As Name
    Set legacyName = Nothing
    For Each legacyName In ThisWorkbook.names
        If StrComp(legacyName.Name, "PartternList", vbTextCompare) = 0 Then
            legacyName.RefersTo = refStr
            Set legacyName = Nothing
            Exit For
        End If
    Next legacyName
    If err.number = 0 Then
        ' if the workbook does not already have PartternList, create it silently
        Dim existsLegacy As Boolean
        existsLegacy = False
        For Each nmObj In ThisWorkbook.names
            If StrComp(nmObj.Name, "PartternList", vbTextCompare) = 0 Then
                existsLegacy = True
                Exit For
            End If
        Next nmObj
        If Not existsLegacy Then
            On Error Resume Next
            ThisWorkbook.names.Add Name:="PartternList", RefersTo:=refStr
            On Error GoTo 0
        End If
    End If
    On Error GoTo 0
    
    ' Optionally hide helper sheet
    On Error Resume Next
    wsList.Visible = xlSheetHidden
    On Error GoTo 0
    
    LogMsg "BuildPatternValidation: Created named range PatternList with " & lastRow & " entries."
End Sub
