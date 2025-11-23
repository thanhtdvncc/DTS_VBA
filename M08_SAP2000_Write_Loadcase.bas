Attribute VB_Name = "M08_SAP2000_Write_Loadcase"
Option Explicit
'===============================================================
' Module: modSAP2000_LoadCases
' Purpose: Write load patterns, load cases and load combos
'===============================================================

Public Sub WriteLoadCases()
    ' Check if SapModel is connected
    If SapModel Is Nothing Then
        LogMsg "WriteLoadCases: SapModel is not connected."
        Exit Sub
    End If
    
    ' Write all load patterns, cases and combinations to sheet
    Dim ws As Worksheet
    Set ws = SheetOrCreate("Loadcase", True)
    If ws Is Nothing Then
        LogMsg "WriteLoadCases: Cannot access Loadcase sheet."
        Exit Sub
    End If
    
    ' Clear existing data AND formatting
    ws.Cells.Clear
    
    Dim rowPtr As Long
    rowPtr = 1
    
    ' ========== SECTION 1: LOAD PATTERNS ==========
    ws.Cells(rowPtr, "A").Value = "LOAD PATTERNS"
    rowPtr = rowPtr + 1
    
    ws.Range("A" & rowPtr & ":C" & rowPtr).Value = Array("Pattern Name", "Load Type", "Self Weight Multiplier")
    rowPtr = rowPtr + 1
    
    ' Get all load pattern names
    Dim NumberPatterns As Long
    Dim patternNames() As String
    Dim ret As Long
    
    ret = SapModel.loadPatterns.GetNameList(NumberPatterns, patternNames)
    
    If ret = 0 And NumberPatterns > 0 Then
        Dim i As Long
        For i = 0 To NumberPatterns - 1
            Dim patternName As String
            patternName = patternNames(i)
            
            ' Get load type
            Dim loadType As Long
            Dim loadTypeStr As String
            loadTypeStr = ""
            
            On Error Resume Next
            ret = SapModel.loadPatterns.GetLoadType(patternName, loadType)
            If err.number = 0 And ret = 0 Then
                loadTypeStr = GetLoadPatternTypeString(loadType)
            End If
            err.Clear
            On Error GoTo 0
            
            ' Get self weight multiplier
            Dim selfWtMult As Double
            selfWtMult = 0
            
            On Error Resume Next
            ret = SapModel.loadPatterns.GetSelfWTMultiplier(patternName, selfWtMult)
            If err.number <> 0 Then
                selfWtMult = 0
                err.Clear
            End If
            On Error GoTo 0
            
            ' Write to sheet
            ws.Cells(rowPtr, "A").Value = patternName
            ws.Cells(rowPtr, "B").Value = loadTypeStr
            ws.Cells(rowPtr, "C").Value = selfWtMult
            ws.Cells(rowPtr, "C").NumberFormat = "0.00"
            
            rowPtr = rowPtr + 1
        Next i
        
        LogMsg "Load patterns written: " & NumberPatterns & " patterns."
    Else
        ws.Cells(rowPtr, "A").Value = "(No load patterns found)"
        rowPtr = rowPtr + 1
        LogMsg "No load patterns found."
    End If
    
    ' Add spacing
    rowPtr = rowPtr + 2
    
    ' ========== SECTION 2: LOAD CASES ==========
    ws.Cells(rowPtr, "A").Value = "LOAD CASES"
    rowPtr = rowPtr + 1
    
    ws.Range("A" & rowPtr & ":D" & rowPtr).Value = Array("Load Case Name", "Case Type", "Design Type", "Notes")
    rowPtr = rowPtr + 1
    
    ' Get all load case names
    Dim NumberNames As Long
    Dim CaseNames() As String
    
    ret = SapModel.LoadCases.GetNameList(NumberNames, CaseNames)
    
    If ret <> 0 Or NumberNames = 0 Then
        ws.Cells(rowPtr, "A").Value = "(No load cases found)"
        LogMsg "WriteLoadCases: No load cases found or error occurred."
        rowPtr = rowPtr + 1
    Else
        ' Write load case data
        For i = 0 To NumberNames - 1
            Dim caseName As String
            caseName = CaseNames(i)
            
            ' Get load case type
            Dim caseTypeStr As String
            caseTypeStr = GetLoadCaseType(caseName)
            
            ' Get design type
            Dim designType As Long, designTypeOption As Long
            Dim designTypeStr As String
            designTypeStr = ""
            
            On Error Resume Next
            ret = SapModel.LoadCases.GetDesignType(caseName, designType, designTypeOption)
            If err.number = 0 And ret = 0 Then
                designTypeStr = GetDesignTypeString(designType)
            End If
            err.Clear
            On Error GoTo 0
            
            ' Get notes
            Dim notes As String, guid As String
            notes = ""
            guid = ""
            
            On Error Resume Next
            ret = SapModel.LoadCases.GetNotes(caseName, notes, guid)
            If err.number <> 0 Then
                notes = ""
                err.Clear
            End If
            On Error GoTo 0
            
            ' Write to sheet
            ws.Cells(rowPtr, "A").Value = caseName
            ws.Cells(rowPtr, "B").Value = caseTypeStr
            ws.Cells(rowPtr, "C").Value = designTypeStr
            ws.Cells(rowPtr, "D").Value = notes
            
            rowPtr = rowPtr + 1
        Next i
        
        LogMsg "Load cases written: " & NumberNames & " cases."
    End If
    
    ' Add spacing
    rowPtr = rowPtr + 2
    
    ' ========== SECTION 3: LOAD COMBINATIONS ==========
    ws.Cells(rowPtr, "A").Value = "LOAD COMBINATIONS"
    rowPtr = rowPtr + 1
    
    ws.Range("A" & rowPtr & ":F" & rowPtr).Value = _
        Array("Combo Name", "Combo Type", "Case/Combo Name", "Type", "Scale Factor", "Notes")
    rowPtr = rowPtr + 1
    
    ' Get all load combinations
    Dim NumberCombos As Long
    Dim ComboNames() As String
    
    ret = SapModel.RespCombo.GetNameList(NumberCombos, ComboNames)
    
    If ret = 0 And NumberCombos > 0 Then
        Call WriteLoadCombinations(ws, rowPtr, ComboNames, NumberCombos)
        LogMsg "Load combinations written: " & NumberCombos & " combinations."
    Else
        ws.Cells(rowPtr, "A").Value = "(No load combinations found)"
        LogMsg "No load combinations found."
    End If
    
    ' Auto-fit columns
    ws.Columns("A:F").AutoFit
    
    ' Summary message
    Dim summaryMsg As String
    summaryMsg = "Export completed:" & vbCrLf & _
                 "- " & NumberPatterns & " load patterns" & vbCrLf & _
                 "- " & NumberNames & " load cases" & vbCrLf & _
                 "- " & NumberCombos & " load combinations"
    
    LogMsg summaryMsg
End Sub

Private Function GetLoadPatternTypeString(loadType As Long) As String
    ' Convert load pattern type enum to string
    Select Case loadType
        Case 1: GetLoadPatternTypeString = "Dead"
        Case 2: GetLoadPatternTypeString = "SuperDead"
        Case 3: GetLoadPatternTypeString = "Live"
        Case 4: GetLoadPatternTypeString = "ReduceLive"
        Case 5: GetLoadPatternTypeString = "Quake"
        Case 6: GetLoadPatternTypeString = "Wind"
        Case 7: GetLoadPatternTypeString = "Snow"
        Case 8: GetLoadPatternTypeString = "Other"
        Case 9: GetLoadPatternTypeString = "Move"
        Case 10: GetLoadPatternTypeString = "Temperature"
        Case 11: GetLoadPatternTypeString = "RoofLive"
        Case 12: GetLoadPatternTypeString = "Notional"
        Case 13: GetLoadPatternTypeString = "PatternLive"
        Case 14: GetLoadPatternTypeString = "Wave"
        Case 15: GetLoadPatternTypeString = "Braking"
        Case 16: GetLoadPatternTypeString = "Centrifugal"
        Case 17: GetLoadPatternTypeString = "Friction"
        Case 18: GetLoadPatternTypeString = "Ice"
        Case 19: GetLoadPatternTypeString = "WindOnLiveLoad"
        Case 20: GetLoadPatternTypeString = "HorizontalEarthPressure"
        Case 21: GetLoadPatternTypeString = "VerticalEarthPressure"
        Case 22: GetLoadPatternTypeString = "EarthSurcharge"
        Case 23: GetLoadPatternTypeString = "DownDrag"
        Case 24: GetLoadPatternTypeString = "VehicleCollision"
        Case 25: GetLoadPatternTypeString = "VesselCollision"
        Case 26: GetLoadPatternTypeString = "TemperatureGradient"
        Case 27: GetLoadPatternTypeString = "Settlement"
        Case 28: GetLoadPatternTypeString = "Shrinkage"
        Case 29: GetLoadPatternTypeString = "Creep"
        Case 30: GetLoadPatternTypeString = "WaterLoadPressure"
        Case 31: GetLoadPatternTypeString = "LiveLoadSurcharge"
        Case 32: GetLoadPatternTypeString = "LockedInForces"
        Case 33: GetLoadPatternTypeString = "PedestrianLL"
        Case 34: GetLoadPatternTypeString = "Prestress"
        Case 35: GetLoadPatternTypeString = "Hyperstatic"
        Case 36: GetLoadPatternTypeString = "Bouyancy"
        Case 37: GetLoadPatternTypeString = "StreamFlow"
        Case 38: GetLoadPatternTypeString = "Impact"
        Case 39: GetLoadPatternTypeString = "Construction"
        Case Else: GetLoadPatternTypeString = "Unknown(" & loadType & ")"
    End Select
End Function

Private Function GetLoadCaseType(caseName As String) As String
    ' Try to determine case type by checking which GetNameList returns it
    If SapModel Is Nothing Then
        GetLoadCaseType = "Unknown"
        Exit Function
    End If
    
    Dim NumberNames As Long
    Dim MyName() As String
    Dim ret As Long
    
    ' Try each case type
    Dim caseTypes As Variant
    Dim typeNames As Variant
    Dim i As Long
    
    caseTypes = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15)
    typeNames = Array("LinearStatic", "NonlinearStatic", "Modal", "ResponseSpectrum", _
                      "LinearHistory", "NonlinearHistory", "LinearDynamic", "NonlinearDynamic", _
                      "MovingLoad", "Buckling", "SteadyState", "PowerSpectralDensity", _
                      "LinearStaticMultistep", "Hyperstatic", "ExternalResults")
    
    For i = 0 To UBound(caseTypes)
        On Error Resume Next
        ret = SapModel.LoadCases.GetNameList(NumberNames, MyName, CLng(caseTypes(i)))
        
        If err.number = 0 And ret = 0 And NumberNames > 0 Then
            ' Check if our case is in this list
            Dim j As Long
            For j = 0 To NumberNames - 1
                If MyName(j) = caseName Then
                    GetLoadCaseType = CStr(typeNames(i))
                    err.Clear
                    On Error GoTo 0
                    Exit Function
                End If
            Next j
        End If
        err.Clear
    Next i
    
    On Error GoTo 0
    GetLoadCaseType = "Unknown"
End Function

Private Sub WriteLoadCombinations(ws As Worksheet, startRow As Long, ComboNames() As String, NumberCombos As Long)
    If SapModel Is Nothing Then Exit Sub
    
    Dim i As Long, j As Long
    Dim rowPtr As Long
    rowPtr = startRow
    
    For i = 0 To NumberCombos - 1
        Dim comboName As String
        comboName = ComboNames(i)
        
        ' Get combination type
        Dim comboType As Long
        Dim ret As Long
        comboType = 0
        
        On Error Resume Next
        ret = SapModel.RespCombo.GetTypeOAPI(comboName, comboType)
        err.Clear
        On Error GoTo 0
        
        Dim comboTypeStr As String
        comboTypeStr = GetComboTypeString(comboType)
        
        ' Get notes
        Dim notes As String, guid As String
        notes = ""
        guid = ""
        
        On Error Resume Next
        ret = SapModel.RespCombo.GetNotes(comboName, notes, guid)
        If err.number <> 0 Then
            notes = ""
            err.Clear
        End If
        On Error GoTo 0
        
        ' Get all cases/combos in this combination
        Dim NumberItems As Long
        Dim CType() As eCNameType
        Dim CName() As String
        Dim sF() As Double
        
        ret = SapModel.RespCombo.GetCaseList(comboName, NumberItems, CType, CName, sF)
        
        If ret = 0 And NumberItems > 0 Then
            ' Build formula string for the whole combination (written once in column G of the first row)
            Dim formulaStr As String
            formulaStr = ""
            Dim k As Long
            For k = 0 To NumberItems - 1
                Dim coef As Double
                coef = sF(k)
                Dim itemName As String
                itemName = CName(k)
                Dim coefStr As String
                coefStr = Format(coef, "0.00")
                
                If k = 0 Then
                    If coef < 0 Then
                        formulaStr = "-" & Format(Abs(coef), "0.00") & itemName
                    Else
                        formulaStr = coefStr & itemName
                    End If
                Else
                    If coef < 0 Then
                        formulaStr = formulaStr & " - " & Format(Abs(coef), "0.00") & itemName
                    Else
                        formulaStr = formulaStr & " + " & coefStr & itemName
                    End If
                End If
            Next k
            
            ' Write first item with combo name
            ws.Cells(rowPtr, "A").Value = comboName
            ws.Cells(rowPtr, "B").Value = comboTypeStr
            ws.Cells(rowPtr, "C").Value = CName(0)
            ws.Cells(rowPtr, "D").Value = IIf(CType(0) = 0, "LoadCase", "LoadCombo")
            ws.Cells(rowPtr, "E").Value = sF(0)
            ws.Cells(rowPtr, "F").Value = notes
            ws.Cells(rowPtr, "G").Value = formulaStr    ' <-- công th?c cho toàn combo
            
            ' Format scale factor
            ws.Cells(rowPtr, "E").NumberFormat = "0.00"
            
            rowPtr = rowPtr + 1
            
            ' Write remaining items (without repeating combo name or formula)
            For j = 1 To NumberItems - 1
                ws.Cells(rowPtr, "A").Value = ""
                ws.Cells(rowPtr, "B").Value = ""
                ws.Cells(rowPtr, "C").Value = CName(j)
                ws.Cells(rowPtr, "D").Value = IIf(CType(j) = 0, "LoadCase", "LoadCombo")
                ws.Cells(rowPtr, "E").Value = sF(j)
                ws.Cells(rowPtr, "E").NumberFormat = "0.00"
                ' leave column G blank for subsequent rows of same combo
                
                rowPtr = rowPtr + 1
            Next j
        Else
            ' Combo has no items (empty combo)
            ws.Cells(rowPtr, "A").Value = comboName
            ws.Cells(rowPtr, "B").Value = comboTypeStr
            ws.Cells(rowPtr, "C").Value = "(empty)"
            ws.Cells(rowPtr, "F").Value = notes
            ws.Cells(rowPtr, "G").Value = ""  ' no formula
            rowPtr = rowPtr + 1
        End If
    Next i
End Sub

Private Function GetDesignTypeString(designType As Long) As String
    ' Convert design type enum to string
    Select Case designType
        Case 0: GetDesignTypeString = "Dead"
        Case 1: GetDesignTypeString = "SuperDead"
        Case 2: GetDesignTypeString = "Live"
        Case 3: GetDesignTypeString = "ReduceLive"
        Case 4: GetDesignTypeString = "Quake"
        Case 5: GetDesignTypeString = "Wind"
        Case 6: GetDesignTypeString = "Snow"
        Case 7: GetDesignTypeString = "Other"
        Case 8: GetDesignTypeString = "LiveStorage"
        Case 9: GetDesignTypeString = "LiveRoof"
        Case 10: GetDesignTypeString = "Notional"
        Case 11: GetDesignTypeString = "PatternLive"
        Case 12: GetDesignTypeString = "Wave"
        Case 13: GetDesignTypeString = "Braking"
        Case 14: GetDesignTypeString = "Centrifugal"
        Case 15: GetDesignTypeString = "Friction"
        Case 16: GetDesignTypeString = "Ice"
        Case 17: GetDesignTypeString = "Windonlive"
        Case 18: GetDesignTypeString = "HorizEarthPressure"
        Case 19: GetDesignTypeString = "VertEarthPressure"
        Case 20: GetDesignTypeString = "EarthSurcharge"
        Case 21: GetDesignTypeString = "DownDrag"
        Case 22: GetDesignTypeString = "VehicleCollision"
        Case 23: GetDesignTypeString = "VesselCollision"
        Case 24: GetDesignTypeString = "Temperature"
        Case 25: GetDesignTypeString = "Settlement"
        Case 26: GetDesignTypeString = "Shrinkage"
        Case 27: GetDesignTypeString = "Creep"
        Case 28: GetDesignTypeString = "WaterLoadH"
        Case 29: GetDesignTypeString = "WaterLoadV"
        Case 30: GetDesignTypeString = "RoofLive"
        Case 31: GetDesignTypeString = "Governor"
        Case 32: GetDesignTypeString = "Construction"
        Case Else: GetDesignTypeString = "Other"
    End Select
End Function

Private Function GetComboTypeString(comboType As Long) As String
    ' Convert combo type enum to string
    Select Case comboType
        Case 0: GetComboTypeString = "Linear Additive"
        Case 1: GetComboTypeString = "Envelope"
        Case 2: GetComboTypeString = "Absolute Additive"
        Case 3: GetComboTypeString = "SRSS"
        Case 4: GetComboTypeString = "Range Additive"
        Case Else: GetComboTypeString = "Unknown"
    End Select
End Function

