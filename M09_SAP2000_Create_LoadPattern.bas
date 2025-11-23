Attribute VB_Name = "M09_SAP2000_Create_LoadPattern"
Option Explicit
'===============================================================
' Module: modSAP2000_AutoCreatePatterns
' Purpose: Auto create load patterns from pasted text
' Notes: Updated overwrite / delete-loadcase logic per user request.
'        Added "ensure load case" step: for every requested pattern,
'        if a load case with the same name is missing then create it
'        (StaticLinear.SetCase) and add the pattern to the case with SF = 1.
'===============================================================

Public isFormOpen As Boolean

Public Sub AutoCreateLoadPatterns()
    On Error GoTo ERR_HANDLER
    
    If isFormOpen Then
        MsgBox "Load Pattern form is already open. Please complete or close it first.", vbExclamation
        Exit Sub
    End If
    
    If SapModel Is Nothing Then
        If Not ConnectSAP2000() Then
            MsgBox "Cannot connect to SAP2000. Please ensure SAP2000 is running.", vbCritical
            Exit Sub
        End If
    End If
    
    isFormOpen = True
    Dim frm As frmPasteLoadPatterns
    Set frm = New frmPasteLoadPatterns
    frm.Show vbModeless
    
    Exit Sub
    
ERR_HANDLER:
    isFormOpen = False
    MsgBox "Error opening form: " & err.description, vbCritical
End Sub

'===============================================================
' MAIN PARSING FUNCTION - Parse original input text
'===============================================================
Public Function ParseLoadPatternText(inputText As String) As Object
    Dim patterns As Object
    Set patterns = CreateObject("Scripting.Dictionary")
    
    inputText = Replace(inputText, vbCrLf, vbLf)
    inputText = Replace(inputText, vbCr, vbLf)
    
    Dim lines As Variant
    lines = Split(inputText, vbLf)
    
    Dim i As Long
    Dim deadPatternCount As Long
    deadPatternCount = 0
    
    For i = LBound(lines) To UBound(lines)
        Dim lineText As String
        lineText = Trim(CStr(lines(i)))
        
        If Len(lineText) = 0 Then GoTo NextLine
        If Left(lineText, 1) = "#" Or Left(lineText, 2) = "//" Then GoTo NextLine
        If InStr(1, lineText, "LOAD PATTERN", vbTextCompare) > 0 Or _
           InStr(1, lineText, "LOAD CASE", vbTextCompare) > 0 Or _
           InStr(1, lineText, "PREVIEW", vbTextCompare) > 0 Or _
           InStr(1, lineText, "NOTE:", vbTextCompare) > 0 Or _
           InStr(1, lineText, "Click", vbTextCompare) > 0 Or _
           InStr(1, lineText, "Format:", vbTextCompare) > 0 Or _
           InStr(1, lineText, "Paste your", vbTextCompare) > 0 Then
            GoTo NextLine
        End If
        If InStr(lineText, "---") > 0 Or InStr(lineText, "===") > 0 Then GoTo NextLine
        If InStr(1, lineText, "Pattern", vbTextCompare) > 0 And _
           InStr(1, lineText, "Type", vbTextCompare) > 0 Then GoTo NextLine
        If InStr(lineText, "=") > 0 And InStr(lineText, ":") = 0 Then GoTo NextLine
        
        Dim colonPos As Long
        colonPos = InStr(lineText, ":")
        If colonPos > 0 Then
            Dim patternName As String
            Dim description As String
            patternName = Trim(Left(lineText, colonPos - 1))
            description = Trim(mid(lineText, colonPos + 1))
            
            patternName = Trim(Replace(patternName, vbTab, ""))
            description = Trim(Replace(description, vbTab, " "))
            
            Do While InStr(description, "  ") > 0
                description = Replace(description, "  ", " ")
            Loop
            
            If Len(patternName) > 0 And Len(patternName) <= 50 And Len(description) > 0 Then
                If Not patterns.exists(patternName) Then
                    Dim loadType As Long
                    Dim loadTypeStr As String
                    loadType = DetermineLoadType(patternName, description)
                    
                    If loadType = 1 Then
                        deadPatternCount = deadPatternCount + 1
                        If deadPatternCount > 1 Then
                            loadType = 2
                        End If
                    End If
                    
                    loadTypeStr = GetLoadPatternTypeString(loadType)
                    
                    Dim selfWtMult As Double
                    If loadType = 2 Then
                        selfWtMult = 0#
                    ElseIf loadType = 1 Then
                        If UCase(patternName) = "BL" Or _
                           InStr(1, description, "Self Weight", vbTextCompare) > 0 Or _
                           InStr(1, description, "Structure Weight", vbTextCompare) > 0 Or _
                           InStr(1, description, "Dead Load", vbTextCompare) > 0 Then
                            selfWtMult = 1#
                        Else
                            selfWtMult = 0#
                        End If
                    Else
                        If InStr(1, description, "Self Weight", vbTextCompare) > 0 Or _
                           InStr(1, description, "Structure Weight", vbTextCompare) > 0 Then
                            selfWtMult = 1#
                        Else
                            selfWtMult = 0#
                        End If
                    End If
                    
                    Dim patternInfo(0 To 4) As Variant
                    patternInfo(0) = patternName
                    patternInfo(1) = description
                    patternInfo(2) = loadTypeStr
                    patternInfo(3) = selfWtMult
                    patternInfo(4) = loadType
                    
                    patterns.Add patternName, patternInfo
                End If
            End If
        End If
NextLine:
    Next i
    
    Set ParseLoadPatternText = patterns
End Function

'===============================================================
' PARSE TABLE FORMAT - For edited preview data
'===============================================================
Public Function ParseTableFormat(inputText As String) As Object
    Dim patterns As Object
    Set patterns = CreateObject("Scripting.Dictionary")
    
    inputText = Replace(inputText, vbCrLf, vbLf)
    inputText = Replace(inputText, vbCr, vbLf)
    
    Dim lines As Variant
    lines = Split(inputText, vbLf)
    
    Dim i As Long
    Dim deadPatternCount As Long
    deadPatternCount = 0
    Dim inTableSection As Boolean
    inTableSection = False
    
    For i = LBound(lines) To UBound(lines)
        Dim lineText As String
        lineText = CStr(lines(i))
        
        If Len(Trim(lineText)) = 0 Then GoTo NextLine
        If InStr(lineText, "---") > 0 Or InStr(lineText, "===") > 0 Then GoTo NextLine
        
        If InStr(1, lineText, "Pattern", vbTextCompare) > 0 And _
           InStr(1, lineText, "Type", vbTextCompare) > 0 And _
           InStr(1, lineText, "SW Mult", vbTextCompare) > 0 Then
            inTableSection = True
            GoTo NextLine
        End If
        
        If InStr(1, lineText, "You can edit", vbTextCompare) > 0 Or _
           InStr(1, lineText, "Keep the table", vbTextCompare) > 0 Or _
           InStr(1, lineText, "Click", vbTextCompare) > 0 Or _
           InStr(1, lineText, "PREVIEW", vbTextCompare) > 0 Or _
           InStr(1, lineText, "Options", vbTextCompare) > 0 Or _
           InStr(1, lineText, "Overwrite", vbTextCompare) > 0 Or _
           InStr(1, lineText, "Delete old", vbTextCompare) > 0 Then
            inTableSection = False
            GoTo NextLine
        End If
        
        If Not inTableSection Then GoTo NextLine
        
        If Len(lineText) >= 41 Then
            Dim patternName As String
            Dim typeStr As String
            Dim swMultStr As String
            
            patternName = Trim(Left(lineText, 15))
            typeStr = Trim(mid(lineText, 17, 25))
            swMultStr = Trim(mid(lineText, 43))
            
            If Len(patternName) > 0 And Len(typeStr) > 0 Then
                If InStr(1, patternName, " ", vbTextCompare) > 3 Or _
                   Len(patternName) > 15 Then
                    GoTo NextLine
                End If
                
                Dim loadType As Long
                loadType = GetLoadTypeFromString(typeStr)
                
                If loadType = 1 Then
                    deadPatternCount = deadPatternCount + 1
                    If deadPatternCount > 1 Then
                        loadType = 2
                        typeStr = GetLoadPatternTypeString(loadType)
                    Else
                        typeStr = GetLoadPatternTypeString(loadType)
                    End If
                Else
                    typeStr = GetLoadPatternTypeString(loadType)
                End If
                
                Dim selfWtMult As Double
                On Error Resume Next
                selfWtMult = CDbl(swMultStr)
                If err.number <> 0 Then selfWtMult = 0#
                err.Clear
                On Error GoTo 0
                
                If loadType = 2 Then
                    selfWtMult = 0#
                End If
                
                If Not patterns.exists(patternName) Then
                    Dim patternInfo(0 To 4) As Variant
                    patternInfo(0) = patternName
                    patternInfo(1) = ""
                    patternInfo(2) = typeStr
                    patternInfo(3) = selfWtMult
                    patternInfo(4) = loadType
                    patterns.Add patternName, patternInfo
                End If
            End If
        End If
NextLine:
    Next i
    
    Set ParseTableFormat = patterns
End Function

'===============================================================
' CREATE PATTERNS IN SAP2000
' Overwrite logic (updated):
' Cases:
' 1) overwritePatterns = True, deleteLoadcases = False:
'    - Delete duplicate loadcases whose names match any new pattern first EXCEPT MODAL (type = 3)
'    - Create one new pattern first (if needed)
'    - Delete ALL existing patterns except that ensured one (no protection for DEAD)
'    - Create all new patterns
'
' 2) overwritePatterns = True, deleteLoadcases = True:
'    - Delete ALL loadcases first EXCEPT MODAL (type = 3)
'    - Create one new pattern first (if needed)
'    - Delete ALL existing patterns except that ensured one (no protection for DEAD)
'    - Create all new patterns
'
' 3) overwritePatterns = False:
'    - Create new patterns only if they don't already exist (skip duplicates)
' Additional: After pattern creation, ensure for each requested pattern there exists a load case
'             with the same name. If missing, create load case with StaticLinear.SetCase and
'             add the pattern to the case with scale factor = 1 using SetLoads.
'===============================================================
Public Function CreatePatternsInSAP(patterns As Object, Optional overwritePatterns As Boolean = False, Optional deleteLoadcases As Boolean = False) As String
    Dim createdCount As Long
    Dim skippedCount As Long
    Dim errorCount As Long
    Dim warningCount As Long
    Dim deletedPatternCount As Long
    Dim deletedCaseCount As Long
    Dim errorLog As String
    Dim warningLog As String
    
    ' Added counters for load-case ensuring step
    Dim createdCaseCount As Long
    Dim skippedCaseCount As Long
    Dim caseErrorCount As Long
    
    createdCount = 0
    skippedCount = 0
    errorCount = 0
    warningCount = 0
    deletedPatternCount = 0
    deletedCaseCount = 0
    createdCaseCount = 0
    skippedCaseCount = 0
    caseErrorCount = 0
    errorLog = ""
    warningLog = ""
    
    ' Build list of new pattern names (uppercase for comparison)
    Dim newNames As Object
    Set newNames = CreateObject("Scripting.Dictionary")
    Dim pName As Variant
    For Each pName In patterns.keys
        newNames.Add UCase(CStr(pName)), True
    Next pName
    
    ' Get existing load pattern list
    Dim NumberNames As Long
    Dim MyName() As String
    Dim ret As Long
    On Error Resume Next
    ret = SapModel.loadPatterns.GetNameList(NumberNames, MyName)
    If err.number <> 0 Or ret <> 0 Then
        errorLog = errorLog & "ERROR: Unable to retrieve existing load patterns list. err=" & err.number & " ret=" & ret & vbCrLf
        errorCount = errorCount + 1
        err.Clear
        NumberNames = 0
        ReDim MyName(-1)
    End If
    On Error GoTo 0
    
    Dim existingNames As Object
    Set existingNames = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = 0 To NumberNames - 1
        existingNames.Add UCase(MyName(i)), True
    Next i
    
    ' ---------- OPTION 1: overwritePatterns = True, deleteLoadcases = False ----------
    ' Delete duplicate loadcases with names matching new patterns first (except modal)
    If overwritePatterns And Not deleteLoadcases Then
        On Error Resume Next
        Dim totalCases1 As Long
        Dim allCases1() As String
        ret = SapModel.LoadCases.GetNameList_1(totalCases1, allCases1)
        If err.number <> 0 Or ret <> 0 Then
            errorLog = errorLog & "ERROR: Unable to retrieve load case list for duplicate deletion. err=" & err.number & " ret=" & ret & vbCrLf
            errorCount = errorCount + 1
            err.Clear
        Else
            ' Get Modal cases (type = 3) to protect
            Dim modalCount1 As Long
            Dim modalCases1() As String
            ret = SapModel.LoadCases.GetNameList_1(modalCount1, modalCases1, 3)
            If err.number <> 0 Or ret <> 0 Then
                modalCount1 = 0
                Erase modalCases1
                err.Clear
            End If
            
            Dim modalDict1 As Object
            Set modalDict1 = CreateObject("Scripting.Dictionary")
            Dim idx1 As Long
            For idx1 = 0 To modalCount1 - 1
                modalDict1.Add UCase(modalCases1(idx1)), True
            Next idx1
            
            ' Delete loadcases that match any new pattern name (except modal)
            For idx1 = 0 To totalCases1 - 1
                Dim caseName1 As String
                caseName1 = CStr(allCases1(idx1))
                If newNames.exists(UCase(caseName1)) Then
                    If Not modalDict1.exists(UCase(caseName1)) Then
                        ret = SapModel.LoadCases.Delete(caseName1)
                        If ret = 0 Then
                            deletedCaseCount = deletedCaseCount + 1
                        Else
                            errorLog = errorLog & "ERROR: Unable to delete duplicate load case '" & caseName1 & "' (code " & ret & ")" & vbCrLf
                            errorCount = errorCount + 1
                        End If
                    End If
                End If
            Next idx1
        End If
        On Error GoTo 0
    End If
    
    ' ---------- OPTION 2: deleteLoadcases = True (delete ALL except modal) ----------
    If deleteLoadcases Then
        On Error Resume Next
        Dim totalCases As Long
        Dim allCases() As String
        ret = SapModel.LoadCases.GetNameList_1(totalCases, allCases)
        If err.number <> 0 Or ret <> 0 Then
            errorLog = errorLog & "ERROR: Unable to retrieve load case list for deletion. err=" & err.number & " ret=" & ret & vbCrLf
            errorCount = errorCount + 1
            err.Clear
        Else
            ' Get Modal cases (type = 3) to protect
            Dim modalCount As Long
            Dim modalCases() As String
            ret = SapModel.LoadCases.GetNameList_1(modalCount, modalCases, 3)
            If err.number <> 0 Or ret <> 0 Then
                modalCount = 0
                Erase modalCases
                err.Clear
            End If
            
            Dim modalDict As Object
            Set modalDict = CreateObject("Scripting.Dictionary")
            Dim idx As Long
            For idx = 0 To modalCount - 1
                modalDict.Add UCase(modalCases(idx)), True
            Next idx
            
            ' Delete all loadcases except modal ones
            For idx = 0 To totalCases - 1
                Dim caseName As String
                caseName = CStr(allCases(idx))
                If Not modalDict.exists(UCase(caseName)) Then
                    ret = SapModel.LoadCases.Delete(caseName)
                    If ret = 0 Then
                        deletedCaseCount = deletedCaseCount + 1
                    Else
                        errorLog = errorLog & "ERROR: Unable to delete load case '" & caseName & "' (code " & ret & ")" & vbCrLf
                        errorCount = errorCount + 1
                    End If
                End If
            Next idx
        End If
        On Error GoTo 0
    End If
    
    '=================================================================
    ' OVERWRITE MODE: DELETE ALL EXISTING PATTERNS (except the temporary/ensured one)
    ' NOTE: Failures to delete individual patterns are now treated as WARNINGS
    '       (not fatal Errors), because some patterns may be protected/locked by SAP.
    '=================================================================
    Dim firstPatternCreated As String
    firstPatternCreated = ""
    
    If overwritePatterns Then
        Dim ensuredOne As Boolean
        ensuredOne = False
        
        ' Step 1: Ensure at least one pattern exists by creating one (if needed)
        For Each pName In patterns.keys
            Dim pnameStr As String
            pnameStr = CStr(pName)
            If Not existingNames.exists(UCase(pnameStr)) Then
                Dim infoArr As Variant
                infoArr = patterns(pName)
                ret = SapModel.loadPatterns.Add(pnameStr, CLng(infoArr(4)), CDbl(infoArr(3)), True)
                If ret = 0 Then
                    createdCount = createdCount + 1
                    ensuredOne = True
                    firstPatternCreated = pnameStr
                    existingNames.Add UCase(pnameStr), True
                Else
                    errorLog = errorLog & "ERROR: Creating initial pattern '" & pnameStr & "' failed (code " & ret & ")" & vbCrLf
                    errorCount = errorCount + 1
                End If
                Exit For
            Else
                ' If it already exists, we can use it as the ensured one
                ensuredOne = True
                firstPatternCreated = pnameStr
                Exit For
            End If
        Next pName
        
        ' If all new names already exist (we couldn't create any), create a temporary pattern to ensure at least one pattern exists
        If Not ensuredOne Then
            Dim tmpName As String
            tmpName = "__TMP_PATTERN__"
            ret = SapModel.loadPatterns.Add(tmpName, 8, 0#, True)
            If ret = 0 Then
                ensuredOne = True
                firstPatternCreated = tmpName
                existingNames.Add UCase(tmpName), True
                createdCount = createdCount + 1
            Else
                errorLog = errorLog & "ERROR: Unable to create temporary pattern (code " & ret & ")" & vbCrLf
                errorCount = errorCount + 1
                ' If cannot create temp, we cannot safely proceed with overwrite - fallback to non-overwrite behavior
                overwritePatterns = False
            End If
        End If
        
        ' Step 2: Refresh current pattern list
        If overwritePatterns Then
            On Error Resume Next
            NumberNames = 0
            Erase MyName
            ret = SapModel.loadPatterns.GetNameList(NumberNames, MyName)
            If err.number = 0 And ret = 0 Then
                Set existingNames = CreateObject("Scripting.Dictionary")
                For i = 0 To NumberNames - 1
                    existingNames.Add UCase(MyName(i)), True
                Next i
            Else
                If err.number <> 0 Then
                    errorLog = errorLog & "ERROR: Unable to refresh pattern list. err=" & err.number & vbCrLf
                    errorCount = errorCount + 1
                    err.Clear
                Else
                    errorLog = errorLog & "ERROR: GetNameList failed (ret=" & ret & ")" & vbCrLf
                    errorCount = errorCount + 1
                End If
                overwritePatterns = False
            End If
            On Error GoTo 0
        End If
        
        ' Step 3: Delete ALL existing patterns except firstPatternCreated (no protection for DEAD or other types)
        If overwritePatterns Then
            Dim toDelete As Collection
            Set toDelete = New Collection
            Dim keyName As Variant
            
            For Each keyName In existingNames.keys
                Dim keyStr As String
                keyStr = CStr(keyName)
                
                ' Delete EVERYTHING except the ensured one
                If UCase(keyStr) <> UCase(firstPatternCreated) Then
                    toDelete.Add keyStr
                End If
            Next keyName
            
            ' Perform deletion
            Dim delName As Variant
            For Each delName In toDelete
                ret = SapModel.loadPatterns.Delete(CStr(delName))
                If ret = 0 Then
                    deletedPatternCount = deletedPatternCount + 1
                Else
                    ' Treat inability to delete an individual pattern as a WARNING (not fatal)
                    warningLog = warningLog & "WARNING: Unable to delete pattern '" & delName & "' (code " & ret & ")" & vbCrLf
                    warningCount = warningCount + 1
                    ' don't increment errorCount
                End If
            Next delName
        End If
    End If
    
    '=================================================================
    ' CREATE NEW PATTERNS (either overwrite mode or normal mode)
    '=================================================================
    For Each pName In patterns.keys
        Dim pNames As String
        pNames = CStr(pName)
        
        ' Check existence now (after potential deletions)
        Dim existsNow As Boolean
        existsNow = False
        On Error Resume Next
        NumberNames = 0
        Erase MyName
        ret = SapModel.loadPatterns.GetNameList(NumberNames, MyName)
        If err.number = 0 And ret = 0 Then
            For i = 0 To NumberNames - 1
                If UCase(MyName(i)) = UCase(pNames) Then
                    existsNow = True
                    Exit For
                End If
            Next i
        End If
        err.Clear
        On Error GoTo 0
        
        If Not existsNow Then
            Dim pInfo As Variant
            pInfo = patterns(pName)
            ret = SapModel.loadPatterns.Add(pNames, CLng(pInfo(4)), CDbl(pInfo(3)), True)
            If ret = 0 Then
                createdCount = createdCount + 1
            Else
                errorLog = errorLog & "ERROR: Creating pattern '" & pNames & "' failed (code " & ret & ")" & vbCrLf
                errorCount = errorCount + 1
            End If
        Else
            skippedCount = skippedCount + 1
        End If
    Next pName
    
    ' If we created a temporary pattern named "__TMP_PATTERN__", attempt to delete it now (only if it wasn't part of requested patterns)
    On Error Resume Next
    If overwritePatterns Then
        If UCase(firstPatternCreated) = "__TMP_PATTERN__" Then
            ' ensure user hasn't also requested it
            If Not newNames.exists(UCase("__TMP_PATTERN__")) Then
                ret = SapModel.loadPatterns.Delete("__TMP_PATTERN__")
                ' ignore result - we don't treat failure as fatal
                If ret = 0 Then
                    deletedPatternCount = deletedPatternCount + 1
                End If
            End If
        End If
    End If
    On Error GoTo 0
    
    '=================================================================
    ' ENSURE LOAD CASES FOR REQUESTED PATTERNS
    ' For each requested pattern name: if a load case with same name doesn't exist,
    ' create a static-linear load case (SetCase) and assign the load (SetLoads)
    ' with LoadType = "Load", LoadName = patternName and SF = 1.0.
    ' If load case already exists, skip. Log errors for failures.
    '=================================================================
    On Error Resume Next
    Dim totalCasesNow As Long
    Dim allCasesNow() As String
    ret = SapModel.LoadCases.GetNameList_1(totalCasesNow, allCasesNow)
    If err.number <> 0 Or ret <> 0 Then
        errorLog = errorLog & "ERROR: Unable to retrieve load case list for ensuring cases. err=" & err.number & " ret=" & ret & vbCrLf
        errorCount = errorCount + 1
        err.Clear
    Else
        Dim caseDict As Object
        Set caseDict = CreateObject("Scripting.Dictionary")
        Dim ci As Long
        For ci = 0 To totalCasesNow - 1
            caseDict.Add UCase(CStr(allCasesNow(ci))), True
        Next ci
        
        For Each pName In patterns.keys
            Dim patternCaseName As String
            patternCaseName = CStr(pName)
            If Not caseDict.exists(UCase(patternCaseName)) Then
                ' Create new static-linear load case with the pattern name
                ret = SapModel.LoadCases.StaticLinear.SetCase(patternCaseName)
                If ret = 0 Then
                    ' Now assign the pattern to the case with SF = 1
                    Dim MyLoadType() As String
                    Dim MyLoadName() As String
                    Dim MySF() As Double
                    ReDim MyLoadType(0)
                    ReDim MyLoadName(0)
                    ReDim MySF(0)
                    MyLoadType(0) = "Load"
                    MyLoadName(0) = patternCaseName
                    MySF(0) = 1#
                    
                    ret = SapModel.LoadCases.StaticLinear.SetLoads(patternCaseName, 1, MyLoadType, MyLoadName, MySF)
                    If ret = 0 Then
                        createdCaseCount = createdCaseCount + 1
                        ' add to caseDict to avoid duplicate effort
                        caseDict.Add UCase(patternCaseName), True
                    Else
                        errorLog = errorLog & "ERROR: SetLoads failed for case '" & patternCaseName & "' (code " & ret & ")" & vbCrLf
                        caseErrorCount = caseErrorCount + 1
                    End If
                Else
                    errorLog = errorLog & "ERROR: SetCase failed for '" & patternCaseName & "' (code " & ret & ")" & vbCrLf
                    caseErrorCount = caseErrorCount + 1
                End If
            Else
                ' Already exists - skip
                skippedCaseCount = skippedCaseCount + 1
            End If
        Next pName
    End If
    err.Clear
    On Error GoTo 0
    
    '=================================================================
    ' BUILD SUMMARY MESSAGE
    '=================================================================
    Dim summary As String
    summary = "===================================" & vbCrLf
    summary = summary & "  LOAD PATTERN CREATION SUMMARY" & vbCrLf
    summary = summary & "===================================" & vbCrLf & vbCrLf
    summary = summary & "Patterns in request: " & patterns.count & vbCrLf
    summary = summary & "Created: " & createdCount & vbCrLf
    summary = summary & "Skipped (already exist): " & skippedCount & vbCrLf
    
    If overwritePatterns Then
        summary = summary & "Deleted (old patterns): " & deletedPatternCount & vbCrLf
    End If
    
    If (Not deleteLoadcases And overwritePatterns) Or deleteLoadcases Then
        summary = summary & "Deleted (load cases): " & deletedCaseCount & vbCrLf
    End If
    
    summary = summary & vbCrLf & "Load cases ensured for patterns:" & vbCrLf
    summary = summary & "  Created cases: " & createdCaseCount & vbCrLf
    summary = summary & "  Skipped (case already existed): " & skippedCaseCount & vbCrLf
    If caseErrorCount > 0 Then
        summary = summary & "  Case errors: " & caseErrorCount & vbCrLf
    End If
    
    summary = summary & vbCrLf & "Errors: " & errorCount & vbCrLf
    summary = summary & "Warnings: " & warningCount & vbCrLf & vbCrLf
    
    If errorCount > 0 Or caseErrorCount > 0 Then
        summary = summary & "Error details:" & vbCrLf
        summary = summary & "-----------------------------------" & vbCrLf
        summary = summary & errorLog
    End If
    
    If warningCount > 0 Then
        summary = summary & vbCrLf & "Warning details (non-fatal):" & vbCrLf
        summary = summary & "-----------------------------------" & vbCrLf
        summary = summary & warningLog
    End If
    
    CreatePatternsInSAP = summary
End Function

'===============================================================
' IMPROVED DetermineLoadType with SCORING SYSTEM
' Replaces the priority-based approach with multi-factor scoring
'===============================================================
Private Function DetermineLoadType(Name As String, description As String) As Long
    Dim nameLower As String
    Dim descLower As String
    
    nameLower = LCase(Trim(Name))
    descLower = LCase(Trim(description))
    
    ' Initialize scores for each load type (39 types total)
    Dim scores(1 To 39) As Long
    Dim i As Long
    For i = 1 To 39
        scores(i) = 0
    Next i
    
    '=================================================================
    ' SCORING PHASE 1: EXACT NAME MATCHES (Highest confidence)
    '=================================================================
    Select Case nameLower
        Case "bl"
            scores(1) = scores(1) + 100  ' Dead
        Case "dl"
            scores(1) = scores(1) + 100  ' Dead
        Case "sdl"
            scores(2) = scores(2) + 100  ' SuperDead
        Case "ll"
            scores(3) = scores(3) + 100  ' Live
        Case "lr"
            scores(11) = scores(11) + 100  ' RoofLive
        Case "th", "temp"
            scores(10) = scores(10) + 100  ' Temperature
        Case "gwp"
            scores(30) = scores(30) + 100  ' WaterLoadPressure
        Case "sp"
            scores(20) = scores(20) + 100  ' HorizontalEarthPressure
        Case "ep"
            scores(20) = scores(20) + 100  ' HorizontalEarthPressure
        Case "sur"
            scores(22) = scores(22) + 100  ' EarthSurcharge
    End Select
    
    '=================================================================
    ' SCORING PHASE 2: DESCRIPTION KEYWORDS (High confidence)
    '=================================================================
    ' NOTIONAL - Must check FIRST with high score to override other weak signals
    If MatchKeywords(descLower, Array("notional", "stability")) Then
        scores(12) = scores(12) + 80  ' Notional
    End If
    
    ' DEAD and SUPERDEAD
    If MatchKeywords(descLower, Array("self weight", "structure weight", "structural weight", "selfweight")) Then
        scores(1) = scores(1) + 70  ' Dead
    End If
    
    If MatchKeywords(descLower, Array("dead load")) And Not MatchKeywords(descLower, Array("super", "additional")) Then
        scores(1) = scores(1) + 60  ' Dead
    End If
    
    If MatchKeywords(descLower, Array("super dead", "superdead", "superimposed", "additional dead")) Then
        scores(2) = scores(2) + 80  ' SuperDead
    End If
    
    ' LIVE loads
    If MatchKeywords(descLower, Array("live")) And MatchKeywords(descLower, Array("roof")) Then
        scores(11) = scores(11) + 75  ' RoofLive
    End If
    
    If MatchKeywords(descLower, Array("live load", "occupancy")) Then
        scores(3) = scores(3) + 65  ' Live
    End If
    
    ' SEISMIC
    If MatchKeywords(descLower, Array("seismic", "earthquake")) Then
        scores(5) = scores(5) + 80  ' Quake
    End If
    
    If MatchKeywords(descLower, Array("dynamic")) Then
        scores(5) = scores(5) + 40  ' Quake (weaker signal)
    End If
    
    ' WIND
    If MatchKeywords(descLower, Array("wind")) Then
        scores(6) = scores(6) + 75  ' Wind
    End If
    
    ' TEMPERATURE
    If MatchKeywords(descLower, Array("thermal", "temperature")) And Not MatchKeywords(descLower, Array("gradient")) Then
        scores(10) = scores(10) + 70  ' Temperature
    End If
    
    If MatchKeywords(descLower, Array("temperature gradient", "gradient")) Then
        scores(26) = scores(26) + 75  ' TemperatureGradient
    End If
    
    ' EQUIPMENT and PIPING
    If MatchKeywords(descLower, Array("equipment", "machinery")) Then
        scores(8) = scores(8) + 70  ' Other
    End If
    
    If MatchKeywords(descLower, Array("piping", "pipe")) Then
        scores(8) = scores(8) + 70  ' Other
    End If
    
    If MatchKeywords(descLower, Array("cable")) Then
        scores(8) = scores(8) + 65  ' Other
    End If
    
    ' WATER and EARTH PRESSURE
    If MatchKeywords(descLower, Array("water pressure", "groundwater", "ground water", "waterload")) Then
        scores(30) = scores(30) + 75  ' WaterLoadPressure
    End If
    
    If MatchKeywords(descLower, Array("soil pressure", "earth pressure", "earthpressure")) And Not MatchKeywords(descLower, Array("vertical")) Then
        scores(20) = scores(20) + 70  ' HorizontalEarthPressure
    End If
    
    If MatchKeywords(descLower, Array("vertical earth", "vertical soil", "vertical")) And MatchKeywords(descLower, Array("pressure", "soil", "earth")) Then
        scores(21) = scores(21) + 75  ' VerticalEarthPressure
    End If
    
    If MatchKeywords(descLower, Array("surcharge", "surchage")) And Not MatchKeywords(descLower, Array("live")) Then
        scores(22) = scores(22) + 70  ' EarthSurcharge
    End If
    
    If MatchKeywords(descLower, Array("live")) And MatchKeywords(descLower, Array("surcharge")) Then
        scores(31) = scores(31) + 75  ' LiveLoadSurcharge
    End If
    
    ' SNOW
    If MatchKeywords(descLower, Array("snow")) Then
        scores(7) = scores(7) + 75  ' Snow
    End If
    
    ' OTHER SPECIALIZED LOADS
    If MatchKeywords(descLower, Array("buoyancy", "uplift")) Then
        scores(36) = scores(36) + 75  ' Bouyancy
    End If
    
    If MatchKeywords(descLower, Array("construction", "temporary")) Then
        scores(39) = scores(39) + 70  ' Construction
    End If
    
    If MatchKeywords(descLower, Array("settlement")) Then
        scores(27) = scores(27) + 75  ' Settlement
    End If
    
    If MatchKeywords(descLower, Array("prestress", "tendon")) Then
        scores(34) = scores(34) + 75  ' Prestress
    End If
    
    If MatchKeywords(descLower, Array("impact", "blast")) Then
        scores(38) = scores(38) + 75  ' Impact
    End If
    
    If MatchKeywords(descLower, Array("wave", "tsunami", "ocean")) Then
        scores(14) = scores(14) + 75  ' Wave
    End If
    
    If MatchKeywords(descLower, Array("ice")) Then
        scores(18) = scores(18) + 75  ' Ice
    End If
    
    If MatchKeywords(descLower, Array("braking", "brake")) Then
        scores(15) = scores(15) + 75  ' Braking
    End If
    
    If MatchKeywords(descLower, Array("centrifugal", "curve")) Then
        scores(16) = scores(16) + 75  ' Centrifugal
    End If
    
    If MatchKeywords(descLower, Array("friction")) Then
        scores(17) = scores(17) + 70  ' Friction
    End If
    
    If MatchKeywords(descLower, Array("vehicle collision", "vehiclecollision", "vehicle")) And MatchKeywords(descLower, Array("collision")) Then
        scores(24) = scores(24) + 75  ' VehicleCollision
    End If
    
    If MatchKeywords(descLower, Array("vessel collision", "vesselcollision", "ship collision")) Then
        scores(25) = scores(25) + 75  ' VesselCollision
    End If
    
    If MatchKeywords(descLower, Array("shrinkage")) Then
        scores(28) = scores(28) + 75  ' Shrinkage
    End If
    
    If MatchKeywords(descLower, Array("creep")) Then
        scores(29) = scores(29) + 75  ' Creep
    End If
    
    If MatchKeywords(descLower, Array("downdrag", "negative skin friction")) Then
        scores(23) = scores(23) + 75  ' DownDrag
    End If
    
    If MatchKeywords(descLower, Array("stream", "river", "current")) Then
        scores(37) = scores(37) + 70  ' StreamFlow
    End If
    
    If MatchKeywords(descLower, Array("pedestrian")) Then
        scores(33) = scores(33) + 75  ' PedestrianLL
    End If
    
    '=================================================================
    ' SCORING PHASE 3: NAME PREFIX PATTERNS (Medium confidence)
    '=================================================================
    ' NOTIONAL - NL prefix
    If Len(nameLower) >= 2 Then
        If Left(nameLower, 2) = "nl" Then
            scores(12) = scores(12) + 50  ' Notional
        End If
    End If
    
    ' SEISMIC - SX, SY, SZ, EQ prefix (but NOT if already scored high for Notional)
    If Len(nameLower) >= 2 Then
        If (Left(nameLower, 2) = "sx") Or (Left(nameLower, 2) = "sy") Or _
           (Left(nameLower, 2) = "sz") Or (Left(nameLower, 2) = "eq") Then
            ' Only add if not already strongly Notional
            If scores(12) < 50 Then
                scores(5) = scores(5) + 45  ' Quake
            End If
        End If
    End If
    
    ' WIND - W prefix (but exclude WP, GWP)
    If Len(nameLower) >= 1 Then
        If Left(nameLower, 1) = "w" And nameLower <> "wp" And nameLower <> "gwp" Then
            ' Only add if not already strongly Notional
            If scores(12) < 50 Then
                scores(6) = scores(6) + 40  ' Wind
            End If
        End If
    End If
    
    ' EQUIPMENT - E prefix (2 chars like EX, EY)
    ' But be careful: NE*** patterns should not trigger this if Notional is strong
    If Len(nameLower) = 2 Then
        If Left(nameLower, 1) = "e" Then
            If scores(12) < 50 Then  ' Not strongly Notional
                scores(8) = scores(8) + 35  ' Other (Equipment)
            End If
        End If
    End If
    
    ' PIPING - P prefix (2 chars like PX, PY)
    ' But be careful: NP*** patterns should not trigger this if Notional is strong
    If Len(nameLower) = 2 Then
        If Left(nameLower, 1) = "p" Then
            If scores(12) < 50 Then  ' Not strongly Notional
                scores(8) = scores(8) + 35  ' Other (Piping)
            End If
        End If
    End If
    
    ' LIVE - L prefix (but be very careful - many false positives)
    ' Only apply if NO strong Notional signal
    If Len(nameLower) >= 1 And Len(nameLower) <= 3 Then
        If Left(nameLower, 1) = "l" And Left(nameLower, 2) <> "lr" Then
            If scores(12) < 30 Then  ' Not Notional
                scores(3) = scores(3) + 25  ' Live (weak signal)
            End If
        End If
    End If
    
    ' ROOF LIVE - LR prefix
    If Len(nameLower) >= 2 Then
        If Left(nameLower, 2) = "lr" Then
            If scores(12) < 30 Then  ' Not Notional
                scores(11) = scores(11) + 40  ' RoofLive
            End If
        End If
    End If
    
    '=================================================================
    ' SCORING PHASE 4: SECONDARY NAME PATTERNS (Lower confidence)
    '=================================================================
    ' Check for common suffixes that indicate direction (x, y, z)
    If Len(nameLower) >= 2 Then
        Dim lastChar As String
        lastChar = Right(nameLower, 1)
        
        If lastChar = "x" Or lastChar = "y" Or lastChar = "z" Then
            ' This is common for directional loads - slight boost for certain types
            ' But don't let this override strong signals
            If scores(5) > 20 Then scores(5) = scores(5) + 5  ' Seismic often has direction
            If scores(6) > 20 Then scores(6) = scores(6) + 5  ' Wind often has direction
            If scores(12) > 20 Then scores(12) = scores(12) + 5  ' Notional often has direction
        End If
    End If
    
    '=================================================================
    ' FINAL DECISION: Return type with highest score
    '=================================================================
    Dim maxScore As Long
    Dim maxType As Long
    maxScore = 0
    maxType = 8  ' Default to Other
    
    For i = 1 To 39
        If scores(i) > maxScore Then
            maxScore = scores(i)
            maxType = i
        End If
    Next i
    
    ' Safety threshold: if max score is too low, default to Other
    If maxScore < 15 Then
        maxType = 8  ' Other
    End If
    
    DetermineLoadType = maxType
End Function

'===============================================================
' Helper function (unchanged)
'===============================================================
Private Function MatchKeywords(searchText As String, keywords As Variant) As Boolean
    Dim i As Long
    For i = LBound(keywords) To UBound(keywords)
        If InStr(1, searchText, CStr(keywords(i)), vbTextCompare) > 0 Then
            MatchKeywords = True
            Exit Function
        End If
    Next i
    MatchKeywords = False
End Function


Public Function GetLoadPatternTypeString(loadType As Long) As String
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
        Case Else: GetLoadPatternTypeString = "Other"
    End Select
End Function

Private Function GetLoadTypeFromString(typeStr As String) As Long
    Select Case LCase(Trim(typeStr))
        Case "dead": GetLoadTypeFromString = 1
        Case "superdead": GetLoadTypeFromString = 2
        Case "live": GetLoadTypeFromString = 3
        Case "reducelive": GetLoadTypeFromString = 4
        Case "quake": GetLoadTypeFromString = 5
        Case "wind": GetLoadTypeFromString = 6
        Case "snow": GetLoadTypeFromString = 7
        Case "other": GetLoadTypeFromString = 8
        Case "move": GetLoadTypeFromString = 9
        Case "temperature": GetLoadTypeFromString = 10
        Case "rooflive": GetLoadTypeFromString = 11
        Case "notional": GetLoadTypeFromString = 12
        Case "patternlive": GetLoadTypeFromString = 13
        Case "wave": GetLoadTypeFromString = 14
        Case "braking": GetLoadTypeFromString = 15
        Case "centrifugal": GetLoadTypeFromString = 16
        Case "friction": GetLoadTypeFromString = 17
        Case "ice": GetLoadTypeFromString = 18
        Case "windonliveload": GetLoadTypeFromString = 19
        Case "horizontalearthpressure": GetLoadTypeFromString = 20
        Case "verticalearthpressure": GetLoadTypeFromString = 21
        Case "earthsurcharge": GetLoadTypeFromString = 22
        Case "downdrag": GetLoadTypeFromString = 23
        Case "vehiclecollision": GetLoadTypeFromString = 24
        Case "vesselcollision": GetLoadTypeFromString = 25
        Case "temperaturegradient": GetLoadTypeFromString = 26
        Case "settlement": GetLoadTypeFromString = 27
        Case "shrinkage": GetLoadTypeFromString = 28
        Case "creep": GetLoadTypeFromString = 29
        Case "waterloadpressure": GetLoadTypeFromString = 30
        Case "liveloadsurcharge": GetLoadTypeFromString = 31
        Case "lockedinforces": GetLoadTypeFromString = 32
        Case "pedestrianll": GetLoadTypeFromString = 33
        Case "prestress": GetLoadTypeFromString = 34
        Case "hyperstatic": GetLoadTypeFromString = 35
        Case "bouyancy": GetLoadTypeFromString = 36
        Case "streamflow": GetLoadTypeFromString = 37
        Case "impact": GetLoadTypeFromString = 38
        Case "construction": GetLoadTypeFromString = 39
        Case Else: GetLoadTypeFromString = 8  ' Default to Other
    End Select
End Function

