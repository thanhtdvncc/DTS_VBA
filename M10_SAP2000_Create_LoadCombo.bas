Attribute VB_Name = "M10_SAP2000_Create_LoadCombo"

Option Explicit
'===============================================================
' Module: modSAP2000_AutoCreateCombos (Enhanced with Database Import)
' Purpose: Auto create load combinations with intelligent scoring
' Features: Multi-factor scoring, pattern detection, sequence analysis
'           Database import for large combo sets (>200)
'===============================================================

Public isFormOpen_Combos As Boolean

' Scoring thresholds
Private Const SCORE_THRESHOLD As Long = 10
Private Const DATABASE_THRESHOLD As Long = 200  ' Switch to database method if combos > 200

' NOTE: Do NOT redeclare database globals here.
' Globals are defined in modSAP2000_DBExport to maintain single source-of-truth:
' g_SelectedTableKey, g_TableVersion, g_FieldKeysIncluded, g_NumberRecords, g_ExportedSheetName, g_ExportedWorkbookName

'--------------------
' Entry point: open form
Public Sub AutoCreateLoadCombinations()
    On Error GoTo ERR_HANDLER

    If isFormOpen_Combos Then
        MsgBox "Load Combination form is already open. Please complete or close it first.", vbExclamation
        Exit Sub
    End If
    
    ConnectSAP2000
    
    If SapModel Is Nothing Then
        On Error Resume Next
        If VBA.err.number = 0 Then
            err.Clear
            'CallByName Application.VBE.VBProjects, "ToIgnore", VbMethod
        End If
        On Error GoTo 0

        If SapModel Is Nothing Then
            MsgBox "SAP2000 model object (SapModel) is not set. Please connect using your connection module first.", vbCritical
            Exit Sub
        End If
    End If

    isFormOpen_Combos = True
    Dim frm As frmPasteLoadCombinations
    Set frm = New frmPasteLoadCombinations
    frm.Show vbModeless

    Exit Sub

ERR_HANDLER:
    isFormOpen_Combos = False
    MsgBox "Error opening form: " & err.description, vbCritical
End Sub

'===============================================================
' CANDIDATE ARRAY STRUCTURE (unchanged)
'===============================================================

'===============================================================
' ENHANCED PARSING - Add Phase 7: Post-filtering
'===============================================================

'===============================================================
' MAIN PARSING FUNCTION - Updated with Phase 7
'===============================================================
Public Function ParseLoadComboText(inputText As String) As Object
    ' Normalize newlines
    inputText = Replace(inputText, vbCrLf, vbLf)
    inputText = Replace(inputText, vbCr, vbLf)
    
    Dim lines As Variant
    lines = Split(inputText, vbLf)
    
    ' PASS 1: Find section headers
    Dim sectionLines As Object
    Set sectionLines = FindSectionHeaders(lines)
    
    ' PASS 2: Extract all candidates with metadata
    Dim candidates As Collection
    Set candidates = ExtractCandidates(lines, sectionLines)
    
    If candidates.count = 0 Then
        Set ParseLoadComboText = CreateObject("Scripting.Dictionary")
        Exit Function
    End If
    
    ' PASS 3: Pattern analysis
    Dim patternInfo As Object
    Set patternInfo = AnalyzePatterns(candidates)
    
    ' PASS 4: Sequence analysis
    Dim sequenceInfo As Object
    Set sequenceInfo = AnalyzeSequences(candidates)
    
    ' PASS 5: Score all candidates
    ScoreCandidates candidates, patternInfo, sequenceInfo
    
    ' PASS 6: Filter by score and create final dictionary
    Dim finalCombos As Object
    Set finalCombos = FilterAndCreateCombos(candidates)
    
    ' *** PASS 7: POST-FILTER - Clean formulas from trailing noise ***
    PostFilterCleanFormulas finalCombos
    
    Set ParseLoadComboText = finalCombos
End Function

'===============================================================
' CREATe COMBOS IN SAP2000 - enhanced with overwrite check
' Updated: Do NOT show MsgBox for returned result; returns status string only
'===============================================================
Public Function CreateCombosInSAP(combos As Object, Optional overwriteCombos As Boolean = False) As String
    On Error GoTo ErrHandler
    
    CreateCombosInSAP = "" ' default empty return
    
    ' Check existing combos count if overwrite mode
    Dim existingCount As Long
    existingCount = 0
    
    If overwriteCombos Then
        Dim NumberNames As Long, MyName() As String
        Dim ret As Long
        ret = SapModel.RespCombo.GetNameList(NumberNames, MyName)
        If ret = 0 Then existingCount = NumberNames
        
        ' Check if existing combos > 200 and offer database deletion
        If existingCount > DATABASE_THRESHOLD Then
            Dim userChoice As VbMsgBoxResult
            userChoice = MsgBox( _
                "Overwrite Mode: " & existingCount & " existing load combinations detected." & vbCrLf & vbCrLf & _
                "For large numbers, You must DELETE their first!" & vbCrLf & vbCrLf & _
                "Choose the DELETE method:" & vbCrLf & _
                "- YES = Use Database Delete (Fast, Recommended)" & vbCrLf & _
                "- NO = Use API Delete (Slower)" & vbCrLf & _
                "- CANCEL = Abort operation" & vbCrLf, _
                vbYesNoCancel + vbQuestion + vbDefaultButton1, "Large Existing Combo Set Detected")
            
            If userChoice = vbCancel Then
                CreateCombosInSAP = "CANCELLED"
                Exit Function
            ElseIf userChoice = vbYes Then
                Call PrepareDeleteViaDatabase(existingCount)
                CreateCombosInSAP = "DATABASE_COMBINATIONS_DELETED" & vbCrLf & "TRY TO CREATE NEW COMBINATIONS AGAIN"
                Exit Function
            End If
        End If
    End If
    
    ' Check new combo count and offer database method
    If combos.count > DATABASE_THRESHOLD Then
        Dim userChoice2 As VbMsgBoxResult
        userChoice2 = MsgBox( _
            "You are creating " & combos.count & " load combinations." & vbCrLf & vbCrLf & _
            "For large numbers, Database Import is much faster!" & vbCrLf & vbCrLf & _
            "Choose method:" & vbCrLf & _
            "- YES = Use Database Import (Fast, Recommended)" & vbCrLf & _
            "- NO = Use API Method (Slower)" & vbCrLf & _
            "- CANCEL = Abort operation" & vbCrLf, _
            vbYesNoCancel + vbQuestion + vbDefaultButton1, "Large Combo Set Detected")
        
        If userChoice2 = vbCancel Then
            CreateCombosInSAP = "CANCELLED"
            Exit Function
        ElseIf userChoice2 = vbYes Then
            Call PrepareImportViaDatabase(combos, overwriteCombos)
            CreateCombosInSAP = "DATABASE_IMPORT_PREPARED"
            Exit Function
        End If
    End If
    
    ' Use original API method
    CreateCombosInSAP = CreateCombosViaAPI(combos, overwriteCombos)
    Exit Function

ErrHandler:
    CreateCombosInSAP = "CRITICAL ERROR: " & err.description
End Function

'===============================================================
' ORIGINAL API METHOD
'===============================================================
Private Function CreateCombosViaAPI(combos As Object, overwriteCombos As Boolean) As String
    Dim createdCount As Long, skippedCount As Long, errorCount As Long, deletedCount As Long
    Dim errorLog As String
    createdCount = 0: skippedCount = 0: errorCount = 0: deletedCount = 0: errorLog = ""

    Dim NumberNames As Long, MyName() As String, ret As Long

    On Error Resume Next
    ret = SapModel.RespCombo.GetNameList(NumberNames, MyName)
    If err.number <> 0 Or ret <> 0 Then
        errorLog = errorLog & "ERROR: Unable to retrieve existing combo list." & vbCrLf
        errorCount = errorCount + 1
        err.Clear
        NumberNames = 0
    End If
    On Error GoTo 0

    Dim existingNames As Object
    Set existingNames = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = 0 To NumberNames - 1
        existingNames.Add UCase(MyName(i)), True
    Next i

    If overwriteCombos Then
        For i = 0 To NumberNames - 1
            ret = SapModel.RespCombo.Delete(MyName(i))
            If ret = 0 Then deletedCount = deletedCount + 1
        Next i
        Set existingNames = CreateObject("Scripting.Dictionary")
    End If

    Dim comboKey As Variant
    For Each comboKey In combos.keys
        Dim comboNameStr As String
        comboNameStr = CStr(comboKey)

        If Not existingNames.exists(UCase(comboNameStr)) Then
            Dim cInfo As Variant
            cInfo = combos(comboKey)
            Dim comboType As Long
            comboType = CLng(cInfo(2))

            ret = SapModel.RespCombo.Add(comboNameStr, comboType)
            If ret = 0 Then
                Dim formula As String
                formula = CStr(cInfo(1))

                If AddCasesToCombo(comboNameStr, formula, errorLog) Then
                    createdCount = createdCount + 1
                Else
                    errorLog = errorLog & "ERROR: Failed to add cases to '" & comboNameStr & "'." & vbCrLf
                    errorCount = errorCount + 1
                    SapModel.RespCombo.Delete comboNameStr
                End If
            Else
                errorLog = errorLog & "ERROR: Creating combo '" & comboNameStr & "' failed (code " & ret & ")" & vbCrLf
                errorCount = errorCount + 1
            End If
        Else
            skippedCount = skippedCount + 1
        End If
    Next comboKey

    Dim summary As String
    summary = "===================================" & vbCrLf & _
              "  LOAD COMBO CREATION SUMMARY (API)" & vbCrLf & _
              "===================================" & vbCrLf & vbCrLf & _
              "Combos in request: " & combos.count & vbCrLf & _
              "Created: " & createdCount & vbCrLf & _
              "Skipped (already exist): " & skippedCount & vbCrLf

    If overwriteCombos Then summary = summary & "Deleted (old combos): " & deletedCount & vbCrLf

    summary = summary & "Errors: " & errorCount & vbCrLf & vbCrLf

    If errorCount > 0 Then
        summary = summary & "Details:" & vbCrLf & "-----------------------------------" & vbCrLf & errorLog
    End If

    CreateCombosViaAPI = summary
End Function

'===============================================================
' PREPARE DELETE VIA DATABASE - Export to Database sheet
' Updated: Sub now exports only ONE combo row for DB Delete, then
'          triggers UpdateToSAP2000 automatically, and finally
'          removes the last combo via API if only one remains.
' All steps run automatically (no user confirmation) - errors show MsgBox.
'===============================================================
Private Sub PrepareDeleteViaDatabase(existingCount As Long)
    On Error GoTo ErrHandler
    
    ' Get existing combo list
    Dim NumberNames As Long, MyName() As String, ret As Long
    Dim FieldKeyArr() As String, FieldNameArr() As String
    Dim DescriptionArr() As String, UnitsStringArr() As String, IsImportableArr() As Boolean
    Dim TableVersion As Long, NumberFields As Long
    Dim ws As Worksheet
    Dim tableKey As String
    Dim idxComboName As Long
    Dim nCols As Long
    Dim c As Long, i As Long
    Dim dataRange As Object
    
    ret = SapModel.RespCombo.GetNameList(NumberNames, MyName)
    
    If ret <> 0 Or NumberNames = 0 Then
        MsgBox "ERROR: Failed to retrieve existing combos", vbCritical, "Error"
        Exit Sub
    End If
    
    ' Find or create Database sheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Database")
    On Error GoTo ErrHandler
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.Name = "Database"
    End If
    
    ' Clear existing data
    ws.Cells.Clear
    ws.Cells.ClearFormats
    
    ' Build table in SAP2000 Database format
    tableKey = "Combination Definitions"
    
    ' Get field structure
    ret = SapModel.DatabaseTables.GetAllFieldsInTable( _
        tableKey, TableVersion, NumberFields, _
        FieldKeyArr, FieldNameArr, DescriptionArr, UnitsStringArr, IsImportableArr)
    
    If ret <> 0 Then
        MsgBox "ERROR: Failed to get database table structure", vbCritical, "Error"
        Exit Sub
    End If
    
    ' Find ComboName field index
    idxComboName = FindFieldIndex(FieldKeyArr, "ComboName")
    
    If idxComboName < 0 Then
        MsgBox "ERROR: ComboName field not found", vbCritical, "Error"
        Exit Sub
    End If
    
    ' ROW 1: Table Title
    nCols = UBound(FieldKeyArr) - LBound(FieldKeyArr) + 1
    
    ws.Cells(1, 1).Value = "SAP2000 Database: " & tableKey
    ws.Range(ws.Cells(1, 1), ws.Cells(1, nCols)).Merge
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 1).Font.Size = 12
    ws.Cells(1, 1).HorizontalAlignment = -4108 ' xlCenter
    ws.Cells(1, 1).Interior.color = RGB(68, 114, 196)
    ws.Cells(1, 1).Font.color = RGB(255, 255, 255)
    ws.rows(1).rowHeight = 25
    
    ' ROW 2: Field Keys
    For c = 0 To nCols - 1
        ws.Cells(2, c + 1).Value = FieldKeyArr(c)
        ws.Cells(2, c + 1).Font.Bold = True
        ws.Cells(2, c + 1).Interior.color = RGB(217, 217, 217)
    Next c
    
    ' ROW 3: Field Names
    For c = 0 To nCols - 1
        ws.Cells(3, c + 1).Value = FieldNameArr(c)
        ws.Cells(3, c + 1).Font.Italic = True
    Next c
    
    ' ROW 4: Units
    For c = 0 To nCols - 1
        ws.Cells(4, c + 1).Value = UnitsStringArr(c)
        ws.Cells(4, c + 1).Font.color = RGB(128, 128, 128)
    Next c
    
    ' ROW 5+: Data - Only ONE ComboName row for deletion (automated approach)
    ' Strategy: Keep only a single record in the Database sheet so that UpdateToSAP2000
    ' will process records and effectively remove data for other combos (avoids keeping all names).
    ' We write only the first combo name.
    ws.Cells(5, idxComboName + 1).Value = MyName(0)
    
    ' NOTE: we intentionally write only one record (row 5). The import/update step
    '       will be executed automatically below.
    
    ' Format
    ws.rows("2:4").Font.Bold = True
    ws.Columns("A:AZ").AutoFit
    
    ' Add borders for header + single data row
    Set dataRange = ws.Range(ws.Cells(2, 1), ws.Cells(5, nCols))
    dataRange.Borders.LineStyle = 1 ' xlContinuous
    dataRange.Borders.Weight = 2 ' xlThin
    
    ' Activate the sheet
    ws.Activate
    ws.Range("A1").Select
    
    ' *** SET GLOBAL VARIABLES (in DBExport module globals) - IMPORTANT ***
    ' Note: these globals are declared in modSAP2000_DBExport
    g_SelectedTableKey = tableKey
    g_TableVersion = TableVersion
    g_ExportedSheetName = ws.Name
    g_ExportedWorkbookName = ws.Parent.Name
    ' We set number records to 1 because we exported only one row intentionally
    g_NumberRecords = 1
    
    ' Copy FieldKeyArr to global
    ReDim g_FieldKeysIncluded(LBound(FieldKeyArr) To UBound(FieldKeyArr))
    For i = LBound(FieldKeyArr) To UBound(FieldKeyArr)
        g_FieldKeysIncluded(i) = FieldKeyArr(i)
    Next i
    
    ' ---- AUTOMATED STEPS ----
    ' 1) Trigger UpdateToSAP2000 to push the Database sheet to SAP2000 (Update operation)
    '    This sub is expected to exist in modSAP2000_DBExport and to perform the import/update.
    On Error Resume Next
    Call UpdateToSAP2000
    If err.number <> 0 Then
        MsgBox "ERROR: UpdateToSAP2000 failed: " & err.description, vbCritical, "Error"
        err.Clear
        Exit Sub
    End If
    On Error GoTo ErrHandler
    
    ' 2) After update, check current number of combos in the model.
    ret = SapModel.RespCombo.GetNameList(NumberNames, MyName)
    If ret <> 0 Then
        MsgBox "ERROR: Unable to retrieve combo list after UpdateToSAP2000", vbCritical, "Error"
        Exit Sub
    End If
    
    ' 3) If only one combo remains in the model, delete it via API (final cleanup)
    If NumberNames = 1 Then
        ret = SapModel.RespCombo.Delete(MyName(0))
        If ret <> 0 Then
            MsgBox "ERROR: Failed to delete last combo '" & MyName(0) & "' via API", vbCritical, "Error"
            Exit Sub
        End If
    Else
        ' If more than one remains, attempt to delete remaining combos via API loop
        ' (This covers edge cases where UpdateToSAP2000 did not remove all combos.)
        Dim delCount As Long
        delCount = 0
        For i = 0 To NumberNames - 1
            On Error Resume Next
            ret = SapModel.RespCombo.Delete(MyName(i))
            If ret = 0 Then delCount = delCount + 1
            On Error GoTo ErrHandler
        Next i
    End If
    
    ' No MsgBox by default - operation completed
    Exit Sub
    
ErrHandler:
    MsgBox "ERROR preparing delete: " & err.description, vbCritical, "Error"
End Sub

'===============================================================
' PREPARE IMPORT VIA DATABASE - Export to Database sheet
' Updated: Sub, only shows MsgBox, sets global variables in modSAP2000_DBExport
'===============================================================
Private Sub PrepareImportViaDatabase(combos As Object, overwriteCombos As Boolean)
    On Error GoTo ErrHandler
    
    Dim tableKey As String
    tableKey = "Combination Definitions"
    
    ' Get database structure
    Dim ret As Long, TableVersion As Long, NumberFields As Long
    Dim FieldKeyArr() As String, FieldNameArr() As String
    Dim DescriptionArr() As String, UnitsStringArr() As String, IsImportableArr() As Boolean
    Dim existingCombos As Object
    Dim FieldKeysIncluded() As String, tableData() As String
    Dim totalRows As Long, rowsBuilt As Long
    Dim ws As Worksheet
    Dim numFields As Long
    Dim c As Long, i As Long, row As Long
    Dim fk As String, foundIdx As Long
    
    ret = SapModel.DatabaseTables.GetAllFieldsInTable( _
        tableKey, TableVersion, NumberFields, _
        FieldKeyArr, FieldNameArr, DescriptionArr, UnitsStringArr, IsImportableArr)
    
    If ret <> 0 Then
        MsgBox "ERROR: Failed to get database table structure", vbCritical, "Error"
        Exit Sub
    End If
    
    ' Get existing combos (if not overwriting)
    Set existingCombos = CreateObject("Scripting.Dictionary")
    
    If Not overwriteCombos Then
        Dim NumberNames As Long, MyName() As String
        ret = SapModel.RespCombo.GetNameList(NumberNames, MyName)
        If ret = 0 And NumberNames > 0 Then
            For i = 0 To NumberNames - 1
                existingCombos.Add UCase(MyName(i)), True
            Next i
        End If
    End If
    
    ' Build table data
    If Not BuildComboTableData(combos, existingCombos, FieldKeyArr, _
                               FieldKeysIncluded, tableData, totalRows, rowsBuilt) Then
        MsgBox "ERROR: Failed to build table data", vbCritical, "Error"
        Exit Sub
    End If
    
    If rowsBuilt = 0 Then
        MsgBox "No new combinations to create (all already exist)", vbInformation, "Info"
        Exit Sub
    End If
    
    ' Export to Database sheet in SAP2000 format
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Database")
    On Error GoTo ErrHandler
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.Name = "Database"
    End If
    
    ' Clear existing data
    ws.Cells.Clear
    ws.Cells.ClearFormats
    
    ' Build in SAP2000 Database format
    numFields = UBound(FieldKeysIncluded) - LBound(FieldKeysIncluded) + 1
    
    ' ROW 1: Table Title
    ws.Cells(1, 1).Value = "SAP2000 Database: " & tableKey
    ws.Range(ws.Cells(1, 1), ws.Cells(1, numFields)).Merge
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 1).Font.Size = 12
    ws.Cells(1, 1).HorizontalAlignment = -4108 ' xlCenter
    ws.Cells(1, 1).Interior.color = RGB(68, 114, 196)
    ws.Cells(1, 1).Font.color = RGB(255, 255, 255)
    ws.rows(1).rowHeight = 25
    
    ' ROW 2: Field Keys
    For c = 0 To numFields - 1
        ws.Cells(2, c + 1).Value = FieldKeysIncluded(c)
        ws.Cells(2, c + 1).Font.Bold = True
        ws.Cells(2, c + 1).Interior.color = RGB(217, 217, 217)
    Next c
    
    ' ROW 3: Field Names (from metadata)
    For c = 0 To numFields - 1
        fk = FieldKeysIncluded(c)
        foundIdx = FindFieldIndex(FieldKeyArr, fk)
        If foundIdx >= 0 Then
            ws.Cells(3, c + 1).Value = FieldNameArr(foundIdx)
            ws.Cells(3, c + 1).Font.Italic = True
        End If
    Next c
    
    ' ROW 4: Units
    For c = 0 To numFields - 1
        fk = FieldKeysIncluded(c)
        foundIdx = FindFieldIndex(FieldKeyArr, fk)
        If foundIdx >= 0 Then
            ws.Cells(4, c + 1).Value = UnitsStringArr(foundIdx)
            ws.Cells(4, c + 1).Font.color = RGB(128, 128, 128)
        End If
    Next c
    
    ' ROW 5+: Data rows
    For row = 0 To rowsBuilt - 1
        For c = 0 To numFields - 1
            ws.Cells(5 + row, c + 1).Value = tableData(row * numFields + c)
        Next c
    Next row
    
    ' Format
    ws.rows("2:4").Font.Bold = True
    ws.Columns("A:AZ").AutoFit
    
    ' Add borders
    If rowsBuilt > 0 Then
        Dim dataRange As Object
        Set dataRange = ws.Range(ws.Cells(2, 1), ws.Cells(rowsBuilt + 4, numFields))
        dataRange.Borders.LineStyle = 1 ' xlContinuous
        dataRange.Borders.Weight = 2 ' xlThin
    End If
    
    ' Activate sheet
    ws.Activate
    ws.Range("A1").Select
    
    ' *** SET GLOBAL VARIABLES (in DBExport module globals) - IMPORTANT ***
    g_SelectedTableKey = tableKey
    g_TableVersion = TableVersion
    g_ExportedSheetName = ws.Name
    g_ExportedWorkbookName = ws.Parent.Name
    g_NumberRecords = rowsBuilt
    
    ' Copy FieldKeysIncluded to global (keep same bounds)
    ReDim g_FieldKeysIncluded(LBound(FieldKeysIncluded) To UBound(FieldKeysIncluded))
    For i = LBound(FieldKeysIncluded) To UBound(FieldKeysIncluded)
        g_FieldKeysIncluded(i) = FieldKeysIncluded(i)
    Next i
    
    ' Show message - single message only
    MsgBox "Prepared " & rowsBuilt & " combination rows for import" & vbCrLf & _
        "Total combos: " & combos.count & vbCrLf & _
        "Skipped (already exist): " & (totalRows - rowsBuilt) & vbCrLf & vbCrLf & _
        "NEXT STEPS:" & vbCrLf & _
        "1. Review data in 'Database' sheet (now active)" & vbCrLf & _
        "2. You can edit the data (row 5 onwards) if needed" & vbCrLf & _
        "3. To import, click 'Import to SAP2000' button" & vbCrLf & vbCrLf & _
        "The data is ready in SAP2000 Database format.", _
        vbInformation, "Database Import Ready"
    
    Exit Sub
    
ErrHandler:
    MsgBox "ERROR: " & err.description, vbCritical, "Error"
End Sub

'===============================================================
' BUILD DATABASE TABLE DATA
'===============================================================
Private Function BuildComboTableData( _
    combos As Object, existingCombos As Object, FieldKeyArr() As String, _
    ByRef FieldKeysIncluded() As String, ByRef tableData() As String, _
    ByRef totalRows As Long, ByRef rowsBuilt As Long) As Boolean
    
    On Error GoTo ErrHandler
    BuildComboTableData = False
    
    ' Find field indices
    Dim idxComboName As Long, idxComboType As Long, idxAutoDesign As Long
    Dim idxCaseType As Long, idxCaseName As Long, idxScaleFactor As Long
    Dim idxSteelDesign As Long, idxConcDesign As Long
    Dim idxAlumDesign As Long, idxColdDesign As Long, idxGUID As Long, idxNotes As Long
    
    idxComboName = FindFieldIndex(FieldKeyArr, "ComboName")
    idxComboType = FindFieldIndex(FieldKeyArr, "ComboType")
    idxAutoDesign = FindFieldIndex(FieldKeyArr, "AutoDesign")
    idxCaseType = FindFieldIndex(FieldKeyArr, "CaseType")
    idxCaseName = FindFieldIndex(FieldKeyArr, "CaseName")
    idxScaleFactor = FindFieldIndex(FieldKeyArr, "ScaleFactor")
    idxSteelDesign = FindFieldIndex(FieldKeyArr, "SteelDesign")
    idxConcDesign = FindFieldIndex(FieldKeyArr, "ConcDesign")
    idxAlumDesign = FindFieldIndex(FieldKeyArr, "AlumDesign")
    idxColdDesign = FindFieldIndex(FieldKeyArr, "ColdDesign")
    idxGUID = FindFieldIndex(FieldKeyArr, "GUID")
    idxNotes = FindFieldIndex(FieldKeyArr, "Notes")
    
    If idxComboName < 0 Or idxCaseName < 0 Or idxScaleFactor < 0 Then Exit Function
    
    ' Build field keys list
    Dim numFields As Long
    numFields = UBound(FieldKeyArr) - LBound(FieldKeyArr) + 1
    ReDim FieldKeysIncluded(0 To numFields - 1)
    
    Dim i As Long
    For i = 0 To numFields - 1
        FieldKeysIncluded(i) = FieldKeyArr(i)
    Next i
    
    ' Count total rows
    totalRows = 0
    Dim comboKey As Variant
    For Each comboKey In combos.keys
        Dim cInfo As Variant
        cInfo = combos(comboKey)
        Dim terms As Collection
        Set terms = ParseFormula(CStr(cInfo(1)))
        If Not terms Is Nothing Then totalRows = totalRows + terms.count
    Next comboKey
    
    If totalRows = 0 Then Exit Function
    
    ' Allocate table
    ReDim tableData(0 To totalRows * numFields - 1)
    
    ' Fill table data
    rowsBuilt = 0
    For Each comboKey In combos.keys
        Dim comboNameStr As String
        comboNameStr = CStr(comboKey)
        
        ' Skip if exists
        If existingCombos.count > 0 Then
            If existingCombos.exists(UCase(comboNameStr)) Then GoTo NextCombo
        End If
        
        cInfo = combos(comboKey)
        Dim comboType As Long
        comboType = CLng(cInfo(2))
        
        Set terms = ParseFormula(CStr(cInfo(1)))
        If terms Is Nothing Or terms.count = 0 Then GoTo NextCombo
        
        ' Add rows for each case
        Dim termIdx As Long
        For termIdx = 1 To terms.count
            Dim termInfo As Variant
            termInfo = terms(termIdx)
            
            Dim scaleFactor As Double, caseName As String
            scaleFactor = CDbl(termInfo(0))
            caseName = CStr(termInfo(1))
            
            ' Initialize all fields
            For i = 0 To numFields - 1
                tableData(rowsBuilt * numFields + i) = ""
            Next i
            
            ' Set fields
            If idxComboName >= 0 Then tableData(rowsBuilt * numFields + idxComboName) = comboNameStr
            If termIdx = 1 And idxComboType >= 0 Then tableData(rowsBuilt * numFields + idxComboType) = GetComboTypeString(comboType)
            If termIdx = 1 And idxAutoDesign >= 0 Then tableData(rowsBuilt * numFields + idxAutoDesign) = "No"
            If idxCaseType >= 0 Then tableData(rowsBuilt * numFields + idxCaseType) = DetectCaseType(caseName)
            If idxCaseName >= 0 Then tableData(rowsBuilt * numFields + idxCaseName) = caseName
            If idxScaleFactor >= 0 Then tableData(rowsBuilt * numFields + idxScaleFactor) = CStr(scaleFactor)
            
            If termIdx = 1 Then
                If idxSteelDesign >= 0 Then tableData(rowsBuilt * numFields + idxSteelDesign) = "None"
                If idxConcDesign >= 0 Then tableData(rowsBuilt * numFields + idxConcDesign) = "None"
                If idxAlumDesign >= 0 Then tableData(rowsBuilt * numFields + idxAlumDesign) = "None"
                If idxColdDesign >= 0 Then tableData(rowsBuilt * numFields + idxColdDesign) = "None"
            End If
            
            rowsBuilt = rowsBuilt + 1
        Next termIdx
NextCombo:
    Next comboKey
    
    BuildComboTableData = True
    Exit Function

ErrHandler:
    BuildComboTableData = False
End Function

'===============================================================
' All other parsing/helper functions unchanged from original
' (FindFieldIndex, DetectCaseType, ParseFormula, etc.)
' For brevity they are included below unchanged.
'===============================================================

Private Function FindFieldIndex(FieldKeyArr() As String, fieldKey As String) As Long
    FindFieldIndex = -1
    If Not IsArray(FieldKeyArr) Then Exit Function
    Dim i As Long
    For i = LBound(FieldKeyArr) To UBound(FieldKeyArr)
        If FieldKeyArr(i) = fieldKey Then
            FindFieldIndex = i
            Exit Function
        End If
    Next i
End Function

Private Function DetectCaseType(caseName As String) As String
    Dim upper As String
    upper = UCase(caseName)
    
    If InStr(upper, "MODAL") > 0 Or InStr(upper, "MODE") > 0 Then
        DetectCaseType = "Modal"
    ElseIf InStr(upper, "RESP") > 0 Or InStr(upper, "SPEC") > 0 Then
        DetectCaseType = "Response Spectrum"
    ElseIf InStr(upper, "TIME") > 0 Or InStr(upper, "HIST") > 0 Then
        DetectCaseType = "Time History"
    ElseIf InStr(upper, "COMBO") > 0 Or InStr(upper, "COMB") > 0 Then
        DetectCaseType = "Linear Add"
    Else
        DetectCaseType = "Linear Static"
    End If
End Function

Private Function FindSectionHeaders(lines As Variant) As Object
    Dim sectionLines As Object
    Set sectionLines = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        Dim lineText As String
        lineText = Trim(CStr(lines(i)))
        
        If IsSectionHeader(lineText) Then
            sectionLines.Add i, True
        End If
    Next i
    
    Set FindSectionHeaders = sectionLines
End Function

Private Function IsSectionHeader(lineText As String) As Boolean
    IsSectionHeader = False
    
    If Len(lineText) = 0 Or Len(lineText) > 100 Then Exit Function
    
    Dim upper As String
    upper = UCase(lineText)
    
    If IsNumericSection(lineText) Then
        IsSectionHeader = True
        Exit Function
    End If
    
    If Left(upper, 1) >= "A" And Left(upper, 1) <= "Z" Then
        If Len(lineText) >= 2 Then
            If mid(lineText, 2, 1) = "." Or mid(lineText, 2, 1) = ")" Then
                IsSectionHeader = True
                Exit Function
            End If
        End If
    End If
    
    If InStr(upper, "LOAD COMBINATION") > 0 Or InStr(upper, "LOAD COMBO") > 0 Or _
       InStr(upper, "MEMBER DESIGN") > 0 Or InStr(upper, "EMPTY CASE") > 0 Or _
       InStr(upper, "OPERATION") > 0 Or InStr(upper, "TEST") > 0 Then
        IsSectionHeader = True
    End If
End Function

Private Function ExtractCandidates(lines As Variant, sectionLines As Object) As Collection
    Dim candidates As Collection
    Set candidates = New Collection
    
    Dim lastSectionLine As Long
    lastSectionLine = -100
    
    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        Dim lineText As String
        lineText = Trim(CStr(lines(i)))
        
        If Len(lineText) = 0 Then GoTo NextLine
        
        If sectionLines.exists(i) Then
            lastSectionLine = i
            GoTo NextLine
        End If
        
        If IsObviousNoise(lineText) Then GoTo NextLine
        
        Dim comboName As String, formula As String
        
        If TryParseComboLine(lineText, comboName, formula) Then
            If IsValidComboFormula(formula) Then
                Dim candidate(0 To 6) As Variant
                candidate(0) = i
                candidate(1) = comboName
                candidate(2) = formula
                candidate(3) = CLng(0)
                candidate(4) = (i - lastSectionLine) <= 20
                
                Dim prefix As String, number As Long
                If ExtractPattern(comboName, prefix, number) Then
                    candidate(5) = prefix
                    candidate(6) = number
                Else
                    candidate(5) = ""
                    candidate(6) = CLng(0)
                End If
                
                candidates.Add candidate
            End If
        End If
NextLine:
    Next i
    
    Set ExtractCandidates = candidates
End Function

Private Function AnalyzePatterns(candidates As Collection) As Object
    Dim patternInfo As Object
    Set patternInfo = CreateObject("Scripting.Dictionary")
    
    Dim patternCounts As Object
    Set patternCounts = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    For i = 1 To candidates.count
        Dim candidate As Variant
        candidate = candidates(i)
        
        Dim patternPrefix As String
        patternPrefix = CStr(candidate(5))
        
        If Len(patternPrefix) > 0 Then
            If Not patternCounts.exists(patternPrefix) Then
                patternCounts.Add patternPrefix, 0
            End If
            patternCounts(patternPrefix) = patternCounts(patternPrefix) + 1
        End If
    Next i
    
    Dim dominantPattern As String, maxCount As Long
    dominantPattern = ""
    maxCount = 0
    
    Dim key As Variant
    For Each key In patternCounts.keys
        If patternCounts(key) > maxCount Then
            maxCount = patternCounts(key)
            dominantPattern = CStr(key)
        End If
    Next key
    
    patternInfo.Add "DominantPattern", dominantPattern
    patternInfo.Add "DominantCount", maxCount
    patternInfo.Add "PatternCounts", patternCounts
    
    Set AnalyzePatterns = patternInfo
End Function

Private Function AnalyzeSequences(candidates As Collection) As Object
    Dim sequenceInfo As Object
    Set sequenceInfo = CreateObject("Scripting.Dictionary")
    
    Dim groups As Object
    Set groups = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    For i = 1 To candidates.count
        Dim candidate As Variant
        candidate = candidates(i)
        
        Dim patternPrefix As String, patternNumber As Long
        patternPrefix = CStr(candidate(5))
        patternNumber = CLng(candidate(6))
        
        If Len(patternPrefix) > 0 Then
            If Not groups.exists(patternPrefix) Then
                Dim newList As Collection
                Set newList = New Collection
                groups.Add patternPrefix, newList
            End If
            
            Dim grp As Collection
            Set grp = groups(patternPrefix)
            grp.Add patternNumber
        End If
    Next i
    
    Dim sequentialGroups As Object
    Set sequentialGroups = CreateObject("Scripting.Dictionary")
    
    Dim key As Variant
    For Each key In groups.keys
        Dim numbers As Collection
        Set numbers = groups(key)
        
        If numbers.count >= 3 Then
            If IsSequential(numbers) Then
                sequentialGroups.Add key, True
            End If
        End If
    Next key
    
    sequenceInfo.Add "SequentialGroups", sequentialGroups
    
    Set AnalyzeSequences = sequenceInfo
End Function

Private Function IsSequential(numbers As Collection) As Boolean
    IsSequential = False
    If numbers.count < 3 Then Exit Function
    
    Dim arr() As Long
    ReDim arr(1 To numbers.count)
    
    Dim i As Long
    For i = 1 To numbers.count
        arr(i) = numbers(i)
    Next i
    
    Dim j As Long, temp As Long
    For i = 1 To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(i) > arr(j) Then
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            End If
        Next j
    Next i
    
    Dim increasing As Long
    increasing = 0
    For i = 1 To UBound(arr) - 1
        If arr(i + 1) > arr(i) Then increasing = increasing + 1
    Next i
    
    If increasing >= (UBound(arr) - 1) * 0.8 Then IsSequential = True
End Function

Private Sub ScoreCandidates(candidates As Collection, patternInfo As Object, sequenceInfo As Object)
    Dim dominantPattern As String, dominantCount As Long
    
    dominantPattern = ""
    dominantCount = 0
    
    If patternInfo.exists("DominantPattern") Then dominantPattern = CStr(patternInfo("DominantPattern"))
    If patternInfo.exists("DominantCount") Then dominantCount = CLng(patternInfo("DominantCount"))
    
    Dim sequentialGroups As Object
    Set sequentialGroups = sequenceInfo("SequentialGroups")
    
    Dim updatedCandidates As Collection
    Set updatedCandidates = New Collection
    
    Dim i As Long
    For i = 1 To candidates.count
        Dim candidate As Variant
        candidate = candidates(i)
        
        Dim comboName As String, formula As String, patternPrefix As String
        Dim patternNumber As Long, afterSection As Boolean
        
        comboName = CStr(candidate(1))
        formula = CStr(candidate(2))
        afterSection = CBool(candidate(4))
        patternPrefix = CStr(candidate(5))
        patternNumber = CLng(candidate(6))
        
        Dim Score As Long
        Score = 5
        
        If Len(patternPrefix) > 0 Then
            Score = Score + 10
            If dominantCount >= 3 And patternPrefix = dominantPattern Then Score = Score + 15
        End If
        
        If sequentialGroups.exists(patternPrefix) Then Score = Score + 10
        
        Dim termCount As Long
        termCount = CountFormulaTerms(formula)
        If termCount >= 3 Then Score = Score + 5
        If termCount >= 5 Then Score = Score + 8
        If HasScaleFactor(formula) Then Score = Score + 5
        If HasMultipleOperators(formula) Then Score = Score + 3
        If afterSection Then Score = Score + 8
        
        If ContainsNoiseKeywords(formula) Then Score = Score - 15
        If HasInvalidCharacters(formula) Then Score = Score - 20
        If IsDocumentNumber(comboName) Then Score = Score - 25
        If Len(comboName) > 20 Then Score = Score - 10
        
        candidate(3) = Score
        updatedCandidates.Add candidate
    Next i
    
    Do While candidates.count > 0
        candidates.Remove 1
    Loop
    
    For i = 1 To updatedCandidates.count
        candidates.Add updatedCandidates(i)
    Next i
End Sub

Private Function FilterAndCreateCombos(candidates As Collection) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    For i = 1 To candidates.count
        Dim candidate As Variant
        candidate = candidates(i)
        
        Dim comboName As String, formula As String, Score As Long
        comboName = CStr(candidate(1))
        formula = CStr(candidate(2))
        Score = CLng(candidate(3))
        
        If Score >= SCORE_THRESHOLD Then
            If Not result.exists(comboName) Then
                Dim comboInfo(0 To 2) As Variant
                comboInfo(0) = comboName
                comboInfo(1) = formula
                comboInfo(2) = 0
                
                result.Add comboName, comboInfo
            End If
        End If
    Next i
    
    Set FilterAndCreateCombos = result
End Function

Private Function CountFormulaTerms(formula As String) As Long
    Dim clean As String
    clean = Replace(formula, " ", "")
    
    Dim count As Long
    count = 1
    
    Dim i As Long
    For i = 1 To Len(clean)
        Dim ch As String
        ch = mid(clean, i, 1)
        
        If ch = "+" Then count = count + 1
        If ch = "-" And i > 1 Then count = count + 1
    Next i
    
    CountFormulaTerms = count
End Function

Private Function HasMultipleOperators(formula As String) As Boolean
    Dim count As Long
    count = 0
    
    Dim i As Long
    For i = 1 To Len(formula)
        Dim ch As String
        ch = mid(formula, i, 1)
        
        If ch = "+" Or (ch = "-" And i > 1) Then count = count + 1
    Next i
    
    HasMultipleOperators = (count >= 2)
End Function

Private Function ContainsNoiseKeywords(text As String) As Boolean
    Dim upper As String
    upper = UCase(text)
    
    ContainsNoiseKeywords = _
        InStr(upper, "STRUCTURE") > 0 Or InStr(upper, "SEISMIC") > 0 Or _
        InStr(upper, "GROUND") > 0 Or InStr(upper, "ABOVE") > 0 Or _
        InStr(upper, "UNDER") > 0 Or InStr(upper, "WEIGHT") > 0 Or _
        InStr(upper, "PRESSURE") > 0 Or InStr(upper, "CALCULATION") > 0 Or _
        InStr(upper, "SHEET") > 0 Or InStr(upper, "DOC") > 0
End Function

Private Function HasInvalidCharacters(text As String) As Boolean
    HasInvalidCharacters = InStr(text, ":") > 0 Or InStr(text, "=") > 0 Or InStr(text, "/") > 0
End Function

Private Function IsObviousNoise(lineText As String) As Boolean
    IsObviousNoise = False
    
    If Left(lineText, 1) = "#" Or Left(lineText, 2) = "//" Then
        IsObviousNoise = True
        Exit Function
    End If
    
    If InStr(lineText, "---") > 0 Or InStr(lineText, "===") > 0 Then
        IsObviousNoise = True
        Exit Function
    End If
    
    Dim upper As String
    upper = UCase(lineText)
    
    If InStr(upper, "COMBO") > 0 And InStr(upper, "TYPE") > 0 And InStr(upper, "FORMULA") > 0 Then
        IsObviousNoise = True
    End If
End Function

'===============================================================
' ENHANCED PARSING FUNCTIONS - Improved noise filtering
'===============================================================

' Main parsing function with improved noise handling
Private Function TryParseComboLine(lineText As String, ByRef comboName As String, ByRef formula As String) As Boolean
    TryParseComboLine = False
    comboName = ""
    formula = ""
    
    ' Normalize whitespace
    lineText = Replace(lineText, vbTab, " ")
    Do While InStr(lineText, "  ") > 0
        lineText = Replace(lineText, "  ", " ")
    Loop
    lineText = Trim(lineText)
    
    If Len(lineText) = 0 Then Exit Function
    
    ' Find the split between name and formula
    Dim splitPos As Long
    splitPos = FindNameFormulaSplit(lineText)
    
    If splitPos = 0 Then Exit Function
    
    comboName = Trim(Left(lineText, splitPos - 1))
    Dim rawFormula As String
    rawFormula = Trim(mid(lineText, splitPos))
    
    If Len(comboName) = 0 Or Len(rawFormula) = 0 Then Exit Function
    If Len(comboName) > 50 Then Exit Function
    
    ' *** NEW: Clean formula by removing trailing noise ***
    formula = CleanFormulaTrailingNoise(rawFormula)
    
    If Len(formula) = 0 Then Exit Function
    
    TryParseComboLine = True
End Function

'===============================================================
' FUNCTION: Clean trailing noise from formula
' Removes isolated characters/words after formula ends
'===============================================================
Private Function CleanFormulaTrailingNoise(rawFormula As String) As String
    CleanFormulaTrailingNoise = rawFormula
    
    ' Strategy: Find the end of valid formula by tracking parentheses and operators
    Dim validEnd As Long
    validEnd = FindFormulaEnd(rawFormula)
    
    If validEnd > 0 And validEnd < Len(rawFormula) Then
        ' Check if remaining text is just noise
        Dim remainder As String
        remainder = Trim(mid(rawFormula, validEnd + 1))
        
        If IsTrailingNoise(remainder) Then
            CleanFormulaTrailingNoise = Trim(Left(rawFormula, validEnd))
        End If
    End If
End Function

'===============================================================
' FUNCTION: Find the actual end of formula
' Returns position of last valid formula character
'===============================================================
Private Function FindFormulaEnd(formula As String) As Long
    FindFormulaEnd = Len(formula)
    
    Dim parenDepth As Long
    Dim lastValidPos As Long
    Dim i As Long
    
    parenDepth = 0
    lastValidPos = 0
    
    For i = 1 To Len(formula)
        Dim ch As String
        ch = mid(formula, i, 1)
        
        ' Track parentheses
        If ch = "(" Then
            parenDepth = parenDepth + 1
            lastValidPos = i
        ElseIf ch = ")" Then
            parenDepth = parenDepth - 1
            lastValidPos = i
        ' Valid formula characters
        ElseIf IsFormulaChar(ch) Then
            lastValidPos = i
        ' Whitespace is allowed within formula
        ElseIf ch = " " Then
            ' Don't update lastValidPos for space
        Else
            ' Invalid character - check if we should stop
            If parenDepth = 0 And lastValidPos > 0 Then
                ' Check if there are multiple spaces before this char
                Dim spacesBeforeNoise As Long
                spacesBeforeNoise = CountSpacesBefore(formula, i)
                
                ' If 2+ spaces or 5+ chars of whitespace, likely noise
                If spacesBeforeNoise >= 2 Or HasLargeGap(formula, lastValidPos, i) Then
                    FindFormulaEnd = lastValidPos
                    Exit Function
                End If
            End If
        End If
    Next i
    
    ' If parentheses are balanced and we found valid chars, use lastValidPos
    If parenDepth = 0 And lastValidPos > 0 Then
        FindFormulaEnd = lastValidPos
    End If
End Function

'===============================================================
' FUNCTION: Check if character is valid in formula
'===============================================================
Private Function IsFormulaChar(ch As String) As Boolean
    IsFormulaChar = False
    
    ' Numbers
    If ch >= "0" And ch <= "9" Then
        IsFormulaChar = True
        Exit Function
    End If
    
    ' Letters
    If (ch >= "A" And ch <= "Z") Or (ch >= "a" And ch <= "z") Then
        IsFormulaChar = True
        Exit Function
    End If
    
    ' Operators and special chars
    If ch = "+" Or ch = "-" Or ch = "*" Or ch = "/" Or ch = "." Or _
       ch = "(" Or ch = ")" Then
        IsFormulaChar = True
        Exit Function
    End If
End Function

'===============================================================
' FUNCTION: Count consecutive spaces before position
'===============================================================
Private Function CountSpacesBefore(text As String, pos As Long) As Long
    CountSpacesBefore = 0
    
    Dim i As Long
    For i = pos - 1 To 1 Step -1
        If mid(text, i, 1) = " " Then
            CountSpacesBefore = CountSpacesBefore + 1
        Else
            Exit For
        End If
    Next i
End Function

'===============================================================
' FUNCTION: Check if there's a large gap between positions
'===============================================================
Private Function HasLargeGap(text As String, StartPos As Long, EndPos As Long) As Boolean
    HasLargeGap = False
    
    If EndPos <= StartPos Then Exit Function
    
    Dim gapText As String
    gapText = mid(text, StartPos + 1, EndPos - StartPos - 1)
    
    ' Remove all spaces
    Dim nonSpaceCount As Long
    nonSpaceCount = 0
    
    Dim i As Long
    For i = 1 To Len(gapText)
        If mid(gapText, i, 1) <> " " Then
            nonSpaceCount = nonSpaceCount + 1
        End If
    Next i
    
    ' If gap is mostly spaces (95%+) and length > 5, it's a large gap
    If nonSpaceCount = 0 And Len(gapText) >= 5 Then
        HasLargeGap = True
    End If
End Function

'===============================================================
' FUNCTION: Check if text is trailing noise
' Returns True if text looks like noise (single chars, keywords, etc.)
'===============================================================
Private Function IsTrailingNoise(text As String) As Boolean
    IsTrailingNoise = False
    
    If Len(text) = 0 Then Exit Function
    
    ' Single character (like "S", "P", etc.)
    If Len(text) = 1 Then
        IsTrailingNoise = True
        Exit Function
    End If
    
    ' Very short text (2-3 chars) that's all letters
    If Len(text) <= 3 Then
        Dim allLetters As Boolean
        allLetters = True
        
        Dim i As Long
        For i = 1 To Len(text)
            Dim ch As String
            ch = mid(text, i, 1)
            If Not ((ch >= "A" And ch <= "Z") Or (ch >= "a" And ch <= "z")) Then
                allLetters = False
                Exit For
            End If
        Next i
        
        If allLetters Then
            IsTrailingNoise = True
            Exit Function
        End If
    End If
    
    ' Common noise keywords
    Dim upper As String
    upper = UCase(text)
    
    If upper = "S" Or upper = "P" Or upper = "N" Or upper = "Y" Or _
       upper = "YES" Or upper = "NO" Or upper = "OK" Or _
       upper = "TRUE" Or upper = "FALSE" Or _
       InStr(upper, "NOTES") > 0 Or InStr(upper, "REMARK") > 0 Then
        IsTrailingNoise = True
        Exit Function
    End If
    
    ' Text that doesn't contain formula operators or numbers
    If Not HasFormulaElements(text) Then
        IsTrailingNoise = True
    End If
End Function

'===============================================================
' FUNCTION: Check if text has formula elements
'===============================================================
Private Function HasFormulaElements(text As String) As Boolean
    HasFormulaElements = False
    
    ' Must have at least numbers or parentheses to be formula-like
    Dim i As Long
    For i = 1 To Len(text)
        Dim ch As String
        ch = mid(text, i, 1)
        
        If (ch >= "0" And ch <= "9") Or ch = "(" Or ch = ")" Or ch = "+" Or ch = "-" Then
            HasFormulaElements = True
            Exit Function
        End If
    Next i
End Function

'===============================================================
' ALTERNATIVE APPROACH: Aggressive trimming
' Use this if above methods are too conservative
'===============================================================
Private Function CleanFormulaAggressiveTrim(rawFormula As String) As String
    CleanFormulaAggressiveTrim = rawFormula
    
    ' Find last closing parenthesis or valid load case name
    Dim lastParen As Long
    lastParen = 0
    
    Dim i As Long
    For i = Len(rawFormula) To 1 Step -1
        If mid(rawFormula, i, 1) = ")" Then
            lastParen = i
            Exit For
        End If
    Next i
    
    ' If found closing paren, check if there's significant text after it
    If lastParen > 0 And lastParen < Len(rawFormula) Then
        Dim afterParen As String
        afterParen = Trim(mid(rawFormula, lastParen + 1))
        
        ' If text after closing paren is short and doesn't start with operator
        If Len(afterParen) > 0 And Len(afterParen) <= 10 Then
            Dim firstChar As String
            firstChar = Left(afterParen, 1)
            
            ' If not an operator, it's likely noise
            If firstChar <> "+" And firstChar <> "-" And firstChar <> "*" And firstChar <> "/" Then
                CleanFormulaAggressiveTrim = Trim(Left(rawFormula, lastParen))
            End If
        End If
    End If
End Function

Private Function FindNameFormulaSplit(lineText As String) As Long
    FindNameFormulaSplit = 0
    
    Dim i As Long
    For i = 2 To Len(lineText) - 1
        If mid(lineText, i, 1) = " " Then
            Dim j As Long
            For j = i + 1 To Len(lineText)
                Dim ch As String
                ch = mid(lineText, j, 1)
                
                If ch <> " " Then
                    If (ch >= "0" And ch <= "9") Or ch = "+" Or ch = "-" Or ch = "(" Or ch = "." Then
                        FindNameFormulaSplit = i + 1
                        Exit Function
                    Else
                        Exit For
                    End If
                End If
            Next j
        End If
    Next i
End Function

Private Function IsValidComboFormula(formula As String) As Boolean
    IsValidComboFormula = False
    
    Dim clean As String
    clean = Replace(formula, " ", "")
    
    If Len(clean) = 0 Then Exit Function
    
    Dim hasOperator As Boolean
    hasOperator = (InStr(clean, "+") > 0) Or (InStr(clean, "-") > 0)
    
    If Not hasLetters(clean) Then Exit Function
    
    If Not hasOperator Then
        If IsSingleLoadCase(clean) Then IsValidComboFormula = True
        Exit Function
    End If
    
    If hasOperator Then
        If HasScaleFactor(clean) Or InStr(clean, "(") > 0 Then
            IsValidComboFormula = True
        End If
    End If
End Function

Private Function HasScaleFactor(formula As String) As Boolean
    HasScaleFactor = False
    
    Dim i As Long
    For i = 1 To Len(formula) - 1
        Dim ch As String, nextCh As String
        ch = mid(formula, i, 1)
        nextCh = mid(formula, i + 1, 1)
        
        If (ch >= "0" And ch <= "9") And ((nextCh >= "A" And nextCh <= "Z") Or (nextCh >= "a" And nextCh <= "z")) Then
            HasScaleFactor = True
            Exit Function
        End If
        
        If ch = "." And i > 1 And i < Len(formula) Then
            Dim prevCh As String
            prevCh = mid(formula, i - 1, 1)
            If (prevCh >= "0" And prevCh <= "9") And (nextCh >= "0" And nextCh <= "9") Then
                HasScaleFactor = True
                Exit Function
            End If
        End If
    Next i
End Function

Private Function hasLetters(text As String) As Boolean
    hasLetters = False
    
    Dim i As Long
    For i = 1 To Len(text)
        Dim ch As String
        ch = mid(text, i, 1)
        
        If (ch >= "A" And ch <= "Z") Or (ch >= "a" And ch <= "z") Then
            hasLetters = True
            Exit Function
        End If
    Next i
End Function

Private Function IsSingleLoadCase(text As String) As Boolean
    IsSingleLoadCase = False
    
    If Len(text) = 0 Or Len(text) > 20 Then Exit Function
    
    Dim i As Long
    For i = 1 To Len(text)
        Dim ch As String
        ch = mid(text, i, 1)
        
        If Not ((ch >= "0" And ch <= "9") Or ch = ".") Then Exit For
    Next i
    
    If i <= Len(text) Then
        Dim letterPart As String
        letterPart = mid(text, i)
        
        Dim j As Long
        For j = 1 To Len(letterPart)
            ch = mid(letterPart, j, 1)
            If Not ((ch >= "A" And ch <= "Z") Or (ch >= "a" And ch <= "z")) Then Exit Function
        Next j
        
        If Len(letterPart) >= 2 And Len(letterPart) <= 10 Then IsSingleLoadCase = True
    End If
End Function

Private Function ExtractPattern(Name As String, ByRef prefix As String, ByRef number As Long) As Boolean
    ExtractPattern = False
    prefix = ""
    number = 0
    
    If Len(Name) = 0 Then Exit Function
    
    Dim i As Long
    For i = Len(Name) To 1 Step -1
        Dim ch As String
        ch = mid(Name, i, 1)
        
        If Not (ch >= "0" And ch <= "9") Then Exit For
    Next i
    
    If i >= Len(Name) Then Exit Function
    
    prefix = Left(Name, i)
    Dim numStr As String
    numStr = mid(Name, i + 1)
    
    If Len(numStr) > 0 Then
        On Error Resume Next
        number = CLng(numStr)
        If err.number = 0 And number > 0 Then ExtractPattern = True
        err.Clear
        On Error GoTo 0
    End If
End Function

Private Function IsNumericSection(lineText As String) As Boolean
    IsNumericSection = False
    
    lineText = Trim(lineText)
    If Len(lineText) > 15 Then Exit Function
    
    If Right(lineText, 1) = "." Or Right(lineText, 1) = ")" Then
        IsNumericSection = True
        Exit Function
    End If
    
    Dim dotCount As Long, digitCount As Long, i As Long
    
    For i = 1 To Len(lineText)
        Dim ch As String
        ch = mid(lineText, i, 1)
        If ch = "." Then dotCount = dotCount + 1
        If ch >= "0" And ch <= "9" Then digitCount = digitCount + 1
    Next i
    
    If dotCount > 0 And dotCount <= 3 And digitCount >= dotCount Then IsNumericSection = True
End Function

Private Function IsDocumentNumber(text As String) As Boolean
    IsDocumentNumber = False
    
    Dim dashCount As Long, digitCount As Long, i As Long
    
    For i = 1 To Len(text)
        Dim ch As String
        ch = mid(text, i, 1)
        
        If ch = "-" Then dashCount = dashCount + 1
        If ch >= "0" And ch <= "9" Then digitCount = digitCount + 1
    Next i
    
    If dashCount >= 2 And digitCount >= 4 Then IsDocumentNumber = True
End Function

Public Function ParseComboTableFormat(inputText As String) As Object
    Dim combos As Object
    Set combos = CreateObject("Scripting.Dictionary")

    inputText = Replace(inputText, vbCrLf, vbLf)
    inputText = Replace(inputText, vbCr, vbLf)

    Dim lines As Variant
    lines = Split(inputText, vbLf)

    Dim i As Long, inTableSection As Boolean
    inTableSection = False

    For i = LBound(lines) To UBound(lines)
        Dim lineText As String
        lineText = CStr(lines(i))

        If Len(Trim(lineText)) = 0 Then GoTo NextLine
        If InStr(lineText, "---") > 0 Or InStr(lineText, "===") > 0 Then GoTo NextLine

        If InStr(1, lineText, "Combo", vbTextCompare) > 0 And _
           InStr(1, lineText, "Type", vbTextCompare) > 0 And _
           InStr(1, lineText, "Formula", vbTextCompare) > 0 Then
            inTableSection = True
            GoTo NextLine
        End If

        If InStr(1, lineText, "You can edit", vbTextCompare) > 0 Or _
           InStr(1, lineText, "Keep the table", vbTextCompare) > 0 Or _
           InStr(1, lineText, "Click", vbTextCompare) > 0 Or _
           InStr(1, lineText, "PREVIEW", vbTextCompare) > 0 Or _
           InStr(1, lineText, "Options", vbTextCompare) > 0 Then
            inTableSection = False
            GoTo NextLine
        End If

        If Not inTableSection Then GoTo NextLine

        If Len(lineText) >= 35 Then
            Dim comboName As String, typeStr As String, formulaStr As String

            comboName = Trim(Left(lineText, 20))
            typeStr = Trim(mid(lineText, 22, 20))
            formulaStr = Trim(mid(lineText, 43))

            If Len(comboName) > 0 And Len(formulaStr) > 0 Then
                Dim comboType As Long
                comboType = GetComboTypeFromString(typeStr)

                If Not combos.exists(comboName) Then
                    Dim comboInfo(0 To 2) As Variant
                    comboInfo(0) = comboName
                    comboInfo(1) = formulaStr
                    comboInfo(2) = comboType
                    combos.Add comboName, comboInfo
                End If
            End If
        End If
NextLine:
    Next i

    Set ParseComboTableFormat = combos
End Function

Private Function AddCasesToCombo(comboName As String, formula As String, ByRef errorLog As String) As Boolean
    AddCasesToCombo = False

    formula = Replace(formula, " ", "")
    formula = Replace(formula, vbTab, "")
    formula = Replace(formula, vbCr, "")
    formula = Replace(formula, vbLf, "")

    Dim terms As Collection
    Set terms = ParseFormula(formula)

    If terms Is Nothing Or terms.count = 0 Then
        errorLog = errorLog & "WARNING: No terms parsed from formula: " & formula & vbCrLf
        Exit Function
    End If

    Dim anyFail As Boolean
    anyFail = False

    Dim i As Long
    For i = 1 To terms.count
        Dim termInfo As Variant
        termInfo = terms(i)

        Dim scaleFactor As Double, caseName As String
        scaleFactor = CDbl(termInfo(0))
        caseName = CStr(termInfo(1))

        Dim ret As Long
        ret = SapModel.RespCombo.SetCaseList(comboName, 0, caseName, scaleFactor)

        If ret <> 0 Then
            ret = SapModel.RespCombo.SetCaseList(comboName, 1, caseName, scaleFactor)
            If ret <> 0 Then
                errorLog = errorLog & "WARNING: Failed to add '" & caseName & "' to '" & comboName & "'" & vbCrLf
                anyFail = True
            End If
        End If
    Next i

    AddCasesToCombo = Not anyFail
End Function

Private Function ParseFormula(formula As String) As Collection
    Dim terms As Collection
    Set terms = New Collection

    Dim currentPos As Long
    currentPos = 1

    Dim currentSign As Double
    currentSign = 1#

    Do While currentPos <= Len(formula)
        Dim ch As String
        ch = mid(formula, currentPos, 1)

        If ch = "+" Then
            currentSign = 1#
            currentPos = currentPos + 1
        ElseIf ch = "-" Then
            currentSign = -1#
            currentPos = currentPos + 1
        ElseIf ch = "(" Then
            Dim closePos As Long
            closePos = FindMatchingParen(formula, currentPos)

            If closePos > currentPos Then
                Dim preScale As Double
                preScale = 1#

                If currentPos > 1 Then
                    Dim scaleStr As String
                    scaleStr = ExtractNumberBefore(formula, currentPos - 1)
                    If Len(scaleStr) > 0 Then
                        On Error Resume Next
                        preScale = CDbl(scaleStr)
                        If err.number <> 0 Then preScale = 1#
                        err.Clear
                        On Error GoTo 0
                    End If
                End If

                preScale = preScale * currentSign

                Dim innerFormula As String
                innerFormula = mid(formula, currentPos + 1, closePos - currentPos - 1)

                Dim innerTerms As Collection
                Set innerTerms = ParseFormula(innerFormula)

                Dim j As Long
                For j = 1 To innerTerms.count
                    Dim innerTerm As Variant
                    innerTerm = innerTerms(j)

                    Dim termData(0 To 1) As Variant
                    termData(0) = CDbl(innerTerm(0)) * preScale
                    termData(1) = CStr(innerTerm(1))

                    terms.Add termData
                Next j

                currentPos = closePos + 1
                currentSign = 1#
            Else
                currentPos = currentPos + 1
            End If
        Else
            Dim termEnd As Long
            termEnd = FindTermEnd(formula, currentPos)

            If termEnd > currentPos Then
                Dim termStr As String
                termStr = mid(formula, currentPos, termEnd - currentPos)

                Dim sF As Double, cn As String

                If ExtractScaleAndName(termStr, sF, cn) Then
                    Dim tD(0 To 1) As Variant
                    tD(0) = sF * currentSign
                    tD(1) = cn

                    terms.Add tD
                End If

                currentPos = termEnd
                currentSign = 1#
            Else
                currentPos = currentPos + 1
            End If
        End If
    Loop

    Set ParseFormula = terms
End Function

Private Function FindMatchingParen(formula As String, openPos As Long) As Long
    FindMatchingParen = 0
    Dim Depth As Long
    Depth = 1

    Dim i As Long
    For i = openPos + 1 To Len(formula)
        Dim ch As String
        ch = mid(formula, i, 1)

        If ch = "(" Then
            Depth = Depth + 1
        ElseIf ch = ")" Then
            Depth = Depth - 1
            If Depth = 0 Then
                FindMatchingParen = i
                Exit Function
            End If
        End If
    Next i
End Function

Private Function ExtractNumberBefore(formula As String, EndPos As Long) As String
    ExtractNumberBefore = ""

    Dim i As Long
    For i = EndPos To 1 Step -1
        Dim ch As String
        ch = mid(formula, i, 1)

        If (ch >= "0" And ch <= "9") Or ch = "." Then
            ExtractNumberBefore = ch & ExtractNumberBefore
        Else
            Exit For
        End If
    Next i
End Function

Private Function FindTermEnd(formula As String, StartPos As Long) As Long
    Dim i As Long
    For i = StartPos To Len(formula)
        Dim ch As String
        ch = mid(formula, i, 1)

        If ch = "+" Or ch = "-" Or ch = "(" Or ch = ")" Then
            FindTermEnd = i
            Exit Function
        End If
    Next i

    FindTermEnd = Len(formula) + 1
End Function

Private Function ExtractScaleAndName(term As String, ByRef scaleFactor As Double, ByRef caseName As String) As Boolean
    ExtractScaleAndName = False
    scaleFactor = 1#
    caseName = ""

    term = Trim(term)
    If Len(term) = 0 Then Exit Function

    Dim i As Long
    For i = 1 To Len(term)
        Dim ch As String
        ch = mid(term, i, 1)

        If Not ((ch >= "0" And ch <= "9") Or ch = ".") Then
            Exit For
        End If
    Next i

    If i > 1 Then
        Dim numStr As String
        numStr = Left(term, i - 1)

        On Error Resume Next
        scaleFactor = CDbl(numStr)
        If err.number <> 0 Then scaleFactor = 1#
        err.Clear
        On Error GoTo 0

        caseName = mid(term, i)
    Else
        scaleFactor = 1#
        caseName = term
    End If

    If Len(caseName) > 0 Then
        ExtractScaleAndName = True
    End If
End Function

Private Function GetComboTypeFromString(typeStr As String) As Long
    Select Case LCase(Trim(typeStr))
        Case "linear", "linear additive", "additive", "linear add"
            GetComboTypeFromString = 0
        Case "envelope"
            GetComboTypeFromString = 1
        Case "absolute", "absolute additive", "absolute add"
            GetComboTypeFromString = 2
        Case "srss"
            GetComboTypeFromString = 3
        Case "range", "range additive", "range add"
            GetComboTypeFromString = 4
        Case Else
            GetComboTypeFromString = 0
    End Select
End Function

Public Function GetComboTypeString(comboType As Long) As String
    Select Case comboType
        Case 0: GetComboTypeString = "Linear Additive"
        Case 1: GetComboTypeString = "Envelope"
        Case 2: GetComboTypeString = "Absolute Additive"
        Case 3: GetComboTypeString = "SRSS"
        Case 4: GetComboTypeString = "Range Additive"
        Case Else: GetComboTypeString = "Linear Additive"
    End Select
End Function

'===============================================================
' PHASE 7: POST-FILTER - Clean all formulas
' This runs AFTER scoring and filtering, as final safety check
'===============================================================
Private Sub PostFilterCleanFormulas(combos As Object)
    On Error Resume Next
    
    If combos Is Nothing Then Exit Sub
    If combos.count = 0 Then Exit Sub
    
    Dim comboKey As Variant
    Dim keysToUpdate As Collection
    Set keysToUpdate = New Collection
    
    ' First pass: identify formulas that need cleaning
    For Each comboKey In combos.keys
        Dim cInfo As Variant
        cInfo = combos(comboKey)
        
        Dim originalFormula As String
        originalFormula = CStr(cInfo(1))
        
        Dim cleanedFormula As String
        cleanedFormula = PostCleanFormula(originalFormula)
        
        ' If formula changed, mark for update
        If cleanedFormula <> originalFormula Then
            Dim updateInfo(0 To 2) As Variant
            updateInfo(0) = comboKey
            updateInfo(1) = cleanedFormula
            updateInfo(2) = cInfo(2) ' Preserve combo type
            
            keysToUpdate.Add updateInfo
        End If
    Next comboKey
    
    ' Second pass: update the dictionary
    Dim i As Long
    For i = 1 To keysToUpdate.count
        Dim updateData As Variant
        updateData = keysToUpdate(i)
        
        Dim keyToUpdate As Variant
        keyToUpdate = updateData(0)
        
        ' Update the combo info
        Dim newInfo(0 To 2) As Variant
        newInfo(0) = CStr(keyToUpdate)
        newInfo(1) = CStr(updateData(1)) ' Cleaned formula
        newInfo(2) = updateData(2) ' Combo type
        
        combos(keyToUpdate) = newInfo
    Next i
End Sub

'===============================================================
' POST-CLEAN FORMULA: Remove trailing/leading noise
' Uses multiple detection strategies
'===============================================================
Private Function PostCleanFormula(formula As String) As String
    PostCleanFormula = formula
    
    If Len(formula) = 0 Then Exit Function
    
    Dim workingFormula As String
    workingFormula = formula
    
    ' Strategy 1: Remove text after large whitespace gap
    workingFormula = RemoveAfterLargeGap(workingFormula)
    
    ' Strategy 2: Remove isolated trailing terms that don't belong
    workingFormula = RemoveIsolatedTrailingTerms(workingFormula)
    
    ' Strategy 3: Remove leading noise (rare but possible)
    workingFormula = RemoveLeadingNoise(workingFormula)
    
    ' Strategy 4: Validate parentheses balance
    workingFormula = FixUnbalancedParentheses(workingFormula)
    
    ' Final validation
    If IsValidFormulaStructure(workingFormula) Then
        PostCleanFormula = Trim(workingFormula)
    End If
End Function

'===============================================================
' STRATEGY 1: Remove text after large whitespace gap
'===============================================================
Private Function RemoveAfterLargeGap(formula As String) As String
    RemoveAfterLargeGap = formula
    
    ' Find sequences of 3+ consecutive spaces (or tabs converted to spaces)
    Dim pos As Long
    pos = InStr(formula, "   ") ' 3 spaces
    
    If pos > 0 Then
        ' Check if what's after the gap looks like noise
        Dim afterGap As String
        afterGap = Trim(mid(formula, pos))
        
        If IsDefinitelyNoise(afterGap) Then
            RemoveAfterLargeGap = Trim(Left(formula, pos - 1))
        End If
    End If
End Function

'===============================================================
' STRATEGY 2: Remove isolated trailing terms
' Checks if last term in formula is suspicious
'===============================================================
Private Function RemoveIsolatedTrailingTerms(formula As String) As String
    RemoveIsolatedTrailingTerms = formula
    
    ' Parse formula into terms
    Dim terms As Collection
    Set terms = SplitFormulaIntoTerms(formula)
    
    If terms.count = 0 Then Exit Function
    
    ' Check last term
    Dim lastTerm As String
    lastTerm = Trim(CStr(terms(terms.count)))
    
    ' If last term is suspicious single char/word
    If IsSuspiciousTrailingTerm(lastTerm, formula) Then
        ' Rebuild formula without last term
        Dim rebuilt As String
        rebuilt = ""
        
        Dim i As Long
        For i = 1 To terms.count - 1
            If i > 1 Then rebuilt = rebuilt & " "
            rebuilt = rebuilt & CStr(terms(i))
        Next i
        
        If Len(rebuilt) > 0 Then
            RemoveIsolatedTrailingTerms = rebuilt
        End If
    End If
End Function

'===============================================================
' STRATEGY 3: Remove leading noise
'===============================================================
Private Function RemoveLeadingNoise(formula As String) As String
    RemoveLeadingNoise = formula
    
    formula = Trim(formula)
    If Len(formula) = 0 Then Exit Function
    
    ' Check if formula starts with isolated single letter/number
    Dim firstChar As String
    firstChar = Left(formula, 1)
    
    ' If starts with letter that's not followed by valid formula chars
    If (firstChar >= "A" And firstChar <= "Z") Or (firstChar >= "a" And firstChar <= "z") Then
        If Len(formula) >= 3 Then
            Dim secondChar As String
            secondChar = mid(formula, 2, 1)
            
            ' If followed by space and then number/operator, remove first char
            If secondChar = " " Then
                Dim thirdChar As String
                thirdChar = mid(formula, 3, 1)
                
                If (thirdChar >= "0" And thirdChar <= "9") Or thirdChar = "(" Or thirdChar = "-" Then
                    RemoveLeadingNoise = Trim(mid(formula, 3))
                End If
            End If
        End If
    End If
End Function

'===============================================================
' STRATEGY 4: Fix unbalanced parentheses
'===============================================================
Private Function FixUnbalancedParentheses(formula As String) As String
    FixUnbalancedParentheses = formula
    
    Dim openCount As Long, closeCount As Long
    Dim i As Long
    
    For i = 1 To Len(formula)
        If mid(formula, i, 1) = "(" Then openCount = openCount + 1
        If mid(formula, i, 1) = ")" Then closeCount = closeCount + 1
    Next i
    
    ' If more closing than opening, likely has noise before formula
    If closeCount > openCount Then
        ' Try to find first opening paren and start from there
        Dim firstOpen As Long
        firstOpen = InStr(formula, "(")
        
        If firstOpen > 1 Then
            Dim beforeParen As String
            beforeParen = Trim(Left(formula, firstOpen - 1))
            
            If IsDefinitelyNoise(beforeParen) Then
                FixUnbalancedParentheses = Trim(mid(formula, firstOpen))
            End If
        End If
    End If
    
    ' If more opening than closing, likely has noise after formula
    If openCount > closeCount Then
        ' Find last valid closing paren position
        Dim lastClose As Long
        lastClose = 0
        
        For i = Len(formula) To 1 Step -1
            If mid(formula, i, 1) = ")" Then
                ' Check if this balances the parentheses
                Dim testFormula As String
                testFormula = Left(formula, i)
                
                If AreParensBalanced(testFormula) Then
                    lastClose = i
                    Exit For
                End If
            End If
        Next i
        
        If lastClose > 0 And lastClose < Len(formula) Then
            Dim afterClose As String
            afterClose = Trim(mid(formula, lastClose + 1))
            
            If IsDefinitelyNoise(afterClose) Then
                FixUnbalancedParentheses = Trim(Left(formula, lastClose))
            End If
        End If
    End If
End Function

'===============================================================
' HELPER: Check if parentheses are balanced
'===============================================================
Private Function AreParensBalanced(formula As String) As Boolean
    Dim Depth As Long
    Depth = 0
    
    Dim i As Long
    For i = 1 To Len(formula)
        If mid(formula, i, 1) = "(" Then Depth = Depth + 1
        If mid(formula, i, 1) = ")" Then Depth = Depth - 1
        
        If Depth < 0 Then
            AreParensBalanced = False
            Exit Function
        End If
    Next i
    
    AreParensBalanced = (Depth = 0)
End Function

'===============================================================
' HELPER: Split formula into terms (by spaces, considering parens)
'===============================================================
Private Function SplitFormulaIntoTerms(formula As String) As Collection
    Dim terms As Collection
    Set terms = New Collection
    
    Dim currentTerm As String
    Dim parenDepth As Long
    Dim i As Long
    
    currentTerm = ""
    parenDepth = 0
    
    For i = 1 To Len(formula)
        Dim ch As String
        ch = mid(formula, i, 1)
        
        If ch = "(" Then parenDepth = parenDepth + 1
        If ch = ")" Then parenDepth = parenDepth - 1
        
        If ch = " " And parenDepth = 0 Then
            ' Space outside parentheses - term boundary
            If Len(currentTerm) > 0 Then
                terms.Add currentTerm
                currentTerm = ""
            End If
        Else
            currentTerm = currentTerm & ch
        End If
    Next i
    
    ' Add last term
    If Len(currentTerm) > 0 Then
        terms.Add currentTerm
    End If
    
    Set SplitFormulaIntoTerms = terms
End Function

'===============================================================
' HELPER: Check if text is definitely noise
'===============================================================
Private Function IsDefinitelyNoise(text As String) As Boolean
    IsDefinitelyNoise = False
    
    text = Trim(text)
    If Len(text) = 0 Then
        IsDefinitelyNoise = True
        Exit Function
    End If
    
    ' Single character that's just a letter
    If Len(text) = 1 Then
        Dim ch As String
        ch = text
        If (ch >= "A" And ch <= "Z") Or (ch >= "a" And ch <= "z") Then
            IsDefinitelyNoise = True
            Exit Function
        End If
    End If
    
    ' Short text (2-5 chars) with no numbers, operators, or parentheses
    If Len(text) <= 5 Then
        If Not HasFormulaCharacters(text) Then
            IsDefinitelyNoise = True
            Exit Function
        End If
    End If
    
    ' Common noise keywords
    Dim upper As String
    upper = UCase(text)
    
    If upper = "S" Or upper = "P" Or upper = "N" Or upper = "Y" Or _
       upper = "YES" Or upper = "NO" Or upper = "OK" Or upper = "TRUE" Or _
       upper = "FALSE" Or upper = "NOTES" Or upper = "NOTE" Or _
       upper = "REMARK" Or upper = "DESC" Or upper = "DESCRIPTION" Then
        IsDefinitelyNoise = True
    End If
End Function

'===============================================================
' HELPER: Check if term is suspicious trailing term
'===============================================================
Private Function IsSuspiciousTrailingTerm(term As String, fullFormula As String) As Boolean
    IsSuspiciousTrailingTerm = False
    
    ' If term is definitely noise
    If IsDefinitelyNoise(term) Then
        IsSuspiciousTrailingTerm = True
        Exit Function
    End If
    
    ' Check if term is isolated (has large space before it)
    Dim termPos As Long
    termPos = InStr(fullFormula, term)
    
    If termPos > 3 Then
        ' Check space before term
        Dim beforeTerm As String
        beforeTerm = mid(fullFormula, termPos - 3, 3)
        
        ' If 3+ spaces before this term, it's isolated
        If beforeTerm = "   " Or InStr(beforeTerm, "  ") > 0 Then
            IsSuspiciousTrailingTerm = True
            Exit Function
        End If
    End If
    
    ' Check if term doesn't start with number or operator (suspicious for formula)
    Dim firstChar As String
    firstChar = Left(term, 1)
    
    If Not ((firstChar >= "0" And firstChar <= "9") Or firstChar = "+" Or _
            firstChar = "-" Or firstChar = "(" Or firstChar = ".") Then
        ' But has no operators inside either
        If Not HasFormulaOperators(term) Then
            IsSuspiciousTrailingTerm = True
        End If
    End If
End Function

'===============================================================
' HELPER: Check if text has formula characters
'===============================================================
Private Function HasFormulaCharacters(text As String) As Boolean
    HasFormulaCharacters = False
    
    Dim i As Long
    For i = 1 To Len(text)
        Dim ch As String
        ch = mid(text, i, 1)
        
        If (ch >= "0" And ch <= "9") Or ch = "(" Or ch = ")" Or _
           ch = "+" Or ch = "-" Or ch = "." Then
            HasFormulaCharacters = True
            Exit Function
        End If
    Next i
End Function

'===============================================================
' HELPER: Check if text has formula operators
'===============================================================
Private Function HasFormulaOperators(text As String) As Boolean
    HasFormulaOperators = False
    
    If InStr(text, "+") > 0 Or InStr(text, "-") > 0 Or _
       InStr(text, "*") > 0 Or InStr(text, "/") > 0 Then
        HasFormulaOperators = True
    End If
End Function

'===============================================================
' HELPER: Validate formula structure
'===============================================================
Private Function IsValidFormulaStructure(formula As String) As Boolean
    IsValidFormulaStructure = False
    
    If Len(formula) = 0 Then Exit Function
    
    ' Must have balanced parentheses
    If Not AreParensBalanced(formula) Then Exit Function
    
    ' Must contain at least some formula elements
    If Not HasFormulaCharacters(formula) Then Exit Function
    
    ' Must have letters (load case names)
    Dim hasLetters As Boolean
    hasLetters = False
    
    Dim i As Long
    For i = 1 To Len(formula)
        Dim ch As String
        ch = mid(formula, i, 1)
        
        If (ch >= "A" And ch <= "Z") Or (ch >= "a" And ch <= "z") Then
            hasLetters = True
            Exit For
        End If
    Next i
    
    If Not hasLetters Then Exit Function
    
    IsValidFormulaStructure = True
End Function


