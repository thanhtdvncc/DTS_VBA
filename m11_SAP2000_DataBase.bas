Attribute VB_Name = "m11_SAP2000_DataBase"
Option Explicit
'===============================================================
' Module: modSAP2000_DatabaseTables
' Purpose: Export/Import SAP2000 Database Tables via Excel
' User-friendly interface with simple function names
'===============================================================

' Module-level variables (single source of truth for globals)
Public g_ExportedSheetName As String
Public g_ExportedWorkbookName As String
Public g_SelectedTableKey As String
Public g_FieldKeysIncluded() As String
Public g_TableVersion As Long
Public g_NumberRecords As Long

'===============================================================
' MAIN USER FUNCTIONS
'===============================================================

' User-friendly function to get database table for editing
Public Sub GetDatabaseToEdit()
    ' Show selection form
    On Error Resume Next
    frmTableSelector.Show
End Sub

' User-friendly function to update changes back to SAP2000
' Optional force: if True, will attempt to deduce table from active sheet even if globals empty
Public Sub UpdateToSAP2000(Optional ByVal force As Boolean = False)
    Call ImportEditedTableFromActiveSheet(force)
End Sub

'===============================================================
' NEW: Export only Grid Lines to sheet "Girdline"
' - Uses SAP2000 DatabaseTables API to get metadata + TableData
' - Writes Title (row1), Field Keys (row2), Field Names (row3), Units (row4)
' - Writes data starting at row5
' - Deletes existing sheet "Girdline" if present, then creates new one
' - Only clears the exact target range (title..data) on new sheet
'===============================================================
Public Sub ExportGridLinesToGirdlineSheet()
    On Error GoTo ErrHandler

    ' Connect to SAP2000
    If Not ConnectSAP2000() Then
        MsgBox "Could not connect to SAP2000.", vbCritical, "Connection Error"
        Exit Sub
    End If

    Dim tableKey    As String
    tableKey = Trim(g_SelectedTableKey)
    If Len(tableKey) = 0 Then tableKey = "Grid Lines"

    Dim ret         As Long
    ' Get metadata for headers
    Dim metaTableVersion As Long
    Dim metaNumberFields As Long
    Dim FieldKeyArr() As String
    Dim FieldNameArr() As String
    Dim DescriptionArr() As String
    Dim UnitsStringArr() As String
    Dim IsImportableArr() As Boolean

    ret = SapModel.DatabaseTables.GetAllFieldsInTable( _
            tableKey, metaTableVersion, metaNumberFields, _
            FieldKeyArr, FieldNameArr, DescriptionArr, _
            UnitsStringArr, IsImportableArr)
    ' If metadata fails, we will still try to get data (names/units may be empty).

    ' Prepare for data request
    Dim FieldKeyListInput() As String
    ReDim FieldKeyListInput(0 To 0) As String
    FieldKeyListInput(0) = ""

    Dim FieldKeysIncludedLocal() As String
    Dim NumberRecords As Long
    Dim tableData() As String
    Dim TableVersion As Long

    ret = SapModel.DatabaseTables.GetTableForDisplayArray( _
            tableKey, FieldKeyListInput, "All", _
            TableVersion, FieldKeysIncludedLocal, NumberRecords, tableData)

    If ret <> 0 Then
        MsgBox "Failed to get table data." & vbCrLf & "Return code: " & ret, vbCritical, "Error"
        Exit Sub
    End If

    If NumberRecords <= 0 Or Not IsArray(tableData) Then
        ' No data -> nothing to write
        Exit Sub
    End If

    ' Determine dimensions
    Dim nCols As Long, nRows As Long
    nCols = UBound(FieldKeysIncludedLocal) - LBound(FieldKeysIncludedLocal) + 1
    nRows = NumberRecords

    ' Prepare Excel
    Dim xlApp As Object, wb As Object, ws As Object
    On Error Resume Next
    Set xlApp = GetObject(, "Excel.Application")
    If xlApp Is Nothing Then Set xlApp = CreateObject("Excel.Application")
    On Error GoTo ErrHandler

    xlApp.Visible = True
    If xlApp.Workbooks.count = 0 Then
        xlApp.Workbooks.Add
    End If

    Set wb = xlApp.ActiveWorkbook
    If wb Is Nothing Then
        MsgBox "Please open an Excel workbook first.", vbExclamation, "No Workbook"
        Exit Sub
    End If

    ' Ensure sheet "Girdline" exists (create if missing) and clear all contents/formats
    Dim shtName     As String: shtName = "Girdline"
    On Error Resume Next
    Set ws = wb.Worksheets(shtName)
    On Error GoTo ErrHandler

    If ws Is Nothing Then
        ' create sheet at end
        On Error Resume Next
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count))
        On Error GoTo ErrHandler
        If Not ws Is Nothing Then
            On Error Resume Next
            ws.Name = shtName
            If err.number <> 0 Then
                err.Clear
                ' If renaming failed (rare), keep default name (ws points to created sheet)
            End If
            On Error GoTo ErrHandler
        Else
            MsgBox "Failed to create sheet for Girdline.", vbExclamation, "Sheet Error"
            Exit Sub
        End If
    End If

    ' Clear entire sheet (values, formulas, formats)
    On Error Resume Next
    ws.Cells.Clear
    ws.Cells.ClearFormats
    On Error GoTo ErrHandler

    ' Define write rows/cols
    Dim titleRow    As Long: titleRow = 1
    Dim keysRow     As Long: keysRow = 2
    Dim namesRow    As Long: namesRow = 3
    Dim unitsRow    As Long: unitsRow = 4
    Dim dataStartRow As Long: dataStartRow = 5
    Dim leftCol     As Long: leftCol = 1    ' column A
    Dim rightCol    As Long: rightCol = leftCol + nCols - 1
    Dim bottomRow   As Long: bottomRow = dataStartRow + nRows - 1
    If bottomRow < dataStartRow Then bottomRow = dataStartRow

    ' Clear only the exact target range on the fresh sheet (should be empty, but keep for safety)
    On Error Resume Next
    ws.Range(ws.Cells(titleRow, leftCol), ws.Cells(bottomRow, rightCol)).Clear
    On Error GoTo ErrHandler

    ' ROW 1: Title (merged)
    ws.Cells(titleRow, leftCol).Value = "SAP2000 Database: " & tableKey
    If rightCol >= leftCol Then
        On Error Resume Next
        ws.Range(ws.Cells(titleRow, leftCol), ws.Cells(titleRow, rightCol)).Merge
        On Error GoTo ErrHandler
    End If
    ws.Cells(titleRow, leftCol).Font.Bold = True
    ws.Cells(titleRow, leftCol).Font.Size = 12
    ws.Cells(titleRow, leftCol).HorizontalAlignment = -4108
    On Error Resume Next
    ws.Cells(titleRow, leftCol).Interior.color = RGB(68, 114, 196)
    ws.Cells(titleRow, leftCol).Font.color = RGB(255, 255, 255)
    On Error GoTo ErrHandler
    ws.rows(titleRow).rowHeight = 25

    ' ROW 2: Field Keys
    Dim c As Long, r As Long
    For c = 0 To nCols - 1
        Dim fldIdx  As Long: fldIdx = LBound(FieldKeysIncludedLocal) + c
        If fldIdx >= LBound(FieldKeysIncludedLocal) And fldIdx <= UBound(FieldKeysIncludedLocal) Then
            ws.Cells(keysRow, leftCol + c).Value = FieldKeysIncludedLocal(fldIdx)
            ws.Cells(keysRow, leftCol + c).Font.Bold = True
            On Error Resume Next
            ws.Cells(keysRow, leftCol + c).Interior.color = RGB(217, 217, 217)
            On Error GoTo ErrHandler
        End If
    Next c

    ' ROW 3: Field Names (match to FieldKeyArr)
    For c = 0 To nCols - 1
        Dim fk      As String: fk = ""
        fldIdx = LBound(FieldKeysIncludedLocal) + c
        If fldIdx >= LBound(FieldKeysIncludedLocal) And fldIdx <= UBound(FieldKeysIncludedLocal) Then
            fk = FieldKeysIncludedLocal(fldIdx)
        End If

        Dim foundIdx As Long: foundIdx = -1
        If IsArray(FieldKeyArr) Then
            For r = LBound(FieldKeyArr) To UBound(FieldKeyArr)
                If FieldKeyArr(r) = fk Then
                    foundIdx = r
                    Exit For
                End If
            Next r
        End If

        If foundIdx >= 0 Then
            If IsArray(FieldNameArr) Then
                ws.Cells(namesRow, leftCol + c).Value = FieldNameArr(foundIdx)
                ws.Cells(namesRow, leftCol + c).Font.Italic = True
            End If
        End If
    Next c

    ' ROW 4: Units
    For c = 0 To nCols - 1
        fk = ""
        fldIdx = LBound(FieldKeysIncludedLocal) + c
        If fldIdx >= LBound(FieldKeysIncludedLocal) And fldIdx <= UBound(FieldKeysIncludedLocal) Then
            fk = FieldKeysIncludedLocal(fldIdx)
        End If

        foundIdx = -1
        If IsArray(FieldKeyArr) Then
            For r = LBound(FieldKeyArr) To UBound(FieldKeyArr)
                If FieldKeyArr(r) = fk Then
                    foundIdx = r
                    Exit For
                End If
            Next r
        End If

        If foundIdx >= 0 Then
            If IsArray(UnitsStringArr) Then
                ws.Cells(unitsRow, leftCol + c).Value = UnitsStringArr(foundIdx)
                ws.Cells(unitsRow, leftCol + c).Font.color = RGB(128, 128, 128)
            End If
        End If
    Next c

    ' ROW 5+: Data (TableData in row-major order)
    Dim baseIdx     As Long
    For r = 0 To nRows - 1
        For c = 0 To nCols - 1
            baseIdx = r * nCols + c
            If baseIdx >= LBound(tableData) And baseIdx <= UBound(tableData) Then
                ws.Cells(dataStartRow + r, leftCol + c).Value = tableData(baseIdx)
            End If
        Next c
    Next r

    ' Light formatting: bold header rows and autofit columns
    On Error Resume Next
    ws.rows(keysRow & ":" & unitsRow).Font.Bold = True
    ws.Columns("A:AZ").AutoFit
    On Error GoTo ErrHandler

    ' Add borders around header+data
    If nRows > 0 Then
        Dim dataRange As Object
        Set dataRange = ws.Range(ws.Cells(keysRow, leftCol), ws.Cells(bottomRow, rightCol))
        On Error Resume Next
        dataRange.Borders.LineStyle = 1
        dataRange.Borders.Weight = 2
        On Error GoTo ErrHandler
    End If

    ' Update module globals for consistency
    g_SelectedTableKey = tableKey
    ReDim g_FieldKeysIncluded(LBound(FieldKeysIncludedLocal) To UBound(FieldKeysIncludedLocal))
    Dim ii          As Long
    For ii = LBound(FieldKeysIncludedLocal) To UBound(FieldKeysIncludedLocal)
        g_FieldKeysIncluded(ii) = FieldKeysIncludedLocal(ii)
    Next ii
    g_TableVersion = TableVersion
    g_NumberRecords = NumberRecords
    g_ExportedSheetName = ws.Name
    g_ExportedWorkbookName = wb.Name

    Exit Sub

ErrHandler:
    MsgBox "ExportGridLinesToGirdlineSheet failed: " & err.description, vbCritical, "Error"
End Sub

'===============================================================
' INTERNAL: Export table to ActiveSheet
'===============================================================
Public Sub ExportTableToActiveSheet(tableKey As String)
    On Error GoTo ErrHandler

    ' Connect to SAP2000
    If Not ConnectSAP2000() Then
        MsgBox "Could not connect to SAP2000.", vbCritical, "Connection Error"
        Exit Sub
    End If

    Dim ret         As Long
    g_SelectedTableKey = tableKey

    ' Get field metadata
    Dim TableVersion As Long
    Dim NumberFields As Long
    Dim FieldKeyArr() As String
    Dim FieldNameArr() As String
    Dim DescriptionArr() As String
    Dim UnitsStringArr() As String
    Dim IsImportableArr() As Boolean

    ret = SapModel.DatabaseTables.GetAllFieldsInTable( _
            g_SelectedTableKey, TableVersion, NumberFields, _
            FieldKeyArr, FieldNameArr, DescriptionArr, _
            UnitsStringArr, IsImportableArr)

    If ret <> 0 Then
        MsgBox "Failed to get table fields." & vbCrLf & "Return code: " & ret, vbCritical, "Error"
        Exit Sub
    End If

    g_TableVersion = TableVersion

    ' Prepare for GetTableForDisplayArray
    Dim FieldKeyListInput() As String
    ReDim FieldKeyListInput(0 To 0) As String
    FieldKeyListInput(0) = ""

    Dim FieldKeysIncludedLocal() As String
    Dim NumberRecords As Long
    Dim tableData() As String

    ret = SapModel.DatabaseTables.GetTableForDisplayArray( _
            g_SelectedTableKey, FieldKeyListInput, "All", _
            TableVersion, FieldKeysIncludedLocal, NumberRecords, tableData)

    If ret <> 0 Then
        MsgBox "Failed to get table data." & vbCrLf & "Return code: " & ret, vbCritical, "Error"
        Exit Sub
    End If

    ' Store metadata
    If Not IsArray(FieldKeysIncludedLocal) Then
        MsgBox "No fields returned from table.", vbExclamation, "Warning"
        Exit Sub
    End If

    Dim i           As Long
    ReDim g_FieldKeysIncluded(LBound(FieldKeysIncludedLocal) To UBound(FieldKeysIncludedLocal))
    For i = LBound(FieldKeysIncludedLocal) To UBound(FieldKeysIncludedLocal)
        g_FieldKeysIncluded(i) = FieldKeysIncludedLocal(i)
    Next i
    g_NumberRecords = NumberRecords

    ' Get Excel ActiveSheet
    Dim xlApp As Object, wb As Object, ws As Object
    On Error Resume Next
    Set xlApp = GetObject(, "Excel.Application")
    If xlApp Is Nothing Then Set xlApp = CreateObject("Excel.Application")
    On Error GoTo ErrHandler

    xlApp.Visible = True
    If xlApp.Workbooks.count = 0 Then
        MsgBox "Please open an Excel workbook first.", vbExclamation, "No Workbook"
        Exit Sub
    End If

    Set wb = xlApp.ActiveWorkbook
    Set ws = xlApp.ActiveSheet

    If ws Is Nothing Then
        MsgBox "No active sheet found.", vbExclamation, "No Sheet"
        Exit Sub
    End If

    ' CLEAR ALL OLD DATA
    ws.Cells.Clear
    ws.Cells.ClearFormats

    g_ExportedSheetName = ws.Name
    g_ExportedWorkbookName = wb.Name

    ' ROW 1: Table Title
    Dim nCols       As Long
    nCols = UBound(g_FieldKeysIncluded) - LBound(g_FieldKeysIncluded) + 1

    ws.Cells(1, 1).Value = "SAP2000 Database: " & g_SelectedTableKey
    ws.Range(ws.Cells(1, 1), ws.Cells(1, nCols)).Merge
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 1).Font.Size = 12
    ws.Cells(1, 1).HorizontalAlignment = -4108
    ws.Cells(1, 1).Interior.color = RGB(68, 114, 196)
    ws.Cells(1, 1).Font.color = RGB(255, 255, 255)
    ws.rows(1).rowHeight = 25

    ' ROW 2: Field Keys (DO NOT CHANGE)
    Dim r As Long, c As Long
    For c = 0 To nCols - 1
        ws.Cells(2, c + 1).Value = g_FieldKeysIncluded(c)
        ws.Cells(2, c + 1).Font.Bold = True
        ws.Cells(2, c + 1).Interior.color = RGB(217, 217, 217)
    Next c

    ' ROW 3: Field Names
    For c = 0 To nCols - 1
        Dim fk      As String: fk = g_FieldKeysIncluded(c)
        Dim foundIdx As Long: foundIdx = -1

        If IsArray(FieldKeyArr) Then
            For r = LBound(FieldKeyArr) To UBound(FieldKeyArr)
                If FieldKeyArr(r) = fk Then
                    foundIdx = r
                    Exit For
                End If
            Next r
        End If

        If foundIdx >= 0 Then
            ws.Cells(3, c + 1).Value = FieldNameArr(foundIdx)
            ws.Cells(3, c + 1).Font.Italic = True
        End If
    Next c

    ' ROW 4: Units
    For c = 0 To nCols - 1
        fk = g_FieldKeysIncluded(c)
        foundIdx = -1

        If IsArray(FieldKeyArr) Then
            For r = LBound(FieldKeyArr) To UBound(FieldKeyArr)
                If FieldKeyArr(r) = fk Then
                    foundIdx = r
                    Exit For
                End If
            Next r
        End If

        If foundIdx >= 0 Then
            ws.Cells(4, c + 1).Value = UnitsStringArr(foundIdx)
            ws.Cells(4, c + 1).Font.color = RGB(128, 128, 128)
        End If
    Next c

    ' ROW 5+: Data
    If NumberRecords > 0 And IsArray(tableData) Then
        Dim baseIdx As Long
        For r = 0 To NumberRecords - 1
            For c = 0 To nCols - 1
                baseIdx = r * nCols + c
                If baseIdx >= LBound(tableData) And baseIdx <= UBound(tableData) Then
                    ws.Cells(r + 5, c + 1).Value = tableData(baseIdx)
                End If
            Next c
        Next r
    End If

    ' Format
    ws.rows("2:4").Font.Bold = True
    ws.Columns("A:AZ").AutoFit

    ' Add borders
    If NumberRecords > 0 Then
        Dim dataRange As Object
        Set dataRange = ws.Range(ws.Cells(2, 1), ws.Cells(NumberRecords + 4, nCols))
        dataRange.Borders.LineStyle = 1
        dataRange.Borders.Weight = 2
    End If

    MsgBox "Database table exported successfully!" & vbCrLf & vbCrLf & _
            "Records: " & NumberRecords & vbCrLf & vbCrLf & _
            "You can now edit the data (row 5 onwards)." & vbCrLf & _
            "Click 'Update to SAP2000' when done.", _
            vbInformation, "Export Complete"

    Exit Sub

ErrHandler:
    MsgBox "Export failed: " & err.description, vbCritical, "Error"
End Sub

'===============================================================
' INTERNAL: Import edited table from ActiveSheet
'===============================================================
Private Sub ImportEditedTableFromActiveSheet(Optional ByVal force As Boolean = False)
    On Error GoTo ErrHandler

    If Not ConnectSAP2000() Then
        MsgBox "Could not connect to SAP2000.", vbCritical, "Connection Error"
        Exit Sub
    End If

    ' Get ActiveSheet
    Dim xlApp As Object, ws As Object
    On Error Resume Next
    Set xlApp = GetObject(, "Excel.Application")
    On Error GoTo ErrHandler

    If xlApp Is Nothing Then
        MsgBox "Excel is not running.", vbExclamation, "Excel Not Found"
        Exit Sub
    End If

    Set ws = xlApp.ActiveSheet
    If ws Is Nothing Then
        MsgBox "No active sheet found.", vbExclamation, "No Sheet"
        Exit Sub
    End If

    ' --- Declare local variables used in this sub (fix Variable not defined compile errors) ---
    Dim title       As String
    Dim lastCol As Long, lastRow As Long
    Dim c As Long, r As Long
    Dim nRows As Long, nCols As Long
    Dim FieldKeysIncludedLocal() As String
    Dim tableData() As String
    Dim baseIdx     As Long
    Dim tmpVer As Long, tmpNum As Long
    Dim tmpFK() As String, tmpFN() As String, tmpDesc() As String
    Dim tmpUnits() As String, tmpImp() As Boolean
    Dim ret         As Long
    Dim FillImportLog As Long
    Dim NumFatalErrors As Long, NumErrorMsgs As Long
    Dim NumWarnMsgs As Long, NumInfoMsgs As Long
    Dim ImportLog   As String

    ' If globals are empty or force requested, attempt to deduce metadata from active sheet
    If Len(Trim(g_SelectedTableKey)) = 0 Or force Then
        title = ""
        On Error Resume Next
        title = CStr(ws.Cells(1, 1).Value)
        On Error GoTo ErrHandler

        If InStr(1, title, "SAP2000 Database:", vbTextCompare) > 0 Then
            Dim deducedKey As String
            deducedKey = Trim(mid(title, Len("SAP2000 Database:") + 1))
            If Len(deducedKey) > 0 Then
                g_SelectedTableKey = deducedKey
                ' Read field keys from row 2 into g_FieldKeysIncluded
                lastCol = 0
                For c = 1 To 512
                    If Trim(CStr(ws.Cells(2, c).Value)) = "" Then
                        lastCol = c - 1
                        Exit For
                    End If
                Next c
                If lastCol = 0 Then lastCol = 512
                ReDim g_FieldKeysIncluded(0 To lastCol - 1)
                For c = 1 To lastCol
                    g_FieldKeysIncluded(c - 1) = CStr(ws.Cells(2, c).Value)
                Next c
                g_ExportedSheetName = ws.Name
                g_ExportedWorkbookName = ws.Parent.Name
                ' Note: g_NumberRecords will be set below once rows counted
            Else
                MsgBox "Active sheet title does not contain a valid table key.", vbExclamation, "No Table Found"
                Exit Sub
            End If
        Else
            MsgBox "No table was exported. Please use 'Get Database to Edit' first or activate the exported sheet.", _
                    vbExclamation, "No Table Selected"
            Exit Sub
        End If
    End If

    ' Find data range
    lastCol = 0
    For c = 1 To 512
        If Trim(CStr(ws.Cells(2, c).Value)) = "" Then
            lastCol = c - 1
            Exit For
        End If
    Next c
    If lastCol = 0 Then lastCol = 512

    lastRow = 4
    For r = 5 To 65536
        Dim anyVal  As Boolean: anyVal = False
        For c = 1 To lastCol
            If Trim(CStr(ws.Cells(r, c).Value)) <> "" Then
                anyVal = True
                Exit For
            End If
        Next c
        If anyVal Then
            lastRow = r
        Else
            If r > 5 And lastRow >= 5 Then Exit For
        End If
    Next r

    If lastRow < 5 Then
        MsgBox "No data found to import.", vbExclamation, "No Data"
        Exit Sub
    End If

    nRows = lastRow - 4
    nCols = lastCol

    ' Read FieldKeysIncluded from row 2 (use sheet values - authoritative)
    ReDim FieldKeysIncludedLocal(0 To nCols - 1) As String
    For c = 1 To nCols
        FieldKeysIncludedLocal(c - 1) = CStr(ws.Cells(2, c).Value)
    Next c

    ' Update module global field list to match what is on sheet (keeps global consistent)
    ReDim g_FieldKeysIncluded(LBound(FieldKeysIncludedLocal) To UBound(FieldKeysIncludedLocal))
    For c = LBound(FieldKeysIncludedLocal) To UBound(FieldKeysIncludedLocal)
        g_FieldKeysIncluded(c) = FieldKeysIncludedLocal(c)
    Next c

    ' Build TableData
    ReDim tableData(0 To nRows * nCols - 1) As String
    For r = 1 To nRows
        For c = 1 To nCols
            baseIdx = (r - 1) * nCols + (c - 1)
            tableData(baseIdx) = CStr(ws.Cells(r + 4, c).Value)
        Next c
    Next r

    ' Update g_NumberRecords for consistency
    g_NumberRecords = nRows

    ' Get TableVersion if needed
    If g_TableVersion = 0 Then
        ret = SapModel.DatabaseTables.GetAllFieldsInTable( _
                g_SelectedTableKey, tmpVer, tmpNum, tmpFK, tmpFN, tmpDesc, tmpUnits, tmpImp)
        If ret = 0 Then g_TableVersion = tmpVer
    End If

    ' Set table for editing
    ret = SapModel.DatabaseTables.SetTableForEditingArray( _
            g_SelectedTableKey, g_TableVersion, FieldKeysIncludedLocal, nRows, tableData)

    If ret <> 0 Then
        MsgBox "Failed to prepare table for import." & vbCrLf & "Return code: " & ret, _
                vbCritical, "Error"
        Exit Sub
    End If

    ' Apply changes
    FillImportLog = 1
    ret = SapModel.DatabaseTables.ApplyEditedTables( _
            FillImportLog, NumFatalErrors, NumErrorMsgs, _
            NumWarnMsgs, NumInfoMsgs, ImportLog)

    If ret <> 0 Then
        MsgBox "Failed to apply changes to SAP2000." & vbCrLf & "Return code: " & ret, _
                vbCritical, "Error"
        Exit Sub
    End If

    ' Show result
    Dim summary     As String
    If NumFatalErrors > 0 Or NumErrorMsgs > 0 Then
        summary = "Update completed WITH ERRORS:" & vbCrLf & vbCrLf & _
                "Fatal Errors: " & NumFatalErrors & vbCrLf & _
                "Errors: " & NumErrorMsgs & vbCrLf & _
                "Warnings: " & NumWarnMsgs & vbCrLf & vbCrLf & _
                "Check the log file for details."

        ' Save log
        Dim logPath As String
        logPath = Environ("TEMP") & "\SAP2000_ImportLog.txt"
        Dim fNum    As Integer
        fNum = FreeFile
        Open logPath For Output As #fNum
        Print #fNum, ImportLog
        Close #fNum

        MsgBox summary & vbCrLf & vbCrLf & "Log saved to: " & logPath, _
                vbExclamation, "Update Complete"
    Else
        summary = "Database updated successfully!" & vbCrLf & vbCrLf & _
                "Warnings: " & NumWarnMsgs & vbCrLf & _
                "Info: " & NumInfoMsgs
        MsgBox summary, vbInformation, "Update Complete"
    End If

    Exit Sub

ErrHandler:
    MsgBox "Import failed: " & err.description, vbCritical, "Error"
End Sub

'===============================================================
' HELPER: Cancel pending edits
'===============================================================
Public Sub CancelTableEditing()
    On Error Resume Next
    If Not SapModel Is Nothing Then
        SapModel.DatabaseTables.CancelTableEditing
    End If
End Sub

'===============================================================
' HELPER: Validate that an active sheet looks like an exported DB sheet
' Returns True if valid, False otherwise.
' Basic checks:
'  - cell(1,1) contains "SAP2000 Database: <tableKey>"
'  - row 2 has at least one non-empty field key
'  - row 3 or 4 exist (optional)
'===============================================================
Public Function IsDatabaseSheetValid(Optional ByVal ws As Object = Nothing) As Boolean
    On Error GoTo ErrHandler
    IsDatabaseSheetValid = False

    Dim xlApp       As Object
    If ws Is Nothing Then
        On Error Resume Next
        Set xlApp = GetObject(, "Excel.Application")
        On Error GoTo ErrHandler
        If xlApp Is Nothing Then Exit Function
        Set ws = xlApp.ActiveSheet
        If ws Is Nothing Then Exit Function
    End If

    Dim title       As String
    title = ""
    title = CStr(ws.Cells(1, 1).Value)
    If InStr(1, title, "SAP2000 Database:", vbTextCompare) = 0 Then Exit Function

    ' check row 2 has at least one non-empty
    Dim c As Long, anyField As Boolean: anyField = False
    For c = 1 To 512
        If Trim(CStr(ws.Cells(2, c).Value)) <> "" Then
            anyField = True
            Exit For
        End If
    Next c
    If Not anyField Then Exit Function

    ' basic pass
    IsDatabaseSheetValid = True
    Exit Function

ErrHandler:
    IsDatabaseSheetValid = False
End Function

