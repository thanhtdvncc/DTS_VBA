Attribute VB_Name = "m06_SAP2000_GroupsValidation"
Option Explicit
'===============================================================
' Module: modSAP2000_GroupsValidation
' Purpose: Extract groups & build validation lists (range-based)
' Notes  :
'   - Avoids Excel's 255-char Data Validation list limit by using
'     a named range as the list source instead of a comma string.
'===============================================================

Public Sub WriteGroups()
    If Not ENABLE_GROUPS Then Exit Sub
    
    Dim ws As Worksheet
    Set ws = SheetOrCreate("Groups", True)
    If ws Is Nothing Then
        LogMsg "WriteGroups: Cannot access Groups sheet."
        Exit Sub
    End If
    ws.Cells.clearContents
    
    Dim groupNames() As String, groupCount As Long, ret As Long
    ret = SapModel.GroupDef.GetNameList(groupCount, groupNames)
    CheckRet ret, "GroupDef.GetNameList"
    If ret <> 0 Or groupCount = 0 Or IsArrayEmpty(groupNames) Then Exit Sub
    
    Dim iCol As Long: iCol = 1
    Dim g As Long, grp As String
    Dim NumberItems As Long
    Dim objectType() As Long
    Dim ObjectName() As String
    
    For g = LBound(groupNames) To UBound(groupNames)
        grp = groupNames(g)
        If UCase$(grp) <> "ALL" Then
            NumberItems = 0
            Erase objectType
            Erase ObjectName
            
            ret = SapModel.GroupDef.GetAssignments(grp, NumberItems, objectType, ObjectName)
            CheckRet ret, "GroupDef.GetAssignments(" & grp & ")"
            
            ws.Cells(1, iCol).Value = grp
            If NumberItems > 0 Then
                ws.Cells(2, iCol).Resize(NumberItems, 1).Value = SafeTranspose(ObjectName)
                ws.Cells(2, iCol + 1).Resize(NumberItems, 1).Value = SafeTranspose(objectType)
            End If
            iCol = iCol + 3
        End If
    Next
End Sub

Public Sub BuildGroupValidation()
    If Not ENABLE_GROUPS Then Exit Sub
    
    ' 1) Get all group names from SAP2000
    Dim groupNames() As String, groupCount As Long, ret As Long
    ret = SapModel.GroupDef.GetNameList(groupCount, groupNames)
    If ret <> 0 Or groupCount = 0 Or IsArrayEmpty(groupNames) Then Exit Sub
    
    ' 2) Filter out "ALL" and remove duplicates (if any)
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim i As Long, grp As String
    For i = LBound(groupNames) To UBound(groupNames)
        grp = CStr(groupNames(i))
        If Len(grp) > 0 Then
            If UCase$(grp) <> "ALL" Then
                If Not dict.exists(grp) Then dict.Add grp, 1
            End If
        End If
    Next i
    If dict.count = 0 Then Exit Sub
    
    ' Convert dictionary keys to an array
    Dim arrGroups() As Variant
    arrGroups = dict.keys   ' 0-based 1D variant array
    
    ' 3) Write list to a helper sheet (GroupList) and create/update a named range
    Dim wsList As Worksheet
    Set wsList = SheetOrCreate("GroupList", True)
    If wsList Is Nothing Then
        LogMsg "BuildGroupValidation: Cannot create/access GroupList sheet."
        Exit Sub
    End If
    
    wsList.Cells.clearContents
    wsList.Range("A1").Resize(UBound(arrGroups) - LBound(arrGroups) + 1, 1).Value = ToVerticalVariant(arrGroups)
    
    Dim lastRow As Long
    lastRow = UBound(arrGroups) - LBound(arrGroups) + 1
    If lastRow <= 0 Then Exit Sub
    
    ' Create or update a workbook-level named range "GroupList"
    Dim refStr As String
    refStr = "=" & wsList.Name & "!$A$1:$A$" & lastRow
    
    Dim nm As Name, found As Boolean
    For Each nm In ThisWorkbook.names
        If StrComp(nm.Name, "GroupList", vbTextCompare) = 0 Then
            nm.RefersTo = refStr
            found = True
            Exit For
        End If
    Next nm
    If Not found Then
        ThisWorkbook.names.Add Name:="GroupList", RefersTo:=refStr
    End If
    
    ' Optionally hide helper sheet (VeryHidden to avoid accidental edits)
    On Error Resume Next
    wsList.Visible = xlSheetHidden ' or xlSheetVeryHidden
    On Error GoTo 0
    
    ' 4) Apply validation to WindIntensity!K4 using the named range
    Dim wsWind As Worksheet
    Set wsWind = SheetOrCreate("WindIntensity", False)
    If Not wsWind Is Nothing Then
        With wsWind.Range("K4").Validation
            On Error Resume Next
            .Delete
            On Error GoTo 0
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                 Operator:=xlBetween, Formula1:="=GroupList"
            .IgnoreBlank = True
            .InCellDropdown = True
        End With
    End If
    
    ' 5) Fill AssignWindArea!H2:H with the same list from the helper sheet
    Dim wsAssign As Worksheet
    Set wsAssign = SheetOrCreate("AssignWindArea", False)
    If Not wsAssign Is Nothing Then
        wsAssign.Range("H2").Resize(lastRow, 1).Value = wsList.Range("A1").Resize(lastRow, 1).Value
    End If
End Sub

