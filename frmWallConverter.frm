VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmWallConverter 
   Caption         =   "SAP2000 Wall Tool"
   ClientHeight    =   6690
   ClientLeft      =   120
   ClientTop       =   675
   ClientWidth     =   5235
   OleObjectBlob   =   "frmWallConverter.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmWallConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ==============================================================================
' listbox config
' ==============================================================================
Private Const COL_IDX_NO As Long = 0
Private Const COL_IDX_HANDLE As Long = 1
Private Const COL_IDX_THICKNESS As Long = 2
Private Const COL_IDX_PATTERN As Long = 3
Private Const COL_IDX_VALUE As Long = 4
Private Const COL_IDX_MAPPING As Long = 5

Private Const LISTBOX_COL_COUNT As Long = 6
Private Const LISTBOX_COL_WIDTHS As String = "20;45;45;35;35;200"

Private acadApp     As Object
Private acadDoc     As Object
Private mSelectedAxes() As Object
Private mAxesCount  As Long
Private m_StoryCount As Long
Private m_StoryElevations As Object  ' Scripting.Dictionary of storyName

' Module-level variables for tracking mouse/selection for lstLoadAssignments
Private mLA_LastMouseX As Single
Private mLA_LastMouseY As Single
Private mLA_LastRowIndex As Long
Private mLA_LastColIndex As Long

' XData application name constant
Private Const XDATA_APP As String = "DTS_APP"
'

' ==========================================================================================
' BUTTON: APPLY LOADS TO CAD (SYNC LISTBOX TO CAD)
' Purpose: Writes data from the ListBox (Pattern/Value) back to the specific Entities (Handle).
' ==========================================================================================
Private Sub btnApplyLoadsToCAD_Click()
    On Error GoTo ErrHandler

    ' 1. Validate ListBox
    If Not HasControl("lstLoadAssignments") Then Exit Sub
    Dim lst         As Object: Set lst = GetControl("lstLoadAssignments")

    If lst.ListCount = 0 Then
        MsgBox "List is empty. Nothing to apply.", vbExclamation, "No Data"
        Exit Sub
    End If

    ' 2. Connect to AutoCAD
    Dim acadApp As Object, acadDoc As Object
    Set acadApp = Nothing: Set acadDoc = Nothing

    On Error Resume Next
    Set acadApp = GetObject(, "AutoCAD.Application")
    If acadApp Is Nothing Then
        MsgBox "AutoCAD not connected!", vbExclamation, "Error"
        Exit Sub
    End If
    On Error GoTo ErrHandler
    Set acadDoc = acadApp.ActiveDocument

    ' 3. Iterate ListBox and Sync to CAD
    Dim i           As Long
    Dim successCount As Long: successCount = 0
    Dim failCount   As Long: failCount = 0

    ' Arrays to collect handles for batch refresh later
    Dim updatedHandles() As String
    ReDim updatedHandles(0 To lst.ListCount - 1)

    For i = 0 To lst.ListCount - 1
        Dim hdl As String, thickStr As String, pat As String, valStr As String

        hdl = Trim(CStr(lst.List(i, COL_IDX_HANDLE)))    ' Col 1
        thickStr = Trim(CStr(lst.List(i, COL_IDX_THICKNESS)))    ' Col 2 (THICKNESS)
        pat = Trim(CStr(lst.List(i, COL_IDX_PATTERN)))   ' Col 3
        valStr = Trim(CStr(lst.List(i, COL_IDX_VALUE)))  ' Col 4

        ' Validate Thickness
        Dim newThick As Double: newThick = 200    ' Default safety
        If IsNumeric(thickStr) Then
            newThick = CDbl(thickStr)
        Else
            ' Try to clean if user typed "W200" by mistake
            thickStr = Replace(UCase(thickStr), "W", "")
            If IsNumeric(thickStr) Then newThick = CDbl(thickStr)
        End If

        ' Validate Load Value
        If IsNumeric(valStr) = False Then valStr = Replace(valStr, ",", ".")
        Dim valNum  As Double: valNum = 0
        If IsNumeric(valStr) Then valNum = CDbl(valStr)

        If Len(hdl) > 0 And newThick > 0 Then
            ' PASS THICKNESS TO FUNCTION
            If UpdateEntityLoadSafe(acadDoc, hdl, newThick, pat, valNum) Then
                updatedHandles(successCount) = hdl
                successCount = successCount + 1
            Else
                failCount = failCount + 1
            End If
        End If
    Next i

    ' 4. Refresh Labels for updated entities
    If successCount > 0 Then
        ReDim Preserve updatedHandles(0 To successCount - 1)
        ' Call the core refresh logic silently
        RefreshLabels_Core acadDoc, updatedHandles, False
        acadDoc.Regen 1
    End If

    ' 5. Report result
    MsgBox "Sync Complete!" & vbCrLf & _
            "Updated: " & successCount & " walls." & vbCrLf & _
            "Failed/Missing: " & failCount & " walls.", vbInformation, "Apply Loads"

    Exit Sub

ErrHandler:
    MsgBox "Error applying loads: " & err.description & " at line " & Erl, vbCritical, "Error"
End Sub

' ==========================================================================================
' HELPER: Update Entity Load Safely (THICKNESS DRIVEN)
' Logic:
' 1. Update Thickness (Index 1).
' 2. Auto-generate WallType (Index 2) as "W" & Thickness.
' 3. Update Load (Index 3,4).
' 4. Preserve Mapping (Index 5+).
' ==========================================================================================
Private Function UpdateEntityLoadSafe(acadDoc As Object, handleStr As String, _
        newThickness As Double, newPattern As String, newValue As Double) As Boolean
    On Error Resume Next
    UpdateEntityLoadSafe = False

    Dim ent         As Object
    Set ent = acadDoc.HandleToObject(handleStr)
    If ent Is Nothing Then Exit Function

    ' Register App
    ent.Application.ActiveDocument.RegisteredApplications.Add "DTS_APP"

    Dim xdType As Variant, xdVal As Variant
    ent.GetXData "DTS_APP", xdType, xdVal

    Dim finalType() As Integer
    Dim finalVal()  As Variant
    Dim currentSize As Long

    ' --- STEP 1: PREPARE ARRAYS ---
    If err.number = 0 And Not IsEmpty(xdVal) And IsArray(xdVal) Then
        currentSize = UBound(xdVal)
        If currentSize < 4 Then currentSize = 4
        ReDim finalType(0 To currentSize)
        ReDim finalVal(0 To currentSize)

        Dim i       As Long
        For i = LBound(xdVal) To UBound(xdVal)
            finalType(i) = CInt(xdType(i))
            finalVal(i) = xdVal(i)
        Next i
    Else
        ' New Structure
        currentSize = 4
        ReDim finalType(0 To 4)
        ReDim finalVal(0 To 4)
        finalType(0) = 1001: finalVal(0) = "DTS_APP"
    End If

    ' --- STEP 2: UPDATE DATA (Thickness Driven) ---

    ' A. Update Thickness (Index 1)
    finalType(1) = 1040
    finalVal(1) = CDbl(newThickness)

    ' B. Auto-Generate Section Name (Index 2)
    finalType(2) = 1000
    finalVal(2) = "W" & CInt(newThickness)    ' e.g., 200 -> "W200"

    ' C. Update Load Data
    finalType(3) = 1000: finalVal(3) = CStr(newPattern)
    finalType(4) = 1040: finalVal(4) = CDbl(newValue)

    ' --- STEP 3: WRITE BACK ---
    ent.SetXData finalType, finalVal

    If err.number = 0 Then UpdateEntityLoadSafe = True
    On Error GoTo 0
End Function
' Add: GetSelectionHandles_OnScreen - returns 0-based array of handles for selected entities (or Empty)
Private Function GetSelectionHandles_OnScreen(acadDoc As Object) As Variant
    On Error GoTo ErrHandler
    If acadDoc Is Nothing Then
        GetSelectionHandles_OnScreen = Empty
        Exit Function
    End If

    Dim ssName      As String: ssName = "WC_GETSEL_" & Format(Now, "hhmmss")
    On Error Resume Next
    acadDoc.SelectionSets.item(ssName).Delete
    On Error GoTo ErrHandler

    Dim ss          As Object
    Set ss = acadDoc.SelectionSets.Add(ssName)
    ss.SelectOnScreen

    If ss.count = 0 Then
        ss.Delete
        GetSelectionHandles_OnScreen = Empty
        Exit Function
    End If

    Dim arr()       As String
    ReDim arr(0 To ss.count - 1)
    Dim i           As Long
    For i = 0 To ss.count - 1
        arr(i) = ss.item(i).Handle
    Next i

    ss.Delete
    GetSelectionHandles_OnScreen = arr
    Exit Function

ErrHandler:
    On Error Resume Next
    If Not ss Is Nothing Then ss.Delete
    GetSelectionHandles_OnScreen = Empty
End Function

Private Sub btnCancel2_Click()
    SaveSettingsToExcel  ' Save settings even when cancel
    Unload Me
End Sub

Public Function CreateWallDictFromEntity(ent As Object) As Object
    ' Returns Scripting.Dictionary with keys:
    ' "Handle","StartX","StartY","StartZ","EndX","EndY","EndZ","Length",
    ' "thickness","wallType","loadPattern","LoadValue"
    On Error GoTo ErrHandler

    Dim dict        As Object
    Set dict = CreateObject("Scripting.Dictionary")

    If ent Is Nothing Then
        Set CreateWallDictFromEntity = dict
        Exit Function
    End If

    ' Initialize defaults
    dict.Add "Handle", ""
    dict.Add "StartX", 0#
    dict.Add "StartY", 0#
    dict.Add "StartZ", 0#
    dict.Add "EndX", 0#
    dict.Add "EndY", 0#
    dict.Add "EndZ", 0#
    dict.Add "Length", 0#
    dict.Add "thickness", 0#
    dict.Add "wallType", ""
    dict.Add "loadPattern", "DL"
    dict.Add "LoadValue", 0#

    ' Geometry - safe read
    On Error Resume Next
    Dim sp As Variant, ep As Variant
    sp = ent.StartPoint
    ep = ent.EndPoint
    If err.number <> 0 Then
        err.Clear
        ' Try alternative properties if needed, but return default dict
        Set CreateWallDictFromEntity = dict
        Exit Function
    End If
    On Error GoTo ErrHandler

    ' Ensure arrays
    If IsArray(sp) And IsArray(ep) Then
        dict("Handle") = CStr(ent.Handle)
        dict("StartX") = CDbl(sp(0))
        dict("StartY") = CDbl(sp(1))
        dict("StartZ") = CDbl(sp(2))
        dict("EndX") = CDbl(ep(0))
        dict("EndY") = CDbl(ep(1))
        dict("EndZ") = CDbl(ep(2))
        dict("Length") = Sqr((dict("EndX") - dict("StartX")) ^ 2 + (dict("EndY") - dict("StartY")) ^ 2)
    End If

    ' Read XData from unified DTS_APP structure (if present)
    On Error Resume Next
    Dim xdType As Variant, xdVal As Variant
    ent.GetXData XDATA_APP, xdType, xdVal
    If err.number = 0 Then
        If Not IsEmpty(xdVal) And IsArray(xdVal) Then
            Dim ub  As Long
            ub = UBound(xdVal)
            ' Indexing follows structure:
            ' 0 = app, 1 = thickness, 2 = wallType, 3 = loadPattern, 4 = loadValue, then mapping groups...
            If ub >= XDATA_OFFSET_THICKNESS Then
                If IsNumeric(xdVal(XDATA_OFFSET_THICKNESS)) Then
                    dict("thickness") = CDbl(xdVal(XDATA_OFFSET_THICKNESS))
                End If
            End If
            If ub >= XDATA_OFFSET_WALLTYPE Then
                If Not IsEmpty(xdVal(XDATA_OFFSET_WALLTYPE)) Then
                    dict("wallType") = CStr(xdVal(XDATA_OFFSET_WALLTYPE))
                End If
            End If
            If ub >= XDATA_OFFSET_LOADPATTERN Then
                If Not IsEmpty(xdVal(XDATA_OFFSET_LOADPATTERN)) Then
                    dict("loadPattern") = CStr(xdVal(XDATA_OFFSET_LOADPATTERN))
                End If
            End If
            If ub >= XDATA_OFFSET_LOADVALUE Then
                If IsNumeric(xdVal(XDATA_OFFSET_LOADVALUE)) Then
                    dict("LoadValue") = CDbl(xdVal(XDATA_OFFSET_LOADVALUE))
                End If
            End If
        End If
    End If
    err.Clear
    On Error GoTo ErrHandler

    Set CreateWallDictFromEntity = dict
    Exit Function

ErrHandler:
    ' On any error, return partial dictionary (if created) to avoid compile/runtime crash
    On Error Resume Next
    Set CreateWallDictFromEntity = dict
End Function

' ==============================================================================
' BUTTON: CLEAR LIST
' Purpose: Clears the ListBox content only (UI only).
'          Does not affect CAD/Excel until "Save" or "OK" is clicked.
' ==============================================================================
Private Sub btnClearList_Click()
    On Error GoTo ErrHandler

    ' 1. Check Control
    If Not HasControl("lstLoadAssignments") Then Exit Sub
    Dim lst         As Object: Set lst = GetControl("lstLoadAssignments")

    ' 2. Check if list is already empty
    If lst.ListCount = 0 Then Exit Sub

    ' 3. Confirm Action
    Dim ans         As VbMsgBoxResult
    ans = MsgBox("Are you sure you want to clear the list?", vbQuestion + vbYesNo, "Clear List")

    If ans = vbYes Then
        ' 4. Clear UI
        lst.Clear
    End If

    Exit Sub

ErrHandler:
    MsgBox "Error clearing list: " & err.description, vbCritical
End Sub

' ==================== btnCombineWithSAP - MAIN MAPPING FUNCTION ====================
Private Sub btnCombineWithSAP_Click()
    On Error GoTo ErrHandler

Debug.Print "========== Starting Combine with SAP2000 =========="

    ' 1. Connect to SAP2000
    ConnectSAP2000
    If SapModel Is Nothing Then
        MsgBox "SAP2000 not connected!", vbExclamation, "Error"
        Exit Sub
    End If

    ' 2. Get Story Info (Priority: Manual -> List)
    Dim storyInfo   As Object
    Set storyInfo = GetSelectedStory()
    If storyInfo Is Nothing Then
        MsgBox "Please select a story from the list OR enter valid Manual Elevation and Height.", vbExclamation, "Missing Input"
        Exit Sub
    End If

    ' 3. Get Insertion Point
    Dim InsertPt    As Variant
    InsertPt = GetInsertionPoint()
    If IsEmpty(InsertPt) Then
        MsgBox "Invalid insertion point!", vbExclamation, "Error"
        Exit Sub
    End If

    ' 4. Select Wall Lines in AutoCAD
    Dim acadApp As Object, acadDoc As Object
    Set acadApp = Nothing
    Set acadDoc = Nothing

    On Error Resume Next
    Set acadApp = GetObject(, "AutoCAD.Application")
    On Error GoTo ErrHandler

    If acadApp Is Nothing Then
        MsgBox "AutoCAD not connected!", vbExclamation, "Error"
        Exit Sub
    End If
    Set acadDoc = acadApp.ActiveDocument

    BringAutoCADToFront

    ' Using existing helper to get handles
    Dim handles     As Variant
    handles = GetSelectionHandles_OnScreen(acadDoc)

    ' ? FIX 1: Validate handles array
    If IsEmpty(handles) Then
        'MsgBox "No entities selected. Operation cancelled.", vbInformation, "Cancelled"
        Exit Sub
    End If

    If Not IsArray(handles) Then
        'MsgBox "Invalid selection data!", vbExclamation, "Error"
        Exit Sub
    End If

    If UBound(handles) < LBound(handles) Then
        'MsgBox "No valid entities selected!", vbExclamation, "Error"
        Exit Sub
    End If

    ' 5. Use New API to Get SAP Frames in Z-Range
    Dim zMin As Double, zMax As Double, currentElev As Double
    currentElev = CDbl(storyInfo("Elevation"))

    zMin = currentElev - 200
    zMax = currentElev + 200

    ' Call API DTS_SAP2000_Getlist2node (Mode 3 = Map)
    Dim propAny(0)  As Variant
    propAny(0) = "Any"

    Dim frameNodeMap As Object
    Set frameNodeMap = Nothing

    On Error Resume Next
    Set frameNodeMap = DTS_SAP2000_Getlist2node("Frame", "Any", propAny, , , , , zMin, zMax, 3)
    On Error GoTo ErrHandler

    ' ? FIX 2: Validate frameNodeMap
    If frameNodeMap Is Nothing Then
        'MsgBox "Failed to retrieve SAP frames!", vbExclamation, "Error"
        Exit Sub
    End If

    If frameNodeMap.count = 0 Then
        'MsgBox "No SAP frames found at elevation range " & zMin & " - " & zMax, vbExclamation, "No Frames"
        Exit Sub
    End If

    ' 6. EXECUTE MAPPING (with validation)
    Dim mappedCount As Long
    mappedCount = 0

    On Error Resume Next
    mappedCount = n02_ACAD_Wall_Force_SAP2000.ExecuteWallMapping_Batch( _
            acadDoc, SapModel, frameNodeMap, handles, InsertPt, CDbl(storyInfo("Elevation")))

    If err.number <> 0 Then
        'MsgBox "Mapping failed: " & err.description, vbCritical, "Error"
        err.Clear
        Exit Sub
    End If
    On Error GoTo ErrHandler

    ' ? FIX 3: Only refresh labels if mapping succeeded
    If mappedCount > 0 Then
        ' 7. AUTO REFRESH LABELS (Reuse handles - No user prompt)
        On Error Resume Next
        RefreshLabels_Core acadDoc, handles, False
        On Error GoTo ErrHandler

        ' 8. Force Regen
        acadDoc.Regen 1
    End If

Debug.Print "Mapping Complete: " & mappedCount & " walls processed"

    Exit Sub

ErrHandler:
Debug.Print "ERROR in btnCombineWithSAP: " & err.description & " (Line: " & Erl & ")"
    MsgBox "Error in Combine: " & err.description, vbCritical, "Error"
End Sub

' ==================== DEBUG: Show handles with DTS_APP XData ====================
' Adds handler for a form button named "btnDebug" (you must add the button control to the form)
' Usage: select entities in AutoCAD and click btnDebug -> opens notepad with details
Private Sub btnDebug_Click()
    On Error GoTo ErrHandler

    ' Connect to AutoCAD instance
    Dim acadAppLocal As Object, acadDocLocal As Object
    Set acadAppLocal = Nothing
    Set acadDocLocal = Nothing

    On Error Resume Next
    Set acadAppLocal = GetObject(, "AutoCAD.Application")
    If acadAppLocal Is Nothing Then
        MsgBox "AutoCAD not connected!", vbExclamation, "Error"
        Exit Sub
    End If
    On Error GoTo ErrHandler
    Set acadDocLocal = acadAppLocal.ActiveDocument

    ' Bring AutoCAD front
    AppActivate acadAppLocal.Caption
    DoEvents

    ' Try to get selection handles first
    Dim handles     As Variant
    handles = GetSelectionHandles_OnScreen(acadDocLocal)

    If IsEmpty(handles) Then
        Dim ans     As VbMsgBoxResult
        ans = MsgBox("No selection detected. Do you want to scan the entire ModelSpace for layer 'DTS_WALL_DIAGRAM'?", vbYesNo + vbQuestion, "Scan ModelSpace?")
        If ans = vbYes Then
            handles = GetAllHandlesInLayer(acadDocLocal, "DTS_WALL_DIAGRAM")
            If IsEmpty(handles) Then
                MsgBox "No entities found on layer 'DTS_WALL_DIAGRAM'.", vbInformation, "No Entities"
                Exit Sub
            End If
        Else
            Exit Sub
        End If
    End If

    ' Show handles with XData details
    ShowHandlesWithXData acadDocLocal, handles

    Exit Sub

ErrHandler:
    MsgBox "Error in btnDebug_Click: " & err.description, vbCritical, "Error"
End Sub

' Return 0-based array of handles for all line entities on the specified layer.
' If layerName is empty, returns all line handles.
Private Function GetAllHandlesInLayer(acadDoc As Object, layerName As String) As Variant
    On Error GoTo ErrHandler
    If acadDoc Is Nothing Then
        GetAllHandlesInLayer = Empty
        Exit Function
    End If

    Dim ms          As Object
    Set ms = acadDoc.ModelSpace

    Dim coll        As Object
    Set coll = CreateObject("Scripting.Dictionary")

    Dim ent         As Object
    Dim h           As String
    For Each ent In ms
        On Error Resume Next
        h = CStr(ent.Handle)
        On Error GoTo ErrHandler

        ' Only include line entities
        If Not IsLineEntity(ent) Then GoTo NextEnt

        ' If layerName provided, do tolerant match (underscore/space / case-insensitive)
        If Len(Trim$(layerName)) > 0 Then
            Dim entLayerLower As String
            entLayerLower = LCase$(Trim$(ent.layer))
            Dim layerLower As String
            layerLower = LCase$(layerName)
            If entLayerLower <> layerLower Then
                ' tolerant token match
                If Not (InStr(1, entLayerLower, "dts", vbTextCompare) > 0 And _
                        InStr(1, entLayerLower, "wall", vbTextCompare) > 0 And _
                        InStr(1, entLayerLower, "diagram", vbTextCompare) > 0) Then
                    GoTo NextEnt
                End If
            End If
        End If

        If Not coll.exists(h) Then coll.Add h, True

NextEnt:
    Next ent

    If coll.count = 0 Then
        GetAllHandlesInLayer = Empty
        Exit Function
    End If

    Dim arr()       As String
    ReDim arr(0 To coll.count - 1)
    Dim i           As Long
    i = 0
    Dim key         As Variant
    For Each key In coll.keys
        arr(i) = CStr(key)
        i = i + 1
    Next key

    GetAllHandlesInLayer = arr
    Exit Function

ErrHandler:
    On Error Resume Next
    GetAllHandlesInLayer = Empty
End Function



' Utility: write text content to file (overwrite) and close
Private Sub WriteTextToFile(filePath As String, content As String)
    On Error Resume Next
    Dim fn          As Integer
    fn = FreeFile
    Open filePath For Output As #fn
    Print #fn, content
    Close #fn
End Sub

' ==========================================================================================
' BUTTON: EDIT MAPPING (SYNC TEXT TO DATA) - MASS UPDATE
' Logic:
' 1. User selects Labels (Text/MText).
' 2. Code parses text -> compares with Wall XData.
' 3. If Load changes -> Update Load (No Override Flag).
' 4. If Mapping changes -> Update Mapping + SET OVERRIDE FLAG.
' 5. If Syntax Error -> Revert label to original data.
' ==========================================================================================
Private Sub btnEditMapping_Click()
    On Error GoTo ErrHandler

    ' 1. Connect to AutoCAD
    Dim acadApp As Object, acadDoc As Object
    Set acadApp = Nothing: Set acadDoc = Nothing
    On Error Resume Next
    Set acadApp = GetObject(, "AutoCAD.Application")
    If acadApp Is Nothing Then Exit Sub
    Set acadDoc = acadApp.ActiveDocument
    
    BringAutoCADToFront

    ' 2. Select Labels
    Dim ss          As Object
    Dim ssName      As String: ssName = "SYNC_LABELS_SEL"
    
    On Error Resume Next
    acadDoc.SelectionSets.item(ssName).Delete
    On Error GoTo 0
    Set ss = acadDoc.SelectionSets.Add(ssName)

    Dim gpCode(0 To 1) As Integer
    Dim dataVal(0 To 1) As Variant
    gpCode(0) = 8: dataVal(0) = "dts_frame_label" ' Only Labels
    gpCode(1) = 0: dataVal(1) = "*TEXT"           ' Text or MText

    On Error Resume Next
    acadDoc.Utility.prompt vbCrLf & "Select labels to Sync/Update (Press Enter for All): "
    ss.SelectOnScreen gpCode, dataVal
    ' Auto select all if user hits enter
    If ss.count = 0 Then ss.Select 5, , , gpCode, dataVal
    On Error GoTo ErrHandler

    If ss.count = 0 Then Exit Sub

    ' 3. Process Each Label
    Dim i As Long
    Dim updateCount As Long: updateCount = 0
    Dim revertCount As Long: revertCount = 0
    Dim overrideCount As Long: overrideCount = 0
    
    ' Array to collect handles for batch refresh
    Dim refreshHandles() As String
    Dim rIdx As Long: rIdx = 0
    ReDim refreshHandles(0 To ss.count - 1)

    For i = 0 To ss.count - 1
        Dim lblObj As Object: Set lblObj = ss.item(i)
        Dim txtStr As String: txtStr = ""
        
        ' Get Text Content
        On Error Resume Next
        txtStr = lblObj.TextString
        If txtStr = "" Then txtStr = lblObj.text
        On Error GoTo 0
        
        ' Get Linked Handle
        Dim hWall As String
        hWall = ReadLabelXData(lblObj)
        
        ' If linked handle exists
        If hWall <> "" Then
            Dim wallEnt As Object: Set wallEnt = Nothing
            On Error Resume Next
            Set wallEnt = acadDoc.HandleToObject(hWall)
            On Error GoTo ErrHandler
            
            If Not wallEnt Is Nothing Then
                ' ANALYZE AND SYNC
                Dim res As Integer
                ' res: 0=NoChange, 1=LoadUpdate, 2=MappingOverride, -1=Error/Revert
                res = SyncLabelToWallData(wallEnt, txtStr)
                
                If res <> 0 Then
                    refreshHandles(rIdx) = hWall
                    rIdx = rIdx + 1
                End If
                
                If res = 1 Then updateCount = updateCount + 1
                If res = 2 Then overrideCount = overrideCount + 1
                If res = -1 Then revertCount = revertCount + 1
            End If
        End If
    Next i
    
    ss.Delete
    
    ' 4. Refresh Visuals (Standardize formatting for updated/reverted labels)
    If rIdx > 0 Then
        ReDim Preserve refreshHandles(0 To rIdx - 1)
        RefreshLabels_Core acadDoc, refreshHandles, False
    End If

    Exit Sub

ErrHandler:
    MsgBox "Error syncing labels: " & err.description, vbCritical
    On Error Resume Next
    If Not ss Is Nothing Then ss.Delete
End Sub

' ==========================================================================================
' CORE LOGIC: Sync Text -> Data (Strict Override Rules)
' Returns: 0=NoChange, 1=LoadUpdate, 2=MappingOverride, -1=Error
' FIX: Compare with BASE mapping, not effective mapping
' ==========================================================================================
Private Function SyncLabelToWallData(wallEnt As Object, labelText As String) As Integer
    On Error GoTo ErrHandler
    SyncLabelToWallData = 0

    ' 1. Parse Text Input
    Dim pPat As String, pVal As Double, pMapStr As String
    If Not ParseLabelContentRobust(labelText, pPat, pVal, pMapStr) Then
        SyncLabelToWallData = -1
        Exit Function
    End If

    ' 2. Read Current XData
    Dim baseWall As WallSegmentMap, baseMappings() As MappingRecord
    Dim ovWall As WallSegmentMap, ovMappings() As MappingRecord
    Dim hasOv As Boolean, ovCount As Long, baseCount As Long
    
    baseCount = n02_ACAD_Wall_Force_SAP2000.ReadWallAllXData( _
            wallEnt, baseWall, baseMappings, ovWall, ovMappings, hasOv, ovCount)

    ' 3. Prepare New Mapping Data from Text
    Dim cleanMapStr As String
    cleanMapStr = CleanMappingString(pMapStr)
    
    Dim parseInput As String
    If InStr(1, LCase(cleanMapStr), "to ") = 0 And Len(Trim(cleanMapStr)) > 0 Then
        parseInput = "to " & cleanMapStr
    Else
        parseInput = cleanMapStr
    End If
    
    Dim newMappings() As MappingRecord
    Dim newCount As Long
    newCount = n02_ACAD_Wall_Force_SAP2000.ParseMappingLabel(parseInput, newMappings)
    
    ' Backfill Length for "Full" mappings
    Dim sp As Variant, ep As Variant
    sp = wallEnt.StartPoint: ep = wallEnt.EndPoint
    Dim wLen As Double
    wLen = Sqr((sp(0) - ep(0)) ^ 2 + (sp(1) - ep(1)) ^ 2)
    
    Dim k As Long
    For k = 0 To newCount - 1
        If Abs(newMappings(k).DistJ - newMappings(k).DistI) < 0.001 Then
            newMappings(k).DistI = 0
            newMappings(k).DistJ = wLen
            newMappings(k).FrameLength = wLen
        End If
    Next k

    ' 4. CRITICAL FIX: Compare with BASE mapping (original from SAP)
    '    NOT with effective (which may already have override)
    Dim mapChanged As Boolean
    mapChanged = Not AreMappingSetsSemanticallyEqual(baseMappings, baseCount, newMappings, newCount)
    
    ' 5. Compare Load (with BASE load)
    Dim loadChanged As Boolean
    loadChanged = False
    If UCase(Trim(baseWall.LoadPattern)) <> UCase(Trim(pPat)) Then loadChanged = True
    If Abs(baseWall.LoadValue - pVal) > 0.001 Then loadChanged = True

    ' 6. RULE LOGIC
    
    ' RULE 1: Mapping changed from BASE -> OVERRIDE (Purple + Yellow Circle)
    If mapChanged Then
        Dim newOvWall As WallSegmentMap
        newOvWall = baseWall
        newOvWall.LoadPattern = pPat
        newOvWall.LoadValue = pVal
        
        n02_ACAD_Wall_Force_SAP2000.WriteWallCompleteXData_WithOverride _
                wallEnt, baseWall, baseMappings, baseCount, True, newOvWall, newMappings, newCount
        
        SyncLabelToWallData = 2
        Exit Function
    End If
    
    ' RULE 2: Only Load changed, Mapping unchanged
    If loadChanged Then
        ' Update BASE (keep Auto color: Yellow/Green/Red)
        UpdateEntityLoadXData wallEnt.Application.ActiveDocument, wallEnt.Handle, pPat, pVal
        
        ' If currently has Override -> Remove Override flag (back to Auto state)
        If hasOv Then
            n02_ACAD_Wall_Force_SAP2000.WriteWallCompleteXData_WithOverride _
                wallEnt, baseWall, baseMappings, baseCount, False, baseWall, baseMappings, 0
        End If
        
        SyncLabelToWallData = 1
        Exit Function
    End If

    ' No Change
    SyncLabelToWallData = 0
    Exit Function

ErrHandler:
    SyncLabelToWallData = -1
End Function

' Helper to clean "(full 8.0m)" from string
Private Function CleanMappingString(raw As String) As String
    Dim s As String
    s = raw
    
    ' Remove anything in parenthesis
    Dim pOpen As Long, pClose As Long
    Do
        pOpen = InStr(1, s, "(")
        pClose = InStr(1, s, ")")
        If pOpen > 0 And pClose > pOpen Then
            s = Left(s, pOpen - 1) & mid(s, pClose + 1)
        Else
            Exit Do
        End If
    Loop
    
    CleanMappingString = Trim(s)
End Function
' ==========================================================================================
' HELPER: Smart Mapping Comparison (Order Independent, Strict Tolerance)
' FIX: Compare "189" vs "189 (full 18m)" must be EQUAL
'      But "189" vs "189 (I=0to2.5)" must be DIFFERENT
' ==========================================================================================
Private Function AreMappingSetsSemanticallyEqual(m1() As MappingRecord, C1 As Long, _
                                                 m2() As MappingRecord, C2 As Long) As Boolean
    On Error Resume Next
    AreMappingSetsSemanticallyEqual = False

    ' 1. Count Check
    If C1 <> C2 Then Exit Function
    If C1 = 0 Then
        AreMappingSetsSemanticallyEqual = True
        Exit Function
    End If

    ' 2. Clone Arrays
    Dim s1() As MappingRecord, s2() As MappingRecord
    ReDim s1(0 To C1 - 1): ReDim s2(0 To C2 - 1)
    Dim i As Long
    For i = 0 To C1 - 1: s1(i) = m1(i): Next i
    For i = 0 To C2 - 1: s2(i) = m2(i): Next i

    ' 3. Sort Arrays
    SortMappingArray s1, C1
    SortMappingArray s2, C2

    ' 4. Deep Compare with STRICT distance check
    For i = 0 To C1 - 1
        ' Compare Frame Name
        If UCase(Trim(s1(i).TargetFrame)) <> UCase(Trim(s2(i).TargetFrame)) Then Exit Function
        
        ' CRITICAL: Compare distances STRICTLY
        ' Tolerance 10mm - If difference > 10mm -> DIFFERENT
        ' Example: (0,18000) vs (0,2500) -> CLEARLY DIFFERENT
        '          (0,18000) vs (0,18005) -> SAME (within tolerance)
        If Abs(s1(i).DistI - s2(i).DistI) > 10 Then Exit Function ' 10mm tolerance
        If Abs(s1(i).DistJ - s2(i).DistJ) > 10 Then Exit Function
        
        ' Compare MapType
        If UCase(Trim(s1(i).MapType)) <> UCase(Trim(s2(i).MapType)) Then Exit Function
    Next i

    AreMappingSetsSemanticallyEqual = True
End Function

' Helper: Bubble Sort for Mapping Records
Private Sub SortMappingArray(arr() As MappingRecord, count As Long)
    Dim i As Long, j As Long
    Dim temp As MappingRecord
    Dim key1 As String, key2 As String
    
    For i = 0 To count - 2
        For j = i + 1 To count - 1
            ' Build comparison keys
            key1 = UCase(Trim(arr(i).TargetFrame)) & "_" & Format(arr(i).DistI, "00000.00")
            key2 = UCase(Trim(arr(j).TargetFrame)) & "_" & Format(arr(j).DistI, "00000.00")
            
            ' Swap if wrong order
            If key1 > key2 Then
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            End If
        Next j
    Next i
End Sub
' ==========================================================================================
' HELPER: Robust String Parser (Tolerates spaces, flexible format)
' Expected: "[Handle] W200 DL=12.5 ..." OR "DL=12.5 ..."
' ==========================================================================================
Private Function ParseLabelContentRobust(rawTxt As String, _
        ByRef pat As String, ByRef val As Double, ByRef mapStr As String) As Boolean
    On Error Resume Next
    ParseLabelContentRobust = False
    
    Dim s As String: s = rawTxt
    
    ' 1. Clean Handle "[...]"
    Dim pClose As Long
    pClose = InStr(1, s, "]")
    If pClose > 0 Then s = mid$(s, pClose + 1)
    s = Trim$(s)
    
    ' 2. Find "to " split (Separator between Load and Mapping)
    Dim pTo As Long
    pTo = InStr(1, s, " to ", vbTextCompare)
    
    Dim loadPart As String
    If pTo > 0 Then
        loadPart = Left$(s, pTo - 1)
        mapStr = mid$(s, pTo) ' Keep "to " part
    Else
        loadPart = s
        mapStr = ""
    End If
    
    ' 3. Parse Load Part
    ' Remove units to avoid confusion
    loadPart = Replace(loadPart, "kN/m2", "", , , vbTextCompare)
    loadPart = Replace(loadPart, "kN/m", "", , , vbTextCompare)
    
    ' Find "="
    Dim pEq As Long
    pEq = InStr(1, loadPart, "=")
    If pEq = 0 Then Exit Function ' Must have "="
    
    ' --- VALUE (Right of =) ---
    Dim vStr As String
    vStr = Trim$(mid$(loadPart, pEq + 1))
    If IsNumeric(vStr) Then
        val = CDbl(vStr)
    Else
        Exit Function
    End If
    
    ' --- PATTERN (Left of =) ---
    Dim leftPart As String
    leftPart = Trim$(Left$(loadPart, pEq - 1))
    
    ' Logic: Pattern is the LAST word before "="
    ' Example: "W200 DL" -> Pattern is DL
    ' Example: "DL" -> Pattern is DL
    Dim parts() As String
    parts = Split(leftPart, " ")
    
    If UBound(parts) >= 0 Then
        pat = Trim$(parts(UBound(parts)))
    Else
        pat = Trim$(leftPart)
    End If
    
    ' Final check
    If Len(pat) > 0 Then ParseLabelContentRobust = True
End Function
' ==========================================================================================
' HELPER: Parse Complete Label Data (Pattern, Value, WallType)
' Input: "W200 DL=7.20 kN/m2 to ..."
' Updates: wallSeg structure
' ==========================================================================================
Private Function ParseCompleteLabelData(labelText As String, ByRef wallSeg As WallSegmentMap) As Boolean
    On Error Resume Next
    ParseCompleteLabelData = False

    ' Remove handle prefix if present
    Dim cleanLabel  As String
    cleanLabel = labelText
    If InStr(1, cleanLabel, "]") > 0 Then
        cleanLabel = Trim$(mid$(cleanLabel, InStr(1, cleanLabel, "]") + 1))
    End If

    ' Find "to" keyword
    Dim toPos       As Long
    toPos = InStr(1, cleanLabel, " to ", vbTextCompare)
    If toPos = 0 Then Exit Function

    ' Extract load part: "W200 DL=7.20 kN/m2"
    Dim loadPart    As String
    loadPart = Trim$(Left$(cleanLabel, toPos - 1))

    ' Split by space
    Dim parts()     As String
    parts = Split(loadPart, " ")
    If UBound(parts) < 1 Then Exit Function

    ' Part 0: WallType (e.g., "W200")
    wallSeg.WallType = Trim$(parts(0))

    ' Extract thickness from WallType if needed
    If Left$(wallSeg.WallType, 1) = "W" And Len(wallSeg.WallType) > 1 Then
        Dim thkStr  As String
        thkStr = mid$(wallSeg.WallType, 2)
        If IsNumeric(thkStr) Then
            wallSeg.Thickness = CDbl(thkStr)
        End If
    End If

    ' Part 1: "Pattern=Value" or "Pattern=Value kN/m2"
    Dim loadStr     As String
    loadStr = Trim$(parts(1))

    Dim eqPos       As Long
    eqPos = InStr(1, loadStr, "=")
    If eqPos = 0 Then Exit Function

    ' Extract Pattern
    wallSeg.LoadPattern = Trim$(Left$(loadStr, eqPos - 1))

    ' Extract Value
    Dim valStr      As String
    valStr = Trim$(mid$(loadStr, eqPos + 1))

    ' Remove units if present
    valStr = Replace(valStr, "kN/m2", "", , , vbTextCompare)
    valStr = Replace(valStr, "kN/m", "", , , vbTextCompare)
    valStr = Trim$(valStr)

    If Not IsNumeric(valStr) Then Exit Function
    wallSeg.LoadValue = CDbl(valStr)

    ParseCompleteLabelData = True
End Function

' ==========================================================================================
' HELPER: Attach Line Handle to Label XData (Using DTS_APP)
' Structure: Index 0 = "DTS_APP", Index 1 = Line Handle (String)
' ==========================================================================================
Private Sub AttachLabelXData(labelObj As Object, lineHandle As String)
    On Error Resume Next
    If labelObj Is Nothing Then Exit Sub
    If Len(Trim$(lineHandle)) = 0 Then Exit Sub

    ' Register app
    labelObj.Application.ActiveDocument.RegisteredApplications.Add XDATA_APP

    ' Build XData: Simple structure with handle only
    Dim xdType(0 To 1) As Integer
    Dim xdVal(0 To 1) As Variant

    xdType(0) = 1001: xdVal(0) = XDATA_APP          ' App name
    xdType(1) = 1000: xdVal(1) = CStr(lineHandle)   ' Line handle

    labelObj.SetXData xdType, xdVal
End Sub

' ==========================================================================================
' HELPER: Read Line Handle from Label XData
' Returns: Line handle string, or "" if not found
' ==========================================================================================
Private Function ReadLabelXData(labelObj As Object) As String
    On Error Resume Next
    ReadLabelXData = ""
    If labelObj Is Nothing Then Exit Function

    Dim xdType As Variant, xdVal As Variant
    labelObj.GetXData XDATA_APP, xdType, xdVal

    If err.number <> 0 Then
        err.Clear
        Exit Function
    End If

    If IsEmpty(xdVal) Or Not IsArray(xdVal) Then Exit Function
    If UBound(xdVal) < 1 Then Exit Function

    ' Read handle from index 1
    ReadLabelXData = CStr(xdVal(1))
End Function

' ==========================================================================================
' HELPER: Find Label by Line Handle (via XData search)
' Returns: Label object, or Nothing if not found
' ==========================================================================================
Private Function FindLabelByHandle(acadDoc As Object, targetHandle As String) As Object
    On Error Resume Next
    Set FindLabelByHandle = Nothing
    If acadDoc Is Nothing Then Exit Function
    If Len(Trim$(targetHandle)) = 0 Then Exit Function

    Dim ms          As Object
    Set ms = acadDoc.ModelSpace

    Dim ent         As Object
    For Each ent In ms
        ' Check layer first (performance)
        If LCase$(Trim$(ent.layer)) = "dts_frame_label" Then
            ' Check object type
            Dim objType As String
            objType = LCase$(Trim$(ent.ObjectName))

            If objType = "acdbtext" Or objType = "acdbmtext" Then
                ' Read XData
                Dim hdl As String
                hdl = ReadLabelXData(ent)

                If hdl <> "" And LCase$(hdl) = LCase$(targetHandle) Then
                    Set FindLabelByHandle = ent
                    Exit Function
                End If
            End If
        End If
    Next ent
End Function
' ==========================================================================================
' ULTRA-FAST: Build Label Cache using SelectionSet Filters
' Instead of iterating ModelSpace (slow), we ask AutoCAD to give us only the labels.
' ==========================================================================================
Private Function BuildLabelCacheFromHandles(acadDoc As Object, lineHandles As Variant) As Object
    On Error Resume Next
    Set BuildLabelCacheFromHandles = CreateObject("Scripting.Dictionary")

    If acadDoc Is Nothing Then Exit Function
    If IsEmpty(lineHandles) Or Not IsArray(lineHandles) Then Exit Function

    ' 1. Build Hashset of target handles for O(1) lookup
    Dim targetHandles As Object
    Set targetHandles = CreateObject("Scripting.Dictionary")
    Dim i           As Long
    For i = LBound(lineHandles) To UBound(lineHandles)
        Dim key     As String
        key = LCase$(Trim$(CStr(lineHandles(i))))
        If Len(key) > 0 Then
            If Not targetHandles.exists(key) Then targetHandles.Add key, True
        End If
    Next i

    If targetHandles.count = 0 Then Exit Function

    ' 2. Use SelectionSet with Filter
    Dim ssName      As String
    ssName = "CACHE_FILTER_" & Format(Now, "hhmmss")

    On Error Resume Next
    acadDoc.SelectionSets.item(ssName).Delete
    On Error GoTo 0

    Dim ss          As Object
    Set ss = acadDoc.SelectionSets.Add(ssName)

    ' Filter Code: Group 0 = Object Type (*TEXT), Group 8 = Layer Name (dts_frame_label)
    Dim gpCode(0 To 1) As Integer
    Dim dataVal(0 To 1) As Variant

    gpCode(0) = 0: dataVal(0) = "*TEXT"             ' Text or MText
    gpCode(1) = 8: dataVal(1) = "dts_frame_label"   ' Only this layer

    ' SelectAll is much faster than iterating because it uses AutoCAD's internal spatial index
    ss.Select 5, , , gpCode, dataVal    ' 5 = acSelectionSetAll

    If ss.count = 0 Then
        ss.Delete
        Exit Function
    End If

    ' 3. Iterate ONLY the labels (much smaller set)
    Dim ent         As Object
    Dim linkedHandle As String

    For i = 0 To ss.count - 1
        Set ent = ss.item(i)

        ' Read XData to get linked handle
        linkedHandle = ReadLabelXData(ent)

        If Len(linkedHandle) > 0 Then
            Dim handleKey As String
            handleKey = LCase$(linkedHandle)

            ' If this label belongs to one of our selected lines, cache it
            If targetHandles.exists(handleKey) Then
                If Not BuildLabelCacheFromHandles.exists(handleKey) Then
                    BuildLabelCacheFromHandles.Add handleKey, ent
                End If
            End If
        End If
    Next i

    ss.Delete
End Function

' ==============================================================================
' BUTTON: REFRESH LABELS (UPDATED)
' ==============================================================================
Private Sub btnRefreshLabels_Click()
    On Error GoTo ErrHandler

    Set acadApp = Nothing: Set acadDoc = Nothing
    On Error Resume Next
    Set acadApp = GetObject(, "AutoCAD.Application")
    If acadApp Is Nothing Then MsgBox "AutoCAD not connected!": Exit Sub
    Set acadDoc = acadApp.ActiveDocument
    On Error GoTo ErrHandler

    BringAutoCADToFront

    Dim handles     As Variant
    handles = GetSelectionHandles_OnScreen(acadDoc)
    If IsEmpty(handles) Then Exit Sub

    ' Call Core -> Auto Delete Old -> Create New -> Silent
    RefreshLabels_Core acadDoc, handles, False

    Exit Sub
ErrHandler:
Debug.Print "Error: " & err.description
End Sub
' ==============================================================================
' OPTIMIZED: Refresh labels (STRICT COLOR RULES)
' Rules:
' 1. Not Mapped -> Yellow (2)
' 2. Mapped (Auto) -> Green (3)
' 3. New Frame -> Red (1)
' 4. Override -> Line/Handle Magenta (6) + Circle Yellow (2)
' ==============================================================================
Private Sub RefreshLabels_Core(acadDoc As Object, handles As Variant, showMsg As Boolean)
    On Error Resume Next

    If IsEmpty(handles) Or Not IsArray(handles) Then Exit Sub

    Core_CAD_Plotter.EnsureLabelLayer acadDoc, 8

    Dim dynHeight   As Double
    dynHeight = CalculateDynamicTextHeight(acadDoc, handles)

    Dim labelCache  As Object
    Set labelCache = BuildLabelCacheFromHandles(acadDoc, handles)

    Dim i           As Long
    Dim createdCount As Long: createdCount = 0
    Dim updatedCount As Long: updatedCount = 0

    Dim prevVar     As Variant
    prevVar = acadDoc.GetVariable("CMDECHO")
    acadDoc.SetVariable "CMDECHO", 0

    For i = LBound(handles) To UBound(handles)
        Dim hdl     As String
        hdl = CStr(handles(i))

        Dim ent     As Object
        Set ent = Nothing
        Set ent = acadDoc.HandleToObject(hdl)

        If Not ent Is Nothing Then
            If ent.ObjectName = "AcDbLine" Then
                If LCase$(Trim$(ent.layer)) = "dts_wall_diagram" Then

                    ' 1. READ DATA
                    Dim effWall As WallSegmentMap
                    Dim effMappings() As MappingRecord
                    Dim effCount As Long
                    Dim hasOv As Boolean

                    If Not ReadEffectiveWallData(ent, effWall, effMappings, effCount, hasOv) Then
                        GoTo NextHandle
                    End If

                    ' 2. DETERMINE COLOR & STATE
                    Dim mainColor As Integer    ' For Line and Handle text
                    Dim drawCircle As Boolean: drawCircle = False
                    Dim circleColor As Integer: circleColor = 2 ' Default Yellow

                    If hasOv Then
                        ' === CASE 4: OVERRIDE ===
                        mainColor = 6       ' Magenta for Line & Handle
                        drawCircle = True   ' Draw Marker
                        circleColor = 2     ' Yellow for Circle
                    Else
                        If effCount > 0 Then
                            ' Check if any mapping is "NEW"
                            Dim isNew As Boolean: isNew = False
                            Dim k As Long
                            For k = 0 To effCount - 1
                                If UCase(Trim(effMappings(k).MapType)) = "NEW" Then
                                    isNew = True
                                    Exit For
                                End If
                            Next k

                            If isNew Then
                                ' === CASE 3: NEW ===
                                mainColor = 1   ' Red
                            Else
                                ' === CASE 2: MAPPED (NORMAL) ===
                                mainColor = 3   ' Green
                            End If
                        Else
                            ' === CASE 1: NOT MAPPED ===
                            mainColor = 2       ' Yellow
                        End If
                    End If

                    ' 3. UPDATE LINE COLOR
                    If ent.color <> mainColor Then ent.color = mainColor

                    ' 4. GENERATE LABEL CONTENT
                    Dim baseText As String
                    baseText = n02_ACAD_Wall_Force_SAP2000.GenerateCompositeLabel(effWall, effMappings, effCount)

                    If Len(Trim$(baseText)) > 0 Then
                        ' Format: Handle gets mainColor, Body text is standard (white/bylayer)
                        Dim formattedText As String
                        formattedText = "{\C" & mainColor & ";[" & hdl & "]} " & baseText

                        Dim sp As Variant, ep As Variant
                        sp = ent.StartPoint: ep = ent.EndPoint
                        
                        Dim midX As Double: midX = (sp(0) + ep(0)) / 2
                        Dim midY As Double: midY = (sp(1) + ep(1)) / 2
                        Dim midZ As Double: midZ = (sp(2) + ep(2)) / 2

                        ' === CLEANUP: Remove old markers (circles) at this spot ===
                        ClearOldMarkersAtPosition acadDoc, midX, midY

                        ' Update or Create Label
                        Dim cacheKey As String
                        cacheKey = LCase$(hdl)

                        If labelCache.exists(cacheKey) Then
                            Dim existingLabel As Object
                            Set existingLabel = labelCache(cacheKey)
                            
                            ' Validate object validity
                            If IsObjectValid(existingLabel) Then
                                If existingLabel.ObjectName = "AcDbMText" Then
                                    Core_CAD_Plotter.UpdateFrameLabelEx existingLabel, _
                                            CDbl(sp(0)), CDbl(sp(1)), CDbl(sp(2)), _
                                            CDbl(ep(0)), CDbl(ep(1)), CDbl(ep(2)), _
                                            formattedText, dynHeight
                                    updatedCount = updatedCount + 1
                                Else
                                    existingLabel.Delete
                                    GoTo CreateNew
                                End If
                            Else
                                labelCache.Remove cacheKey
                                GoTo CreateNew
                            End If
                        Else
CreateNew:
                            Dim newLabel As Object
                            Set newLabel = Core_CAD_Plotter.PlotFrameLabelEx(acadDoc, _
                                    CDbl(sp(0)), CDbl(sp(1)), CDbl(sp(2)), _
                                    CDbl(ep(0)), CDbl(ep(1)), CDbl(ep(2)), _
                                    formattedText, dynHeight)

                            If Not newLabel Is Nothing Then
                                AttachLabelXData newLabel, hdl
                                createdCount = createdCount + 1
                            End If
                        End If

                        ' 5. DRAW CIRCLE (If needed)
                        If drawCircle Then
                            Dim center(0 To 2) As Double
                            center(0) = midX: center(1) = midY: center(2) = midZ
                            
                            ' Radius relative to text height
                            Dim rad As Double
                            rad = dynHeight / 1.8

                            Dim circ As Object
                            Set circ = acadDoc.ModelSpace.AddCircle(center, rad)
                            If Not circ Is Nothing Then
                                circ.layer = "dts_frame_label"
                                circ.color = circleColor ' Specific color for circle (Yellow=2)
                            End If
                        End If
                    End If
                End If
            End If
        End If
NextHandle:
    Next i

    acadDoc.SetVariable "CMDECHO", prevVar
    acadDoc.Regen 1

    If showMsg Then
        MsgBox "Refresh Complete!" & vbCrLf & _
               "Created: " & createdCount & vbCrLf & _
               "Updated: " & updatedCount, vbInformation
    End If
End Sub

' Helper to safely check if object is valid (not deleted)
Private Function IsObjectValid(obj As Object) As Boolean
    On Error Resume Next
    IsObjectValid = False
    If obj Is Nothing Then Exit Function
    Dim h As String
    h = obj.Handle ' Try accessing a property
    If err.number = 0 Then IsObjectValid = True
    On Error GoTo 0
End Function
' ==============================================================================
' BUTTON: DELETE LABELS (UPDATED - Deletes Text AND Circles)
' ==============================================================================
Private Sub btnDeleteLabels_Click()
    On Error GoTo ErrHandler

    ' 1. Connect to AutoCAD
    Set acadApp = Nothing: Set acadDoc = Nothing
    On Error Resume Next
    Set acadApp = GetObject(, "AutoCAD.Application")
    If acadApp Is Nothing Then
        MsgBox "AutoCAD not connected!", vbExclamation, "Error"
        Exit Sub
    End If
    On Error GoTo ErrHandler
    Set acadDoc = acadApp.ActiveDocument

    ' 2. Bring AutoCAD to Front
    BringAutoCADToFront

    ' 3. Prompt User to Select Region
    Dim ssName      As String
    ssName = "DEL_LABELS_SEL_" & Format(Now, "hhmmss")

    On Error Resume Next
    acadDoc.SelectionSets.item(ssName).Delete
    On Error GoTo ErrHandler

    Dim ss          As Object
    Set ss = acadDoc.SelectionSets.Add(ssName)

    ' Filter: Text/MText AND CIRCLES on dts_frame_label layer
    Dim gpCode(0 To 1) As Integer
    Dim dataVal(0 To 1) As Variant
    gpCode(0) = 8: dataVal(0) = "dts_frame_label"  ' Layer
    gpCode(1) = 0: dataVal(1) = "*TEXT,CIRCLE"     ' Object Types (Added CIRCLE)

    acadDoc.Utility.prompt vbCrLf & "Select labels/markers to delete (window selection): "

    On Error Resume Next
    ss.SelectOnScreen gpCode, dataVal

    If err.number <> 0 Or ss.count = 0 Then
        If Not ss Is Nothing Then ss.Delete
        Exit Sub
    End If
    On Error GoTo ErrHandler

    ' 4. Delete selected objects
    Dim deletedCount As Long: deletedCount = 0
    Dim i           As Long
    For i = ss.count - 1 To 0 Step -1
        On Error Resume Next
        ss.item(i).Delete
        If err.number = 0 Then deletedCount = deletedCount + 1
        err.Clear
    Next i

    ' 5. Cleanup
    On Error Resume Next
    ss.Delete
    acadDoc.Regen 1
    On Error GoTo ErrHandler

    Exit Sub

ErrHandler:
    MsgBox "Error deleting labels: " & err.description, vbCritical, "Error"
    On Error Resume Next
    If Not ss Is Nothing Then ss.Delete
End Sub
' ==============================================================================
' HELPER: CLEAN OLD MARKERS (Text & Circles)
' Purpose: Deletes circles and unlinked text near position to prevent artifacts
' ==============================================================================
Private Sub ClearOldMarkersAtPosition(acadDoc As Object, midX As Double, midY As Double)
    On Error Resume Next

    ' Search radius (tolerance)
    Dim tol         As Double: tol = 100
    Dim p1(0 To 2)  As Double
    Dim p2(0 To 2)  As Double

    p1(0) = midX - tol: p1(1) = midY - tol: p1(2) = 0
    p2(0) = midX + tol: p2(1) = midY + tol: p2(2) = 0

    Dim ssName      As String: ssName = "CLR_MRK_" & Format(Now, "hhmmss") & "_" & Int(Timer)
    Dim ss          As Object
    Set ss = acadDoc.SelectionSets.Add(ssName)

    ' Filter: Circles on dts_frame_label layer
    ' We only aggressively auto-delete CIRCLES here.
    ' Text is handled by the Cache logic in RefreshLabels_Core mostly.
    Dim gpCode(0 To 1) As Integer
    Dim dataVal(0 To 1) As Variant

    gpCode(0) = 8: dataVal(0) = "dts_frame_label"
    gpCode(1) = 0: dataVal(1) = "CIRCLE"

    ss.Select 1, p1, p2, gpCode, dataVal    ' 1 = Crossing Window

    Dim i As Long
    For i = 0 To ss.count - 1
        ss.item(i).Delete
    Next i

    ss.Delete
    On Error GoTo 0
End Sub

Private Sub UserForm_Initialize()
    On Error Resume Next

    ConnectSAP2000
    SapModel.SetPresentUnits (5)    '(KN_m_C)

    Set m_StoryElevations = Nothing
    m_StoryCount = 0

    ' Load saved settings to Excel Define Names
    LoadSettingsFromExcel

    mAxesCount = 0
    ReDim mSelectedAxes(0)
    UpdateAxesLabel

    Set acadApp = GetObject(, "AutoCAD.Application")
    If Not acadApp Is Nothing Then Set acadDoc = acadApp.ActiveDocument
End Sub

Private Sub LoadSettingsFromExcel()
    On Error Resume Next

    ' Load to Excel Define Names or use default
    Me.txtThickness.text = GetExcelSetting("WC_Thickness", "150,200,220,250")
    Me.txtLayers.text = GetExcelSetting("WC_Layers", "")
    Me.txtDoorWidths.text = GetExcelSetting("WC_DoorWidths", "700,800,900,1000,1200")
    Me.txtColumnWidths.text = GetExcelSetting("WC_ColumnWidths", "200,250,300,400,500")
    Me.txtAutoJoinGap.text = GetExcelSetting("WC_AutoJoinGap", "300")
    Me.txtAxisSnapDistance.text = GetExcelSetting("WC_AxisSnapDistance", "50")
    Me.txtAngleTol.text = GetExcelSetting("WC_AngleTol", "20")
    Me.txtExtendMult.text = GetExcelSetting("WC_ExtendMult", "2")
    Me.txtDistTol.text = GetExcelSetting("WC_DistTol", "10")
    Me.chkAutoExtend.Value = CBool(GetExcelSetting("WC_AutoExtend", "True"))
    Me.chkIntersection.Value = CBool(GetExcelSetting("WC_Intersection", "True"))
    Me.chkBreakAtGrid.Value = CBool(GetExcelSetting("WC_BreakAtGrid", "True"))
    Me.chkExtendToGrid.Value = CBool(GetExcelSetting("WC_ExtendToGrid", "True"))

    ' New: insertion point and manual elevation/height
    If HasControl("txtInsertX") Then GetControl("txtInsertX").text = GetExcelSetting("WC_InsertX", "")
    If HasControl("txtInsertY") Then GetControl("txtInsertY").text = GetExcelSetting("WC_InsertY", "")
    If HasControl("txtManualElevation") Then GetControl("txtManualElevation").text = GetExcelSetting("WC_ManualElevation", "")
    If HasControl("txtManualHeight") Then GetControl("txtManualHeight").text = GetExcelSetting("WC_ManualHeight", "")

    ' New: load assignments listbox - stored as rows joined with "||", columns joined with "|"
    Dim loadsSerialized As String
    loadsSerialized = GetExcelSetting("WC_LoadAssignments", "")

    If HasControl("lstLoadAssignments") Then
        Dim lst     As Object: Set lst = GetControl("lstLoadAssignments")
        lst.Clear

        ' 1. Setup ListBox Columns
        lst.ColumnCount = LISTBOX_COL_COUNT
        lst.ColumnWidths = LISTBOX_COL_WIDTHS

        ' 2. Parse and Load Data
        If Len(Trim$(loadsSerialized)) > 0 Then
            Dim rows() As String
            rows = Split(loadsSerialized, "||")    ' Split rows by ||

            Dim r   As Long
            For r = 0 To UBound(rows)
                If Len(Trim$(rows(r))) = 0 Then GoTo NextRow

                Dim cols() As String
                cols = Split(rows(r), "|")    ' Split columns by |

                lst.AddItem
                ' Load data into columns safely
                If UBound(cols) >= 0 Then lst.List(lst.ListCount - 1, COL_IDX_NO) = cols(0)
                If UBound(cols) >= 1 Then lst.List(lst.ListCount - 1, COL_IDX_HANDLE) = cols(1)
                If UBound(cols) >= 2 Then lst.List(lst.ListCount - 1, COL_IDX_THICKNESS) = cols(2)
                If UBound(cols) >= 3 Then lst.List(lst.ListCount - 1, COL_IDX_PATTERN) = cols(3)
                If UBound(cols) >= 4 Then lst.List(lst.ListCount - 1, COL_IDX_VALUE) = cols(4)
                If UBound(cols) >= 5 Then lst.List(lst.ListCount - 1, COL_IDX_MAPPING) = cols(5)

NextRow:
            Next r
        End If
    End If


    ' Tooltips for textboxes and checkboxes (more detailed, multiline)
    Me.txtThickness.ControlTipText = "Enter wall thicknesses in millimetres (mm)." & vbCrLf & _
            "Comma-separated positive numbers. Example: 150,200,220."
    Me.txtLayers.ControlTipText = "Comma-separated list of layer names to process." & vbCrLf & _
            "Leave blank to process all layers. Example: WALLS,DOORS,WINDOWS."
    Me.txtDoorWidths.ControlTipText = "Enter door/window widths in millimetres (mm)." & vbCrLf & _
            "Used to help auto-join broken walls. Example: 700,800,900."
    Me.txtColumnWidths.ControlTipText = "Enter column widths in millimetres (mm)." & vbCrLf & _
            "Used to recognize columns when auto-joining. Example: 200,300,400."
    Me.txtAutoJoinGap.ControlTipText = "Auto-join walls when gap between endpoints <= this value (mm)." & vbCrLf & _
            "Set 0 to disable auto-join. Example: 300."
    Me.txtAxisSnapDistance.ControlTipText = "Maximum distance in millimetres (mm) to snap endpoints to an axis." & vbCrLf & _
            "Set 0 to disable snapping. Example: 50."
    Me.txtAngleTol.ControlTipText = "Perpendicular angle tolerance in degrees (1-45)." & vbCrLf & _
            "Smaller = stricter. Example: 20."
    Me.txtExtendMult.ControlTipText = "Extend multiplier (0.5 - 10) used for auto-extend near-perpendicular walls." & vbCrLf & _
            "1 = no extra extension. Example: 2."
    Me.txtDistTol.ControlTipText = "Distance tolerance in millimetres (mm) used for merging/comparisons." & vbCrLf & _
            "Must be >= 0. Example: 10."
    Me.chkAutoExtend.ControlTipText = "When checked, near-perpendicular walls will be auto-extended." & vbCrLf & _
            "Uses the Extend Mult value to reach intersections."
    Me.chkIntersection.ControlTipText = "When checked, centerline intersections are detected and adjusted." & vbCrLf & _
            "Useful for accurate load diagram nodes."
    Me.chkBreakAtGrid.ControlTipText = "When checked, centerlines will be split at structural grid intersections." & vbCrLf & _
            "Grid crossing points become separate segments."
            Me.chkExtendToGrid.ControlTipText = "Extend centerlines on axes to nearest grid intersections" & vbCrLf & _
    "(Only applies to lines parallel to and lying on structural axes)"

    ' Tooltips for new controls
    On Error Resume Next
    If HasControl("txtInsertX") Then
        GetControl("txtInsertX").ControlTipText = "X coordinate of insertion point for placing the diagram in AutoCAD (mm)." & vbCrLf & _
                "Use 'Pick insertion point' to set this automatically."
    End If
    If HasControl("txtInsertY") Then
        GetControl("txtInsertY").ControlTipText = "Y coordinate of insertion point for placing the diagram in AutoCAD (mm)." & vbCrLf & _
                "Use 'Pick insertion point' to set this automatically."
    End If
    If HasControl("txtManualElevation") Then
        GetControl("txtManualElevation").ControlTipText = "Optional manual elevation for the selected story (use when list value is not suitable)." & vbCrLf & _
                "Units same as model elevations (e.g., mm)."
    End If
    If HasControl("txtManualHeight") Then
        GetControl("txtManualHeight").ControlTipText = "Optional manual story height (used to convert kN/m2 to kN/m)." & vbCrLf & _
                "Enter in mm, e.g. 3000."
    End If
    If HasControl("lstLoadAssignments") Then
        GetControl("lstLoadAssignments").ControlTipText = "List of wall load assignments." & vbCrLf & _
                "Columns: No | WallType | LoadPattern | Load_kN/m2." & vbCrLf & _
                "Double-click a row or the LoadPattern/LoadValue column to edit."
    End If
    On Error GoTo 0

End Sub

Private Function GetExcelSetting(settingName As String, defaultValue As String) As String
    On Error Resume Next

    Dim wb          As Workbook
    Set wb = ThisWorkbook

    Dim nm          As Name
    Set nm = Nothing
    Set nm = wb.names(settingName)

    If Not nm Is Nothing Then
        Dim v       As Variant
        v = nm.RefersToRange.Value
        If IsArray(v) Then
            ' If a 2D array, get first cell
            GetExcelSetting = CStr(v(1, 1))
        Else
            GetExcelSetting = CStr(v)
        End If
        ' If empty, use default
        If Len(Trim(GetExcelSetting)) = 0 Then GetExcelSetting = defaultValue
    Else
        GetExcelSetting = defaultValue
    End If

    If err.number <> 0 Then
        GetExcelSetting = defaultValue
        err.Clear
    End If

    On Error GoTo 0
End Function

' Replace: SaveSettingsToExcel - extended to save new controls and listbox content
Private Sub SaveSettingsToExcel()
    On Error Resume Next

    Dim wb          As Workbook
    Set wb = ThisWorkbook

    ' Delete and create Define Names for existing settings
    SaveExcelSetting wb, "WC_Thickness", Trim(Me.txtThickness.text)
    SaveExcelSetting wb, "WC_Layers", Trim(Me.txtLayers.text)
    SaveExcelSetting wb, "WC_DoorWidths", Trim(Me.txtDoorWidths.text)
    SaveExcelSetting wb, "WC_ColumnWidths", Trim(Me.txtColumnWidths.text)
    SaveExcelSetting wb, "WC_AutoJoinGap", Trim(Me.txtAutoJoinGap.text)
    SaveExcelSetting wb, "WC_AxisSnapDistance", Trim(Me.txtAxisSnapDistance.text)
    SaveExcelSetting wb, "WC_AngleTol", Trim(Me.txtAngleTol.text)
    SaveExcelSetting wb, "WC_ExtendMult", Trim(Me.txtExtendMult.text)
    SaveExcelSetting wb, "WC_DistTol", Trim(Me.txtDistTol.text)
    SaveExcelSetting wb, "WC_AutoExtend", CStr(Me.chkAutoExtend.Value)
    SaveExcelSetting wb, "WC_Intersection", CStr(Me.chkIntersection.Value)
    SaveExcelSetting wb, "WC_BreakAtGrid", CStr(Me.chkBreakAtGrid.Value)
    SaveExcelSetting wb, "WC_ExtendToGrid", CStr(Me.chkExtendToGrid.Value)

    ' New: insertion point and manual elevation/height
    If HasControl("txtInsertX") Then SaveExcelSetting wb, "WC_InsertX", Trim(GetControl("txtInsertX").text)
    If HasControl("txtInsertY") Then SaveExcelSetting wb, "WC_InsertY", Trim(GetControl("txtInsertY").text)
    If HasControl("txtManualElevation") Then SaveExcelSetting wb, "WC_ManualElevation", Trim(GetControl("txtManualElevation").text)
    If HasControl("txtManualHeight") Then SaveExcelSetting wb, "WC_ManualHeight", Trim(GetControl("txtManualHeight").text)

    ' New: save lstLoadAssignments content as serialized string: rows separated by "||", columns by "|"
    If HasControl("lstLoadAssignments") Then
        Dim lst     As Object: Set lst = GetControl("lstLoadAssignments")

        If lst.ListCount > 0 Then
            Dim rowsArr() As String
            ReDim rowsArr(0 To lst.ListCount - 1)

            Dim r   As Long
            For r = 0 To lst.ListCount - 1
                ' Serialize 6 columns separated by |
                ' Note: Ensure Handle and Mapping columns are saved
                rowsArr(r) = CStr(lst.List(r, COL_IDX_NO)) & "|" & _
                        CStr(lst.List(r, COL_IDX_HANDLE)) & "|" & _
                        CStr(lst.List(r, COL_IDX_THICKNESS)) & "|" & _
                        CStr(lst.List(r, COL_IDX_PATTERN)) & "|" & _
                        CStr(lst.List(r, COL_IDX_VALUE)) & "|" & _
                        CStr(lst.List(r, COL_IDX_MAPPING))
            Next r

            ' Join rows by ||
            Dim serialized As String
            serialized = Join(rowsArr, "||")

            SaveExcelSetting wb, "WC_LoadAssignments", serialized
        Else
            ' Clear setting if list is empty
            SaveExcelSetting wb, "WC_LoadAssignments", ""
        End If
    End If

    On Error GoTo 0
End Sub

Private Sub SaveExcelSetting(wb As Workbook, settingName As String, settingValue As String)
    On Error Resume Next

    ' Delete existing name if present
    On Error Resume Next
    wb.names(settingName).Delete
    On Error GoTo 0

    ' Find or create settings sheet (use hidden sheet)
    Dim ws          As Worksheet
    Set ws = GetOrCreateSettingsSheet(wb)

    ' Find next empty row
    Dim lastRow     As Long
    lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row + 1
    If lastRow = 1 Then lastRow = 2  ' Skip header

    ' Ensure cell is formatted as Text to preserve commas and prevent scientific notation
    With ws.Cells(lastRow, 1)
        .NumberFormat = "@"        ' Force Text format
        .Value = settingValue      ' Write value as text
    End With

    ' Create Define Name that refers to that cell
    wb.names.Add Name:=settingName, RefersTo:="=" & ws.Name & "!$A$" & lastRow

    On Error GoTo 0
End Sub

Private Function GetOrCreateSettingsSheet(wb As Workbook) As Worksheet
    On Error Resume Next

    Dim ws          As Worksheet
    Set ws = Nothing
    Set ws = wb.Worksheets("WallConverterSettings")

    If ws Is Nothing Then
        ' Create new sheet
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count))
        ws.Name = "WallConverterSettings"
        ws.Cells(1, 1).Value = "Settings Storage (Do not edit)"
        ws.Visible = xlSheetVeryHidden  ' hide sheet
    End If

    Set GetOrCreateSettingsSheet = ws
    On Error GoTo 0
End Function

Private Sub UpdateAxesLabel()
    If mAxesCount = 0 Then
        Me.lblAxesCount.Caption = "No axes selected (optional)"
        Me.lblAxesCount.ForeColor = RGB(128, 128, 128)
    Else
        Me.lblAxesCount.Caption = mAxesCount & " axes selected"
        Me.lblAxesCount.ForeColor = RGB(0, 128, 0)
    End If
End Sub

'==================================================================
' VALIDATION
'==================================================================
Private Sub txtThickness_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidateNumericList Me.txtThickness, "Wall thickness"
End Sub

Private Sub txtDoorWidths_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidateNumericList Me.txtDoorWidths, "Door widths"
End Sub

Private Sub txtColumnWidths_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidateNumericList Me.txtColumnWidths, "Column widths"
End Sub

Private Sub txtAutoJoinGap_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Trim(Me.txtAutoJoinGap.text) = "" Then Exit Sub
    If Not IsNumeric(Me.txtAutoJoinGap.text) Or CDbl(Me.txtAutoJoinGap.text) < 0 Then
        MsgBox "Auto-join gap must be a positive number or 0!", vbExclamation, "Input Error"
        Me.txtAutoJoinGap.SetFocus
    End If
End Sub

Private Sub txtAxisSnapDistance_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Trim(Me.txtAxisSnapDistance.text) = "" Then Exit Sub
    If Not IsNumeric(Me.txtAxisSnapDistance.text) Or CDbl(Me.txtAxisSnapDistance.text) < 0 Then
        MsgBox "Axis snap distance must be a positive number or 0!", vbExclamation, "Input Error"
        Me.txtAxisSnapDistance.SetFocus
    End If
End Sub

Private Sub txtAngleTol_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidateAngleTol
End Sub

Private Sub txtExtendMult_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidateExtendMult
End Sub

Private Sub txtDistTol_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidateDistTol
End Sub

Private Sub ValidateNumericList(txtBox As MSForms.TextBox, fieldName As String)
    Dim parts() As String, i As Long, val As String
    If Trim(txtBox.text) = "" Then Exit Sub

    parts = Split(txtBox.text, ",")
    For i = 0 To UBound(parts)
        val = Trim(parts(i))
        If val <> "" Then
            If Not IsNumeric(val) Or CDbl(val) <= 0 Then
                MsgBox "Invalid " & fieldName & " value: " & val & vbCrLf & "Only positive numbers allowed!", vbExclamation, "Input Error"
                txtBox.SetFocus
                Exit Sub
            End If
        End If
    Next i
End Sub

Private Sub ValidateAngleTol()
    If Trim(Me.txtAngleTol.text) = "" Then Exit Sub
    If Not IsNumeric(Me.txtAngleTol.text) Or CDbl(Me.txtAngleTol.text) < 1 Or CDbl(Me.txtAngleTol.text) > 45 Then
        MsgBox "Angle tolerance must be between 1 and 45 degrees!", vbExclamation, "Input Error"
        Me.txtAngleTol.SetFocus
    End If
End Sub

Private Sub ValidateExtendMult()
    If Trim(Me.txtExtendMult.text) = "" Then Exit Sub
    If Not IsNumeric(Me.txtExtendMult.text) Or CDbl(Me.txtExtendMult.text) < 0.5 Or CDbl(Me.txtExtendMult.text) > 10 Then
        MsgBox "Extend multiplier must be between 0.5 and 10!", vbExclamation, "Input Error"
        Me.txtExtendMult.SetFocus
    End If
End Sub

Private Sub ValidateDistTol()
    If Trim(Me.txtDistTol.text) = "" Then Exit Sub
    If Not IsNumeric(Me.txtDistTol.text) Or CDbl(Me.txtDistTol.text) < 0 Then
        MsgBox "Distance tolerance must be a positive number!", vbExclamation, "Input Error"
        Me.txtDistTol.SetFocus
    End If
End Sub

'==================================================================
' PICK LAYER
'==================================================================
Private Sub btnPickLayer_Click()
    On Error GoTo ErrHandler

    ' Ensure we refer to the current active AutoCAD instance and document
    Set acadApp = Nothing
    Set acadDoc = Nothing
    On Error Resume Next
    Set acadApp = GetObject(, "AutoCAD.Application")
    If acadApp Is Nothing Then
        MsgBox "AutoCAD not connected!", vbExclamation, "Error"
        Exit Sub
    End If
    On Error GoTo ErrHandler

    ' Refresh active document reference so selection happens in the current drawing tab
    Set acadDoc = acadApp.ActiveDocument

    DoEvents

    On Error Resume Next
    ' Activate AutoCAD window (bring to foreground)
    AppActivate acadApp.Caption
    DoEvents
    On Error GoTo ErrHandler

    Dim selectedObj As Object, pickPoint As Variant
    On Error Resume Next
    ' Use the live active document for GetEntity to ensure selection occurs in the active drawing
    acadApp.ActiveDocument.Utility.GetEntity selectedObj, pickPoint, vbCrLf & "Pick an object to get its layer: "

    Dim pickSuccess As Boolean
    pickSuccess = (err.number = 0 And Not selectedObj Is Nothing)
    On Error GoTo ErrHandler

    If pickSuccess Then
        Dim layerName As String
        layerName = selectedObj.layer

        Dim current As String
        current = Trim(Me.txtLayers.text)
        If Len(current) = 0 Then
            Me.txtLayers.text = layerName
        ElseIf InStr(1, "," & current & ",", "," & layerName & ",") = 0 Then
            Me.txtLayers.text = current & "," & layerName
        End If
    End If

    Exit Sub

ErrHandler:
    MsgBox "Error picking layer: " & err.description, vbCritical, "Error"
End Sub

Private Sub btnClearLayers_Click()
    Me.txtLayers.text = ""
    Me.txtLayers.SetFocus
End Sub

'==================================================================
' PICK AXES
'==================================================================
Private Sub btnPickAxes_Click()
    On Error GoTo ErrHandler

    ' Ensure AutoCAD application and active document are refreshed
    Set acadApp = Nothing
    Set acadDoc = Nothing
    On Error Resume Next
    Set acadApp = GetObject(, "AutoCAD.Application")
    If acadApp Is Nothing Then
        MsgBox "AutoCAD not connected!", vbExclamation, "Error"
        Exit Sub
    End If
    On Error GoTo ErrHandler

    Set acadDoc = acadApp.ActiveDocument

    DoEvents

    On Error Resume Next
    AppActivate acadApp.Caption
    DoEvents
    On Error GoTo ErrHandler

    Dim ssName      As String
    ssName = "AXES_TEMP_" & Format(Now, "hhmmss")

    On Error Resume Next
    ' Use the active document's SelectionSets to avoid using stale document reference
    acadApp.ActiveDocument.SelectionSets.item(ssName).Delete
    On Error GoTo ErrHandler

    Dim ss          As Object
    Set ss = acadApp.ActiveDocument.SelectionSets.Add(ssName)

    acadApp.ActiveDocument.Utility.prompt vbCrLf & "Select structural axes (lines): "

    On Error Resume Next
    ss.SelectOnScreen
    If err.number <> 0 Or ss.count = 0 Then
        ss.Delete
        Exit Sub
    End If
    On Error GoTo ErrHandler

    mAxesCount = 0
    ReDim mSelectedAxes(0 To ss.count - 1)

    Dim i           As Long
    For i = 0 To ss.count - 1
        If IsLineEntity(ss.item(i)) Then
            Set mSelectedAxes(mAxesCount) = ss.item(i)
            mAxesCount = mAxesCount + 1
        End If
    Next i

    If mAxesCount > 0 Then
        ReDim Preserve mSelectedAxes(0 To mAxesCount - 1)
        'MsgBox "Selected " & mAxesCount & " axes successfully!", vbInformation, "Axes Selected"
    Else
        ReDim mSelectedAxes(0)
        MsgBox "No valid line objects found in selection!", vbExclamation, "No Axes"
    End If

    ss.Delete
    UpdateAxesLabel

    Exit Sub

ErrHandler:
    MsgBox "Error picking axes: " & err.description, vbCritical, "Error"
    On Error Resume Next
    If Not ss Is Nothing Then ss.Delete
    On Error GoTo 0
End Sub

'==================================================================
' CLEAR AXES BUTTON
'==================================================================
Private Sub btnClearAxes_Click()
    If mAxesCount = 0 Then
        MsgBox "No axes to clear.", vbInformation, "Info"
        Exit Sub
    End If

    Dim result      As VbMsgBoxResult
    result = MsgBox("Clear all " & mAxesCount & " selected axes?", vbQuestion + vbYesNo, "Confirm Clear")

    If result = vbYes Then
        mAxesCount = 0
        ReDim mSelectedAxes(0)
        UpdateAxesLabel
        MsgBox "All selected axes have been cleared.", vbInformation, "Axes Cleared"
    End If
End Sub

Private Function IsLineEntity(ent As Object) As Boolean
    On Error Resume Next
    Dim tName       As String
    tName = LCase(ent.ObjectName)
    IsLineEntity = (tName = "acdbline")
    If err.number <> 0 Then IsLineEntity = False
    On Error GoTo 0
End Function

'==================================================================
' OK BUTTON
'==================================================================
Private Sub btnOK_Click()
    On Error GoTo ErrHandler

    If Trim(Me.txtThickness.text) = "" Then
        MsgBox "Please enter at least one wall thickness value!", vbExclamation, "Required Field"
        Me.txtThickness.SetFocus
        Exit Sub
    End If

    ValidateNumericList Me.txtThickness, "Wall thickness"
    ValidateNumericList Me.txtDoorWidths, "Door widths"
    ValidateNumericList Me.txtColumnWidths, "Column widths"
    ValidateAngleTol
    ValidateExtendMult
    ValidateDistTol

    ' Validate auto-join gap
    If Trim(Me.txtAutoJoinGap.text) <> "" Then
        If Not IsNumeric(Me.txtAutoJoinGap.text) Or CDbl(Me.txtAutoJoinGap.text) < 0 Then
            MsgBox "Auto-join gap must be a positive number or 0!", vbExclamation, "Input Error"
            Me.txtAutoJoinGap.SetFocus
            Exit Sub
        End If
    End If

    ' Validate axis snap distance
    If Trim(Me.txtAxisSnapDistance.text) <> "" Then
        If Not IsNumeric(Me.txtAxisSnapDistance.text) Or CDbl(Me.txtAxisSnapDistance.text) < 0 Then
            MsgBox "Axis snap distance must be a positive number or 0!", vbExclamation, "Input Error"
            Me.txtAxisSnapDistance.SetFocus
            Exit Sub
        End If
    End If

    ' Save settings before processing
    SaveSettingsToExcel

    DoEvents

    ' Call processing function
    ProcessWallConversion _
            Trim(Me.txtThickness.text), _
            Trim(Me.txtLayers.text), _
            Trim(Me.txtDoorWidths.text), _
            Trim(Me.txtColumnWidths.text), _
            Trim(Me.txtAutoJoinGap.text), _
            Trim(Me.txtAxisSnapDistance.text), _
            CDbl(Me.txtAngleTol.text), _
            CDbl(Me.txtExtendMult.text), _
            CDbl(Me.txtDistTol.text), _
            Me.chkAutoExtend.Value, _
            Me.chkIntersection.Value, _
            Me.chkBreakAtGrid.Value, _
            Me.chkExtendToGrid.Value, _
            mSelectedAxes, _
            mAxesCount

    Exit Sub

ErrHandler:
    MsgBox "Processing failed: " & err.description, vbCritical, "Wall Converter Error"
End Sub

'==================================================================
' DEFAULTS BUTTON - set default values and save to Excel
'==================================================================
Private Sub btnDefaults_Click()
    On Error GoTo ErrHandler

    ' Set default values (must be strings with commas where appropriate)
    Me.txtThickness.text = "150,200,220,250"
    Me.txtLayers.text = ""
    Me.txtDoorWidths.text = "700,800,900,1000,1200"
    Me.txtColumnWidths.text = "200,250,300,400,500"
    Me.txtAutoJoinGap.text = "300"
    Me.txtAxisSnapDistance.text = "50"
    Me.txtAngleTol.text = "20"
    Me.txtExtendMult.text = "2"
    Me.txtDistTol.text = "10"
    Me.chkAutoExtend.Value = True
    Me.chkIntersection.Value = True
    Me.chkBreakAtGrid.Value = False

    ' Save defaults to Excel so next load uses them
    SaveSettingsToExcel

    MsgBox "Default settings have been applied and saved.", vbInformation, "Defaults Applied"
    Exit Sub
ErrHandler:
    MsgBox "Failed to apply defaults: " & err.description, vbCritical, "Error"
End Sub

'==================================================================
' CANCEL
'==================================================================
Private Sub btnCancel_Click()
    SaveSettingsToExcel  ' Save settings even when cancel
    Unload Me
End Sub


Private Sub btnLoadStoryList_Click()
    On Error GoTo ErrHandler

    If Not HasControl("lstStoryInfo") Then Exit Sub
    Dim lst         As Object: Set lst = GetControl("lstStoryInfo")

    lst.Clear
    lst.ColumnCount = 4
    lst.ColumnWidths = "20;45;35;35"

    ' Try to get story data from M12_DrawColumnAreas first (may return different formats)
    Dim rawDict     As Object
    Set rawDict = Nothing
    On Error Resume Next
    Set rawDict = M12_DrawColumnAreas.GetGridlineElevations()
    On Error GoTo 0

    ' Fallback to existing function if nothing returned
    If rawDict Is Nothing Or rawDict.count = 0 Then
        Set rawDict = GetStoriesFromSAP()
    End If

    If rawDict Is Nothing Or rawDict.count = 0 Then
        MsgBox "No stories found in SAP2000 model!", vbExclamation, "No Data"
        Exit Sub
    End If

    ' Normalize into name->elevation dictionary (m_StoryElevations)
    Dim normDict    As Object
    Set normDict = CreateObject("Scripting.Dictionary")

    Dim k           As Variant
    For Each k In rawDict.keys
        Dim v       As Variant
        v = rawDict(k)

        Dim nm      As String
        Dim elev    As Double
        Dim height  As Double
        nm = ""
        elev = 0
        height = 0

        ' Case: value is a dictionary/object with fields
        If Not IsEmpty(v) Then
            Dim tName As String
            tName = TypeName(v)
        Else
            tName = ""
        End If

        If LCase$(tName) = "dictionary" Or LCase$(tName) = "scripting.dictionary" Then
            ' value contains details
            On Error Resume Next
            If Not IsError(v("Name")) Then nm = CStr(v("Name"))
            If Not IsError(v("Elevation")) Then elev = CDbl(v("Elevation"))
            If Not IsError(v("Height")) Then height = CDbl(v("Height"))
            On Error GoTo 0

            If Trim$(nm) = "" Then
                ' maybe key is the name
                If Not IsNumeric(k) Then nm = CStr(k)
            End If
        Else
            ' value is not dictionary
            If IsNumeric(k) And Not IsNumeric(v) Then
                ' key = elevation, value = name
                elev = CDbl(k)
                nm = CStr(v)
            ElseIf Not IsNumeric(k) And IsNumeric(v) Then
                ' key = name, value = elevation
                nm = CStr(k)
                elev = CDbl(v)
            Else
                ' both are strings or unknown: try to parse numeric parts
                If IsNumeric(CDbl(val(k))) And Len(Trim$(CStr(v))) > 0 Then
                    elev = CDbl(val(k))
                    nm = CStr(v)
                ElseIf IsNumeric(CDbl(val(v))) And Len(Trim$(CStr(k))) > 0 Then
                    elev = CDbl(val(v))
                    nm = CStr(k)
                Else
                    ' fallback: treat key as name and value as elevation if possible, else generate name
                    nm = CStr(k)
                    If IsNumeric(CDbl(val(v))) Then elev = CDbl(val(v)) Else elev = 0
                End If
            End If
        End If

        If Trim$(nm) = "" Then nm = "Story_" & Format(elev, "0.0")

        ' Ensure unique story names
        Dim baseName As String: baseName = nm
        Dim seq     As Long: seq = 1
        Do While normDict.exists(nm)
            nm = baseName & "_" & CStr(seq)
            seq = seq + 1
        Loop

        normDict.Add nm, elev
    Next k

    ' Populate module-level m_StoryElevations and m_StoryCount for other helpers
    Set m_StoryElevations = CreateObject("Scripting.Dictionary")
    Dim key         As Variant
    For Each key In normDict.keys
        m_StoryElevations.Add CStr(key), CDbl(normDict(key))
    Next key
    m_StoryCount = m_StoryElevations.count

    If m_StoryCount = 0 Then
        MsgBox "No stories found after normalization!", vbExclamation, "No Data"
        Exit Sub
    End If

    ' Build sorted array (ascending elevation) and compute heights
    Dim keysArr     As Variant
    keysArr = m_StoryElevations.keys

    Dim tmpArr()    As Variant
    ReDim tmpArr(0 To UBound(keysArr))
    Dim i           As Long
    For i = 0 To UBound(keysArr)
        tmpArr(i) = Array(CStr(keysArr(i)), CDbl(m_StoryElevations(keysArr(i))))
    Next i

    ' Sort by elevation ascending
    Dim j As Long, p As Long
    For j = 0 To UBound(tmpArr) - 1
        For p = j + 1 To UBound(tmpArr)
            If CDbl(tmpArr(j)(1)) > CDbl(tmpArr(p)(1)) Then
                Dim tmp As Variant: tmp = tmpArr(j): tmpArr(j) = tmpArr(p): tmpArr(p) = tmp
            End If
        Next p
    Next j

    ' Compute heights (distance to next elevation), last = default 3000
    Dim defaultHeight As Double: defaultHeight = 3000
    Dim heights()   As Double
    ReDim heights(0 To UBound(tmpArr))
    For i = 0 To UBound(tmpArr)
        If i < UBound(tmpArr) Then
            heights(i) = CDbl(tmpArr(i + 1)(1)) - CDbl(tmpArr(i)(1))
            If heights(i) <= 0 Then heights(i) = defaultHeight
        Else
            heights(i) = defaultHeight
        End If
    Next i

    ' Populate listbox (No | Name | Elevation | Height)
    Dim idx         As Long: idx = 1
    For i = 0 To UBound(tmpArr)
        lst.AddItem
        lst.List(lst.ListCount - 1, 0) = CStr(idx)
        lst.List(lst.ListCount - 1, 1) = CStr(tmpArr(i)(0))
        lst.List(lst.ListCount - 1, 2) = Format(CDbl(tmpArr(i)(1)), "0.00")
        lst.List(lst.ListCount - 1, 3) = Format(heights(i), "0.00")
        idx = idx + 1
    Next i

    'MsgBox "Loaded " & lst.ListCount & " stories from SAP2000", vbInformation, "Load Complete"
    Exit Sub

ErrHandler:
    MsgBox "Error loading stories: " & err.description, vbCritical, "Error"
End Sub
' Populate manual elevation/height when a story is clicked in the listbox
Private Sub lstStoryInfo_Click()
    On Error Resume Next
    If Not HasControl("lstStoryInfo") Then Exit Sub
    Dim lst         As Object: Set lst = GetControl("lstStoryInfo")
    If lst.ListIndex < 0 Then Exit Sub

    ' Read elevation and height from the selected list row (columns: 0=No,1=Name,2=Elevation,3=Height)
    Dim selElev As String, selHeight As String
    selElev = CStr(lst.List(lst.ListIndex, 2))
    selHeight = CStr(lst.List(lst.ListIndex, 3))

    ' Populate manual textboxes if they exist
    If HasControl("txtManualElevation") Then
        On Error Resume Next
        GetControl("txtManualElevation").text = selElev
        On Error GoTo 0
    End If

    If HasControl("txtManualHeight") Then
        On Error Resume Next
        GetControl("txtManualHeight").text = selHeight
        On Error GoTo 0
    End If
End Sub
' Get stories from SAP2000
Private Function GetStoriesFromSAP() As Object
    On Error Resume Next

    ' ? FIX: Single declaration
    Dim tempDict    As Object
    Set tempDict = CreateObject("Scripting.Dictionary")
    Set GetStoriesFromSAP = tempDict

    ' 1) Try M12_DrawColumnAreas first
    On Error Resume Next
    Dim gridElevDict As Object
    Set gridElevDict = M12_DrawColumnAreas.GetGridlineElevations()
    On Error GoTo 0

    If Not gridElevDict Is Nothing Then
        If gridElevDict.count > 0 Then
            Set GetStoriesFromSAP = gridElevDict

            ' Update module-level storage
            If m_StoryElevations Is Nothing Then Set m_StoryElevations = CreateObject("Scripting.Dictionary")
            m_StoryElevations.RemoveAll

            Dim key As Variant
            For Each key In gridElevDict.keys
                m_StoryElevations.Add CStr(key), CDbl(gridElevDict(key))
            Next key

            m_StoryCount = m_StoryElevations.count
            Exit Function
        End If
    End If

    ' 2) Fallback: Read from SAP Points
    Dim SapModel    As Object
    Set SapModel = GetSAPModel()
    If SapModel Is Nothing Then Exit Function

    Dim numPoints   As Long
    Dim pointNames() As String

    If SapModel.pointObj.GetNameList(numPoints, pointNames) <> 0 Then Exit Function

    ' Collect unique elevations
    Dim elevations  As Object
    Set elevations = CreateObject("Scripting.Dictionary")

    Dim idx         As Long
    For idx = 0 To numPoints - 1
        Dim px As Double, py As Double, pz As Double
        If SapModel.pointObj.GetCoordCartesian(pointNames(idx), px, py, pz) = 0 Then
            ' Check if elevation already exists (10mm tolerance)
            Dim found As Boolean
            found = False

            Dim ek  As Variant
            For Each ek In elevations.keys
                If Abs(CDbl(elevations(ek)) - pz) < 10 Then
                    found = True
                    Exit For
                End If
            Next ek

            If Not found Then
                Dim keyName As String
                keyName = "Story_" & Format(pz, "0.00")
                elevations.Add keyName, pz
            End If
        End If
    Next idx

    ' Convert to result dictionary
    If elevations.count > 0 Then
        ' Sort elevations
        Dim elevKeys As Variant
        elevKeys = elevations.keys

        Dim elevArr() As Variant
        ReDim elevArr(0 To elevations.count - 1)

        For idx = 0 To elevations.count - 1
            elevArr(idx) = Array(CStr(elevKeys(idx)), CDbl(elevations(elevKeys(idx))))
        Next idx

        ' Sort ascending
        Dim i As Long, j As Long
        For i = 0 To UBound(elevArr) - 1
            For j = i + 1 To UBound(elevArr)
                If elevArr(i)(1) > elevArr(j)(1) Then
                    Dim tmp As Variant
                    tmp = elevArr(i)
                    elevArr(i) = elevArr(j)
                    elevArr(j) = tmp
                End If
            Next j
        Next i

        ' Build result
        Dim defaultHeight As Double
        defaultHeight = 3000

        For idx = 0 To UBound(elevArr)
            Dim storyName As String
            Dim storyElev As Double
            Dim storyHeight As Double

            storyName = CStr(elevArr(idx)(0))
            storyElev = CDbl(elevArr(idx)(1))

            If idx < UBound(elevArr) Then
                storyHeight = CDbl(elevArr(idx + 1)(1)) - storyElev
                If storyHeight <= 0 Then storyHeight = defaultHeight
            Else
                storyHeight = defaultHeight
            End If

            Dim info As Object
            Set info = CreateObject("Scripting.Dictionary")
            info.Add "Name", storyName
            info.Add "Elevation", storyElev
            info.Add "Height", storyHeight
            info.Add "Index", idx + 1

            tempDict.Add storyName, info
        Next idx

        Set GetStoriesFromSAP = tempDict
    End If

    On Error GoTo 0
End Function
' Replace: GetStoryNameByIndex - use m_StoryElevations and return top-to-bottom name
Private Function GetStoryNameByIndex(storyIndex As Long) As String
    On Error Resume Next
    GetStoryNameByIndex = CStr(storyIndex)  ' Default to index string if not found

    ' Ensure m_StoryElevations populated
    If m_StoryElevations Is Nothing Or m_StoryElevations.count = 0 Then
        Dim tmp     As Object
        Set tmp = GetStoriesFromSAP()
        If tmp Is Nothing Then Exit Function
        Set m_StoryElevations = CreateObject("Scripting.Dictionary")
        Dim k       As Variant
        For Each k In tmp.keys
            m_StoryElevations.Add k, tmp(k)("Elevation")
        Next k
        m_StoryCount = m_StoryElevations.count
    End If

    If m_StoryElevations.count = 0 Then Exit Function

    ' Build sorted list of keys by elevation descending (highest first)
    Dim keys        As Variant
    keys = m_StoryElevations.keys

    Dim arr()       As Variant
    ReDim arr(0 To UBound(keys))
    Dim i           As Long
    For i = 0 To UBound(keys)
        arr(i) = Array(keys(i), CDbl(m_StoryElevations(keys(i))))
    Next i

    Dim j As Long, p As Long
    For j = 0 To UBound(arr) - 1
        For p = j + 1 To UBound(arr)
            If arr(j)(1) < arr(p)(1) Then
                Dim tmp As Variant: tmp = arr(j): arr(j) = arr(p): arr(p) = tmp
            End If
        Next p
    Next j

    ' storyIndex expected 1-based (top to bottom)
    If storyIndex >= 1 And storyIndex <= UBound(arr) + 1 Then
        GetStoryNameByIndex = CStr(arr(storyIndex - 1)(0))
    End If
End Function
' Replace: GetStoryIndexByName - return 1-based top-to-bottom index given a story name
Private Function GetStoryIndexByName(storyName As String) As Long
    On Error Resume Next
    GetStoryIndexByName = 0

    If Trim$(storyName) = "" Then Exit Function

    ' Ensure m_StoryElevations populated
    If m_StoryElevations Is Nothing Or m_StoryElevations.count = 0 Then
        Dim tmp     As Object
        Set tmp = GetStoriesFromSAP()
        If tmp Is Nothing Then Exit Function
        Set m_StoryElevations = CreateObject("Scripting.Dictionary")
        Dim k       As Variant
        For Each k In tmp.keys
            m_StoryElevations.Add k, tmp(k)("Elevation")
        Next k
        m_StoryCount = m_StoryElevations.count
    End If

    If m_StoryElevations.count = 0 Then Exit Function

    ' Build sorted list of keys by elevation descending
    Dim keys        As Variant
    keys = m_StoryElevations.keys

    Dim arr()       As Variant
    ReDim arr(0 To UBound(keys))
    Dim i           As Long
    For i = 0 To UBound(keys)
        arr(i) = Array(keys(i), CDbl(m_StoryElevations(keys(i))))
    Next i

    Dim j As Long, p As Long
    For j = 0 To UBound(arr) - 1
        For p = j + 1 To UBound(arr)
            If arr(j)(1) < arr(p)(1) Then
                Dim tmp As Variant: tmp = arr(j): arr(j) = arr(p): arr(p) = tmp
            End If
        Next p
    Next j

    ' Search for name
    For i = 0 To UBound(arr)
        If CStr(arr(i)(0)) = storyName Then
            GetStoryIndexByName = i + 1
            Exit Function
        End If
    Next i
End Function
' Pick insertion point button
Private Sub btnPickInsertionPoint_Click()
    On Error GoTo ErrHandler

    Dim acadApp As Object, acadDoc As Object
    Set acadApp = GetObject(, "AutoCAD.Application")

    If acadApp Is Nothing Then
        MsgBox "AutoCAD not connected!", vbExclamation, "Error"
        Exit Sub
    End If

    Set acadDoc = acadApp.ActiveDocument

    DoEvents

    On Error Resume Next
    AppActivate acadApp.Caption
    DoEvents
    On Error GoTo ErrHandler

    Dim pt          As Variant
    pt = acadDoc.Utility.GetPoint(, vbCrLf & "Pick model origin point in AutoCAD: ")

    If IsEmpty(pt) Then
        Exit Sub
    End If

    ' Store insertion point
    If HasControl("txtInsertX") Then GetControl("txtInsertX").text = Format(pt(0), "0.00")
    If HasControl("txtInsertY") Then GetControl("txtInsertY").text = Format(pt(1), "0.00")


    'MsgBox "Insertion point set: (" & Format(pt(0), "0.00") & ", " & Format(pt(1), "0.00") & ")", _
     vbInformation , "Point Set"

    Exit Sub

ErrHandler:
    MsgBox "Error picking point: " & err.description, vbCritical, "Error"
End Sub

' ==========================================================================================
' BUTTON: LOAD WALL TYPES (OPTIMIZED & CORRECTED)
' Logic:
' 1. Bring AutoCAD to front.
' 2. Prompt user to Select walls immediately (using AutoCAD Layer/Type Filter).
' 3. If user presses Enter (no selection), automatically Select All walls on the layer.
' 4. Scan for Load Texts internally.
' 5. Run Nearest Neighbor (Voronoi) logic on the selected walls -> Update XData.
' 6. Refresh Labels and Populate ListBox.
' ==========================================================================================
Private Sub btnLoadWallTypes_Click()
    On Error GoTo ErrHandler

    ' 1. Connect to AutoCAD
    Dim acadApp As Object, acadDoc As Object
    Set acadApp = Nothing: Set acadDoc = Nothing
    On Error Resume Next
    Set acadApp = GetObject(, "AutoCAD.Application")
    If acadApp Is Nothing Then Exit Sub
    Set acadDoc = acadApp.ActiveDocument

    ' 2. Bring AutoCAD Window to Front (Important!)
    BringAutoCADToFront
    
    ' 3. USER SELECTION (First Step)
    ' We use AutoCAD filters here to prevent selecting wrong objects
    Dim ssNameSel   As String: ssNameSel = "WALL_USER_SEL"
    Dim ssWalls     As Object
    
    On Error Resume Next
    acadDoc.SelectionSets.item(ssNameSel).Delete
    On Error GoTo 0
    Set ssWalls = acadDoc.SelectionSets.Add(ssNameSel)

    ' Define Filter: Only Lines on Layer "DTS_WALL_DIAGRAM"
    Dim gpCode(0 To 1) As Integer
    Dim dataVal(0 To 1) As Variant
    gpCode(0) = 8: dataVal(0) = "DTS_WALL_DIAGRAM" ' Layer Name
    gpCode(1) = 0: dataVal(1) = "LINE"             ' Object Type

    ' Prompt User
    On Error Resume Next
    acadDoc.Utility.prompt vbCrLf & "Select walls to Update & Load (Press Enter for ALL): "
    
    ' This activates the crosshair immediately with the filter applied
    ssWalls.SelectOnScreen gpCode, dataVal
    
    ' Fallback: If user hits Enter (Count=0), Select ALL valid walls programmatically
    If ssWalls.count = 0 Then
        ssWalls.Select 5, , , gpCode, dataVal ' 5 = acSelectionSetAll
    End If
    On Error GoTo ErrHandler

    If ssWalls.count = 0 Then
        MsgBox "No walls found on layer DTS_WALL_DIAGRAM.", vbExclamation
        Exit Sub
    End If

    ' 4. COLLECT LOAD TEXTS (Internal Scan)
    ' We only do this AFTER we have a valid wall selection
    Dim loadTexts As Collection
    Set loadTexts = New Collection
    
    Dim ssTexts As Object
    Dim ssNameT As String: ssNameT = "TEXTS_TEMP_LOAD"
    On Error Resume Next
    acadDoc.SelectionSets.item(ssNameT).Delete
    On Error GoTo 0
    Set ssTexts = acadDoc.SelectionSets.Add(ssNameT)
    
    Dim gpCodeT(0 To 1) As Integer: Dim dataValT(0 To 1) As Variant
    gpCodeT(0) = 8: dataValT(0) = "dts_wall_loading"
    gpCodeT(1) = 0: dataValT(1) = "*TEXT"
    
    On Error Resume Next
    ssTexts.Select 5, , , gpCodeT, dataValT ' Scan all texts in drawing
    
    ' Parse Text Data
    Dim i As Long
    For i = 0 To ssTexts.count - 1
        Dim ent As Object: Set ent = ssTexts.item(i)
        Dim txtStr As String: txtStr = ""
        
        On Error Resume Next
        txtStr = ent.TextString
        If txtStr = "" Then txtStr = ent.text
        
        Dim pt As Variant: pt = Empty
        pt = ent.insertionPoint
        If IsEmpty(pt) Then pt = ent.TextAlignmentPoint
        If IsEmpty(pt) Then pt = ent.Coordinates
        On Error GoTo 0
        
        If Len(Trim$(txtStr)) > 0 And IsArray(pt) Then
            Dim pat As String, val As Double
            If ParseLoadText(txtStr, pat, val) Then
                Dim lt As Object: Set lt = CreateObject("Scripting.Dictionary")
                lt.Add "Pattern", pat
                lt.Add "Value", val
                lt.Add "X", CDbl(pt(0))
                lt.Add "Y", CDbl(pt(1))
                loadTexts.Add lt
            End If
        End If
    Next i
    ssTexts.Delete

    ' 5. PROCESS WALLS (Calculation & ListBox Update)
    If Not HasControl("lstLoadAssignments") Then Exit Sub
    Dim lst As Object: Set lst = GetControl("lstLoadAssignments")
    
    ' Clear ListBox
    lst.Clear
    If lst.ColumnCount = 0 Then
        lst.ColumnCount = LISTBOX_COL_COUNT
        lst.ColumnWidths = LISTBOX_COL_WIDTHS
    End If

    Dim rowIdx As Long: rowIdx = 1
    Dim updatedHandles() As String
    Dim updateCount As Long: updateCount = 0
    ReDim updatedHandles(0 To ssWalls.count - 1)

    ' Loop through selected walls
    For i = 0 To ssWalls.count - 1
        Set ent = ssWalls.item(i)
        Dim hWall As String: hWall = CStr(ent.Handle)
        
        ' === 5a. NEAREST NEIGHBOR LOGIC (Voronoi) ===
        If loadTexts.count > 0 Then
            Dim sp As Variant, ep As Variant
            sp = ent.StartPoint: ep = ent.EndPoint
            Dim midX As Double: midX = (sp(0) + ep(0)) / 2
            Dim midY As Double: midY = (sp(1) + ep(1)) / 2
            
            Dim bestText As Object: Set bestText = Nothing
            Dim minDistSq As Double: minDistSq = -1
            
            Dim k As Long
            For k = 1 To loadTexts.count
                Dim tObj As Object: Set tObj = loadTexts(k)
                Dim d2 As Double
                d2 = (midX - CDbl(tObj("X"))) ^ 2 + (midY - CDbl(tObj("Y"))) ^ 2
                
                If minDistSq < 0 Or d2 < minDistSq Then
                    minDistSq = d2
                    Set bestText = tObj
                End If
            Next k
            
            ' Assign nearest text to wall
            If Not bestText Is Nothing Then
                UpdateEntityLoadXData acadDoc, hWall, CStr(bestText("Pattern")), CDbl(bestText("Value"))
                updatedHandles(updateCount) = hWall
                updateCount = updateCount + 1
            End If
        End If
        
        ' === 5b. ADD TO LISTBOX ===
        Call AddWallRowFromEntity(ent, lst, rowIdx)
        rowIdx = rowIdx + 1
    Next i
    
    ' Clean up selection set
    ssWalls.Delete

    ' 6. REFRESH LABELS (Only for updated entities)
    If updateCount > 0 Then
        ReDim Preserve updatedHandles(0 To updateCount - 1)
        RefreshLabels_Core acadDoc, updatedHandles, False
    End If

    Exit Sub

ErrHandler:
    Debug.Print "Error LoadWalls: " & err.description
    On Error Resume Next
    If Not ssWalls Is Nothing Then ssWalls.Delete
End Sub

' Detect wall types from AutoCAD
Private Function DetectWallTypesFromCAD() As Collection
    On Error Resume Next

    Set DetectWallTypesFromCAD = New Collection

    Dim acadApp As Object, acadDoc As Object
    Set acadApp = GetObject(, "AutoCAD.Application")

    If acadApp Is Nothing Then Exit Function

    Set acadDoc = acadApp.ActiveDocument

    Dim ms          As Object
    Set ms = acadDoc.ModelSpace

    Dim uniqueTypes As Object
    Set uniqueTypes = CreateObject("Scripting.Dictionary")

    Dim ent         As Object
    For Each ent In ms
        If LCase(ent.layer) = LCase("DTS_WALL_DIAGRAM") Then
            ' Try read XData
            Dim xdType As Variant, xdVal As Variant
            On Error Resume Next
            ent.GetXData XDATA_APP, xdType, xdVal

            If Not IsEmpty(xdVal) And IsArray(xdVal) Then
                If UBound(xdVal) >= 1 Then
                    Dim Thickness As Double
                    Thickness = CDbl(xdVal(1))

                    Dim WallType As String
                    WallType = "W" & CStr(CInt(Thickness))

                    If Not uniqueTypes.exists(WallType) Then
                        uniqueTypes.Add WallType, True
                    End If
                End If
            End If
            On Error GoTo 0
        End If
    Next ent

    ' Convert to collection
    Dim k           As Variant
    For Each k In uniqueTypes.keys
        DetectWallTypesFromCAD.Add k
    Next k

    On Error GoTo 0
End Function

' ==================== ASSIGN LOADS BUTTON (FIXED: NEW Frames + Units) ====================
Private Sub btnAssignLoads_Click()
    On Error GoTo ErrHandler

    ' Validate inputs
    Dim manualElevStr As String, manualHeightStr As String
    Dim manualElev As Double, manualHeight As Double
    manualElev = 0#: manualHeight = 0#

    If HasControl("txtManualElevation") Then
        manualElevStr = Trim$(GetControl("txtManualElevation").text)
    Else
        manualElevStr = ""
    End If
    If HasControl("txtManualHeight") Then
        manualHeightStr = Trim$(GetControl("txtManualHeight").text)
    Else
        manualHeightStr = ""
    End If

    If manualElevStr = "" Or Not IsNumeric(manualElevStr) Then
        MsgBox "Please provide Manual Elevation", vbExclamation, "Missing Input"
        Exit Sub
    End If
    If manualHeightStr = "" Or Not IsNumeric(manualHeightStr) Then
        MsgBox "Please provide Manual Height", vbExclamation, "Missing Input"
        Exit Sub
    End If

    manualElev = CDbl(manualElevStr)
    manualHeight = CDbl(manualHeightStr)

    ' Build storyInfo
    Dim storyInfo   As Object
    Set storyInfo = CreateObject("Scripting.Dictionary")
    storyInfo("Name") = "Manual_Story_" & Format(manualElev, "0.00")
    storyInfo("Elevation") = manualElev
    storyInfo("Height") = manualHeight

    ' Connect to AutoCAD
    Dim acadApp As Object, acadDoc As Object
    On Error Resume Next
    Set acadApp = GetObject(, "AutoCAD.Application")
    If acadApp Is Nothing Then
        MsgBox "AutoCAD not connected!", vbExclamation, "Error"
        Exit Sub
    End If
    On Error GoTo ErrHandler
    Set acadDoc = acadApp.ActiveDocument

    ' Prompt user to select walls
    AppActivate acadApp.Caption
    DoEvents

    Dim ssName      As String
    ssName = "WALL_ASSIGN_SEL_" & Format(Now, "hhmmss")
    On Error Resume Next
    acadDoc.SelectionSets.item(ssName).Delete
    On Error GoTo ErrHandler

    Dim ss          As Object
    Set ss = acadDoc.SelectionSets.Add(ssName)

    acadDoc.Utility.prompt vbCrLf & "Select walls to assign loads (must have mapping XData): "
    On Error Resume Next
    ss.SelectOnScreen
    If err.number <> 0 Or ss.count = 0 Then
        If Not ss Is Nothing Then ss.Delete
        MsgBox "No walls selected", vbInformation, "Cancelled"
        Exit Sub
    End If
    On Error GoTo ErrHandler

    ' Connect to SAP
    ConnectSAP2000
    If SapModel Is Nothing Then
        MsgBox "SAP2000 not connected!", vbExclamation, "Error"
        ss.Delete
        Exit Sub
    End If

    ' Process each selected wall
    Dim assignedCount As Long, failedCount As Long, createdCount As Long
    assignedCount = 0
    failedCount = 0
    createdCount = 0

    Dim i           As Long
    For i = 0 To ss.count - 1
        Dim ent     As Object
        Set ent = ss.item(i)

        On Error Resume Next
        Dim layerOK As Boolean
        layerOK = (LCase$(Trim$(ent.layer)) = LCase$("DTS_WALL_DIAGRAM"))
        On Error GoTo ErrHandler

        If Not layerOK Then GoTo NextWallAssign

        Dim baseWall As WallSegmentMap
        Dim baseMappings() As MappingRecord
        Dim ovWall  As WallSegmentMap
        Dim ovMappings() As MappingRecord
        Dim hasOverride As Boolean
        Dim ovCount As Long

        Dim baseCount As Long
        baseCount = n02_ACAD_Wall_Force_SAP2000.ReadWallAllXData( _
                ent, baseWall, baseMappings, ovWall, ovMappings, hasOverride, ovCount)

        Dim useWall As WallSegmentMap
        Dim useMappings() As MappingRecord
        Dim useCount As Long

        If hasOverride And ovCount > 0 Then
            useWall = ovWall
            useMappings = ovMappings
            useCount = ovCount
        Else
            useWall = baseWall
            useMappings = baseMappings
            useCount = baseCount
        End If

        If useCount = 0 Then
Debug.Print "Wall " & ent.Handle & " has no mapping XData, skipped"
            failedCount = failedCount + 1
            GoTo NextWallAssign
        End If

        ' CRITICAL FIX: Load on line is already in kN/m, DO NOT multiply by wall height
        ' Only convert from kN/m -> kN/mm (SAP uses kN, mm, C)
        Dim loadPerMeter As Double
        loadPerMeter = useWall.LoadValue ' kN/m (already available)
        
        Dim loadPerMM As Double
        loadPerMM = loadPerMeter / 1000# ' kN/mm (SAP units)

        Dim j       As Long
        For j = 0 To useCount - 1
            Dim frameName As String
            frameName = useMappings(j).TargetFrame

            Dim distStart As Double, distEnd As Double
            distStart = useMappings(j).DistI / 1000 ' mm -> m
            distEnd = useMappings(j).DistJ / 1000

            Dim guid As String
            guid = "DTS_" & storyInfo("Name") & "_" & frameName & "_" & _
                    Format(distStart, "0.0") & "to" & Format(distEnd, "0.0")

            Dim ret As Long
            
            ' HANDLE "NEW" FRAMES
            If UCase(Trim(useMappings(j).MapType)) = "NEW" Then
                ' Create new frame in SAP
                Dim sp As Variant, ep As Variant
                sp = ent.StartPoint: ep = ent.EndPoint
                
                ' Calculate frame start/end points based on DistI/DistJ along wall
                Dim wLen As Double
                wLen = Sqr((ep(0) - sp(0)) ^ 2 + (ep(1) - sp(1)) ^ 2)
                
                Dim ratio1 As Double, ratio2 As Double
                ratio1 = useMappings(j).DistI / wLen
                ratio2 = useMappings(j).DistJ / wLen
                
                Dim x1 As Double, y1 As Double, z1 As Double
                Dim x2 As Double, y2 As Double, z2 As Double
                
                x1 = sp(0) + ratio1 * (ep(0) - sp(0))
                y1 = sp(1) + ratio1 * (ep(1) - sp(1))
                z1 = manualElev ' Use story elevation
                
                x2 = sp(0) + ratio2 * (ep(0) - sp(0))
                y2 = sp(1) + ratio2 * (ep(1) - sp(1))
                z2 = manualElev
                
                ' Add frame to SAP
                Dim newFrameName As String
                ret = SapModel.frameObj.AddByCoord(x1, y1, z1, x2, y2, z2, newFrameName, "Default")
                
                If ret = 0 Then
Debug.Print "  Created NEW frame: " & newFrameName
                    frameName = newFrameName
                    createdCount = createdCount + 1
                Else
Debug.Print "  Failed to create NEW frame, skipped"
                    failedCount = failedCount + 1
                    GoTo NextMapping
                End If
            End If
            
            ' Assign load to frame (EXISTING or NEW)
            ret = SapModel.frameObj.SetLoadDistributedWithGUID(frameName, useWall.LoadPattern, 1, 10, _
                    distStart, distEnd, loadPerMM, loadPerMM, guid, "Global", True, False)

            If ret = 0 Then
Debug.Print "  Assigned " & Format(loadPerMeter, "0.00") & " kN/m to " & frameName & _
        " [" & Format(distStart, "0.0") & " to " & Format(distEnd, "0.0") & "]"
                assignedCount = assignedCount + 1
            Else
Debug.Print "  Failed to assign to " & frameName
                failedCount = failedCount + 1
            End If

NextMapping:
        Next j

NextWallAssign:
    Next i

    ss.Delete

    ' Refresh SAP view
    On Error Resume Next
    SapModel.View.RefreshView
    On Error GoTo ErrHandler

    MsgBox "Load assignment complete!" & vbCrLf & vbCrLf & _
            "Assigned: " & assignedCount & vbCrLf & _
            "Created NEW frames: " & createdCount & vbCrLf & _
            "Failed: " & failedCount, vbInformation, "Complete"

    Exit Sub

ErrHandler:
    MsgBox "Error in btnAssignLoads: " & err.description, vbCritical, "Error"
    On Error Resume Next
    If Not ss Is Nothing Then ss.Delete
End Sub
' ==============================================================================
' BUTTON: DELETE LOADS IN CAD ONLY (Modified)
' Purpose: Reset XData to 0, Color to Yellow, Update Labels in AutoCAD only.
' ==============================================================================
Private Sub btnDeleteLoads_Click()
    On Error GoTo ErrHandler

    ' 1. Connect to AutoCAD
    Dim acadApp As Object, acadDoc As Object
    Set acadApp = Nothing: Set acadDoc = Nothing

    On Error Resume Next
    Set acadApp = GetObject(, "AutoCAD.Application")
    If acadApp Is Nothing Then
        MsgBox "AutoCAD not connected!", vbExclamation, "Error"
        Exit Sub
    End If
    On Error GoTo ErrHandler
    Set acadDoc = acadApp.ActiveDocument

    ' 2. Bring AutoCAD to Front
    BringAutoCADToFront

    ' 3. Get Selection Handles
    Dim handles     As Variant
    handles = GetSelectionHandles_OnScreen(acadDoc)

    If IsEmpty(handles) Then
        ' MsgBox "No selection.", vbInformation
        Exit Sub
    End If

    ' 4. Process Reset Loop
    Dim i           As Long
    Dim resetCount  As Long: resetCount = 0

    For i = LBound(handles) To UBound(handles)
        Dim ent     As Object
        Set ent = Nothing
        On Error Resume Next
        Set ent = acadDoc.HandleToObject(CStr(handles(i)))
        On Error GoTo 0

        If Not ent Is Nothing Then
            ' Only process DTS_WALL_DIAGRAM layer
            If LCase$(Trim$(ent.layer)) = LCase$("DTS_WALL_DIAGRAM") Then
                ResetWallEntityInCAD ent
                resetCount = resetCount + 1
            End If
        End If
    Next i

    ' 5. AUTO UPDATE LABELS & COLOR (Silent)
    RefreshLabels_Core acadDoc, handles, False

    ' 6. Regen to show changes
    acadDoc.Regen 1

    Exit Sub
ErrHandler:
    MsgBox "Error in DeleteLoads (CAD): " & err.description, vbCritical
End Sub
' ==========================================================================================
' HELPER: Update Base Load Data (Safely preserves Override indices if they exist)
' ==========================================================================================
Private Sub UpdateEntityLoadXData(acadDoc As Object, handleStr As String, newPattern As String, newValue As Double)
    On Error Resume Next

    Dim ent As Object
    Set ent = acadDoc.HandleToObject(handleStr)
    If ent Is Nothing Then Exit Sub

    ent.Application.ActiveDocument.RegisteredApplications.Add "DTS_APP"

    Dim xdType As Variant, xdVal As Variant
    ent.GetXData "DTS_APP", xdType, xdVal

    If err.number = 0 And Not IsEmpty(xdVal) And IsArray(xdVal) Then
        If UBound(xdVal) >= 4 Then
            ' Update Index 3 & 4 (Base Load)
            xdVal(3) = CStr(newPattern)
            xdVal(4) = CDbl(newValue)
            
            ' Fix Types
            xdType(3) = CInt(1000)
            xdType(4) = CInt(1040)

            ' Write back - This preserves indices 5+ (Mapping/Override) if they exist
            ent.SetXData xdType, xdVal
        End If
    End If
    On Error GoTo 0
End Sub

' Helper: Cleanup Temporary Sheet
Private Sub CleanupTempSheet()
    On Error Resume Next
    Dim ws          As Worksheet
    Set ws = Nothing
    Set ws = ThisWorkbook.Worksheets("WallLoadData")

    If Not ws Is Nothing Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If
    On Error GoTo 0
End Sub
' ==============================================================================
' BUTTON: DELETE SAP LOADS (New - Replacing btnUpdateLoads)
' Purpose: Removes wall distributed loads from SAP2000 model for the specific story.
' ==============================================================================
Private Sub btnDeleteSAPLoads_Click()
    On Error GoTo ErrHandler

    ' 1. Validate Inputs (Need Elevation/Story Name to find GUIDs)
    If Not HasControl("txtManualElevation") Then Exit Sub

    Dim manElevStr  As String: manElevStr = Trim$(GetControl("txtManualElevation").text)
    If manElevStr = "" Or Not IsNumeric(manElevStr) Then
        MsgBox "Please enter Manual Elevation to identify the story.", vbExclamation, "Missing Input"
        Exit Sub
    End If

    ' Construct Story Name used in GUIDs (Must match generation logic)
    Dim storyName   As String
    storyName = "Manual_Story_" & Format(CDbl(manElevStr), "0.00")

    ' 2. Confirm Action
    Dim ans         As VbMsgBoxResult
    ans = MsgBox("Are you sure you want to delete ALL wall loads created by this tool" & vbCrLf & _
            "for Story: " & storyName & " in SAP2000?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm SAP Deletion")
    If ans = vbNo Then Exit Sub

    ' 3. Connect to SAP2000
    ConnectSAP2000
    If SapModel Is Nothing Then
        MsgBox "SAP2000 not connected!", vbExclamation, "Error"
        Exit Sub
    End If

    ' 4. Get Load Patterns to clean up
    Dim patterns    As Collection
    Set patterns = GetLoadPatterns()

    If patterns.count = 0 Then
        ' If listbox is empty, try to define at least default DL/Dead
        Set patterns = New Collection
        patterns.Add "DL"
        patterns.Add "DEAD"
    End If

    ' 5. Call Module Function to Delete by GUID Prefix
    ' Ensure n02 module has public Sub DeleteWallLoadsForStory
    n02_ACAD_Wall_Force_SAP2000.DeleteWallLoadsForStory SapModel, storyName, patterns

    Exit Sub

ErrHandler:
    MsgBox "Error deleting SAP loads: " & err.description, vbCritical
End Sub

' ==============================================================================
' HELPER: Reset a single entity in CAD (PRESERVE THICKNESS & TYPE)
' Purpose: Clears Load (3,4) and Mapping (5+), keeps Geometry Data (1,2)
' ==============================================================================
Private Sub ResetWallEntityInCAD(ent As Object)
    On Error Resume Next

    ' 1. Read Existing Data first
    Dim xdType As Variant, xdVal As Variant
    ent.GetXData "DTS_APP", xdType, xdVal

    Dim keepThickness As Double: keepThickness = 0
    Dim keepType    As String: keepType = ""

    ' Recover existing geometry data if available
    If err.number = 0 And Not IsEmpty(xdVal) And IsArray(xdVal) Then
        If UBound(xdVal) >= 1 Then
            If IsNumeric(xdVal(1)) Then keepThickness = CDbl(xdVal(1))
        End If
        If UBound(xdVal) >= 2 Then
            If Not IsEmpty(xdVal(2)) Then keepType = CStr(xdVal(2))
        End If
    End If

    ' Fallback: If Type is empty but Thickness exists, generate Type
    If keepType = "" And keepThickness > 0 Then
        keepType = "W" & CInt(keepThickness)
    End If

    ' 2. Reset Visuals
    ent.color = 2    ' Reset to yellow (processed but no mapping)

    ' 3. Write CLEAN structure (Indices 0 to 4 only)
    Dim newType(0 To 4) As Integer
    Dim newVal(0 To 4) As Variant

    newType(0) = 1001: newVal(0) = "DTS_APP"
    newType(1) = 1040: newVal(1) = keepThickness   ' PRESERVED
    newType(2) = 1000: newVal(2) = keepType        ' PRESERVED
    newType(3) = 1000: newVal(3) = "DL"            ' Reset Pattern
    newType(4) = 1040: newVal(4) = 0#              ' Reset Value

    ent.SetXData newType, newVal

    On Error GoTo 0
End Sub

' ==========================================================================================
' EXPORT TO EXCEL (6 Columns)
' ==========================================================================================
Private Sub btnExportLoads_Click()
    On Error GoTo ErrHandler

    If Not HasControl("lstLoadAssignments") Then Exit Sub
    Dim lst         As Object: Set lst = GetControl("lstLoadAssignments")

    If lst.ListCount = 0 Then
        MsgBox "List is empty.", vbExclamation
        Exit Sub
    End If

    Dim ws          As Worksheet
    Set ws = PrepareLoadSheet()

    ws.Cells.Clear
    ws.Range("A1:F1").Value = Array("No", "Handle", "Thickness", "Pattern", "Value_kN/m2", "Mapping_Info")
    ws.Range("A1:F1").Font.Bold = True
    ws.Range("B:B").NumberFormat = "@"

    Dim r           As Long: r = 2
    Dim i           As Long

    For i = 0 To lst.ListCount - 1
        ws.Cells(r, 1).Value = lst.List(i, COL_IDX_NO)
        ws.Cells(r, 2).Value = lst.List(i, COL_IDX_HANDLE)
        ws.Cells(r, 3).Value = lst.List(i, COL_IDX_THICKNESS)
        ws.Cells(r, 4).Value = lst.List(i, COL_IDX_PATTERN)
        ws.Cells(r, 5).Value = lst.List(i, COL_IDX_VALUE)
        ws.Cells(r, 6).Value = lst.List(i, COL_IDX_MAPPING)
        r = r + 1
    Next i

    ws.Columns.AutoFit
    ws.Activate
    Exit Sub

ErrHandler:
    MsgBox "Export failed: " & err.description, vbCritical
End Sub
' ==========================================================================================
' IMPORT FROM EXCEL (UPDATED: THICKNESS DRIVEN)
' Purpose: Reads data from Excel and updates AutoCAD Entities + ListBox.
' Excel Columns:
'   A: No
'   B: Handle
'   C: Thickness (Master Data) - can be "200" or "W200"
'   D: Load Pattern
'   E: Load Value
'   F: Mapping Info (String)
' ==========================================================================================
Private Sub btnImportLoads_Click()
    On Error GoTo ErrHandler

    ' 1. Connect to AutoCAD
    Dim acadApp As Object, acadDoc As Object
    Set acadApp = Nothing: Set acadDoc = Nothing

    On Error Resume Next
    Set acadApp = GetObject(, "AutoCAD.Application")
    If acadApp Is Nothing Then
        MsgBox "AutoCAD not connected! Cannot sync import.", vbExclamation
        Exit Sub
    End If
    Set acadDoc = acadApp.ActiveDocument
    On Error GoTo ErrHandler

    ' 2. Check Worksheet
    Dim ws          As Worksheet
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("WallLoadData")
    On Error GoTo ErrHandler

    If ws Is Nothing Then
        MsgBox "Sheet 'WallLoadData' not found!", vbExclamation
        Exit Sub
    End If

    ' 3. Prepare ListBox
    If Not HasControl("lstLoadAssignments") Then Exit Sub
    Dim lst         As Object: Set lst = GetControl("lstLoadAssignments")

    lst.Clear
    lst.ColumnCount = LISTBOX_COL_COUNT
    lst.ColumnWidths = LISTBOX_COL_WIDTHS

    ' 4. Import Loop
    Dim r           As Long: r = 2
    Dim updatedCount As Long: updatedCount = 0

    ' Loop until empty Handle in Column B
    Do While Len(Trim$(CStr(ws.Cells(r, 2).Value))) > 0

        ' --- READ DATA FROM EXCEL ---
        Dim hdl     As String
        hdl = Trim$(CStr(ws.Cells(r, 2).Value))    ' Col B: Handle

        ' --- READ THICKNESS (Col C) ---
        Dim thickVal As Double
        Dim rawThick As String
        rawThick = Trim$(CStr(ws.Cells(r, 3).Value))

        ' Auto-clean input (e.g. "W200" -> "200") to ensure we get a number
        If UCase(Left(rawThick, 1)) = "W" Then rawThick = mid(rawThick, 2)

        If IsNumeric(rawThick) And Len(rawThick) > 0 Then
            thickVal = CDbl(rawThick)
        Else
            thickVal = 200    ' Default safe fallback if invalid/empty
        End If

        ' Read Load Data
        Dim pat     As String
        pat = Trim$(CStr(ws.Cells(r, 4).Value))    ' Col D: Pattern

        Dim valStr  As String
        valStr = Trim$(CStr(ws.Cells(r, 5).Value))    ' Col E: Value
        Dim valNum  As Double: valNum = 0#
        If IsNumeric(valStr) Then valNum = CDbl(valStr)

        Dim mapStr  As String
        mapStr = Trim$(CStr(ws.Cells(r, 6).Value))    ' Col F: Mapping info

        ' --- A. ADD TO LISTBOX ---
        lst.AddItem
        lst.List(lst.ListCount - 1, COL_IDX_NO) = CStr(ws.Cells(r, 1).Value)
        lst.List(lst.ListCount - 1, COL_IDX_HANDLE) = hdl

        ' Display Thickness directly (formatted as integer string)
        lst.List(lst.ListCount - 1, COL_IDX_THICKNESS) = Format(thickVal, "0")

        lst.List(lst.ListCount - 1, COL_IDX_PATTERN) = pat
        lst.List(lst.ListCount - 1, COL_IDX_VALUE) = Format(valNum, "0.00")
        lst.List(lst.ListCount - 1, COL_IDX_MAPPING) = mapStr

        ' --- B. SYNC TO AUTOCAD (Thickness Driven) ---
        ' This function updates Thickness (Index 1) and auto-generates WallType (Index 2)
        If UpdateEntityLoadSafe(acadDoc, hdl, thickVal, pat, valNum) Then

            ' Handle mapping override strings if present in Excel
            If Len(mapStr) > 0 Then
                Dim ent As Object
                Set ent = Nothing
                On Error Resume Next
                Set ent = acadDoc.HandleToObject(hdl)
                On Error GoTo ErrHandler

                If Not ent Is Nothing Then
                    ' Read current state to get Base Mappings
                    Dim baseWall As WallSegmentMap
                    Dim baseMappings() As MappingRecord
                    Dim ovWall As WallSegmentMap
                    Dim ovMappings() As MappingRecord
                    Dim hasOv As Boolean
                    Dim ovCount As Long
                    Dim baseCount As Long

                    baseCount = n02_ACAD_Wall_Force_SAP2000.ReadWallAllXData( _
                            ent, baseWall, baseMappings, ovWall, ovMappings, hasOv, ovCount)

                    ' Parse the new mapping string from Excel
                    Dim newOvMappings() As MappingRecord
                    Dim newOvCount As Long

                    ' Construct a fake label to utilize existing parser logic
                    Dim fakeLabel As String
                    ' If mapStr starts with "to ", prepend dummy load info
                    If InStr(1, LCase$(mapStr), "to ", vbTextCompare) = 1 Then
                        fakeLabel = "W" & thickVal & " " & pat & "=" & valNum & " " & mapStr
                    Else
                        fakeLabel = mapStr
                    End If

                    newOvCount = n02_ACAD_Wall_Force_SAP2000.ParseMappingLabel(fakeLabel, newOvMappings)

                    ' Write back with Override
                    ' We use baseWall as the override source to ensure Thickness consistency
                    n02_ACAD_Wall_Force_SAP2000.WriteWallCompleteXData_WithOverride _
                            ent, baseWall, baseMappings, baseCount, True, baseWall, newOvMappings, newOvCount
                End If
            End If

            updatedCount = updatedCount + 1
        End If

        r = r + 1
    Loop

    ' 5. Cleanup & Refresh Visuals
    CleanupTempSheet

    ' Refresh Visual Labels in CAD for all imported items
    If lst.ListCount > 0 Then
        Dim hArr()  As String
        ReDim hArr(0 To lst.ListCount - 1)
        Dim n       As Long
        For n = 0 To lst.ListCount - 1
            hArr(n) = CStr(lst.List(n, COL_IDX_HANDLE))
        Next n
        RefreshLabels_Core acadDoc, hArr, False
    End If

    MsgBox "Import Complete! Updated " & updatedCount & " entities.", vbInformation
    Exit Sub

ErrHandler:
    MsgBox "Import Error: " & err.description, vbCritical
End Sub

' Helper functions for Tab 2
' Get selected story info; PRIORITY: Manual Input > List Selection
Private Function GetSelectedStory() As Object
    On Error Resume Next
    Set GetSelectedStory = Nothing

    Dim storyInfo   As Object
    Set storyInfo = CreateObject("Scripting.Dictionary")

    ' --- PRIORITY 1: CHECK MANUAL INPUTS ---
    Dim hasManual   As Boolean
    hasManual = False
    Dim manElev As Double, manHeight As Double

    Dim txtE As String, txtH As String
    If HasControl("txtManualElevation") Then txtE = Trim$(GetControl("txtManualElevation").text)
    If HasControl("txtManualHeight") Then txtH = Trim$(GetControl("txtManualHeight").text)

    If txtE <> "" And IsNumeric(txtE) And txtH <> "" And IsNumeric(txtH) Then
        manElev = CDbl(txtE)
        manHeight = CDbl(txtH)
        hasManual = True
    End If

    ' --- CHECK LISTBOX SELECTION ---
    Dim lst         As Object
    Dim hasSelection As Boolean
    hasSelection = False
    Dim lstName As String, lstElev As Double, lstHeight As Double

    If HasControl("lstStoryInfo") Then
        Set lst = GetControl("lstStoryInfo")
        If lst.ListIndex >= 0 Then
            hasSelection = True
            lstName = CStr(lst.List(lst.ListIndex, 1))
            On Error Resume Next
            lstElev = CDbl(lst.List(lst.ListIndex, 2))
            lstHeight = CDbl(lst.List(lst.ListIndex, 3))
            On Error GoTo 0
        End If
    End If

    ' --- DECISION LOGIC ---

    If hasManual Then
        ' Case A: Manual input exists (Override everything)
        storyInfo("Elevation") = manElev
        storyInfo("Height") = manHeight

        If hasSelection Then
            ' If list is also selected, keep the name from list for reference
            storyInfo("Name") = lstName
        Else
            ' If only manual, generate a generic name
            storyInfo("Name") = "Manual_Story_" & Format(manElev, "0.00")
        End If

        Set GetSelectedStory = storyInfo
        Exit Function

    ElseIf hasSelection Then
        ' Case B: Only List selection exists
        storyInfo("Name") = lstName
        storyInfo("Elevation") = lstElev
        storyInfo("Height") = lstHeight

        Set GetSelectedStory = storyInfo
        Exit Function
    End If

    ' Case C: Neither exists -> Return Nothing (Caller will handle error)
End Function

Private Function GetLoadAssignments() As Object
    On Error Resume Next

    Set GetLoadAssignments = CreateObject("Scripting.Dictionary")

    If Not HasControl("lstLoadAssignments") Then Exit Function
    Dim lst         As Object: Set lst = GetControl("lstLoadAssignments")

    Dim i           As Long
    For i = 0 To lst.ListCount - 1
        Dim WallType As String
        WallType = CStr(lst.List(i, 1))

        Dim loadData As Object
        Set loadData = CreateObject("Scripting.Dictionary")
        loadData("Pattern") = CStr(lst.List(i, 2))
        loadData("Value") = CDbl(lst.List(i, 3))

        If Not GetLoadAssignments.exists(WallType) Then
            GetLoadAssignments.Add WallType, loadData
        End If
    Next i
End Function

Private Function GetLoadPatterns() As Collection
    On Error Resume Next

    Set GetLoadPatterns = New Collection

    If Not HasControl("lstLoadAssignments") Then Exit Function
    Dim lst         As Object: Set lst = GetControl("lstLoadAssignments")

    Dim added       As Object
    Set added = CreateObject("Scripting.Dictionary")

    Dim i           As Long
    For i = 0 To lst.ListCount - 1
        Dim pattern As String
        pattern = CStr(lst.List(i, 2))

        If Not added.exists(pattern) Then
            GetLoadPatterns.Add pattern
            added.Add pattern, True
        End If
    Next i
End Function

Private Function ValidateLoadInputs() As Boolean
    ValidateLoadInputs = False

    ' Check story selection
    If Not HasControl("lstStoryInfo") Then Exit Function
    Dim lstStory    As Object: Set lstStory = GetControl("lstStoryInfo")

    If lstStory.ListIndex < 0 Then
        MsgBox "Please select a story!", vbExclamation, "No Selection"
        Exit Function
    End If

    ' Check insertion point
    If HasControl("txtInsertX") And HasControl("txtInsertY") Then
        If Not IsNumeric(GetControl("txtInsertX").text) Or Not IsNumeric(GetControl("txtInsertY").text) Then
            MsgBox "Invalid insertion point coordinates!", vbExclamation, "Invalid Input"
            Exit Function
        End If
    Else
        MsgBox "Insertion point not set!", vbExclamation, "No Point"
        Exit Function
    End If

    ' Check load assignments
    If Not HasControl("lstLoadAssignments") Then Exit Function
    Dim lstLoad     As Object: Set lstLoad = GetControl("lstLoadAssignments")

    If lstLoad.ListCount = 0 Then
        MsgBox "No load assignments defined!", vbExclamation, "No Data"
        Exit Function
    End If

    ValidateLoadInputs = True
End Function

Private Function PrepareLoadSheet() As Worksheet
    On Error Resume Next

    Dim ws          As Worksheet
    Set ws = ThisWorkbook.Worksheets("WallLoadData")

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "WallLoadData"
    End If

    Set PrepareLoadSheet = ws
End Function

Private Function GetSAPModel() As Object
    On Error Resume Next

    Set GetSAPModel = Nothing

    ' Try to get existing SAP instance
    Dim SapObj      As Object
    Set SapObj = GetObject(, "CSI.SAP2000.API.SapObject")

    If Not SapObj Is Nothing Then
        Set GetSAPModel = SapObj.SapModel
    End If
End Function


' MouseDown handler for lstLoadAssignments: capture mouse position and selected row/col
Private Sub lstLoadAssignments_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    On Error Resume Next
    If Not HasControl("lstLoadAssignments") Then Exit Sub
    Dim lst         As Object: Set lst = GetControl("lstLoadAssignments")
    If lst Is Nothing Then Exit Sub

    mLA_LastMouseX = X
    mLA_LastMouseY = Y

    ' Determine row under mouse if no ListIndex selected
    Dim rowHeight   As Single
    ' Estimate row height: using font size heuristic (may vary slightly)
    rowHeight = lst.Font.Size * 15

    Dim clickedIndex As Long
    clickedIndex = lst.ListIndex
    If clickedIndex < 0 Then
        clickedIndex = Int(Y / rowHeight)
        If clickedIndex < 0 Then clickedIndex = -1
        If clickedIndex > lst.ListCount - 1 Then clickedIndex = -1
        If clickedIndex >= 0 Then lst.ListIndex = clickedIndex
    End If

    mLA_LastRowIndex = lst.ListIndex
    mLA_LastColIndex = GetColumnIndexFromX(lst, X)
End Sub

' ==========================================================================================
' LISTBOX INTERACTION: Double Click to Edit
' ==========================================================================================
Private Sub lstLoadAssignments_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    On Error Resume Next

    If Not HasControl("lstLoadAssignments") Then Exit Sub
    Dim lst         As Object: Set lst = GetControl("lstLoadAssignments")

    Dim rowIdx      As Long
    rowIdx = lst.ListIndex
    If rowIdx < 0 Then Exit Sub

    ' Calculate Column Index based on X position (using helper from previous code)
    Dim colIdx      As Long
    ' Fallback to mLA_LastColIndex if available, or calculate
    If mLA_LastColIndex >= 0 Then
        colIdx = mLA_LastColIndex
    Else
        colIdx = GetColumnIndexFromX(lst, mLA_LastMouseX)
    End If

    ' Columns: 0:No | 1:Handle | 2:Section | 3:Pattern | 4:Value | 5:Mapping

    ' Allow editing for Pattern (3), Value (4), or General Row Click
    Dim editPattern As Boolean: editPattern = False
    Dim editValue   As Boolean: editValue = False

    If colIdx = 3 Then
        editPattern = True
    ElseIf colIdx = 4 Then
        editValue = True
    Else
        ' Clicked elsewhere -> Ask to edit everything
        If MsgBox("Edit load data for this wall?", vbQuestion + vbYesNo, "Edit") = vbYes Then
            editPattern = True
            editValue = True
        Else
            Exit Sub
        End If
    End If

    ' Get Current Values
    Dim curHandle   As String: curHandle = CStr(lst.List(rowIdx, 1))
    Dim curPat      As String: curPat = CStr(lst.List(rowIdx, 3))
    Dim curVal      As String: curVal = CStr(lst.List(rowIdx, 4))

    Dim newPat      As String: newPat = curPat
    Dim newValStr   As String: newValStr = curVal

    ' Input Dialogs
    If editPattern Then
        newPat = InputBox("Enter Load Pattern:", "Edit Pattern", curPat)
        If StrPtr(newPat) = 0 Then Exit Sub    ' Cancelled
    End If

    If editValue Then
        newValStr = InputBox("Enter Load Value (kN/m2):", "Edit Value", curVal)
        If StrPtr(newValStr) = 0 Then Exit Sub    ' Cancelled
        If Not IsNumeric(newValStr) Then
            MsgBox "Invalid number!", vbExclamation
            Exit Sub
        End If
    End If

    ' UPDATE 1: ListBox UI
    lst.List(rowIdx, 3) = newPat
    lst.List(rowIdx, 4) = Format(CDbl(newValStr), "0.00")

    ' UPDATE 2: AutoCAD Entity XData (Crucial!)
    Set acadApp = GetObject(, "AutoCAD.Application")
    If Not acadApp Is Nothing Then
        Set acadDoc = acadApp.ActiveDocument
        UpdateEntityLoadXData acadDoc, curHandle, newPat, CDbl(newValStr)

        ' Refresh visual label immediately
        Dim handles(0) As String
        handles(0) = curHandle
        RefreshLabels_Core acadDoc, handles, False
    End If

End Sub

' Helper: edit a single cell (column 2 or 3) for lstLoadAssignments
Private Sub EditLoadAssignmentCell(rowIndex As Long, colIndex As Long)
    On Error GoTo ErrHandler
    If Not HasControl("lstLoadAssignments") Then Exit Sub
    Dim lst         As Object: Set lst = GetControl("lstLoadAssignments")
    If lst Is Nothing Then Exit Sub
    If rowIndex < 0 Or rowIndex > lst.ListCount - 1 Then Exit Sub
    If colIndex < 2 Or colIndex > 3 Then Exit Sub  ' only allow Edit for columns 2 and 3

    Dim currentVal  As String
    currentVal = CStr(lst.List(rowIndex, colIndex))

    Dim prompt      As String
    If colIndex = 2 Then
        prompt = "Edit Load Pattern (e.g. DEAD, LIVE):"
    Else
        prompt = "Edit Load Value (kN/m2 or numeric):"
    End If

    Dim newVal      As String
    newVal = InputBox(prompt, "Edit Load Assignment", currentVal)
    ' If user pressed Cancel, InputBox returns zero-length string? use StrPtr to detect Cancel in VBA forms? InputBox returns "" on cancel.
    If StrPtr(newVal) = 0 Then
        ' Some hosts may return pointer 0; treat as cancel
        Exit Sub
    End If
    ' Accept empty as clearing value
    lst.List(rowIndex, colIndex) = CStr(newVal)
    ' Optionally update preserved mapping or other state here (not required)

    Exit Sub
ErrHandler:
    Resume Next
End Sub

' Helper: edit both LoadPattern and LoadValue for a row sequentially
Private Sub EditLoadAssignmentRow(rowIndex As Long)
    On Error GoTo ErrHandler
    If Not HasControl("lstLoadAssignments") Then Exit Sub
    Dim lst         As Object: Set lst = GetControl("lstLoadAssignments")
    If lst Is Nothing Then Exit Sub
    If rowIndex < 0 Or rowIndex > lst.ListCount - 1 Then Exit Sub

    Dim curPattern As String, curValue As String
    curPattern = CStr(lst.List(rowIndex, 2))
    curValue = CStr(lst.List(rowIndex, 3))

    Dim newPattern  As String
    Dim newValue    As String

    newPattern = InputBox("Edit Load Pattern (e.g. DEAD, LIVE):", "Edit Load Pattern", curPattern)
    If StrPtr(newPattern) = 0 Then Exit Sub

    newValue = InputBox("Edit Load Value (kN/m2 or numeric):", "Edit Load Value", curValue)
    If StrPtr(newValue) = 0 Then Exit Sub

    lst.List(rowIndex, 2) = CStr(newPattern)
    lst.List(rowIndex, 3) = CStr(newValue)

    Exit Sub
ErrHandler:
    Resume Next
End Sub

' Utility: determine column index from x coordinate for a ListBox with ColumnWidths property
' Returns 0-based column index
Private Function GetColumnIndexFromX(lst As Object, X As Single) As Long
    On Error Resume Next
    Dim widths()    As String
    widths = Split(lst.ColumnWidths, ";")
    Dim i           As Long
    Dim total       As Double: total = 0
    For i = LBound(widths) To UBound(widths)
        total = total + CDbl(Trim$(widths(i)))
    Next i

    If total <= 0 Then
        ' fallback: estimate by equal columns
        GetColumnIndexFromX = Int((X / lst.Width) * lst.ColumnCount)
        If GetColumnIndexFromX < 0 Then GetColumnIndexFromX = 0
        If GetColumnIndexFromX > lst.ColumnCount - 1 Then GetColumnIndexFromX = lst.ColumnCount - 1
        Exit Function
    End If

    Dim acc         As Double: acc = 0
    Dim lstWidth    As Double: lstWidth = lst.Width
    For i = LBound(widths) To UBound(widths)
        acc = acc + CDbl(Trim$(widths(i)))
        If X <= (acc / total) * lstWidth Then
            GetColumnIndexFromX = i
            Exit Function
        End If
    Next i

    GetColumnIndexFromX = UBound(widths)
    If GetColumnIndexFromX < 0 Then GetColumnIndexFromX = 0
End Function
' Usage:
'   If HasControl("txtInsertX") Then Set ctrl = GetControl("txtInsertX")
' These functions are safe: they return False / Nothing when control not found.
' Helper: build load assignment dictionary from lstLoadAssignments
' Returns dictionary: key = Upper(wallType) -> value = dictionary("Pattern", "Value")
Private Function BuildLoadAssignmentDictFromList() As Object
    On Error Resume Next
    Dim dict        As Object
    Set dict = CreateObject("Scripting.Dictionary")

    If Not HasControl("lstLoadAssignments") Then
        Set BuildLoadAssignmentDictFromList = dict
        Exit Function
    End If

    Dim lst         As Object: Set lst = GetControl("lstLoadAssignments")
    If lst Is Nothing Then
        Set BuildLoadAssignmentDictFromList = dict
        Exit Function
    End If

    Dim i           As Long
    For i = 0 To lst.ListCount - 1
        Dim wt      As String
        Dim pat     As String
        Dim valStr  As String
        Dim valNum  As Double
        wt = ""
        pat = ""
        valStr = ""
        valNum = 0#

        On Error Resume Next
        wt = Trim$(CStr(lst.List(i, 1)))
        pat = Trim$(CStr(lst.List(i, 2)))
        valStr = Trim$(CStr(lst.List(i, 3)))
        On Error GoTo 0

        If wt <> "" Then
            If IsNumeric(valStr) Then
                valNum = CDbl(valStr)
            Else
                valNum = 0#
            End If

            Dim info As Object
            Set info = CreateObject("Scripting.Dictionary")
            info.Add "Pattern", pat
            info.Add "Value", valNum

            Dim key As String
            key = UCase$(wt)
            If Not dict.exists(key) Then
                dict.Add key, info
            Else
                dict(key) = info
            End If
        End If
    Next i

    Set BuildLoadAssignmentDictFromList = dict
End Function

' ==========================================================================================
' HELPER: Attach Wall Force XData (NON-DESTRUCTIVE to Mapping)
' ==========================================================================================
Private Sub AttachWallForceXData(entObj As Object, ByVal pattern As String, ByVal LoadValue As Double)
    On Error GoTo ErrHandler
    If entObj Is Nothing Then Exit Sub

    Dim roundedVal  As Double
    roundedVal = Round(CDbl(LoadValue), 4)

    ' Register application
    On Error Resume Next
    entObj.Application.ActiveDocument.RegisteredApplications.Add "DTS_APP"
    On Error GoTo ErrHandler

    ' Read existing
    Dim xdType As Variant, xdVal As Variant
    entObj.GetXData "DTS_APP", xdType, xdVal

    Dim finalType() As Integer
    Dim finalVal()  As Variant
    Dim ub          As Long

    If err.number = 0 And Not IsEmpty(xdVal) And IsArray(xdVal) Then
        ' Case: Update Existing - Preserve Mapping (Index 5+)
        ub = UBound(xdVal)
        If ub < 4 Then ub = 4    ' Ensure space for load data

        ReDim finalType(0 To ub)
        ReDim finalVal(0 To ub)

        Dim i       As Long
        For i = LBound(xdVal) To UBound(xdVal)
            finalType(i) = CInt(xdType(i))
            finalVal(i) = xdVal(i)
        Next i
    Else
        ' Case: New
        ub = 4
        ReDim finalType(0 To 4)
        ReDim finalVal(0 To 4)
        finalType(0) = 1001: finalVal(0) = "DTS_APP"
        finalType(1) = 1040: finalVal(1) = 0#    ' Default Thickness
        finalType(2) = 1000: finalVal(2) = ""    ' Default Type
    End If

    ' Update Load Data
    finalType(3) = 1000: finalVal(3) = CStr(pattern)
    finalType(4) = 1040: finalVal(4) = CDbl(roundedVal)

    entObj.SetXData finalType, finalVal

    Exit Sub
ErrHandler:
Debug.Print "AttachWallForceXData error: " & err.description
End Sub

Private Function ApplyLoadToEntity(entObj As Object, loadDict As Object) As Boolean
    On Error GoTo ErrHandler
    ApplyLoadToEntity = False
    If entObj Is Nothing Then Exit Function
    If loadDict Is Nothing Then Exit Function

    ' Only process lines on the DTS_WALL_DIAGRAM layer
    On Error Resume Next
    Dim entLayer    As String
    entLayer = LCase$(Trim$(entObj.layer))
    On Error GoTo ErrHandler

    If entLayer <> LCase$("DTS_WALL_DIAGRAM") Then Exit Function
    If Not IsLineEntity(entObj) Then Exit Function

    ' Determine wall type from existing DTS_APP XData or thickness
    Dim WallType    As String
    WallType = GetWallTypeFromEntity(entObj)
    If Len(Trim$(WallType)) = 0 Then Exit Function

    Dim key         As String
    key = UCase$(WallType)
    If Not loadDict.exists(key) Then Exit Function

    Dim info        As Object
    Set info = loadDict(key)

    Dim pat         As String
    Dim val         As Double
    On Error Resume Next
    pat = CStr(info("Pattern"))
    val = CDbl(info("Value"))
    On Error GoTo ErrHandler

    ' Attach using the central AttachWallForceXData (which rounds)
    AttachWallForceXData entObj, pat, val

    ApplyLoadToEntity = True
    Exit Function

ErrHandler:
Debug.Print "ApplyLoadToEntity ERROR: " & err.description
    ApplyLoadToEntity = False
End Function
Private Function ApplyLoadDictToSelection(acadDoc As Object, handles As Variant, loadDict As Object) As Long
    On Error GoTo ErrHandler
    ApplyLoadDictToSelection = 0
    If acadDoc Is Nothing Then Exit Function
    If IsEmpty(handles) Then Exit Function
    If loadDict Is Nothing Or loadDict.count = 0 Then Exit Function

    Dim i           As Long
    For i = LBound(handles) To UBound(handles)
        On Error Resume Next
        Dim h       As String
        h = CStr(handles(i))
        Dim ent     As Object
        Set ent = acadDoc.HandleToObject(h)
        On Error GoTo ErrHandler

        If Not ent Is Nothing Then
            If ApplyLoadToEntity(ent, loadDict) Then
                ApplyLoadDictToSelection = ApplyLoadDictToSelection + 1
            End If
        End If
    Next i

    Exit Function
ErrHandler:
Debug.Print "ApplyLoadDictToSelection ERROR: " & err.description
    Resume Next
End Function
' Returns True if present and sets patternOut and valueOut
Private Function ReadWallForceXData(entObj As Object, ByRef patternOut As String, ByRef valueOut As Double) As Boolean
    On Error Resume Next
    ReadWallForceXData = False
    patternOut = ""
    valueOut = 0#

    If entObj Is Nothing Then Exit Function

    Dim xdType As Variant, xdVal As Variant
    entObj.GetXData "DTS_APP", xdType, xdVal

    If err.number <> 0 Then
        err.Clear
        Exit Function
    End If

    If IsArray(xdVal) Then
        ' NEW Format: (0)=App, (1)=Thickness, (2)=WallType, (3)=Pattern, (4)=LoadValue
        If UBound(xdVal) >= 4 Then
            If Not IsEmpty(xdVal(3)) Then patternOut = CStr(xdVal(3))
            If IsNumeric(xdVal(4)) Then valueOut = CDbl(xdVal(4))
            ReadWallForceXData = True
        Else
Debug.Print "ReadWallForceXData: Old format (no load data)"
            ReadWallForceXData = False
        End If
    End If
End Function

' Helper: determine wall type for an entity by reading XData
' Returns wallType string (e.g., "W200") or "" if not found
Private Function GetWallTypeFromEntity(entObj As Object) As String
    On Error Resume Next
    GetWallTypeFromEntity = ""
    If entObj Is Nothing Then Exit Function

    Dim xdType As Variant, xdVal As Variant
    entObj.GetXData XDATA_APP, xdType, xdVal
    If err.number = 0 Then
        If Not IsEmpty(xdVal) And IsArray(xdVal) Then
            Dim Thickness As Double
            Thickness = 0#
            If UBound(xdVal) >= 1 Then
                If IsNumeric(xdVal(1)) Then Thickness = CDbl(xdVal(1))
            End If
            If UBound(xdVal) >= 2 Then
                If Not IsEmpty(xdVal(2)) Then
                    GetWallTypeFromEntity = CStr(xdVal(2))
                    Exit Function
                End If
            End If
            If Thickness > 0 Then
                GetWallTypeFromEntity = "W" & CStr(CInt(Thickness))
                Exit Function
            End If
        End If
    End If
    err.Clear
End Function
Private Function HasControl(ctrlName As String) As Boolean
    On Error Resume Next
    Dim c           As Object
    Set c = Nothing
    ' Try direct lookup
    Set c = Me.Controls(ctrlName)
    If err.number <> 0 Then
        err.Clear
        HasControl = False
        Exit Function
    End If
    HasControl = Not c Is Nothing
    On Error GoTo 0
End Function

Private Function GetControl(ctrlName As String) As Object
    On Error Resume Next
    Dim c           As Object
    Set c = Nothing
    Set c = Me.Controls(ctrlName)
    If err.number <> 0 Then
        err.Clear
        Set GetControl = Nothing
        Exit Function
    End If
    Set GetControl = c
    On Error GoTo 0
End Function
' Get insertion point from textboxes
Private Function GetInsertionPoint() As Variant
    On Error Resume Next

    If Not HasControl("txtInsertX") Or Not HasControl("txtInsertY") Then
        GetInsertionPoint = Empty
        Exit Function
    End If

    Dim X As String, Y As String
    X = Trim$(GetControl("txtInsertX").text)
    Y = Trim$(GetControl("txtInsertY").text)

    If Not IsNumeric(X) Or Not IsNumeric(Y) Then
        GetInsertionPoint = Empty
        Exit Function
    End If

    Dim pt(0 To 2)  As Double
    pt(0) = CDbl(X)
    pt(1) = CDbl(Y)
    pt(2) = 0

    GetInsertionPoint = pt

End Function
' ==============================================================================
' HELPER: BRING AUTOCAD TO FRONT
' ==============================================================================
Private Sub BringAutoCADToFront()
    On Error Resume Next
    If Not acadApp Is Nothing Then
        AppActivate acadApp.Caption
        DoEvents
    End If
    On Error GoTo 0
End Sub

' ==============================================================================
' HELPER: CLEAR OLD LABEL AT POSITION
' Purpose: Deletes text on 'dts_frame_label' layer near a specific point
' ==============================================================================
Private Sub ClearOldLabelAtPosition(acadDoc As Object, midPt As Variant)
    On Error Resume Next

    ' Define a search window (e.g., +/- 200mm around midpoint)
    Dim tol         As Double: tol = 200

    Dim p1(0 To 2)  As Double
    Dim p2(0 To 2)  As Double

    p1(0) = midPt(0) - tol: p1(1) = midPt(1) - tol: p1(2) = 0
    p2(0) = midPt(0) + tol: p2(1) = midPt(1) + tol: p2(2) = 0

    ' Create a temporary selection set
    Dim ssName      As String: ssName = "CLR_LBL_" & Format(Now, "hhmmss") & "_" & Int(Timer)
    Dim ss          As Object
    Set ss = acadDoc.SelectionSets.Add(ssName)

    ' Filter for Text/MText on specific layer
    Dim gpCode(0 To 1) As Integer
    Dim dataVal(0 To 1) As Variant

    gpCode(0) = 8: dataVal(0) = "dts_frame_label"    ' Layer filter
    gpCode(1) = 0: dataVal(1) = "*TEXT"           ' Object type filter (TEXT, MTEXT)

    ' Select
    ss.Select 1, p1, p2, gpCode, dataVal    ' 1 = acSelectionSetCrossing

    ' Delete found items
    Dim i           As Long
    For i = 0 To ss.count - 1
        ss.item(i).Delete
    Next i

    ' Cleanup
    ss.Delete
    On Error GoTo 0
End Sub
Private Sub ShowHandlesWithXData(acadDoc As Object, handles As Variant)
    On Error GoTo ErrHandler
    If acadDoc Is Nothing Then Exit Sub
    If IsEmpty(handles) Or Not IsArray(handles) Then
        MsgBox "No handles to inspect.", vbInformation, "No Data"
        Exit Sub
    End If

    ' ===== BUILD CACHE FOR SELECTED HANDLES ONLY =====
    Dim labelCache  As Object
    Set labelCache = BuildLabelCacheFromHandles(acadDoc, handles)

    Dim totalHandles As Long: totalHandles = 0
    Dim foundCount  As Long: foundCount = 0
    Dim createdLabels As Long: createdLabels = 0
    Dim skippedCount As Long: skippedCount = 0

    Dim i           As Long
    For i = LBound(handles) To UBound(handles)
        totalHandles = totalHandles + 1
        Dim h       As String
        h = CStr(handles(i))

        Dim ent     As Object
        Set ent = Nothing
        On Error Resume Next
        Set ent = acadDoc.HandleToObject(h)
        If err.number <> 0 Or ent Is Nothing Then
            err.Clear
            skippedCount = skippedCount + 1
            GoTo NextHandleLoop
        End If
        On Error GoTo ErrHandler

        If Not IsLineEntity(ent) Then
            skippedCount = skippedCount + 1
            GoTo NextHandleLoop
        End If

        ' Get geometry
        Dim sp As Variant, ep As Variant
        On Error Resume Next
        sp = ent.StartPoint
        ep = ent.EndPoint
        If err.number <> 0 Or Not IsArray(sp) Or Not IsArray(ep) Then
            err.Clear
            skippedCount = skippedCount + 1
            GoTo NextHandleLoop
        End If
        On Error GoTo ErrHandler

        Dim startX As Double, startY As Double, startZ As Double
        Dim endX As Double, endY As Double, endZ As Double
        startX = CDbl(sp(0)): startY = CDbl(sp(1)): startZ = CDbl(sp(2))
        endX = CDbl(ep(0)): endY = CDbl(ep(1)): endZ = CDbl(ep(2))

        ' Delete old label (from cache)
        Dim cacheKey As String
        cacheKey = LCase$(h)

        If labelCache.exists(cacheKey) Then
            On Error Resume Next
            labelCache(cacheKey).Delete
            err.Clear
            On Error GoTo ErrHandler
            labelCache.Remove cacheKey
        End If

        ' Read mapping data
        Dim mappingCount As Long
        Dim wallSeg As WallSegmentMap
        Dim mappings() As MappingRecord
        Dim labelText As String

        On Error Resume Next
        mappingCount = n02_ACAD_Wall_Force_SAP2000.ReadWallCompleteXData(ent, wallSeg, mappings)
        If err.number <> 0 Then
            err.Clear
            mappingCount = 0
        End If
        On Error GoTo ErrHandler

        If mappingCount > 0 Then
            labelText = n02_ACAD_Wall_Force_SAP2000.GenerateCompositeLabel(wallSeg, mappings, mappingCount)
            labelText = "[" & h & "] " & labelText
            foundCount = foundCount + 1
        Else
            labelText = "[" & h & "] (No mapping data)"
        End If

        ' Plot label
        If Len(Trim$(labelText)) > 0 Then
            On Error Resume Next
            Dim newLabel As Object
            Set newLabel = Core_CAD_Plotter.PlotFrameLabelEx(acadDoc, _
                    CDbl(startX), CDbl(startY), CDbl(startZ), _
                    CDbl(endX), CDbl(endY), CDbl(endZ), _
                    labelText, 80)

            If Not newLabel Is Nothing Then
                AttachLabelXData newLabel, h
                If err.number = 0 Then createdLabels = createdLabels + 1
                err.Clear
            End If
            On Error GoTo ErrHandler
        End If

NextHandleLoop:
    Next i

    On Error Resume Next
    acadDoc.Regen 1
    On Error GoTo ErrHandler

    Exit Sub

ErrHandler:
    MsgBox "Error: " & err.description, vbCritical, "Error"
End Sub
' ==============================================================================
' HELPER: Calculate Dynamic Text Height (FIXED LAYER FILTER)
' Logic: Height = 1/60 of the Longest Selected WALL (Layer: DTS_WALL_DIAGRAM)
'        Ignores Grid lines or other objects in selection
' ==============================================================================
Private Function CalculateDynamicTextHeight(acadDoc As Object, handles As Variant) As Double
    On Error Resume Next
    CalculateDynamicTextHeight = 150    ' Default fallback

    If IsEmpty(handles) Or Not IsArray(handles) Then Exit Function

    Dim maxLen      As Double: maxLen = 0
    Dim i           As Long
    Dim ent         As Object
    Dim sp As Variant, ep As Variant
    Dim currLen     As Double
    Dim layerName   As String

    ' Iterate through handles to find the longest WALL element
    For i = LBound(handles) To UBound(handles)
        Set ent = Nothing
        Set ent = acadDoc.HandleToObject(CStr(handles(i)))

        If Not ent Is Nothing Then
            ' 1. Check Object Type (Line)
            If ent.ObjectName = "AcDbLine" Then
                ' 2. CRITICAL FIX: Check Layer Name (Must be Wall Layer)
                layerName = LCase$(Trim$(ent.layer))
                
                If layerName = "dts_wall_diagram" Then
                    sp = ent.StartPoint
                    ep = ent.EndPoint
                    
                    ' Calculate Euclidean distance
                    currLen = Sqr((sp(0) - ep(0)) ^ 2 + (sp(1) - ep(1)) ^ 2 + (sp(2) - ep(2)) ^ 2)
                    
                    If currLen > maxLen Then maxLen = currLen
                End If
            End If
        End If
    Next i

    ' If no valid walls found on layer DTS_WALL_DIAGRAM, stick to default
    If maxLen <= 0 Then Exit Function

    ' Rule: 1/60 of the longest WALL
    Dim calcHeight  As Double
    calcHeight = maxLen / 60#

    ' Safety Clamps
    If calcHeight < 50 Then calcHeight = 50
    If calcHeight > 600 Then calcHeight = 600

    CalculateDynamicTextHeight = calcHeight
Debug.Print "Dynamic Text Height: " & calcHeight & " (Max Wall Len: " & Format(maxLen, "0") & ")"
End Function
' ==================== UPDATED READ FUNCTION (Robust Schema) ====================
Public Function ReadWallAllXData(entObj As Object, _
        ByRef wallSeg As WallSegmentMap, _
        ByRef baseMappings() As MappingRecord, _
        ByRef ovWallSeg As WallSegmentMap, _
        ByRef ovMappings() As MappingRecord, _
        ByRef hasOverride As Boolean, _
        ByRef ovCount As Long) As Long

    On Error Resume Next
    ReadWallAllXData = 0
    hasOverride = False
    ovCount = 0

    If entObj Is Nothing Then Exit Function

    Dim xdType As Variant, xdVal As Variant
    entObj.GetXData XDATA_APP, xdType, xdVal

    If IsEmpty(xdVal) Or Not IsArray(xdVal) Then Exit Function

    Dim maxIndex    As Long
    maxIndex = UBound(xdVal)

    ' --- READ HEADER (Standard Indices 0-4) ---
    ' Initialize Defaults
    wallSeg.Thickness = 200
    wallSeg.WallType = ""
    wallSeg.LoadPattern = "DL"
    wallSeg.LoadValue = 0

    ' Read Index 1: Thickness
    If maxIndex >= XDATA_OFFSET_THICKNESS Then
        If IsNumeric(xdVal(XDATA_OFFSET_THICKNESS)) Then wallSeg.Thickness = CDbl(xdVal(XDATA_OFFSET_THICKNESS))
    End If

    ' Read Index 2: WallType
    If maxIndex >= XDATA_OFFSET_WALLTYPE Then
        If Not IsEmpty(xdVal(XDATA_OFFSET_WALLTYPE)) Then wallSeg.WallType = CStr(xdVal(XDATA_OFFSET_WALLTYPE))
    End If

    ' Read Index 3: LoadPattern (If exists)
    If maxIndex >= XDATA_OFFSET_LOADPATTERN Then
        If Not IsEmpty(xdVal(XDATA_OFFSET_LOADPATTERN)) Then wallSeg.LoadPattern = CStr(xdVal(XDATA_OFFSET_LOADPATTERN))
    End If

    ' Read Index 4: LoadValue (If exists)
    If maxIndex >= XDATA_OFFSET_LOADVALUE Then
        If IsNumeric(xdVal(XDATA_OFFSET_LOADVALUE)) Then wallSeg.LoadValue = CDbl(xdVal(XDATA_OFFSET_LOADVALUE))
    End If

    ' --- READ MAPPINGS (Index 5+) ---
    Dim remainingSize As Long
    remainingSize = maxIndex - XDATA_OFFSET_MAPPING_START + 1

    Dim baseCount   As Long
    baseCount = 0
    Dim cursor      As Long
    cursor = XDATA_OFFSET_MAPPING_START

    If remainingSize >= XDATA_MAPPING_RECORD_SIZE Then
        Dim totalSlots As Long
        totalSlots = remainingSize
        Dim possibleCount As Long
        possibleCount = totalSlots \ XDATA_MAPPING_RECORD_SIZE

        If possibleCount > 0 Then
            ReDim baseMappings(0 To possibleCount - 1)
            Dim i   As Long
            For i = 0 To possibleCount - 1
                If cursor + 3 > maxIndex Then Exit For

                ' Basic check: TargetFrame and MapType must be strings
                If TypeName(xdVal(cursor)) = "String" And TypeName(xdVal(cursor + 1)) = "String" Then
                    baseMappings(baseCount).TargetFrame = CStr(xdVal(cursor))
                    baseMappings(baseCount).MapType = CStr(xdVal(cursor + 1))
                    baseMappings(baseCount).DistI = CDbl(xdVal(cursor + 2))
                    baseMappings(baseCount).DistJ = CDbl(xdVal(cursor + 3))
                    baseMappings(baseCount).FrameLength = baseMappings(baseCount).DistJ - baseMappings(baseCount).DistI

                    baseCount = baseCount + 1
                    cursor = cursor + XDATA_MAPPING_RECORD_SIZE
                Else
                    ' Hit override flag or end of clean mapping
                    Exit For
                End If
            Next i

            If baseCount > 0 Then ReDim Preserve baseMappings(0 To baseCount - 1)
        End If
    End If

    ReadWallAllXData = baseCount

    ' --- READ OVERRIDE (If exists after base mappings) ---
    If cursor > maxIndex Then Exit Function

    ' Override flag is Integer/Long (1071)
    If IsNumeric(xdVal(cursor)) Then
        Dim ovFlag  As Long
        ovFlag = CLng(xdVal(cursor))
        If ovFlag <> 0 Then hasOverride = True
        cursor = cursor + 1
    Else
        Exit Function
    End If

    If Not hasOverride Then Exit Function

    ' Read Override Header (WallType, Pattern, Value) - 3 fields
    ' Note: Thickness implies reuse of base thickness usually, but we need to check cursor bounds

    ' OV_WallType
    If cursor <= maxIndex Then ovWallSeg.WallType = CStr(xdVal(cursor)): cursor = cursor + 1
    ' OV_Pattern
    If cursor <= maxIndex Then ovWallSeg.LoadPattern = CStr(xdVal(cursor)): cursor = cursor + 1
    ' OV_Value
    If cursor <= maxIndex Then ovWallSeg.LoadValue = CDbl(xdVal(cursor)): cursor = cursor + 1

    ovWallSeg.Thickness = wallSeg.Thickness    ' Inherit

    ' Read OV Mappings
    Dim ovSlotCount As Long
    ovSlotCount = (maxIndex - cursor + 1)
    If ovSlotCount < XDATA_MAPPING_RECORD_SIZE Then Exit Function

    Dim maxOVCount  As Long
    maxOVCount = ovSlotCount \ XDATA_MAPPING_RECORD_SIZE
    If maxOVCount <= 0 Then Exit Function

    ReDim ovMappings(0 To maxOVCount - 1)
    Dim j           As Long
    For j = 0 To maxOVCount - 1
        If cursor + 3 > maxIndex Then Exit For
        ovMappings(ovCount).TargetFrame = CStr(xdVal(cursor))
        ovMappings(ovCount).MapType = CStr(xdVal(cursor + 1))
        ovMappings(ovCount).DistI = CDbl(xdVal(cursor + 2))
        ovMappings(ovCount).DistJ = CDbl(xdVal(cursor + 3))
        ovMappings(ovCount).FrameLength = ovMappings(ovCount).DistJ - ovMappings(ovCount).DistI

        ovCount = ovCount + 1
        cursor = cursor + XDATA_MAPPING_RECORD_SIZE
    Next j

    If ovCount > 0 Then ReDim Preserve ovMappings(0 To ovCount - 1)
End Function
Private Function AreMappingsEqual(m1() As MappingRecord, C1 As Long, m2() As MappingRecord, C2 As Long) As Boolean
    On Error Resume Next
    AreMappingsEqual = False

    If C1 <= 0 And C2 <= 0 Then
        AreMappingsEqual = True
        Exit Function
    End If

    If C1 <> C2 Then Exit Function

    Dim i           As Long
    For i = 0 To C1 - 1
        If Trim$(UCase$(m1(i).TargetFrame)) <> Trim$(UCase$(m2(i).TargetFrame)) Then Exit Function
        If Trim$(UCase$(m1(i).MapType)) <> Trim$(UCase$(m2(i).MapType)) Then Exit Function
        If Abs(m1(i).DistI - m2(i).DistI) > 0.001 Then Exit Function
        If Abs(m1(i).DistJ - m2(i).DistJ) > 0.001 Then Exit Function
    Next i

    AreMappingsEqual = True
End Function
' ==========================================================================
' HELPER: ReadEffectiveWallData
' Purpose: Determines whether to use Base data or Override data for display.
'          This ensures the UI shows exactly what will be used for SAP.
' ==========================================================================
Private Function ReadEffectiveWallData(ent As Object, _
        ByRef effWall As WallSegmentMap, _
        ByRef effMappings() As MappingRecord, _
        ByRef effCount As Long, _
        ByRef hasOverride As Boolean) As Boolean

    On Error GoTo ErrHandler
    ReadEffectiveWallData = False
    effCount = 0
    hasOverride = False

    If ent Is Nothing Then Exit Function

    ' Variables to hold raw data read from n02 module
    Dim baseWall    As WallSegmentMap
    Dim baseMappings() As MappingRecord
    Dim ovWall      As WallSegmentMap
    Dim ovMappings() As MappingRecord
    Dim ovCount     As Long
    Dim baseCount   As Long

    ' 1. Call the core reader from Module n02
    ' This reads the raw XData into Base and Override components
    baseCount = n02_ACAD_Wall_Force_SAP2000.ReadWallAllXData( _
            ent, baseWall, baseMappings, ovWall, ovMappings, hasOverride, ovCount)

    ' 2. Decision Logic: Which one is "Effective"?
    If hasOverride And ovCount > 0 Then
        ' If Override exists (and has mappings), it TAKES PRECEDENCE
        effWall = ovWall
        effMappings = ovMappings
        effCount = ovCount
    Else
        ' Otherwise use Base data
        effWall = baseWall
        effMappings = baseMappings
        effCount = baseCount

        ' Explicitly set false just in case
        hasOverride = False
    End If

    ' 3. Success
    ReadEffectiveWallData = True
    Exit Function

ErrHandler:
    ' In case of error (e.g. types not defined), return False
    ReadEffectiveWallData = False
Debug.Print "Error in ReadEffectiveWallData: " & err.description
End Function
' Ensure that a layer exists with given name and color index (COM-compatible)
Private Sub EnsureLoadingLayer(acadDoc As Object, layerName As String, colorIndex As Integer)
    On Error Resume Next

    If acadDoc Is Nothing Then Exit Sub

    Dim lay         As Object
    err.Clear
    ' Try get existing layer
    Set lay = Nothing
    Set lay = acadDoc.layers.item(layerName)
    If err.number <> 0 Then
        ' Layer not found -> create it
        err.Clear
        Set lay = acadDoc.layers.Add(layerName)
        If err.number <> 0 Then
            ' Could not create layer, give up silently
            err.Clear
            On Error GoTo 0
            Exit Sub
        End If
    End If

    ' Set color if possible (use direct numeric ACI value - avoid using CShort)
    On Error Resume Next
    If Not lay Is Nothing Then
        ' Try direct assignment first
        lay.color = colorIndex
        If err.number <> 0 Then
            err.Clear
            ' Some COM versions use lowercase property name
            On Error Resume Next
            lay.color = colorIndex
            err.Clear
        End If
    End If

    err.Clear
    On Error GoTo 0
End Sub
' Parse load text like "2.5" or "DL2.5" or "DL 2.5"
' Returns True if parsed ok, patternOut = "DL" by default, valueOut in kN/m
Private Function ParseLoadText(ByVal rawText As String, ByRef patternOut As String, ByRef valueOut As Double) As Boolean
    On Error GoTo ErrHandler

    ParseLoadText = False
    patternOut = "DL"
    valueOut = 0#

    Dim s           As String
    s = Trim$(rawText)

    If Len(s) = 0 Then Exit Function

    ' Replace comma with dot for decimal
    s = Replace(s, ",", ".")

    ' Case 1: pure numeric -> default DL
    If IsNumeric(s) Then
        valueOut = CDbl(s)
        ParseLoadText = (valueOut <> 0#)
        Exit Function
    End If

    ' Case 2: something like "DL2.5", "DL 2.5", "LL3.2", "LL 3.2"
    Dim i           As Long
    Dim prefix      As String
    Dim numPart     As String

    ' Split at first digit
    For i = 1 To Len(s)
        Dim ch      As String
        ch = mid$(s, i, 1)
        If (ch >= "0" And ch <= "9") Or ch = "." Then
            prefix = Trim$(Left$(s, i - 1))
            numPart = Trim$(mid$(s, i))
            Exit For
        End If
    Next i

    If numPart = "" Then
        ' Try splitting by space: "DL 2.5"
        Dim parts() As String
        parts = Split(s, " ")
        If UBound(parts) >= 1 Then
            prefix = Trim$(parts(0))
            numPart = Trim$(parts(1))
        End If
    End If

    If numPart = "" Then Exit Function

    If prefix <> "" Then
        patternOut = UCase$(prefix)
    Else
        patternOut = "DL"
    End If

    If Not IsNumeric(numPart) Then Exit Function

    valueOut = CDbl(numPart)
    ParseLoadText = (valueOut <> 0#)
    Exit Function

ErrHandler:
    ParseLoadText = False
End Function
' Distance squared from point (px,py) to line segment (x1,y1)-(x2,y2)
Private Function PointToSegmentDistanceSquared(px As Double, py As Double, _
        x1 As Double, y1 As Double, _
        x2 As Double, y2 As Double) As Double
    Dim dx As Double, dy As Double
    dx = x2 - x1
    dy = y2 - y1

    If dx = 0# And dy = 0# Then
        ' Degenerate segment
        PointToSegmentDistanceSquared = (px - x1) ^ 2 + (py - y1) ^ 2
        Exit Function
    End If

    Dim t           As Double
    t = ((px - x1) * dx + (py - y1) * dy) / (dx * dx + dy * dy)

    If t < 0# Then t = 0#
    If t > 1# Then t = 1#

    Dim projX As Double, projY As Double
    projX = x1 + t * dx
    projY = y1 + t * dy

    PointToSegmentDistanceSquared = (px - projX) ^ 2 + (py - projY) ^ 2
End Function
' Add one wall row into lstLoadAssignments from an AutoCAD line entity
Private Sub AddWallRowFromEntity(ent As Object, lst As Object, ByRef rowIdx As Long)
    On Error GoTo ErrHandler
    If lst Is Nothing Then Exit Sub
    If ent Is Nothing Then Exit Sub

    Dim effWall     As WallSegmentMap
    Dim effMappings() As MappingRecord
    Dim effCount    As Long
    Dim hasOv       As Boolean

    If Not ReadEffectiveWallData(ent, effWall, effMappings, effCount, hasOv) Then
        Exit Sub
    End If

    Dim sectionName As String
    sectionName = Trim$(effWall.WallType)
    If sectionName = "" And effWall.Thickness > 0 Then
        sectionName = "W" & CInt(effWall.Thickness)
    End If
    If sectionName = "" Then sectionName = "Unknown"

    Dim fullLabel   As String
    fullLabel = n02_ACAD_Wall_Force_SAP2000.GenerateCompositeLabel(effWall, effMappings, effCount)

    Dim mapStr      As String
    mapStr = "(Not Mapped)"
    If effCount > 0 Then
        Dim pTo     As Long
        pTo = InStr(1, fullLabel, " to ", vbTextCompare)
        If pTo > 0 Then
            mapStr = mid$(fullLabel, pTo + 1)
        Else
            mapStr = fullLabel
        End If
    End If

    lst.AddItem
    lst.List(lst.ListCount - 1, COL_IDX_NO) = CStr(rowIdx)
    lst.List(lst.ListCount - 1, COL_IDX_HANDLE) = CStr(ent.Handle)
    lst.List(lst.ListCount - 1, COL_IDX_THICKNESS) = Format(effWall.Thickness, "0")
    lst.List(lst.ListCount - 1, COL_IDX_PATTERN) = effWall.LoadPattern
    lst.List(lst.ListCount - 1, COL_IDX_VALUE) = Format(effWall.LoadValue, "0.00")
    lst.List(lst.ListCount - 1, COL_IDX_MAPPING) = mapStr

    rowIdx = rowIdx + 1
    Exit Sub

ErrHandler:
    ' Skip this entity on error
End Sub
' ==========================================================================================
' FORM CLOSE: Cleanup
' ==========================================================================================
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' Save user settings
    SaveSettingsToExcel

    ' Remove temporary data sheet
    CleanupTempSheet

    If CloseMode = vbFormControlMenu Then
        Cancel = False
    End If
End Sub
