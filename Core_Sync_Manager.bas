Attribute VB_Name = "Core_Sync_Manager"

Option Explicit
'===============================================================
' Core Module: Core_Sync_Manager (OPTIMIZED VERSION v3 - refactored import)
' Purpose: Intelligent 2-way sync between AutoCAD and SAP2000
' Notes: Refactored ImportSelectedEntitiesToSAP to reuse CreateSAPPoint
' and reduce duplicated AddCartesian calls. Kept module name unchanged.
'===============================================================

Private Const DEFAULT_TOLERANCE As Double = 1#
Private Const DEFAULT_SCALE As Double = 1#
Private Const LOG_SHEET_NAME As String = "CADtoSAP_Map"
Private Const COORD_TOLERANCE As Double = 0.01

' Label layer constants
Private Const LAYER_NODE_LABEL As String = "dts_node_label"
Private Const LAYER_FRAME_LABEL As String = "dts_frame_label"
Private Const LAYER_AREA_LABEL As String = "dts_area_label"
Private Const LAYER_SPRING As String = "spring"

Private StatusCallback As Object

' NEW: Auto-sync flags
Private mIsAutoSyncMode As Boolean
Private mSuppressDialogs As Boolean

' ----------------------------
' Public API
' ----------------------------

Public Sub SetSyncStatusCallback(callback As Object)
    Set StatusCallback = callback
End Sub

Public Sub SetAutoSyncMode(enabled As Boolean)
    mIsAutoSyncMode = enabled
    mSuppressDialogs = enabled
    LogStatus "Auto-sync mode: " & IIf(enabled, "ENABLED", "DISABLED")
End Sub

Public Function IsAutoSyncMode() As Boolean
    IsAutoSyncMode = mIsAutoSyncMode
End Function
' ----------------------------
' NEW: Filter Rule Checker (supports field:pattern)
' Fields supported: name, guid, section
' Default field when omitted: name
' Match logic: "match any" => if any rule matches -> filtered out (exclude)
Private Function IsFilteredOut(elementInfo As Object, filterRules As String) As Boolean
    ' elementInfo: Dictionary with possible keys "Name", "GUID", "Section"
    ' filterRules: multiline string, each line is either "field:pattern" or "pattern" (defaults to name)
    On Error GoTo ErrHandler
    IsFilteredOut = False
    
    If elementInfo Is Nothing Then Exit Function
    If Len(Trim$(filterRules)) = 0 Then Exit Function
    
    Dim lines() As String
    lines = Split(filterRules, vbCrLf)
    
    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        Dim raw As String
        raw = Trim$(lines(i))
        If Len(raw) = 0 Then GoTo NextLine
        
        Dim fieldName As String
        Dim pattern As String
        Dim colonPos As Long
        colonPos = InStr(raw, ":")
        If colonPos > 0 Then
            fieldName = LCase$(Trim$(Left$(raw, colonPos - 1)))
            pattern = Trim$(mid$(raw, colonPos + 1))
        Else
            ' default field = name
            fieldName = "name"
            pattern = raw
        End If
        
        ' Normalize supported field names
        If fieldName = "name" Or fieldName = "n" Then
            fieldName = "Name"
        ElseIf fieldName = "guid" Or fieldName = "id" Then
            fieldName = "GUID"
        ElseIf fieldName = "section" Or fieldName = "prop" Or fieldName = "property" Then
            fieldName = "Section"
        Else
            ' Unsupported field -> treat as Name (fallback)
            fieldName = "Name"
        End If
        
        ' Get value from elementInfo (case-insensitive keys)
        Dim fieldValue As String
        fieldValue = ""
        On Error Resume Next
        If elementInfo.exists(fieldName) Then
            fieldValue = CStr(elementInfo(fieldName))
        Else
            ' try case-insensitive lookup
            Dim k As Variant
            For Each k In elementInfo.keys
                If LCase$(CStr(k)) = LCase$(fieldName) Then
                    fieldValue = CStr(elementInfo(k))
                    Exit For
                End If
            Next k
        End If
        On Error GoTo ErrHandler
        
        ' Perform pattern match (case-insensitive)
        Dim pat As String
        pat = pattern
        
        If Len(pat) = 0 Then
            ' empty pattern -> skip
            GoTo NextLine
        End If
        
        Dim lv As String
        lv = LCase$(fieldValue)
        Dim lp As String
        lp = LCase$(pat)
        
        ' Support patterns:
        '  *both* -> contains
        '  prefix* -> startswith
        '  *suffix -> endswith
        '  exact -> exact match
        If Left$(lp, 1) = "*" And Right$(lp, 1) = "*" And Len(lp) >= 2 Then
            Dim inner As String
            inner = mid$(lp, 2, Len(lp) - 2)
            If Len(inner) = 0 Then
                ' pattern = ** -> treat as match all
                IsFilteredOut = True
                Exit Function
            End If
            If InStr(1, lv, inner, vbBinaryCompare) > 0 Then
                IsFilteredOut = True
                Exit Function
            End If
        ElseIf Right$(lp, 1) = "*" Then
            Dim pref As String
            pref = Left$(lp, Len(lp) - 1)
            If Len(pref) = 0 Then
                ' pattern = * -> match all
                IsFilteredOut = True
                Exit Function
            End If
            If Left$(lv, Len(pref)) = pref Then
                IsFilteredOut = True
                Exit Function
            End If
        ElseIf Left$(lp, 1) = "*" Then
            Dim suf As String
            suf = mid$(lp, 2)
            If Len(suf) = 0 Then
                ' pattern = * -> match all
                IsFilteredOut = True
                Exit Function
            End If
            If Right$(lv, Len(suf)) = suf Then
                IsFilteredOut = True
                Exit Function
            End If
        Else
            ' exact match
            If StrComp(lv, lp, vbTextCompare) = 0 Then
                IsFilteredOut = True
                Exit Function
            End If
        End If
        
NextLine:
    Next i

    Exit Function
ErrHandler:
    ' On error, do not filter out by default (fail-safe)
    IsFilteredOut = False
    Resume Next
End Function

' NEW: Get SAP Selection
Private Function GetSAPSelection(SapModel As Object, ByRef selectedFrames As Object, ByRef selectedAreas As Object) As Boolean
    On Error GoTo ErrHandler
    
    GetSAPSelection = False
    Set selectedFrames = CreateObject("Scripting.Dictionary")
    Set selectedAreas = CreateObject("Scripting.Dictionary")
    
    ' Get selected objects count and types
    Dim numSelected As Long
    Dim objType() As Long
    Dim objName() As String
    
    Dim ret As Long
    ret = SapModel.SelectObj.GetSelected(numSelected, objType, objName)
    
    If ret <> 0 Or numSelected = 0 Then
        LogStatus "No objects selected in SAP2000 or selection read failed."
        Exit Function
    End If
    
    LogStatus "SAP2000 Selection: " & numSelected & " objects"
    
    ' Process selection
    ' objType: 2 = Frame, 5 = Area
    Dim i As Long
    For i = 0 To numSelected - 1
        If objType(i) = 2 Then
            ' Frame
            If Not selectedFrames.exists(objName(i)) Then
                selectedFrames.Add objName(i), True
            End If
        ElseIf objType(i) = 5 Then
            ' Area
            If Not selectedAreas.exists(objName(i)) Then
                selectedAreas.Add objName(i), True
            End If
        End If
    Next i
    
    LogStatus "  Selected Frames: " & selectedFrames.count
    LogStatus "  Selected Areas: " & selectedAreas.count
    
    GetSAPSelection = True
    Exit Function
    
ErrHandler:
    LogStatus "ERROR reading SAP selection: " & err.description
    GetSAPSelection = False
End Function
' SMART SYNC: SAP -> AutoCAD with filters and selection
Public Sub SyncSAPToCADWithFilters(acadDoc As Object, SapModel As Object, _
        Optional framesOnly As Boolean = True, _
        Optional showNodeNames As Boolean = True, _
        Optional showFrameNames As Boolean = True, _
        Optional showShellNames As Boolean = True, _
        Optional sapOnlyMode As Boolean = False, _
        Optional frameFilterRules As String = "", _
        Optional areaFilterRules As String = "", _
        Optional silentMode As Boolean = False)

    On Error GoTo ErrHandler

    If Not silentMode Then
        LogStatus "========== Starting OPTIMIZED SAP -> CAD Sync (v4) =========="
    End If

    ' Register data application
    RegisterDataApp acadDoc

    ' Create geometry layers
    EnsureLayerExists acadDoc, "dts_point", 3
    EnsureLayerExists acadDoc, "dts_frame", 7
    EnsureLayerExists acadDoc, "dts_area", 2

    ' Create label layers
    EnsureLayerExists acadDoc, LAYER_NODE_LABEL, 3
    EnsureLayerExists acadDoc, LAYER_FRAME_LABEL, 7
    EnsureLayerExists acadDoc, LAYER_AREA_LABEL, 2

    ' Clear old labels before sync
    ClearLabelLayers acadDoc

    ' Get SAP selection if needed
    Dim sapSelectedFrames As Object
    Dim sapSelectedAreas As Object
    Dim useSelection As Boolean
    useSelection = False
    
    If sapOnlyMode Then
        If GetSAPSelection(SapModel, sapSelectedFrames, sapSelectedAreas) Then
            useSelection = True
            LogStatus "Using SAP selection filter."
        Else
            LogStatus "WARNING: SAP Only mode enabled but no selection found. Plotting all elements."
        End If
    End If

    ' Build SAP dictionaries
    Dim nodeDict    As Object
    Set nodeDict = BuildNodeDictFromSAP(SapModel)
    If Not silentMode Then LogStatus "Extracted " & nodeDict.count & " nodes from SAP2000"

    Dim frameDict   As Object
    Set frameDict = BuildFrameDictFromSAPWithFilter(SapModel, frameFilterRules, useSelection, sapSelectedFrames)
    If Not silentMode Then LogStatus "Extracted " & frameDict.count & " frames from SAP2000 (after filters)"

    Dim areaDict    As Object
    If Not framesOnly Then
        Set areaDict = BuildAreaDictFromSAPWithFilter(SapModel, areaFilterRules, useSelection, sapSelectedAreas)
        If Not silentMode Then LogStatus "Extracted " & areaDict.count & " areas from SAP2000 (after filters)"
    End If

    ' Read existing CAD entities
    Dim existingPoints() As CADPoint
    Dim existingFrames() As CADFrame
    Dim existingAreas() As CADArea

    Dim cadPointCount As Long, cadFrameCount As Long, cadAreaCount As Long
    cadPointCount = Core_XData_Reader.ReadPointsFromCAD(acadDoc, existingPoints)
    cadFrameCount = Core_XData_Reader.ReadFramesFromCAD(acadDoc, existingFrames)
    cadAreaCount = Core_XData_Reader.ReadAreasFromCAD(acadDoc, existingAreas)

    If Not silentMode Then
        LogStatus "Found in CAD: " & cadPointCount & " points, " & cadFrameCount & " frames, " & cadAreaCount & " areas"
    End If

    ' Stats tracking
    Dim stats       As Object
    Set stats = CreateObject("Scripting.Dictionary")
    stats("PointsAdded") = 0
    stats("PointsUpdated") = 0
    stats("PointsDeleted") = 0
    stats("FramesAdded") = 0
    stats("FramesUpdated") = 0
    stats("FramesDeleted") = 0
    stats("AreasAdded") = 0
    stats("AreasUpdated") = 0
    stats("AreasDeleted") = 0

    ' Sync geometry only (no labels)
    SyncPointsSmart acadDoc, nodeDict, existingPoints, stats, silentMode
    SyncFramesSmart acadDoc, frameDict, nodeDict, existingFrames, stats, silentMode

    If Not framesOnly Then
        SyncAreasSmart acadDoc, areaDict, nodeDict, existingAreas, stats, silentMode
    End If

    ' Refresh view
    On Error Resume Next
    acadDoc.Regen 0
    On Error GoTo ErrHandler

    ' Now draw labels separately in dedicated layers
    If showNodeNames Then
        DrawNodeLabels acadDoc, nodeDict
    End If

    If showFrameNames Then
        DrawFrameLabels acadDoc, frameDict, nodeDict
    End If

    If showShellNames And Not framesOnly Then
        DrawAreaLabels acadDoc, areaDict, nodeDict
    End If

    ' Final refresh
    On Error Resume Next
    acadDoc.Regen 0
    On Error GoTo ErrHandler

    ' Show summary only if not silent
    If Not silentMode Then
        Dim summary As String
        summary = "========== OPTIMIZED Sync Completed ==========" & vbCrLf & _
                "Points: +" & stats("PointsAdded") & " ~" & stats("PointsUpdated") & " -" & stats("PointsDeleted") & vbCrLf & _
                "Frames: +" & stats("FramesAdded") & " ~" & stats("FramesUpdated") & " -" & stats("FramesDeleted")

        If Not framesOnly Then
            summary = summary & vbCrLf & "Areas: +" & stats("AreasAdded") & " ~" & stats("AreasUpdated") & " -" & stats("AreasDeleted")
        End If

        LogStatus summary
    End If

    Exit Sub

ErrHandler:
    LogStatus "ERROR in SyncSAPToCADWithFilters: " & err.description
    If Not mSuppressDialogs Then
        MsgBox "Error during sync: " & err.description, vbCritical, "Sync Error"
    End If
End Sub
' ----------------------------
' Build Frame Dictionary with filter by name/guid/section
Private Function BuildFrameDictFromSAPWithFilter(SapModel As Object, filterRules As String, _
        useSelection As Boolean, selectedFrames As Object) As Object
    Set BuildFrameDictFromSAPWithFilter = CreateObject("Scripting.Dictionary")

    On Error Resume Next
    ExtractFrames

    If gFrameCount <= 0 Then Exit Function

    Dim i As Long
    For i = 0 To gFrameCount - 1
        Dim frameName As String
        frameName = CStr(gFrameNames(i))
        
        ' Build element info dictionary for filtering
        Dim fInfoTemp As Object
        Set fInfoTemp = CreateObject("Scripting.Dictionary")
        fInfoTemp.Add "Name", frameName
        ' GUID fallback: use frameName if no separate GUID available
        On Error Resume Next
        If (Not IsEmpty(gFrameGUIDs)) Then
            ' If global array gFrameGUIDs exists, try to read it
            fInfoTemp.Add "GUID", CStr(gFrameGUIDs(i))
        Else
            fInfoTemp.Add "GUID", frameName
        End If
        On Error GoTo 0
        fInfoTemp.Add "Section", CStr(gFrameProp(i))
        
        ' Apply filter rules (if any)
        If Len(Trim$(filterRules)) > 0 Then
            If IsFilteredOut(fInfoTemp, filterRules) Then
                GoTo NextFrame
            End If
        End If
        
        ' Apply selection filter
        If useSelection Then
            If Not selectedFrames.exists(frameName) Then
                GoTo NextFrame
            End If
        End If

        ' Build final fInfo dictionary to add to return collection
        Dim finfo As Object
        Set finfo = CreateObject("Scripting.Dictionary")
        finfo("Name") = frameName
        finfo("P1") = CStr(gFrameP1(i))
        finfo("P2") = CStr(gFrameP2(i))
        finfo("Section") = CStr(gFrameProp(i))
        finfo("Angle") = CDbl(gFrameAngle(i))
        ' Add GUID too
        If fInfoTemp.exists("GUID") Then finfo("GUID") = fInfoTemp("GUID") Else finfo("GUID") = frameName

        BuildFrameDictFromSAPWithFilter.Add frameName, finfo
NextFrame:
    Next i

    On Error GoTo 0
End Function

' ----------------------------
' Build Area Dictionary with filter by name/guid/section
Private Function BuildAreaDictFromSAPWithFilter(SapModel As Object, filterRules As String, _
        useSelection As Boolean, selectedAreas As Object) As Object
    Set BuildAreaDictFromSAPWithFilter = CreateObject("Scripting.Dictionary")

    On Error Resume Next
    ExtractAreas

    If gAreaCount <= 0 Then Exit Function

    Dim i As Long
    For i = 0 To gAreaCount - 1
        Dim areaName As String
        areaName = CStr(gAreaNames(i))
        
        ' Build element info dictionary for filtering
        Dim aInfoTemp As Object
        Set aInfoTemp = CreateObject("Scripting.Dictionary")
        aInfoTemp.Add "Name", areaName
        ' GUID fallback: use areaName if no separate GUID available
        On Error Resume Next
        If (Not IsEmpty(gAreaGUIDs)) Then
            aInfoTemp.Add "GUID", CStr(gAreaGUIDs(i))
        Else
            aInfoTemp.Add "GUID", areaName
        End If
        On Error GoTo 0
        aInfoTemp.Add "Section", CStr(gAreaProp(i))
        
        ' Apply filter rules (if any)
        If Len(Trim$(filterRules)) > 0 Then
            If IsFilteredOut(aInfoTemp, filterRules) Then
                GoTo NextArea
            End If
        End If
        
        ' Apply selection filter
        If useSelection Then
            If Not selectedAreas.exists(areaName) Then
                GoTo NextArea
            End If
        End If

        Dim aInfo As Object
        Set aInfo = CreateObject("Scripting.Dictionary")
        aInfo("Name") = areaName
        aInfo("Section") = CStr(gAreaProp(i))
        aInfo("PointList") = CStr(gAreaPointStr(i))
        aInfo("NumPoints") = gAreaNumPts(i)
        If aInfoTemp.exists("GUID") Then aInfo("GUID") = aInfoTemp("GUID") Else aInfo("GUID") = areaName

        BuildAreaDictFromSAPWithFilter.Add areaName, aInfo
NextArea:
    Next i

    On Error GoTo 0
End Function

' NEW: Import Floor Plan from AutoCAD with Z elevation
Public Sub ImportFloorPlanFromCAD(acadDoc As Object, SapModel As Object, _
        zElevation As Double, _
        Optional tolerance As Double = DEFAULT_TOLERANCE, _
        Optional scaleFactor As Double = DEFAULT_SCALE)
    
    On Error GoTo ErrHandler
    
    LogStatus "========== Starting Floor Plan Import =========="
    LogStatus "Z Elevation: " & zElevation & " mm"
    
    ' Activate AutoCAD for selection
    On Error Resume Next
    Dim acadApp As Object
    Set acadApp = acadDoc.Application
    If Not acadApp Is Nothing Then
        acadApp.Visible = True
        AppActivate acadApp.Caption
    End If
    On Error GoTo ErrHandler
    
    LogStatus "Please select floor plan entities in AutoCAD..."
    DoEvents
    
    ' Get user selection
    Dim handles As Variant
    handles = GetSelectionHandles_OnScreen(acadDoc)
    
    If IsEmpty(handles) Then
        LogStatus "Floor plan import cancelled or no entities selected."
        On Error Resume Next
        AppActivate Application.Caption
        On Error GoTo ErrHandler
        Exit Sub
    End If
    
    ' Return focus to Excel
    On Error Resume Next
    AppActivate Application.Caption
    On Error GoTo ErrHandler
    
    LogStatus "Processing " & (UBound(handles) - LBound(handles) + 1) & " selected entities..."
    
    ' Build SAP node collection and sets
    Dim sapNodeDict As Object
    Set sapNodeDict = BuildNodeDictFromSAP(SapModel)
    Dim sapNodeCol As Collection
    Set sapNodeCol = New Collection
    Dim k As Variant
    For Each k In sapNodeDict.keys
        sapNodeCol.Add sapNodeDict(k)
    Next k
    
    Dim sapFrameSet As Object
    Set sapFrameSet = CreateObject("Scripting.Dictionary")
    BuildSAPFrameSet SapModel, sapFrameSet
    
    Dim sapAreaSet As Object
    Set sapAreaSet = CreateObject("Scripting.Dictionary")
    BuildSAPAreaSet SapModel, sapAreaSet
    
    ' Prepare log sheet
    Dim wsLog As Worksheet
    Set wsLog = PrepareLogSheet()
    Dim logRow As Long
    logRow = wsLog.Cells(wsLog.rows.count, "A").End(xlUp).row + 1
    
    ' Counters
    Dim createdNodes As Long: createdNodes = 0
    Dim createdFrames As Long: createdFrames = 0
    Dim createdAreas As Long: createdAreas = 0
    
    ' Process each selected entity
    Dim i As Long
    For i = LBound(handles) To UBound(handles)
        Dim h As String: h = CStr(handles(i))
        Dim ent As Object
        Set ent = acadDoc.HandleToObject(h)
        If ent Is Nothing Then GoTo NextHandle2
        
        Dim tName As String: tName = LCase$(TypeName(ent))
        
        ' Process Lines as Frames
        If InStr(tName, "line") > 0 And InStr(tName, "polyline") = 0 Then
            Dim spP As Variant, epP As Variant
            spP = ent.StartPoint: epP = ent.EndPoint
            
            Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
            x1 = CDbl(spP(0)): y1 = CDbl(spP(1))
            x2 = CDbl(epP(0)): y2 = CDbl(epP(1))
            
            ' Use specified Z elevation
            Dim z1 As Double, z2 As Double
            z1 = zElevation
            z2 = zElevation
            
            ' Get or create nodes
            Dim n1 As String, n2 As String
            n1 = GetNearestSAPNode(x1, y1, z1, sapNodeCol, tolerance)
            If Len(n1) = 0 Then n1 = CreateSAPPoint(SapModel, x1, y1, z1, sapNodeCol, createdNodes)
            
            n2 = GetNearestSAPNode(x2, y2, z2, sapNodeCol, tolerance)
            If Len(n2) = 0 Then n2 = CreateSAPPoint(SapModel, x2, y2, z2, sapNodeCol, createdNodes)
            
            ' Create frame
            Dim fKey As String
            fKey = NormalizeFrameKey(n1, n2)
            If Not sapFrameSet.exists(fKey) Then
                Dim newF As String: newF = ""
                Dim rf As Long
                rf = SapModel.frameObj.AddByPoint(n1, n2, newF, "Default", "")
                If rf = 0 Then
                    sapFrameSet.Add fKey, newF
                    createdFrames = createdFrames + 1
                    
                    ' Attach XData
                    AttachFrameXData ent, newF, n1, n2, "Default"
                    
                    LogToSheet wsLog, logRow, h, TypeName(ent), "", "Frame", newF, "Floor plan import"
                    logRow = logRow + 1
                    
                    If createdFrames Mod 10 = 0 Then
                        LogStatus "Created " & createdFrames & " frames..."
                    End If
                End If
            End If
            
        ' Process Closed Polylines as Areas
        ElseIf InStr(tName, "polyline") > 0 Or InStr(tName, "lwpolyline") > 0 Then
            Dim isClosed As Boolean: isClosed = False
            On Error Resume Next
            If HasProperty(ent, "Closed") Then
                isClosed = CBool(ent.Closed)
            End If
            On Error GoTo ErrHandler
            
            If Not isClosed Then
                ' Check if coordinates form closed shape
                Dim coords As Variant
                coords = Empty
                On Error Resume Next
                If HasProperty(ent, "Coordinates") Then
                    coords = ent.Coordinates
                End If
                On Error GoTo ErrHandler
                
                If Not IsEmpty(coords) And IsArray(coords) Then
                    Dim nCoords As Long: nCoords = UBound(coords) - LBound(coords) + 1
                    If nCoords >= 4 Then
                        Dim x0 As Double, y0 As Double, xn As Double, yn As Double
                        x0 = coords(LBound(coords))
                        y0 = coords(LBound(coords) + 1)
                        xn = coords(UBound(coords) - 1)
                        yn = coords(UBound(coords))
                        If Abs(x0 - xn) < 0.0001 And Abs(y0 - yn) < 0.0001 Then
                            isClosed = True
                        End If
                    End If
                End If
            End If
            
            If isClosed Then
                Dim verts As Collection
                Set verts = GetPolylineVerticesWithZ(ent, zElevation)
                
                If Not verts Is Nothing And verts.count >= 3 Then
                    Dim ptsArr() As String
                    ReDim ptsArr(0 To verts.count - 1)
                    
                    Dim vi As Long
                    For vi = 1 To verts.count
                        Dim vpt As Variant
                        vpt = verts(vi)
                        Dim vx As Double, vy As Double, vz As Double
                        vx = CDbl(vpt(0)): vy = CDbl(vpt(1)): vz = CDbl(vpt(2))
                        
                        Dim pName As String
                        pName = GetNearestSAPNode(vx, vy, vz, sapNodeCol, tolerance)
                        If Len(pName) = 0 Then
                            pName = CreateSAPPoint(SapModel, vx, vy, vz, sapNodeCol, createdNodes)
                        End If
                        ptsArr(vi - 1) = pName
                    Next vi
                    
                    ' Create area
                    Dim aKey As String
                    aKey = NormalizeAreaKey(ptsArr)
                    If Not sapAreaSet.exists(aKey) Then
                        Dim newArea As String
                        Dim retA As Long
                        retA = SapModel.AreaObj.AddByPoint(UBound(ptsArr) + 1, ptsArr, newArea, "Default", "")
                        If retA = 0 Then
                            sapAreaSet.Add aKey, newArea
                            createdAreas = createdAreas + 1
                            
                            ' Attach XData
                            AttachAreaXData ent, newArea, "Default", Join(ptsArr, ",")
                            
                            LogToSheet wsLog, logRow, h, TypeName(ent), "", "Area", newArea, "Floor plan import"
                            logRow = logRow + 1
                            
                            If createdAreas Mod 5 = 0 Then
                                LogStatus "Created " & createdAreas & " areas..."
                            End If
                        End If
                    End If
                End If
            End If
        End If
        
NextHandle2:
    Next i
    
    ' Refresh SAP view
    On Error Resume Next
    SapModel.View.RefreshView
    On Error GoTo ErrHandler
    
    ' Summary
    Dim msg As String
    msg = "========== Floor Plan Import Complete ==========" & vbCrLf & vbCrLf & _
          "Nodes created: " & createdNodes & vbCrLf & _
          "Frames created: " & createdFrames & vbCrLf & _
          "Areas created: " & createdAreas
    
    LogStatus msg
    MsgBox msg, vbInformation, "Floor Plan Import"
    
    Exit Sub
    
ErrHandler:
    LogStatus "ERROR in ImportFloorPlanFromCAD: " & err.description
    MsgBox "Error during floor plan import: " & err.description, vbCritical, "Import Error"
End Sub

' Helper function to get polyline vertices with specified Z
Private Function GetPolylineVerticesWithZ(ent As Object, zValue As Double) As Collection
    On Error GoTo ErrHandler
    
    Dim verts As Collection
    Set verts = New Collection
    
    Dim coords As Variant
    coords = Empty
    
    If HasProperty(ent, "Coordinates") Then
        On Error Resume Next
        coords = ent.Coordinates
        On Error GoTo 0
    End If
    
    If IsEmpty(coords) Then
        Set GetPolylineVerticesWithZ = verts
        Exit Function
    End If
    
    Dim lb As Long, ub As Long
    lb = LBound(coords): ub = UBound(coords)
    Dim nNums As Long
    nNums = ub - lb + 1
    
    ' Determine step (2 for LWPolyline, 3 for 3DPolyline, or 2 default)
    Dim objType As String
    objType = LCase$(TypeName(ent))
    
    Dim stepN As Long
    If InStr(objType, "lwpolyline") > 0 Then
        stepN = 2
    Else
        If (nNums Mod 3) = 0 Then
            stepN = 3
        Else
            stepN = 2
        End If
    End If
    
    Dim nVerts As Long
    nVerts = nNums \ stepN
    
    Dim i As Long
    For i = 0 To nVerts - 1
        Dim vx As Double, vy As Double
        vx = CDbl(coords(lb + i * stepN))
        vy = CDbl(coords(lb + i * stepN + 1))
        
        Dim vpt(0 To 2) As Double
        vpt(0) = vx
        vpt(1) = vy
        vpt(2) = zValue ' Use specified Z elevation
        verts.Add vpt
    Next i
    
    ' Remove duplicate last vertex if closed
    If verts.count > 1 Then
        Dim firstV As Variant, lastV As Variant
        firstV = verts(1)
        lastV = verts(verts.count)
        If Abs(firstV(0) - lastV(0)) < 0.0001 And Abs(firstV(1) - lastV(1)) < 0.0001 Then
            verts.Remove verts.count
        End If
    End If
    
    Set GetPolylineVerticesWithZ = verts
    Exit Function
    
ErrHandler:
    Set GetPolylineVerticesWithZ = Nothing
End Function

' ----------------------------
' Label Management Functions
' ----------------------------

Private Sub ClearLabelLayers(acadDoc As Object)
    On Error Resume Next

    LogStatus "Clearing old labels..."

    Dim layers      As Variant
    layers = Array(LAYER_NODE_LABEL, LAYER_FRAME_LABEL, LAYER_AREA_LABEL)

    Dim layerName   As Variant
    For Each layerName In layers
        ClearLayer acadDoc, CStr(layerName)
    Next layerName

    On Error GoTo 0
End Sub

Private Sub ClearLayer(acadDoc As Object, layerName As String)
    On Error Resume Next

    Dim ms          As Object
    Set ms = acadDoc.ModelSpace

    Dim ent         As Object
    Dim toDelete    As Collection
    Set toDelete = New Collection

    ' Collect entities to delete
    For Each ent In ms
        If LCase(ent.layer) = LCase(layerName) Then
            toDelete.Add ent
        End If
    Next ent

    ' Delete collected entities
    Dim obj         As Variant
    For Each obj In toDelete
        obj.Delete
    Next obj

    LogStatus "Cleared " & toDelete.count & " objects from layer '" & layerName & "'"

    On Error GoTo 0
End Sub

Private Sub DrawNodeLabels(acadDoc As Object, nodeDict As Object)
    On Error Resume Next

    LogStatus "Drawing node labels..."

    Dim ms          As Object
    Set ms = acadDoc.ModelSpace

    Dim nodeName    As Variant
    Dim count       As Long: count = 0

    For Each nodeName In nodeDict.keys
        Dim nInfo   As Object
        Set nInfo = nodeDict(nodeName)

        Dim X As Double, Y As Double, Z As Double
        X = nInfo("X")
        Y = nInfo("Y")
        Z = nInfo("Z")

        Dim textPt(0 To 2) As Double
        textPt(0) = X + 20
        textPt(1) = Y + 20
        textPt(2) = Z

        Dim txtObj  As Object
        Set txtObj = ms.AddText(CStr(nodeName), textPt, 60)

        If Not txtObj Is Nothing Then
            txtObj.layer = LAYER_NODE_LABEL
            txtObj.color = 3
            count = count + 1
        End If
    Next nodeName

    LogStatus "Created " & count & " node labels"

    On Error GoTo 0
End Sub

Private Sub DrawFrameLabels(acadDoc As Object, frameDict As Object, nodeDict As Object)
    On Error Resume Next

    LogStatus "Drawing frame labels..."

    Dim ms          As Object
    Set ms = acadDoc.ModelSpace

    Dim frameName   As Variant
    Dim count       As Long: count = 0

    For Each frameName In frameDict.keys
        Dim finfo   As Object
        Set finfo = frameDict(frameName)

        Dim p1 As String, p2 As String
        p1 = finfo("P1")
        p2 = finfo("P2")

        If Not nodeDict.exists(p1) Or Not nodeDict.exists(p2) Then
            GoTo NextFrame
        End If

        Dim n1 As Object, n2 As Object
        Set n1 = nodeDict(p1)
        Set n2 = nodeDict(p2)

        Dim midX As Double, midY As Double, midZ As Double
        midX = (n1("X") + n2("X")) / 2
        midY = (n1("Y") + n2("Y")) / 2
        midZ = (n1("Z") + n2("Z")) / 2

        Dim txtPt(0 To 2) As Double
        txtPt(0) = midX
        txtPt(1) = midY
        txtPt(2) = midZ

        Dim txtObj  As Object
        Set txtObj = ms.AddText(CStr(frameName), txtPt, 80)

        If Not txtObj Is Nothing Then
            txtObj.layer = LAYER_FRAME_LABEL
            txtObj.color = 7
            count = count + 1
        End If

NextFrame:
    Next frameName

    LogStatus "Created " & count & " frame labels"

    On Error GoTo 0
End Sub

Private Sub DrawAreaLabels(acadDoc As Object, areaDict As Object, nodeDict As Object)
    On Error Resume Next

    LogStatus "Drawing area labels..."

    Dim ms          As Object
    Set ms = acadDoc.ModelSpace

    Dim areaName    As Variant
    Dim count       As Long: count = 0

    For Each areaName In areaDict.keys
        Dim aInfo   As Object
        Set aInfo = areaDict(areaName)

        Dim pointList As String
        pointList = aInfo("PointList")

        If Len(pointList) = 0 Then GoTo NextArea

        Dim pts()   As String
        pts = Split(pointList, ",")

        ' Calculate centroid
        Dim cx As Double, cy As Double, cz As Double
        cx = 0: cy = 0: cz = 0
        Dim validCount As Long: validCount = 0

        Dim j       As Long
        For j = 0 To UBound(pts)
            Dim pName As String
            pName = Trim$(pts(j))

            If nodeDict.exists(pName) Then
                Dim nInfo As Object
                Set nInfo = nodeDict(pName)
                cx = cx + nInfo("X")
                cy = cy + nInfo("Y")
                cz = cz + nInfo("Z")
                validCount = validCount + 1
            End If
        Next j

        If validCount > 0 Then
            cx = cx / validCount
            cy = cy / validCount
            cz = cz / validCount

            Dim txtPt(0 To 2) As Double
            txtPt(0) = cx
            txtPt(1) = cy
            txtPt(2) = cz

            Dim txtObj As Object
            Set txtObj = ms.AddText(CStr(areaName), txtPt, 80)

            If Not txtObj Is Nothing Then
                txtObj.layer = LAYER_AREA_LABEL
                txtObj.color = 2
                count = count + 1
            End If
        End If

NextArea:
    Next areaName

    LogStatus "Created " & count & " area labels"

    On Error GoTo 0
End Sub

' ----------------------------
' Optimized Smart Sync Functions
' ----------------------------

Private Sub SyncPointsSmart(acadDoc As Object, sapNodes As Object, cadPoints() As CADPoint, stats As Object, Optional silentMode As Boolean = False)
    On Error Resume Next

    If Not silentMode Then LogStatus "Syncing points (optimized mode)..."

    Dim ms          As Object
    Set ms = acadDoc.ModelSpace

    ' Build fast lookup dictionary
    Dim cadLookup   As Object
    Set cadLookup = CreateObject("Scripting.Dictionary")

    Dim i           As Long
    On Error Resume Next
    i = LBound(cadPoints)
    If err.number = 0 Then
        err.Clear
        For i = LBound(cadPoints) To UBound(cadPoints)
            Dim nm  As String
            nm = Trim$(cadPoints(i).nodeName)
            If Len(nm) > 0 Then
                If Not cadLookup.exists(nm) Then
                    cadLookup(nm) = i
                End If
            End If
        Next i
    Else
        err.Clear
    End If
    On Error GoTo 0

    ' Process SAP nodes
    Dim nodeName    As Variant
    For Each nodeName In sapNodes.keys
        Dim nInfo   As Object
        Set nInfo = sapNodes(nodeName)

        Dim sapX As Double, sapY As Double, sapZ As Double
        sapX = nInfo("X")
        sapY = nInfo("Y")
        sapZ = nInfo("Z")

        If cadLookup.exists(CStr(nodeName)) Then
            ' Node exists - check if update needed
            Dim idx As Long
            idx = cadLookup(CStr(nodeName))

            With cadPoints(idx)
                Dim actualX As Double, actualY As Double, actualZ As Double
                actualX = .X: actualY = .Y: actualZ = .Z

                ' Try get actual coordinates from entity
                On Error Resume Next
                Dim entObj As Object
                Set entObj = acadDoc.HandleToObject(.entityHandle)
                If Not entObj Is Nothing Then
                    If HasProperty(entObj, "Center") Then
                        Dim ctr As Variant
                        ctr = entObj.center
                        actualX = CDbl(ctr(0))
                        actualY = CDbl(ctr(1))
                        actualZ = CDbl(ctr(2))
                    End If
                End If
                On Error GoTo 0

                ' Compare coordinates
                If PointChanged(actualX, actualY, actualZ, sapX, sapY, sapZ) Then
                    ' Update needed - delete and recreate
                    DeleteEntityByHandle acadDoc, .entityHandle
                    CreatePointEntity ms, CStr(nodeName), sapX, sapY, sapZ, nInfo("Spring")
                    stats("PointsUpdated") = stats("PointsUpdated") + 1

                    If stats("PointsUpdated") Mod 50 = 0 And Not silentMode Then
                        LogStatus "Updated Node " & nodeName
                    End If
                Else
                    ' Geometry unchanged - update XData only
                    UpdatePointXDataOnly entObj, CStr(nodeName), sapX, sapY, sapZ, nInfo("Spring")
                End If
            End With

            cadLookup.Remove CStr(nodeName)
        Else
            ' New node - create
            CreatePointEntity ms, CStr(nodeName), sapX, sapY, sapZ, nInfo("Spring")
            stats("PointsAdded") = stats("PointsAdded") + 1

            If stats("PointsAdded") Mod 50 = 0 And Not silentMode Then
                LogStatus "Created Node " & nodeName
            End If
        End If
    Next nodeName

    ' Delete remaining CAD points
    Dim remaining   As Variant
    For Each remaining In cadLookup.keys
        idx = cadLookup(remaining)
        DeleteEntityByHandle acadDoc, cadPoints(idx).entityHandle
        stats("PointsDeleted") = stats("PointsDeleted") + 1
    Next remaining

    On Error GoTo 0
End Sub

Private Sub SyncFramesSmart(acadDoc As Object, sapFrames As Object, sapNodes As Object, _
        cadFrames() As CADFrame, stats As Object, Optional silentMode As Boolean = False)
    On Error Resume Next

    If Not silentMode Then LogStatus "Syncing frames (optimized mode)..."

    Dim ms          As Object
    Set ms = acadDoc.ModelSpace

    ' Fast lookup
    Dim cadLookup   As Object
    Set cadLookup = CreateObject("Scripting.Dictionary")

    Dim i           As Long
    On Error Resume Next
    i = LBound(cadFrames)
    If err.number = 0 Then
        err.Clear
        For i = LBound(cadFrames) To UBound(cadFrames)
            cadLookup(cadFrames(i).frameName) = i
        Next i
    Else
        err.Clear
    End If
    On Error GoTo 0

    ' Process SAP frames
    Dim frameName   As Variant
    For Each frameName In sapFrames.keys
        Dim finfo   As Object
        Set finfo = sapFrames(frameName)

        Dim p1 As String, p2 As String, Section As String
        p1 = finfo("P1")
        p2 = finfo("P2")
        Section = finfo("Section")

        If Not sapNodes.exists(p1) Or Not sapNodes.exists(p2) Then
            GoTo NextFrame
        End If

        If cadLookup.exists(CStr(frameName)) Then
            ' Frame exists - check if update needed
            Dim idx As Long
            idx = cadLookup(CStr(frameName))

            With cadFrames(idx)
                If .Point1Name <> p1 Or .Point2Name <> p2 Or .sectionName <> Section Then
                    ' Update needed
                    DeleteEntityByHandle acadDoc, .entityHandle
                    CreateFrameEntity ms, frameName, p1, p2, Section, sapNodes
                    stats("FramesUpdated") = stats("FramesUpdated") + 1

                    If stats("FramesUpdated") Mod 50 = 0 And Not silentMode Then
                        LogStatus "Updated Frame " & frameName
                    End If
                Else
                    ' Check coordinate changes
                    Dim n1 As Object, n2 As Object
                    Set n1 = sapNodes(p1)
                    Set n2 = sapNodes(p2)

                    If NodeCoordsChanged(p1, p2, n1, n2, acadDoc, .entityHandle) Then
                        DeleteEntityByHandle acadDoc, .entityHandle
                        CreateFrameEntity ms, frameName, p1, p2, Section, sapNodes
                        stats("FramesUpdated") = stats("FramesUpdated") + 1
                    Else
                        ' Geometry unchanged - update XData only
                        On Error Resume Next
                        Dim entObj As Object
                        Set entObj = acadDoc.HandleToObject(.entityHandle)
                        If Not entObj Is Nothing Then
                            AttachFrameXData entObj, frameName, p1, p2, Section
                        End If
                        On Error GoTo 0
                    End If
                End If
            End With

            cadLookup.Remove CStr(frameName)
        Else
            ' New frame
            CreateFrameEntity ms, frameName, p1, p2, Section, sapNodes
            stats("FramesAdded") = stats("FramesAdded") + 1

            If stats("FramesAdded") Mod 50 = 0 And Not silentMode Then
                LogStatus "Created Frame " & frameName
            End If
        End If

NextFrame:
    Next frameName

    ' Delete remaining
    Dim remaining   As Variant
    For Each remaining In cadLookup.keys
        idx = cadLookup(remaining)
        DeleteEntityByHandle acadDoc, cadFrames(idx).entityHandle
        stats("FramesDeleted") = stats("FramesDeleted") + 1
    Next remaining

    On Error GoTo 0
End Sub

Private Sub SyncAreasSmart(acadDoc As Object, sapAreas As Object, sapNodes As Object, _
        cadAreas() As CADArea, stats As Object, Optional silentMode As Boolean = False)
    On Error Resume Next

    If Not silentMode Then LogStatus "Syncing areas (optimized mode)..."

    Dim ms          As Object
    Set ms = acadDoc.ModelSpace

    ' Fast lookup
    Dim cadLookup   As Object
    Set cadLookup = CreateObject("Scripting.Dictionary")

    Dim i           As Long
    On Error Resume Next
    i = LBound(cadAreas)
    If err.number = 0 Then
        err.Clear
        For i = LBound(cadAreas) To UBound(cadAreas)
            cadLookup(cadAreas(i).areaName) = i
        Next i
    Else
        err.Clear
    End If
    On Error GoTo 0

    ' Process SAP areas
    Dim areaName    As Variant
    For Each areaName In sapAreas.keys
        Dim aInfo   As Object
        Set aInfo = sapAreas(areaName)

        Dim pointList As String, Section As String
        pointList = aInfo("PointList")
        Section = aInfo("Section")

        If cadLookup.exists(CStr(areaName)) Then
            Dim idx As Long
            idx = cadLookup(CStr(areaName))

            With cadAreas(idx)
                If .pointList <> pointList Or .sectionName <> Section Then
                    DeleteEntityByHandle acadDoc, .entityHandle
                    CreateAreaEntity ms, areaName, pointList, Section, sapNodes
                    stats("AreasUpdated") = stats("AreasUpdated") + 1

                    If stats("AreasUpdated") Mod 20 = 0 And Not silentMode Then
                        LogStatus "Updated Area " & areaName
                    End If
                Else
                    ' Update XData only
                    On Error Resume Next
                    Dim entObj As Object
                    Set entObj = acadDoc.HandleToObject(.entityHandle)
                    If Not entObj Is Nothing Then
                        AttachAreaXData entObj, areaName, Section, pointList
                    End If
                    On Error GoTo 0
                End If
            End With

            cadLookup.Remove CStr(areaName)
        Else
            CreateAreaEntity ms, areaName, pointList, Section, sapNodes
            stats("AreasAdded") = stats("AreasAdded") + 1

            If stats("AreasAdded") Mod 20 = 0 And Not silentMode Then
                LogStatus "Created Area " & areaName
            End If
        End If
    Next areaName

    ' Delete remaining
    Dim remaining   As Variant
    For Each remaining In cadLookup.keys
        idx = cadLookup(remaining)
        DeleteEntityByHandle acadDoc, cadAreas(idx).entityHandle
        stats("AreasDeleted") = stats("AreasDeleted") + 1
    Next remaining

    On Error GoTo 0
End Sub

' ----------------------------
' XData-Only Update Functions
' ----------------------------

Private Sub UpdatePointXDataOnly(entObj As Object, nodeName As String, X As Double, Y As Double, Z As Double, springData As String)
    On Error Resume Next

    If entObj Is Nothing Then Exit Sub

    AttachPointXData entObj, nodeName, X, Y, Z, springData

    On Error GoTo 0
End Sub

' ----------------------------
' Helper Functions
' ----------------------------

Private Function PointChanged(x1 As Double, y1 As Double, z1 As Double, _
        x2 As Double, y2 As Double, z2 As Double) As Boolean
    PointChanged = (Abs(x1 - x2) > COORD_TOLERANCE Or _
            Abs(y1 - y2) > COORD_TOLERANCE Or _
            Abs(z1 - z2) > COORD_TOLERANCE)
End Function

Private Function NodeCoordsChanged(p1Name As String, p2Name As String, _
        n1 As Object, n2 As Object, _
        acadDoc As Object, frameHandle As String) As Boolean
    On Error GoTo ErrHandler
    NodeCoordsChanged = False

    Dim ent         As Object
    Set ent = acadDoc.HandleToObject(frameHandle)
    If ent Is Nothing Then Exit Function

    Dim sp As Variant, ep As Variant
    sp = ent.StartPoint
    ep = ent.EndPoint

    If Not IsArray(sp) Or Not IsArray(ep) Then Exit Function
    If UBound(sp) < 2 Or UBound(ep) < 2 Then Exit Function

    Dim spX As Double, spY As Double, spZ As Double
    Dim epX As Double, epY As Double, epZ As Double

    spX = CDbl(sp(0)): spY = CDbl(sp(1)): spZ = CDbl(sp(2))
    epX = CDbl(ep(0)): epY = CDbl(ep(1)): epZ = CDbl(ep(2))

    Dim n1x As Double, n1y As Double, n1Z As Double
    Dim n2x As Double, n2y As Double, n2Z As Double

    On Error Resume Next
    If Not (n1 Is Nothing) Then
        n1x = CDbl(n1("X"))
        n1y = CDbl(n1("Y"))
        n1Z = CDbl(n1("Z"))
    End If
    If Not (n2 Is Nothing) Then
        n2x = CDbl(n2("X"))
        n2y = CDbl(n2("Y"))
        n2Z = CDbl(n2("Z"))
    End If
    On Error GoTo ErrHandler

    If PointChanged(spX, spY, spZ, n1x, n1y, n1Z) Or _
            PointChanged(epX, epY, epZ, n2x, n2y, n2Z) Then
        NodeCoordsChanged = True
    End If

    Exit Function

ErrHandler:
    NodeCoordsChanged = False
End Function

Private Sub DeleteEntityByHandle(acadDoc As Object, Handle As String)
    On Error Resume Next

    Dim ent         As Object
    Set ent = acadDoc.HandleToObject(Handle)

    If Not ent Is Nothing Then
        ent.Delete
    End If

    On Error GoTo 0
End Sub

Private Sub CreatePointEntity(ms As Object, ByVal nodeName As String, X As Double, Y As Double, Z As Double, springData As String)
    On Error Resume Next

    Dim centerPt(0 To 2) As Double
    centerPt(0) = X: centerPt(1) = Y: centerPt(2) = Z

    Dim radius      As Double
    radius = 10

    Dim circObj     As Object
    Set circObj = ms.AddCircle(centerPt, radius)

    If Not circObj Is Nothing Then
        circObj.layer = "dts_point"
        circObj.color = 3
        AttachPointXData circObj, nodeName, X, Y, Z, springData
    End If

    On Error GoTo 0
End Sub

Private Sub CreateFrameEntity(ms As Object, ByVal frameName As String, p1 As String, p2 As String, _
        Section As String, nodeDict As Object)
    On Error Resume Next

    Dim n1 As Object, n2 As Object
    Set n1 = nodeDict(p1)
    Set n2 = nodeDict(p2)

    Dim startPt(0 To 2) As Double
    Dim endPt(0 To 2) As Double

    startPt(0) = n1("X"): startPt(1) = n1("Y"): startPt(2) = n1("Z")
    endPt(0) = n2("X"): endPt(1) = n2("Y"): endPt(2) = n2("Z")

    Dim lineObj     As Object
    Set lineObj = ms.AddLine(startPt, endPt)

    If Not lineObj Is Nothing Then
        lineObj.layer = "dts_frame"
        lineObj.color = 7
        AttachFrameXData lineObj, frameName, p1, p2, Section
    End If

    On Error GoTo 0
End Sub

Private Sub CreateAreaEntity(ms As Object, ByVal areaName As String, pointList As String, _
        Section As String, nodeDict As Object)
    On Error Resume Next

    If Len(pointList) = 0 Then Exit Sub

    Dim pts()       As String
    pts = Split(pointList, ",")

    Dim nPts        As Long
    nPts = UBound(pts) + 1

    If nPts < 3 Then Exit Sub

    Dim validCount  As Long: validCount = 0
    Dim coords3D()  As Double
    ReDim coords3D(0 To nPts * 3 - 1)

    Dim j           As Long
    Dim minZ As Double, maxZ As Double
    Dim firstZ      As Boolean: firstZ = True

    For j = 0 To nPts - 1
        Dim pName   As String
        pName = Trim$(pts(j))

        If nodeDict.exists(pName) Then
            Dim nInfo As Object
            Set nInfo = nodeDict(pName)

            coords3D(validCount * 3 + 0) = nInfo("X")
            coords3D(validCount * 3 + 1) = nInfo("Y")
            coords3D(validCount * 3 + 2) = nInfo("Z")

            If firstZ Then
                minZ = nInfo("Z")
                maxZ = nInfo("Z")
                firstZ = False
            Else
                If nInfo("Z") < minZ Then minZ = nInfo("Z")
                If nInfo("Z") > maxZ Then maxZ = nInfo("Z")
            End If

            validCount = validCount + 1
        End If
    Next j

    If validCount < 3 Then Exit Sub

    ReDim Preserve coords3D(0 To validCount * 3 - 1)

    Dim zDiff       As Double
    zDiff = Abs(maxZ - minZ)

    Dim isHorizontal As Boolean
    isHorizontal = (zDiff < 10)

    Dim polyObj     As Object

    If isHorizontal Then
        Dim coords2D() As Double
        ReDim coords2D(0 To validCount * 2 - 1)

        For j = 0 To validCount - 1
            coords2D(j * 2 + 0) = coords3D(j * 3 + 0)
            coords2D(j * 2 + 1) = coords3D(j * 3 + 1)
        Next j

        Set polyObj = ms.AddLightWeightPolyline(coords2D)
        If Not polyObj Is Nothing Then
            polyObj.elevation = (minZ + maxZ) / 2
        End If
    Else
        Set polyObj = ms.Add3DPoly(coords3D)
    End If

    If Not polyObj Is Nothing Then
        polyObj.layer = "dts_area"
        polyObj.color = 2
        polyObj.Closed = True
        AttachAreaXData polyObj, areaName, Section, pointList
    End If

    On Error GoTo 0
End Sub

' ----------------------------
' XData Attachment Functions
' ----------------------------

Private Sub AttachPointXData(entObj As Object, nodeName As Variant, X As Double, Y As Double, Z As Double, springData As Variant)
    On Error GoTo ErrHandler

    On Error Resume Next
    entObj.Application.ActiveDocument.RegisteredApplications.Add "DTS_APP"
    On Error GoTo ErrHandler

    Dim xdType()    As Integer
    Dim xdVal()     As Variant

    If springData = "" Then
        ReDim xdType(0 To 4)
        ReDim xdVal(0 To 4)

        xdType(0) = 1001: xdVal(0) = "DTS_APP"
        xdType(1) = 1000: xdVal(1) = CStr(nodeName)
        xdType(2) = 1040: xdVal(2) = X
        xdType(3) = 1040: xdVal(3) = Y
        xdType(4) = 1040: xdVal(4) = Z
    Else
        ReDim xdType(0 To 5)
        ReDim xdVal(0 To 5)

        xdType(0) = 1001: xdVal(0) = "DTS_APP"
        xdType(1) = 1000: xdVal(1) = CStr(nodeName)
        xdType(2) = 1040: xdVal(2) = X
        xdType(3) = 1040: xdVal(3) = Y
        xdType(4) = 1040: xdVal(4) = Z
        xdType(5) = 1000: xdVal(5) = CStr(springData)
    End If

    entObj.SetXData xdType, xdVal

    Exit Sub

ErrHandler:
Debug.Print "ERROR in AttachPointXData: " & err.description
End Sub

Private Sub AttachFrameXData(entObj As Object, frameName As Variant, p1 As Variant, p2 As Variant, Section As Variant)
    On Error GoTo ErrHandler

    On Error Resume Next
    entObj.Application.ActiveDocument.RegisteredApplications.Add "DTS_APP"
    On Error GoTo ErrHandler

    Dim xdType(0 To 4) As Integer
    Dim xdVal(0 To 4) As Variant

    xdType(0) = 1001: xdVal(0) = "DTS_APP"
    xdType(1) = 1000: xdVal(1) = CStr(frameName)
    xdType(2) = 1000: xdVal(2) = CStr(p1)
    xdType(3) = 1000: xdVal(3) = CStr(p2)
    xdType(4) = 1000: xdVal(4) = CStr(Section)

    entObj.SetXData xdType, xdVal

    Exit Sub

ErrHandler:
Debug.Print "ERROR in AttachFrameXData: " & err.description
End Sub

Private Sub AttachAreaXData(entObj As Object, areaName As Variant, Section As Variant, pointList As Variant)
    On Error GoTo ErrHandler

    On Error Resume Next
    entObj.Application.ActiveDocument.RegisteredApplications.Add "DTS_APP"
    On Error GoTo ErrHandler

    Dim xdType(0 To 3) As Integer
    Dim xdVal(0 To 3) As Variant

    xdType(0) = 1001: xdVal(0) = "DTS_APP"
    xdType(1) = 1000: xdVal(1) = CStr(areaName)
    xdType(2) = 1000: xdVal(2) = CStr(Section)
    xdType(3) = 1000: xdVal(3) = CStr(pointList)

    entObj.SetXData xdType, xdVal

    Exit Sub

ErrHandler:
Debug.Print "ERROR in AttachAreaXData: " & err.description
End Sub

' ----------------------------
' SAP Data Builders
' ----------------------------

Private Function BuildNodeDictFromSAP(SapModel As Object) As Object
    Set BuildNodeDictFromSAP = CreateObject("Scripting.Dictionary")

    On Error Resume Next
    ExtractPoints

    If gPointCount <= 0 Then Exit Function

    Dim i           As Long
    For i = 0 To gPointCount - 1
        Dim nodeName As String
        nodeName = CStr(gPointNames(i))

        Dim nInfo   As Object
        Set nInfo = CreateObject("Scripting.Dictionary")
        nInfo("Name") = nodeName
        nInfo("X") = CDbl(gDictX(nodeName))
        nInfo("Y") = CDbl(gDictY(nodeName))
        nInfo("Z") = CDbl(gDictZ(nodeName))
        nInfo("Spring") = GetSpringData(SapModel, nodeName)

        BuildNodeDictFromSAP.Add nodeName, nInfo
    Next i

    On Error GoTo 0
End Function

Private Function BuildFrameDictFromSAP(SapModel As Object) As Object
    Set BuildFrameDictFromSAP = CreateObject("Scripting.Dictionary")

    On Error Resume Next
    ExtractFrames

    If gFrameCount <= 0 Then Exit Function

    Dim i           As Long
    For i = 0 To gFrameCount - 1
        Dim frameName As String
        frameName = CStr(gFrameNames(i))

        Dim finfo   As Object
        Set finfo = CreateObject("Scripting.Dictionary")
        finfo("Name") = frameName
        finfo("P1") = CStr(gFrameP1(i))
        finfo("P2") = CStr(gFrameP2(i))
        finfo("Section") = CStr(gFrameProp(i))
        finfo("Angle") = CDbl(gFrameAngle(i))

        BuildFrameDictFromSAP.Add frameName, finfo
    Next i

    On Error GoTo 0
End Function

Private Function BuildAreaDictFromSAP(SapModel As Object) As Object
    Set BuildAreaDictFromSAP = CreateObject("Scripting.Dictionary")

    On Error Resume Next
    ExtractAreas

    If gAreaCount <= 0 Then Exit Function

    Dim i           As Long
    For i = 0 To gAreaCount - 1
        Dim areaName As String
        areaName = CStr(gAreaNames(i))

        Dim aInfo   As Object
        Set aInfo = CreateObject("Scripting.Dictionary")
        aInfo("Name") = areaName
        aInfo("Section") = CStr(gAreaProp(i))
        aInfo("PointList") = CStr(gAreaPointStr(i))
        aInfo("NumPoints") = gAreaNumPts(i)

        BuildAreaDictFromSAP.Add areaName, aInfo
    Next i

    On Error GoTo 0
End Function

' ----------------------------
' CAD to SAP Sync (SILENT MODE FOR AUTO-SYNC)
' ----------------------------

' NEW: Silent import for auto-sync (no dialogs, no user interaction)
Public Sub ImportCADEntityToSAP_Silent(acadDoc As Object, SapModel As Object, _
        entityHandle As String, _
        Optional tolerance As Double = DEFAULT_TOLERANCE, _
        Optional scaleFactor As Double = DEFAULT_SCALE)
    On Error GoTo ErrHandler

    If Not mIsAutoSyncMode Then Exit Sub    ' Only in auto-sync mode

    LogStatus "[AUTO-SYNC] Processing CAD entity: " & entityHandle

    ' Get entity
    Dim ent         As Object
    Set ent = acadDoc.HandleToObject(entityHandle)
    If ent Is Nothing Then Exit Sub

    ' Build SAP lookups (lightweight)
    Dim sapNodeList As Collection
    Set sapNodeList = New Collection
    BuildSAPNodeList SapModel, sapNodeList, scaleFactor

    Dim sapFrameSet As Object
    Set sapFrameSet = CreateObject("Scripting.Dictionary")
    BuildSAPFrameSet SapModel, sapFrameSet

    Dim sapAreaSet  As Object
    Set sapAreaSet = CreateObject("Scripting.Dictionary")
    BuildSAPAreaSet SapModel, sapAreaSet

    ' Classify entity
    Dim tName       As String
    tName = LCase(TypeName(ent))

    Dim createdNodes As Long: createdNodes = 0

    ' Prepare log sheet and row
    Dim wsLog       As Worksheet
    Dim logRow      As Long
    Set wsLog = PrepareLogSheet()
    logRow = wsLog.Cells(wsLog.rows.count, "A").End(xlUp).row + 1

    ' Process based on type
    If InStr(tName, "circle") > 0 Then
        ' Point (Circle)
        Dim c       As Variant
        c = ent.center
        Dim px As Double, py As Double, pz As Double
        px = CDbl(c(0)): py = CDbl(c(1)): pz = CDbl(c(2))

        Dim existingNode As String
        existingNode = GetNearestSAPNode(px, py, pz, sapNodeList, tolerance)

        If Len(existingNode) = 0 Then
            Dim newNode As String
            newNode = CreateSAPPoint(SapModel, px, py, pz, sapNodeList, createdNodes)

            ' AUTO-SYNC: Apply restraint if in "spring" layer
            On Error Resume Next
            Dim entLayer As String
            entLayer = LCase(Trim(ent.layer))
            If entLayer = LCase(LAYER_SPRING) Then
                ApplyPinnedRestraint SapModel, newNode
                LogStatus "[AUTO-SYNC] Applied pinned restraint (U1,U2,U3) to node: " & newNode
            End If
            On Error GoTo ErrHandler

            ' AUTO-SYNC XDATA: Get SAP coordinates back
            Dim sapX As Double, sapY As Double, sapZ As Double
            SapModel.pointObj.GetCoordCartesian newNode, sapX, sapY, sapZ

            ' Attach XData to CAD entity
            AttachPointXData ent, newNode, sapX, sapY, sapZ, ""

            LogStatus "[AUTO-SYNC] Created SAP node: " & newNode

            ' Log mapping to worksheet for session tracking
            On Error Resume Next
            LogToSheet wsLog, logRow, CStr(ent.Handle), TypeName(ent), ent.layer, "Point", newNode, "Auto-created"
            logRow = logRow + 1
            On Error GoTo 0
        End If

    ElseIf InStr(tName, "block") > 0 Or InStr(tName, "insert") > 0 Then
        ' Block Reference - Get insertion point (centroid)
        Dim insPoint As Variant
        On Error Resume Next
        If HasProperty(ent, "InsertionPoint") Then
            insPoint = ent.insertionPoint
        ElseIf HasProperty(ent, "Position") Then
            insPoint = ent.Position
        End If
        On Error GoTo ErrHandler

        If Not IsEmpty(insPoint) And IsArray(insPoint) Then
            Dim bx As Double, by As Double, bz As Double
            bx = CDbl(insPoint(0)): by = CDbl(insPoint(1)): bz = CDbl(insPoint(2))

            Dim existingBlockNode As String
            existingBlockNode = GetNearestSAPNode(bx, by, bz, sapNodeList, tolerance)

            If Len(existingBlockNode) = 0 Then
                Dim newBlockNode As String
                newBlockNode = CreateSAPPoint(SapModel, bx, by, bz, sapNodeList, createdNodes)

                ' AUTO-SYNC: Apply restraint if in "spring" layer
                On Error Resume Next
                Dim blockLayer As String
                blockLayer = LCase(Trim(ent.layer))
                If blockLayer = LCase(LAYER_SPRING) Then
                    ApplyPinnedRestraint SapModel, newBlockNode
                    LogStatus "[AUTO-SYNC] Applied pinned restraint (U1,U2,U3) to block node: " & newBlockNode
                End If
                On Error GoTo ErrHandler

                ' AUTO-SYNC XDATA
                Dim sapBx As Double, sapBy As Double, sapBz As Double
                SapModel.pointObj.GetCoordCartesian newBlockNode, sapBx, sapBy, sapBz

                ' Attach XData
                AttachPointXData ent, newBlockNode, sapBx, sapBy, sapBz, ""

                LogStatus "[AUTO-SYNC] Created SAP node from block: " & newBlockNode

                ' Log mapping
                On Error Resume Next
                LogToSheet wsLog, logRow, CStr(ent.Handle), TypeName(ent), ent.layer, "Point", newBlockNode, "Auto-created from block"
                logRow = logRow + 1
                On Error GoTo 0
            End If
        End If

    ElseIf InStr(tName, "line") > 0 And InStr(tName, "polyline") = 0 Then
        ' Frame
        Dim sp As Variant, ep As Variant
        sp = ent.StartPoint: ep = ent.EndPoint

        Dim x1 As Double, y1 As Double, z1 As Double
        Dim x2 As Double, y2 As Double, z2 As Double
        x1 = CDbl(sp(0)): y1 = CDbl(sp(1)): z1 = CDbl(sp(2))
        x2 = CDbl(ep(0)): y2 = CDbl(ep(1)): z2 = CDbl(ep(2))

        ' Get or create nodes
        Dim n1 As String, n2 As String
        n1 = GetNearestSAPNode(x1, y1, z1, sapNodeList, tolerance)
        If Len(n1) = 0 Then
            n1 = CreateSAPPoint(SapModel, x1, y1, z1, sapNodeList, createdNodes)
        End If

        n2 = GetNearestSAPNode(x2, y2, z2, sapNodeList, tolerance)
        If Len(n2) = 0 Then
            n2 = CreateSAPPoint(SapModel, x2, y2, z2, sapNodeList, createdNodes)
        End If

        ' Check if frame exists
        Dim fKey    As String
        fKey = NormalizeFrameKey(n1, n2)

        If Not sapFrameSet.exists(fKey) Then
            Dim newFrame As String
            Dim ret As Long
            ret = SapModel.frameObj.AddByPoint(n1, n2, newFrame, "Default", "")

            If ret = 0 Then
                sapFrameSet.Add fKey, newFrame

                ' AUTO-SYNC XDATA: Get section name from SAP
                Dim sectionName As String
                SapModel.frameObj.GetSection newFrame, sectionName

                ' Attach XData to CAD entity
                AttachFrameXData ent, newFrame, n1, n2, sectionName

                LogStatus "[AUTO-SYNC] Created SAP frame: " & newFrame

                ' Log mapping to worksheet for session tracking
                On Error Resume Next
                LogToSheet wsLog, logRow, CStr(ent.Handle), TypeName(ent), ent.layer, "Frame", newFrame, "Auto-created"
                logRow = logRow + 1
                On Error GoTo 0
            End If

            ' Refresh SAP view immediately
            On Error Resume Next
            SapModel.View.RefreshView
            On Error GoTo 0
        End If

    ElseIf InStr(tName, "polyline") > 0 Then
        ' Area
        Dim isClosed As Boolean
        On Error Resume Next
        isClosed = ent.Closed
        On Error GoTo ErrHandler

        If isClosed Then
            Dim verts As Collection
            Set verts = GetPolylineVertices(ent)

            If Not verts Is Nothing And verts.count >= 3 Then
                ' Get or create nodes for all vertices
                Dim ptsArr() As String
                ReDim ptsArr(0 To verts.count - 1)

                Dim vi As Long
                For vi = 1 To verts.count
                    Dim vpt As Variant
                    vpt = verts(vi)

                    Dim vx As Double, vy As Double, vz As Double
                    vx = CDbl(vpt(0)): vy = CDbl(vpt(1)): vz = CDbl(vpt(2))

                    Dim pName As String
                    pName = GetNearestSAPNode(vx, vy, vz, sapNodeList, tolerance)
                    If Len(pName) = 0 Then
                        pName = CreateSAPPoint(SapModel, vx, vy, vz, sapNodeList, createdNodes)
                    End If

                    ptsArr(vi - 1) = pName
                Next vi

                ' Check if area exists
                Dim aKey As String
                aKey = NormalizeAreaKey(ptsArr)

                If Not sapAreaSet.exists(aKey) Then
                    Dim newArea As String
                    Dim retA As Long
                    retA = SapModel.AreaObj.AddByPoint(UBound(ptsArr) + 1, ptsArr, newArea, "Default", "")

                    If retA = 0 Then
                        sapAreaSet.Add aKey, newArea

                        ' AUTO-SYNC XDATA: Get section name from SAP
                        Dim areaSectionName As String
                        SapModel.AreaObj.GetProperty newArea, areaSectionName

                        ' Build point list string
                        Dim pointListStr As String
                        pointListStr = Join(ptsArr, ",")

                        ' Attach XData to CAD entity
                        AttachAreaXData ent, newArea, areaSectionName, pointListStr

                        LogStatus "[AUTO-SYNC] Created SAP area: " & newArea

                        ' Log mapping to worksheet for session tracking
                        On Error Resume Next
                        LogToSheet wsLog, logRow, CStr(ent.Handle), TypeName(ent), ent.layer, "Area", newArea, "Auto-created"
                        logRow = logRow + 1
                        On Error GoTo 0
                    End If

                    ' Refresh SAP view
                    On Error Resume Next
                    SapModel.View.RefreshView
                    On Error GoTo 0
                End If
            End If
        End If
    End If

    ' Refresh SAP view at end as well (redundant safe)
    On Error Resume Next
    SapModel.View.RefreshView
    On Error GoTo 0

    Exit Sub

ErrHandler:
    LogStatus "[AUTO-SYNC] ERROR: " & err.description
End Sub

' ----------------------------
' MISSING / RESTORED FUNCTIONS FROM v1
' (CAD -> SAP import, logging, selection helpers, ExecuteImport, ClearSAPModel, ApplySpring)
' ----------------------------

Public Sub SyncCADToSAP(acadDoc As Object, SapModel As Object, _
        Optional tolerance As Double = DEFAULT_TOLERANCE, _
        Optional scaleFactor As Double = DEFAULT_SCALE, _
        Optional overwriteMode As Boolean = False)

    On Error GoTo ErrHandler

    LogStatus "========== Starting CAD -> SAP Sync =========="

    ' Overwrite mode: clear SAP model first
    If overwriteMode Then
        LogStatus "OVERWRITE MODE: Clearing SAP model..."
        If Not mSuppressDialogs Then
            If MsgBox("Overwrite mode will DELETE ALL objects in SAP2000 model. Continue?", vbYesNo + vbExclamation, "Confirm Overwrite") <> vbYes Then
                LogStatus "Import cancelled by user"
                Exit Sub
            End If
        End If
        ClearSAPModel SapModel
    End If

    ' Read entities from CAD
    Dim pointsArr() As CADPoint
    Dim framesArr() As CADFrame
    Dim areasArr()  As CADArea

    Dim pCount As Long, fCount As Long, aCount As Long

    pCount = Core_XData_Reader.ReadPointsFromCAD(acadDoc, pointsArr)
    fCount = Core_XData_Reader.ReadFramesFromCAD(acadDoc, framesArr)
    aCount = Core_XData_Reader.ReadAreasFromCAD(acadDoc, areasArr)

    LogStatus "Read from CAD: " & pCount & " points, " & fCount & " frames, " & aCount & " areas"

    ' Build SAP lookups
    Dim sapNodeList As Collection
    Set sapNodeList = New Collection
    BuildSAPNodeList SapModel, sapNodeList, scaleFactor

    Dim sapFrameSet As Object
    Set sapFrameSet = CreateObject("Scripting.Dictionary")
    BuildSAPFrameSet SapModel, sapFrameSet

    Dim sapAreaSet  As Object
    Set sapAreaSet = CreateObject("Scripting.Dictionary")
    BuildSAPAreaSet SapModel, sapAreaSet

    ' Prepare log sheet
    Dim wsLog       As Worksheet
    Set wsLog = PrepareLogSheet()
    Dim logRow      As Long
    logRow = wsLog.Cells(wsLog.rows.count, "A").End(xlUp).row + 1

    ' Execute import
    Dim results     As Object
    Set results = ExecuteImport(SapModel, pointsArr, framesArr, areasArr, sapNodeList, sapFrameSet, sapAreaSet, tolerance, wsLog, logRow, acadDoc)

    ' Refresh SAP view
    On Error Resume Next
    SapModel.View.RefreshView
    On Error GoTo ErrHandler

    ' Show results
    Dim msg         As String
    msg = "========== Import Complete ==========" & vbCrLf & vbCrLf & _
            "Points created: " & results("PointsCreated") & vbCrLf & _
            "Frames created: " & results("FramesCreated") & vbCrLf & _
            "Areas created: " & results("AreasCreated") & vbCrLf & _
            "Failures: " & results("Failures")

    If Not mSuppressDialogs Then
        MsgBox msg, vbInformation, "Import Results"
    Else
        LogStatus msg
    End If
    LogStatus "========== CAD -> SAP Sync Completed =========="

    Exit Sub

ErrHandler:
    LogStatus "ERROR in SyncCADToSAP: " & err.description
    If Not mSuppressDialogs Then
        MsgBox "Error during sync: " & err.description, vbCritical, "Sync Error"
    End If
End Sub

Private Sub ClearSAPModel(SapModel As Object)
    On Error Resume Next

    ' Delete all frames
    Dim fCount As Long, fNames() As String
    SapModel.frameObj.GetNameList fCount, fNames
    Dim i           As Long
    For i = 0 To fCount - 1
        SapModel.frameObj.Delete fNames(i)
    Next i
    LogStatus "Deleted " & fCount & " frames"

    ' Delete all areas
    Dim aCount As Long, aNames() As String
    SapModel.AreaObj.GetNameList aCount, aNames
    For i = 0 To aCount - 1
        SapModel.AreaObj.Delete aNames(i)
    Next i
    LogStatus "Deleted " & aCount & " areas"

    ' Delete all points
    Dim pCount As Long, pNames() As String
    SapModel.pointObj.GetNameList pCount, pNames
    For i = 0 To pCount - 1
        SapModel.pointObj.Delete pNames(i)
    Next i
    LogStatus "Deleted " & pCount & " points"

    ' Refresh SAP view after clearing model
    On Error Resume Next
    SapModel.View.RefreshView
    On Error GoTo 0

End Sub

Private Function PrepareLogSheet() As Worksheet
    On Error Resume Next

    Dim ws          As Worksheet
    Set ws = ThisWorkbook.Worksheets(LOG_SHEET_NAME)

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.Name = LOG_SHEET_NAME
        ws.Range("A1:F1").Value = Array("CAD_Handle", "CAD_Type", "Layer", "SAP_Type", "SAP_Name", "Notes")
    End If

    Set PrepareLogSheet = ws
    On Error GoTo 0
End Function

Private Sub LogToSheet(ws As Worksheet, row As Long, Handle As String, cadType As String, layer As String, sapType As String, sapName As String, notes As String)
    On Error Resume Next
    ws.Cells(row, 1).Value = Handle
    ws.Cells(row, 2).Value = cadType
    ws.Cells(row, 3).Value = layer
    ws.Cells(row, 4).Value = sapType
    ws.Cells(row, 5).Value = sapName
    ws.Cells(row, 6).Value = notes
    On Error GoTo 0
End Sub

Private Function ExecuteImport(SapModel As Object, _
        pointsArr() As CADPoint, framesArr() As CADFrame, areasArr() As CADArea, _
        sapNodes As Collection, sapFrames As Object, sapAreas As Object, _
        tolerance As Double, wsLog As Worksheet, ByRef logRow As Long, _
        Optional acadDoc As Object = Nothing) As Object

    Set ExecuteImport = CreateObject("Scripting.Dictionary")

    Dim pCreated As Long, fCreated As Long, aCreated As Long
    Dim fSkipped As Long, aSkipped As Long, failures As Long

    pCreated = 0: fCreated = 0: aCreated = 0
    fSkipped = 0: aSkipped = 0: failures = 0

    Dim i           As Long

    ' Create points
    On Error Resume Next
    i = LBound(pointsArr)
    If err.number = 0 Then
        err.Clear
        For i = LBound(pointsArr) To UBound(pointsArr)
            Dim pt  As CADPoint
            pt = pointsArr(i)

            Dim existingNode As String
            existingNode = GetNearestSAPNode(pt.X, pt.Y, pt.Z, sapNodes, tolerance)

            If Len(existingNode) = 0 Then
                Dim newNode As String
                newNode = CreateSAPPoint(SapModel, pt.X, pt.Y, pt.Z, sapNodes, pCreated)

                ' Apply spring data if present
                If Len(pt.springData) > 0 Then
                    ApplySpring SapModel, newNode, pt.springData
                End If

                ' Apply pinned restraint if in "spring" layer
                On Error Resume Next
                Dim ptLayer As String
                ptLayer = LCase(Trim(pt.layerName))
                If ptLayer = LCase(LAYER_SPRING) Then
                    ApplyPinnedRestraint SapModel, newNode
                    LogStatus "Applied pinned restraint to node: " & newNode & " (from layer: " & pt.layerName & ")"
                End If
                On Error GoTo 0

                LogToSheet wsLog, logRow, pt.entityHandle, "Circle", pt.layerName, "Point", newNode, _
                        "Created (" & Format(pt.X, "0.00") & "," & Format(pt.Y, "0.00") & "," & Format(pt.Z, "0.00") & ")"
                logRow = logRow + 1
            End If
        Next i
    Else
        err.Clear
    End If
    On Error GoTo 0

    ' Create frames
    On Error Resume Next
    i = LBound(framesArr)
    If err.number = 0 Then
        err.Clear
        For i = LBound(framesArr) To UBound(framesArr)
            Dim fr  As CADFrame
            fr = framesArr(i)

            Dim fKey As String
            fKey = NormalizeFrameKey(fr.Point1Name, fr.Point2Name)

            If sapFrames.exists(fKey) Then
                fSkipped = fSkipped + 1
                LogToSheet wsLog, logRow, fr.entityHandle, "Line", fr.layerName, "Frame", "SKIPPED", "Already exists"
                logRow = logRow + 1
            Else
                Dim newFrame As String
                Dim retF As Long
                retF = SapModel.frameObj.AddByPoint(fr.Point1Name, fr.Point2Name, newFrame, fr.sectionName, "")

                If err.number <> 0 Or retF <> 0 Then
                    failures = failures + 1
                    LogToSheet wsLog, logRow, fr.entityHandle, "Line", fr.layerName, "Frame", "FAILED", "ret=" & retF
                    logRow = logRow + 1
                    err.Clear
                Else
                    fCreated = fCreated + 1
                    sapFrames.Add fKey, newFrame
                    LogToSheet wsLog, logRow, fr.entityHandle, "Line", fr.layerName, "Frame", newFrame, "Section: " & fr.sectionName
                    logRow = logRow + 1
                    ' Refresh SAP view per creation
                    On Error Resume Next
                    SapModel.View.RefreshView
                    On Error GoTo 0
                End If
            End If
        Next i
    Else
        err.Clear
    End If
    On Error GoTo 0

    ' Create areas (with geometry extraction support)
    On Error Resume Next
    i = LBound(areasArr)
    If err.number = 0 Then
        err.Clear
        For i = LBound(areasArr) To UBound(areasArr)
            Dim ar  As CADArea
            ar = areasArr(i)

            Dim pts() As String
            Dim ptsCount As Long
            ptsCount = 0

            If Len(Trim$(ar.pointList)) > 0 Then
                pts = Split(ar.pointList, ",")
                ptsCount = UBound(pts) - LBound(pts) + 1
            Else
                If Not acadDoc Is Nothing Then
                    On Error Resume Next
                    Dim entObj As Object
                    Set entObj = acadDoc.HandleToObject(ar.entityHandle)
                    On Error GoTo 0
                    If Not entObj Is Nothing Then
                        Dim verts As Collection
                        Set verts = GetPolylineVertices(entObj)
                        Dim nVerts As Long
                        nVerts = 0
                        If Not verts Is Nothing Then nVerts = verts.count
                        If nVerts >= 3 Then
                            ReDim pts(0 To nVerts - 1)
                            Dim vi As Long
                            For vi = 1 To nVerts
                                Dim vpt As Variant
                                vpt = verts(vi)
                                Dim vx As Double, vy As Double, vz As Double
                                vx = CDbl(vpt(0)): vy = CDbl(vpt(1)): vz = CDbl(vpt(2))

                                Dim pName As String
                                pName = GetNearestSAPNode(vx, vy, vz, sapNodes, tolerance)
                                If Len(pName) = 0 Then
                                    Dim tmpCreated As Long
                                    tmpCreated = pCreated
                                    pName = CreateSAPPoint(SapModel, vx, vy, vz, sapNodes, tmpCreated)
                                    pCreated = tmpCreated
                                End If
                                pts(vi - 1) = pName
                            Next vi
                            ptsCount = nVerts
                        End If
                    End If
                End If
            End If

            If ptsCount = 0 Then
                aSkipped = aSkipped + 1
                LogToSheet wsLog, logRow, ar.entityHandle, "Polyline", ar.layerName, "Area", "SKIPPED", "No vertex data"
                logRow = logRow + 1
                GoTo NextAreaLoop
            End If

            Dim aKey As String
            aKey = NormalizeAreaKey(pts)

            If sapAreas.exists(aKey) Then
                aSkipped = aSkipped + 1
                LogToSheet wsLog, logRow, ar.entityHandle, "Polyline", ar.layerName, "Area", "SKIPPED", "Already exists"
                logRow = logRow + 1
            Else
                Dim newArea As String
                Dim retA As Long
                retA = SapModel.AreaObj.AddByPoint(UBound(pts) - LBound(pts) + 1, pts, newArea, ar.sectionName, "")

                If err.number <> 0 Or retA <> 0 Then
                    failures = failures + 1
                    LogToSheet wsLog, logRow, ar.entityHandle, "Polyline", ar.layerName, "Area", "FAILED", "ret=" & retA
                    logRow = logRow + 1
                    err.Clear
                Else
                    aCreated = aCreated + 1
                    sapAreas.Add aKey, newArea
                    LogToSheet wsLog, logRow, ar.entityHandle, "Polyline", ar.layerName, "Area", newArea, "Section: " & ar.sectionName
                    logRow = logRow + 1
                    ' Refresh SAP view after creating area
                    On Error Resume Next
                    SapModel.View.RefreshView
                    On Error GoTo 0
                End If
            End If
NextAreaLoop:
        Next i
    Else
        err.Clear
    End If
    On Error GoTo 0

    ExecuteImport("PointsCreated") = pCreated
    ExecuteImport("FramesCreated") = fCreated
    ExecuteImport("AreasCreated") = aCreated
    ExecuteImport("FramesSkipped") = fSkipped
    ExecuteImport("AreasSkipped") = aSkipped
    ExecuteImport("Failures") = failures
End Function

Private Sub ApplySpring(SapModel As Object, nodeName As String, springData As String)
    On Error Resume Next

    If Len(springData) = 0 Then Exit Sub

    Dim parts()     As String
    parts = Split(springData, ",")

    If UBound(parts) >= 5 Then
        Dim k1 As Double, k2 As Double, k3 As Double
        Dim k4 As Double, k5 As Double, k6 As Double

        k1 = CDbl(parts(0))
        k2 = CDbl(parts(1))
        k3 = CDbl(parts(2))
        k4 = CDbl(parts(3))
        k5 = CDbl(parts(4))
        k6 = CDbl(parts(5))

        SapModel.pointObj.SetSpring nodeName, k1, k2, k3, k4, k5, k6
    End If

    On Error GoTo 0
End Sub

' Apply pinned restraint (U1, U2, U3 locked, R1, R2, R3 free)
Private Sub ApplyPinnedRestraint(SapModel As Object, nodeName As String)
    On Error Resume Next

    Dim restraints(0 To 5) As Boolean
    restraints(0) = True  ' U1 - locked
    restraints(1) = True  ' U2 - locked
    restraints(2) = True  ' U3 - locked
    restraints(3) = False    ' R1 - free
    restraints(4) = False    ' R2 - free
    restraints(5) = False    ' R3 - free

    Dim ret         As Long
    ret = SapModel.pointObj.SetRestraint(nodeName, restraints, 0)    ' 0 = eItemType.Object

    If ret <> 0 Then
        LogStatus "WARNING: Failed to apply restraint to node " & nodeName & " (ret=" & ret & ")"
    End If

    On Error GoTo 0
End Sub

Public Sub ImportSelectedEntitiesToSAP(acadDoc As Object, SapModel As Object, _
        Optional ByVal tolerance As Double = DEFAULT_TOLERANCE, _
        Optional ByVal scaleFactor As Double = DEFAULT_SCALE)
    On Error GoTo ErrHandler

    If acadDoc Is Nothing Then
        LogStatus "ERROR: AutoCAD document not provided."
        Exit Sub
    End If
    If SapModel Is Nothing Then
        LogStatus "ERROR: SAP2000 not connected."
        Exit Sub
    End If

    On Error Resume Next
    Dim acadApp     As Object
    Set acadApp = acadDoc.Application
    If Not acadApp Is Nothing Then
        acadApp.Visible = True
        AppActivate acadApp.Caption
    End If
    On Error GoTo ErrHandler

    LogStatus "Import: Please select entities in AutoCAD (layer '0', 'dts_*', or 'spring')."

    DoEvents

    Dim handles     As Variant
    handles = GetSelectionHandles_OnScreen(acadDoc)

    If IsEmpty(handles) Then
        LogStatus "Import cancelled or no entities selected."
        On Error Resume Next
        AppActivate Application.Caption
        On Error GoTo ErrHandler
        Exit Sub
    End If

    On Error Resume Next
    AppActivate Application.Caption
    On Error GoTo ErrHandler

    ' Build SAP node collection (Collection), and sets for frames/areas
    Dim sapNodeDict As Object
    Set sapNodeDict = BuildNodeDictFromSAP(SapModel)
    Dim sapNodeCol  As Collection: Set sapNodeCol = New Collection
    Dim k           As Variant
    For Each k In sapNodeDict.keys
        sapNodeCol.Add sapNodeDict(k)
    Next k

    Dim sapFrameSet As Object
    Set sapFrameSet = CreateObject("Scripting.Dictionary")
    BuildSAPFrameSet SapModel, sapFrameSet

    Dim sapAreaSet  As Object
    Set sapAreaSet = CreateObject("Scripting.Dictionary")
    BuildSAPAreaSet SapModel, sapAreaSet

    ' Prepare log sheet
    Dim wsLog       As Worksheet
    Set wsLog = PrepareLogSheet()
    Dim logRow      As Long
    logRow = wsLog.Cells(wsLog.rows.count, "A").End(xlUp).row + 1

    ' Counters
    Dim createdNodes As Long: createdNodes = 0
    Dim createdFrames As Long: createdFrames = 0
    Dim createdAreas As Long: createdAreas = 0

    ' Process selection handles one by one using a unified logic (reduce duplication)
    Dim i           As Long
    For i = LBound(handles) To UBound(handles)
        Dim h       As String: h = CStr(handles(i))
        Dim ent     As Object
        Set ent = acadDoc.HandleToObject(h)
        If ent Is Nothing Then GoTo NextHandle

        On Error Resume Next
        Dim lyr     As String: lyr = ""
        lyr = LCase$(Trim$(ent.layer))
        On Error GoTo ErrHandler

        Dim isValidLayer As Boolean
        isValidLayer = (lyr = "0" Or InStr(lyr, "dts_") > 0 Or lyr = LCase(LAYER_SPRING))
        If Not isValidLayer Then GoTo NextHandle

        Dim tName   As String: tName = LCase$(TypeName(ent))
        Dim oName   As String: oName = ""
        If HasProperty(ent, "ObjectName") Then
            On Error Resume Next
            oName = LCase$(ent.ObjectName)
            On Error GoTo ErrHandler
        End If

        ' Handle Block/Insert and Circle as points
        If InStr(tName, "block") > 0 Or InStr(tName, "insert") > 0 Then
            Dim ip  As Variant
            On Error Resume Next
            If HasProperty(ent, "InsertionPoint") Then
                ip = ent.insertionPoint
            ElseIf HasProperty(ent, "Position") Then
                ip = ent.Position
            End If
            On Error GoTo ErrHandler

            If Not IsEmpty(ip) And IsArray(ip) Then
                Dim bx As Double, by As Double, bz As Double
                bx = CDbl(ip(0)): by = CDbl(ip(1)): bz = CDbl(ip(2))

                Dim existingBlockNode As String
                existingBlockNode = GetNearestSAPNode(bx, by, bz, sapNodeCol, tolerance)

                If Len(existingBlockNode) = 0 Then
                    Dim newBlockNode As String
                    newBlockNode = CreateSAPPoint(SapModel, bx, by, bz, sapNodeCol, createdNodes)
                    If Len(newBlockNode) > 0 Then
                        ' Apply pinned restraint if in spring layer
                        If lyr = LCase(LAYER_SPRING) Then
                            ApplyPinnedRestraint SapModel, newBlockNode
                            LogStatus "Applied pinned restraint to node: " & newBlockNode
                        End If
                        ' Attach XData to CAD entity
                        AttachPointXData ent, newBlockNode, bx, by, bz, ""
                        LogToSheet wsLog, logRow, h, TypeName(ent), ent.layer, "Point", newBlockNode, "Created from block"
                        logRow = logRow + 1
                    End If
                End If
            End If

        ElseIf InStr(tName, "circle") > 0 Or InStr(oName, "circle") > 0 Then
            Dim cc  As Variant
            cc = ent.center
            Dim px As Double, py As Double, pz As Double
            px = CDbl(cc(0)): py = CDbl(cc(1)): pz = CDbl(cc(2))

            Dim nearest As String
            nearest = GetNearestSAPNode(px, py, pz, sapNodeCol, tolerance)
            If Len(nearest) = 0 Then
                Dim outName As String
                outName = CreateSAPPoint(SapModel, px, py, pz, sapNodeCol, createdNodes)
                If Len(outName) > 0 Then
                    If lyr = LCase(LAYER_SPRING) Then
                        ApplyPinnedRestraint SapModel, outName
                        LogStatus "Applied pinned restraint to node: " & outName
                    End If
                    AttachPointXData ent, outName, px, py, pz, ""
                    LogToSheet wsLog, logRow, h, TypeName(ent), ent.layer, "Point", outName, "Created from circle"
                    logRow = logRow + 1
                End If
            End If

            ' Polyline/Polyline-like as area (if closed)
        ElseIf InStr(tName, "polyline") > 0 Or InStr(tName, "lwpolyline") > 0 Or InStr(oName, "polyline") > 0 Then
            Dim isClosed As Boolean: isClosed = False
            On Error Resume Next
            If HasProperty(ent, "Closed") Then
                isClosed = CBool(ent.Closed)
            End If
            On Error GoTo ErrHandler

            If Not isClosed Then
                Dim coords As Variant
                coords = Empty
                On Error Resume Next
                If HasProperty(ent, "Coordinates") Then
                    coords = ent.Coordinates
                End If
                On Error GoTo ErrHandler

                If Not IsEmpty(coords) Then
                    If IsArray(coords) Then
                        Dim n As Long: n = UBound(coords)
                        If n >= 5 Then
                            Dim x0 As Double, y0 As Double, xn As Double, yn As Double
                            x0 = coords(0): y0 = coords(1)
                            If n Mod 3 = 2 Then
                                xn = coords(n - 2): yn = coords(n - 1)
                            Else
                                xn = coords(n - 1): yn = coords(n)
                            End If
                            If Abs(x0 - xn) < 0.0001 And Abs(y0 - yn) < 0.0001 Then
                                isClosed = True
                            End If
                        End If
                    End If
                End If
            End If

            If isClosed Then
                Dim verts As Collection
                Set verts = GetPolylineVertices(ent)
                If Not verts Is Nothing And verts.count >= 3 Then
                    Dim ptsArr() As String
                    ReDim ptsArr(0 To verts.count - 1)
                    Dim vi As Long
                    For vi = 1 To verts.count
                        Dim vpt As Variant
                        vpt = verts(vi)
                        Dim vx As Double, vy As Double, vz As Double
                        vx = CDbl(vpt(0)): vy = CDbl(vpt(1)): vz = CDbl(vpt(2))

                        Dim pName As String
                        pName = GetNearestSAPNode(vx, vy, vz, sapNodeCol, tolerance)
                        If Len(pName) = 0 Then
                            pName = CreateSAPPoint(SapModel, vx, vy, vz, sapNodeCol, createdNodes)
                        End If
                        ptsArr(vi - 1) = pName
                    Next vi

                    Dim aKey As String
                    aKey = NormalizeAreaKey(ptsArr)
                    If Not sapAreaSet.exists(aKey) Then
                        Dim newArea As String
                        Dim retA As Long
                        retA = SapModel.AreaObj.AddByPoint(UBound(ptsArr) + 1, ptsArr, newArea, "Default", "")
                        If retA = 0 Then
                            sapAreaSet.Add aKey, newArea
                            AttachAreaXData ent, newArea, "", Join(ptsArr, ",")
                            LogToSheet wsLog, logRow, h, TypeName(ent), ent.layer, "Area", newArea, "Created from polyline"
                            logRow = logRow + 1
                        End If
                    End If
                End If
            End If

            ' Line as frame
        ElseIf InStr(tName, "line") > 0 Or InStr(oName, "line") > 0 Then
            If InStr(tName, "polyline") = 0 Then
                Dim spP As Variant, epP As Variant
                spP = ent.StartPoint: epP = ent.EndPoint
                Dim x1 As Double, y1 As Double, z1 As Double, x2 As Double, y2 As Double, z2 As Double
                x1 = CDbl(spP(0)): y1 = CDbl(spP(1)): z1 = CDbl(spP(2))
                x2 = CDbl(epP(0)): y2 = CDbl(epP(1)): z2 = CDbl(epP(2))

                Dim n1 As String, n2 As String
                n1 = GetNearestSAPNode(x1, y1, z1, sapNodeCol, tolerance)
                If Len(n1) = 0 Then n1 = CreateSAPPoint(SapModel, x1, y1, z1, sapNodeCol, createdNodes)
                n2 = GetNearestSAPNode(x2, y2, z2, sapNodeCol, tolerance)
                If Len(n2) = 0 Then n2 = CreateSAPPoint(SapModel, x2, y2, z2, sapNodeCol, createdNodes)

                Dim fKey As String
                fKey = NormalizeFrameKey(n1, n2)
                If Not sapFrameSet.exists(fKey) Then
                    Dim newF As String: newF = ""
                    Dim rf As Long
                    rf = SapModel.frameObj.AddByPoint(n1, n2, newF, "Default", "")
                    If rf = 0 Then
                        sapFrameSet.Add fKey, newF
                        AttachFrameXData ent, newF, n1, n2, ""
                        LogToSheet wsLog, logRow, h, TypeName(ent), ent.layer, "Frame", newF, "Created from line"
                        logRow = logRow + 1
                    End If
                End If
            End If
        End If

NextHandle:
    Next i

    LogStatus "IMPORT SUMMARY: Nodes created (session) >= " & createdNodes

    Exit Sub

ErrHandler:
    LogStatus "ERROR in ImportSelectedEntitiesToSAP: " & err.number & " - " & err.description

    On Error Resume Next
    AppActivate Application.Caption
    On Error GoTo 0
End Sub

Private Function GetSelectionHandles_OnScreen(acadDoc As Object) As Variant
    On Error GoTo ErrHandler
    Dim ssName      As String: ssName = "DTS_IMPORT_SEL"
    On Error Resume Next
    acadDoc.SelectionSets.item(ssName).Delete
    On Error GoTo 0

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
    GetSelectionHandles_OnScreen = Empty
End Function

Private Sub ClassifySelectionHandles(acadDoc As Object, handles As Variant, selPoints As Collection, selFrames As Collection, selAreas As Collection)
    On Error GoTo ErrHandler
    Dim i           As Long
    For i = LBound(handles) To UBound(handles)
        Dim h       As String: h = CStr(handles(i))
        Dim ent     As Object
        Set ent = acadDoc.HandleToObject(h)
        If ent Is Nothing Then GoTo NextH

        On Error Resume Next
        Dim lyr     As String: lyr = ""
        lyr = LCase$(Trim$(ent.layer))
        On Error GoTo ErrHandler

        Dim isValidLayer As Boolean
        isValidLayer = (lyr = "0" Or InStr(lyr, "dts_") > 0 Or lyr = LCase(LAYER_SPRING))
        If Not isValidLayer Then GoTo NextH

        Dim tName   As String: tName = LCase$(TypeName(ent))
        Dim oName   As String: oName = ""
        If HasProperty(ent, "ObjectName") Then
            On Error Resume Next
            oName = LCase$(ent.ObjectName)
            On Error GoTo ErrHandler
        End If

        ' 1) Block/Insert as point (centroid)
        If InStr(tName, "block") > 0 Or InStr(tName, "insert") > 0 Then
            Dim insPoint As Variant
            On Error Resume Next
            If HasProperty(ent, "InsertionPoint") Then
                insPoint = ent.insertionPoint
            ElseIf HasProperty(ent, "Position") Then
                insPoint = ent.Position
            End If
            On Error GoTo ErrHandler

            If Not IsEmpty(insPoint) And IsArray(insPoint) Then
                Dim pb As Object: Set pb = CreateObject("Scripting.Dictionary")
                pb("Handle") = h
                pb("X") = CDbl(insPoint(0))
                pb("Y") = CDbl(insPoint(1))
                pb("Z") = CDbl(insPoint(2))
                pb("Layer") = lyr
                selPoints.Add pb
            End If

            ' 2) Circle as point
        ElseIf InStr(tName, "circle") > 0 Or InStr(oName, "circle") > 0 Then
            Dim p   As Object: Set p = CreateObject("Scripting.Dictionary")
            Dim c   As Variant
            c = ent.center
            p("Handle") = h
            p("X") = CDbl(c(0)): p("Y") = CDbl(c(1)): p("Z") = CDbl(c(2))
            p("Layer") = lyr
            selPoints.Add p

            ' 3) Polyline as area (if closed)
        ElseIf InStr(tName, "polyline") > 0 Or InStr(tName, "lwpolyline") > 0 Or InStr(oName, "polyline") > 0 Then
            Dim isClosed As Boolean: isClosed = False
            On Error Resume Next
            If HasProperty(ent, "Closed") Then
                isClosed = CBool(ent.Closed)
            End If
            On Error GoTo ErrHandler

            If Not isClosed Then
                Dim coords As Variant
                coords = Empty
                On Error Resume Next
                If HasProperty(ent, "Coordinates") Then
                    coords = ent.Coordinates
                End If
                On Error GoTo ErrHandler

                If Not IsEmpty(coords) Then
                    If IsArray(coords) Then
                        Dim nNums As Long
                        Dim lbC As Long, ubC As Long
                        lbC = LBound(coords): ubC = UBound(coords)
                        nNums = ubC - lbC + 1

                        Dim objTypeLocal As String
                        objTypeLocal = LCase$(TypeName(ent))
                        Dim objNameLocal As String
                        objNameLocal = ""
                        If HasProperty(ent, "ObjectName") Then
                            On Error Resume Next
                            objNameLocal = LCase$(ent.ObjectName)
                            On Error GoTo 0
                        End If

                        Dim isLWLocal As Boolean
                        isLWLocal = (InStr(objTypeLocal, "lwpolyline") > 0) Or (InStr(objNameLocal, "lwpolyline") > 0)

                        Dim stepNLocal As Long
                        If isLWLocal Then
                            stepNLocal = 2
                        Else
                            If (nNums Mod 3) = 0 Then
                                stepNLocal = 3
                            Else
                                stepNLocal = 2
                            End If
                        End If

                        If nNums >= (stepNLocal * 2 + 1) Then
                            Dim x0 As Double, y0 As Double, xn As Double, yn As Double
                            x0 = coords(lbC)
                            y0 = coords(lbC + 1)
                            xn = coords(ubC - (stepNLocal - 1))
                            yn = coords(ubC - (stepNLocal - 2))
                            If Abs(x0 - xn) < 0.0001 And Abs(y0 - yn) < 0.0001 Then
                                isClosed = True
                            End If
                        End If
                    End If
                End If
            End If

            If isClosed Then
                Dim a As Object: Set a = CreateObject("Scripting.Dictionary")
                a("Handle") = h
                Dim verts As Collection
                Set verts = GetPolylineVertices(ent)

                If Not verts Is Nothing Then
                    a("Vertices") = verts
                    a("Name") = ""
                    selAreas.Add a
                Else
                    a("Vertices") = Nothing
                    a("Name") = ""
                    selAreas.Add a
                End If
            End If

            ' 4) Line as frame (but not polyline)
        ElseIf InStr(tName, "line") > 0 Or InStr(oName, "line") > 0 Then
            If InStr(tName, "polyline") = 0 Then
                Dim f As Object: Set f = CreateObject("Scripting.Dictionary")
                Dim sp As Variant, ep As Variant
                sp = ent.StartPoint: ep = ent.EndPoint
                f("Handle") = h
                f("X1") = CDbl(sp(0)): f("Y1") = CDbl(sp(1)): f("Z1") = CDbl(sp(2))
                f("X2") = CDbl(ep(0)): f("Y2") = CDbl(ep(1)): f("Z2") = CDbl(ep(2))
                selFrames.Add f
            End If
        End If

NextH:
    Next i
    Exit Sub

ErrHandler:
Debug.Print "ERROR in ClassifySelectionHandles: " & err.description
    Resume Next
End Sub

' ----------------------------
' Existing helper functions
' ----------------------------

Private Function GetSpringData(SapModel As Object, nodeName As String) As String
    On Error Resume Next

    Dim k1 As Double, k2 As Double, k3 As Double
    Dim k4 As Double, k5 As Double, k6 As Double
    Dim ret         As Long

    ret = SapModel.pointObj.GetSpring(nodeName, k1, k2, k3, k4, k5, k6)

    If err.number = 0 And ret = 0 Then
        If k1 <> 0 Or k2 <> 0 Or k3 <> 0 Or k4 <> 0 Or k5 <> 0 Or k6 <> 0 Then
            GetSpringData = k1 & "," & k2 & "," & k3 & "," & k4 & "," & k5 & "," & k6
        End If
    End If

    On Error GoTo 0
End Function

Private Sub BuildSAPNodeList(SapModel As Object, sapNodes As Collection, scaleFactor As Double)
    On Error Resume Next

    Dim nCount As Long, names() As String
    Dim ret         As Long

    ret = SapModel.pointObj.GetNameList(nCount, names)
    If ret <> 0 Then Exit Sub

    Dim i           As Long
    For i = 0 To nCount - 1
        Dim X As Double, Y As Double, Z As Double
        SapModel.pointObj.GetCoordCartesian names(i), X, Y, Z

        Dim info    As Object
        Set info = CreateObject("Scripting.Dictionary")
        info("Name") = names(i)
        info("X") = X * scaleFactor
        info("Y") = Y * scaleFactor
        info("Z") = Z * scaleFactor

        sapNodes.Add info
    Next i

    On Error GoTo 0
End Sub

Private Sub BuildSAPFrameSet(SapModel As Object, frameSet As Object)
    On Error Resume Next

    Dim fCount As Long, fNames() As String
    SapModel.frameObj.GetNameList fCount, fNames

    Dim i           As Long
    For i = 0 To fCount - 1
        Dim p1 As String, p2 As String
        SapModel.frameObj.GetPoints fNames(i), p1, p2

        If err.number = 0 Then
            Dim k   As String
            k = NormalizeFrameKey(p1, p2)
            If Not frameSet.exists(k) Then frameSet.Add k, fNames(i)
        End If
    Next i

    On Error GoTo 0
End Sub

Private Sub BuildSAPAreaSet(SapModel As Object, areaSet As Object)
    On Error Resume Next

    Dim aCount As Long, aNames() As String
    SapModel.AreaObj.GetNameList aCount, aNames

    Dim i           As Long
    For i = 0 To aCount - 1
        Dim nPts As Long, pts() As String
        SapModel.AreaObj.GetPoints aNames(i), nPts, pts

        If err.number = 0 And nPts > 0 Then
            Dim k   As String
            k = NormalizeAreaKey(pts)
            If Not areaSet.exists(k) Then areaSet.Add k, aNames(i)
        End If
    Next i

    On Error GoTo 0
End Sub

Private Function NormalizeFrameKey(p1 As String, p2 As String) As String
    If StrComp(Trim(p1), Trim(p2), vbTextCompare) <= 0 Then
        NormalizeFrameKey = Trim(p1) & "|" & Trim(p2)
    Else
        NormalizeFrameKey = Trim(p2) & "|" & Trim(p1)
    End If
End Function

Private Function NormalizeAreaKey(pts() As String) As String
    Dim arr         As Variant
    arr = pts

    Dim i As Long, j As Long
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If StrComp(arr(i), arr(j), vbTextCompare) > 0 Then
                Dim tmp As String
                tmp = arr(i)
                arr(i) = arr(j)
                arr(j) = tmp
            End If
        Next j
    Next i

    Dim key         As String
    For i = LBound(arr) To UBound(arr)
        If i > LBound(arr) Then key = key & "|"
        key = key & arr(i)
    Next i

    NormalizeAreaKey = key
End Function

Private Function GetNearestSAPNode(X As Double, Y As Double, Z As Double, sapNodes As Collection, tolerance As Double) As String
    Dim bestName    As String
    bestName = ""
    Dim bestD2      As Double
    bestD2 = tolerance * tolerance

    Dim i           As Long
    For i = 1 To sapNodes.count
        Dim ni      As Object
        Set ni = sapNodes(i)

        Dim dx As Double, dy As Double, dz As Double
        dx = ni("X") - X
        dy = ni("Y") - Y
        dz = ni("Z") - Z

        Dim d2      As Double
        d2 = dx * dx + dy * dy + dz * dz

        If d2 <= bestD2 Then
            bestD2 = d2
            bestName = ni("Name")
        End If
    Next i

    GetNearestSAPNode = bestName
End Function

Private Function CreateSAPPoint(SapModel As Object, X As Double, Y As Double, Z As Double, sapNodes As Collection, ByRef createdCount As Long) As String
    On Error GoTo ErrHandler

    ' Try letting SAP2000 assign the node name automatically first (preferred).
    ' If SAP fails to assign a name, fall back to generating a timestamp-based name.

    Dim baseName    As String
    baseName = "CAD_N_" & Format(Now, "yyyymmddHHMMSS") & "_" & CStr(createdCount + 1)

    Dim ret         As Long
    Dim outName     As String
    outName = ""

    ' First attempt: ask SAP to auto-generate the name (pass empty suggested name)
    ret = SapModel.pointObj.AddCartesian(X, Y, Z, outName, "", "Global", False, 0)
    If ret = 0 Then
        ' SAP returned success - use outName if provided
        If Len(Trim$(outName)) > 0 Then
            CreateSAPPoint = outName
            Dim niAuto As Object
            Set niAuto = CreateObject("Scripting.Dictionary")
            niAuto("Name") = outName
            niAuto("X") = X
            niAuto("Y") = Y
            niAuto("Z") = Z
            On Error Resume Next
            sapNodes.Add niAuto
            On Error GoTo ErrHandler
            createdCount = createdCount + 1
            Exit Function
        End If
    End If

    ' If we reach here, SAP did not provide a name or AddCartesian failed.
    ' Try adding with a generated baseName as fallback.
    ret = SapModel.pointObj.AddCartesian(X, Y, Z, outName, baseName, "Global", False, 0)
    If ret <> 0 Then
        ' second fallback: try with empty suggested name again but ignore ret
        ret = SapModel.pointObj.AddCartesian(X, Y, Z, outName, "", "Global", False, 0)
    End If

    Dim newName     As String
    If ret = 0 And Len(Trim$(outName)) > 0 Then
        newName = outName
    Else
        newName = baseName
    End If

    Dim ni          As Object
    Set ni = CreateObject("Scripting.Dictionary")
    ni("Name") = newName
    ni("X") = X
    ni("Y") = Y
    ni("Z") = Z
    On Error Resume Next
    sapNodes.Add ni
    On Error GoTo ErrHandler

    createdCount = createdCount + 1
    CreateSAPPoint = newName

    On Error GoTo 0
    Exit Function

ErrHandler:
    On Error Resume Next
    Dim ni2         As Object
    Set ni2 = CreateObject("Scripting.Dictionary")
    ni2("Name") = baseName
    ni2("X") = X
    ni2("Y") = Y
    ni2("Z") = Z
    sapNodes.Add ni2
    createdCount = createdCount + 1
    CreateSAPPoint = baseName
    On Error GoTo 0
End Function

Private Function GetPolylineVertices(ent As Object) As Collection
    On Error GoTo ErrHandler
    Dim verts       As Collection
    Set verts = New Collection

    Dim coords      As Variant
    coords = Empty

    If HasProperty(ent, "Coordinates") Then
        On Error Resume Next
        coords = ent.Coordinates
        On Error GoTo 0
    End If

    If IsEmpty(coords) Then
        Set GetPolylineVertices = verts
        Exit Function
    End If

    Dim lb As Long, ub As Long
    lb = LBound(coords): ub = UBound(coords)
    Dim nNums       As Long
    nNums = ub - lb + 1

    ' Decide stepN robustly:
    ' - If this is an LWPolyline (or object name suggests lwpolyline), treat as 2D pairs
    '   (third coordinate in Coordinates for LWPolyline can be bulge, not Z).
    ' - Otherwise, if total numbers divisible by 3 -> assume 3D (x,y,z)
    ' - Else treat as 2D pairs (x,y)
    Dim objType     As String
    objType = LCase$(TypeName(ent))
    Dim objName     As String
    objName = ""
    If HasProperty(ent, "ObjectName") Then
        On Error Resume Next
        objName = LCase$(ent.ObjectName)
        On Error GoTo 0
    End If

    Dim isLW        As Boolean
    isLW = (InStr(objType, "lwpolyline") > 0) Or (InStr(objName, "lwpolyline") > 0) Or (InStr(objName, "lwp") > 0)

    Dim stepN       As Long
    If isLW Then
        stepN = 2    ' treat as X,Y pairs and use entity Elevation for Z
    Else
        If (nNums Mod 3) = 0 Then
            stepN = 3    ' x,y,z triple
        Else
            stepN = 2    ' x,y pairs
        End If
    End If

    Dim nVerts      As Long
    nVerts = nNums \ stepN

    Dim i           As Long
    For i = 0 To nVerts - 1
        Dim vx As Double, vy As Double, vz As Double
        vx = CDbl(coords(lb + i * stepN))
        vy = CDbl(coords(lb + i * stepN + 1))
        If stepN = 3 Then
            vz = CDbl(coords(lb + i * stepN + 2))
        Else
            ' stepN = 2: use entity Elevation if available, otherwise 0
            If HasProperty(ent, "Elevation") Then
                On Error Resume Next
                vz = CDbl(ent.elevation)
                On Error GoTo 0
            Else
                vz = 0
            End If
        End If

        Dim vpt(0 To 2) As Double
        vpt(0) = vx: vpt(1) = vy: vpt(2) = vz
        verts.Add vpt
    Next i

    ' Remove duplicate last vertex if closed (compare with tolerance)
    If verts.count > 1 Then
        Dim firstV As Variant, lastV As Variant
        firstV = verts(1)
        lastV = verts(verts.count)
        If Abs(firstV(0) - lastV(0)) < 0.0001 And Abs(firstV(1) - lastV(1)) < 0.0001 And Abs(firstV(2) - lastV(2)) < 0.0001 Then
            verts.Remove verts.count
        End If
    End If

    Set GetPolylineVertices = verts
    Exit Function

ErrHandler:
    Set GetPolylineVertices = Nothing
    Exit Function
End Function
Private Sub RegisterDataApp(acadDoc As Object)
    On Error Resume Next
    acadDoc.RegisteredApplications.Add "DTS_APP"
    On Error GoTo 0
End Sub

Private Sub EnsureLayerExists(acadDoc As Object, layerName As String, colorIndex As Long)
    On Error Resume Next
    Dim lay         As Object
    Set lay = acadDoc.layers.item(layerName)
    If err.number <> 0 Then
        err.Clear
        Set lay = acadDoc.layers.Add(layerName)
        lay.color = colorIndex
    End If
    On Error GoTo 0
End Sub

Private Function HasProperty(obj As Object, propName As String) As Boolean
    On Error Resume Next
    Dim tmp
    tmp = CallByName(obj, propName, VbGet)
    If err.number = 0 Then
        HasProperty = True
    Else
        HasProperty = False
        err.Clear
    End If
    On Error GoTo 0
End Function

Public Sub LogStatus(msg As String)
    On Error Resume Next
    If Not StatusCallback Is Nothing Then
        StatusCallback.SetStatus msg
    End If
Debug.Print msg
    On Error GoTo 0
End Sub
' ============================
' HELPER - convert any string/variant array to 1-based Variant array
' ============================
Private Function ConvertTo1BasedVariantArray(arr As Variant) As Variant
    On Error GoTo ErrHandler
    If IsEmpty(arr) Then
        ConvertTo1BasedVariantArray = arr
        Exit Function
    End If

    If Not IsArray(arr) Then
        Dim singleV(1 To 1) As Variant
        singleV(1) = arr
        ConvertTo1BasedVariantArray = singleV
        Exit Function
    End If

    Dim lb As Long, ub As Long
    lb = LBound(arr): ub = UBound(arr)
    Dim n           As Long: n = ub - lb + 1
    Dim out()       As Variant
    ReDim out(1 To n)

    Dim i           As Long
    For i = lb To ub
        out(i - lb + 1) = arr(i)
    Next i

    ConvertTo1BasedVariantArray = out
    Exit Function

ErrHandler:
    ConvertTo1BasedVariantArray = arr
End Function

