Attribute VB_Name = "M12_DrawColumnAreas"
Option Explicit
'===============================================================
' Module: M12_DrawColumnAreas
' Purpose: Draw Area objects representing column cross-sections at story levels
' FIXED VERSION:
' - Fixed UserClickedOK check (must call as function with parentheses)
' - Added default settings for standalone execution
' - Added graceful handling when dialog is cancelled
'===============================================================

' Constants
Private Const COORD_TOLERANCE As Double = 0.001  ' mm tolerance for coordinate comparison
Private Const AREA_SECTION_NONE As String = "None"
Private Const COLXSEC_PREFIX As String = "ColXSec_"
Private Const MIN_COLUMN_ZLEN As Double = 1#   ' minimal vertical length (mm) to be considered a column
Private Const PI    As Double = 3.14159265358979
Private Const THICKNESS_SCALE As Double = 0.1  ' Scale factor for thickness visibility (% of max(B,H))

'===============================================================
' PUBLIC SETTINGS TYPE & GLOBAL
'===============================================================
Public Type ColumnDrawSettings
    storyRange      As String
    AreaPrefix      As String
    AddFloorSuffix  As Boolean
    DrawPunching    As Boolean
    PunchingOutputType As String   ' "AREA" or "LINES"
    CoverSettings   As Object     ' Scripting.Dictionary: keys "AUTO", "STORY_n", "SECTION_name"
End Type

Public g_Settings   As Object
Public g_StoryElevations As Object  ' dictionary of elevations -> story names (set at runtime)

'===============================================================
' PUBLIC API - Main Entry Point
'===============================================================

Public Sub DrawColumnCrossSectionsWithSettings(settingsDict As Object)
    On Error GoTo ErrorHandler

    ' Validate input
    If settingsDict Is Nothing Then
        MsgBox "No settings provided.", vbExclamation, "Settings Error"
        Exit Sub
    End If

    ' Connect to SAP2000
    If Not ConnectSAP2000() Then
        MsgBox "Could not connect to SAP2000.", vbCritical, "Connection Error"
        Exit Sub
    End If

    ' Set units to mm
    SapModel.SetPresentUnits 5

    ' Get gridline data (elevation in mm)
    Dim gridData    As Object
    Set gridData = GetGridlineElevations()

    If gridData Is Nothing Or gridData.count = 0 Then
        MsgBox "No gridline elevations found. Please ensure model contains vertical frames or gridline data.", vbExclamation, "No Data"
        Exit Sub
    End If
    Set g_StoryElevations = gridData

    ' Store settings
    Set g_Settings = settingsDict

    ' Debug: Log settings
    LogMsg "Settings received:"
    LogMsg "  - storyRange: " & GetSettingString("storyRange", "")
    LogMsg "  - AreaPrefix: " & GetSettingString("AreaPrefix", "")
    LogMsg "  - AddFloorSuffix: " & GetSettingBool("AddFloorSuffix", False)
    LogMsg "  - DrawPunching: " & GetSettingBool("DrawPunching", False)

    If g_Settings.exists("SectionPrefixes") Then
        Dim prefixDict As Object
        Set prefixDict = g_Settings("SectionPrefixes")
        LogMsg "  - SectionPrefixes count: " & prefixDict.count

        ' Log all prefixes
        Dim k       As Variant
        For Each k In prefixDict.keys
            LogMsg "    * " & k & " = " & prefixDict(k)
        Next k
    Else
        LogMsg "  - WARNING: SectionPrefixes NOT FOUND in settings!"
    End If

    ' Determine story selection and draw
    Dim storyRangeStr As String
    storyRangeStr = GetSettingString("storyRange", "")

    If Trim$(storyRangeStr) = "" Then
        Call DrawAllStories(gridData)
    Else
        Call DrawSpecificStory(storyRangeStr, gridData)
    End If

    SapModel.View.RefreshView 0, False
    MsgBox "Column cross-section drawing completed successfully!", vbInformation, "Complete"
    Exit Sub

ErrorHandler:
    MsgBox "ERROR in DrawColumnCrossSectionsWithSettings: " & err.description, vbCritical, "Error"
End Sub


Public Sub DrawColumnCrossSections()
    MsgBox "Please use the Column Section Form to configure settings." & vbCrLf & _
            "Run: show_frmColumnSection", vbInformation, "Use Form Instead"
End Sub
Private Sub InitializeDefaultSettings()
    On Error Resume Next

    Set g_Settings = CreateObject("Scripting.Dictionary")

    ' General settings - defaults for standalone mode
    g_Settings.Add "storyRange", ""  ' Draw all stories
    g_Settings.Add "AreaPrefix", "ColSec_"
    g_Settings.Add "AddFloorSuffix", True
    g_Settings.Add "DrawPunching", False  ' Disabled by default
    g_Settings.Add "PunchingOutputType", "AREA"

    ' Default section prefixes
    Dim prefixDict  As Object
    Set prefixDict = CreateObject("Scripting.Dictionary")
    prefixDict.Add "RC", "RC_COL_"
    prefixDict.Add "I", "STEEL_I_"
    prefixDict.Add "H", "STEEL_H_"
    prefixDict.Add "PIPE", "STEEL_PIPE_"
    prefixDict.Add "BOX", "STEEL_BOX_"
    prefixDict.Add "CHANNEL", "STEEL_CH_"
    prefixDict.Add "TEE", "STEEL_TEE_"
    prefixDict.Add "ANGLE", "STEEL_ANG_"
    prefixDict.Add "CIRCLE", "STEEL_CIR_"
    prefixDict.Add "DEFAULT", "ColSec_"
    g_Settings.Add "SectionPrefixes", prefixDict

    ' Default cover settings (for punching shear if enabled later)
    Dim coverDict   As Object
    Set coverDict = CreateObject("Scripting.Dictionary")
    Dim arrAuto(0 To 1) As Double
    arrAuto(0) = 50  ' Top cover
    arrAuto(1) = 50  ' Bottom cover
    coverDict.Add "AUTO", arrAuto
    g_Settings.Add "CoverSettings", coverDict

    LogMsg "Initialized default settings for standalone execution"
End Sub
'===============================================================
' ALTERNATIVE: Standalone entry point with default settings
'===============================================================
Public Sub DrawColumnCrossSection_NoDialog()
    On Error GoTo ErrorHandler

    MsgBox "This function is for testing only." & vbCrLf & _
            "Please use: show_frmColumnSection", vbInformation, "Use Form Instead"
    Exit Sub

ErrorHandler:
    MsgBox "ERROR in DrawColumnCrossSection_NoDialog: " & err.description, vbCritical, "Error"
End Sub

'===============================================================
' GRIDLINE DATA FUNCTIONS
'===============================================================

' Get gridline elevations from model endpoints first, then map names from Gridline sheet (if present).
' Returns Dictionary: Key = elevation (mm), Value = story name
Public Function GetGridlineElevations() As Object
    On Error Resume Next

    Set GetGridlineElevations = CreateObject("Scripting.Dictionary")
    Dim elevDict    As Object
    Set elevDict = GetGridlineElevations

    ' Ensure SAP units expected to be mm by caller (DrawColumnCrossSections sets it)

    ' 1) Extract elevations from model frame endpoints (only endpoints are added)
    Call ExtractElevationsFromPointsAndFrames(elevDict)

    ' 2) Try to read Gridline sheet (support both "Girdline" and "Gridline")
    Dim ws          As Worksheet
    Set ws = Nothing
    On Error Resume Next
    If SheetExists(ThisWorkbook, "Girdline") Then
        Set ws = ThisWorkbook.Worksheets("Girdline")
    ElseIf SheetExists(ThisWorkbook, "Gridline") Then
        Set ws = ThisWorkbook.Worksheets("Gridline")
    End If
    On Error GoTo 0

    If Not ws Is Nothing Then
        ' Map sheet names to existing model elevations (do not add sheet-only elevations)
        Call ReadGridlineFromSheet(ws, elevDict)
    Else
        ' Optionally attempt exporter that may create sheet; then map
        If VbProcedureExists("ExportGridLinesToGirdlineSheet") Then
            On Error Resume Next
            ExportGridLinesToGirdlineSheet
            On Error GoTo 0
            If SheetExists(ThisWorkbook, "Girdline") Then
                Set ws = ThisWorkbook.Worksheets("Girdline")
                Call ReadGridlineFromSheet(ws, elevDict)
            ElseIf SheetExists(ThisWorkbook, "Gridline") Then
                Set ws = ThisWorkbook.Worksheets("Gridline")
                Call ReadGridlineFromSheet(ws, elevDict)
            End If
        End If
    End If

    ' 3) If still empty (no frame endpoints found), try database tables as fallback
    If elevDict.count = 0 Then
        Call TryReadStoriesFromDatabase(elevDict)
    End If

    On Error GoTo 0
End Function

' Helper: check if sheet exists
Private Function SheetExists(wb As Workbook, sheetName As String) As Boolean
    On Error Resume Next
    Dim sh          As Worksheet
    Set sh = wb.Worksheets(sheetName)
    SheetExists = Not sh Is Nothing
    On Error GoTo 0
End Function

' Read gridline data from Excel sheet BUT map sheet elevations to existing model elevations only.
' If sheet elevation doesn't match any model endpoint elevation (within COORD_TOLERANCE) it's ignored.
Private Sub ReadGridlineFromSheet(ws As Worksheet, elevDict As Object)
    On Error Resume Next

    If elevDict Is Nothing Then Exit Sub

    Dim lastRow     As Long
    lastRow = ws.Cells(ws.rows.count, "B").End(xlUp).row
    If lastRow < 5 Then Exit Sub

    Dim i           As Long
    For i = 5 To lastRow
        Dim axisDir As String
        axisDir = Trim$(CStr(ws.Cells(i, 2).Value))    ' column B

        If UCase$(axisDir) = "Z" Then
            Dim gridName As String
            Dim coordVal As Variant
            gridName = Trim$(CStr(ws.Cells(i, 3).Value))    ' column C
            coordVal = ws.Cells(i, 4).Value    ' column D

            If IsNumeric(coordVal) Then
                Dim coordMM As Double
                coordMM = CDbl(coordVal)    ' assume sheet values in mm

                ' Find nearest existing elevation in elevDict within tolerance
                Dim matchedKey As Variant
                matchedKey = Empty
                Dim k As Variant
                For Each k In elevDict.keys
                    If Abs(CDbl(k) - coordMM) <= COORD_TOLERANCE Then
                        matchedKey = k
                        Exit For
                    End If
                Next k

                If Not IsEmpty(matchedKey) Then
                    ' Update name for matched elevation (prefer sheet name)
                    If Trim$(gridName) = "" Then
                        gridName = "Story_" & Format(CDbl(matchedKey), "0.0")
                    End If
                    elevDict.Remove matchedKey
                    elevDict.Add CDbl(matchedKey), gridName
                    LogMsg "Mapped sheet grid '" & gridName & "' to model elevation " & Format(CDbl(matchedKey), "0.0") & " mm"
                Else
                    ' No match: sheet elevation does not correspond to any model column endpoint -> ignore
                    LogMsg "Ignored sheet grid at " & Format(coordMM, "0.0") & " mm (no column endpoint found at this elevation)"
                End If
            End If
        End If
    Next i

    On Error GoTo 0
End Sub

' Read gridline data directly from SAP2000 (try DB tables then fallback to points/frames)
Private Sub ReadGridlineFromSAP(elevDict As Object)
    On Error Resume Next

    ' 1) Try to read story table from DatabaseTables (best effort)
    If TryReadStoriesFromDatabase(elevDict) Then
        ' success
        Exit Sub
    End If

    ' 2) Fallback: collect elevations from points and frame endpoints
    Call ExtractElevationsFromPointsAndFrames(elevDict)

    On Error GoTo 0
End Sub

' Try to read stories / elevations from DatabaseTables
' Returns True if succeeded and added at least one elevation
Private Function TryReadStoriesFromDatabase(elevDict As Object) As Boolean
    On Error Resume Next
    TryReadStoriesFromDatabase = False

    Dim numRows As Long, numCols As Long
    Dim tableData   As Variant
    Dim triedNames  As Variant
    triedNames = Array("Stories", "Story Definitions", "Story", "Story Data")

    Dim i           As Long
    For i = LBound(triedNames) To UBound(triedNames)
        numRows = 0: numCols = 0
        tableData = Empty
        err.Clear
        SapModel.DatabaseTables.GetTableForDisplay CStr(triedNames(i)), numRows, numCols, tableData
        If err.number = 0 Then
            If Not IsEmpty(tableData) Then
                Dim r As Long
                For r = 0 To numRows - 1
                    ' Try to find a numeric elevation within the row data
                    Dim c As Long
                    For c = 0 To numCols - 1
                        Dim cellVal As Variant
                        cellVal = tableData(r, c)
                        If IsNumeric(cellVal) Then
                            Dim elevMM As Double
                            ' Database table values are assumed in current SAP units; we set units to mm earlier
                            elevMM = CDbl(cellVal)
                            If Not elevDict.exists(elevMM) Then
                                elevDict.Add elevMM, "Story_" & Format(elevMM, "0.0")
                            End If
                            ' we break c loop and continue next row
                            Exit For
                        End If
                    Next c
                Next r
                If elevDict.count > 0 Then
                    TryReadStoriesFromDatabase = True
                    Exit Function
                End If
            End If
        End If
        err.Clear
    Next i

    On Error GoTo 0
End Function

' Extract elevations (Z) from frame endpoints as a robust model-based source.
' Only elevations that are actual frame endpoints (i.e., a column starts or ends there) are added.
' This prevents "floating" elevations that merely cross mid-span from being treated as stories.
Private Sub ExtractElevationsFromPointsAndFrames(elevDict As Object)
    On Error Resume Next

    If elevDict Is Nothing Then Set elevDict = CreateObject("Scripting.Dictionary")

    ' Temporary dictionaries:
    Dim endpointDict As Object
    Set endpointDict = CreateObject("Scripting.Dictionary")  ' key: rounded string, value: representative double z

    Dim numFrames   As Long
    Dim frameNames() As String
    Dim ret         As Long
    ret = SapModel.frameObj.GetNameList(numFrames, frameNames)
    If ret <> 0 Or numFrames = 0 Then
        ' No frames found -> nothing to add
        Exit Sub
    End If

    Dim i           As Long
    For i = 0 To numFrames - 1
        Dim fn      As String
        fn = frameNames(i)

        Dim p1 As String, p2 As String
        ret = SapModel.frameObj.GetPoints(fn, p1, p2)
        If ret <> 0 Then GoTo NextFrame

        Dim x1 As Double, y1 As Double, z1 As Double
        Dim x2 As Double, y2 As Double, z2 As Double
        If SapModel.pointObj.GetCoordCartesian(p1, x1, y1, z1) <> 0 Then GoTo NextFrame
        If SapModel.pointObj.GetCoordCartesian(p2, x2, y2, z2) <> 0 Then GoTo NextFrame

        ' Only consider vertical frames (columns)
        If (Abs(x1 - x2) > COORD_TOLERANCE) Or (Abs(y1 - y2) > COORD_TOLERANCE) Then GoTo NextFrame

        ' Add both endpoints Z (rounded key)
        Dim za As Double, zb As Double
        za = z1: zb = z2

        Dim kStrA As String, kStrB As String
        kStrA = Format(za, "0.000000")
        kStrB = Format(zb, "0.000000")

        If Not endpointDict.exists(kStrA) Then
            endpointDict.Add kStrA, za
        End If
        If Not endpointDict.exists(kStrB) Then
            endpointDict.Add kStrB, zb
        End If

NextFrame:
    Next i

    ' Populate elevDict only with these endpoint elevations (unique)
    Dim k           As Variant
    For Each k In endpointDict.keys
        Dim zVal    As Double
        zVal = CDbl(endpointDict(k))
        ' Add using AddElevationToDictRounded to preserve merging behavior
        Call AddElevationToDictRounded(elevDict, zVal)
    Next k

    On Error GoTo 0
End Sub

' Add elevation to dictionary using rounding/merging within tolerance to avoid duplicates
Private Sub AddElevationToDictRounded(elevDict As Object, ByVal elev As Double, Optional ByVal gridName As String = "")
    On Error Resume Next

    Dim key         As Variant
    For Each key In elevDict.keys
        If Abs(CDbl(key) - elev) <= COORD_TOLERANCE Then
            ' existing key is close: update name if current name is empty or looks like auto-generated
            Dim existingName As String
            existingName = CStr(elevDict(key))
            If Trim$(existingName) = "" Or Left$(existingName, 6) = "Story_" Then
                If Trim$(gridName) <> "" Then
                    elevDict.Remove key
                    elevDict.Add elev, gridName
                End If
            End If
            Exit Sub
        End If
    Next key

    ' Not found; add with provided name or generated name
    If gridName = "" Then
        gridName = "Story_" & Format(elev, "0.0")
    End If
    elevDict.Add elev, gridName

    On Error GoTo 0
End Sub

' Get formatted story list for display
Public Function GetStoryListText(gridData As Object) As String
    Dim result      As String
    Dim n           As Long
    n = gridData.count
    If n = 0 Then
        GetStoryListText = "No stories found."
        Exit Function
    End If

    Dim elevations() As Double
    ReDim elevations(1 To n)
    Dim idx         As Long: idx = 1
    Dim k           As Variant
    For Each k In gridData.keys
        elevations(idx) = CDbl(k)
        idx = idx + 1
    Next k

    ' sort ascending
    Dim i As Long, j As Long, tmp As Double
    For i = 1 To n - 1
        For j = i + 1 To n
            If elevations(i) > elevations(j) Then
                tmp = elevations(i)
                elevations(i) = elevations(j)
                elevations(j) = tmp
            End If
        Next j
    Next i

    ' Build numbered list
    For i = 1 To n
        Dim elev    As Double
        elev = elevations(i)
        Dim Name    As String
        Name = CStr(gridData(elev))
        result = result & i & ": " & Name & " (" & Format(elev, "0.0") & " mm)" & vbCrLf
    Next i

    GetStoryListText = result
End Function

' Get section dimensions (width and depth in mm)
' Also returns section shape type and detail parameters
' Uses proper ByRef variables for SAP2000 PropFrame Get* APIs to avoid ByRef argument mismatch.
Private Function GetSectionDimensions(sectionName As String, ByRef Width As Double, ByRef Depth As Double, _
        Optional ByRef shapeType As String = "", _
        Optional ByRef shapeParams As Variant) As Boolean
    On Error Resume Next
    GetSectionDimensions = False
    Dim ret         As Long
    shapeType = "RECT"  ' Default to rectangle

    ' Initialize shapeParams as empty array
    ReDim shapeParams(0 To 0)

    ' Common out variables for PropFrame API calls (ByRef)
    Dim FileName    As String
    Dim matProp     As String
    Dim color       As Long
    Dim notes       As String
    Dim guid        As String

    ' Local dimension variables
    Dim t3 As Double, t2 As Double
    Dim tf As Double, tw As Double, t2b As Double, tfb As Double
    Dim Diameter    As Double
    Dim outerDiam As Double, Thickness As Double

    ' Try Rectangle: GetRectangle(Name, FileName, MatProp, t3, t2, Color, Notes, GUID)
    ret = SapModel.PropFrame.GetRectangle(sectionName, FileName, matProp, t3, t2, color, notes, guid)
    If ret = 0 Then
        Depth = t3   ' local 3 axis (depth)
        Width = t2   ' local 2 axis (width)
        shapeType = "RECT"
        GetSectionDimensions = True
        Exit Function
    End If
    err.Clear

    ' Try Circle: GetCircle(Name, FileName, MatProp, Diameter, Color, Notes, GUID)
    ret = SapModel.PropFrame.GetCircle(sectionName, FileName, matProp, Diameter, color, notes, guid)
    If ret = 0 Then
        Depth = Diameter
        Width = Diameter
        shapeType = "CIRCLE"
        ReDim shapeParams(0 To 0)
        shapeParams(0) = Diameter
        GetSectionDimensions = True
        Exit Function
    End If
    err.Clear

    ' Try Pipe: GetPipe(Name, FileName, MatProp, OuterDiam, Thickness, Color, Notes, GUID)
    ret = SapModel.PropFrame.GetPipe(sectionName, FileName, matProp, outerDiam, Thickness, color, notes, guid)
    If ret = 0 Then
        Depth = outerDiam
        Width = outerDiam
        shapeType = "PIPE"
        ReDim shapeParams(0 To 1)
        shapeParams(0) = outerDiam
        shapeParams(1) = Thickness
        GetSectionDimensions = True
        Exit Function
    End If
    err.Clear

    ' Try I-Section:
    ' GetISection(Name, FileName, MatProp, t3, t2, tf, tw, t2b, tfb, Color, Notes, GUID)
    ret = SapModel.PropFrame.GetISection(sectionName, FileName, matProp, t3, t2, tf, tw, t2b, tfb, color, notes, guid)
    If ret = 0 Then
        Depth = t3
        Width = IIf(t2 > t2b, t2, t2b)
        shapeType = "I"
        ReDim shapeParams(0 To 5)
        shapeParams(0) = t3   ' depth
        shapeParams(1) = t2   ' top flange width
        shapeParams(2) = tf   ' top flange thickness
        shapeParams(3) = tw   ' web thickness
        shapeParams(4) = t2b  ' bottom flange width
        shapeParams(5) = tfb  ' bottom flange thickness
        GetSectionDimensions = True
        Exit Function
    End If
    err.Clear

    ' Try Tee:
    ' GetTee(Name, FileName, MatProp, t3, t2, tf, tw, Color, Notes, GUID)
    ret = SapModel.PropFrame.GetTee(sectionName, FileName, matProp, t3, t2, tf, tw, color, notes, guid)
    If ret = 0 Then
        Depth = t3
        Width = t2
        shapeType = "TEE"
        ReDim shapeParams(0 To 3)
        shapeParams(0) = t3
        shapeParams(1) = t2
        shapeParams(2) = tf
        shapeParams(3) = tw
        GetSectionDimensions = True
        Exit Function
    End If
    err.Clear

    ' Try Angle:
    ' GetAngle(Name, FileName, MatProp, t3, t2, tf, tw, Color, Notes, GUID)
    ret = SapModel.PropFrame.GetAngle(sectionName, FileName, matProp, t3, t2, tf, tw, color, notes, guid)
    If ret = 0 Then
        Depth = t3
        Width = t2
        shapeType = "ANGLE"
        ReDim shapeParams(0 To 3)
        shapeParams(0) = t3
        shapeParams(1) = t2
        shapeParams(2) = tf
        shapeParams(3) = tw
        GetSectionDimensions = True
        Exit Function
    End If
    err.Clear

    ' Try Channel:
    ' GetChannel(Name, FileName, MatProp, t3, t2, tf, tw, Color, Notes, GUID)
    ret = SapModel.PropFrame.GetChannel(sectionName, FileName, matProp, t3, t2, tf, tw, color, notes, guid)
    If ret = 0 Then
        Depth = t3
        Width = t2
        shapeType = "CHANNEL"
        ReDim shapeParams(0 To 3)
        shapeParams(0) = t3
        shapeParams(1) = t2
        shapeParams(2) = tf
        shapeParams(3) = tw
        GetSectionDimensions = True
        Exit Function
    End If
    err.Clear

    ' Try Tube (box):
    ' GetTube(Name, FileName, MatProp, t3, t2, tf, tw, Color, Notes, GUID)
    ret = SapModel.PropFrame.GetTube(sectionName, FileName, matProp, t3, t2, tf, tw, color, notes, guid)
    If ret = 0 Then
        Depth = t3
        Width = t2
        shapeType = "BOX"
        ReDim shapeParams(0 To 3)
        shapeParams(0) = t3
        shapeParams(1) = t2
        shapeParams(2) = tf  ' thickness in 3 direction
        shapeParams(3) = tw  ' thickness in 2 direction
        GetSectionDimensions = True
        Exit Function
    End If
    err.Clear

    ' No supported type matched
    GetSectionDimensions = False
    On Error GoTo 0
End Function

'===============================================================
' DRAWING FUNCTIONS
'===============================================================

' Draw column cross-sections for all stories
' Process elevations in descending order (top-down)
Private Sub DrawAllStories(gridData As Object)
    Dim elevations() As Double
    Dim n           As Long
    n = gridData.count
    If n = 0 Then Exit Sub

    ReDim elevations(1 To n)
    Dim i As Long, idx As Long
    idx = 1
    Dim k           As Variant
    For Each k In gridData.keys
        elevations(idx) = CDbl(k)
        idx = idx + 1
    Next k

    ' sort descending (top-down)
    Dim j As Long, tmp As Double
    For i = 1 To n - 1
        For j = i + 1 To n
            If elevations(i) < elevations(j) Then
                tmp = elevations(i): elevations(i) = elevations(j): elevations(j) = tmp
            End If
        Next j
    Next i

    ' Process each elevation from top to bottom
    For i = 1 To n
        Call DrawColumnAreasAtElevation(elevations(i), CStr(gridData(elevations(i))))
    Next i
End Sub

' Draw column cross-sections for specific story or elevation
Private Sub DrawSpecificStory(userInput As String, gridData As Object)
    On Error Resume Next

    Dim elevationsToDraw As Collection
    Set elevationsToDraw = ResolveSelectionToElevations(userInput, gridData)

    If elevationsToDraw Is Nothing Or elevationsToDraw.count = 0 Then
        MsgBox "No matching stories found for input: " & userInput, vbExclamation, "Not Found"
        Exit Sub
    End If

    Dim ev          As Variant
    For Each ev In elevationsToDraw
        Call DrawColumnAreasAtElevation(CDbl(ev), CStr(gridData(CDbl(ev))))
    Next ev

    On Error GoTo 0
End Sub

' Draw column cross-sections at specific elevation
Private Sub DrawColumnAreasAtElevation(elevation As Double, storyName As String)
    On Error GoTo ErrorHandler

    LogMsg "Drawing column cross-sections for " & storyName & " at elevation " & Format(elevation, "0.0") & " mm"

    ' Find all column frames at this elevation
    Dim columnFrames As Collection
    Set columnFrames = FindColumnFramesAtElevation(elevation)

    If columnFrames.count = 0 Then
        LogMsg "No column frames found at elevation " & Format(elevation, "0.0") & " mm"
        Exit Sub
    End If

    ' Create dictionary to track created area centers for this elevation to avoid duplicates
    Dim createdCenters As Object
    Set createdCenters = CreateObject("Scripting.Dictionary")

    ' Process each column frame
    Dim frameName   As Variant
    Dim successCount As Long
    successCount = 0

    For Each frameName In columnFrames
        ' Determine if frame ends at this elevation (maxZ ~= elevation)
        Dim useBelow As Boolean
        useBelow = False
        On Error Resume Next
        Dim pt1 As String, pt2 As String
        Dim x1 As Double, y1 As Double, z1 As Double
        Dim x2 As Double, y2 As Double, z2 As Double
        If SapModel.frameObj.GetPoints(CStr(frameName), pt1, pt2) = 0 Then
            If SapModel.pointObj.GetCoordCartesian(pt1, x1, y1, z1) = 0 And SapModel.pointObj.GetCoordCartesian(pt2, x2, y2, z2) = 0 Then
                Dim minZ As Double, maxZ As Double
                If z1 <= z2 Then
                    minZ = z1: maxZ = z2
                Else
                    minZ = z2: maxZ = z1
                End If
                If Abs(maxZ - elevation) <= COORD_TOLERANCE Then
                    useBelow = True
                Else
                    useBelow = False
                End If
            End If
        End If
        On Error GoTo ErrorHandler

        If ProcessColumnFrame(CStr(frameName), elevation, useBelow, createdCenters) Then
            successCount = successCount + 1
        End If
    Next frameName

    LogMsg "Created " & successCount & " column cross-section areas at " & storyName
    Exit Sub

ErrorHandler:
    LogMsg "ERROR in DrawColumnAreasAtElevation: " & err.description
End Sub

' Find column frames at elevation.
' Returns frames in order: starters -> passThrough -> enders (but excludes enders that have continuation above)
Private Function FindColumnFramesAtElevation(elevation As Double) As Collection
    Set FindColumnFramesAtElevation = New Collection
    On Error Resume Next

    Dim ret         As Long
    Dim numFrames   As Long
    Dim frameNames() As String

    ret = SapModel.frameObj.GetNameList(numFrames, frameNames)
    If ret <> 0 Or numFrames = 0 Then Exit Function

    ' Dictionary to map location -> frame type: "S" (starter), "P" (passthrough), "E" (ender)
    Dim starterLocations As Object: Set starterLocations = CreateObject("Scripting.Dictionary")

    ' Collections to hold categorized frames with their locations
    Dim starters    As Collection: Set starters = New Collection
    Dim passThrough As Collection: Set passThrough = New Collection
    Dim enders      As Collection: Set enders = New Collection

    Dim i           As Long
    For i = 0 To numFrames - 1
        Dim frameName As String
        frameName = frameNames(i)

        ' Get frame start/end point names
        Dim pt1 As String, pt2 As String
        ret = SapModel.frameObj.GetPoints(frameName, pt1, pt2)
        If ret <> 0 Then GoTo NextFrame

        ' Get coordinates
        Dim x1 As Double, y1 As Double, z1 As Double
        Dim x2 As Double, y2 As Double, z2 As Double
        ret = SapModel.pointObj.GetCoordCartesian(pt1, x1, y1, z1)
        If ret <> 0 Then GoTo NextFrame
        ret = SapModel.pointObj.GetCoordCartesian(pt2, x2, y2, z2)
        If ret <> 0 Then GoTo NextFrame

        ' Ensure vertical (column) by checking X/Y nearly equal
        If (Abs(x1 - x2) > COORD_TOLERANCE) Or (Abs(y1 - y2) > COORD_TOLERANCE) Then GoTo NextFrame

        ' Z extents
        Dim minZ As Double, maxZ As Double
        If z1 <= z2 Then
            minZ = z1: maxZ = z2
        Else
            minZ = z2: maxZ = z1
        End If

        ' Check if frame relates to this elevation (within tolerance)
        Dim includesLevel As Boolean
        includesLevel = (elevation >= minZ - COORD_TOLERANCE) And (elevation <= maxZ + COORD_TOLERANCE)
        If Not includesLevel Then GoTo NextFrame

        ' Calculate location key once
        Dim centerX As Double, centerY As Double
        centerX = (x1 + x2) / 2
        centerY = (y1 + y2) / 2
        Dim locKey  As String
        locKey = Format(centerX, "0.000") & "|" & Format(centerY, "0.000")

        ' Classify and store:
        ' Starter: frame starts at level (minZ ~= elevation)
        If Abs(minZ - elevation) <= COORD_TOLERANCE Then
            starters.Add frameName
            starterLocations(locKey) = True  ' Mark this location has a starter
            GoTo NextFrame
        End If

        ' Pass-through: strictly passes through (minZ < elevation < maxZ)
        If (elevation > minZ + COORD_TOLERANCE) And (elevation < maxZ - COORD_TOLERANCE) Then
            passThrough.Add frameName
            GoTo NextFrame
        End If

        ' Ender: frame ends at level (maxZ ~= elevation)
        If Abs(maxZ - elevation) <= COORD_TOLERANCE Then
            enders.Add frameName & "|" & locKey  ' Store with location for later filtering
            GoTo NextFrame
        End If

NextFrame:
    Next i

    ' Combine results: starters -> passThrough -> filtered enders
    Dim unique      As Object
    Set unique = CreateObject("Scripting.Dictionary")

    ' Add starters
    Dim fn          As Variant
    For Each fn In starters
        If Not unique.exists(CStr(fn)) Then
            FindColumnFramesAtElevation.Add fn
            unique.Add CStr(fn), True
        End If
    Next fn

    ' Add passthrough
    For Each fn In passThrough
        If Not unique.exists(CStr(fn)) Then
            FindColumnFramesAtElevation.Add fn
            unique.Add CStr(fn), True
        End If
    Next fn

    ' Add enders (only if no starter at same location)
    For Each fn In enders
        Dim parts() As String
        parts = Split(CStr(fn), "|", 2)  ' Split frameName|locKey
        Dim enderName As String
        enderName = parts(0)

        If UBound(parts) >= 1 Then
            Dim enderLoc As String
            enderLoc = parts(1)

            ' Check if this location has a starter (section change)
            If starterLocations.exists(enderLoc) Then
                LogMsg "Detected section change at elevation " & Format(elevation, "0.0") & _
                        " mm: excluding ender frame " & enderName & " (starter exists at same location)"
            Else
                ' No starter at this location -> include ender
                If Not unique.exists(enderName) Then
                    FindColumnFramesAtElevation.Add enderName
                    unique.Add enderName, True
                End If
            End If
        End If
    Next fn

    On Error GoTo 0
End Function

' Process single column frame and create area
' Optional useBelow flag: when True, area label will include "_F" suffix
' Optional createdCenters: dictionary to track created centers and avoid duplicates
Private Function ProcessColumnFrame(frameName As String, elevation As Double, _
                                    Optional useBelow As Boolean = False, _
                                    Optional createdCenters As Object = Nothing) As Boolean
    On Error GoTo ErrorHandler

    ProcessColumnFrame = False

    ' Get frame section
    Dim sectionName As String
    Dim sectionAuto As String
    Dim ret As Long
    ret = SapModel.frameObj.GetSection(frameName, sectionName, sectionAuto)
    If ret <> 0 Then Exit Function

    If Trim$(sectionName) = "" Then
        sectionName = Trim$(sectionAuto)
    End If

    ' Get section dimensions
    Dim sectionT3 As Double, sectionT2 As Double
    Dim shapeType As String
    Dim shapeParams As Variant
    
    If Not GetSectionDimensions(sectionName, sectionT2, sectionT3, shapeType, shapeParams) Then
        LogMsg "WARNING: Unsupported section type for frame " & frameName & ", section " & sectionName
        Exit Function
    End If
    
    LogMsg "Frame " & frameName & " section: " & shapeType & ", t2=" & Format(sectionT2, "0.0") & ", t3=" & Format(sectionT3, "0.0")

    ' =====================================================================
    ' CRITICAL FIX: Determine prefix from settings based on section name and shape type
    ' =====================================================================
    Dim prefix As String
    prefix = GetPrefixForSection(sectionName, shapeType)
    
    LogMsg "ProcessColumnFrame: frame=" & frameName & ", section=" & sectionName & _
           ", shape=" & shapeType & ", prefix=" & prefix

    ' Build area name using determined prefix (NOT hardcoded constant!)
    Dim areaName As String
    areaName = prefix & frameName
    If useBelow Then areaName = areaName & "_F"
    
    LogMsg "  -> Final area name: " & areaName
    ' =====================================================================

    ' Get frame location
    Dim pt1 As String, pt2 As String
    ret = SapModel.frameObj.GetPoints(frameName, pt1, pt2)
    If ret <> 0 Then Exit Function

    ' Get coordinates
    Dim x1 As Double, y1 As Double, z1 As Double
    Dim x2 As Double, y2 As Double, z2 As Double
    ret = SapModel.pointObj.GetCoordCartesian(pt1, x1, y1, z1)
    If ret <> 0 Then Exit Function
    ret = SapModel.pointObj.GetCoordCartesian(pt2, x2, y2, z2)
    If ret <> 0 Then Exit Function

    ' Use column center
    Dim centerX As Double, centerY As Double
    centerX = (x1 + x2) / 2
    centerY = (y1 + y2) / 2

    ' Get insertion point offsets
    Dim offsetT2 As Double, offsetT3 As Double
    Call GetInsertionPointOffsets(frameName, sectionName, sectionT2, sectionT3, offsetT2, offsetT3)
    
    ' Get frame transformation matrix
    Dim transformMatrix() As Double
    ReDim transformMatrix(0 To 8)
    Dim local2AngleRad As Double
    Dim local3AngleRad As Double
    
    ret = SapModel.frameObj.GetTransformationMatrix(frameName, transformMatrix, True)
    
    If ret = 0 Then
        Dim local2X As Double, local2Y As Double
        local2X = transformMatrix(1)
        local2Y = transformMatrix(4)
        
        Dim local3X As Double, local3Y As Double
        local3X = transformMatrix(2)
        local3Y = transformMatrix(5)
        
        local2AngleRad = Atan2(local2Y, local2X)
        local3AngleRad = Atan2(local3Y, local3X)
    Else
        Dim localAxisAngle As Double
        Dim advanced As Boolean
        ret = SapModel.frameObj.GetLocalAxes(frameName, localAxisAngle, advanced)
        If ret = 0 Then
            local2AngleRad = localAxisAngle * PI / 180#
            local3AngleRad = (localAxisAngle + 90#) * PI / 180#
        Else
            Exit Function
        End If
    End If

    ' Create area
    Call CreateColumnAreaWithShape(frameName, centerX, centerY, elevation, _
                                   sectionT2, sectionT3, offsetT2, offsetT3, _
                                   local3AngleRad, local2AngleRad, _
                                   shapeType, shapeParams, areaName, createdCenters)

    ' Draw punching perimeter if enabled
    On Error Resume Next
    If GetSettingBool("DrawPunching", False) Then
        Dim storyIdx As Long
        storyIdx = GetStoryIndexByElevation(elevation)
        M12A_DrawPunchingPerimeter frameName, centerX, centerY, elevation, _
            sectionT2, sectionT3, offsetT2, offsetT3, local3AngleRad, local2AngleRad, _
            shapeType, storyIdx, sectionName, areaName
    End If
    On Error GoTo ErrorHandler

    ProcessColumnFrame = True
    Exit Function

ErrorHandler:
    LogMsg "ERROR in ProcessColumnFrame(" & frameName & "): " & err.description
    ProcessColumnFrame = False
End Function
Private Function GetPrefixForSection(sectionName As String, shapeType As String) As String
    On Error Resume Next
    GetPrefixForSection = "ColSec_"  ' Default fallback

    If g_Settings Is Nothing Then Exit Function
    If Not g_Settings.exists("SectionPrefixes") Then Exit Function

    Dim prefixDict  As Object
    Set prefixDict = g_Settings("SectionPrefixes")
    If prefixDict Is Nothing Then Exit Function

    ' Check section name for concrete/RC keywords
    Dim upsec       As String
    upsec = UCase$(Trim$(sectionName))

    If InStr(upsec, "CONC") > 0 Or InStr(upsec, "CONCRETE") > 0 Or InStr(upsec, "RC") > 0 Or InStr(upsec, "RECT") > 0 Then
        ' Concrete section
        If prefixDict.exists("RC") Then
            GetPrefixForSection = CStr(prefixDict("RC"))
        End If
    Else
        ' Steel section - use shape type
        Dim shapeKey As String
        shapeKey = UCase$(Trim$(shapeType))

        If prefixDict.exists(shapeKey) Then
            GetPrefixForSection = CStr(prefixDict(shapeKey))
        ElseIf prefixDict.exists("DEFAULT") Then
            GetPrefixForSection = CStr(prefixDict("DEFAULT"))
        End If
    End If
End Function
Private Function GetStoryIndexByElevation(elevation As Double) As Long
    On Error Resume Next
    GetStoryIndexByElevation = 0
    If g_StoryElevations Is Nothing Then Exit Function

    Dim elevationsArr() As Double
    elevationsArr = GetElevationsOrdered(g_StoryElevations)    ' 1-based ascending
    If UBound(elevationsArr) < 1 Then Exit Function

    Dim i           As Long
    For i = 1 To UBound(elevationsArr)
        If Abs(elevationsArr(i) - elevation) <= COORD_TOLERANCE Then
            GetStoryIndexByElevation = i
            Exit Function
        End If
    Next i
End Function

Private Sub CreateColumnAreaWithShape(frameName As String, centerX As Double, centerY As Double, _
        elevation As Double, sectionT2 As Double, sectionT3 As Double, _
        offsetT2 As Double, offsetT3 As Double, _
        t2AngleRad As Double, t3AngleRad As Double, _
        shapeType As String, shapeParams As Variant, _
        ByVal areaName As String, Optional createdCenters As Object = Nothing)
    On Error GoTo ErrorHandler

    ' Dispatch to appropriate shape drawing function
    Select Case UCase$(shapeType)
        Case "I"
            Call CreateIShapeArea(frameName, centerX, centerY, elevation, sectionT2, sectionT3, _
                    offsetT2, offsetT3, t2AngleRad, t3AngleRad, shapeParams, areaName, createdCenters)
        Case "PIPE"
            Call CreatePipeShapeArea(frameName, centerX, centerY, elevation, sectionT2, sectionT3, _
                    offsetT2, offsetT3, t2AngleRad, t3AngleRad, shapeParams, areaName, createdCenters)
        Case "BOX"
            Call CreateBoxShapeArea(frameName, centerX, centerY, elevation, sectionT2, sectionT3, _
                    offsetT2, offsetT3, t2AngleRad, t3AngleRad, shapeParams, areaName, createdCenters)
        Case "CHANNEL"
            Call CreateChannelShapeArea(frameName, centerX, centerY, elevation, sectionT2, sectionT3, _
                    offsetT2, offsetT3, t2AngleRad, t3AngleRad, shapeParams, areaName, createdCenters)
        Case Else
            ' Default: rectangle for RECT, CIRCLE, TEE, ANGLE, and unknown shapes
            Call CreateRectangleShapeArea(frameName, centerX, centerY, elevation, sectionT2, sectionT3, _
                    offsetT2, offsetT3, t2AngleRad, t3AngleRad, areaName, createdCenters)
    End Select
    Exit Sub

ErrorHandler:
    LogMsg "ERROR in CreateColumnAreaWithShape: " & err.description
End Sub

' Draw Rectangle (default for most shapes)
Private Sub CreateRectangleShapeArea(frameName As String, centerX As Double, centerY As Double, _
        elevation As Double, sectionT2 As Double, sectionT3 As Double, _
        offsetT2 As Double, offsetT3 As Double, _
        t2AngleRad As Double, t3AngleRad As Double, _
        ByVal areaName As String, Optional createdCenters As Object = Nothing)
    On Error GoTo ErrorHandler

    ' Calculate sin and cos for the direction angles
    ' Note: t2AngleRad is the angle for section t2 dimension, t3AngleRad for t3 dimension
    Dim cos2 As Double, sin2 As Double
    Dim cos3 As Double, sin3 As Double
    cos2 = Cos(t2AngleRad)
    sin2 = Sin(t2AngleRad)
    cos3 = Cos(t3AngleRad)
    sin3 = Sin(t3AngleRad)

    ' Transform offset from section coordinates to global coordinates
    Dim rotatedOffsetX As Double, rotatedOffsetY As Double
    rotatedOffsetX = offsetT2 * cos2 + offsetT3 * cos3
    rotatedOffsetY = offsetT2 * sin2 + offsetT3 * sin3

    ' Apply rotated offset to center
    Dim actualX As Double, actualY As Double
    actualX = centerX + rotatedOffsetX
    actualY = centerY + rotatedOffsetY

    ' If createdCenters provided, check duplicate at same (x,y,elev) to avoid drawing twice for section change
    Dim centerKey   As String
    centerKey = Format(actualX, "0.000") & "|" & Format(actualY, "0.000") & "|" & Format(elevation, "0.000")
    If Not createdCenters Is Nothing Then
        On Error Resume Next
        If createdCenters.exists(centerKey) Then
            LogMsg "Skipped duplicate area at " & centerKey & " for frame " & frameName & " (likely section-change continuation)"
            Exit Sub
        End If
        On Error GoTo ErrorHandler
    End If

    ' If area with same name already exists, skip creation (non-destructive)
    Dim propName    As String
    Dim ret         As Long
    On Error Resume Next
    ret = SapModel.AreaObj.GetProperty(areaName, propName)
    If ret = 0 Then
        ' Area exists -> skip creating new one to avoid overwriting previous (preserve top-first)
        LogMsg "Area " & areaName & " already exists; skip creation."
        ' Mark createdCenters as well to avoid duplicates from other frames
        If Not createdCenters Is Nothing Then
            On Error Resume Next
            If Not createdCenters.exists(centerKey) Then createdCenters.Add centerKey, areaName
            On Error GoTo ErrorHandler
        End If
        Exit Sub
    End If
    On Error GoTo ErrorHandler

    ' Define rectangle corners in SECTION coordinate system (centered at origin)
    Dim sectionX1 As Double, sectionY1 As Double
    Dim sectionX2 As Double, sectionY2 As Double
    Dim sectionX3 As Double, sectionY3 As Double
    Dim sectionX4 As Double, sectionY4 As Double

    ' Rectangle corners: half extents along t2 and t3 dimensions
    sectionX1 = -sectionT2 / 2: sectionY1 = -sectionT3 / 2
    sectionX2 = sectionT2 / 2: sectionY2 = -sectionT3 / 2
    sectionX3 = sectionT2 / 2: sectionY3 = sectionT3 / 2
    sectionX4 = -sectionT2 / 2: sectionY4 = sectionT3 / 2

    ' Transform each corner from section coordinates to global XY coordinates
    ' Note: t2AngleRad and t3AngleRad already account for the mapping (potentially swapped)
    Dim x1 As Double, y1 As Double
    Dim x2 As Double, y2 As Double
    Dim x3 As Double, y3 As Double
    Dim x4 As Double, y4 As Double

    x1 = actualX + sectionX1 * cos2 + sectionY1 * cos3
    y1 = actualY + sectionX1 * sin2 + sectionY1 * sin3

    x2 = actualX + sectionX2 * cos2 + sectionY2 * cos3
    y2 = actualY + sectionX2 * sin2 + sectionY2 * sin3

    x3 = actualX + sectionX3 * cos2 + sectionY3 * cos3
    y3 = actualY + sectionX3 * sin2 + sectionY3 * sin3

    x4 = actualX + sectionX4 * cos2 + sectionY4 * cos3
    y4 = actualY + sectionX4 * sin2 + sectionY4 * sin3

    ' Create or get corner point names using areaName to ensure uniqueness per naming rule
    Dim pt1Name As String, pt2Name As String, pt3Name As String, pt4Name As String
    pt1Name = areaName & "_P1"
    pt2Name = areaName & "_P2"
    pt3Name = areaName & "_P3"
    pt4Name = areaName & "_P4"

    ' Delete existing points if any (only those exact names) to avoid duplicates
    On Error Resume Next
    SapModel.pointObj.DeleteSpecialPoint pt1Name
    SapModel.pointObj.DeleteSpecialPoint pt2Name
    SapModel.pointObj.DeleteSpecialPoint pt3Name
    SapModel.pointObj.DeleteSpecialPoint pt4Name
    On Error GoTo ErrorHandler

    ' Create corner points at requested elevation (using rotated global coordinates)
    ret = SapModel.pointObj.AddCartesian(x1, y1, elevation, pt1Name, pt1Name)
    ret = SapModel.pointObj.AddCartesian(x2, y2, elevation, pt2Name, pt2Name)
    ret = SapModel.pointObj.AddCartesian(x3, y3, elevation, pt3Name, pt3Name)
    ret = SapModel.pointObj.AddCartesian(x4, y4, elevation, pt4Name, pt4Name)

    ' Create area from 4 points (order consistent - counter-clockwise)
    Dim pointNames() As String
    ReDim pointNames(0 To 3)
    pointNames(0) = pt1Name
    pointNames(1) = pt2Name
    pointNames(2) = pt3Name
    pointNames(3) = pt4Name

    ' Create area: use AREA_SECTION_NONE as property
    ret = SapModel.AreaObj.AddByPoint(4, pointNames, areaName, AREA_SECTION_NONE, areaName)
    If ret <> 0 Then
        LogMsg "ERROR: Failed to create area " & areaName
        Exit Sub
    End If

    ' Set area property to "None" explicitly
    ret = SapModel.AreaObj.SetProperty(areaName, AREA_SECTION_NONE)

    ' Optional: Assign to group for identification
    On Error Resume Next
    ret = SapModel.GroupDef.SetGroup("ColumnCrossSections")
    ret = SapModel.AreaObj.SetGroupAssign(areaName, "ColumnCrossSections", False)
    On Error GoTo 0

    ' Record created center to avoid duplicates (for section-change cases)
    If Not createdCenters Is Nothing Then
        On Error Resume Next
        If Not createdCenters.exists(centerKey) Then createdCenters.Add centerKey, areaName
        On Error GoTo 0
    End If

    ' Log with rotation info
    Dim angle2Deg As Double, angle3Deg As Double
    angle2Deg = t2AngleRad * 180 / PI
    angle3Deg = t3AngleRad * 180 / PI
    If Abs(angle2Deg) > 0.01 Or Abs(angle3Deg - 90) > 0.01 Then
        LogMsg "Created area " & areaName & " for frame " & frameName & _
                " (t2=" & Format(sectionT2, "0.0") & "mm @ " & Format(angle2Deg, "0.0") & _
                "°, t3=" & Format(sectionT3, "0.0") & "mm @ " & Format(angle3Deg, "0.0") & "°)"
    Else
        LogMsg "Created area " & areaName & " for frame " & frameName & _
                " (t2=" & Format(sectionT2, "0.0") & "mm, t3=" & Format(sectionT3, "0.0") & "mm)"
    End If
    Exit Sub

ErrorHandler:
    LogMsg "ERROR in CreateColumnAreaRectangle: " & err.description
End Sub

' Helper function: Atan2 (VBA doesn't have it built-in)
Private Function Atan2(Y As Double, X As Double) As Double
    If X > 0 Then
        Atan2 = Atn(Y / X)
    ElseIf X < 0 Then
        If Y >= 0 Then
            Atan2 = Atn(Y / X) + PI
        Else
            Atan2 = Atn(Y / X) - PI
        End If
    Else    ' x = 0
        If Y > 0 Then
            Atan2 = PI / 2
        ElseIf Y < 0 Then
            Atan2 = -PI / 2
        Else
            Atan2 = 0    ' undefined, return 0
        End If
    End If
End Function

' Helper: Check and mark duplicate area location
Private Function CheckAndMarkDuplicate(actualX As Double, actualY As Double, elevation As Double, _
        frameName As String, areaName As String, _
        createdCenters As Object) As Boolean
    CheckAndMarkDuplicate = False
    Dim centerKey   As String
    centerKey = Format(actualX, "0.000") & "|" & Format(actualY, "0.000") & "|" & Format(elevation, "0.000")

    If Not createdCenters Is Nothing Then
        On Error Resume Next
        If createdCenters.exists(centerKey) Then
            LogMsg "Skipped duplicate area at " & centerKey & " for frame " & frameName
            CheckAndMarkDuplicate = True
            Exit Function
        End If
        On Error GoTo 0
    End If

    ' Check if area already exists
    Dim propName    As String
    Dim ret         As Long
    On Error Resume Next
    ret = SapModel.AreaObj.GetProperty(areaName, propName)
    If ret = 0 Then
        LogMsg "Area " & areaName & " already exists; skip creation."
        If Not createdCenters Is Nothing Then
            On Error Resume Next
            If Not createdCenters.exists(centerKey) Then createdCenters.Add centerKey, areaName
        End If
        CheckAndMarkDuplicate = True
        Exit Function
    End If
    On Error GoTo 0
End Function

' Helper: Create area from array of local coordinates
' pointsLocal: array of (x,y) pairs in section local coordinates (0-based, even indices = x, odd = y)
' numPoints: number of points
Private Sub CreateAreaFromLocalPoints(pointsLocal() As Double, numPoints As Long, _
        actualX As Double, actualY As Double, elevation As Double, _
        cos2 As Double, sin2 As Double, cos3 As Double, sin3 As Double, _
        areaName As String, frameName As String, _
        sectionT2 As Double, sectionT3 As Double)
    On Error GoTo ErrorHandler

    ' Transform points to global coordinates
    Dim pointNames() As String
    ReDim pointNames(0 To numPoints - 1)

    Dim i           As Long
    Dim ret         As Long
    For i = 0 To numPoints - 1
        Dim localX As Double, localY As Double
        localX = pointsLocal(i * 2)
        localY = pointsLocal(i * 2 + 1)

        ' Transform to global
        Dim globalX As Double, globalY As Double
        globalX = actualX + localX * cos2 + localY * cos3
        globalY = actualY + localX * sin2 + localY * sin3

        ' Create point
        Dim ptName  As String
        ptName = areaName & "_P" & CStr(i + 1)

        ' Delete if exists
        On Error Resume Next
        SapModel.pointObj.DeleteSpecialPoint ptName
        On Error GoTo ErrorHandler

        ' Create point
        ret = SapModel.pointObj.AddCartesian(globalX, globalY, elevation, ptName, ptName)
        pointNames(i) = ptName
    Next i

    ' Create area from points
    ret = SapModel.AreaObj.AddByPoint(numPoints, pointNames, areaName, AREA_SECTION_NONE, areaName)
    If ret <> 0 Then
        LogMsg "ERROR: Failed to create area " & areaName
        Exit Sub
    End If

    ' Set area property
    ret = SapModel.AreaObj.SetProperty(areaName, AREA_SECTION_NONE)

    ' Assign to group
    On Error Resume Next
    ret = SapModel.GroupDef.SetGroup("ColumnCrossSections")
    ret = SapModel.AreaObj.SetGroupAssign(areaName, "ColumnCrossSections", False)
    On Error GoTo 0

    LogMsg "Created area " & areaName & " for frame " & frameName & _
            " (t2=" & Format(sectionT2, "0.0") & "mm, t3=" & Format(sectionT3, "0.0") & "mm)"
    Exit Sub

ErrorHandler:
    LogMsg "ERROR in CreateAreaFromLocalPoints: " & err.description
End Sub

' Helper: Calculate scaled thickness for visibility (non-scale display)
Private Function GetScaledThickness(actualThickness As Double, maxDim As Double) As Double
    ' Scale thickness to be visible but not too large
    Dim minThickness As Double
    minThickness = maxDim * THICKNESS_SCALE
    If actualThickness < minThickness Then
        GetScaledThickness = minThickness
    Else
        GetScaledThickness = actualThickness
    End If
End Function

'===============================================================
' SHAPE-SPECIFIC DRAWING FUNCTIONS
'===============================================================

' Draw I-Section (H-beam)
' shapeParams: (0)=t3, (1)=t2, (2)=tf, (3)=tw, (4)=t2b, (5)=tfb
Private Sub CreateIShapeArea(frameName As String, centerX As Double, centerY As Double, _
        elevation As Double, sectionT2 As Double, sectionT3 As Double, _
        offsetT2 As Double, offsetT3 As Double, _
        t2AngleRad As Double, t3AngleRad As Double, _
        shapeParams As Variant, areaName As String, _
        Optional createdCenters As Object = Nothing)
    On Error GoTo ErrorHandler

    Dim cos2 As Double, sin2 As Double, cos3 As Double, sin3 As Double
    cos2 = Cos(t2AngleRad): sin2 = Sin(t2AngleRad)
    cos3 = Cos(t3AngleRad): sin3 = Sin(t3AngleRad)

    Dim rotatedOffsetX As Double, rotatedOffsetY As Double
    rotatedOffsetX = offsetT2 * cos2 + offsetT3 * cos3
    rotatedOffsetY = offsetT2 * sin2 + offsetT3 * sin3

    Dim actualX As Double, actualY As Double
    actualX = centerX + rotatedOffsetX
    actualY = centerY + rotatedOffsetY

    If CheckAndMarkDuplicate(actualX, actualY, elevation, frameName, areaName, createdCenters) Then Exit Sub

    ' Extract parameters
    Dim h As Double, bf As Double, tf As Double, tw As Double, bfb As Double, tfb As Double
    h = shapeParams(0)
    bf = shapeParams(1)
    tf = shapeParams(2)
    tw = shapeParams(3)
    bfb = shapeParams(4)
    tfb = shapeParams(5)

    ' Scale thicknesses for visibility
    Dim maxDim      As Double
    maxDim = IIf(h > bf, h, bf)
    If maxDim < bfb Then maxDim = bfb
    tf = GetScaledThickness(tf, maxDim)
    tw = GetScaledThickness(tw, maxDim)
    tfb = GetScaledThickness(tfb, maxDim)

    ' Define I-shape outline (12 points, counter-clockwise from bottom-left)
    Dim pointsLocal() As Double
    ReDim pointsLocal(0 To 23)  ' 12 points x 2 coords

    Dim h2 As Double, bf2 As Double, bfb2 As Double, tw2 As Double
    h2 = h / 2: bf2 = bf / 2: bfb2 = bfb / 2: tw2 = tw / 2

    ' Bottom flange (4 points)
    pointsLocal(0) = -bfb2: pointsLocal(1) = -h2                    ' P1: left-bottom
    pointsLocal(2) = bfb2: pointsLocal(3) = -h2                     ' P2: right-bottom
    pointsLocal(4) = bfb2: pointsLocal(5) = -h2 + tfb               ' P3: right-bottom inner
    pointsLocal(6) = tw2: pointsLocal(7) = -h2 + tfb                ' P4: web right-bottom

    ' Web to top flange (4 points)
    pointsLocal(8) = tw2: pointsLocal(9) = h2 - tf                  ' P5: web right-top
    pointsLocal(10) = bf2: pointsLocal(11) = h2 - tf                ' P6: top flange inner-right
    pointsLocal(12) = bf2: pointsLocal(13) = h2                     ' P7: top flange outer-right
    pointsLocal(14) = -bf2: pointsLocal(15) = h2                    ' P8: top flange outer-left

    ' Top flange to web (4 points)
    pointsLocal(16) = -bf2: pointsLocal(17) = h2 - tf               ' P9: top flange inner-left
    pointsLocal(18) = -tw2: pointsLocal(19) = h2 - tf               ' P10: web left-top
    pointsLocal(20) = -tw2: pointsLocal(21) = -h2 + tfb             ' P11: web left-bottom
    pointsLocal(22) = -bfb2: pointsLocal(23) = -h2 + tfb            ' P12: left-bottom inner

    Call CreateAreaFromLocalPoints(pointsLocal, 12, actualX, actualY, elevation, _
            cos2, sin2, cos3, sin3, areaName, frameName, sectionT2, sectionT3)

    ' Mark as created
    If Not createdCenters Is Nothing Then
        Dim centerKey As String
        centerKey = Format(actualX, "0.000") & "|" & Format(actualY, "0.000") & "|" & Format(elevation, "0.000")
        On Error Resume Next
        If Not createdCenters.exists(centerKey) Then createdCenters.Add centerKey, areaName
    End If
    Exit Sub

ErrorHandler:
    LogMsg "ERROR in CreateIShapeArea: " & err.description
End Sub

' Draw Pipe (circular hollow section)
' shapeParams: (0)=outerDiam, (1)=thickness
Private Sub CreatePipeShapeArea(frameName As String, centerX As Double, centerY As Double, _
        elevation As Double, sectionT2 As Double, sectionT3 As Double, _
        offsetT2 As Double, offsetT3 As Double, _
        t2AngleRad As Double, t3AngleRad As Double, _
        shapeParams As Variant, areaName As String, _
        Optional createdCenters As Object = Nothing)
    On Error GoTo ErrorHandler

    Dim cos2 As Double, sin2 As Double, cos3 As Double, sin3 As Double
    cos2 = Cos(t2AngleRad): sin2 = Sin(t2AngleRad)
    cos3 = Cos(t3AngleRad): sin3 = Sin(t3AngleRad)

    Dim rotatedOffsetX As Double, rotatedOffsetY As Double
    rotatedOffsetX = offsetT2 * cos2 + offsetT3 * cos3
    rotatedOffsetY = offsetT2 * sin2 + offsetT3 * sin3

    Dim actualX As Double, actualY As Double
    actualX = centerX + rotatedOffsetX
    actualY = centerY + rotatedOffsetY

    If CheckAndMarkDuplicate(actualX, actualY, elevation, frameName, areaName, createdCenters) Then Exit Sub

    ' Extract parameters
    Dim outerD As Double, thick As Double
    outerD = shapeParams(0)
    thick = shapeParams(1)

    ' Scale thickness for visibility
    thick = GetScaledThickness(thick, outerD)

    Dim innerD      As Double
    innerD = outerD - 2 * thick
    If innerD < 0 Then innerD = outerD * 0.5  ' Safety check

    ' Create octagonal approximation (16 points: 8 outer + 8 inner)
    Dim numSides    As Long
    numSides = 8
    Dim pointsLocal() As Double
    ReDim pointsLocal(0 To 31)  ' 16 points x 2 coords

    Dim i           As Long
    Dim angle       As Double
    Dim outerR As Double, innerR As Double
    outerR = outerD / 2: innerR = innerD / 2

    ' Outer octagon (counter-clockwise)
    For i = 0 To numSides - 1
        angle = 2 * PI * i / numSides
        pointsLocal(i * 2) = outerR * Cos(angle)
        pointsLocal(i * 2 + 1) = outerR * Sin(angle)
    Next i

    ' Inner octagon (clockwise to create hole)
    For i = 0 To numSides - 1
        angle = 2 * PI * (numSides - 1 - i) / numSides
        pointsLocal((numSides + i) * 2) = innerR * Cos(angle)
        pointsLocal((numSides + i) * 2 + 1) = innerR * Sin(angle)
    Next i

    Call CreateAreaFromLocalPoints(pointsLocal, 16, actualX, actualY, elevation, _
            cos2, sin2, cos3, sin3, areaName, frameName, sectionT2, sectionT3)

    ' Mark as created
    If Not createdCenters Is Nothing Then
        Dim centerKey As String
        centerKey = Format(actualX, "0.000") & "|" & Format(actualY, "0.000") & "|" & Format(elevation, "0.000")
        On Error Resume Next
        If Not createdCenters.exists(centerKey) Then createdCenters.Add centerKey, areaName
    End If
    Exit Sub

ErrorHandler:
    LogMsg "ERROR in CreatePipeShapeArea: " & err.description
End Sub

' Draw Box (rectangular hollow section)
' shapeParams: (0)=t3, (1)=t2, (2)=tf (thick-3), (3)=tw (thick-2)
Private Sub CreateBoxShapeArea(frameName As String, centerX As Double, centerY As Double, _
        elevation As Double, sectionT2 As Double, sectionT3 As Double, _
        offsetT2 As Double, offsetT3 As Double, _
        t2AngleRad As Double, t3AngleRad As Double, _
        shapeParams As Variant, areaName As String, _
        Optional createdCenters As Object = Nothing)
    On Error GoTo ErrorHandler

    Dim cos2 As Double, sin2 As Double, cos3 As Double, sin3 As Double
    cos2 = Cos(t2AngleRad): sin2 = Sin(t2AngleRad)
    cos3 = Cos(t3AngleRad): sin3 = Sin(t3AngleRad)

    Dim rotatedOffsetX As Double, rotatedOffsetY As Double
    rotatedOffsetX = offsetT2 * cos2 + offsetT3 * cos3
    rotatedOffsetY = offsetT2 * sin2 + offsetT3 * sin3

    Dim actualX As Double, actualY As Double
    actualX = centerX + rotatedOffsetX
    actualY = centerY + rotatedOffsetY

    If CheckAndMarkDuplicate(actualX, actualY, elevation, frameName, areaName, createdCenters) Then Exit Sub

    ' Extract parameters
    Dim h As Double, b As Double, t3 As Double, t2 As Double
    h = shapeParams(0)
    b = shapeParams(1)
    t3 = shapeParams(2)  ' thickness in 3-direction (vertical walls)
    t2 = shapeParams(3)  ' thickness in 2-direction (horizontal walls)

    ' Scale thicknesses for visibility
    Dim maxDim      As Double
    maxDim = IIf(h > b, h, b)
    t3 = GetScaledThickness(t3, maxDim)
    t2 = GetScaledThickness(t2, maxDim)

    ' Define box outline (8 points: 4 outer + 4 inner)
    Dim pointsLocal() As Double
    ReDim pointsLocal(0 To 15)  ' 8 points x 2 coords

    Dim h2 As Double, B2 As Double
    h2 = h / 2: B2 = b / 2

    ' Outer rectangle (counter-clockwise)
    pointsLocal(0) = -B2: pointsLocal(1) = -h2      ' P1: bottom-left
    pointsLocal(2) = B2: pointsLocal(3) = -h2       ' P2: bottom-right
    pointsLocal(4) = B2: pointsLocal(5) = h2        ' P3: top-right
    pointsLocal(6) = -B2: pointsLocal(7) = h2       ' P4: top-left

    ' Inner rectangle (clockwise to create hole)
    pointsLocal(8) = -B2 + t2: pointsLocal(9) = h2 - t3      ' P5: top-left inner
    pointsLocal(10) = B2 - t2: pointsLocal(11) = h2 - t3     ' P6: top-right inner
    pointsLocal(12) = B2 - t2: pointsLocal(13) = -h2 + t3    ' P7: bottom-right inner
    pointsLocal(14) = -B2 + t2: pointsLocal(15) = -h2 + t3   ' P8: bottom-left inner

    Call CreateAreaFromLocalPoints(pointsLocal, 8, actualX, actualY, elevation, _
            cos2, sin2, cos3, sin3, areaName, frameName, sectionT2, sectionT3)

    ' Mark as created
    If Not createdCenters Is Nothing Then
        Dim centerKey As String
        centerKey = Format(actualX, "0.000") & "|" & Format(actualY, "0.000") & "|" & Format(elevation, "0.000")
        On Error Resume Next
        If Not createdCenters.exists(centerKey) Then createdCenters.Add centerKey, areaName
    End If
    Exit Sub

ErrorHandler:
    LogMsg "ERROR in CreateBoxShapeArea: " & err.description
End Sub

' Draw Channel (C-section)
' shapeParams: (0)=t3, (1)=t2, (2)=tf, (3)=tw
Private Sub CreateChannelShapeArea(frameName As String, centerX As Double, centerY As Double, _
        elevation As Double, sectionT2 As Double, sectionT3 As Double, _
        offsetT2 As Double, offsetT3 As Double, _
        t2AngleRad As Double, t3AngleRad As Double, _
        shapeParams As Variant, areaName As String, _
        Optional createdCenters As Object = Nothing)
    On Error GoTo ErrorHandler

    Dim cos2 As Double, sin2 As Double, cos3 As Double, sin3 As Double
    cos2 = Cos(t2AngleRad): sin2 = Sin(t2AngleRad)
    cos3 = Cos(t3AngleRad): sin3 = Sin(t3AngleRad)

    Dim rotatedOffsetX As Double, rotatedOffsetY As Double
    rotatedOffsetX = offsetT2 * cos2 + offsetT3 * cos3
    rotatedOffsetY = offsetT2 * sin2 + offsetT3 * sin3

    Dim actualX As Double, actualY As Double
    actualX = centerX + rotatedOffsetX
    actualY = centerY + rotatedOffsetY

    If CheckAndMarkDuplicate(actualX, actualY, elevation, frameName, areaName, createdCenters) Then Exit Sub

    ' Extract parameters
    Dim h As Double, b As Double, tf As Double, tw As Double
    h = shapeParams(0)
    b = shapeParams(1)
    tf = shapeParams(2)
    tw = shapeParams(3)

    ' Scale thicknesses for visibility
    Dim maxDim      As Double
    maxDim = IIf(h > b, h, b)
    tf = GetScaledThickness(tf, maxDim)
    tw = GetScaledThickness(tw, maxDim)

    ' Define C-shape outline (8 points, counter-clockwise from bottom-left)
    Dim pointsLocal() As Double
    ReDim pointsLocal(0 To 15)  ' 8 points x 2 coords

    Dim h2          As Double
    h2 = h / 2

    ' Start from left edge (web)
    pointsLocal(0) = -b: pointsLocal(1) = -h2                    ' P1: bottom-left outer
    pointsLocal(2) = 0: pointsLocal(3) = -h2                     ' P2: bottom-right outer
    pointsLocal(4) = 0: pointsLocal(5) = -h2 + tf                ' P3: bottom flange inner
    pointsLocal(6) = -b + tw: pointsLocal(7) = -h2 + tf          ' P4: web inner-bottom
    pointsLocal(8) = -b + tw: pointsLocal(9) = h2 - tf           ' P5: web inner-top
    pointsLocal(10) = 0: pointsLocal(11) = h2 - tf               ' P6: top flange inner
    pointsLocal(12) = 0: pointsLocal(13) = h2                    ' P7: top-right outer
    pointsLocal(14) = -b: pointsLocal(15) = h2                   ' P8: top-left outer

    Call CreateAreaFromLocalPoints(pointsLocal, 8, actualX, actualY, elevation, _
            cos2, sin2, cos3, sin3, areaName, frameName, sectionT2, sectionT3)

    ' Mark as created
    If Not createdCenters Is Nothing Then
        Dim centerKey As String
        centerKey = Format(actualX, "0.000") & "|" & Format(actualY, "0.000") & "|" & Format(elevation, "0.000")
        On Error Resume Next
        If Not createdCenters.exists(centerKey) Then createdCenters.Add centerKey, areaName
    End If
    Exit Sub

ErrorHandler:
    LogMsg "ERROR in CreateChannelShapeArea: " & err.description
End Sub

' Get insertion point offsets in local coordinates (before rotation)
Private Sub GetInsertionPointOffsets(frameName As String, sectionName As String, _
        Width As Double, Depth As Double, _
        ByRef offsetX As Double, ByRef offsetY As Double)
    On Error Resume Next

    offsetX = 0
    offsetY = 0

    Dim cardinalPoint As Long
    Dim mirror2     As Boolean
    Dim mirror3     As Boolean
    Dim stiffTransform As Boolean
    Dim ret         As Long
    Dim CSys        As String

    ' Arrays for offsets (ByRef arrays required by API)
    Dim Offset1()   As Double
    Dim Offset2()   As Double

    ' Initialize arrays with 3 elements as examples (API expects 0..2)
    ReDim Offset1(0 To 2)
    ReDim Offset2(0 To 2)

    ' Try the newer API GetInsertionPoint_1 (has Mirror3)
    ret = -1
    err.Clear
    On Error Resume Next
    ret = SapModel.frameObj.GetInsertionPoint_1(frameName, cardinalPoint, mirror2, mirror3, stiffTransform, Offset1, Offset2, CSys)
    If err.number <> 0 Or ret <> 0 Then
        ' Reset error and try older GetInsertionPoint (without Mirror3)
        err.Clear
        mirror3 = False
        ReDim Offset1(0 To 2)
        ReDim Offset2(0 To 2)
        ret = SapModel.frameObj.GetInsertionPoint(frameName, cardinalPoint, mirror2, stiffTransform, Offset1, Offset2, CSys)
        If err.number <> 0 Or ret <> 0 Then
            ' Both attempts failed -> default to centroid
            err.Clear
            cardinalPoint = 10
            mirror2 = False
            mirror3 = False
            stiffTransform = False
        End If
    End If
    On Error GoTo 0

    ' Map cardinal point to offsets in LOCAL coordinates
    ' Local-2 axis -> width, Local-3 axis -> depth
    Select Case cardinalPoint
        Case 1  ' bottom-left
            offsetX = -Width / 2
            offsetY = -Depth / 2
        Case 2  ' bottom-center
            offsetX = 0
            offsetY = -Depth / 2
        Case 3  ' bottom-right
            offsetX = Width / 2
            offsetY = -Depth / 2
        Case 4  ' middle-left
            offsetX = -Width / 2
            offsetY = 0
        Case 5  ' middle-center
            offsetX = 0
            offsetY = 0
        Case 6  ' middle-right
            offsetX = Width / 2
            offsetY = 0
        Case 7  ' top-left
            offsetX = -Width / 2
            offsetY = Depth / 2
        Case 8  ' top-center
            offsetX = 0
            offsetY = Depth / 2
        Case 9  ' top-right
            offsetX = Width / 2
            offsetY = Depth / 2
        Case Else  ' 10 (centroid) or 11 (shear center)
            offsetX = 0
            offsetY = 0
    End Select

    ' Apply mirror flags in LOCAL coordinate system
    ' Mirror2 flips about local-2 axis -> invert local-3 (depth/offsetY)
    ' Mirror3 flips about local-3 axis -> invert local-2 (width/offsetX)
    If mirror2 Then
        offsetY = -offsetY
    End If
    If mirror3 Then
        offsetX = -offsetX
    End If
End Sub

'===============================================================
' SELECTION / RESOLVE FUNCTIONS
'===============================================================

Private Function ResolveSelectionToElevations(ByVal userInput As String, gridData As Object) As Collection
    On Error Resume Next
    Dim result      As New Collection
    If gridData Is Nothing Or gridData.count = 0 Then
        Set ResolveSelectionToElevations = result
        Exit Function
    End If

    ' Build ordered elevations array
    Dim ordered()   As Double
    ordered = GetElevationsOrdered(gridData)    ' 1-based array

    ' Normalize input, split by comma
    Dim tokens()    As String
    tokens = Split(Replace(userInput, " ", ""), ",")

    Dim t           As Variant
    For Each t In tokens
        Dim token   As String
        token = Trim(CStr(t))
        If token = "" Then GoTo NextToken

        ' check for range using "to" (case-insensitive) or hyphen '-'
        Dim lowerStr As String, upperStr As String
        Dim posTo   As Long
        posTo = InStr(LCase(token), "to")
        If posTo > 0 Then
            lowerStr = Left(token, posTo - 1)
            upperStr = mid(token, posTo + 2)
            Call ProcessIndexRangeOrSingle(lowerStr, upperStr, ordered, result, gridData)
            GoTo NextToken
        End If

        If InStr(token, "-") > 0 Then
            Dim parts() As String
            parts = Split(token, "-")
            If UBound(parts) = 1 Then
                lowerStr = Trim(parts(0))
                upperStr = Trim(parts(1))
                Call ProcessIndexRangeOrSingle(lowerStr, upperStr, ordered, result, gridData)
                GoTo NextToken
            End If
        End If

        ' If token is numeric:
        If IsNumeric(token) Then
            ' If integer and within index range -> treat as index
            If InStr(token, ".") = 0 Then
                Dim idxVal As Long
                idxVal = CLng(token)
                If idxVal >= 1 And idxVal <= UBound(ordered) Then
                    On Error Resume Next
                    result.Add ordered(idxVal)
                    On Error GoTo 0
                    GoTo NextToken
                End If
            End If
            ' Otherwise treat as elevation value (match within tolerance)
            Dim elevVal As Double
            elevVal = CDbl(token)
            Dim found As Boolean
            found = False
            Dim k   As Variant
            For Each k In gridData.keys
                If Abs(CDbl(k) - elevVal) <= COORD_TOLERANCE Then
                    result.Add CDbl(k)
                    found = True
                    Exit For
                End If
            Next k
            ' if not found, skip silently (could warn)
            GoTo NextToken
        End If

        ' Otherwise treat token as story name (case-insensitive exact match)
        Dim nameFound As Boolean
        nameFound = False
        Dim key     As Variant
        For Each key In gridData.keys
            If UCase(Trim$(CStr(gridData(key)))) = UCase(Trim(token)) Then
                result.Add CDbl(key)
                nameFound = True
                Exit For
            End If
        Next key
        ' end token processing
NextToken:
    Next t

    ' Remove duplicates (keep first occurrence)
    Dim unique      As Object
    Set unique = CreateObject("Scripting.Dictionary")
    Dim outColl     As New Collection
    Dim item        As Variant
    For Each item In result
        Dim sKey    As String
        sKey = Format(CDbl(item), "0.000000")
        If Not unique.exists(sKey) Then
            unique.Add sKey, True
            outColl.Add CDbl(item)
        End If
    Next item

    Set ResolveSelectionToElevations = outColl
    On Error GoTo 0
End Function

' Helper used by ResolveSelectionToElevations: process a range of indices or single index strings
Private Sub ProcessIndexRangeOrSingle(ByVal lowerStr As String, ByVal upperStr As String, _
        ByRef ordered() As Double, ByRef result As Collection, gridData As Object)
    On Error Resume Next
    If lowerStr = "" Or upperStr = "" Then Exit Sub
    If IsNumeric(lowerStr) And IsNumeric(upperStr) Then
        Dim l As Long, u As Long
        l = CLng(lowerStr): u = CLng(upperStr)
        If l > u Then
            Dim tmp As Long: tmp = l: l = u: u = tmp
        End If
        Dim i       As Long
        For i = l To u
            If i >= 1 And i <= UBound(ordered) Then
                result.Add ordered(i)
            End If
        Next i
    Else
        ' if values are non-numeric, do nothing here
    End If
    On Error GoTo 0
End Sub

' Returns a 1-based ascending array of elevations from gridData
Private Function GetElevationsOrdered(gridData As Object) As Double()
    Dim n           As Long
    n = gridData.count
    Dim arr()       As Double
    If n = 0 Then
        ReDim arr(0)
        GetElevationsOrdered = arr
        Exit Function
    End If

    ReDim arr(1 To n)
    Dim idx         As Long
    idx = 1
    Dim k           As Variant
    For Each k In gridData.keys
        arr(idx) = CDbl(k)
        idx = idx + 1
    Next k

    ' sort ascending
    Dim i As Long, j As Long, tmp As Double
    For i = 1 To n - 1
        For j = i + 1 To n
            If arr(i) > arr(j) Then
                tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
            End If
        Next j
    Next i

    GetElevationsOrdered = arr
End Function

'===============================================================
' UTILITY FUNCTIONS
'===============================================================

' Helper: check whether a VBA procedure exists in this project (by name)
Private Function VbProcedureExists(procName As String) As Boolean
    On Error Resume Next
    ' crude check: try to get address of procedure via Application.Run and handle error
    ' We'll simply test for existence by checking for error when getting the vbcomponent collection (best-effort)
    Dim exists      As Boolean
    exists = False
    ' Try to call the procedure name in a protected way - we won't actually execute heavy work.
    ' Instead check type name in VBProject components for Module containing procName - but VBProject may be protected.
    ' We'll fallback to Try/Catch approach: call Application.Run and trap "Sub or Function not defined" if missing.
    On Error Resume Next
    Application.Run procName
    If err.number = 0 Then
        ' If it ran without error it exists (but may have executed). Clear error.
        exists = True
    Else
        ' If error is "Sub or Function not defined" then doesn't exist; other errors ignored.
        exists = (err.description <> "Sub or Function not defined")
    End If
    err.Clear
    On Error GoTo 0
    VbProcedureExists = exists
End Function

' Helper to safely get a string setting
Private Function GetSettingString(key As String, Optional defaultVal As String = "") As String
    On Error Resume Next
    If g_Settings Is Nothing Then
        GetSettingString = defaultVal
        Exit Function
    End If
    If g_Settings.exists(key) Then
        GetSettingString = CStr(g_Settings(key))
    Else
        GetSettingString = defaultVal
    End If
End Function

Private Function GetSettingBool(key As String, Optional defaultVal As Boolean = False) As Boolean
    On Error Resume Next
    If g_Settings Is Nothing Then
        GetSettingBool = defaultVal
        Exit Function
    End If
    If g_Settings.exists(key) Then
        GetSettingBool = CBool(g_Settings(key))
    Else
        GetSettingBool = defaultVal
    End If
End Function

' Simple logging function
Private Sub LogMsg(msg As String)
Debug.Print "modDrawColumnAreas: " & msg
End Sub

'===============================================================
' END OF MODULE
'===============================================================

