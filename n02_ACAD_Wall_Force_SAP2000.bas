Attribute VB_Name = "n02_ACAD_Wall_Force_SAP2000"
Option Explicit
'===============================================================
' Module: n02_ACAD_Wall_Force_SAP2000
' Purpose: Assign wall loads from AutoCAD to SAP2000 model
' Author: DTS System (by thanhtdvncc)
' Version: 2.1 - User Mapping Override Support
'===============================================================

' ==================== CONSTANTS ====================
Private Const XDATA_APP As String = "DTS_APP"
Private Const WALL_LAYER As String = "DTS_WALL_DIAGRAM"
Private Const COORD_TOLERANCE As Double = 250    ' mm - Tube radius
Private Const ANGLE_TOLERANCE As Double = 0.174533  ' 10 degrees in radians
Private Const MIN_OVERLAP_RATIO As Double = 0.15
Private Const PI    As Double = 3.14159265358979
Private Const GRID_CELL_SIZE As Double = 10000   ' 10m spatial hash grid
Private Const APP_NAME As String = "DTS_APP"

' XData structure offsets
Public Const XDATA_OFFSET_THICKNESS As Long = 1
Public Const XDATA_OFFSET_WALLTYPE As Long = 2
Public Const XDATA_OFFSET_LOADPATTERN As Long = 3
Public Const XDATA_OFFSET_LOADVALUE As Long = 4
Public Const XDATA_OFFSET_MAPPING_START As Long = 5
Public Const XDATA_MAPPING_RECORD_SIZE As Long = 4

' === USER OVERRIDE EXTENSION IN XDATA ===
' After base mappings, we append:
'   OV_FLAG        : 1 or 0
'   OV_WallType    : string
'   OV_Pattern     : string
'   OV_Value       : double
' Then OV mappings with SAME record size as base mappings.
'
' Layout:
'   0  : AppName
'   1  : Thickness
'   2  : WallType
'   3  : LoadPattern
'   4  : LoadValue
'   5..N: Base mappings (system)
'   N+1: OV_Flag (Int, 0/1)          (DXF=1071)
'   N+2: OV_WallType (string)        (DXF=1000)
'   N+3: OV_Pattern (string)         (DXF=1000)
'   N+4: OV_Value (double)           (DXF=1040)
'   Then OV mappings (TargetFrame, MapType, DistI, DistJ) like base.

Private Const XD_CODE_APPNAME As Integer = 1001
Private Const XD_CODE_STRING As Integer = 1000
Private Const XD_CODE_REAL As Integer = 1040
Private Const XD_CODE_INT32 As Integer = 1071

' ==================== TYPE DEFINITIONS ====================
Private Type Point2D
    X               As Double
    Y               As Double
End Type

Private Type IntervalRangeVS
    Start           As Double    ' local coordinate along frame (mm)
    End             As Double    ' local coordinate along frame (mm)
End Type

Private Type SapFrameStatus
    FrameIndex      As Long     ' index in mFrameDB
    Length          As Double
    Occupied()      As IntervalRangeVS
    OccupiedCount   As Long
End Type

Private Type CadLoadSegment
    WallIndex       As Long     ' index in wallLines array (Dictionary) OR internal wall index
    StartGlobal     As Double   ' global 1D coordinate on strip (mm)
    EndGlobal       As Double
End Type

Private Type CadLoadStrip
    StripID         As Long
    angle           As Double
    UnitX           As Double
    UnitY           As Double
    OriginX         As Double
    OriginY         As Double
    segments()      As CadLoadSegment
    segmentCount    As Long
    TotalLength     As Double
End Type

Private Type SapStripInfo
    StripID         As Long
    angle           As Double
    UnitX           As Double
    UnitY           As Double
    OriginX         As Double
    OriginY         As Double
    FrameIndices()  As Long
    FrameCount      As Long
    TotalLength     As Double
End Type

Private Type StripMatchPair
    CadStripID      As Long
    SapStripID      As Long
    Score           As Double
    overlapRatio    As Double
    DistanceScore   As Double
End Type

Private Type BoundingBox
    minX            As Double
    maxX            As Double
    minY            As Double
    maxY            As Double
End Type

Private Type FrameElement
    frameName       As String
    startPt         As Point2D
    endPt           As Point2D
    Z               As Double
    Length          As Double
    UnitVector      As Point2D
    BBox            As BoundingBox
    IsActive        As Boolean
End Type

Public Type WallSegmentMap
    startX          As Double
    startY          As Double
    endX            As Double
    endY            As Double
    Length          As Double
    angle           As Double
    Thickness       As Double
    WallType        As String
    LoadPattern     As String
    LoadValue       As Double
    UnitX           As Double
    UnitY           As Double
End Type

Public Type MappingRecord
    TargetFrame     As String
    MapType         As String    ' "EXACT", "PARTIAL", "NEW"
    DistI           As Double    ' Distance from I-end (mm)
    DistJ           As Double    ' Distance from J-end (mm)
    FrameLength     As Double
End Type

Private Type MatchResult
    FrameIndex      As Long
    MatchType       As String
    overlapStart    As Double    ' Relative [0..1]
    overlapEnd      As Double    ' Relative [0..1]
    OverlapLength   As Double    ' mm
    overlapRatio    As Double    ' Coverage of wall
    Distance        As Double    ' Perpendicular distance
    Score           As Double
End Type


Private Type OverlapMatchResult
    FrameIndex      As Long
    MatchType       As String    ' "EXACT", "PARTIAL", "NEW"
    Score           As Double
    Distance        As Double
    overlapRatio    As Double
    overlapStart    As Double    ' Relative [0..1]
    overlapEnd      As Double    ' Relative [0..1]
    DistI           As Double    ' Distance in mm from Start
    DistJ           As Double    ' Distance in mm from Start
    MappedLength    As Double    ' Length in mm
End Type

Private Type IntervalNode_Map
    Low             As Double    ' Min on Range
    High            As Double    ' Max on Range
    MaxHigh         As Double    ' Max on Map
    FrameIndex      As Long    ' Index of mFrameDB
    LeftChild       As Long    ' Index Left Point (-1 if non exist)
    RightChild      As Long    ' Index Right Point (-1 if non exist)
End Type

Private mLogFilePath As String
Dim gLogFilePath    As String
Dim gLogEnabled     As Boolean


' ==================== MODULE VARIABLES ====================
Private mFrameDB()  As FrameElement
Private mFrameCount As Long
Private mSpatialHash As Object  ' Dictionary of SpatialCell
Private mIsInitialized As Boolean

Private mIntervalTreeHorizontal() As IntervalNode_Map
Private mIntervalTreeVertical() As IntervalNode_Map
Private mIntervalCountH As Long
Private mIntervalCountV As Long

' Vector strip
Private mSapFrameStatus() As SapFrameStatus
Private mSapFrameStatusCount As Long

Private mCadStripsVS() As CadLoadStrip
Private mCadStripCountVS As Long

Private mSapStripsVS() As SapStripInfo
Private mSapStripCountVS As Long

' For some legacy functions
Private Type FrameSegmentInfoDummy
    frameName As String
    startX As Double
    startY As Double
    endX As Double
    endY As Double
    Length As Double
    angle As Double
    BoundMinX As Double
    BoundMaxX As Double
    BoundMinY As Double
    BoundMaxY As Double
    UnitX As Double
    UnitY As Double
End Type

Private mFrameSegments() As FrameSegmentInfoDummy
' ===============================================================
' MAIN EXECUTOR: Batch Mapping
' ===============================================================

Public Function ExecuteWallMapping_Batch( _
        acadDoc As Object, _
        SapModel As Object, _
        frameNodeMap As Object, _
        selectedHandles As Variant, _
        InsertPt As Variant, _
        storyElev As Double _
        ) As Long

    On Error GoTo ErrHandler
    ExecuteWallMapping_Batch = 0

    ' Initialize logging for this batch
    Log_Init ""
    Log_Write "ExecuteWallMapping_Batch: START. storyElev=" & CStr(storyElev)

    ' ========== PHASE 1: VALIDATION ==========
    If acadDoc Is Nothing Or SapModel Is Nothing Or frameNodeMap Is Nothing Then
        Log_Write "ERROR: Null objects passed to ExecuteWallMapping_Batch"
        Debug.Print "ERROR: Null objects passed"
        Exit Function
    End If

    ' Validate insertPt
    Dim offsetX As Double, offsetY As Double
    offsetX = 0: offsetY = 0
    If IsArray(InsertPt) Then
        On Error Resume Next
        offsetX = CDbl(InsertPt(0))
        offsetY = CDbl(InsertPt(1))
        On Error GoTo ErrHandler
    End If

    ' Validate selectedHandles bounds
    Dim iterLB As Long, iterUB As Long
    Dim hasSelectionArray As Boolean
    hasSelectionArray = False
    On Error Resume Next
    If Not IsEmpty(selectedHandles) Then
        If IsArray(selectedHandles) Then
            iterLB = LBound(selectedHandles)
            iterUB = UBound(selectedHandles)
            hasSelectionArray = True
        End If
    End If
    On Error GoTo ErrHandler

    If Not hasSelectionArray Then
        Log_Write "No valid selectedHandles array passed to ExecuteWallMapping_Batch"
        Exit Function
    End If

    ' Check if 2D array
    Dim is2D As Boolean
    is2D = False
    On Error Resume Next
    Dim test2D As Variant
    test2D = selectedHandles(iterLB, 0)
    If err.number = 0 Then is2D = True
    err.Clear
    On Error GoTo ErrHandler

    Log_Write "PHASE: Building frame spatial database..."

    ' ========== PHASE 2: BUILD FRAME DATABASE ==========
    If Not mIsInitialized Then
        Debug.Print "Building frame spatial database..."
        If Not BuildFrameSpatialDatabase(SapModel, frameNodeMap, storyElev) Then
            Log_Write "ERROR: Failed to build frame database at elevation: " & CStr(storyElev)
            Debug.Print "ERROR: Failed to build frame database"
            Exit Function
        Else
            Log_Write "Frame spatial database built. FrameCount=" & CStr(mFrameCount)
        End If
    End If

    If mFrameCount = 0 Then
        Log_Write "ERROR: No frames at elevation " & CStr(storyElev)
        Debug.Print "ERROR: No frames at elevation " & storyElev
        Exit Function
    End If

    ' ========== PHASE 3: BATCH PROCESS WALLS ==========
    Dim processCount As Long, failCount As Long, skippedCount As Long
    Dim colorNew As Long, colorPartial As Long, colorExact As Long
    processCount = 0: failCount = 0: skippedCount = 0
    colorNew = 0: colorPartial = 0: colorExact = 0

    Dim i As Long
    For i = iterLB To iterUB
        On Error Resume Next

        ' Get handle string depending on 1D/2D passed selection
        Dim hStr As String
        If is2D Then
            hStr = CStr(selectedHandles(i, 0))
        Else
            hStr = CStr(selectedHandles(i))
        End If

        Dim ent As Object
        Set ent = Nothing
        On Error Resume Next
        Set ent = acadDoc.HandleToObject(hStr)
        If err.number <> 0 Then
            err.Clear
            Set ent = Nothing
        End If
        On Error GoTo ErrHandler

        If ent Is Nothing Then
            skippedCount = skippedCount + 1
            Log_Write "Skip: handle " & hStr & " not found in AutoCAD."
            GoTo NextWallLoop
        End If

        ' === COMMON FILTER: require correct layer and valid DTS_APP xdata ===
        If Not IsValidWallWithXData(ent) Then
            skippedCount = skippedCount + 1
            Log_Write "Skip: handle " & hStr & " is not valid wall (wrong layer or missing XData)."
            GoTo NextWallLoop
        End If

        ' Extract wall geometry and header XData (strict)
        Dim wallSeg As WallSegmentMap
        If Not ExtractWallFromEntity(ent, wallSeg, offsetX, offsetY) Then
            skippedCount = skippedCount + 1
            Log_Write "Skip: handle " & hStr & " failed ExtractWallFromEntity."
            GoTo NextWallLoop
        End If

        Log_Write "Processing handle " & hStr & " -> wall len=" & Format(wallSeg.Length, "0.00") & " mm; thickness=" & CStr(wallSeg.Thickness) & "; type=" & wallSeg.WallType

        ' Find matches
        Dim matches() As MatchResult
        Dim matchCount As Long
        matchCount = FindMatchingFrames(wallSeg, storyElev, matches)
        Log_Write "Found " & CStr(matchCount) & " candidate matches for handle " & hStr

        ' Convert to mappings
        Dim mappings() As MappingRecord
        Dim mappingCount As Long
        mappingCount = ConvertToMappings(matches, matchCount, mappings)

        ' Determine color and decision
        Dim colorIdx As Integer
        If matchCount = 0 Then
            colorIdx = 1  ' Red - New frame needed
            colorNew = colorNew + 1
            Log_Write "Decision: No matches -> CREATE NEW frame for handle " & hStr

            ReDim mappings(0 To 0)
            mappings(0).TargetFrame = "New"
            mappings(0).MapType = "NEW"
            mappingCount = 1
        Else
            Dim totalCoverage As Double
            totalCoverage = CalculateTotalCoverage(matches, matchCount, wallSeg.Length)

            If totalCoverage >= 0.95 Then
                colorIdx = 3  ' Green - Full coverage
                colorExact = colorExact + 1
                Log_Write "Decision: Full coverage (" & Format(totalCoverage, "0.00") & ") for handle " & hStr
            Else
                colorIdx = 2  ' Yellow - Partial coverage
                colorPartial = colorPartial + 1
                Log_Write "Decision: Partial coverage (" & Format(totalCoverage, "0.00") & ") for handle " & hStr
            End If
        End If

        ' Apply color
        On Error Resume Next
        ent.color = colorIdx
        On Error GoTo ErrHandler

        ' Write XData (Update the mapping info)
        ' Preserve base header; no override here -> pass an empty MappingRecord array (count 0)
        Dim dummyOv() As MappingRecord
        ReDim dummyOv(0 To 0)
        WriteWallCompleteXData_WithOverride ent, wallSeg, mappings, mappingCount, False, wallSeg, dummyOv, 0
        Log_Write "Wrote XData for handle " & hStr & " with " & CStr(mappingCount) & " mapping records."

        processCount = processCount + 1

NextWallLoop:
        ' continue loop
    Next i

    ExecuteWallMapping_Batch = processCount

    Log_Write "BATCH COMPLETE. Processed=" & CStr(processCount) & ", Skipped=" & CStr(skippedCount) & _
              ", New=" & CStr(colorNew) & ", Partial=" & CStr(colorPartial) & ", Exact=" & CStr(colorExact)

    ' Open notepad to view log
    Log_OpenNotepad

    Debug.Print "========== BATCH COMPLETE =========="
    Debug.Print "Processed: " & processCount & ", Skipped: " & skippedCount
    Debug.Print "New: " & colorNew & ", Partial: " & colorPartial & ", Exact: " & colorExact

    Exit Function

ErrHandler:
    Log_Write "CRITICAL ERROR in ExecuteWallMapping_Batch: " & err.description & " (Err# " & CStr(err.number) & ")"
    Debug.Print "CRITICAL ERROR in ExecuteWallMapping_Batch: " & err.description
    ExecuteWallMapping_Batch = processCount
End Function
Private Function CalculateTotalCoverageFromRecords(recs() As MappingRecord, count As Long, wallLength As Double) As Double
    CalculateTotalCoverageFromRecords = 0
    If count <= 0 Or wallLength <= 0 Then Exit Function

    Dim totalLen As Double
    Dim i As Long
    For i = 0 To count - 1
        totalLen = totalLen + (recs(i).DistJ - recs(i).DistI)
    Next i

    CalculateTotalCoverageFromRecords = totalLen / wallLength
End Function

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

    If err.number <> 0 Then
        err.Clear
        Exit Function
    End If

    If IsEmpty(xdVal) Or Not IsArray(xdVal) Then Exit Function

    Dim maxIndex As Long
    maxIndex = UBound(xdVal)

    ' Read base wall properties
    If maxIndex >= XDATA_OFFSET_THICKNESS Then wallSeg.Thickness = CDbl(xdVal(XDATA_OFFSET_THICKNESS))
    If maxIndex >= XDATA_OFFSET_WALLTYPE Then wallSeg.WallType = CStr(xdVal(XDATA_OFFSET_WALLTYPE))
    If maxIndex >= XDATA_OFFSET_LOADPATTERN Then wallSeg.LoadPattern = CStr(xdVal(XDATA_OFFSET_LOADPATTERN))
    If maxIndex >= XDATA_OFFSET_LOADVALUE Then wallSeg.LoadValue = CDbl(xdVal(XDATA_OFFSET_LOADVALUE))

    ' Base mappings
    Dim remainingSize As Long
    remainingSize = maxIndex - XDATA_OFFSET_MAPPING_START + 1

    Dim baseCount As Long
    baseCount = 0
    Dim cursor As Long
    cursor = XDATA_OFFSET_MAPPING_START

    If remainingSize >= XDATA_MAPPING_RECORD_SIZE Then
        Dim totalSlots As Long
        totalSlots = (maxIndex - XDATA_OFFSET_MAPPING_START + 1)

        Dim possibleCount As Long
        possibleCount = totalSlots \ XDATA_MAPPING_RECORD_SIZE

        If possibleCount > 0 Then
            ReDim baseMappings(0 To possibleCount - 1)
            Dim i As Long
            For i = 0 To possibleCount - 1
                If cursor + 3 > maxIndex Then Exit For
                If TypeName(xdVal(cursor)) = "String" And _
                   TypeName(xdVal(cursor + 1)) = "String" Then
                    baseMappings(baseCount).TargetFrame = CStr(xdVal(cursor))
                    baseMappings(baseCount).MapType = CStr(xdVal(cursor + 1))
                    baseMappings(baseCount).DistI = CDbl(xdVal(cursor + 2))
                    baseMappings(baseCount).DistJ = CDbl(xdVal(cursor + 3))
                    baseMappings(baseCount).FrameLength = baseMappings(baseCount).DistJ - baseMappings(baseCount).DistI
                    baseCount = baseCount + 1
                    cursor = cursor + XDATA_MAPPING_RECORD_SIZE
                Else
                    Exit For
                End If
            Next i

            If baseCount > 0 Then
                ReDim Preserve baseMappings(0 To baseCount - 1)
            Else
                Erase baseMappings
            End If
        End If
    End If

    ReadWallAllXData = baseCount

    ' Now try to read override block (if present)
    If cursor > maxIndex Then Exit Function

    ' Override flag stored as Int32 at current cursor index
    If TypeName(xdVal(cursor)) = "Long" Or TypeName(xdVal(cursor)) = "Integer" Then
        Dim ovFlag As Long
        ovFlag = CLng(xdVal(cursor))
        If ovFlag <> 0 Then
            hasOverride = True
        End If
        cursor = cursor + 1
    Else
        Exit Function
    End If

    If Not hasOverride Then Exit Function

    ' Override wall type/pattern/value
    If cursor > maxIndex Then Exit Function
    If TypeName(xdVal(cursor)) = "String" Then
        ovWallSeg.WallType = CStr(xdVal(cursor))
        cursor = cursor + 1
    End If

    If cursor > maxIndex Then Exit Function
    If TypeName(xdVal(cursor)) = "String" Then
        ovWallSeg.LoadPattern = CStr(xdVal(cursor))
        cursor = cursor + 1
    End If

    If cursor > maxIndex Then Exit Function
    If IsNumeric(xdVal(cursor)) Then
        ovWallSeg.LoadValue = CDbl(xdVal(cursor))
        cursor = cursor + 1
    End If

    ' Thickness for override = base thickness
    ovWallSeg.Thickness = wallSeg.Thickness

    ' Override mappings
    If cursor > maxIndex Then Exit Function

    Dim ovSlotCount As Long
    ovSlotCount = (maxIndex - cursor + 1)

    If ovSlotCount < XDATA_MAPPING_RECORD_SIZE Then Exit Function

    Dim maxOVCount As Long
    maxOVCount = ovSlotCount \ XDATA_MAPPING_RECORD_SIZE

    If maxOVCount <= 0 Then Exit Function

    ReDim ovMappings(0 To maxOVCount - 1)

    Dim j As Long
    For j = 0 To maxOVCount - 1
        If cursor + 3 > maxIndex Then Exit For
        If TypeName(xdVal(cursor)) = "String" And TypeName(xdVal(cursor + 1)) = "String" Then
            ovMappings(ovCount).TargetFrame = CStr(xdVal(cursor))
            ovMappings(ovCount).MapType = CStr(xdVal(cursor + 1))
            ovMappings(ovCount).DistI = CDbl(xdVal(cursor + 2))
            ovMappings(ovCount).DistJ = CDbl(xdVal(cursor + 3))
            ovMappings(ovCount).FrameLength = ovMappings(ovCount).DistJ - ovMappings(ovCount).DistI
            ovCount = ovCount + 1
            cursor = cursor + XDATA_MAPPING_RECORD_SIZE
        Else
            Exit For
        End If
    Next j

    If ovCount > 0 Then
        ReDim Preserve ovMappings(0 To ovCount - 1)
    Else
        Erase ovMappings
        hasOverride = False
    End If
End Function
' ==================== UPDATED WRITE FUNCTION (Standard Header) ====================
Public Sub WriteWallCompleteXData_WithOverride( _
    entObj As Object, _
    baseWall As WallSegmentMap, _
    baseMappings() As MappingRecord, _
    baseCount As Long, _
    hasOverride As Boolean, _
    ovWallSeg As WallSegmentMap, _
    ovMappings() As MappingRecord, _
    ovCount As Long)

    On Error Resume Next
    If entObj Is Nothing Then Exit Sub

    Dim acadDoc As Object
    Set acadDoc = entObj.Application.ActiveDocument
    acadDoc.RegisteredApplications.Add XDATA_APP

    ' Calculate Size
    ' Header: 0-4 (5 items)
    Dim totalSize As Long
    totalSize = XDATA_OFFSET_MAPPING_START + (baseCount * XDATA_MAPPING_RECORD_SIZE)

    If hasOverride Then
        ' Flag(1) + Type(1) + Pat(1) + Val(1) = 4 items overhead for override header
        totalSize = totalSize + 4
        totalSize = totalSize + (ovCount * XDATA_MAPPING_RECORD_SIZE)
    End If

    Dim xdType() As Integer
    Dim xdVal() As Variant
    ReDim xdType(0 To totalSize - 1)
    ReDim xdVal(0 To totalSize - 1)

    ' --- WRITE HEADER (Indices 0-4) ---
    ' 0: App
    xdType(0) = XD_CODE_APPNAME: xdVal(0) = XDATA_APP
    ' 1: Thickness
    xdType(XDATA_OFFSET_THICKNESS) = XD_CODE_REAL: xdVal(XDATA_OFFSET_THICKNESS) = baseWall.Thickness
    ' 2: WallType
    xdType(XDATA_OFFSET_WALLTYPE) = XD_CODE_STRING: xdVal(XDATA_OFFSET_WALLTYPE) = baseWall.WallType
    ' 3: LoadPattern
    xdType(XDATA_OFFSET_LOADPATTERN) = XD_CODE_STRING
    If baseWall.LoadPattern = "" Then xdVal(XDATA_OFFSET_LOADPATTERN) = "DL" Else xdVal(XDATA_OFFSET_LOADPATTERN) = baseWall.LoadPattern
    ' 4: LoadValue
    xdType(XDATA_OFFSET_LOADVALUE) = XD_CODE_REAL: xdVal(XDATA_OFFSET_LOADVALUE) = baseWall.LoadValue

    ' --- WRITE BASE MAPPINGS ---
    Dim i As Long, Offset As Long
    For i = 0 To baseCount - 1
        Offset = XDATA_OFFSET_MAPPING_START + (i * XDATA_MAPPING_RECORD_SIZE)
        xdType(Offset) = XD_CODE_STRING: xdVal(Offset) = baseMappings(i).TargetFrame
        xdType(Offset + 1) = XD_CODE_STRING: xdVal(Offset + 1) = baseMappings(i).MapType
        xdType(Offset + 2) = XD_CODE_REAL: xdVal(Offset + 2) = baseMappings(i).DistI
        xdType(Offset + 3) = XD_CODE_REAL: xdVal(Offset + 3) = baseMappings(i).DistJ
    Next i

    ' --- WRITE OVERRIDE ---
    Offset = XDATA_OFFSET_MAPPING_START + (baseCount * XDATA_MAPPING_RECORD_SIZE)

    If hasOverride Then
        ' Flag
        xdType(Offset) = XD_CODE_INT32: xdVal(Offset) = 1
        ' OV Header
        xdType(Offset + 1) = XD_CODE_STRING: xdVal(Offset + 1) = ovWallSeg.WallType
        xdType(Offset + 2) = XD_CODE_STRING: xdVal(Offset + 2) = ovWallSeg.LoadPattern
        xdType(Offset + 3) = XD_CODE_REAL: xdVal(Offset + 3) = ovWallSeg.LoadValue
        
        Offset = Offset + 4

        ' OV Mappings
        For i = 0 To ovCount - 1
            xdType(Offset) = XD_CODE_STRING: xdVal(Offset) = ovMappings(i).TargetFrame
            xdType(Offset + 1) = XD_CODE_STRING: xdVal(Offset + 1) = ovMappings(i).MapType
            xdType(Offset + 2) = XD_CODE_REAL: xdVal(Offset + 2) = ovMappings(i).DistI
            xdType(Offset + 3) = XD_CODE_REAL: xdVal(Offset + 3) = ovMappings(i).DistJ
            Offset = Offset + XDATA_MAPPING_RECORD_SIZE
        Next i
    End If

    entObj.SetXData xdType, xdVal
End Sub
' ==================== BUILD SAP STRIPS FOR VECTOR STRIP PROJECTION ====================
Private Sub BuildSapStripsVS()
    On Error Resume Next

    mSapStripCountVS = 0
    If mFrameCount <= 0 Then Exit Sub

    ReDim mSapStripsVS(0 To mFrameCount - 1)

    Dim i As Long, j As Long
    For i = 0 To mFrameCount - 1
        If Not mFrameDB(i).IsActive Then GoTo NextFrameVS
        If mFrameDB(i).Length <= 1 Then GoTo NextFrameVS

        ' Already assigned to a strip?
        Dim already As Boolean
        already = False
        For j = 0 To mSapStripCountVS - 1
            If IsFrameCollinearToStrip(i, j) Then
                ' Add to existing strip
                With mSapStripsVS(j)
                    .FrameIndices(.FrameCount) = i
                    .FrameCount = .FrameCount + 1
                    .TotalLength = .TotalLength + mFrameDB(i).Length
                End With
                already = True
                Exit For
            End If
        Next j

        If Not already Then
            ' Create new strip
            With mSapStripsVS(mSapStripCountVS)
                .StripID = mSapStripCountVS
                Dim dx As Double, dy As Double, l As Double
                dx = mFrameDB(i).endPt.X - mFrameDB(i).startPt.X
                dy = mFrameDB(i).endPt.Y - mFrameDB(i).startPt.Y
                l = Sqr(dx * dx + dy * dy)
                If l <= 0.001 Then GoTo NextFrameVS

                .angle = Atan2(dy, dx)
                .UnitX = dx / l
                .UnitY = dy / l
                .OriginX = mFrameDB(i).startPt.X
                .OriginY = mFrameDB(i).startPt.Y

                ReDim .FrameIndices(0 To mFrameCount - 1)
                .FrameIndices(0) = i
                .FrameCount = 1
                .TotalLength = mFrameDB(i).Length
            End With

            mSapStripCountVS = mSapStripCountVS + 1
        End If

NextFrameVS:
    Next i

    If mSapStripCountVS > 0 Then
        ReDim Preserve mSapStripsVS(0 To mSapStripCountVS - 1)
    End If

    ' Build frame status map
    ReDim mSapFrameStatus(0 To mFrameCount - 1)
    mSapFrameStatusCount = mFrameCount
    For i = 0 To mFrameCount - 1
        With mSapFrameStatus(i)
            .FrameIndex = i
            .Length = mFrameDB(i).Length
            .OccupiedCount = 0
            ReDim .Occupied(0 To 10)
        End With
    Next i
End Sub

Private Function IsFrameCollinearToStrip(frameIdx As Long, stripIdx As Long) As Boolean
    On Error Resume Next
    IsFrameCollinearToStrip = False

    Dim dx As Double, dy As Double, l As Double
    dx = mFrameDB(frameIdx).endPt.X - mFrameDB(frameIdx).startPt.X
    dy = mFrameDB(frameIdx).endPt.Y - mFrameDB(frameIdx).startPt.Y
    l = Sqr(dx * dx + dy * dy)
    If l <= 0.001 Then Exit Function

    Dim ang         As Double
    ang = Atan2(dy, dx)

    Dim angDiff     As Double
    angDiff = Abs(ang - mSapStripsVS(stripIdx).angle)
    If angDiff > PI Then angDiff = 2 * PI - angDiff

    ' === FIX START: Allow 180 degree opposite vectors ===
    If angDiff > ANGLE_TOLERANCE Then
        ' Check if it is opposite direction (approx PI)
        If Abs(angDiff - PI) > ANGLE_TOLERANCE Then Exit Function
    End If
    ' === FIX END ===

    ' Perpendicular distance from frame mid to strip line
    Dim midX As Double, midY As Double
    midX = (mFrameDB(frameIdx).startPt.X + mFrameDB(frameIdx).endPt.X) / 2
    midY = (mFrameDB(frameIdx).startPt.Y + mFrameDB(frameIdx).endPt.Y) / 2

    Dim dist        As Double
    dist = PointToLineDistanceInfinite(midX, midY, _
            mSapStripsVS(stripIdx).OriginX, mSapStripsVS(stripIdx).OriginY, _
            mSapStripsVS(stripIdx).OriginX + mSapStripsVS(stripIdx).UnitX * 1000#, _
            mSapStripsVS(stripIdx).OriginY + mSapStripsVS(stripIdx).UnitY * 1000#)

    IsFrameCollinearToStrip = (dist <= COORD_TOLERANCE)
End Function


' ===============================================================
' CORE: Build Frame Spatial Database
' ===============================================================
Private Function BuildFrameSpatialDatabase(SapModel As Object, frameNodeMap As Object, zLevel As Double) As Boolean
    On Error GoTo ErrHandler

    BuildFrameSpatialDatabase = False
    mFrameCount = 0
    mIsInitialized = False

    If frameNodeMap Is Nothing Then Exit Function
    If frameNodeMap.count = 0 Then Exit Function

    ' Initialize spatial hash
    Set mSpatialHash = CreateObject("Scripting.Dictionary")

    ' Allocate frame array
    ReDim mFrameDB(0 To frameNodeMap.count - 1)

    Dim frameKey    As Variant
    Dim nodes       As Collection
    Dim pt1 As String, pt2 As String
    Dim x1 As Double, y1 As Double, z1 As Double
    Dim x2 As Double, y2 As Double, z2 As Double
    Dim ret         As Long

    Dim idx         As Long
    idx = 0

    For Each frameKey In frameNodeMap.keys
        On Error Resume Next
        Set nodes = frameNodeMap(frameKey)

        If err.number <> 0 Then
            err.Clear
            GoTo NextFrame
        End If

        If nodes Is Nothing Then GoTo NextFrame
        If nodes.count < 2 Then GoTo NextFrame

        pt1 = CStr(nodes(1))
        pt2 = CStr(nodes(2))

        ret = SapModel.pointObj.GetCoordCartesian(pt1, x1, y1, z1)
        If ret <> 0 Then GoTo NextFrame

        ret = SapModel.pointObj.GetCoordCartesian(pt2, x2, y2, z2)
        If ret <> 0 Then GoTo NextFrame

        On Error GoTo ErrHandler

        ' Check Z-level (with tolerance)
        Dim avgZ    As Double
        avgZ = (z1 + z2) / 2
        If Abs(avgZ - zLevel) > 100 Then GoTo NextFrame  ' 100mm tolerance

        ' Store frame data
        With mFrameDB(idx)
            .frameName = CStr(frameKey)
            .startPt.X = x1
            .startPt.Y = y1
            .endPt.X = x2
            .endPt.Y = y2
            .Z = avgZ

            Dim dx As Double, dy As Double
            dx = x2 - x1
            dy = y2 - y1
            .Length = Sqr(dx * dx + dy * dy)

            If .Length > 0.001 Then
                .UnitVector.X = dx / .Length
                .UnitVector.Y = dy / .Length

                ' Bounding box
                .BBox.minX = IIf(x1 < x2, x1, x2)
                .BBox.maxX = IIf(x1 > x2, x1, x2)
                .BBox.minY = IIf(y1 < y2, y1, y2)
                .BBox.maxY = IIf(y1 > y2, y1, y2)

                .IsActive = True

                ' Add to spatial hash
                AddFrameToSpatialHash idx

                idx = idx + 1
            End If
        End With

NextFrame:
    Next frameKey

    mFrameCount = idx

    If mFrameCount > 0 Then
        ReDim Preserve mFrameDB(0 To mFrameCount - 1)
        mIsInitialized = True
        BuildFrameSpatialDatabase = True
Debug.Print "Spatial database built: " & mFrameCount & " frames, " & mSpatialHash.count & " cells"
    End If

    Exit Function

ErrHandler:
Debug.Print "ERROR in BuildFrameSpatialDatabase: " & err.description
    BuildFrameSpatialDatabase = False
End Function
' ===============================================================
' SPATIAL HASH: Add frame to grid cells (FIXED FOR VARIANT ERROR)
' ===============================================================
Private Sub AddFrameToSpatialHash(frameIdx As Long)
    On Error Resume Next

    With mFrameDB(frameIdx)
        ' Get cell range
        Dim cellX1 As Long, cellY1 As Long
        Dim cellX2 As Long, cellY2 As Long

        cellX1 = Int(.BBox.minX / GRID_CELL_SIZE)
        cellY1 = Int(.BBox.minY / GRID_CELL_SIZE)
        cellX2 = Int(.BBox.maxX / GRID_CELL_SIZE)
        cellY2 = Int(.BBox.maxY / GRID_CELL_SIZE)

        ' Add to all intersecting cells
        Dim cx As Long, cy As Long
        For cx = cellX1 To cellX2
            For cy = cellY1 To cellY2
                Dim cellKey As String
                cellKey = cx & "," & cy

                Dim indices() As Long

                If mSpatialHash.exists(cellKey) Then
                    ' Retrieve existing array
                    indices = mSpatialHash(cellKey)

                    ' Expand array
                    ReDim Preserve indices(0 To UBound(indices) + 1)
                    indices(UBound(indices)) = frameIdx
                Else
                    ' Create new array
                    ReDim indices(0 To 0)
                    indices(0) = frameIdx
                End If

                ' Store back to dictionary
                mSpatialHash(cellKey) = indices
            Next cy
        Next cx
    End With
End Sub

' ===============================================================
' CORE: Find Matching Frames (Spatial Query)
' ===============================================================
Private Function FindMatchingFrames(wallSeg As WallSegmentMap, zLevel As Double, ByRef matches() As MatchResult) As Long
    On Error Resume Next

    FindMatchingFrames = 0

    ' Get candidate frames from spatial hash
    Dim candidates() As Long
    Dim candidateCount As Long
    candidateCount = GetSpatialCandidates(wallSeg, candidates)

    Log_Write "FindMatchingFrames: candidateCount=" & CStr(candidateCount)

    If candidateCount = 0 Then
        ' Detailed hint: spatial search returned 0 (GetSpatialCandidates already logged bbox/cells)
        Log_Write "FindMatchingFrames: No spatial candidates -> cannot match. Check GRID_CELL_SIZE/COORD_TOLERANCE or frame positions."
        Exit Function
    End If

    ' Analyze each candidate
    ReDim matches(0 To candidateCount - 1)
    Dim matchCount  As Long
    matchCount = 0

    ' Counters for rejection reasons
    Dim cntAngle    As Long: cntAngle = 0
    Dim cntDist     As Long: cntDist = 0
    Dim cntOverlap  As Long: cntOverlap = 0
    Dim cntOther    As Long: cntOther = 0

    Dim i           As Long
    For i = 0 To candidateCount - 1
        Dim result  As MatchResult
        Dim reason  As String
        reason = ""
        If AnalyzeWallFrameMatch(wallSeg, candidates(i), result, reason) Then
            If result.overlapRatio >= MIN_OVERLAP_RATIO Then
                matches(matchCount) = result
                matchCount = matchCount + 1
            Else
                ' Rejected: low overlap despite passing checks
                Log_Write "CandidateRejected: frameIdx=" & CStr(candidates(i)) & " (" & mFrameDB(candidates(i)).frameName & ") reason=low_overlap overlapRatio=" & Format(result.overlapRatio, "0.000")
                cntOverlap = cntOverlap + 1
            End If
        Else
            ' AnalyzeWallFrameMatch returned False; reason should be set
            If Len(reason) = 0 Then reason = "unknown"
            Log_Write "CandidateRejected: frameIdx=" & CStr(candidates(i)) & " (" & mFrameDB(candidates(i)).frameName & ") reason=" & reason

            If InStr(1, reason, "angle_diff", vbTextCompare) > 0 Then
                cntAngle = cntAngle + 1
            ElseIf InStr(1, reason, "distMid", vbTextCompare) > 0 Then
                cntDist = cntDist + 1
            ElseIf InStr(1, reason, "overlapLen", vbTextCompare) > 0 Then
                cntOverlap = cntOverlap + 1
            Else
                cntOther = cntOther + 1
            End If
        End If
    Next i

    If matchCount > 0 Then
        ReDim Preserve matches(0 To matchCount - 1)

        ' Sort by score (best first)
        SortMatchesByScore matches, matchCount

        FindMatchingFrames = matchCount
        Log_Write "FindMatchingFrames: matchedCount=" & CStr(matchCount)
    Else
        Erase matches
        FindMatchingFrames = 0
        Log_Write "FindMatchingFrames: NO matches after analyzing " & CStr(candidateCount) & " candidates. RejectionSummary: angle=" & CStr(cntAngle) & ", distance=" & CStr(cntDist) & ", overlap=" & CStr(cntOverlap) & ", other=" & CStr(cntOther)
    End If
End Function
' ===============================================================
' SPATIAL HASH: Get candidate frames (FIXED FOR ARRAY STORAGE)
' ===============================================================
Private Function GetSpatialCandidates(wallSeg As WallSegmentMap, ByRef candidates() As Long) As Long
    On Error Resume Next

    GetSpatialCandidates = 0

    Dim searchRadius As Double
    searchRadius = 1000   ' Increased from COORD_TOLERANCE (250)

    ' Get wall bounding box (with tolerance)
    Dim wMinX As Double, wMaxX As Double
    Dim wMinY As Double, wMaxY As Double

    wMinX = IIf(wallSeg.startX < wallSeg.endX, wallSeg.startX, wallSeg.endX) - searchRadius
    wMaxX = IIf(wallSeg.startX > wallSeg.endX, wallSeg.startX, wallSeg.endX) + searchRadius
    wMinY = IIf(wallSeg.startY < wallSeg.endY, wallSeg.startY, wallSeg.endY) - searchRadius
    wMaxY = IIf(wallSeg.startY > wallSeg.endY, wallSeg.startY, wallSeg.endY) + searchRadius

    Log_Write "GetSpatialCandidates: wall bbox X[" & Format(wMinX, "0.0") & "," & Format(wMaxX, "0.0") & "] Y[" & Format(wMinY, "0.0") & "," & Format(wMaxY, "0.0") & "]"

    ' Get cell range
    Dim cellX1 As Long, cellY1 As Long
    Dim cellX2 As Long, cellY2 As Long

    cellX1 = Int(wMinX / GRID_CELL_SIZE)
    cellY1 = Int(wMinY / GRID_CELL_SIZE)
    cellX2 = Int(wMaxX / GRID_CELL_SIZE)
    cellY2 = Int(wMaxY / GRID_CELL_SIZE)

    Log_Write "GetSpatialCandidates: checking grid cells X[" & CStr(cellX1) & ".." & CStr(cellX2) & "] Y[" & CStr(cellY1) & ".." & CStr(cellY2) & "] (GRID_CELL_SIZE=" & CStr(GRID_CELL_SIZE) & ")"

    ' Collect unique frame indices
    Dim uniqueDict  As Object
    Set uniqueDict = CreateObject("Scripting.Dictionary")

    Dim cx As Long, cy As Long
    Dim cellsScanned As Long: cellsScanned = 0
    Dim cellsWithFrames As Long: cellsWithFrames = 0
    For cx = cellX1 To cellX2
        For cy = cellY1 To cellY2
            cellsScanned = cellsScanned + 1
            Dim cellKey As String
            cellKey = cx & "," & cy

            If mSpatialHash.exists(cellKey) Then
                cellsWithFrames = cellsWithFrames + 1
                ' Retrieve array directly
                Dim indices() As Long
                indices = mSpatialHash(cellKey)

                Dim j As Long
                Dim frameIdx As Long
                For j = LBound(indices) To UBound(indices)
                    frameIdx = indices(j)
                    If Not uniqueDict.exists(frameIdx) Then
                        uniqueDict.Add frameIdx, True
                    End If
                Next j
            End If
        Next cy
    Next cx

    Log_Write "GetSpatialCandidates: scannedCells=" & CStr(cellsScanned) & ", cellsWithFrames=" & CStr(cellsWithFrames) & ", spatialHashCells=" & CStr(IIf(Not mSpatialHash Is Nothing, mSpatialHash.count, 0))

    ' Convert to array
    If uniqueDict.count > 0 Then
        ReDim candidates(0 To uniqueDict.count - 1)

        Dim idx     As Long
        idx = 0
        Dim key     As Variant
        For Each key In uniqueDict.keys
            candidates(idx) = CLng(key)
            idx = idx + 1
        Next key

        GetSpatialCandidates = uniqueDict.count

        ' Log sample of candidates (up to 20)
        Dim sampleCount As Long
        sampleCount = IIf(GetSpatialCandidates > 20, 20, GetSpatialCandidates)
        Dim samp    As Long
        Dim sampList As String
        sampList = ""
        For samp = 0 To sampleCount - 1
            sampList = sampList & CStr(candidates(samp)) & "(" & mFrameDB(candidates(samp)).frameName & ")"
            If samp < sampleCount - 1 Then sampList = sampList & ", "
        Next samp
        Log_Write "GetSpatialCandidates: uniqueCandidates=" & CStr(GetSpatialCandidates) & " sample=[" & sampList & "]"
    Else
        Log_Write "GetSpatialCandidates: NO candidates found in spatial hash for this wall."
    End If
End Function
' ===============================================================
' GEOMETRY: Analyze Wall-Frame Match (FIXED LOGIC v7 - OVERLAP PRIORITY)
' 1. CHECK OVERLAP FIRST: If physical overlap exists (>200mm), ACCEPT regardless of lengths.
'    (This fixes the issue where 5m wall on 10m beam was rejected).
' 2. CHECK GAP SECOND: If no overlap, allow Gap Match ONLY IF lengths are similar.
'    (This prevents 2m orphan wall from snapping to 8m beam).
' ===============================================================
Private Function AnalyzeWallFrameMatch(wallSeg As WallSegmentMap, frameIdx As Long, ByRef result As MatchResult, ByRef debugReason As String) As Boolean
    On Error Resume Next
    AnalyzeWallFrameMatch = False
    debugReason = ""

    With mFrameDB(frameIdx)
        ' ---------------------------------------------------------
        ' STEP 1: CHECK PARALLELISM (ANGLE)
        ' ---------------------------------------------------------
        Dim wallAngle As Double
        wallAngle = Atan2(wallSeg.endY - wallSeg.startY, wallSeg.endX - wallSeg.startX)

        Dim frameAngle As Double
        frameAngle = Atan2(.endPt.Y - .startPt.Y, .endPt.X - .startPt.X)

        Dim angleDiff As Double
        angleDiff = Abs(wallAngle - frameAngle)

        If angleDiff > PI Then angleDiff = 2 * PI - angleDiff

        ' Allow 0 rad (Same) OR PI rad (Opposite)
        Dim isParallel As Boolean
        isParallel = (angleDiff <= ANGLE_TOLERANCE) Or (Abs(angleDiff - PI) <= ANGLE_TOLERANCE)

        If Not isParallel Then
            debugReason = "angle_diff=" & Format(angleDiff, "0.000") & " > TOL"
            Exit Function
        End If

        ' ---------------------------------------------------------
        ' STEP 2: CHECK DISTANCE (Perpendicular)
        ' ---------------------------------------------------------
        Dim distMid As Double
        Dim wallMidX As Double, wallMidY As Double
        wallMidX = (wallSeg.startX + wallSeg.endX) / 2
        wallMidY = (wallSeg.startY + wallSeg.endY) / 2

        distMid = PointToLineDistance(wallMidX, wallMidY, .startPt.X, .startPt.Y, .endPt.X, .endPt.Y)

        Dim dynamicDistTol As Double
        dynamicDistTol = wallSeg.Thickness * 5#
        If dynamicDistTol < 250 Then dynamicDistTol = 250
        If dynamicDistTol > 1500 Then dynamicDistTol = 1500

        If distMid > dynamicDistTol Then
            debugReason = "distMid=" & Format(distMid, "0.0") & " > DynTol(" & Format(dynamicDistTol, "0.0") & ")"
            Exit Function
        End If

        ' ---------------------------------------------------------
        ' STEP 3: PROJECT & CALCULATE OVERLAP (CORE LOGIC)
        ' ---------------------------------------------------------
        Dim t1 As Double, t2 As Double
        t1 = ((wallSeg.startX - .startPt.X) * .UnitVector.X + (wallSeg.startY - .startPt.Y) * .UnitVector.Y)
        t2 = ((wallSeg.endX - .startPt.X) * .UnitVector.X + (wallSeg.endY - .startPt.Y) * .UnitVector.Y)

        If t1 > t2 Then
            Dim tmp As Double: tmp = t1: t1 = t2: t2 = tmp
        End If

        Dim frameLen As Double: frameLen = .Length

        ' Calculate Physical Intersection with Frame [0, frameLen]
        Dim overlapStart As Double, overlapEnd As Double
        overlapStart = IIf(t1 > 0, t1, 0)
        overlapEnd = IIf(t2 < frameLen, t2, frameLen)

        Dim overlapLen As Double
        If overlapEnd > overlapStart Then
            overlapLen = overlapEnd - overlapStart
        Else
            overlapLen = 0
        End If

        ' ---------------------------------------------------------
        ' STEP 4: DECISION LOGIC (The Fix)
        ' ---------------------------------------------------------
        Dim isValidMatch As Boolean
        isValidMatch = False

        ' === PRIORITY A: PHYSICAL OVERLAP ===
        ' If wall sits ON the frame significantly, it's a match.
        ' Even if Wall=5m and Frame=10m, overlapLen=5m -> Ratio=1.0 -> ACCEPT.
        If overlapLen > 200 Then
            ' Check if overlap covers a significant portion of the WALL
            If (overlapLen / wallSeg.Length) >= 0.15 Then
                isValidMatch = True
            End If
        End If

        ' === PRIORITY B: GAP MATCH (ONLY IF NO OVERLAP) ===
        ' Only apply Length Check here to prevent "Orphan Suction"
        If Not isValidMatch Then

            ' 1. Calculate Gap
            Dim gap As Double: gap = 0
            If t2 < 0 Then gap = 0 - t2          ' Wall before frame
            If t1 > frameLen Then gap = t1 - frameLen  ' Wall after frame

            ' 2. Check Gap Tolerance
            If gap <= 2000 Then    ' Max gap 2000mm

                ' 3. STRICT LENGTH CHECK (Only for Gap Match)
                ' Prevent 2m wall from snapping to 8m beam via gap
                Dim lenDiff As Double
                lenDiff = Abs(frameLen - wallSeg.Length)

                Dim lenRatio As Double: lenRatio = 0
                If frameLen > 0 Then lenRatio = wallSeg.Length / frameLen

                Dim isSimilar As Boolean: isSimilar = False

                ' Logic: Lengths must be close (+/- 1m) OR Ratio close to 1 (+/- 30%)
                If lenDiff <= 1000 Then
                    isSimilar = True
                ElseIf lenRatio >= 0.7 And lenRatio <= 1.3 Then
                    isSimilar = True
                End If

                If isSimilar Then
                    isValidMatch = True
                    ' Fix overlap for assignment
                    If overlapLen <= 0 Then
                        overlapLen = 100
                        If t2 < 0 Then
                            overlapStart = 0: overlapEnd = 100
                        Else
                            overlapStart = frameLen - 100: overlapEnd = frameLen
                        End If
                    End If
                Else
                    debugReason = "Gap OK but Length Mismatch (Gap Logic): Wall=" & Format(wallSeg.Length, "0") & " Frame=" & Format(frameLen, "0")
                End If
            Else
                debugReason = "Too far: Gap=" & Format(gap, "0") & " > 2000"
            End If
        End If

        If Not isValidMatch Then
            If debugReason = "" Then debugReason = "low_overlap overlapRatio=" & Format(overlapLen / wallSeg.Length, "0.000")
            Exit Function
        End If

        ' ---------------------------------------------------------
        ' STEP 5: RESULT
        ' ---------------------------------------------------------
        result.FrameIndex = frameIdx
        result.overlapStart = overlapStart / frameLen
        result.overlapEnd = overlapEnd / frameLen
        result.OverlapLength = overlapLen
        result.overlapRatio = overlapLen / wallSeg.Length
        result.Distance = distMid

        If result.overlapStart < 0 Then result.overlapStart = 0
        If result.overlapEnd > 1 Then result.overlapEnd = 1

        ' Scoring
        Dim distPenalty As Double
        distPenalty = distMid / dynamicDistTol
        If distPenalty > 1 Then distPenalty = 1

        result.Score = (result.overlapRatio * 0.7) + ((1 - distPenalty) * 0.3)

        If (result.overlapStart <= 0.05) And (result.overlapEnd >= 0.95) Then
            result.MatchType = "EXACT"
            result.Score = result.Score + 0.5
        Else
            result.MatchType = "PARTIAL"
        End If

        debugReason = "OK"
        AnalyzeWallFrameMatch = True
    End With
End Function
' ===============================================================
' HELPER: Convert matches to mapping records
' ===============================================================
Private Function ConvertToMappings(matches() As MatchResult, matchCount As Long, ByRef mappings() As MappingRecord) As Long
    ConvertToMappings = 0

    If matchCount = 0 Then Exit Function

    ReDim mappings(0 To matchCount - 1)

    Dim i           As Long
    For i = 0 To matchCount - 1
        With matches(i)
            mappings(i).TargetFrame = mFrameDB(.FrameIndex).frameName
            mappings(i).MapType = .MatchType
            mappings(i).FrameLength = mFrameDB(.FrameIndex).Length
            mappings(i).DistI = .overlapStart * mappings(i).FrameLength
            mappings(i).DistJ = .overlapEnd * mappings(i).FrameLength
        End With
    Next i

    ConvertToMappings = matchCount
End Function

' ===============================================================
' HELPER: Calculate total coverage
' ===============================================================
Private Function CalculateTotalCoverage(matches() As MatchResult, matchCount As Long, wallLength As Double) As Double
    CalculateTotalCoverage = 0

    If matchCount = 0 Or wallLength <= 0 Then Exit Function

    Dim totalLen    As Double
    totalLen = 0

    Dim i           As Long
    For i = 0 To matchCount - 1
        totalLen = totalLen + matches(i).OverlapLength
    Next i

    CalculateTotalCoverage = totalLen / wallLength
End Function
' ===============================================================
' HELPER: Sort matches by score (descending)
' ===============================================================
Private Sub SortMatchesByScore(ByRef matches() As MatchResult, count As Long)
    If count <= 1 Then Exit Sub

    Dim i As Long, j As Long
    For i = 0 To count - 2
        For j = i + 1 To count - 1
            If matches(j).Score > matches(i).Score Then
                Dim tmp As MatchResult
                tmp = matches(i)
                matches(i) = matches(j)
                matches(j) = tmp
            End If
        Next j
    Next i
End Sub

' ===============================================================
' CORE ALGORITHM: FindMatches_StrictGeometry (FIXED & TYPE MATCHED)
' Purpose: Find frames under the wall using vector projection
' ===============================================================
Function FindMatches_StrictGeometry(wall As WallSegmentMap, ByRef matches() As OverlapMatchResult, zLevel As Double) As Long
    FindMatches_StrictGeometry = 0

    ' 1. Validate frame database
    If mFrameCount = 0 Then
Debug.Print "WARNING: Frame database is empty"
        Exit Function
    End If

    Dim tempMatches() As OverlapMatchResult
    ReDim tempMatches(0 To 10)    ' Initial buffer
    Dim count       As Long
    count = 0

    Dim i           As Long
    Dim distTol     As Double
    distTol = 250    ' mm (Tube radius)

    Dim zTol        As Double
    zTol = 100    ' mm (Z-level tolerance)

    ' 2. Iterate through frames (Using correct variable mFrameDB)
    For i = 0 To mFrameCount - 1
        With mFrameDB(i)
            If Not .IsActive Then GoTo NextFrame

            ' Fast Z-Check
            If Abs(.Z - zLevel) > zTol Then GoTo NextFrame

            ' Fast Bounding Box Check (2D)
            Dim wMinX As Double, wMaxX As Double, wMinY As Double, wMaxY As Double
            wMinX = IIf(wall.startX < wall.endX, wall.startX, wall.endX)
            wMaxX = IIf(wall.startX > wall.endX, wall.startX, wall.endX)
            wMinY = IIf(wall.startY < wall.endY, wall.startY, wall.endY)
            wMaxY = IIf(wall.startY > wall.endY, wall.startY, wall.endY)

            ' Expand Frame BBox by tolerance
            If wMaxX < (.BBox.minX - distTol) Or wMinX > (.BBox.maxX + distTol) Then GoTo NextFrame
            If wMaxY < (.BBox.minY - distTol) Or wMinY > (.BBox.maxY + distTol) Then GoTo NextFrame

            ' Strict Geometry Check
            Dim result As OverlapMatchResult
            If CheckWallOnFrameIntersection(wall, i, result, distTol) Then

                ' Resize buffer if needed
                If count > UBound(tempMatches) Then
                    ReDim Preserve tempMatches(0 To count + 10)
                End If

                tempMatches(count) = result
                count = count + 1
            End If
        End With
NextFrame:
    Next i

    ' 3. Return results safely
    If count > 0 Then
        ReDim Preserve tempMatches(0 To count - 1)
        matches = tempMatches
        FindMatches_StrictGeometry = count
    Else
        Erase matches
        FindMatches_StrictGeometry = 0
    End If
End Function

' ===============================================================
' MATH: CheckWallOnFrameIntersection (FIXED VARIABES)
' ===============================================================
Private Function CheckWallOnFrameIntersection(wall As WallSegmentMap, fIdx As Long, ByRef res As OverlapMatchResult, distTol As Double) As Boolean
    CheckWallOnFrameIntersection = False

    ' Use mFrameDB instead of mFrames
    With mFrameDB(fIdx)

        ' Frame Geometry
        Dim fx1 As Double, fy1 As Double, fx2 As Double, fy2 As Double
        fx1 = .startPt.X
        fy1 = .startPt.Y
        fx2 = .endPt.X
        fy2 = .endPt.Y

        Dim fLen    As Double: fLen = .Length
        If fLen < 1 Then Exit Function

        ' 1. Check Parallelism
        Dim wx As Double, wy As Double
        wx = wall.endX - wall.startX
        wy = wall.endY - wall.startY

        Dim vfx As Double, vfy As Double
        vfx = fx2 - fx1
        vfy = fy2 - fy1

        Dim dotP    As Double
        dotP = Abs((wx * vfx + wy * vfy) / (wall.Length * fLen))
        If dotP < 0.98 Then Exit Function    ' < 11 degrees deviation

        ' 2. Check Perpendicular Distance
        Dim wMidX As Double, wMidY As Double
        wMidX = (wall.startX + wall.endX) / 2
        wMidY = (wall.startY + wall.endY) / 2

        Dim perpDist As Double
        perpDist = PointToLineDistanceInfinite(wMidX, wMidY, fx1, fy1, fx2, fy2)
        If perpDist > distTol Then Exit Function

        ' 3. PROJECT WALL ONTO FRAME
        ' Fix: Use .UnitVector.X instead of .UnitX
        Dim ux As Double, uy As Double
        ux = .UnitVector.X
        uy = .UnitVector.Y

        Dim tStart As Double, tEnd As Double
        tStart = (wall.startX - fx1) * ux + (wall.startY - fy1) * uy
        tEnd = (wall.endX - fx1) * ux + (wall.endY - fy1) * uy

        If tStart > tEnd Then
            Dim tmp As Double: tmp = tStart: tStart = tEnd: tEnd = tmp
        End If

        ' 4. Intersection with [0, fLen]
        Dim overlapStart As Double, overlapEnd As Double
        overlapStart = IIf(tStart > 0, tStart, 0)    ' Max function equivalent
        overlapEnd = IIf(tEnd < fLen, tEnd, fLen)    ' Min function equivalent

        Dim ovLen   As Double
        ovLen = overlapEnd - overlapStart

        ' 5. Validate and Populate Result
        If ovLen > 50 Then
            res.FrameIndex = fIdx
            res.DistI = overlapStart
            res.DistJ = overlapEnd
            res.MappedLength = ovLen
            res.overlapStart = overlapStart / fLen
            res.overlapEnd = overlapEnd / fLen
            res.overlapRatio = ovLen / wall.Length
            res.Distance = perpDist

            ' Simple score logic
            res.Score = res.overlapRatio

            If (overlapStart <= 50) And (overlapEnd >= fLen - 50) Then
                res.MatchType = "EXACT"
                res.Score = res.Score * 1.5    ' Boost exact match
            Else
                res.MatchType = "PARTIAL"
            End If

            CheckWallOnFrameIntersection = True
        End If

    End With
End Function

' ===============================================================
' DATABASE HELPERS
' ===============================================================
Private Function BuildFrameDatabase(SapModel As Object, mapData As Object) As Boolean
    On Error Resume Next

    BuildFrameDatabase = False
    mFrameCount = 0

    ' ? FIX 1: Validate inputs
    If mapData Is Nothing Then
Debug.Print "ERROR: mapData is Nothing"
        Exit Function
    End If

    If mapData.count = 0 Then
Debug.Print "ERROR: mapData is empty"
        Exit Function
    End If

    ReDim mFrames(0 To mapData.count - 1)

    Dim frameKey    As Variant
    Dim nodes       As Collection
    Dim pt1 As String, pt2 As String
    Dim x1 As Double, y1 As Double, z1 As Double
    Dim x2 As Double, y2 As Double, z2 As Double
    Dim ret1 As Long, ret2 As Long

    Dim idx         As Long
    idx = 0

    For Each frameKey In mapData.keys
        On Error Resume Next
        Set nodes = Nothing
        Set nodes = mapData(frameKey)

        If err.number <> 0 Then
Debug.Print "WARNING: Cannot get nodes for frame " & frameKey
            err.Clear
            GoTo NextFrameKey
        End If

        If nodes Is Nothing Then GoTo NextFrameKey
        If nodes.count < 2 Then GoTo NextFrameKey

        pt1 = CStr(nodes(1))
        pt2 = CStr(nodes(2))

        ret1 = SapModel.pointObj.GetCoordCartesian(pt1, x1, y1, z1)
        ret2 = SapModel.pointObj.GetCoordCartesian(pt2, x2, y2, z2)

        If err.number <> 0 Then
Debug.Print "WARNING: Cannot get coords for frame " & frameKey
            err.Clear
            GoTo NextFrameKey
        End If

        If ret1 = 0 And ret2 = 0 Then
            mFrames(idx).frameName = CStr(frameKey)
            mFrames(idx).startX = x1
            mFrames(idx).startY = y1
            mFrames(idx).endX = x2
            mFrames(idx).endY = y2
            mFrames(idx).Z = (z1 + z2) / 2

            Dim dx As Double, dy As Double
            dx = x2 - x1
            dy = y2 - y1
            mFrames(idx).Length = Sqr(dx * dx + dy * dy)

            If mFrames(idx).Length > 0.001 Then
                mFrames(idx).UnitX = dx / mFrames(idx).Length
                mFrames(idx).UnitY = dy / mFrames(idx).Length
                mFrames(idx).IsActive = True
                idx = idx + 1
            End If
        End If

NextFrameKey:
    Next frameKey

    mFrameCount = idx

    If mFrameCount > 0 Then
        ReDim Preserve mFrames(0 To mFrameCount - 1)
        BuildFrameDatabase = True
Debug.Print "Built frame database: " & mFrameCount & " frames"
    Else
Debug.Print "ERROR: No valid frames added to database"
    End If

    On Error GoTo 0
End Function

' ==================== UTILS ====================
Private Function PointToLineDistanceInfinite(px As Double, py As Double, x1 As Double, y1 As Double, x2 As Double, y2 As Double) As Double
    Dim dx As Double, dy As Double
    dx = x2 - x1
    dy = y2 - y1
    Dim l2          As Double: l2 = dx * dx + dy * dy
    If l2 = 0 Then
        PointToLineDistanceInfinite = Sqr((px - x1) ^ 2 + (py - y1) ^ 2)
    Else
        PointToLineDistanceInfinite = Abs(dy * (px - x1) - dx * (py - y1)) / Sqr(l2)
    End If
End Function

Private Function Min(a As Double, b As Double) As Double
    If a < b Then Min = a Else Min = b
End Function
Private Function Max(a As Double, b As Double) As Double
    If a > b Then Max = a Else Max = b
End Function

' ===============================================================
' INTERNAL HELPER: BuildGeometricFrameCache
' Purpose: Converts the API Node Map (Names) to Geometric Map (Coords)
'          Moves the logic from UserForm to Module n02
' ===============================================================
Private Function BuildGeometricFrameCache(SapModel As Object, mapData As Object) As Object
    Dim result      As Object
    Set result = CreateObject("Scripting.Dictionary")

    If SapModel Is Nothing Then
        Set BuildGeometricFrameCache = result
        Exit Function
    End If

    Dim frameName   As Variant
    Dim nodes       As Collection
    Dim pt1 As String, pt2 As String

    Dim x1 As Double, y1 As Double, z1 As Double
    Dim x2 As Double, y2 As Double, z2 As Double
    Dim ret1 As Long, ret2 As Long

    For Each frameName In mapData.keys
        Set nodes = mapData(frameName)
        If nodes.count >= 2 Then
            pt1 = CStr(nodes(1))
            pt2 = CStr(nodes(2))

            On Error Resume Next
            ret1 = SapModel.pointObj.GetCoordCartesian(pt1, x1, y1, z1)
            ret2 = SapModel.pointObj.GetCoordCartesian(pt2, x2, y2, z2)
            On Error GoTo 0

            If ret1 = 0 And ret2 = 0 Then
                Dim frameInfo As Object
                Set frameInfo = CreateObject("Scripting.Dictionary")
                frameInfo.Add "Name", CStr(frameName)
                frameInfo.Add "X1", x1
                frameInfo.Add "Y1", y1
                frameInfo.Add "X2", x2
                frameInfo.Add "Y2", y2
                ' Note: Z is strictly not needed for 2D logic but kept for ref
                result.Add CStr(frameName), frameInfo
            End If
        End If
    Next frameName

    Set BuildGeometricFrameCache = result
End Function

' ===============================================================
' UPDATED MAIN FUNCTION WITH LOGGING
' ===============================================================
' --- REPLACE the existing AssignWallLoadsToSAP Sub with this updated version ---
Public Sub AssignWallLoadsToSAP(acadDoc As Object, SapModel As Object, _
        storyInfo As Object, insertionPoint As Variant, _
        loadAssignments As Object, Optional suppressDialogs As Boolean = False, Optional selHandles As Variant)

    On Error GoTo ErrHandler

    ' Initialize logging for UI-initiated assignment
    Log_Init    ''
    Log_Write "AssignWallLoadsToSAP: START"

    If Not suppressDialogs Then
Debug.Print "========== Starting Wall Load Assignment =========="
    End If

    ' Validate inputs
    If acadDoc Is Nothing Or SapModel Is Nothing Then
        Log_Write "ERROR: AutoCAD or SAP2000 not connected when calling AssignWallLoadsToSAP"
        MsgBox "AutoCAD or SAP2000 not connected!", vbExclamation, "Error"
        Exit Sub
    End If

    If storyInfo Is Nothing Then
        Log_Write "ERROR: storyInfo missing in AssignWallLoadsToSAP"
        MsgBox "Story information not provided!", vbExclamation, "Error"
        Exit Sub
    End If

    If IsEmpty(insertionPoint) Or Not IsArray(insertionPoint) Then
        Log_Write "ERROR: Invalid insertionPoint in AssignWallLoadsToSAP"
        MsgBox "Invalid insertion point!", vbExclamation, "Error"
        Exit Sub
    End If

    ' Extract story data (expect keys "Name","Elevation","Height")
    Dim storyName   As String
    Dim storyElev   As Double
    Dim storyHeight As Double

    On Error Resume Next
    storyName = CStr(storyInfo("Name"))
    storyElev = CDbl(storyInfo("Elevation"))
    storyHeight = CDbl(storyInfo("Height"))
    On Error GoTo ErrHandler

    Log_Write "Story: " & storyName & ", Elevation: " & CStr(storyElev) & ", Height: " & CStr(storyHeight)

    ' DEBUG: Check selHandles before calling ReadWallLinesFromCAD
Debug.Print "=== BEFORE calling ReadWallLinesFromCAD ==="
Debug.Print "IsMissing(selHandles) = " & IsMissing(selHandles)

    If Not IsMissing(selHandles) Then
Debug.Print "IsEmpty(selHandles) = " & IsEmpty(selHandles)
Debug.Print "IsArray(selHandles) = " & IsArray(selHandles)
Debug.Print "TypeName(selHandles) = " & TypeName(selHandles)

        If IsArray(selHandles) Then
            On Error Resume Next
Debug.Print "LBound(selHandles) = " & LBound(selHandles)
Debug.Print "UBound(selHandles) = " & UBound(selHandles)
Debug.Print "Array element count = " & (UBound(selHandles) - LBound(selHandles) + 1)

            ' Print first few handles
            Dim debugIdx As Long
            For debugIdx = LBound(selHandles) To LBound(selHandles) + 2
                If debugIdx <= UBound(selHandles) Then
Debug.Print "  selHandles(" & debugIdx & ") = " & selHandles(debugIdx)
                End If
            Next debugIdx

            On Error GoTo ErrHandler
        End If
    End If
Debug.Print "=== END pre-check ==="

    ' Read wall lines from AutoCAD into Variant array of Dictionaries
    Dim wallLines   As Variant
    Dim wallCount   As Long

Debug.Print "About to call ReadWallLinesFromCAD..."

    ' IMPORTANT: Initialize wallLines before passing to function
    wallLines = Empty

    ' Call with selHandles
    wallCount = ReadWallLinesFromCAD(acadDoc, wallLines, selHandles)
    Log_Write "ReadWallLinesFromCAD returned: wallCount=" & CStr(wallCount)

Debug.Print "Returned from ReadWallLinesFromCAD, wallCount = " & wallCount

    If wallCount = 0 Then
        Log_Write "No wall lines found for story: " & storyName
        MsgBox "No wall lines found on layer '" & WALL_LAYER & "' or in selection!", vbInformation, "No Data"
        Exit Sub
    End If

    If Not suppressDialogs Then
Debug.Print "Found " & wallCount & " wall lines in AutoCAD (from selection or layer)"
    End If

    ' Apply coordinate transformation (insertion point offset)
    Dim offsetX As Double, offsetY As Double
    offsetX = CDbl(insertionPoint(0))
    offsetY = CDbl(insertionPoint(1))

    TransformWallCoordinates wallLines, wallCount, offsetX, offsetY, storyElev
    Log_Write "Transformed wall coordinates by offsetX=" & CStr(offsetX) & " offsetY=" & CStr(offsetY)

    ' Build SAP frame lookup at story elevation
    Dim sapFrames   As Object
    Set sapFrames = BuildSAPFrameLookup(SapModel, storyElev, storyHeight)
    Log_Write "Built SAP frame lookup. Found frames: " & CStr(IIf(sapFrames Is Nothing, 0, sapFrames.count))

    If Not suppressDialogs Then
Debug.Print "Found " & sapFrames.count & " SAP frames at story elevation"
    End If

    ' Process each wall line
    ' ==================== VECTOR STRIP BASED ALLOCATION ====================
    Dim stats       As Object
    Set stats = CreateObject("Scripting.Dictionary")
    stats("Assigned") = 0
    stats("Created") = 0
    stats("Failed") = 0

    ' Build SAP geometric cache (mFrameDB + strips)
    ResetFrameDatabase
    Call BuildFrameSpatialDatabase(SapModel, sapFrames, storyElev)
    Call BuildSapStripsVS
    Log_Write "Built SAP strip structures: mFrameCount=" & CStr(mFrameCount) & ", mSapStripCountVS=" & CStr(mSapStripCountVS)

    ' Build CAD strips from wallLines
    BuildCadStripsVS wallLines, wallCount
    Log_Write "Built CAD strips: mCadStripCountVS=" & CStr(mCadStripCountVS)

    ' Global strip matching + allocation
    Dim pairArr()   As StripMatchPair
    Dim pairCount   As Long
    pairCount = MatchCadStripsToSapStrips(pairArr)
    Log_Write "Matched CAD strips to SAP strips: pairCount=" & CStr(pairCount)

    Dim p           As Long
    For p = 0 To pairCount - 1
        ApplyCadStripToSapStrip pairArr(p), wallLines, SapModel, loadAssignments, _
                storyName, storyElev, storyHeight, stats, suppressDialogs
    Next p

    ' Any wall segments not assigned (leftovers) -> fall back to old per-wall logic (relaxed)
    Dim i           As Long
    For i = 0 To wallCount - 1
        Dim w       As Object
        Set w = wallLines(i)
        If Not w.exists("VS_Assigned") Or w("VS_Assigned") = False Then
            Log_Write "Fallback: processing unassigned wall handle: " & CStr(w("Handle"))
            ' Use old engine as fallback
            ProcessWallLine w, SapModel, sapFrames, loadAssignments, _
                    storyName, storyElev, storyHeight, stats, True   ' suppressDialogs = True for fallback
        End If
    Next i

    ' Refresh SAP view (best-effort)
    On Error Resume Next
    SapModel.View.RefreshView
    On Error GoTo ErrHandler

    ' Show summary
    If Not suppressDialogs Then
        Dim msg     As String
        msg = "Wall Load Assignment Complete!" & vbCrLf & vbCrLf & _
                "Story: " & storyName & vbCrLf & _
                "Assigned to existing frames: " & stats("Assigned") & vbCrLf & _
                "New frames created: " & stats("Created") & vbCrLf & _
                "Failed: " & stats("Failed")

        MsgBox msg, vbInformation, "Assignment Complete"
Debug.Print "========== Assignment Complete =========="
    End If

    Log_Write "Assignment Summary: Assigned=" & CStr(stats("Assigned")) & ", Created=" & CStr(stats("Created")) & ", Failed=" & CStr(stats("Failed"))

    ' Open notepad to view log
    Log_OpenNotepad

    Exit Sub

ErrHandler:
Debug.Print "ERROR in AssignWallLoadsToSAP: " & err.description & " (Number: " & err.number & ")"
    Log_Write "ERROR in AssignWallLoadsToSAP: " & err.description & " (Number: " & CStr(err.number) & ")"
    MsgBox "Error in AssignWallLoadsToSAP: " & err.description, vbCritical, "Error"
End Sub
' Read wall lines from AutoCAD layer or from provided handles
' UPDATED: STRICTLY FILTER FOR DTS_APP XDATA
Private Function ReadWallLinesFromCAD(acadDoc As Object, ByRef wallLines As Variant, Optional handles As Variant) As Long
    ' Temporarily DISABLE error handler for debugging
    ' On Error GoTo ErrHandler

Debug.Print "=== ReadWallLinesFromCAD START ==="

    ReadWallLinesFromCAD = 0

    ' Initialize wallLines as empty array
Debug.Print "Initializing wallLines array..."
    On Error Resume Next
    wallLines = Array()
    If err.number <> 0 Then
        err.Clear
        Dim tempArr() As Variant
        ReDim tempArr(0 To -1)
        wallLines = tempArr
    End If
    On Error GoTo 0

    If acadDoc Is Nothing Then
Debug.Print "ERROR: acadDoc is Nothing"
        Exit Function
    End If

    Dim coll        As Object
    Set coll = CreateObject("Scripting.Dictionary")

    Dim useHandles  As Boolean
    useHandles = False

    ' Check if handles parameter is provided and valid
    If Not IsMissing(handles) Then
        If Not IsEmpty(handles) Then
            If IsArray(handles) Then
                On Error Resume Next
                Dim testLB As Long, testUB As Long
                testLB = LBound(handles)
                testUB = UBound(handles)
                If err.number = 0 Then
                    If testUB >= testLB Then useHandles = True
                End If
                err.Clear
                On Error GoTo 0
            Else
                ' single handle
                If Len(Trim(CStr(handles))) > 0 Then useHandles = True
            End If
        End If
    End If

    Dim ent         As Object
    Dim totalSelected As Long: totalSelected = 0
    Dim processedCount As Long: processedCount = 0
    Dim skippedCount As Long: skippedCount = 0

    ' Flag to check if XData exists
    Dim hasXData    As Boolean

    If useHandles Then
        ' ===== CASE 1: Read from provided handles =====
Debug.Print "=== CASE 1: Reading from selection handles ==="

        Dim hArr()  As String
        If IsArray(handles) Then
            Dim origLB As Long, origUB As Long
            origLB = LBound(handles)
            origUB = UBound(handles)
            ReDim hArr(origLB To origUB)
            Dim copyIdx As Long
            For copyIdx = origLB To origUB
                hArr(copyIdx) = CStr(handles(copyIdx))
            Next copyIdx
        Else
            ReDim hArr(0 To 0)
            hArr(0) = CStr(handles)
        End If

        totalSelected = (UBound(hArr) - LBound(hArr) + 1)

        Dim iHandle As Long
        For iHandle = LBound(hArr) To UBound(hArr)
            hasXData = False    ' Reset flag

            On Error Resume Next
            Dim handleStr As String
            handleStr = Trim$(hArr(iHandle))

            Set ent = Nothing
            Set ent = acadDoc.HandleToObject(handleStr)

            If err.number <> 0 Or ent Is Nothing Then
                err.Clear
                skippedCount = skippedCount + 1
                GoTo NextHandleLoop
            End If
            On Error GoTo 0

            ' Check if it's a LINE entity
            Dim objType As String
            objType = LCase$(ent.ObjectName)
            If objType <> "acdbline" Then
                skippedCount = skippedCount + 1
                GoTo NextHandleLoop
            End If

            ' Get line geometry
            Dim ptStart As Variant, ptEnd As Variant
            On Error Resume Next
            ptStart = ent.StartPoint
            ptEnd = ent.EndPoint

            If err.number <> 0 Or Not IsArray(ptStart) Or Not IsArray(ptEnd) Then
                err.Clear
                skippedCount = skippedCount + 1
                GoTo NextHandleLoop
            End If
            On Error GoTo 0

            Dim x1 As Double, y1 As Double, z1 As Double
            Dim x2 As Double, y2 As Double, z2 As Double
            x1 = CDbl(ptStart(0)): y1 = CDbl(ptStart(1)): z1 = CDbl(ptStart(2))
            x2 = CDbl(ptEnd(0)): y2 = CDbl(ptEnd(1)): z2 = CDbl(ptEnd(2))

            Dim deltaX As Double, deltaY As Double
            deltaX = x2 - x1
            deltaY = y2 - y1
            Dim lineLen As Double
            lineLen = Sqr(deltaX * deltaX + deltaY * deltaY)

            ' Create dictionary for this line
            Dim dictInfo As Object
            Set dictInfo = CreateObject("Scripting.Dictionary")

            dictInfo.Add "Handle", handleStr
            dictInfo.Add "StartX", x1
            dictInfo.Add "StartY", y1
            dictInfo.Add "StartZ", z1
            dictInfo.Add "EndX", x2
            dictInfo.Add "EndY", y2
            dictInfo.Add "EndZ", z2
            dictInfo.Add "Length", lineLen
            ' Default values
            dictInfo.Add "thickness", 200#
            dictInfo.Add "wallType", ""
            dictInfo.Add "LoadPattern", ""
            dictInfo.Add "LoadValue", 0#

            ' Read XData - NEW FORMAT
            Dim xDataType As Variant, xDataValue As Variant
            On Error Resume Next
            ent.GetXData APP_NAME, xDataType, xDataValue

            If err.number = 0 Then
                If Not IsEmpty(xDataValue) Then
                    If IsArray(xDataValue) Then
                        ' STRICT CHECK: Mark as valid only if we have array data
                        hasXData = True

                        Dim xdUB As Long
                        xdUB = UBound(xDataValue)

                        ' Thickness
                        If xdUB >= 1 Then
                            If IsNumeric(xDataValue(1)) Then dictInfo("thickness") = CDbl(xDataValue(1))
                        End If
                        ' WallType
                        If xdUB >= 2 Then
                            If Not IsEmpty(xDataValue(2)) Then dictInfo("wallType") = CStr(xDataValue(2))
                        End If
                        ' LoadPattern
                        If xdUB >= 3 Then
                            If Not IsEmpty(xDataValue(3)) Then dictInfo("LoadPattern") = CStr(xDataValue(3))
                        End If
                        ' LoadValue
                        If xdUB >= 4 Then
                            If IsNumeric(xDataValue(4)) Then dictInfo("LoadValue") = CDbl(xDataValue(4))
                        End If
                    End If
                End If
            End If
            err.Clear
            On Error GoTo 0

            ' Generate wallType if not set
            If Len(Trim$(CStr(dictInfo("wallType")))) = 0 Then
                If CDbl(dictInfo("thickness")) > 0 Then
                    dictInfo("wallType") = "W" & CStr(CInt(CDbl(dictInfo("thickness"))))
                End If
            End If

            ' === CRITICAL CHANGE: Only add if XData was found ===
            If hasXData Then
                If Not coll.exists(handleStr) Then
                    coll.Add handleStr, dictInfo
                    processedCount = processedCount + 1
                End If
            Else
Debug.Print "Handle " & handleStr & " skipped (No DTS_APP XData)"
                skippedCount = skippedCount + 1
            End If

NextHandleLoop:
        Next iHandle

    Else
        ' ===== CASE 2: Scan ModelSpace by layer =====
Debug.Print "=== CASE 2: Scanning ModelSpace by layer ==="

        Dim ms      As Object
        Set ms = acadDoc.ModelSpace

        Dim entityCount As Long: entityCount = 0
         Dim ent As Object
        For Each ent In ms
            hasXData = False

            ' Use the SAME filter for all: must be valid wall with DTS_APP XData
            If Not IsValidWallWithXData(ent) Then
                GoTo NextEntityLoop
            End If

            ' At this point, layer and XData are OK; read geometry + header
            Dim pt1 As Variant, pt2 As Variant
            On Error Resume Next
            pt1 = ent.StartPoint
            pt2 = ent.EndPoint
            If err.number <> 0 Or Not IsArray(pt1) Or Not IsArray(pt2) Then
                err.Clear
                GoTo NextEntityLoop
            End If
            On Error GoTo ErrHandler

            Dim xStart As Double, yStart As Double, zStart As Double
            Dim xEnd As Double, yEnd As Double, zEnd As Double
            xStart = CDbl(pt1(0)): yStart = CDbl(pt1(1)): zStart = CDbl(pt1(2))
            xEnd = CDbl(pt2(0)): yEnd = CDbl(pt2(1)): zEnd = CDbl(pt2(2))

            Dim dx As Double, dy As Double, alen As Double
            dx = xEnd - xStart
            dy = yEnd - yStart
            alen = Sqr(dx * dx + dy * dy)

            Dim info As Object
            Set info = CreateObject("Scripting.Dictionary")
            Dim entHandle As String
            entHandle = CStr(ent.Handle)

            info.Add "Handle", entHandle
            info.Add "StartX", xStart
            info.Add "StartY", yStart
            info.Add "StartZ", zStart
            info.Add "EndX", xEnd
            info.Add "EndY", yEnd
            info.Add "EndZ", zEnd
            info.Add "Length", alen
            info.Add "thickness", 0#
            info.Add "wallType", ""
            info.Add "LoadPattern", ""
            info.Add "LoadValue", 0#

            ' Read XData header again (we know it exists from IsValidWallWithXData)
            Dim xType As Variant, xValue As Variant
            On Error Resume Next
            ent.GetXData APP_NAME, xType, xValue
            If err.number = 0 And IsArray(xValue) Then
                hasXData = True
                Dim ubIdx As Long
                ubIdx = UBound(xValue)

                If ubIdx >= XDATA_OFFSET_THICKNESS And IsNumeric(xValue(XDATA_OFFSET_THICKNESS)) Then
                    info("thickness") = CDbl(xValue(XDATA_OFFSET_THICKNESS))
                End If
                If ubIdx >= XDATA_OFFSET_WALLTYPE And Not IsEmpty(xValue(XDATA_OFFSET_WALLTYPE)) Then
                    info("wallType") = CStr(xValue(XDATA_OFFSET_WALLTYPE))
                End If
                If ubIdx >= XDATA_OFFSET_LOADPATTERN And Not IsEmpty(xValue(XDATA_OFFSET_LOADPATTERN)) Then
                    info("LoadPattern") = CStr(xValue(XDATA_OFFSET_LOADPATTERN))
                End If
                If ubIdx >= XDATA_OFFSET_LOADVALUE And IsNumeric(xValue(XDATA_OFFSET_LOADVALUE)) Then
                    info("LoadValue") = CDbl(xValue(XDATA_OFFSET_LOADVALUE))
                End If
            End If
            err.Clear
            On Error GoTo ErrHandler

            ' wallType fallback from thickness
            If Trim$(CStr(info("wallType"))) = "" Then
                If CDbl(info("thickness")) > 0 Then
                    info("wallType") = "W" & CStr(CInt(CDbl(info("thickness"))))
                End If
            End If

            If hasXData Then
                If Not coll.exists(entHandle) Then
                    coll.Add entHandle, info
                    processedCount = processedCount + 1
                End If
            End If

NextEntityLoop:
        Next ent

Debug.Print "Scanned " & entityCount & " entities in ModelSpace"
    End If

    ' Convert collection to array
    Dim totalCount  As Long
    totalCount = coll.count

Debug.Print "=== SUMMARY ==="
Debug.Print "totalSelected=" & totalSelected
Debug.Print "processed=" & processedCount
Debug.Print "skipped=" & skippedCount
Debug.Print "found=" & totalCount

    If totalCount = 0 Then
Debug.Print "=== ReadWallLinesFromCAD END (no results) ==="
        ReadWallLinesFromCAD = 0
        Exit Function
    End If

Debug.Print "Converting collection to array..."
    ReDim wallLines(0 To totalCount - 1)

    Dim arrIdx      As Long
    arrIdx = 0
    Dim key         As Variant
    For Each key In coll.keys
        Set wallLines(arrIdx) = coll(key)
        arrIdx = arrIdx + 1
    Next key

Debug.Print "Array conversion complete"
Debug.Print "=== ReadWallLinesFromCAD END (success) ==="
    ReadWallLinesFromCAD = totalCount
    Exit Function

ErrHandler:
Debug.Print "ReadWallLinesFromCAD ERROR: " & err.description & " (Error #" & err.number & ")"
    ReadWallLinesFromCAD = 0
End Function
' Transform coordinates based on insertion point
' wallLines is Variant array of Dictionary objects
Private Sub TransformWallCoordinates(ByRef wallLines As Variant, count As Long, _
        offsetX As Double, offsetY As Double, elevationZ As Double)

    Dim i           As Long
    For i = 0 To count - 1
        On Error Resume Next
        Dim w       As Object
        Set w = wallLines(i)
        If w Is Nothing Then GoTo NextT
        w("StartX") = CDbl(w("StartX")) - offsetX
        w("StartY") = CDbl(w("StartY")) - offsetY
        w("StartZ") = elevationZ

        w("EndX") = CDbl(w("EndX")) - offsetX
        w("EndY") = CDbl(w("EndY")) - offsetY
        w("EndZ") = elevationZ
NextT:
    Next i
End Sub

' --- REPLACE the existing ProcessWallLine Sub with this updated version ---
Private Sub ProcessWallLine(wall As Variant, SapModel As Object, sapFrames As Object, _
        loadAssignments As Object, storyName As String, storyElev As Double, storyHeight As Double, _
        stats As Object, suppressDialogs As Boolean)

    On Error Resume Next

    If IsEmpty(wall) Then
        stats("Failed") = stats("Failed") + 1
        Log_Write "ProcessWallLine: wall is empty -> failed"
        Exit Sub
    End If

    Dim dobj        As Object
    Set dobj = wall

    If dobj Is Nothing Then
        stats("Failed") = stats("Failed") + 1
        Log_Write "ProcessWallLine: dobj Is Nothing -> failed"
        Exit Sub
    End If

    ' ===== USE ADVANCED MAPPING ALGORITHM =====
    Dim matches     As Collection
    Set matches = MapWallToFrames(dobj, sapFrames)

    Log_Write "ProcessWallLine: Handle=" & CStr(dobj("Handle")) & " -> matches found=" & CStr(matches.count)

    If matches.count = 0 Then
Debug.Print "No matching frames found for wall, creating new frame..."
        Log_Write "No matching frames -> creating new frame for handle " & CStr(dobj("Handle"))

        ' Create new frame
        Dim frameName As String
        frameName = CreateNewFrame(SapModel, dobj, storyElev)
        Log_Write "CreateNewFrame returned: " & frameName

        If frameName <> "" Then
            stats("Created") = stats("Created") + 1

            ' Assign load to new frame
            If AssignLoadToFrame(dobj, SapModel, frameName, loadAssignments, _
                    storyName, storyHeight, 0, 1, stats, suppressDialogs) Then
                stats("Assigned") = stats("Assigned") + 1
                Log_Write "Assigned load to newly created frame " & frameName & " for handle " & CStr(dobj("Handle"))
            Else
                stats("Failed") = stats("Failed") + 1
                Log_Write "Failed to assign load to newly created frame " & frameName & " for handle " & CStr(dobj("Handle"))
            End If
        Else
            stats("Failed") = stats("Failed") + 1
            Log_Write "Failed to create new frame for handle " & CStr(dobj("Handle"))
        End If

        Exit Sub
    End If

    ' ===== PROCESS EACH MATCH =====
Debug.Print "Found " & matches.count & " matching frame(s)"
    Dim matchItem   As Variant
    Dim match       As Object
    For Each matchItem In matches
        Set match = matchItem

        ' Get frame name from index
        Dim matchFrameName As String
        matchFrameName = GetFrameName(match.FrameIndex)

        If Len(matchFrameName) > 0 Then
Debug.Print "  Assigning to frame: " & matchFrameName & _
        " (overlap=" & Format(match.overlapRatio * 100, "0.0") & "%" & _
        ", range=" & Format(match.overlapStart, "0.00") & " to " & Format(match.overlapEnd, "0.00") & ")"

            Log_Write "Attempt assign: Handle=" & CStr(dobj("Handle")) & " -> Frame=" & matchFrameName & _
                    " overlapRatio=" & Format(match.overlapRatio, "0.000") & " overlapStart=" & Format(match.overlapStart, "0.000") & " overlapEnd=" & Format(match.overlapEnd, "0.000")

            ' Assign load using calculated overlap range
            If AssignLoadToFrame(dobj, SapModel, matchFrameName, loadAssignments, _
                    storyName, storyHeight, match.overlapStart, match.overlapEnd, _
                    stats, suppressDialogs) Then
                stats("Assigned") = stats("Assigned") + 1
                Log_Write "Success assign: Handle=" & CStr(dobj("Handle")) & " -> Frame=" & matchFrameName
            Else
                stats("Failed") = stats("Failed") + 1
                Log_Write "Failed assign: Handle=" & CStr(dobj("Handle")) & " -> Frame=" & matchFrameName
            End If
        Else
            stats("Failed") = stats("Failed") + 1
            Log_Write "Match found but cannot resolve frame name for frame index " & CStr(match.FrameIndex)
        End If
    Next matchItem

    On Error GoTo 0
End Sub

' --- REPLACE the existing AssignLoadToFrame Function with this updated version ---
Private Function AssignLoadToFrame(dobj As Object, SapModel As Object, frameName As String, _
        loadAssignments As Object, storyName As String, storyHeight As Double, _
        overlapStart As Double, overlapEnd As Double, _
        stats As Object, suppressDialogs As Boolean) As Boolean

    On Error Resume Next
    AssignLoadToFrame = False

    ' Determine wall type
    Dim wallTypeStr As String
    wallTypeStr = ""

    If Len(Trim$(CStr(dobj("wallType")))) > 0 Then
        wallTypeStr = CStr(dobj("wallType"))
    Else
        If CDbl(dobj("thickness")) > 0 Then
            wallTypeStr = "W" & CStr(CInt(CDbl(dobj("thickness"))))
        End If
    End If

    If wallTypeStr = "" Then
Debug.Print "Cannot determine wall type"
        Log_Write "AssignLoadToFrame: Cannot determine wall type for handle " & CStr(dobj("Handle"))
        Exit Function
    End If

    ' Get load pattern & value from XData first
    Dim LoadPattern As String
    Dim loadPerMeter As Double
    LoadPattern = ""
    loadPerMeter = 0#

    On Error Resume Next
    If dobj.exists("LoadPattern") Then
        LoadPattern = CStr(dobj("LoadPattern"))
    End If
    If dobj.exists("LoadValue") Then
        loadPerMeter = CDbl(dobj("LoadValue"))  ' kN/m² from XData
    End If
    On Error GoTo 0

    ' Fallback to loadAssignments if no XData
    If loadPerMeter <= 0 Then
        Dim loadData As Object
        Set loadData = GetLoadAssignment(loadAssignments, wallTypeStr)

        If Not loadData Is Nothing Then
            LoadPattern = CStr(loadData("Pattern"))
            loadPerMeter = CDbl(loadData("Value"))  ' kN/m²
        End If
    End If

    If loadPerMeter <= 0 Then
Debug.Print "No load value found for wall type: " & wallTypeStr
        Log_Write "AssignLoadToFrame: No load value for wallType=" & wallTypeStr & " handle=" & CStr(dobj("Handle"))
        Exit Function
    End If

    ' ===== CONVERT LOAD: kN/m² ? kN/m ? kN/mm =====
    ' Step 1: kN/m² × height(m) = kN/m
    Dim heightInMeters As Double
    heightInMeters = storyHeight / 1000#

    Dim loadPerMeterLine As Double
    loadPerMeterLine = loadPerMeter * heightInMeters

    ' Step 2: kN/m ÷ 1000 = kN/mm (if model uses mm)
    Dim loadPerModelUnit As Double
    loadPerModelUnit = loadPerMeterLine / 1000#

    ' Log computed loads and parameters
    Log_Write "AssignLoadToFrame: Frame=" & frameName & " wallType=" & wallTypeStr & _
            " LoadPattern=" & LoadPattern & " load_kN_per_m2=" & Format(loadPerMeter, "0.000") & _
            " height_m=" & Format(heightInMeters, "0.000") & " => load_kN_per_m=" & Format(loadPerMeterLine, "0.000") & _
            " => modelUnit_kN_per_mm=" & Format(loadPerModelUnit, "0.000000") & _
            " overlap=[" & Format(overlapStart, "0.000") & "," & Format(overlapEnd, "0.000") & "]"

    ' Generate GUID
    Dim guid        As String
    guid = GenerateWallGUID(storyName, wallTypeStr, storyHeight, frameName)

    ' Assign distributed load to frame
    Dim ret         As Long
    ret = SapModel.frameObj.SetLoadDistributedWithGUID(frameName, LoadPattern, 1, 10, _
            overlapStart, overlapEnd, loadPerModelUnit, loadPerModelUnit, guid, "Global", True, False)

    Log_Write "SAP SetLoadDistributedWithGUID returned ret=" & CStr(ret) & " for Frame=" & frameName

    If ret = 0 Then
        UpdateFrameGUID SapModel, frameName, guid

        If Not suppressDialogs Then
Debug.Print "  ? Assigned " & Format(loadPerModelUnit * 1000, "0.00") & " kN/m to frame " & frameName & _
        " [" & Format(overlapStart, "0.00") & " to " & Format(overlapEnd, "0.00") & "]"
        End If

        AssignLoadToFrame = True
    Else
Debug.Print "  ? Failed to assign load (ret=" & ret & ")"
        Log_Write "AssignLoadToFrame: FAILED ret=" & CStr(ret) & " for Frame=" & frameName & " handle=" & CStr(dobj("Handle"))
    End If

    On Error GoTo 0
End Function

' ==================== HELPER: GET LOAD ASSIGNMENT ====================
Private Function GetLoadAssignment(loadAssignments As Object, WallType As String) As Object
    On Error Resume Next

    Set GetLoadAssignment = Nothing
    If loadAssignments Is Nothing Then Exit Function
    If Trim$(WallType) = "" Then Exit Function

    ' Direct lookup
    If loadAssignments.exists(WallType) Then
        Set GetLoadAssignment = loadAssignments(WallType)
        Exit Function
    End If

    ' Case-insensitive lookup
    Dim key         As Variant
    For Each key In loadAssignments.keys
        If StrComp(CStr(key), WallType, vbTextCompare) = 0 Then
            Set GetLoadAssignment = loadAssignments(key)
            Exit Function
        End If
    Next key
End Function

' ==================== HELPER: GENERATE GUID ====================
Private Function GenerateWallGUID(storyName As String, WallType As String, height As Double, frameName As String) As String
    On Error Resume Next

    Dim sStory As String, sWall As String, sFrame As String
    sStory = Trim$(CStr(storyName))
    sWall = Trim$(CStr(WallType))
    sFrame = Trim$(CStr(frameName))

    Dim raw         As String
    raw = "DTS_" & sStory & "_W_" & sWall & "_H_" & CStr(CInt(height)) & "_F_" & sFrame

    Dim cleaned     As String
    cleaned = CleanGUIDString(raw)

    If Len(cleaned) = 0 Then cleaned = "DTS" & Format(Now, "yyyymmddHHMMSS")

    If Len(cleaned) > 120 Then
        cleaned = Left(cleaned, 120)
    End If

    GenerateWallGUID = cleaned
End Function


' ==================== HELPER: UPDATE FRAME GUID ====================
Private Sub UpdateFrameGUID(SapModel As Object, frameName As String, guid As String)
    On Error Resume Next
    SapModel.frameObj.SetGUID frameName, guid
End Sub

' ==================== HELPER: CREATE NEW FRAME ====================
' --- REPLACE the existing CreateNewFrame Function with this updated version ---
Private Function CreateNewFrame(SapModel As Object, wall As Object, elevation As Double) As String
    On Error Resume Next

    CreateNewFrame = ""
    If wall Is Nothing Then
        Log_Write "CreateNewFrame: wall is Nothing"
        Exit Function
    End If

    Dim sx As Double, sy As Double, ex As Double, ey As Double
    sx = CDbl(wall("StartX")): sy = CDbl(wall("StartY"))
    ex = CDbl(wall("EndX")): ey = CDbl(wall("EndY"))

    Log_Write "CreateNewFrame: attempt create frame at elevation=" & CStr(elevation) & _
            " from (" & Format(sx, "0.00") & "," & Format(sy, "0.00") & ") to (" & Format(ex, "0.00") & "," & Format(ey, "0.00") & ")"

    Dim p1Name As String, p2Name As String
    p1Name = GetOrCreatePoint(SapModel, sx, sy, elevation)
    p2Name = GetOrCreatePoint(SapModel, ex, ey, elevation)

    If p1Name = "" Or p2Name = "" Then
        Log_Write "CreateNewFrame: failed to get/create points. p1=" & p1Name & " p2=" & p2Name
        Exit Function
    End If

    Dim frameName   As String
    Dim ret         As Long
    ret = SapModel.frameObj.AddByPoint(p1Name, p2Name, frameName, "None", "")

    If ret = 0 Then
        CreateNewFrame = frameName
        Log_Write "CreateNewFrame: success created frame " & frameName & " between " & p1Name & " and " & p2Name
    Else
        Log_Write "CreateNewFrame: SAP AddByPoint failed ret=" & CStr(ret)
    End If

    On Error GoTo 0
End Function

Private Function GetOrCreatePoint(SapModel As Object, X As Double, Y As Double, Z As Double) As String
    On Error Resume Next

    GetOrCreatePoint = ""
    If SapModel Is Nothing Then Exit Function

    Dim numPoints   As Long
    Dim pointNames() As String

    If SapModel.pointObj.GetNameList(numPoints, pointNames) = 0 Then
        Dim i       As Long
        For i = 0 To numPoints - 1
            Dim px As Double, py As Double, pz As Double
            SapModel.pointObj.GetCoordCartesian pointNames(i), px, py, pz

            Dim dist As Double
            dist = Sqr((X - px) ^ 2 + (Y - py) ^ 2 + (Z - pz) ^ 2)

            If dist < 50 Then  ' 50mm tolerance
                GetOrCreatePoint = pointNames(i)
                Exit Function
            End If
        Next i
    End If

    Dim newName     As String
    If SapModel.pointObj.AddCartesian(X, Y, Z, newName, "", "Global") = 0 Then
        GetOrCreatePoint = newName
    End If

    On Error GoTo 0
End Function

' Find matching SAP frame for wall line (wall is Dictionary)
Private Function FindMatchingSAPFrame(wall As Object, sapFrames As Object) As Object
    On Error Resume Next

    Set FindMatchingSAPFrame = Nothing
    If wall Is Nothing Then Exit Function
    If sapFrames Is Nothing Then Exit Function

    Dim frame       As Variant
    For Each frame In sapFrames.items
        Dim fx1 As Double, fy1 As Double, fx2 As Double, fy2 As Double
        fx1 = CDbl(frame("X1"))
        fy1 = CDbl(frame("Y1"))
        fx2 = CDbl(frame("X2"))
        fy2 = CDbl(frame("Y2"))

        ' Check if wall line overlaps with frame (with tolerance)
        Dim dist1 As Double, dist2 As Double
        dist1 = PointToLineDistance(CDbl(wall("StartX")), CDbl(wall("StartY")), fx1, fy1, fx2, fy2)
        dist2 = PointToLineDistance(CDbl(wall("EndX")), CDbl(wall("EndY")), fx1, fy1, fx2, fy2)

        If dist1 <= COORD_TOLERANCE And dist2 <= COORD_TOLERANCE Then
            Set FindMatchingSAPFrame = frame
            Exit Function
        End If
    Next frame

    On Error GoTo 0
End Function

' Calculate overlap between wall and frame (wall is Dictionary)
Private Sub CalculateOverlap(wall As Object, frame As Object, _
        ByRef overlapStart As Double, ByRef overlapEnd As Double)

    If wall Is Nothing Or frame Is Nothing Then
        overlapStart = 0
        overlapEnd = 1
        Exit Sub
    End If

    ' Project wall endpoints onto frame line
    Dim fx1 As Double, fy1 As Double, fx2 As Double, fy2 As Double
    fx1 = CDbl(frame("X1"))
    fy1 = CDbl(frame("Y1"))
    fx2 = CDbl(frame("X2"))
    fy2 = CDbl(frame("Y2"))

    Dim FrameLength As Double
    Dim dx As Double, dy As Double
    dx = fx2 - fx1
    dy = fy2 - fy1
    FrameLength = Sqr(dx * dx + dy * dy)

    If FrameLength < 0.001 Then
        overlapStart = 0
        overlapEnd = 1
        Exit Sub
    End If

    ' Project wall start point onto frame
    Dim t1 As Double, t2 As Double
    t1 = ((CDbl(wall("StartX")) - fx1) * dx + (CDbl(wall("StartY")) - fy1) * dy) / (FrameLength * FrameLength)
    t2 = ((CDbl(wall("EndX")) - fx1) * dx + (CDbl(wall("EndY")) - fy1) * dy) / (FrameLength * FrameLength)

    ' Clamp to [0, 1]
    If t1 < 0 Then t1 = 0
    If t1 > 1 Then t1 = 1
    If t2 < 0 Then t2 = 0
    If t2 > 1 Then t2 = 1

    overlapStart = IIf(t1 < t2, t1, t2)
    overlapEnd = IIf(t1 < t2, t2, t1)
End Sub



' Clean string for GUID (a-z, A-Z, 0-9 only)
Private Function CleanGUIDString(s As String) As String
    Dim result      As String
    Dim i           As Long

    For i = 1 To Len(s)
        Dim ch      As String
        ch = mid(s, i, 1)

        If (ch >= "a" And ch <= "z") Or (ch >= "A" And ch <= "Z") Or (ch >= "0" And ch <= "9") Then
            result = result & ch
        End If
    Next i

    CleanGUIDString = result
End Function

' Delete wall loads for story
Public Sub DeleteWallLoadsForStory(SapModel As Object, storyName As String, loadPatterns As Collection)
    On Error Resume Next

Debug.Print "========== Deleting Wall Loads for Story: " & storyName & " =========="

    Dim cleanStory  As String
    cleanStory = CleanGUIDString(storyName)
    If Len(cleanStory) = 0 Then
        MsgBox "Invalid story name for deletion.", vbExclamation, "Error"
        Exit Sub
    End If

    Dim guidPrefix  As String
    guidPrefix = LCase("DTS" & cleanStory)

    Dim numFrames   As Long
    Dim frameNames() As String

    If SapModel.frameObj.GetNameList(numFrames, frameNames) <> 0 Then Exit Sub

    Dim deleteCount As Long
    deleteCount = 0

    Dim i           As Long
    For i = 0 To numFrames - 1
        Dim guid    As String
        guid = ""
        If SapModel.frameObj.GetGUID(frameNames(i), guid) = 0 Then
            If Len(guid) > 0 Then
                If LCase(Left$(guid, Len(guidPrefix))) = guidPrefix Then
                    Dim pat As Variant
                    For Each pat In loadPatterns
                        On Error Resume Next
                        SapModel.frameObj.DeleteLoadDistributed frameNames(i), CStr(pat)
                        If err.number = 0 Then deleteCount = deleteCount + 1
                        err.Clear
                    Next pat
                End If
            End If
        End If
    Next i

    SapModel.View.RefreshView

    MsgBox "Deleted " & deleteCount & " wall load assignments for story: " & storyName, vbInformation, "Delete Complete"
Debug.Print "========== Delete Complete =========="
End Sub

' Helper functions
Private Function IsLineEntity(ent As Object) As Boolean
    On Error Resume Next
    IsLineEntity = (LCase(ent.ObjectName) = "acdbline")
    If err.number <> 0 Then IsLineEntity = False
End Function

' ===============================================================
' UTILITY: Point to line distance
' ===============================================================
Private Function PointToLineDistance(px As Double, py As Double, x1 As Double, y1 As Double, x2 As Double, y2 As Double) As Double
    Dim dx As Double, dy As Double, l2 As Double
    dx = x2 - x1
    dy = y2 - y1
    l2 = dx * dx + dy * dy

    If l2 = 0 Then
        PointToLineDistance = Sqr((px - x1) ^ 2 + (py - y1) ^ 2)
    Else
        PointToLineDistance = Abs(dy * (px - x1) - dx * (py - y1)) / Sqr(l2)
    End If
End Function

' ==================== MAIN MAPPING FUNCTION ====================
Public Function MapWallToFrames(wall As Object, sapFrames As Object) As Collection
    ' Returns Collection of OverlapMatch objects
    ' Each match represents a frame (or frame segment) that should receive load

    On Error GoTo ErrHandler

    Set MapWallToFrames = New Collection

    If wall Is Nothing Or sapFrames Is Nothing Then Exit Function

    ' Step 1: Build frame database (only once per story)
    If mFrameCount = 0 Then
        BuildFrameDatabase sapFrames
    End If

    If mFrameCount = 0 Then Exit Function

    ' Step 2: Extract wall info
    Dim wallSeg     As WallSegmentMap
    ExtractWallSegment wall, wallSeg

    ' Step 3: Find candidate frames using spatial search
    Dim candidates() As Long
    Dim candidateCount As Long
    candidateCount = FindCandidateFrames(wallSeg, candidates)

Debug.Print "Found " & candidateCount & " candidate frames for wall"

    If candidateCount = 0 Then Exit Function

    ' Step 4: Analyze each candidate and compute overlap
    Dim i           As Long
    For i = 0 To candidateCount - 1
        Dim match   As Object
        If AnalyzeFrameOverlap(wallSeg, candidates(i), match) Then
            ' Only add if overlap is significant
            If match.overlapRatio >= MIN_OVERLAP_RATIO Then
                MapWallToFrames.Add match
Debug.Print "  Match: Frame #" & candidates(i) & _
        ", overlap=" & Format(match.overlapRatio * 100, "0.0") & "%" & _
        ", score=" & Format(match.Score, "0.00")
            End If
        End If
    Next i

    ' Step 5: Sort matches by score (best first)
    If MapWallToFrames.count > 1 Then
        SortMatchesByScore MapWallToFrames
    End If

Debug.Print "Final matches: " & MapWallToFrames.count

    Exit Function

ErrHandler:
Debug.Print "ERROR in MapWallToFrames: " & err.description
End Function


Private Sub BuildIntervalTrees()
    ' Build two interval trees: one for horizontal projection, one for vertical

    ' Separate frames into near-horizontal and near-vertical
    Dim horizFrames() As Long, vertFrames() As Long
    Dim hCount As Long, vCount As Long
    hCount = 0: vCount = 0

    ReDim horizFrames(0 To mFrameCount - 1)
    ReDim vertFrames(0 To mFrameCount - 1)

    Dim i           As Long
    For i = 0 To mFrameCount - 1
        Dim absAngle As Double
        absAngle = Abs(mFrameSegments(i).angle)

        ' Horizontal if angle close to 0 or PI
        If absAngle < PI / 4 Or absAngle > 3 * PI / 4 Then
            horizFrames(hCount) = i
            hCount = hCount + 1
        Else
            vertFrames(vCount) = i
            vCount = vCount + 1
        End If
    Next i

    ' Build horizontal interval tree (based on X projection)
    If hCount > 0 Then
        ReDim mIntervalTreeHorizontal(0 To hCount * 2)
        mIntervalCountH = 0
        BuildIntervalTreeRecursive horizFrames, 0, hCount - 1, mIntervalTreeHorizontal, mIntervalCountH, True
    End If

    ' Build vertical interval tree (based on Y projection)
    If vCount > 0 Then
        ReDim mIntervalTreeVertical(0 To vCount * 2)
        mIntervalCountV = 0
        BuildIntervalTreeRecursive vertFrames, 0, vCount - 1, mIntervalTreeVertical, mIntervalCountV, False
    End If
End Sub

Private Sub BuildIntervalTreeRecursive(indices() As Long, Left As Long, Right As Long, _
        ByRef tree() As IntervalNode_Map, ByRef nodeCount As Long, _
        useX As Boolean)
    If Left > Right Then Exit Sub

    Dim mid         As Long
    mid = (Left + Right) \ 2

    Dim nodeIdx     As Long
    nodeIdx = nodeCount
    nodeCount = nodeCount + 1

    Dim frameIdx    As Long
    frameIdx = indices(mid)

    ' Set interval based on frame bounds
    If useX Then
        tree(nodeIdx).Low = mFrameSegments(frameIdx).BoundMinX
        tree(nodeIdx).High = mFrameSegments(frameIdx).BoundMaxX
    Else
        tree(nodeIdx).Low = mFrameSegments(frameIdx).BoundMinY
        tree(nodeIdx).High = mFrameSegments(frameIdx).BoundMaxY
    End If

    tree(nodeIdx).FrameIndex = frameIdx
    tree(nodeIdx).LeftChild = -1
    tree(nodeIdx).RightChild = -1

    ' Recursively build children
    If Left < mid Then
        tree(nodeIdx).LeftChild = nodeCount
        BuildIntervalTreeRecursive indices, Left, mid - 1, tree, nodeCount, useX
    End If

    If mid < Right Then
        tree(nodeIdx).RightChild = nodeCount
        BuildIntervalTreeRecursive indices, mid + 1, Right, tree, nodeCount, useX
    End If

    ' Update MaxHigh
    tree(nodeIdx).MaxHigh = tree(nodeIdx).High

    If tree(nodeIdx).LeftChild >= 0 Then
        If tree(tree(nodeIdx).LeftChild).MaxHigh > tree(nodeIdx).MaxHigh Then
            tree(nodeIdx).MaxHigh = tree(tree(nodeIdx).LeftChild).MaxHigh
        End If
    End If

    If tree(nodeIdx).RightChild >= 0 Then
        If tree(tree(nodeIdx).RightChild).MaxHigh > tree(nodeIdx).MaxHigh Then
            tree(nodeIdx).MaxHigh = tree(tree(nodeIdx).RightChild).MaxHigh
        End If
    End If
End Sub

' ==================== STEP 2: EXTRACT WALL INFO ====================
Private Sub ExtractWallSegment(wall As Object, ByRef seg As WallSegmentMap)
    On Error Resume Next

    seg.startX = CDbl(wall("StartX"))
    seg.startY = CDbl(wall("StartY"))
    seg.endX = CDbl(wall("EndX"))
    seg.endY = CDbl(wall("EndY"))

    Dim dx As Double, dy As Double
    dx = seg.endX - seg.startX
    dy = seg.endY - seg.startY

    seg.Length = Sqr(dx * dx + dy * dy)

    If seg.Length > 0.001 Then
        seg.UnitX = dx / seg.Length
        seg.UnitY = dy / seg.Length
        seg.angle = Atan2(dy, dx)
    Else
        seg.UnitX = 0
        seg.UnitY = 0
        seg.angle = 0
    End If

    seg.Thickness = CDbl(wall("thickness"))
    seg.WallType = CStr(wall("wallType"))
    seg.LoadPattern = CStr(wall("LoadPattern"))
    seg.LoadValue = CDbl(wall("LoadValue"))

    On Error GoTo 0
End Sub

' ==================== STEP 3: FIND CANDIDATE FRAMES ====================
Private Function FindCandidateFrames(wallSeg As WallSegmentMap, ByRef candidates() As Long) As Long
    On Error Resume Next

    FindCandidateFrames = 0
    ReDim candidates(0 To mFrameCount - 1)

    Dim tempCandidates() As Long
    Dim tempCount   As Long
    ReDim tempCandidates(0 To mFrameCount - 1)
    tempCount = 0

    ' Determine if wall is horizontal or vertical
    Dim absAngle    As Double
    absAngle = Abs(wallSeg.angle)

    Dim useHorizontal As Boolean
    useHorizontal = (absAngle < PI / 4 Or absAngle > 3 * PI / 4)

    ' Search appropriate interval tree
    Dim queryLow As Double, queryHigh As Double

    If useHorizontal Then
        queryLow = minVal(wallSeg.startX, wallSeg.endX) - COORD_TOLERANCE
        queryHigh = maxVal(wallSeg.startX, wallSeg.endX) + COORD_TOLERANCE

        If mIntervalCountH > 0 Then
            SearchIntervalTree mIntervalTreeHorizontal, 0, queryLow, queryHigh, tempCandidates, tempCount
        End If
    Else
        queryLow = minVal(wallSeg.startY, wallSeg.endY) - COORD_TOLERANCE
        queryHigh = maxVal(wallSeg.startY, wallSeg.endY) + COORD_TOLERANCE

        If mIntervalCountV > 0 Then
            SearchIntervalTree mIntervalTreeVertical, 0, queryLow, queryHigh, tempCandidates, tempCount
        End If
    End If

    ' Filter candidates by additional criteria
    Dim i           As Long
    For i = 0 To tempCount - 1
        Dim frameIdx As Long
        frameIdx = tempCandidates(i)

        ' Check angle similarity
        Dim angleDiff As Double
        angleDiff = Abs(wallSeg.angle - mFrameSegments(frameIdx).angle)
        If angleDiff > PI Then angleDiff = 2 * PI - angleDiff

        If angleDiff <= ANGLE_TOLERANCE Then
            ' Check bounding box overlap in other dimension
            Dim overlapOK As Boolean
            overlapOK = False

            If useHorizontal Then
                ' Check Y overlap
                Dim wallMinY As Double, wallMaxY As Double
                wallMinY = minVal(wallSeg.startY, wallSeg.endY) - COORD_TOLERANCE
                wallMaxY = maxVal(wallSeg.startY, wallSeg.endY) + COORD_TOLERANCE

                If Not (wallMaxY < mFrameSegments(frameIdx).BoundMinY Or wallMinY > mFrameSegments(frameIdx).BoundMaxY) Then
                    overlapOK = True
                End If
            Else
                ' Check X overlap
                Dim wallMinX As Double, wallMaxX As Double
                wallMinX = minVal(wallSeg.startX, wallSeg.endX) - COORD_TOLERANCE
                wallMaxX = maxVal(wallSeg.startX, wallSeg.endX) + COORD_TOLERANCE

                If Not (wallMaxX < mFrameSegments(frameIdx).BoundMinX Or wallMinX > mFrameSegments(frameIdx).BoundMaxX) Then
                    overlapOK = True
                End If
            End If

            If overlapOK Then
                candidates(FindCandidateFrames) = frameIdx
                FindCandidateFrames = FindCandidateFrames + 1
            End If
        End If
    Next i

    On Error GoTo 0
End Function

Private Sub SearchIntervalTree(tree() As IntervalNode_Map, nodeIdx As Long, _
        queryLow As Double, queryHigh As Double, _
        ByRef results() As Long, ByRef resultCount As Long)
    If nodeIdx < 0 Then Exit Sub

    ' Check if current interval overlaps query
    If Not (tree(nodeIdx).High < queryLow Or tree(nodeIdx).Low > queryHigh) Then
        results(resultCount) = tree(nodeIdx).FrameIndex
        resultCount = resultCount + 1
    End If

    ' Search left subtree if it might contain overlapping intervals
    If tree(nodeIdx).LeftChild >= 0 Then
        If tree(tree(nodeIdx).LeftChild).MaxHigh >= queryLow Then
            SearchIntervalTree tree, tree(nodeIdx).LeftChild, queryLow, queryHigh, results, resultCount
        End If
    End If

    ' Search right subtree
    If tree(nodeIdx).RightChild >= 0 Then
        SearchIntervalTree tree, tree(nodeIdx).RightChild, queryLow, queryHigh, results, resultCount
    End If
End Sub

' ==================== STEP 4: ANALYZE OVERLAP ====================
Private Function AnalyzeFrameOverlap(wallSeg As WallSegmentMap, frameIdx As Long, ByRef match As OverlapMatchResult) As Boolean
    On Error Resume Next

    AnalyzeFrameOverlap = False

    ' Calculate perpendicular distance
    Dim distStart As Double, distEnd As Double
    distStart = PointToSegmentDistance(wallSeg.startX, wallSeg.startY, frameIdx)
    distEnd = PointToSegmentDistance(wallSeg.endX, wallSeg.endY, frameIdx)

    match.Distance = (distStart + distEnd) / 2

    ' If too far, reject
    If match.Distance > COORD_TOLERANCE Then Exit Function

    ' Project wall onto frame axis
    Dim wallProj1 As Double, wallProj2 As Double
    wallProj1 = ProjectPointOnFrame(wallSeg.startX, wallSeg.startY, frameIdx)
    wallProj2 = ProjectPointOnFrame(wallSeg.endX, wallSeg.endY, frameIdx)

    ' Normalize so proj1 < proj2
    If wallProj1 > wallProj2 Then
        Dim temp    As Double
        temp = wallProj1
        wallProj1 = wallProj2
        wallProj2 = temp
    End If

    ' Frame projection is [0, frameLength]
    Dim frameLen    As Double
    frameLen = mFrameSegments(frameIdx).Length

    ' Calculate overlap
    Dim overlapStart As Double, overlapEnd As Double
    overlapStart = maxVal(wallProj1, 0)
    overlapEnd = minVal(wallProj2, frameLen)

    Dim OverlapLength As Double
    OverlapLength = overlapEnd - overlapStart

    If OverlapLength <= 0 Then Exit Function

    ' Calculate relative positions on frame [0, 1]
    match.overlapStart = overlapStart / frameLen
    match.overlapEnd = overlapEnd / frameLen

    ' Calculate overlap ratio (relative to wall length)
    Dim wallLen     As Double
    wallLen = wallProj2 - wallProj1

    If wallLen > 0.001 Then
        match.overlapRatio = OverlapLength / wallLen
    Else
        match.overlapRatio = 0
    End If

    ' Calculate score
    Dim normalizedDist As Double
    normalizedDist = match.Distance / COORD_TOLERANCE

    match.Score = match.overlapRatio * (1 - normalizedDist)
    match.FrameIndex = frameIdx

    AnalyzeFrameOverlap = True

    On Error GoTo 0
End Function


Private Function PointToSegmentDistance(px As Double, py As Double, frameIdx As Long) As Double
    Dim fx1 As Double, fy1 As Double, fx2 As Double, fy2 As Double
    fx1 = mFrameSegments(frameIdx).startX
    fy1 = mFrameSegments(frameIdx).startY
    fx2 = mFrameSegments(frameIdx).endX
    fy2 = mFrameSegments(frameIdx).endY

    Dim dx As Double, dy As Double, len2 As Double
    dx = fx2 - fx1
    dy = fy2 - fy1
    len2 = dx * dx + dy * dy

    If len2 < 0.000001 Then
        PointToSegmentDistance = Sqr((px - fx1) ^ 2 + (py - fy1) ^ 2)
    Else
        PointToSegmentDistance = Abs(dy * (px - fx1) - dx * (py - fy1)) / Sqr(len2)
    End If
End Function

Private Function ProjectPointOnFrame(px As Double, py As Double, frameIdx As Long) As Double
    ' Returns projection distance along frame axis (0 = start, length = end)

    Dim fx1 As Double, fy1 As Double
    fx1 = mFrameSegments(frameIdx).startX
    fy1 = mFrameSegments(frameIdx).startY

    Dim dx As Double, dy As Double
    dx = px - fx1
    dy = py - fy1

    ProjectPointOnFrame = dx * mFrameSegments(frameIdx).UnitX + dy * mFrameSegments(frameIdx).UnitY
End Function


' ==================== HELPER FUNCTIONS ====================
Private Function minVal(a As Double, b As Double) As Double
    If a < b Then minVal = a Else minVal = b
End Function

Private Function maxVal(a As Double, b As Double) As Double
    If a > b Then maxVal = a Else maxVal = b
End Function

' ===============================================================
' UTILITY: Atan2 function
' ===============================================================
Private Function Atan2(Y As Double, X As Double) As Double
    If X > 0 Then
        Atan2 = Atn(Y / X)
    ElseIf X < 0 And Y >= 0 Then
        Atan2 = Atn(Y / X) + PI
    ElseIf X < 0 And Y < 0 Then
        Atan2 = Atn(Y / X) - PI
    ElseIf X = 0 And Y > 0 Then
        Atan2 = PI / 2
    ElseIf X = 0 And Y < 0 Then
        Atan2 = -PI / 2
    Else
        Atan2 = 0
    End If
End Function

' ==================== GET FRAME NAME BY INDEX ====================
Public Function GetFrameName(frameIdx As Long) As String
    If frameIdx >= 0 And frameIdx < mFrameCount Then
        GetFrameName = mFrameSegments(frameIdx).frameName
    Else
        GetFrameName = ""
    End If
End Function

' ===============================================================
' PUBLIC: Reset database (call when switching stories)
' ===============================================================
Public Sub ResetFrameDatabase()
    mFrameCount = 0
    mIsInitialized = False
    Erase mFrameDB
    Set mSpatialHash = Nothing
Debug.Print "Frame database reset"
End Sub

' ===============================================================
' Wrapper XDATA: Write complete mapping data
' ===============================================================
Public Sub WriteWallCompleteXData(entObj As Object, wallSeg As WallSegmentMap, mappings() As MappingRecord, mappingCount As Long)
    ' Wrapper
    Dim dummyOv() As MappingRecord
    WriteWallCompleteXData_WithOverride entObj, wallSeg, mappings, mappingCount, False, wallSeg, dummyOv, 0
End Sub
'Wrapper ReadWallCompleteXData
Public Function ReadWallCompleteXData(entObj As Object, ByRef wallSeg As WallSegmentMap, ByRef mappings() As MappingRecord) As Long
    Dim baseMappings() As MappingRecord
    Dim ovWallSeg As WallSegmentMap
    Dim ovMappings() As MappingRecord
    Dim hasOverride As Boolean
    Dim ovCount As Long

    ReadWallCompleteXData = ReadWallAllXData(entObj, wallSeg, baseMappings, ovWallSeg, ovMappings, hasOverride, ovCount)

    If ReadWallCompleteXData > 0 Then
        mappings = baseMappings
    End If
End Function

' ==================== LABEL GENERATION FUNCTIONS (THICKNESS DRIVEN) ====================
Public Function GenerateCompositeLabel(wallSeg As WallSegmentMap, mappings() As MappingRecord, mappingCount As Long) As String
    On Error Resume Next

    ' --- LOGIC CHANGE: Always derive display name from Thickness ---
    Dim DisplayName As String
    If wallSeg.Thickness > 0 Then
        DisplayName = "W" & CInt(wallSeg.Thickness)
    Else
        DisplayName = "W?"
    End If
    ' -------------------------------------------------------------

    Dim LoadPattern As String
    LoadPattern = wallSeg.LoadPattern
    If LoadPattern = "" Then LoadPattern = "DL"

    Dim LoadValue As Double
    LoadValue = wallSeg.LoadValue
    If LoadValue < 0 Then LoadValue = 0

    Dim baseLabel As String
    baseLabel = DisplayName & " " & LoadPattern & "=" & Format(LoadValue, "0.00") & " kN/m2"

    Dim mappingPart As String
    If mappingCount <= 0 Then
        mappingPart = " (No Mapping)"
    ElseIf mappingCount = 1 Then
        mappingPart = " " & GetMappingDisplayLabel(mappings(0))
    Else
        mappingPart = " " & GroupMappingsLabel(mappings, mappingCount)
    End If

    GenerateCompositeLabel = baseLabel & mappingPart
End Function

Private Function GetMappingDisplayLabel(rec As MappingRecord) As String
    ' Returns: "to B120 (full 2.5m)" or "to B125 I=0.8to2.1" or "to New"
    If rec.MapType = "NEW" Then
        GetMappingDisplayLabel = "to New"
        Exit Function
    End If

    If rec.MapType = "EXACT" Then
        GetMappingDisplayLabel = "to " & rec.TargetFrame & " (full " & Format(rec.FrameLength / 1000, "0.0") & "m)"
        Exit Function
    End If

    ' PARTIAL
    If rec.FrameLength > 0 Then
        Dim relStart As Double, relEnd As Double
        relStart = rec.DistI / 1000
        relEnd = rec.DistJ / 1000
        GetMappingDisplayLabel = "to " & rec.TargetFrame & " I=" & Format(relStart, "0.0") & "to" & Format(relEnd, "0.0")
    Else
        GetMappingDisplayLabel = "to " & rec.TargetFrame & " I=0to0"
    End If
End Function

Private Function GroupMappingsLabel(mappings() As MappingRecord, count As Long) As String
    On Error Resume Next

    GroupMappingsLabel = "to "

    Dim frameNames As String
    Dim TotalLength As Double
    frameNames = ""
    TotalLength = 0

    Dim i As Long
    For i = 0 To count - 1
        If mappings(i).MapType = "EXACT" Then
            If Len(frameNames) > 0 Then frameNames = frameNames & ","
            frameNames = frameNames & mappings(i).TargetFrame
            TotalLength = TotalLength + mappings(i).FrameLength
        ElseIf mappings(i).MapType = "PARTIAL" Then
            If Len(frameNames) > 0 Then
                GroupMappingsLabel = GroupMappingsLabel & frameNames & " (full " & Format(TotalLength / 1000, "0.0") & "m), "
                frameNames = ""
                TotalLength = 0
            End If
            Dim relStart As Double, relEnd As Double
            relStart = mappings(i).DistI / 1000
            relEnd = mappings(i).DistJ / 1000
            GroupMappingsLabel = GroupMappingsLabel & mappings(i).TargetFrame & " I=" & Format(relStart, "0.0") & "to" & Format(relEnd, "0.0") & ", "
        Else
            If Len(frameNames) > 0 Then
                GroupMappingsLabel = GroupMappingsLabel & frameNames & " (full " & Format(TotalLength / 1000, "0.0") & "m), "
                frameNames = ""
                TotalLength = 0
            End If
            GroupMappingsLabel = GroupMappingsLabel & "New, "
        End If
    Next i

    If Len(frameNames) > 0 Then
        GroupMappingsLabel = GroupMappingsLabel & frameNames & " (full " & Format(TotalLength / 1000, "0.0") & "m)"
    End If

    If Right$(GroupMappingsLabel, 2) = ", " Then
        GroupMappingsLabel = Left$(GroupMappingsLabel, Len(GroupMappingsLabel) - 2)
    End If
End Function

' ==================== PARSE LABEL FUNCTION ====================
Public Function ParseMappingLabel(labelText As String, ByRef mappings() As MappingRecord) As Long
    ' Returns: number of mapping records parsed
    ' Input formats (after load part "W200 DL=..."):
    '   "to New"
    '   "to 179 (full 8.0m)"
    '   "to 190"
    '   "to 19,18,17 I=0.5to9.0"
    '   "to B125,B126 (full 5.0m)"
    '   "to B125,B126 I=0.8to2.1"

    On Error Resume Next

    ParseMappingLabel = 0
    Erase mappings

    If Len(Trim$(labelText)) = 0 Then Exit Function

    ' Find " to " after load part
    Dim toPos As Long
    toPos = InStr(1, labelText, " to ", vbTextCompare)
    If toPos = 0 Then Exit Function

    Dim mappingPart As String
    mappingPart = Trim$(mid$(labelText, toPos + 4))

    ' Case: "New"
    If LCase$(mappingPart) = "new" Then
        ReDim mappings(0 To 0)
        mappings(0).TargetFrame = "New"
        mappings(0).MapType = "NEW"
        mappings(0).DistI = 0#
        mappings(0).DistJ = 0#
        mappings(0).FrameLength = 0#
        ParseMappingLabel = 1
        Exit Function
    End If

    ' Split by first "(" to separate names and range text, if any
    Dim openParen As Long
    openParen = InStr(1, mappingPart, "(")

    Dim frameNamesPart As String
    Dim rangeText As String

    If openParen > 0 Then
        frameNamesPart = Trim$(Left$(mappingPart, openParen - 1))
        rangeText = Trim$(mid$(mappingPart, openParen + 1))
        If Right$(rangeText, 1) = ")" Then rangeText = Left$(rangeText, Len(rangeText) - 1)
    Else
        ' No parentheses: maybe "... I=0.5to9" or just "179"
        Dim spacePos As Long
        spacePos = InStr(1, mappingPart, " ")
        If spacePos > 0 Then
            frameNamesPart = Trim$(Left$(mappingPart, spacePos - 1))
            rangeText = Trim$(mid$(mappingPart, spacePos + 1))
        Else
            frameNamesPart = Trim$(mappingPart)
            rangeText = ""
        End If
    End If

    If Len(frameNamesPart) = 0 Then Exit Function

    ' Names can be "179" or "19,18,17"
    Dim frameArr() As String
    frameArr = Split(frameNamesPart, ",")

    Dim isExact As Boolean
    Dim isPartial As Boolean
    isExact = (InStr(1, rangeText, "full", vbTextCompare) > 0)
    isPartial = (InStr(1, rangeText, "I=", vbTextCompare) > 0)

    Dim i As Long
    Dim n As Long
    n = UBound(frameArr) + 1

    ReDim mappings(0 To n - 1)

    If isExact Or isPartial Then
        For i = 0 To n - 1
            mappings(i).TargetFrame = Trim$(frameArr(i))
            If isExact Then
                mappings(i).MapType = "EXACT"
                Dim lengthStr As String
                lengthStr = Replace(rangeText, "full", "", , , vbTextCompare)
                lengthStr = Replace(lengthStr, "m", "", , , vbTextCompare)
                lengthStr = Trim$(lengthStr)
                If IsNumeric(lengthStr) Then
                    mappings(i).FrameLength = CDbl(lengthStr) * 1000#
                    mappings(i).DistI = 0#
                    mappings(i).DistJ = mappings(i).FrameLength
                Else
                    mappings(i).FrameLength = 0#
                    mappings(i).DistI = 0#
                    mappings(i).DistJ = 0#
                End If
            ElseIf isPartial Then
                mappings(i).MapType = "PARTIAL"
                Dim rangeOnly As String
                rangeOnly = rangeText
                If InStr(1, rangeOnly, "I=", vbTextCompare) > 0 Then
                    rangeOnly = Replace(rangeOnly, "I=", "", , , vbTextCompare)
                End If
                Dim toPos2 As Long
                toPos2 = InStr(1, rangeOnly, "to", vbTextCompare)
                Dim startStr As String, endStr As String
                If toPos2 > 0 Then
                    startStr = Trim$(Left$(rangeOnly, toPos2 - 1))
                    endStr = Trim$(mid$(rangeOnly, toPos2 + 2))
                Else
                    startStr = "0"
                    endStr = "0"
                End If

                If IsNumeric(startStr) Then mappings(i).DistI = CDbl(startStr) * 1000#
                If IsNumeric(endStr) Then mappings(i).DistJ = CDbl(endStr) * 1000#
                mappings(i).FrameLength = mappings(i).DistJ - mappings(i).DistI
            End If
        Next i

        ParseMappingLabel = n
        Exit Function
    End If

    '   "to 190"  => EXACT full
    '   "to 19,18,17" => full, DistI=0, DistJ=0
    For i = 0 To n - 1
        mappings(i).TargetFrame = Trim$(frameArr(i))
        mappings(i).MapType = "EXACT"
        mappings(i).DistI = 0#
        mappings(i).DistJ = 0#
        mappings(i).FrameLength = 0#
    Next i

    ParseMappingLabel = n
End Function
Private Function NormalizeSpaces(ByVal txt As String) As String
    ' Replace multiple spaces with single space
    Do While InStr(txt, "  ") > 0
        txt = Replace(txt, "  ", " ")
    Loop
    NormalizeSpaces = Trim(txt)
End Function

' ==================== MAIN MAPPING FUNCTION (UPDATED) ====================
Public Function MapWallToFramesWithLabel(wall As Object, sapFrames As Object) As String
    ' Returns composite mapping label for display
    ' Updates wall XData with mapping records

    On Error GoTo ErrHandler

    MapWallToFramesWithLabel = ""

    If wall Is Nothing Or sapFrames Is Nothing Then Exit Function

    ' Build frame database if not exists
    If mFrameCount = 0 Then
        BuildFrameDatabase sapFrames
    End If

    If mFrameCount = 0 Then
        MapWallToFramesWithLabel = "to New (no frames)"
        Exit Function
    End If

    ' Extract wall segment
    Dim wallSeg     As WallSegmentMap
    ExtractWallSegment wall, wallSeg

    ' Find matching frames
    Dim matches()   As OverlapMatchResult
    Dim matchCount  As Long
    matchCount = FindAndClassifyMatches(wallSeg, matches)

    ' Convert matches to MappingRecord array
    Dim mappings()  As MappingRecord
    Dim mappingCount As Long
    mappingCount = ConvertMatchesToMappings(matches, matchCount, mappings)

    ' Generate composite label
    Dim compositeLabel As String
    compositeLabel = GenerateCompositeLabel(wallSeg, mappings, mappingCount)

    ' Update wall XData
    Dim acadDoc     As Object
    Set acadDoc = GetActiveACADDocument()
    If Not acadDoc Is Nothing Then
        Dim entObj  As Object
        Set entObj = acadDoc.HandleToObject(CStr(wall("Handle")))
        If Not entObj Is Nothing Then
            WriteWallCompleteXData entObj, wallSeg, mappings, mappingCount
        End If
    End If

    MapWallToFramesWithLabel = compositeLabel

    Exit Function

ErrHandler:
Debug.Print "ERROR in MapWallToFramesWithLabel: " & err.description
    MapWallToFramesWithLabel = "ERROR"
End Function

' Convert OverlapMatchResult array to MappingRecord array
Private Function ConvertMatchesToMappings(matches() As OverlapMatchResult, matchCount As Long, ByRef mappings() As MappingRecord) As Long
    ConvertMatchesToMappings = 0

    If matchCount = 0 Then Exit Function

    ReDim mappings(0 To matchCount - 1)

    Dim i           As Long
    For i = 0 To matchCount - 1
        mappings(i).MapType = matches(i).MatchType

        If matches(i).MatchType = "NEW" Then
            mappings(i).TargetFrame = "New"
            mappings(i).DistI = 0
            mappings(i).DistJ = 0
            mappings(i).FrameLength = 0
        Else
            ' Get frame info
            mappings(i).TargetFrame = GetFrameName(matches(i).FrameIndex)
            mappings(i).FrameLength = mFrameSegments(matches(i).FrameIndex).Length

            ' Calculate absolute distances from I-end
            mappings(i).DistI = matches(i).overlapStart * mappings(i).FrameLength
            mappings(i).DistJ = matches(i).overlapEnd * mappings(i).FrameLength
        End If
    Next i

    ConvertMatchesToMappings = matchCount

End Function

' Find and classify matches (returns array)
Private Function FindAndClassifyMatches(wallSeg As WallSegmentMap, ByRef matches() As OverlapMatchResult) As Long
    FindAndClassifyMatches = 0

    ' Find candidate frames
    Dim candidates() As Long
    Dim candidateCount As Long
    candidateCount = FindCandidateFrames(wallSeg, candidates)

    If candidateCount = 0 Then
        ' No candidates - create new
        ReDim matches(0 To 0)
        matches(0).MatchType = "NEW"
        matches(0).Score = 0
        FindAndClassifyMatches = 1
        Exit Function
    End If

    ' Analyze each candidate
    ReDim matches(0 To candidateCount - 1)
    Dim matchCount  As Long
    matchCount = 0

    Dim i           As Long
    For i = 0 To candidateCount - 1
        Dim match   As OverlapMatchResult
        If AnalyzeFrameOverlap(wallSeg, candidates(i), match) Then
            ' Classify match
            If match.overlapRatio >= 0.95 Then
                match.MatchType = "EXACT"
                matches(matchCount) = match
                matchCount = matchCount + 1
            ElseIf match.overlapRatio >= MIN_OVERLAP_RATIO Then
                match.MatchType = "PARTIAL"
                matches(matchCount) = match
                matchCount = matchCount + 1
            End If
        End If
    Next i

    If matchCount = 0 Then
        ' No valid matches - create new
        ReDim matches(0 To 0)
        matches(0).MatchType = "NEW"
        matches(0).Score = 0
        FindAndClassifyMatches = 1
    Else
        If matchCount < candidateCount Then
            ReDim Preserve matches(0 To matchCount - 1)
        End If
        FindAndClassifyMatches = matchCount
    End If

End Function
Private Function GetActiveACADDocument() As Object
    On Error Resume Next
    Dim acadApp     As Object
    Set acadApp = GetObject(, "AutoCAD.Application")
    If Not acadApp Is Nothing Then
        Set GetActiveACADDocument = acadApp.ActiveDocument
    End If
End Function

' ==================== BUILD CAD STRIPS FROM wallLines (Variant of Dictionary) ====================
Private Sub BuildCadStripsVS(ByRef wallLines As Variant, ByVal wallCount As Long)
    On Error Resume Next

    mCadStripCountVS = 0
    If wallCount <= 0 Then Exit Sub

    Dim i           As Long
    ReDim mCadStripsVS(0 To wallCount - 1)

    ' Temporary: store assigned strip id per wall
    Dim stripIDPerWall() As Long
    ReDim stripIDPerWall(0 To wallCount - 1)
    For i = 0 To wallCount - 1
        stripIDPerWall(i) = -1
    Next i

    Dim w           As Object
    For i = 0 To wallCount - 1
        Set w = wallLines(i)
        If w Is Nothing Then GoTo NextWallVS
        If CDbl(w("Length")) <= 1 Then GoTo NextWallVS

        If stripIDPerWall(i) >= 0 Then GoTo NextWallVS

        Dim newStripID As Long
        newStripID = mCadStripCountVS

        Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
        x1 = CDbl(w("StartX")): y1 = CDbl(w("StartY"))
        x2 = CDbl(w("EndX")): y2 = CDbl(w("EndY"))

        Dim dx As Double, dy As Double, l As Double
        dx = x2 - x1: dy = y2 - y1
        l = Sqr(dx * dx + dy * dy)
        If l <= 0.001 Then GoTo NextWallVS

        With mCadStripsVS(newStripID)
            .StripID = newStripID
            .angle = Atan2(dy, dx)
            .UnitX = dx / l
            .UnitY = dy / l
            .OriginX = x1
            .OriginY = y1
            .TotalLength = l

            ReDim .segments(0 To wallCount - 1)
            .segmentCount = 0

            ' Add first segment
            .segments(0).WallIndex = i
            .segments(0).StartGlobal = 0
            .segments(0).EndGlobal = l
            .segmentCount = 1

            ' Use load from first wall as base
            .TotalLength = l
        End With
        stripIDPerWall(i) = newStripID

        ' Try to search other collinear walls
        Dim j       As Long
        For j = i + 1 To wallCount - 1
            If stripIDPerWall(j) >= 0 Then GoTo NextWallCandidate

            Dim w2  As Object
            Set w2 = wallLines(j)
            If w2 Is Nothing Then GoTo NextWallCandidate

            Dim x3 As Double, y3 As Double, x4 As Double, y4 As Double
            x3 = CDbl(w2("StartX")): y3 = CDbl(w2("StartY"))
            x4 = CDbl(w2("EndX")): y4 = CDbl(w2("EndY"))

            Dim dx2 As Double, dy2 As Double, l2 As Double
            dx2 = x4 - x3: dy2 = y4 - y3
            l2 = Sqr(dx2 * dx2 + dy2 * dy2)
            If l2 <= 1 Then GoTo NextWallCandidate

            ' Check angle
            Dim ang2 As Double, angDiff As Double
            ang2 = Atan2(dy2, dx2)
            angDiff = Abs(ang2 - mCadStripsVS(newStripID).angle)
            If angDiff > PI Then angDiff = 2 * PI - angDiff
            If angDiff > ANGLE_TOLERANCE Then GoTo NextWallCandidate

            ' Check perpendicular distance to strip
            Dim midX As Double, midY As Double, dist As Double
            midX = (x3 + x4) / 2
            midY = (y3 + y4) / 2
            dist = PointToLineDistanceInfinite(midX, midY, _
                    mCadStripsVS(newStripID).OriginX, mCadStripsVS(newStripID).OriginY, _
                    mCadStripsVS(newStripID).OriginX + mCadStripsVS(newStripID).UnitX * 1000#, _
                    mCadStripsVS(newStripID).OriginY + mCadStripsVS(newStripID).UnitY * 1000#)
            If dist > COORD_TOLERANCE Then GoTo NextWallCandidate

            ' Project x3,y3 and x4,y4 onto strip
            Dim s1 As Double, s2 As Double
            s1 = (x3 - mCadStripsVS(newStripID).OriginX) * mCadStripsVS(newStripID).UnitX + _
                    (y3 - mCadStripsVS(newStripID).OriginY) * mCadStripsVS(newStripID).UnitY
            s2 = (x4 - mCadStripsVS(newStripID).OriginX) * mCadStripsVS(newStripID).UnitX + _
                    (y4 - mCadStripsVS(newStripID).OriginY) * mCadStripsVS(newStripID).UnitY
            If s1 > s2 Then
                Dim tmp As Double
                tmp = s1: s1 = s2: s2 = tmp
            End If

            Dim idxSeg As Long
            idxSeg = mCadStripsVS(newStripID).segmentCount
            mCadStripsVS(newStripID).segments(idxSeg).WallIndex = j
            mCadStripsVS(newStripID).segments(idxSeg).StartGlobal = s1
            mCadStripsVS(newStripID).segments(idxSeg).EndGlobal = s2
            mCadStripsVS(newStripID).segmentCount = mCadStripsVS(newStripID).segmentCount + 1

            ' Update total length approximation (not critical)
            mCadStripsVS(newStripID).TotalLength = _
                    maxVal(mCadStripsVS(newStripID).TotalLength, s2)

            stripIDPerWall(j) = newStripID

NextWallCandidate:
        Next j

        ' Trim segment array
        With mCadStripsVS(newStripID)
            If .segmentCount > 0 Then
                ReDim Preserve .segments(0 To .segmentCount - 1)
            End If
        End With

        mCadStripCountVS = mCadStripCountVS + 1

NextWallVS:
    Next i

    If mCadStripCountVS > 0 Then
        ReDim Preserve mCadStripsVS(0 To mCadStripCountVS - 1)
    End If
End Sub

' ==================== MATCH CAD STRIPS TO SAP STRIPS ====================
Private Function MatchCadStripsToSapStrips(ByRef pairs() As StripMatchPair) As Long
    On Error Resume Next

    MatchCadStripsToSapStrips = 0
    If mCadStripCountVS = 0 Or mSapStripCountVS = 0 Then Exit Function

    ReDim pairs(0 To mCadStripCountVS * mSapStripCountVS - 1)
    Dim count       As Long
    count = 0

    Dim i As Long, j As Long
    For i = 0 To mCadStripCountVS - 1
        For j = 0 To mSapStripCountVS - 1
            Dim pr  As StripMatchPair
            If EvaluateStripPairVS(i, j, pr) Then
                pairs(count) = pr
                count = count + 1
            End If
        Next j
    Next i

    If count = 0 Then Exit Function

    ReDim Preserve pairs(0 To count - 1)
    SortStripMatchPairs pairs, count

    MatchCadStripsToSapStrips = count
End Function

Private Function EvaluateStripPairVS(cadID As Long, sapID As Long, ByRef pr As StripMatchPair) As Boolean
    On Error Resume Next
    EvaluateStripPairVS = False

    ' Detailed logging for debugging specific cases can be enabled here
    ' LogTrace "Checking CAD Strip " & cadID & " vs SAP Strip " & sapID

    ' Angle check
    Dim angDiff     As Double
    angDiff = Abs(mCadStripsVS(cadID).angle - mSapStripsVS(sapID).angle)
    If angDiff > PI Then angDiff = 2 * PI - angDiff

    If angDiff > ANGLE_TOLERANCE Then
        ' LogTrace "  -> Rejected: Angle mismatch (" & Format(angDiff * 180 / PI, "0.0") & " deg)"
        Exit Function
    End If

    ' Distance check
    Dim midCad      As Double
    midCad = (mCadStripsVS(cadID).TotalLength) / 2

    Dim midX As Double, midY As Double
    midX = mCadStripsVS(cadID).OriginX + midCad * mCadStripsVS(cadID).UnitX
    midY = mCadStripsVS(cadID).OriginY + midCad * mCadStripsVS(cadID).UnitY

    Dim dist        As Double
    dist = PointToLineDistanceInfinite(midX, midY, _
            mSapStripsVS(sapID).OriginX, mSapStripsVS(sapID).OriginY, _
            mSapStripsVS(sapID).OriginX + mSapStripsVS(sapID).UnitX * 1000#, _
            mSapStripsVS(sapID).OriginY + mSapStripsVS(sapID).UnitY * 1000#)

    If dist > COORD_TOLERANCE Then
        ' LogTrace "  -> Rejected: Too far (" & Format(dist, "0.0") & " mm)"
        Exit Function
    End If

    ' Overlap check (1D)
    Dim overlapRatio As Double
    overlapRatio = Strip1DOverlapRatio(cadID, sapID)

    If overlapRatio < MIN_OVERLAP_RATIO Then
        ' LogTrace "  -> Rejected: Low overlap (" & Format(overlapRatio * 100, "0.0") & "%)"
        Exit Function
    End If

    pr.CadStripID = cadID
    pr.SapStripID = sapID
    pr.overlapRatio = overlapRatio
    pr.DistanceScore = 1 - (dist / COORD_TOLERANCE)
    pr.Score = pr.overlapRatio * 0.6 + pr.DistanceScore * 0.4

    EvaluateStripPairVS = True
    ' LogTrace "  -> MATCHED! Score: " & Format(pr.Score, "0.00")
End Function

Private Function Strip1DOverlapRatio(cadID As Long, sapID As Long) As Double
    On Error Resume Next

    ' CAD global range
    Dim cadMin As Double, cadMax As Double
    cadMin = 0
    cadMax = mCadStripsVS(cadID).TotalLength

    ' SAP range: sum frame lengths
    Dim totalSap    As Double
    totalSap = mSapStripsVS(sapID).TotalLength

    Dim ovStart As Double, ovEnd As Double
    ovStart = maxVal(0, 0)    ' both start at 0
    ovEnd = minVal(cadMax, totalSap)

    Dim ovLen       As Double
    ovLen = ovEnd - ovStart
    If ovLen <= 0 Then
        Strip1DOverlapRatio = 0
    Else
        Strip1DOverlapRatio = ovLen / cadMax
    End If
End Function

Private Sub SortStripMatchPairs(ByRef arr() As StripMatchPair, ByVal count As Long)
    If count <= 1 Then Exit Sub

    Dim i As Long, j As Long
    For i = 0 To count - 2
        For j = i + 1 To count - 1
            If arr(j).Score > arr(i).Score Then
                Dim tmp As StripMatchPair
                tmp = arr(i)
                arr(i) = arr(j)
                arr(j) = tmp
            End If
        Next j
    Next i
End Sub

' ===============================================================
' VECTOR STRIP: Apply CAD Strip Load to SAP Strip Frames
' Updated: Includes Slot Availability Check & Logging
' ===============================================================
Private Sub ApplyCadStripToSapStrip(pr As StripMatchPair, ByRef wallLines As Variant, _
        SapModel As Object, loadAssignments As Object, _
        storyName As String, storyElev As Double, storyHeight As Double, _
        stats As Object, suppressDialogs As Boolean)

    On Error Resume Next

    Dim cadID As Long, sapID As Long
    cadID = pr.CadStripID
    sapID = pr.SapStripID

    ' Log the start of processing for this matched pair
    Log_Write "VS_Process: Processing Match CAD_Strip=" & cadID & " -> SAP_Strip=" & sapID & " (Score=" & Format(pr.Score, "0.00") & ")"

    ' Loop through all segments (walls) in the CAD Strip
    Dim iSeg        As Long
    For iSeg = 0 To mCadStripsVS(cadID).segmentCount - 1
        Dim wIdx    As Long
        wIdx = mCadStripsVS(cadID).segments(iSeg).WallIndex

        ' Retrieve wall object dictionary
        Dim w       As Object
        Set w = wallLines(wIdx)
        If w Is Nothing Then GoTo NextSegVS

        ' Get global start/end of this wall segment on the strip axis
        Dim segStart As Double, segEnd As Double
        segStart = mCadStripsVS(cadID).segments(iSeg).StartGlobal
        segEnd = mCadStripsVS(cadID).segments(iSeg).EndGlobal

        Dim cumLen  As Double
        cumLen = 0

        Dim assignedThisSeg As Boolean
        assignedThisSeg = False

        ' Loop through all frames in the SAP Strip to find overlap
        Dim iFa     As Long
        For iFa = 0 To mSapStripsVS(sapID).FrameCount - 1
            Dim fIdx As Long
            fIdx = mSapStripsVS(sapID).FrameIndices(iFa)

            ' Calculate Frame's global range on the strip
            Dim fStart As Double, fEnd As Double
            fStart = cumLen
            fEnd = cumLen + mFrameDB(fIdx).Length

            ' Calculate intersection between Wall Segment and Frame
            Dim ovStart As Double, ovEnd As Double
            ovStart = maxVal(segStart, fStart)
            ovEnd = minVal(segEnd, fEnd)

            ' Check if there is a valid overlap
            If ovEnd > ovStart Then
                ' Convert to Local Coordinates relative to the Frame Start
                Dim localS As Double, localE As Double
                localS = ovStart - fStart
                localE = ovEnd - fStart

                Dim frameName As String
                frameName = mFrameDB(fIdx).frameName

                ' -------------------------------------------------------
                ' SLOT CHECK LOGIC (CRITICAL STEP)
                ' -------------------------------------------------------
                ' Check if this specific interval on the frame is free (not yet assigned)
                If IsIntervalAvailableVS(fIdx, localS, localE) Then

                    ' --- CASE: SLOT AVAILABLE ---
                    ' Attempt to assign load via API
                    If AssignLoadToFrame_VS(w, SapModel, fIdx, loadAssignments, _
                            storyName, storyHeight, localS, localE, stats, suppressDialogs) Then

                        ' Success: Mark this interval as occupied to prevent double counting
                        MarkIntervalOccupiedVS fIdx, localS, localE

                        ' Mark wall as assigned
                        w("VS_Assigned") = True
                        stats("Assigned") = stats("Assigned") + 1
                        assignedThisSeg = True

                        Log_Write "  -> OK: Assigned to " & frameName & " Local=[" & Format(localS, "0") & "-" & Format(localE, "0") & "]"
                    Else
                        ' API Error
                        Log_Write "  -> ERR: API Failed for " & frameName
                        stats("Failed") = stats("Failed") + 1
                    End If

                Else
                    ' --- CASE: SLOT OCCUPIED ---
                    ' The interval is already taken by another wall (or part of it).
                    ' We SKIP to avoid double loading the same location.
                    Log_Write "  -> SKIP: Slot Occupied on " & frameName & " Local=[" & Format(localS, "0") & "-" & Format(localE, "0") & "]"
                End If
            End If

            ' Advance cumulative length for the next frame in strip
            cumLen = fEnd
        Next iFa

        ' Log warning if wall was not assigned to any frame
        If Not assignedThisSeg Then
            Log_Write "  -> WARN: Wall Handle " & CStr(w("Handle")) & " matched strip but found no valid/free frames."
        End If

NextSegVS:
    Next iSeg
End Sub

Private Function IsIntervalAvailableVS(frameIdx As Long, sLoc As Double, eLoc As Double) As Boolean
    IsIntervalAvailableVS = True

    Dim i           As Long
    With mSapFrameStatus(frameIdx)
        For i = 0 To .OccupiedCount - 1
            If Not (eLoc <= .Occupied(i).Start Or sLoc >= .Occupied(i).[End]) Then
                IsIntervalAvailableVS = False
                Exit Function
            End If
        Next i
    End With
End Function

Private Sub MarkIntervalOccupiedVS(frameIdx As Long, sLoc As Double, eLoc As Double)
    With mSapFrameStatus(frameIdx)
        If .OccupiedCount >= UBound(.Occupied) Then
            ReDim Preserve .Occupied(0 To .OccupiedCount + 10)
        End If
        .Occupied(.OccupiedCount).Start = sLoc
        .Occupied(.OccupiedCount).[End] = eLoc
        .OccupiedCount = .OccupiedCount + 1
    End With
End Sub

' Assign load using existing AssignLoadToFrame logic but with local coordinates
Private Function AssignLoadToFrame_VS(dobj As Object, SapModel As Object, frameIdx As Long, _
        loadAssignments As Object, storyName As String, storyHeight As Double, _
        localStart As Double, localEnd As Double, _
        stats As Object, suppressDialogs As Boolean) As Boolean

    On Error Resume Next
    AssignLoadToFrame_VS = False

    Dim frameName   As String
    frameName = mFrameDB(frameIdx).frameName
    If Len(frameName) = 0 Then Exit Function

    Dim frameLen    As Double
    frameLen = mFrameDB(frameIdx).Length
    If frameLen <= 0.001 Then Exit Function

    Dim relStart As Double, relEnd As Double
    relStart = localStart / frameLen
    relEnd = localEnd / frameLen

    ' Reuse AssignLoadToFrame by building fake overlapStart/End
    AssignLoadToFrame_VS = AssignLoadToFrame(dobj, SapModel, frameName, _
            loadAssignments, storyName, storyHeight, _
            relStart, relEnd, stats, suppressDialogs)
End Function
' ===============================================================
' UNIFIED VALIDATION: Check if entity is a valid wall
' COMPATIBLE WITH USERFORM LOGIC (Relaxed Requirements)
' ===============================================================
Private Function IsValidWallWithXData(ent As Object) As Boolean
    On Error GoTo ErrHandler
    IsValidWallWithXData = False

    If ent Is Nothing Then
        Debug.Print "DEBUG: ent Is Nothing"
        Exit Function
    End If

    ' 1. Layer check
    Dim layName As String
    layName = LCase$(Trim$(ent.layer))
    Debug.Print "DEBUG: Layer = '" & layName & "' (Expected: '" & LCase$(WALL_LAYER) & "')"
    
    If layName <> LCase$(WALL_LAYER) Then
        Debug.Print "DEBUG: REJECTED - Wrong layer!"
        Exit Function
    End If

    ' 2. Object type check
    Dim objType As String
    objType = LCase$(ent.ObjectName)
    Debug.Print "DEBUG: ObjectType = '" & objType & "'"
    
    If objType <> "acdbline" Then
        Debug.Print "DEBUG: REJECTED - Not a line!"
        Exit Function
    End If

    ' 3. XData check
    Dim xdType As Variant, xdVal As Variant
    On Error Resume Next
    ent.GetXData APP_NAME, xdType, xdVal
    
    If err.number <> 0 Then
        Debug.Print "DEBUG: REJECTED - GetXData error: " & err.description
        err.Clear
        Exit Function
    End If
    On Error GoTo ErrHandler

    If IsEmpty(xdVal) Then
        Debug.Print "DEBUG: REJECTED - XData IsEmpty"
        Exit Function
    End If
    
    If Not IsArray(xdVal) Then
        Debug.Print "DEBUG: REJECTED - XData Not IsArray"
        Exit Function
    End If

    Dim maxIdx As Long
    maxIdx = UBound(xdVal)
    Debug.Print "DEBUG: XData array size = " & (maxIdx + 1) & " (Need >= 2 for Thickness)"

    If maxIdx < XDATA_OFFSET_THICKNESS Then
        Debug.Print "DEBUG: REJECTED - XData too small (has " & (maxIdx + 1) & ", need >= 2)"
        Exit Function
    End If

    Debug.Print "DEBUG: ? ACCEPTED!"
    IsValidWallWithXData = True
    Exit Function

ErrHandler:
    Debug.Print "DEBUG: EXCEPTION - " & err.description
    IsValidWallWithXData = False
End Function

' ===============================================================
' UNIFIED EXTRACTION: Extract wall data with safe defaults
' COMPATIBLE WITH USERFORM LOGIC
' ===============================================================
Private Function ExtractWallFromEntity(ent As Object, ByRef wOut As WallSegmentMap, offX As Double, offY As Double) As Boolean
    On Error GoTo ErrHandler
    ExtractWallFromEntity = False

    If ent Is Nothing Then Exit Function

    ' 1. Validate using unified function
    If Not IsValidWallWithXData(ent) Then Exit Function

    ' 2. GET GEOMETRY
    Dim sp As Variant, ep As Variant
    On Error Resume Next
    sp = ent.StartPoint
    ep = ent.EndPoint
    If err.number <> 0 Then
        err.Clear
        Exit Function
    End If
    On Error GoTo ErrHandler

    If Not IsArray(sp) Or Not IsArray(ep) Then Exit Function

    wOut.startX = CDbl(sp(0)) - offX
    wOut.startY = CDbl(sp(1)) - offY
    wOut.endX = CDbl(ep(0)) - offX
    wOut.endY = CDbl(ep(1)) - offY

    Dim dx As Double, dy As Double
    dx = wOut.endX - wOut.startX
    dy = wOut.endY - wOut.startY
    wOut.Length = Sqr(dx * dx + dy * dy)

    If wOut.Length < 1# Then Exit Function

    wOut.UnitX = dx / wOut.Length
    wOut.UnitY = dy / wOut.Length
    wOut.angle = Atan2(dy, dx)

    ' 3. READ XDATA HEADER WITH SAFE DEFAULTS
    Dim xdType As Variant, xdVal As Variant
    On Error Resume Next
    ent.GetXData APP_NAME, xdType, xdVal
    On Error GoTo ErrHandler

    Dim maxIdx As Long
    maxIdx = UBound(xdVal)

    ' Initialize defaults (matching UserForm CreateWallDictFromEntity)
    wOut.Thickness = 200#
    wOut.WallType = ""
    wOut.LoadPattern = "DL"
    wOut.LoadValue = 0#

    ' Read available fields safely
    If maxIdx >= XDATA_OFFSET_THICKNESS Then
        If IsNumeric(xdVal(XDATA_OFFSET_THICKNESS)) Then
            wOut.Thickness = CDbl(xdVal(XDATA_OFFSET_THICKNESS))
        End If
    End If

    If maxIdx >= XDATA_OFFSET_WALLTYPE Then
        If Not IsEmpty(xdVal(XDATA_OFFSET_WALLTYPE)) Then
            wOut.WallType = CStr(xdVal(XDATA_OFFSET_WALLTYPE))
        End If
    End If

    If maxIdx >= XDATA_OFFSET_LOADPATTERN Then
        If Not IsEmpty(xdVal(XDATA_OFFSET_LOADPATTERN)) Then
            wOut.LoadPattern = CStr(xdVal(XDATA_OFFSET_LOADPATTERN))
        End If
    End If

    If maxIdx >= XDATA_OFFSET_LOADVALUE Then
        If IsNumeric(xdVal(XDATA_OFFSET_LOADVALUE)) Then
            wOut.LoadValue = CDbl(xdVal(XDATA_OFFSET_LOADVALUE))
        End If
    End If

    ' Derive default WallType if empty but thickness exists
    If wOut.WallType = "" And wOut.Thickness > 0 Then
        wOut.WallType = "W" & CStr(CInt(wOut.Thickness))
    End If

    ExtractWallFromEntity = True
    Exit Function

ErrHandler:
    ExtractWallFromEntity = False
End Function
' ===============================================================
' LOGGING UTILITIES (NEW ADDITION)
' Purpose: Write detailed execution trace to a text file
' ===============================================================


Private Sub InitLog()
    mLogFilePath = Environ("TEMP") & "\DTS_Wall_Mapping_Log.txt"
    Dim fNum As Integer
    fNum = FreeFile
    Open mLogFilePath For Output As #fNum
    Print #fNum, "========== DTS WALL MAPPING LOG START: " & Now & " =========="
    Print #fNum, "User: " & Environ("USERNAME")
    Print #fNum, "==========================================================="
    Close #fNum
End Sub

Private Sub LogTrace(msg As String)
    Dim fNum As Integer
    fNum = FreeFile
    On Error Resume Next
    Open mLogFilePath For Append As #fNum
    Print #fNum, "[" & Format(Now, "HH:mm:ss") & "] " & msg
    Close #fNum
    Debug.Print msg
End Sub

Private Sub ShowLog()
    On Error Resume Next
    'Shell "notepad.exe """ & mLogFilePath & """", vbNormalFocus
End Sub

Private Sub Log_Init(Optional logPath As String)
    On Error Resume Next
    If Len(Trim$(logPath)) > 0 Then
        gLogFilePath = logPath
    Else
        gLogFilePath = Environ("TEMP") & "\DTS_Mapping_Log.txt"
    End If
    gLogEnabled = True
    Dim fn As Integer: fn = FreeFile
    Open gLogFilePath For Append As #fn
    Print #fn, "===================="
    Print #fn, "DTS Mapping Log Started: " & Format(Now, "yyyy-mm-dd HH:nn:ss")
    Print #fn, "===================="
    Close #fn
End Sub

' Write log message with timestamp
Private Sub Log_Write(msg As String)
    On Error Resume Next
    If Not gLogEnabled Then Exit Sub
    If Len(Trim$(gLogFilePath)) = 0 Then
        gLogFilePath = Environ("TEMP") & "\DTS_Mapping_Log.txt"
    End If
    Dim fn As Integer: fn = FreeFile
    Open gLogFilePath For Append As #fn
    Print #fn, Format(Now, "yyyy-mm-dd HH:nn:ss") & " - " & msg
    Close #fn
End Sub

Private Sub Log_OpenNotepad()
    On Error Resume Next
    If Len(Trim$(gLogFilePath)) = 0 Then Exit Sub
    'Shell "notepad.exe " & Chr(34) & gLogFilePath & Chr(34), vbNormalFocus
End Sub
