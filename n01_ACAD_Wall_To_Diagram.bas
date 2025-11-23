Attribute VB_Name = "n01_ACAD_Wall_To_Diagram"
Option Explicit
'===============================================================
' Module: n01_ACAD_Wall_To_Diagram
' Purpose: Advanced Wall-to-Centerline Converter with Enhanced Single Line Support
' Author: DTS System (by thanhtdvncc)
' Version: 3.1 - Enhanced with Single Line Processing & Exhaustive Recovery
'===============================================================

'==================================================================
' PUBLIC SETTINGS VARIABLES
'==================================================================
Public mThicknessInput As String
Public mLayerFilter As String
Public mAngleTol    As Double
Public mExtendMult  As Double
Public mDistTol     As Double
Public mEnableExtend As Boolean
Public mEnableIntersect As Boolean
Public mBreakAtGridIntersection As Boolean

Public mSettingsReady As Boolean
Private mEnableSingleLineProcessing As Boolean

Private mAxes()     As AxisLine
Private mAxesCount  As Long
Private mDoorWidths() As Double
Private mDoorWidthCount As Long
Private mColumnWidths() As Double
Private mColumnWidthCount As Long
Private mAxisSnapDistance As Double
Private mAutoJoinGapDistance As Double
Private mExtendLinesToGridIntersections As Boolean


' ==================== TYPE DEFINITIONS ====================
Private Type Point2D
    X               As Double
    Y               As Double
End Type

Private Type EndpointNode
    lineID          As Long
    IsStart         As Boolean      ' True = StartPt, False = EndPt
    X               As Double
    Y               As Double
    IsDangling      As Boolean      ' True = free end, False = connected to another line
    ConnectedLineID As Long
End Type

Private Type GridSegment
    StartPoint      As Point2D      ' Intersection with perpendicular axis 1
    EndPoint        As Point2D      ' Intersection with perpendicular axis 2
    StartProj       As Double       ' Projection value at start
    endProj         As Double       ' Projection value at end
End Type


Private Type AxisLine
    startPt         As Point2D
    endPt           As Point2D
    Length          As Double
    angle           As Double
    isHorizontal    As Boolean
    IsVertical      As Boolean
End Type

Private Type AxisCluster
    Coordinate      As Double   ' The representative coordinate (X or Y)
    TotalLength     As Double   ' Sum of lengths (weight)
    lineIndices()   As Long     ' List of centerline IDs in this cluster
    count           As Long
End Type

Private Type WallSegment
    Handle          As String
    startPt         As Point2D
    endPt           As Point2D
    Length          As Double
    angle           As Double
    layer           As String
    IsProcessed     As Boolean
    GroupID         As Long
    PairSegmentID   As Long
    VectorID        As Long
    MergedIntoID    As Long
    IsSingleLine    As Boolean    ' NEW: Flag for single-line walls (centerlines)
    Thickness       As Double    ' NEW: thickness read from XData if available (mm)
End Type

Private Type VectorLine
    angle           As Double
    BasePoint       As Point2D
    SegmentIDs()    As Long
    segmentCount    As Long
End Type

Private Type WallPair
    Seg1ID          As Long
    Seg2ID          As Long
    Thickness       As Double
    IsValid         As Boolean
End Type

Private Type CenterLine
    startPt         As Point2D
    endPt           As Point2D
    Length          As Double
    angle           As Double
    SourcePairID    As Long
    VectorID        As Long
    IsActive        As Boolean
    uniqueID        As String    ' NEW: Unique identifier to prevent duplicates
    Thickness       As Double    ' NEW: thickness associated to this centerline (mm)
    WallType        As String    ' NEW: e.g., "W200" or "" when unknown
End Type

Private Type overlapResult
    HasOverlap      As Boolean
    OverlapPercent  As Double
End Type

Private Type KDNode
    segmentID       As Long
    splitDim        As Integer
    splitValue      As Double
    LeftChild       As Long
    RightChild      As Long
End Type

Private Type BoundingBox
    minX            As Double
    maxX            As Double
    minY            As Double
    maxY            As Double
End Type

Private Type CornerNode
    PointX          As Double
    PointY          As Double
    ConnectedLineIDs() As Long
    ConnectedCount  As Long
    IsProcessed     As Boolean
End Type

Private Type ParallelBand
    RepresentativeID As Long
    CenterLineIDs() As Long
    count           As Long
End Type

Private Type IntervalNode
    Low             As Double
    High            As Double
    segmentID       As Long
    MaxHigh         As Double
    LeftChild       As Long
    RightChild      As Long
    height          As Integer
End Type

Private Type ProjectionData
    segmentID       As Long
    StartProj       As Double
    endProj         As Double
End Type

' ==================== MODULE VARIABLES ====================
Private mAcadApp    As Object
Private mAcadDoc    As Object
Private mWallThicknesses() As Double
Private mWallThicknessCount As Long
Private mTargetLayers As String
Private mAngleTolerance As Double
Private mExtendMultiplier As Double
Private mDistanceTolerance As Double
Private mEnableAutoExtend As Boolean
Private mEnableIntersection As Boolean

Private mSegments() As WallSegment
Private mSegmentCount As Long
Private mKDTree()   As KDNode
Private mKDNodeCount As Long
Private mUnionParent() As Long
Private mUnionRank() As Long

Private mVectorLines() As VectorLine
Private mVectorCount As Long
Private mWallPairs() As WallPair
Private mPairCount  As Long
Private mCenterLines() As CenterLine
Private mCenterCount As Long

Private mIntervalTree() As IntervalNode
Private mIntervalTreeCount As Long

Private mCornerNodes() As CornerNode
Private mCornerNodeCount As Long

Private mPerfStartTime As Double

' NEW: Track processed combinations to prevent duplicates
Private mProcessedCombinations As Object  ' Dictionary

' ==================== CONSTANTS ====================
Private Const PI    As Double = 3.14159265358979
Private Const HALF_PI As Double = 1.5707963267949
Private Const DEG_TO_RAD As Double = 1.74532925199433E-02
Private Const MIN_SEGMENT_LENGTH As Double = 1#
Private Const EPSILON As Double = 0.0000001

' ==================== INITIALIZATION ====================
Private Function InitializeAutoCAD() As Boolean
    On Error Resume Next
    Set mAcadApp = GetObject(, "AutoCAD.Application")
    If mAcadApp Is Nothing Then
        Set mAcadApp = CreateObject("AutoCAD.Application")
    End If
    If Not mAcadApp Is Nothing Then
        Set mAcadDoc = mAcadApp.ActiveDocument
        mAcadApp.Visible = True
        InitializeAutoCAD = True
    Else
        InitializeAutoCAD = False
    End If
    On Error GoTo 0
End Function

Public Sub ProcessWallConversion(thicknessInput As String, _
        layerFilter As String, _
        doorWidthsInput As String, _
        columnWidthsInput As String, _
        autoJoinGapInput As String, _
        axisSnapDistInput As String, _
        angleTol As Double, _
        extendMult As Double, _
        distTol As Double, _
        enableExtend As Boolean, _
        enableIntersect As Boolean, _
        breakAtGrid As Boolean, _
        extendToGrid As Boolean, _
        axesObjects() As Object, _
        axesCount As Long)
    On Error GoTo ErrHandler

Debug.Print "========== WALL CONVERTER v3.1 ENHANCED START =========="
Debug.Print "Time: " & Now

    mPerfStartTime = Timer
    ' Store setting
    mExtendLinesToGridIntersections = extendToGrid

    ' Initialize Dictionary for duplicate tracking
    Set mProcessedCombinations = CreateObject("Scripting.Dictionary")

    ' Initialize AutoCAD
    On Error Resume Next
    Set mAcadApp = GetObject(, "AutoCAD.Application")
    If mAcadApp Is Nothing Then
        Set mAcadApp = CreateObject("AutoCAD.Application")
    End If
    On Error GoTo ErrHandler

    If mAcadApp Is Nothing Then
        MsgBox "Cannot connect to AutoCAD!", vbCritical, "Error"
        Exit Sub
    End If

    Set mAcadDoc = mAcadApp.ActiveDocument
    mAcadApp.Visible = True

    ' Store settings
    mTargetLayers = Trim(layerFilter)
    mAngleTolerance = angleTol
    mExtendMultiplier = extendMult
    mDistanceTolerance = distTol
    mEnableAutoExtend = enableExtend
    mEnableIntersection = enableIntersect
    mBreakAtGridIntersection = breakAtGrid

    ' Determine single-line processing mode:
    ' Single-line mode is enabled ONLY when user explicitly sets thicknessInput = "1"
    If Len(Trim$(thicknessInput)) > 0 And Trim$(thicknessInput) = "1" Then
        mEnableSingleLineProcessing = True
    Else
        mEnableSingleLineProcessing = False
    End If

    ' Parse inputs
    If Not mEnableSingleLineProcessing Then
        If Not ParseThicknesses(thicknessInput) Then
            MsgBox "Invalid wall thickness input!", vbExclamation, "Input Error"
            Exit Sub
        End If
    Else
        ' In single-line mode we do not parse the thickness list: single lines are treated as centerlines
        ReDim mWallThicknesses(0 To 0)
        mWallThicknessCount = 0
    End If

    ParseBreakWidths doorWidthsInput, mDoorWidths, mDoorWidthCount
    ParseBreakWidths columnWidthsInput, mColumnWidths, mColumnWidthCount

    If Len(Trim(axisSnapDistInput)) > 0 And IsNumeric(axisSnapDistInput) Then
        mAxisSnapDistance = CDbl(axisSnapDistInput)
    Else
        mAxisSnapDistance = 50
    End If

    If Len(Trim(autoJoinGapInput)) > 0 And IsNumeric(autoJoinGapInput) Then
        mAutoJoinGapDistance = CDbl(autoJoinGapInput)
    Else
        mAutoJoinGapDistance = 0
    End If

    If axesCount > 0 Then
        ParseAxes axesObjects, axesCount
Debug.Print "Loaded " & mAxesCount & " structural axes"
    Else
        mAxesCount = 0
    End If

Debug.Print "Settings loaded successfully"

    Dim t1          As Double

    ' ===== STEP 0: Collect raw segments (including single lines) =====
    t1 = Timer
    If Not CollectLineSegments() Then
        MsgBox "No valid line segments found!", vbExclamation, "No Data"
        Exit Sub
    End If
Debug.Print "STEP 0: Collected " & mSegmentCount & " raw segments (" & Format(Timer - t1, "0.00") & "s)"

    ' If single-line processing mode is explicitly requested, mark all collected segments as single-line
    If mEnableSingleLineProcessing Then
        Dim si      As Long
        For si = 0 To mSegmentCount - 1
            mSegments(si).IsSingleLine = True
        Next si
Debug.Print "Single-line mode active: all segments marked as single-line."
    End If

    ' ===== STEP 0A: Normalizing angles =====

Debug.Print "STEP 0A: Normalizing angles..."
    t1 = Timer
    NormalizeSegmentAngles
Debug.Print "STEP 0A: Completed (" & Format(Timer - t1, "0.00") & "s)"



    ' ===== STEP 1: Merge overlapping segments =====
    t1 = Timer
    Dim mergedCount As Long
    mergedCount = MergeOverlappingSegments()
Debug.Print "STEP 1: Merged " & mergedCount & " overlapping segments (" & Format(Timer - t1, "0.00") & "s)"

    ' ===== STEP 2: Detect wall pairs (skip when single-line mode) =====
    t1 = Timer
    Dim pairCount   As Long
    If Not mEnableSingleLineProcessing Then
        BuildKDTree
        pairCount = DetectWallPairs()
    Else
        ' In single-line mode we do not detect pairs; leave pairCount = 0
        pairCount = 0
    End If
Debug.Print "STEP 2: Detected " & pairCount & " wall pairs (" & Format(Timer - t1, "0.00") & "s)"

    ' Generate centerlines from pairs (if any)
    If pairCount > 0 Then
        GenerateCenterLinesFromPairs
    Else
        ' Ensure centerlines array is initialized
        ReDim mCenterLines(0 To 0)
        mCenterCount = 0
    End If

    ' NEW: Add single-line segments directly as centerlines
    Dim singleLineCount As Long
    singleLineCount = AddSingleLinesToCenterLines()
Debug.Print "        Added " & singleLineCount & " single-line centerlines"
Debug.Print "        Total initial centerlines: " & mCenterCount

    If mCenterCount = 0 Then
        MsgBox "No centerlines generated!", vbInformation, "No Results"
        Exit Sub
    End If

    ' ===== STEP 2A GLOBAL AXIS ALIGNMENT (Make "Straight XLines") =====
    ' This merges close parallel vectors and aligns far-apart segments onto a single axis
    t1 = Timer
    Dim alignedCount As Long
    ' Tolerance for alignment: roughly 1.5x max wall thickness or 300mm
    Dim alignTol    As Double
    alignTol = 300
    If mWallThicknessCount > 0 Then
        alignTol = mWallThicknesses(mWallThicknessCount - 1) * 1.5
    End If

    alignedCount = AlignCenterLinesToSmartGrid(alignTol)
Debug.Print "STEP 2A: Aligned " & alignedCount & " lines to global smart axes (" & Format(Timer - t1, "0.00") & "s)"


    ' ===== STEP 3: TIER 1 - VECTOR-BASED EXACT RECOVERY (STRICT) =====
    t1 = Timer
    Dim exactJoined As Long
    exactJoined = RecoverWallLengthExactExhaustive()
Debug.Print "STEP 3 (Tier 1): Vector-based recovery: " & exactJoined & " gaps filled (" & Format(Timer - t1, "0.00") & "s)"

    ' ===== STEP 3A: TIER 2 - GLOBAL GAP RECOVERY (RELAXED) =====
    t1 = Timer
    Dim globalRecovered As Long
    globalRecovered = RecoverGapsGlobal()
Debug.Print "STEP 3A (Tier 2): Global gap recovery: " & globalRecovered & " additional gaps filled (" & Format(Timer - t1, "0.00") & "s)"

    ' ===== STEP 3B: TIER 2 - GLOBAL MERGE (RELAXED) =====
    t1 = Timer
    Dim globalMerged As Long
    globalMerged = MergeOverlappingCenterLinesGlobal()
Debug.Print "STEP 3B (Tier 2): Global merge: " & globalMerged & " overlaps removed (" & Format(Timer - t1, "0.00") & "s)"

    ' ===== STEP 3C: CORNER-AWARE GAP RECOVERY (NEW) =====
    t1 = Timer
    BuildCornerNodeGraph
    Dim cornerGapsRecovered As Long
    cornerGapsRecovered = RecoverCornerGaps()
Debug.Print "STEP 3C (Corner): Recovered " & cornerGapsRecovered & " corner gaps (" & Format(Timer - t1, "0.00") & "s)"

    ' ===== STEP 4: ENHANCED - Auto-join by gap with iteration =====
    t1 = Timer
    Dim gapJoined   As Long
    gapJoined = 0
    If mAutoJoinGapDistance > 0 Then
        gapJoined = RecoverWallLengthByGapExhaustive()
Debug.Print "STEP 4: Joined " & gapJoined & " segments by gap (" & Format(Timer - t1, "0.00") & "s)"
    Else
Debug.Print "STEP 4: Skipped (gap distance = 0)"
    End If

    ' ===== STEP 5: ENHANCED - Auto-extend with iteration =====
    t1 = Timer
    Dim extendCount As Long
    extendCount = 0
    If mEnableAutoExtend Then
        extendCount = AutoExtendCenterLinesExhaustive()
Debug.Print "STEP 5: Extended " & extendCount & " centerlines (" & Format(Timer - t1, "0.00") & "s)"
    Else
Debug.Print "STEP 5: Skipped (extend disabled)"
    End If

    ' ===== STEP 6: Snap to structural axes =====
    t1 = Timer
    Dim snappedCount As Long
    snappedCount = 0
    If mAxesCount > 0 And mAxisSnapDistance > 0 Then
        snappedCount = SnapCenterLinesToAxes()
Debug.Print "STEP 6: Snapped " & snappedCount & " centerlines to axes (" & Format(Timer - t1, "0.00") & "s)"
    Else
Debug.Print "STEP 6: Skipped (no axes)"
    End If

    ' ===== STEP 6A: NEW - Merge parallel close centerlines (< 2*wall thickness) =====
    t1 = Timer
    Dim mergedParallel As Long
    mergedParallel = MergeParallelCloseCenterLines()
Debug.Print "STEP 6A: Merged " & mergedParallel & " parallel close centerlines (" & Format(Timer - t1, "0.00") & "s)"

    ' ===== STEP 6B: APPLY AUTOCAD OVERKILL =====
    t1 = Timer
    Dim overkillRemoved As Long
    overkillRemoved = ApplyAutoCADOverkill()
Debug.Print "STEP 6B: AutoCAD OVERKILL removed " & overkillRemoved & " duplicates (" & Format(Timer - t1, "0.00") & "s)"

    ' ===== STEP 7: BREAK AT GRID (option) =====
    If mBreakAtGridIntersection And mAxesCount > 0 Then
        t1 = Timer

        Dim breakCount As Long
        breakCount = BreakCenterLinesAtGridIntersections()
Debug.Print "STEP 7: Broke " & breakCount & " centerlines at grid (" & Format(Timer - t1, "0.00") & "s)"

        ' OVERKILL after breaking
        overkillRemoved = ApplyAutoCADOverkill()
Debug.Print "        OVERKILL after break: Removed " & overkillRemoved & " segments"
    End If

    ' ===== STEP 8: EXTEND LINES TO GRID INTERSECTIONS (NEW OPTION) =====
    If mExtendLinesToGridIntersections And mAxesCount > 0 Then
        t1 = Timer
        Dim extendedCount As Long
        extendedCount = ExtendCenterLinesToGridIntersections()
Debug.Print "STEP 8: Extended " & extendedCount & " lines to grid intersections (" & Format(Timer - t1, "0.00") & "s)"

        ' CRITICAL: Compact after extend+merge
        CompactCenterLineArray

        ' OVERKILL after extend (final cleanup)
        overkillRemoved = ApplyAutoCADOverkill()
Debug.Print "        OVERKILL after extend: Removed " & overkillRemoved & " segments"
    End If

    ' ===== Get insertion point and draw =====
    Dim InsertPt    As Variant
    InsertPt = GetInsertionPoint()
    If IsEmpty(InsertPt) Then
        MsgBox "Operation cancelled by user.", vbInformation, "Cancelled"
        Exit Sub
    End If

    ' Calculate bounding box of all centerlines to determine offset
    Dim minX As Double, minY As Double, maxX As Double, maxY As Double
    minX = 1E+300: minY = 1E+300
    maxX = -1E+300: maxY = -1E+300

    Dim i           As Long
    For i = 0 To mCenterCount - 1
        If mCenterLines(i).IsActive Then
            If mCenterLines(i).startPt.X < minX Then minX = mCenterLines(i).startPt.X
            If mCenterLines(i).startPt.Y < minY Then minY = mCenterLines(i).startPt.Y
            If mCenterLines(i).endPt.X < minX Then minX = mCenterLines(i).endPt.X
            If mCenterLines(i).endPt.Y < minY Then minY = mCenterLines(i).endPt.Y

            If mCenterLines(i).startPt.X > maxX Then maxX = mCenterLines(i).startPt.X
            If mCenterLines(i).startPt.Y > maxY Then maxY = mCenterLines(i).startPt.Y
            If mCenterLines(i).endPt.X > maxX Then maxX = mCenterLines(i).endPt.X
            If mCenterLines(i).endPt.Y > maxY Then maxY = mCenterLines(i).endPt.Y
        End If
    Next i

    ' Calculate offset to place bottom-left corner at insertion point
    Dim offsetX As Double, offsetY As Double
    offsetX = InsertPt(0) - minX  ' Move left edge to insertion X
    offsetY = InsertPt(1) - minY  ' Move bottom edge to insertion Y

Debug.Print "Bounding box: [" & Format(minX, "0.00") & ", " & Format(minY, "0.00") & "] to [" & Format(maxX, "0.00") & ", " & Format(maxY, "0.00") & "]"
Debug.Print "Offset: (" & Format(offsetX, "0.00") & ", " & Format(offsetY, "0.00") & ")"

    ApplyOffsetToCenterLines offsetX, offsetY

    If mBreakAtGridIntersection And mAxesCount > 0 Then
        t1 = Timer
        ApplyOffsetToAxes offsetX, offsetY
        breakCount = BreakCenterLinesAtGridIntersections()
Debug.Print "BREAK AT GRID: Broke " & breakCount & " centerlines (" & Format(Timer - t1, "0.00") & "s)"
    ElseIf mAxesCount > 0 Then
        ApplyOffsetToAxes offsetX, offsetY
    End If

    ' Draw results
    t1 = Timer
    Dim drawnAxesCount As Long
    drawnAxesCount = 0
    If mAxesCount > 0 Then
        EnsureLayerExists "DTS_AXIS_LINE", 5
        drawnAxesCount = DrawAxes()
    End If

    EnsureLayerExists "DTS_WALL_DIAGRAM", 2
    Dim drawnCenterCount As Long
    drawnCenterCount = DrawCenterLines()
Debug.Print "DRAW: Drew " & drawnCenterCount & " centerlines (" & Format(Timer - t1, "0.00") & "s)"

    overkillRemoved = ApplyAutoCADOverkill()

    Dim totalTime   As Double
    totalTime = Timer - mPerfStartTime
Debug.Print "========== COMPLETED in " & Format(totalTime, "0.00") & "s =========="

    On Error Resume Next
    mAcadDoc.Regen 0
    mAcadApp.ZoomExtents
    Set mProcessedCombinations = Nothing
    On Error GoTo ErrHandler

    Exit Sub
ErrHandler:
    MsgBox "Processing error:" & vbCrLf & vbCrLf & _
            err.description & vbCrLf & _
            "Error #" & err.number, vbCritical, "Wall Converter Error"
Debug.Print "ERROR: " & err.description
End Sub

' ==================== PARSE FUNCTIONS (unchanged) ====================
Private Function ParseThicknesses(thicknessInput As String) As Boolean
    On Error GoTo ErrHandler
    ParseThicknesses = False

    If Len(Trim(thicknessInput)) = 0 Then Exit Function

    Dim parts()     As String
    parts = Split(thicknessInput, ",")

    ReDim mWallThicknesses(0 To UBound(parts))
    mWallThicknessCount = 0

    Dim i As Long, val As Double, s As String
    For i = 0 To UBound(parts)
        s = Trim(parts(i))
        If Len(s) > 0 And IsNumeric(s) Then
            val = CDbl(s)
            If val > 0 Then
                mWallThicknesses(mWallThicknessCount) = val
                mWallThicknessCount = mWallThicknessCount + 1
            End If
        End If
    Next i

    If mWallThicknessCount > 0 Then
        ReDim Preserve mWallThicknesses(0 To mWallThicknessCount - 1)
        QuickSortDoubles mWallThicknesses, 0, mWallThicknessCount - 1
        ParseThicknesses = True
    End If

    Exit Function
ErrHandler:
    ParseThicknesses = False
End Function

Private Sub ParseBreakWidths(widthInput As String, ByRef widths() As Double, ByRef count As Long)
    On Error Resume Next
    count = 0
    If Len(Trim(widthInput)) = 0 Then Exit Sub

    Dim parts()     As String
    parts = Split(widthInput, ",")
    ReDim widths(0 To UBound(parts))

    Dim i As Long, val As Double, s As String
    For i = 0 To UBound(parts)
        s = Trim(parts(i))
        If Len(s) > 0 And IsNumeric(s) Then
            val = CDbl(s)
            If val > 0 Then
                widths(count) = val
                count = count + 1
            End If
        End If
    Next i

    If count > 0 Then
        ReDim Preserve widths(0 To count - 1)
        QuickSortDoubles widths, 0, count - 1
    End If
    On Error GoTo 0
End Sub

Private Sub ParseAxes(axesObjects() As Object, axesCount As Long)
    On Error Resume Next
    mAxesCount = 0
    If axesCount = 0 Then Exit Sub

    ReDim mAxes(0 To axesCount - 1)

    Dim i As Long, sp As Variant, ep As Variant
    For i = 0 To axesCount - 1
        sp = axesObjects(i).StartPoint
        ep = axesObjects(i).EndPoint

        mAxes(mAxesCount).startPt.X = sp(0)
        mAxes(mAxesCount).startPt.Y = sp(1)
        mAxes(mAxesCount).endPt.X = ep(0)
        mAxes(mAxesCount).endPt.Y = ep(1)

        Dim dx As Double, dy As Double
        dx = mAxes(mAxesCount).endPt.X - mAxes(mAxesCount).startPt.X
        dy = mAxes(mAxesCount).endPt.Y - mAxes(mAxesCount).startPt.Y

        mAxes(mAxesCount).Length = Sqr(dx * dx + dy * dy)
        mAxes(mAxesCount).angle = Atan2(dy, dx)

        Dim absAngle As Double
        absAngle = Abs(mAxes(mAxesCount).angle)
        mAxes(mAxesCount).isHorizontal = (absAngle < 0.0873 Or absAngle > 3.0543)
        mAxes(mAxesCount).IsVertical = (Abs(absAngle - HALF_PI) < 0.0873)

        mAxesCount = mAxesCount + 1
    Next i

    If mAxesCount > 0 Then ReDim Preserve mAxes(0 To mAxesCount - 1)
    On Error GoTo 0
End Sub

' ==================== COLLECT SEGMENTS ====================
Private Function CollectLineSegments() As Boolean
    On Error GoTo ErrHandler
    CollectLineSegments = False

    Dim ssName      As String
    ssName = "WALL_CONV_TEMP_" & Format(Now, "hhmmss")

    On Error Resume Next
    Set mAcadApp = GetObject(, "AutoCAD.Application")
    If Not mAcadApp Is Nothing Then
        Set mAcadDoc = mAcadApp.ActiveDocument
    End If
    On Error GoTo ErrHandler

    On Error Resume Next
    mAcadApp.ActiveDocument.SelectionSets.item(ssName).Delete
    On Error GoTo ErrHandler

    Dim ss          As Object
    Set ss = mAcadApp.ActiveDocument.SelectionSets.Add(ssName)

    On Error Resume Next
    AppActivate mAcadApp.Caption
    DoEvents
    On Error GoTo ErrHandler

    mAcadApp.ActiveDocument.Utility.prompt vbCrLf & "Select wall LINE objects (double walls OR single centerlines): "

    On Error Resume Next
    ss.SelectOnScreen
    If err.number <> 0 Then
        ss.Delete
        CollectLineSegments = False
        Exit Function
    End If
    On Error GoTo ErrHandler

    If ss.count = 0 Then
        ss.Delete
        Exit Function
    End If

    ReDim mSegments(0 To ss.count - 1)
    mSegmentCount = 0

    Dim i As Long, ent As Object
    For i = 0 To ss.count - 1
        Set ent = ss.item(i)
        If IsLineEntity(ent) And isValidLayer(ent) Then
            If ExtractSegmentInfo(ent, mSegments(mSegmentCount)) Then
                mSegments(mSegmentCount).GroupID = mSegmentCount
                mSegments(mSegmentCount).PairSegmentID = -1
                mSegments(mSegmentCount).VectorID = -1
                mSegments(mSegmentCount).MergedIntoID = -1
                mSegments(mSegmentCount).IsSingleLine = False  ' Will be determined later
                mSegmentCount = mSegmentCount + 1
            End If
        End If
    Next i

    ss.Delete

    If mSegmentCount > 0 Then
        ReDim Preserve mSegments(0 To mSegmentCount - 1)
        CollectLineSegments = True
    End If

    Exit Function
ErrHandler:
    CollectLineSegments = False
    On Error Resume Next
    If Not ss Is Nothing Then ss.Delete
    On Error GoTo 0
End Function

Private Function IsLineEntity(ent As Object) As Boolean
    On Error Resume Next
    IsLineEntity = (LCase(ent.ObjectName) = "acdbline")
    If err.number <> 0 Then IsLineEntity = False
    On Error GoTo 0
End Function

Private Function isValidLayer(ent As Object) As Boolean
    On Error Resume Next
    isValidLayer = True
    If Len(mTargetLayers) = 0 Then Exit Function

    Dim entLayer    As String
    entLayer = LCase(Trim(ent.layer))

    Dim layers()    As String
    layers = Split(mTargetLayers, ",")

    isValidLayer = False
    Dim i           As Long
    For i = 0 To UBound(layers)
        If entLayer = LCase(Trim(layers(i))) Then
            isValidLayer = True
            Exit Function
        End If
    Next i
    On Error GoTo 0
End Function

Private Function ExtractSegmentInfo(ent As Object, ByRef seg As WallSegment) As Boolean
    On Error GoTo ErrHandler
    ExtractSegmentInfo = False

    Dim sp As Variant, ep As Variant
    sp = ent.StartPoint
    ep = ent.EndPoint

    seg.Handle = ent.Handle
    seg.startPt.X = sp(0): seg.startPt.Y = sp(1)
    seg.endPt.X = ep(0): seg.endPt.Y = ep(1)

    Dim dx As Double, dy As Double
    dx = seg.endPt.X - seg.startPt.X
    dy = seg.endPt.Y - seg.startPt.Y
    seg.Length = Sqr(dx * dx + dy * dy)

    If seg.Length < MIN_SEGMENT_LENGTH Then Exit Function

    seg.angle = Atan2(dy, dx)
    seg.layer = ent.layer
    seg.IsProcessed = False
    seg.PairSegmentID = -1
    seg.Thickness = 0    ' default: unknown

    ' Try to read thickness XData from wall drawing
    On Error Resume Next
    Dim xdType As Variant, xdVal As Variant
    ent.GetXData "DTS_APP", xdType, xdVal
    If err.number = 0 Then
        If Not IsEmpty(xdVal) And IsArray(xdVal) Then
            If UBound(xdVal) >= 1 Then
                If IsNumeric(xdVal(1)) Then
                    seg.Thickness = CDbl(xdVal(1))
                End If
            End If
        End If
    End If
    err.Clear
    On Error GoTo ErrHandler

    ExtractSegmentInfo = True
    Exit Function
ErrHandler:
    ExtractSegmentInfo = False
End Function
' ==================== STEP 0A: NORMALIZE ANGLES ====================
Private Sub NormalizeSegmentAngles()
    On Error Resume Next

    If mSegmentCount = 0 Then Exit Sub

    Dim i           As Long
    Dim normalizedCount As Long
    normalizedCount = 0

    For i = 0 To mSegmentCount - 1
        Dim originalAngle As Double
        originalAngle = mSegments(i).angle

        ' Normalize to 0-2p range
        Dim normalizedAngle As Double
        normalizedAngle = originalAngle

        Do While normalizedAngle < 0
            normalizedAngle = normalizedAngle + 2 * PI
        Loop
        Do While normalizedAngle >= 2 * PI
            normalizedAngle = normalizedAngle - 2 * PI
        Loop

        ' Snap to nearest cardinal direction (0°, 90°, 180°, 270°)
        ' with tolerance of ±5° (0.0873 radians)
        Dim snapTolerance As Double
        snapTolerance = 5 * DEG_TO_RAD  ' 5 degrees

        Dim targetAngle As Double
        targetAngle = normalizedAngle

        ' Check 0° (horizontal right)
        If Abs(normalizedAngle) < snapTolerance Or _
                Abs(normalizedAngle - 2 * PI) < snapTolerance Then
            targetAngle = 0
            normalizedCount = normalizedCount + 1

            ' Check 90° (vertical up)
        ElseIf Abs(normalizedAngle - HALF_PI) < snapTolerance Then
            targetAngle = HALF_PI
            normalizedCount = normalizedCount + 1

            ' Check 180° (horizontal left)
        ElseIf Abs(normalizedAngle - PI) < snapTolerance Then
            targetAngle = PI
            normalizedCount = normalizedCount + 1

            ' Check 270° (vertical down)
        ElseIf Abs(normalizedAngle - (3 * HALF_PI)) < snapTolerance Then
            targetAngle = 3 * HALF_PI
            normalizedCount = normalizedCount + 1
        End If

        ' Only modify if normalized
        If Abs(targetAngle - normalizedAngle) > EPSILON Then
            ' Recalculate endpoints based on normalized angle
            Dim Length As Double
            Length = mSegments(i).Length

            Dim dx As Double, dy As Double
            dx = Cos(targetAngle) * Length
            dy = Sin(targetAngle) * Length

            ' Update endpoint while keeping start point
            mSegments(i).endPt.X = mSegments(i).startPt.X + dx
            mSegments(i).endPt.Y = mSegments(i).startPt.Y + dy
            mSegments(i).angle = targetAngle

Debug.Print "  Normalized segment " & i & ": " & _
        Format(originalAngle / DEG_TO_RAD, "0.000") & "° to " & _
        Format(targetAngle / DEG_TO_RAD, "0.000") & "°"
        End If
    Next i

Debug.Print "Angle normalization: " & normalizedCount & " segments adjusted"

    On Error GoTo 0
End Sub
' ==================== MERGE OVERLAPPING (ITERATIVE VERSION) ====================
Private Function MergeOverlappingSegments() As Long
    On Error Resume Next
    MergeOverlappingSegments = 0

    If mSegmentCount <= 1 Then Exit Function

    ' ITERATIVE APPROACH: Keep merging until no more changes
    Dim maxIterations As Long
    maxIterations = 5

    Dim iteration   As Long
    iteration = 0

    Dim totalMerged As Long
    totalMerged = 0

    Do While iteration < maxIterations
        iteration = iteration + 1

        ' Rebuild vector groups each iteration
        GroupSegmentsByVector

        Dim mergedThisIteration As Long
        mergedThisIteration = 0

        Dim v       As Long
        For v = 0 To mVectorCount - 1
            If mVectorLines(v).segmentCount <= 1 Then GoTo NextVector

            Dim merged As Long
            merged = MergeOverlappingInVector(v)
            mergedThisIteration = mergedThisIteration + merged

NextVector:
        Next v

        If mergedThisIteration = 0 Then
Debug.Print "    Merge iteration " & iteration & ": converged (no changes)"
            Exit Do
        End If

        totalMerged = totalMerged + mergedThisIteration
Debug.Print "    Merge iteration " & iteration & ": merged " & mergedThisIteration & " segments"

        ' Compact array after each iteration
        CompactSegmentArray
    Loop

    If iteration >= maxIterations Then
Debug.Print "    WARNING: Merge reached max iterations (" & maxIterations & ")"
    End If

    MergeOverlappingSegments = totalMerged
    On Error GoTo 0
End Function

Private Sub GroupSegmentsByVector()
    ReDim mVectorLines(0 To mSegmentCount - 1)
    mVectorCount = 0

    Dim i           As Long
    For i = 0 To mSegmentCount - 1
        If mSegments(i).VectorID >= 0 Then GoTo NextSeg

        mVectorLines(mVectorCount).angle = mSegments(i).angle
        mVectorLines(mVectorCount).BasePoint = mSegments(i).startPt
        ReDim mVectorLines(mVectorCount).SegmentIDs(0 To mSegmentCount - 1)
        mVectorLines(mVectorCount).segmentCount = 0

        mSegments(i).VectorID = mVectorCount
        mVectorLines(mVectorCount).SegmentIDs(0) = i
        mVectorLines(mVectorCount).segmentCount = 1

        Dim j       As Long
        For j = i + 1 To mSegmentCount - 1
            If mSegments(j).VectorID >= 0 Then GoTo NextJ

            If AreSegmentsCollinear(i, j) Then
                mSegments(j).VectorID = mVectorCount
                mVectorLines(mVectorCount).SegmentIDs(mVectorLines(mVectorCount).segmentCount) = j
                mVectorLines(mVectorCount).segmentCount = mVectorLines(mVectorCount).segmentCount + 1
            End If
NextJ:
        Next j

        If mVectorLines(mVectorCount).segmentCount > 0 Then
            ReDim Preserve mVectorLines(mVectorCount).SegmentIDs(0 To mVectorLines(mVectorCount).segmentCount - 1)
            mVectorCount = mVectorCount + 1
        End If

NextSeg:
    Next i

    If mVectorCount > 0 Then ReDim Preserve mVectorLines(0 To mVectorCount - 1)
End Sub

Private Function AreSegmentsCollinear(id1 As Long, id2 As Long) As Boolean
    Dim angleDiff   As Double
    angleDiff = Abs(mSegments(id1).angle - mSegments(id2).angle)
    If angleDiff > PI Then angleDiff = 2 * PI - angleDiff

    If angleDiff > (mAngleTolerance * DEG_TO_RAD) Then
        AreSegmentsCollinear = False
        Exit Function
    End If

    Dim dist1 As Double, dist2 As Double
    dist1 = PointToLineDistanceInfinite(mSegments(id2).startPt, mSegments(id1))
    dist2 = PointToLineDistanceInfinite(mSegments(id2).endPt, mSegments(id1))

    AreSegmentsCollinear = (dist1 <= mDistanceTolerance And dist2 <= mDistanceTolerance)
End Function

Private Function PointToLineDistanceInfinite(pt As Point2D, line As WallSegment) As Double
    Dim dx As Double, dy As Double
    dx = line.endPt.X - line.startPt.X
    dy = line.endPt.Y - line.startPt.Y

    Dim dlen        As Double
    dlen = Sqr(dx * dx + dy * dy)

    If dlen < EPSILON Then
        PointToLineDistanceInfinite = PointDistance2D(pt, line.startPt)
    Else
        PointToLineDistanceInfinite = Abs(dy * (pt.X - line.startPt.X) - dx * (pt.Y - line.startPt.Y)) / dlen
    End If
End Function

Private Function MergeOverlappingInVector(VectorID As Long) As Long
    MergeOverlappingInVector = 0

    Dim segCount    As Long
    segCount = mVectorLines(VectorID).segmentCount

    If segCount <= 1 Then Exit Function

    Dim projections() As ProjectionData
    ReDim projections(0 To segCount - 1)

    Dim i As Long, segID As Long
    For i = 0 To segCount - 1
        segID = mVectorLines(VectorID).SegmentIDs(i)

        projections(i).segmentID = segID
        projections(i).StartProj = ProjectPointOnVector(mSegments(segID).startPt, VectorID)
        projections(i).endProj = ProjectPointOnVector(mSegments(segID).endPt, VectorID)

        If projections(i).StartProj > projections(i).endProj Then
            Dim tmp As Double
            tmp = projections(i).StartProj
            projections(i).StartProj = projections(i).endProj
            projections(i).endProj = tmp
        End If
    Next i

    QuickSortProjections projections, 0, segCount - 1

    Dim j           As Long
    For i = 0 To segCount - 1
        segID = projections(i).segmentID
        If mSegments(segID).MergedIntoID >= 0 Then GoTo NextMerge

        For j = i + 1 To segCount - 1
            Dim otherID As Long
            otherID = projections(j).segmentID

            If mSegments(otherID).MergedIntoID >= 0 Then GoTo NextOther

            Dim overlapStart As Double, overlapEnd As Double
            overlapStart = Max(projections(i).StartProj, projections(j).StartProj)
            overlapEnd = Min(projections(i).endProj, projections(j).endProj)

            If overlapEnd >= overlapStart - mDistanceTolerance Then
                MergeTwoSegments segID, otherID
                mSegments(otherID).MergedIntoID = segID

                projections(i).StartProj = Min(projections(i).StartProj, projections(j).StartProj)
                projections(i).endProj = Max(projections(i).endProj, projections(j).endProj)

                MergeOverlappingInVector = MergeOverlappingInVector + 1
            End If

NextOther:
        Next j
NextMerge:
    Next i
End Function

Private Function ProjectPointOnVector(pt As Point2D, VectorID As Long) As Double
    Dim basePt      As Point2D
    basePt = mVectorLines(VectorID).BasePoint

    Dim angle       As Double
    angle = mVectorLines(VectorID).angle

    Dim dx As Double, dy As Double
    dx = pt.X - basePt.X
    dy = pt.Y - basePt.Y

    ProjectPointOnVector = dx * Cos(angle) + dy * Sin(angle)
End Function

Private Sub QuickSortProjections(ByRef arr() As ProjectionData, Left As Long, Right As Long)
    If Left >= Right Then Exit Sub

    Dim pivot       As Double
    pivot = arr((Left + Right) \ 2).StartProj

    Dim i As Long, j As Long
    i = Left: j = Right

    Do While i <= j
        Do While arr(i).StartProj < pivot
            i = i + 1
        Loop
        Do While arr(j).StartProj > pivot
            j = j - 1
        Loop
        If i <= j Then
            Dim tmp As ProjectionData
            tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
            i = i + 1: j = j - 1
        End If
    Loop

    If Left < j Then QuickSortProjections arr, Left, j
    If i < Right Then QuickSortProjections arr, i, Right
End Sub

Private Sub MergeTwoSegments(targetID As Long, sourceID As Long)
    ' Determine dominant vector based on the longer segment
    Dim baseAngle   As Double
    Dim refPt       As Point2D

    If mSegments(targetID).Length >= mSegments(sourceID).Length Then
        baseAngle = mSegments(targetID).angle
        refPt = mSegments(targetID).startPt
    Else
        baseAngle = mSegments(sourceID).angle
        refPt = mSegments(sourceID).startPt
    End If

    ' --- FIX: FORCE ANGLE TO STRICT CARDINAL DIRECTION (0, 90, 180, 270) ---
    ' This prevents the "center rotation" when merging slightly offset lines.
    baseAngle = SnapToCardinalAngle(baseAngle)

    Dim cosA As Double, sinA As Double
    cosA = Cos(baseAngle)
    sinA = Sin(baseAngle)

    ' Collect all 4 points
    Dim points(0 To 3) As Point2D
    points(0) = mSegments(targetID).startPt
    points(1) = mSegments(targetID).endPt
    points(2) = mSegments(sourceID).startPt
    points(3) = mSegments(sourceID).endPt

    ' Project all points onto the base vector
    Dim minProj As Double, maxProj As Double
    minProj = 1E+300: maxProj = -1E+300

    Dim i           As Long
    For i = 0 To 3
        Dim dx As Double, dy As Double
        dx = points(i).X - refPt.X
        dy = points(i).Y - refPt.Y

        Dim proj    As Double
        proj = dx * cosA + dy * sinA

        If proj < minProj Then minProj = proj
        If proj > maxProj Then maxProj = proj
    Next i

    ' Reconstruct the merged segment on the dominant axis
    mSegments(targetID).startPt.X = refPt.X + minProj * cosA
    mSegments(targetID).startPt.Y = refPt.Y + minProj * sinA
    mSegments(targetID).endPt.X = refPt.X + maxProj * cosA
    mSegments(targetID).endPt.Y = refPt.Y + maxProj * sinA

    ' Recalculate properties
    mSegments(targetID).Length = maxProj - minProj
    mSegments(targetID).angle = baseAngle    ' Use the SNAPPED angle

    ' Force coordinate cleanup (remove floating point noise)
    If Abs(cosA) > 0.99 Then    ' Horizontal
        mSegments(targetID).startPt.Y = refPt.Y
        mSegments(targetID).endPt.Y = refPt.Y
    Else    ' Vertical
        mSegments(targetID).startPt.X = refPt.X
        mSegments(targetID).endPt.X = refPt.X
    End If
End Sub
Private Sub CompactSegmentArray()
    Dim writeIdx    As Long
    writeIdx = 0

    Dim i           As Long
    For i = 0 To mSegmentCount - 1
        If mSegments(i).MergedIntoID < 0 Then
            If writeIdx <> i Then
                mSegments(writeIdx) = mSegments(i)
            End If
            writeIdx = writeIdx + 1
        End If
    Next i

    mSegmentCount = writeIdx
    If mSegmentCount > 0 Then
        ReDim Preserve mSegments(0 To mSegmentCount - 1)
    End If
End Sub

' ==================== EXTEND CENTERLINES TO GRID INTERSECTIONS (COMPLETE REWRITE) ====================
Private Function ExtendCenterLinesToGridIntersections() As Long
    On Error Resume Next
    ExtendCenterLinesToGridIntersections = 0

    If mAxesCount = 0 Then Exit Function

Debug.Print "  Extending centerlines to grid intersections (STRICT)..."

    ' STEP 1: Build endpoint connection graph to identify dangling ends
    Dim endpointGraph() As EndpointNode
    Dim graphCount  As Long
    graphCount = BuildEndpointConnectionGraph(endpointGraph)

    If graphCount = 0 Then
Debug.Print "    No endpoints found for extension"
        Exit Function
    End If

    ' STEP 2: Process each axis to find grid segments
    Dim axisID      As Long
    Dim totalExtended As Long
    totalExtended = 0

    For axisID = 0 To mAxesCount - 1
        Dim gridSegments() As GridSegment
        Dim segmentCount As Long

        ' Find perpendicular intersections for this axis
        segmentCount = BuildGridSegmentsForAxis(axisID, gridSegments)

        If segmentCount > 0 Then
Debug.Print "    Axis " & axisID & ": " & segmentCount & " grid segments"

            ' Process each grid segment (cell between two perpendicular axes)
            Dim s   As Long
            For s = 0 To segmentCount - 1
                Dim extended As Long
                extended = ExtendLinesInGridSegment(axisID, gridSegments(s), endpointGraph, graphCount)
                totalExtended = totalExtended + extended
            Next s
        End If
    Next axisID

    ExtendCenterLinesToGridIntersections = totalExtended
Debug.Print "  Extended " & totalExtended & " dangling endpoints to grid"
End Function
' ==================== BUILD ENDPOINT CONNECTION GRAPH ====================
Private Function BuildEndpointConnectionGraph(ByRef graph() As EndpointNode) As Long
    On Error Resume Next

    ' Allocate space for all endpoints (2 per line)
    Dim nodeCount   As Long
    nodeCount = 0
    ReDim graph(0 To mCenterCount * 2 - 1)

    Dim connectionTolerance As Double
    connectionTolerance = 50  ' 50mm tolerance for connection detection

    Dim i           As Long
    For i = 0 To mCenterCount - 1
        If Not mCenterLines(i).IsActive Then GoTo NextEndpointBuild
        If mCenterLines(i).Length < EPSILON Then GoTo NextEndpointBuild

        ' Add start endpoint
        graph(nodeCount).lineID = i
        graph(nodeCount).IsStart = True
        graph(nodeCount).X = mCenterLines(i).startPt.X
        graph(nodeCount).Y = mCenterLines(i).startPt.Y
        graph(nodeCount).IsDangling = True
        graph(nodeCount).ConnectedLineID = -1
        nodeCount = nodeCount + 1

        ' Add end endpoint
        graph(nodeCount).lineID = i
        graph(nodeCount).IsStart = False
        graph(nodeCount).X = mCenterLines(i).endPt.X
        graph(nodeCount).Y = mCenterLines(i).endPt.Y
        graph(nodeCount).IsDangling = True
        graph(nodeCount).ConnectedLineID = -1
        nodeCount = nodeCount + 1

NextEndpointBuild:
    Next i

    If nodeCount > 0 Then ReDim Preserve graph(0 To nodeCount - 1)

    ' Detect connections between endpoints
    Dim n1 As Long, n2 As Long
    For n1 = 0 To nodeCount - 1
        For n2 = n1 + 1 To nodeCount - 1
            ' Skip if same line
            If graph(n1).lineID = graph(n2).lineID Then GoTo NextConnection

            ' Check distance
            Dim dist As Double
            dist = Sqr((graph(n1).X - graph(n2).X) ^ 2 + (graph(n1).Y - graph(n2).Y) ^ 2)

            Dim line1Angle As Double, line2Angle As Double
            line1Angle = mCenterLines(graph(n1).lineID).angle
            line2Angle = mCenterLines(graph(n2).lineID).angle
            
            Dim angleDiff As Double
            angleDiff = Abs(line1Angle - line2Angle)
            If angleDiff > PI Then angleDiff = 2 * PI - angleDiff

            If dist <= connectionTolerance And angleDiff > (45 * DEG_TO_RAD) Then
                graph(n1).IsDangling = False
                graph(n2).IsDangling = False
            End If




            If dist <= connectionTolerance Then
                ' Mark both as connected
                graph(n1).IsDangling = False
                graph(n1).ConnectedLineID = graph(n2).lineID
                graph(n2).IsDangling = False
                graph(n2).ConnectedLineID = graph(n1).lineID
            End If

NextConnection:
        Next n2
    Next n1

Debug.Print "    Built connection graph: " & nodeCount & " endpoints"
    BuildEndpointConnectionGraph = nodeCount
End Function

' ==================== BUILD GRID SEGMENTS FOR AXIS ====================
Private Function BuildGridSegmentsForAxis(axisID As Long, ByRef segments() As GridSegment) As Long
    BuildGridSegmentsForAxis = 0

    ' Find all perpendicular axes that intersect this axis
    Dim intersections() As Point2D
    Dim intCount    As Long
    intCount = 0
    ReDim intersections(0 To mAxesCount - 1)

    Dim otherID     As Long
    For otherID = 0 To mAxesCount - 1
        If otherID = axisID Then GoTo NextAxisCheck

        ' Check perpendicular
        Dim angleDiff As Double
        angleDiff = Abs(mAxes(axisID).angle - mAxes(otherID).angle)
        Do While angleDiff > PI
            angleDiff = angleDiff - PI
        Loop

        If Abs(angleDiff - HALF_PI) > (mAngleTolerance * DEG_TO_RAD * 2) Then GoTo NextAxisCheck

        ' Calculate intersection
        Dim intPt   As Point2D
        If GetAxisAxisIntersection(axisID, otherID, intPt) Then
            intersections(intCount) = intPt
            intCount = intCount + 1
        End If

NextAxisCheck:
    Next otherID

    If intCount <= 1 Then Exit Function

    ' Sort intersections by projection along axis
    If intCount > 1 Then
        SortIntersectionsByProjection intersections, intCount, axisID
    End If

    ' Build grid segments between consecutive intersections
    ReDim segments(0 To intCount - 2)

    Dim i           As Long
    For i = 0 To intCount - 2
        segments(i).StartPoint = intersections(i)
        segments(i).EndPoint = intersections(i + 1)
        segments(i).StartProj = ProjectPointOnAxis(intersections(i), axisID)
        segments(i).endProj = ProjectPointOnAxis(intersections(i + 1), axisID)
    Next i

    BuildGridSegmentsForAxis = intCount - 1
End Function

' ==================== EXTEND LINES IN GRID SEGMENT (COMPLETE) ====================
Private Function ExtendLinesInGridSegment(axisID As Long, segment As GridSegment, _
        endpointGraph() As EndpointNode, graphCount As Long) As Long
    ExtendLinesInGridSegment = 0

    Dim snapTolerance As Double
    snapTolerance = 100  ' Lines must be within 100mm of axis to be considered "on axis"

    ' PHASE 1: EXTEND dangling endpoints to grid boundaries
    Dim i           As Long
    For i = 0 To mCenterCount - 1
        If Not mCenterLines(i).IsActive Then GoTo NextLineInSegment

        ' Check if line is parallel to axis
        Dim angleDiff As Double
        angleDiff = Abs(mCenterLines(i).angle - mAxes(axisID).angle)
        If angleDiff > PI Then angleDiff = 2 * PI - angleDiff

        If angleDiff > (mAngleTolerance * DEG_TO_RAD) Then GoTo NextLineInSegment

        ' Check if line midpoint is near axis
        Dim midX As Double, midY As Double
        midX = (mCenterLines(i).startPt.X + mCenterLines(i).endPt.X) / 2
        midY = (mCenterLines(i).startPt.Y + mCenterLines(i).endPt.Y) / 2

        Dim midPt   As Point2D
        midPt.X = midX: midPt.Y = midY

        Dim distToAxis As Double
        distToAxis = PointToAxisDistance(midPt, axisID)

        If distToAxis > snapTolerance Then GoTo NextLineInSegment

        ' Check if line is within grid segment bounds
        Dim lineStartProj As Double, lineEndProj As Double
        lineStartProj = ProjectPointOnAxis(mCenterLines(i).startPt, axisID)
        lineEndProj = ProjectPointOnAxis(mCenterLines(i).endPt, axisID)

        Dim lineMinProj As Double, lineMaxProj As Double
        lineMinProj = Min(lineStartProj, lineEndProj)
        lineMaxProj = Max(lineStartProj, lineEndProj)

        Dim segMinProj As Double, segMaxProj As Double
        segMinProj = Min(segment.StartProj, segment.endProj)
        segMaxProj = Max(segment.StartProj, segment.endProj)

        ' Line must overlap with segment
        If lineMaxProj < segMinProj - snapTolerance Then GoTo NextLineInSegment
        If lineMinProj > segMaxProj + snapTolerance Then GoTo NextLineInSegment

        ' Attempt to extend endpoints to grid boundaries
        Dim extendedStart As Boolean, extendedEnd As Boolean
        extendedStart = TryExtendEndpointToGrid(i, True, segment, axisID, endpointGraph, graphCount)
        extendedEnd = TryExtendEndpointToGrid(i, False, segment, axisID, endpointGraph, graphCount)

        If extendedStart Or extendedEnd Then
            ExtendLinesInGridSegment = ExtendLinesInGridSegment + 1
            mCenterLines(i).uniqueID = GenerateUniqueID(mCenterLines(i))
        End If

NextLineInSegment:
    Next i

    ' PHASE 2: MERGE all extended lines in this grid segment
    Dim mergedCount As Long
    mergedCount = MergeExtendedLinesInGridSegment(axisID, segment)

    If mergedCount > 0 Then
Debug.Print "      Merged " & mergedCount & " extended lines in grid segment"
    End If
End Function

' ==================== TRY EXTEND ENDPOINT TO GRID (FIXED - 30% LENGTH LIMIT) ====================
Private Function TryExtendEndpointToGrid(lineID As Long, isStartPoint As Boolean, _
        segment As GridSegment, axisID As Long, _
        endpointGraph() As EndpointNode, graphCount As Long) As Boolean
    TryExtendEndpointToGrid = False

    ' RULE 1: Check if endpoint is dangling
    If Not IsEndpointDangling(lineID, isStartPoint, endpointGraph, graphCount) Then
        Exit Function
    End If

    ' Get endpoint position
    Dim endPt       As Point2D
    If isStartPoint Then
        endPt = mCenterLines(lineID).startPt
    Else
        endPt = mCenterLines(lineID).endPt
    End If

    ' Calculate projection on the axis
    Dim endProj     As Double
    endProj = ProjectPointOnAxis(endPt, axisID)

    ' Determine which grid boundary is closer (Start or End of the grid segment)
    Dim segMinProj As Double, segMaxProj As Double
    segMinProj = Min(segment.StartProj, segment.endProj)
    segMaxProj = Max(segment.StartProj, segment.endProj)

    Dim targetProj  As Double
    Dim targetPoint As Point2D

    Dim distToMin As Double, distToMax As Double
    distToMin = Abs(endProj - segMinProj)
    distToMax = Abs(endProj - segMaxProj)

    ' Choose closer boundary
    If distToMin < distToMax Then
        targetProj = segMinProj
        targetPoint = segment.StartPoint
    Else
        targetProj = segMaxProj
        targetPoint = segment.EndPoint
    End If

    ' Calculate required extension distance
    Dim extendDist  As Double
    extendDist = Abs(targetProj - endProj)

    If extendDist < 0.1 Then Exit Function  ' Already at boundary

    ' === NEW LOGIC: LIMIT EXTENSION BY 30% OF LINE LENGTH ===
    Dim lineLen     As Double
    lineLen = mCenterLines(lineID).Length

    Dim maxLimit    As Double
    maxLimit = lineLen * 0.3    ' Limit to 30% of the line length

    ' Safety Floor: If line is very short, 30% might be too small to be useful.
    ' We allow at least 200mm extension if the line is short,
    ' provided it doesn't exceed the grid cell size significantly.
    If maxLimit < 300 Then maxLimit = 300

    ' Apply the limit
    If extendDist > maxLimit Then Exit Function

    ' === OBSTRUCTION CHECK (Relaxed) ===
    ' Only check for obstruction if extension is significant (> 100mm)
    If extendDist > 100 Then
        Dim obstruction As Point2D
        Dim obstructionFound As Boolean
        obstructionFound = FindNearestObstruction(endPt, targetPoint, lineID, obstruction)

        If obstructionFound Then
            ' Extend to obstruction point instead of grid
            If isStartPoint Then
                mCenterLines(lineID).startPt = obstruction
            Else
                mCenterLines(lineID).endPt = obstruction
            End If
            GoTo FinalizeUpdate
        End If
    End If

    ' Clear path - extend to grid boundary
    If isStartPoint Then
        mCenterLines(lineID).startPt = targetPoint
    Else
        mCenterLines(lineID).endPt = targetPoint
    End If

FinalizeUpdate:
    ' Recalculate line geometry
    Dim dx As Double, dy As Double
    dx = mCenterLines(lineID).endPt.X - mCenterLines(lineID).startPt.X
    dy = mCenterLines(lineID).endPt.Y - mCenterLines(lineID).startPt.Y
    mCenterLines(lineID).Length = Sqr(dx * dx + dy * dy)
    mCenterLines(lineID).angle = Atan2(dy, dx)

    TryExtendEndpointToGrid = True
End Function

' ==================== CHECK IF ENDPOINT IS DANGLING ====================
Private Function IsEndpointDangling(lineID As Long, isStartPoint As Boolean, _
        endpointGraph() As EndpointNode, graphCount As Long) As Boolean
    IsEndpointDangling = True

    On Error Resume Next
    Dim i           As Long
    For i = 0 To graphCount - 1
        If endpointGraph(i).lineID = lineID And endpointGraph(i).IsStart = isStartPoint Then
            IsEndpointDangling = endpointGraph(i).IsDangling
            Exit Function
        End If
    Next i
    On Error GoTo 0
End Function

' ==================== FIND NEAREST OBSTRUCTION (RAYCAST) ====================
Private Function FindNearestObstruction(startPt As Point2D, endPt As Point2D, _
        excludeLineID As Long, ByRef obstruction As Point2D) As Boolean
    FindNearestObstruction = False

    Dim minDist     As Double
    minDist = 999999

    Dim rayLength   As Double
    rayLength = Sqr((endPt.X - startPt.X) ^ 2 + (endPt.Y - startPt.Y) ^ 2)

    If rayLength < EPSILON Then Exit Function

    Dim i           As Long
    For i = 0 To mCenterCount - 1

        Dim angleDiff As Double
        angleDiff = Abs(mCenterLines(i).angle - mCenterLines(excludeLineID).angle)
        If angleDiff < (5 * DEG_TO_RAD) Then GoTo NextObstruction

        If i = excludeLineID Then GoTo NextObstruction
        If Not mCenterLines(i).IsActive Then GoTo NextObstruction

        ' Check intersection with this line
        Dim intPt   As Point2D
        If RayIntersectsLine(startPt, endPt, i, intPt) Then
            Dim distToInt As Double
            distToInt = Sqr((intPt.X - startPt.X) ^ 2 + (intPt.Y - startPt.Y) ^ 2)

            ' Must be along ray direction and closer than current minimum
            If distToInt > 1 And distToInt < rayLength And distToInt < minDist Then
                obstruction = intPt
                minDist = distToInt
                FindNearestObstruction = True
            End If
        End If

NextObstruction:
    Next i
End Function

' ==================== RAY-LINE INTERSECTION TEST ====================
Private Function RayIntersectsLine(rayStart As Point2D, rayEnd As Point2D, _
        lineID As Long, ByRef intPt As Point2D) As Boolean
    RayIntersectsLine = False

    Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
    Dim x3 As Double, y3 As Double, x4 As Double, y4 As Double

    x1 = rayStart.X: y1 = rayStart.Y
    x2 = rayEnd.X: y2 = rayEnd.Y
    x3 = mCenterLines(lineID).startPt.X: y3 = mCenterLines(lineID).startPt.Y
    x4 = mCenterLines(lineID).endPt.X: y4 = mCenterLines(lineID).endPt.Y

    Dim denom       As Double
    denom = (x1 - x2) * (y3 - y4) - (y1 - y2) * (x3 - x4)

    If Abs(denom) < EPSILON Then Exit Function

    Dim t As Double, u As Double
    t = ((x1 - x3) * (y3 - y4) - (y1 - y3) * (x3 - x4)) / denom
    u = -((x1 - x2) * (y1 - y3) - (y1 - y2) * (x1 - x3)) / denom

    ' Ray parameter t: 0 to 1, Line parameter u: 0 to 1
    If t >= 0 And t <= 1 And u >= 0 And u <= 1 Then
        intPt.X = x1 + t * (x2 - x1)
        intPt.Y = y1 + t * (y2 - y1)
        RayIntersectsLine = True
    End If
End Function

' ==================== MERGE EXTENDED LINES IN GRID SEGMENT (FIXED OVERLAP) ====================
Private Function MergeExtendedLinesInGridSegment(axisID As Long, segment As GridSegment) As Long
    MergeExtendedLinesInGridSegment = 0

    Dim snapTolerance As Double
    snapTolerance = 100  ' Lines must be within 100mm of axis

    ' Collect lines roughly aligned with this grid segment
    Dim lineIndices() As Long
    Dim lineCount   As Long
    lineCount = 0
    ReDim lineIndices(0 To mCenterCount - 1)

    Dim i           As Long
    For i = 0 To mCenterCount - 1
        If Not mCenterLines(i).IsActive Then GoTo NextCollect

        ' Check angle parallel to axis
        Dim angleDiff As Double
        angleDiff = Abs(mCenterLines(i).angle - mAxes(axisID).angle)
        If angleDiff > PI Then angleDiff = 2 * PI - angleDiff
        If angleDiff > (mAngleTolerance * DEG_TO_RAD) Then GoTo NextCollect

        ' Check distance to axis
        Dim midX As Double, midY As Double
        midX = (mCenterLines(i).startPt.X + mCenterLines(i).endPt.X) / 2
        midY = (mCenterLines(i).startPt.Y + mCenterLines(i).endPt.Y) / 2
        Dim midPt   As Point2D: midPt.X = midX: midPt.Y = midY

        If PointToAxisDistance(midPt, axisID) > snapTolerance Then GoTo NextCollect

        ' Check bounds: Line must be relevant to this grid segment
        ' We expand bounds slightly to catch lines extending just outside
        Dim lStart As Double, lEnd As Double
        lStart = ProjectPointOnAxis(mCenterLines(i).startPt, axisID)
        lEnd = ProjectPointOnAxis(mCenterLines(i).endPt, axisID)

        Dim sMin As Double, sMax As Double
        sMin = Min(segment.StartProj, segment.endProj) - snapTolerance
        sMax = Max(segment.StartProj, segment.endProj) + snapTolerance

        ' Check intersection of intervals
        If Max(lStart, lEnd) < sMin Or Min(lStart, lEnd) > sMax Then GoTo NextCollect

        lineIndices(lineCount) = i
        lineCount = lineCount + 1
NextCollect:
    Next i

    If lineCount <= 1 Then Exit Function

    ' Sort lines by projection start to facilitate linear merge
    Dim projections() As ProjectionData
    ReDim projections(0 To lineCount - 1)

    For i = 0 To lineCount - 1
        Dim lineID  As Long: lineID = lineIndices(i)
        projections(i).segmentID = lineID
        projections(i).StartProj = ProjectPointOnAxis(mCenterLines(lineID).startPt, axisID)
        projections(i).endProj = ProjectPointOnAxis(mCenterLines(lineID).endPt, axisID)

        ' Normalize Start < End
        If projections(i).StartProj > projections(i).endProj Then
            Dim tmp As Double
            tmp = projections(i).StartProj
            projections(i).StartProj = projections(i).endProj
            projections(i).endProj = tmp
        End If
    Next i

    QuickSortProjections projections, 0, lineCount - 1

    ' Iterative merge pass
    ' We loop until no more merges occur to handle multi-segment chains A-B-C
    Dim mergedSomething As Boolean
    Do
        mergedSomething = False
        Dim j       As Long

        For i = 0 To lineCount - 1
            Dim id1 As Long: id1 = projections(i).segmentID
            If Not mCenterLines(id1).IsActive Then GoTo NextMerge

            For j = i + 1 To lineCount - 1
                Dim id2 As Long: id2 = projections(j).segmentID
                If Not mCenterLines(id2).IsActive Then GoTo NextInnerMerge

                ' Calculate Gap: Start of Next - End of Current
                Dim gap As Double
                gap = projections(j).StartProj - projections(i).endProj

                ' MERGE CONDITION:
                ' gap <= 50mm allows for:
                ' 1. Slight gaps (up to 50mm)
                ' 2. Touching (gap = 0)
                ' 3. Overlapping (gap < 0) - CRITICAL FIX
                If gap <= 50 Then
                    ' Merge id2 into id1
                    JoinTwoCenterLinesOnAxis id1, id2, mAxes(axisID).angle, mAxes(axisID).startPt
                    mCenterLines(id2).IsActive = False

                    ' Update the active line's projection to cover the new merged extent
                    projections(i).StartProj = Min(projections(i).StartProj, projections(j).StartProj)
                    projections(i).endProj = Max(projections(i).endProj, projections(j).endProj)

                    MergeExtendedLinesInGridSegment = MergeExtendedLinesInGridSegment + 1
                    mergedSomething = True

                    ' Note: We don't exit inner loop; current 'id1' might assume more lines 'id3', etc.
                Else
                    ' Since list is sorted, if gap is large positive, subsequent lines won't match either
                    Exit For
                End If
NextInnerMerge:
            Next j
NextMerge:
        Next i
    Loop While mergedSomething
End Function

' ==================== JOIN TWO CENTERLINES ON AXIS ====================
Private Sub JoinTwoCenterLinesOnAxis(targetID As Long, sourceID As Long, _
        axisAngle As Double, axisBase As Point2D)
    ' Project all 4 endpoints onto axis
    Dim points(0 To 3) As Point2D
    points(0) = mCenterLines(targetID).startPt
    points(1) = mCenterLines(targetID).endPt
    points(2) = mCenterLines(sourceID).startPt
    points(3) = mCenterLines(sourceID).endPt

    Dim cosA As Double, sinA As Double
    cosA = Cos(axisAngle)
    sinA = Sin(axisAngle)

    Dim minProj As Double, maxProj As Double
    minProj = 1E+300: maxProj = -1E+300

    Dim i           As Long
    For i = 0 To 3
        Dim dx As Double, dy As Double
        dx = points(i).X - axisBase.X
        dy = points(i).Y - axisBase.Y

        Dim proj    As Double
        proj = dx * cosA + dy * sinA

        If proj < minProj Then minProj = proj
        If proj > maxProj Then maxProj = proj
    Next i

    ' Reconstruct merged line on axis
    mCenterLines(targetID).startPt.X = axisBase.X + minProj * cosA
    mCenterLines(targetID).startPt.Y = axisBase.Y + minProj * sinA
    mCenterLines(targetID).endPt.X = axisBase.X + maxProj * cosA
    mCenterLines(targetID).endPt.Y = axisBase.Y + maxProj * sinA

    ' Update geometry
    Dim finalDx As Double, finalDy As Double
    finalDx = mCenterLines(targetID).endPt.X - mCenterLines(targetID).startPt.X
    finalDy = mCenterLines(targetID).endPt.Y - mCenterLines(targetID).startPt.Y
    mCenterLines(targetID).Length = Sqr(finalDx * finalDx + finalDy * finalDy)
    mCenterLines(targetID).angle = Atan2(finalDy, finalDx)

    ' PRESERVE PROPERTIES: Merge thickness and WallType
    On Error Resume Next
    Dim tTh As Double, sTh As Double
    tTh = mCenterLines(targetID).Thickness
    sTh = mCenterLines(sourceID).Thickness

    ' Choose larger thickness
    If sTh > tTh Then
        mCenterLines(targetID).Thickness = sTh
    End If

    ' Merge WallType (prefer non-empty)
    Dim tType As String, sType As String
    tType = Trim$(mCenterLines(targetID).WallType)
    sType = Trim$(mCenterLines(sourceID).WallType)

    If Len(tType) = 0 And Len(sType) > 0 Then
        mCenterLines(targetID).WallType = sType
    ElseIf Len(tType) > 0 And Len(sType) > 0 Then
        ' Both have types, choose one with larger thickness
        If sTh > tTh Then
            mCenterLines(targetID).WallType = sType
        End If
    End If

    ' Rebuild WallType from thickness if empty
    If Len(Trim$(mCenterLines(targetID).WallType)) = 0 And mCenterLines(targetID).Thickness > 0 Then
        mCenterLines(targetID).WallType = "W" & CStr(CInt(mCenterLines(targetID).Thickness))
    End If

    ' Update unique ID
    mCenterLines(targetID).uniqueID = GenerateUniqueID(mCenterLines(targetID))
    On Error GoTo 0
End Sub

' ==================== HELPER: PROJECT POINT ON AXIS ====================
Private Function ProjectPointOnAxis(pt As Point2D, axisID As Long) As Double
    Dim dx As Double, dy As Double
    dx = pt.X - mAxes(axisID).startPt.X
    dy = pt.Y - mAxes(axisID).startPt.Y

    Dim cosA As Double, sinA As Double
    cosA = Cos(mAxes(axisID).angle)
    sinA = Sin(mAxes(axisID).angle)

    ProjectPointOnAxis = dx * cosA + dy * sinA
End Function

' ==================== HELPER: SORT INTERSECTIONS BY PROJECTION ====================
Private Sub SortIntersectionsByProjection(ByRef intersections() As Point2D, count As Long, axisID As Long)
    Dim projections() As Double
    ReDim projections(0 To count - 1)

    Dim i           As Long
    For i = 0 To count - 1
        projections(i) = ProjectPointOnAxis(intersections(i), axisID)
    Next i

    ' Bubble sort (simple for small arrays)
    Dim j           As Long
    Dim swapped     As Boolean
    Do
        swapped = False
        For i = 0 To count - 2
            If projections(i) > projections(i + 1) Then
                ' Swap projections
                Dim tmpProj As Double
                tmpProj = projections(i)
                projections(i) = projections(i + 1)
                projections(i + 1) = tmpProj

                ' Swap points
                Dim tmpPt As Point2D
                tmpPt = intersections(i)
                intersections(i) = intersections(i + 1)
                intersections(i + 1) = tmpPt

                swapped = True
            End If
        Next i
    Loop While swapped
End Sub

' Helper: Find axis parallel to line
Private Function FindParallelAxis(lineObj As Object, tolerance As Double) As Long
    FindParallelAxis = -1

    Dim sp As Variant, ep As Variant
    sp = lineObj.StartPoint
    ep = lineObj.EndPoint

    Dim dx As Double, dy As Double
    dx = ep(0) - sp(0)
    dy = ep(1) - sp(1)

    Dim lineAngle   As Double
    lineAngle = Atan2(dy, dx)

    Dim i           As Long
    For i = 0 To mAxesCount - 1
        Dim angleDiff As Double
        angleDiff = Abs(lineAngle - mAxes(i).angle)
        If angleDiff > PI Then angleDiff = 2 * PI - angleDiff

        If angleDiff <= (mAngleTolerance * DEG_TO_RAD) Then
            FindParallelAxis = i
            Exit Function
        End If
    Next i
End Function

' Helper: Extend line to nearest perpendicular axis intersections
Private Function ExtendLineToAxisIntersections(lineObj As Object, axisID As Long, maxExtend As Double) As Boolean
    ExtendLineToAxisIntersections = False

    Dim sp As Variant, ep As Variant
    sp = lineObj.StartPoint
    ep = lineObj.EndPoint

    Dim modified    As Boolean: modified = False

    ' Find perpendicular axes
    Dim i           As Long
    For i = 0 To mAxesCount - 1
        If i = axisID Then GoTo NextAxisCheck

        ' Check if perpendicular
        Dim angleDiff As Double
        angleDiff = Abs(mAxes(axisID).angle - mAxes(i).angle)
        Do While angleDiff > PI
            angleDiff = angleDiff - PI
        Loop

        If Abs(angleDiff - HALF_PI) > (mAngleTolerance * DEG_TO_RAD * 2) Then GoTo NextAxisCheck

        ' Calculate intersection point
        Dim intPt   As Point2D
        If GetAxisAxisIntersection(axisID, i, intPt) Then
            ' Check if intersection is near start point
            Dim distToStart As Double
            distToStart = Sqr((sp(0) - intPt.X) ^ 2 + (sp(1) - intPt.Y) ^ 2)

            If distToStart > 0.1 And distToStart <= maxExtend Then
                sp(0) = intPt.X: sp(1) = intPt.Y
                modified = True
            End If

            ' Check if intersection is near end point
            Dim distToEnd As Double
            distToEnd = Sqr((ep(0) - intPt.X) ^ 2 + (ep(1) - intPt.Y) ^ 2)

            If distToEnd > 0.1 And distToEnd <= maxExtend Then
                ep(0) = intPt.X: ep(1) = intPt.Y
                modified = True
            End If
        End If

NextAxisCheck:
    Next i

    If modified Then
        lineObj.StartPoint = sp
        lineObj.EndPoint = ep
        ExtendLineToAxisIntersections = True
    End If
End Function

' Helper: Get intersection between 2 axes
Private Function GetAxisAxisIntersection(axis1ID As Long, axis2ID As Long, ByRef intPt As Point2D) As Boolean
    GetAxisAxisIntersection = False

    Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
    Dim x3 As Double, y3 As Double, x4 As Double, y4 As Double

    x1 = mAxes(axis1ID).startPt.X: y1 = mAxes(axis1ID).startPt.Y
    x2 = mAxes(axis1ID).endPt.X: y2 = mAxes(axis1ID).endPt.Y
    x3 = mAxes(axis2ID).startPt.X: y3 = mAxes(axis2ID).startPt.Y
    x4 = mAxes(axis2ID).endPt.X: y4 = mAxes(axis2ID).endPt.Y

    Dim denom       As Double
    denom = (x1 - x2) * (y3 - y4) - (y1 - y2) * (x3 - x4)

    If Abs(denom) < EPSILON Then Exit Function

    Dim t           As Double
    t = ((x1 - x3) * (y3 - y4) - (y1 - y3) * (x3 - x4)) / denom

    intPt.X = x1 + t * (x2 - x1)
    intPt.Y = y1 + t * (y2 - y1)
    GetAxisAxisIntersection = True
End Function

' Helper: Collect lines from layer
Private Function CollectLinesFromLayer(layerName As String) As Collection
    Set CollectLinesFromLayer = New Collection

    Dim ms          As Object
    Set ms = mAcadDoc.ModelSpace

    Dim ent         As Object
    For Each ent In ms
        On Error Resume Next
        If LCase(ent.ObjectName) = "acdbline" Then
            If LCase(Trim(ent.layer)) = LCase(layerName) Then
                CollectLinesFromLayer.Add ent
            End If
        End If
        On Error GoTo 0
    Next ent
End Function


' ==================== BUILD KD TREE (from original) ====================
Private Sub BuildKDTree()
    On Error Resume Next
    If mSegmentCount = 0 Then Exit Sub

    ReDim mKDTree(0 To mSegmentCount * 2 - 1)
    mKDNodeCount = 0

    Dim indices()   As Long
    ReDim indices(0 To mSegmentCount - 1)

    Dim i           As Long
    For i = 0 To mSegmentCount - 1
        indices(i) = i
    Next i

    BuildKDTreeRecursive indices, 0, mSegmentCount - 1, 0
    On Error GoTo 0
End Sub

Private Function BuildKDTreeRecursive(indices() As Long, Left As Long, Right As Long, Depth As Long) As Long
    If Left > Right Then
        BuildKDTreeRecursive = -1
        Exit Function
    End If

    Dim splitDim    As Integer
    splitDim = Depth Mod 2

    QuickSortIndicesByDim indices, Left, Right, splitDim

    Dim median      As Long
    median = (Left + Right) \ 2

    Dim nodeID      As Long
    nodeID = mKDNodeCount
    mKDNodeCount = mKDNodeCount + 1

    mKDTree(nodeID).segmentID = indices(median)
    mKDTree(nodeID).splitDim = splitDim

    If splitDim = 0 Then
        mKDTree(nodeID).splitValue = mSegments(indices(median)).startPt.X
    Else
        mKDTree(nodeID).splitValue = mSegments(indices(median)).startPt.Y
    End If

    mKDTree(nodeID).LeftChild = BuildKDTreeRecursive(indices, Left, median - 1, Depth + 1)
    mKDTree(nodeID).RightChild = BuildKDTreeRecursive(indices, median + 1, Right, Depth + 1)

    BuildKDTreeRecursive = nodeID
End Function

Private Sub QuickSortIndicesByDim(ByRef arr() As Long, Left As Long, Right As Long, dimSplit As Integer)
    If Left >= Right Then Exit Sub

    Dim pivot       As Double
    pivot = GetSegmentCoord(arr((Left + Right) \ 2), dimSplit)

    Dim i As Long, j As Long
    i = Left: j = Right

    Do While i <= j
        Do While GetSegmentCoord(arr(i), dimSplit) < pivot
            i = i + 1
        Loop
        Do While GetSegmentCoord(arr(j), dimSplit) > pivot
            j = j - 1
        Loop
        If i <= j Then
            Dim tmp As Long
            tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
            i = i + 1: j = j - 1
        End If
    Loop

    If Left < j Then QuickSortIndicesByDim arr, Left, j, dimSplit
    If i < Right Then QuickSortIndicesByDim arr, i, Right, dimSplit
End Sub

Private Function GetSegmentCoord(segID As Long, dimSplit As Integer) As Double
    If dimSplit = 0 Then
        GetSegmentCoord = mSegments(segID).startPt.X
    Else
        GetSegmentCoord = mSegments(segID).startPt.Y
    End If
End Function

' ==================== NEW: USE AUTOCAD OVERKILL INSTEAD OF CUSTOM MERGE ====================
Private Function ApplyAutoCADOverkill() As Long
    On Error GoTo ErrHandler
    ApplyAutoCADOverkill = 0

Debug.Print "  Applying AutoCAD OVERKILL to centerlines..."

    ' Create selection set of all DTS_WALL_DIAGRAM objects
    Dim ssName      As String: ssName = "OVERKILL_" & Format(Now, "hhmmss")
    On Error Resume Next
    mAcadDoc.SelectionSets.item(ssName).Delete
    On Error GoTo ErrHandler

    Dim ss          As Object
    Set ss = mAcadDoc.SelectionSets.Add(ssName)

    ' Filter: Only lines on DTS_WALL_DIAGRAM layer
    Dim gpCode(0 To 1) As Integer
    Dim dataVal(0 To 1) As Variant
    gpCode(0) = 0: dataVal(0) = "LINE"
    gpCode(1) = 8: dataVal(1) = "DTS_WALL_DIAGRAM"

    ss.Select 5, , , gpCode, dataVal  ' 5 = acSelectionSetAll

    Dim countBefore As Long
    countBefore = ss.count
    ss.Delete    ' Clean up selection set

    If countBefore = 0 Then Exit Function

    ' Run OVERKILL command
    ' -OVERKILL = non-dialog version
    ' Options: Combine co-linear, Ignore layer, 0.001 tolerance
    Dim cmd         As String
    cmd = "_-OVERKILL" & vbCr & _
            "(ssget ""X"" '((8 . ""DTS_WALL_DIAGRAM"")))" & vbCr & vbCr & _
            "_Tolerance" & vbCr & "1" & vbCr & _
            "_Done" & vbCr
    mAcadDoc.SendCommand cmd

    ' Wait for command completion
    Dim i           As Integer
    For i = 1 To 20    ' Wait up to 2 seconds
        DoEvents
        ' A crude wait mechanism, referencing the application object usually helps sync
        mAcadApp.Update
    Next i

    ' 5. Count objects after
    Set ss = mAcadDoc.SelectionSets.Add(ssName & "_AFTER")
    ss.Select 5, , , gpCode, dataVal
    Dim countAfter  As Long
    countAfter = ss.count
    ss.Delete

    ApplyAutoCADOverkill = countBefore - countAfter
Debug.Print "  OVERKILL: Removed " & ApplyAutoCADOverkill & " objects"

    Exit Function
ErrHandler:
Debug.Print "ERROR in ApplyAutoCADOverkill: " & err.description
    On Error Resume Next
    If Not ss Is Nothing Then ss.Delete
End Function

' ==================== DETECT WALL PAIRS ====================
Private Function DetectWallPairs() As Long
    DetectWallPairs = 0

    ReDim mWallPairs(0 To mSegmentCount - 1)
    mPairCount = 0

    Dim i           As Long
    For i = 0 To mSegmentCount - 1
        If mSegments(i).PairSegmentID >= 0 Then GoTo NextSeg

        Dim searchRadius As Double
        searchRadius = mWallThicknesses(mWallThicknessCount - 1) + mDistanceTolerance

        Dim nearbyIDs() As Long, nearbyCount As Long
        nearbyCount = SearchNearbySegments(i, searchRadius, nearbyIDs)

        Dim j       As Long
        For j = 0 To nearbyCount - 1
            Dim otherID As Long
            otherID = nearbyIDs(j)

            If otherID = i Or mSegments(otherID).PairSegmentID >= 0 Then GoTo NextNearby

            If AreSegmentsParallel(i, otherID) Then
                Dim dist As Double
                dist = CalculateSegmentDistance(i, otherID)

                Dim k As Long
                For k = 0 To mWallThicknessCount - 1
                    If Abs(dist - mWallThicknesses(k)) <= mDistanceTolerance Then
                        mSegments(i).PairSegmentID = otherID
                        mSegments(otherID).PairSegmentID = i

                        mWallPairs(mPairCount).Seg1ID = i
                        mWallPairs(mPairCount).Seg2ID = otherID
                        mWallPairs(mPairCount).Thickness = mWallThicknesses(k)
                        mWallPairs(mPairCount).IsValid = True
                        mPairCount = mPairCount + 1

                        DetectWallPairs = DetectWallPairs + 1
                        Exit For
                    End If
                Next k
            End If
NextNearby:
        Next j
NextSeg:
    Next i

    If mPairCount > 0 Then ReDim Preserve mWallPairs(0 To mPairCount - 1)

    ' Mark unpaired segments as single lines
    For i = 0 To mSegmentCount - 1
        If mSegments(i).PairSegmentID < 0 Then
            mSegments(i).IsSingleLine = True
        End If
    Next i
End Function

Private Function SearchNearbySegments(segmentID As Long, radius As Double, ByRef resultIDs() As Long) As Long
    Dim tempIDs()   As Long
    ReDim tempIDs(0 To mSegmentCount - 1)
    Dim count       As Long
    count = 0

    Dim BBox        As BoundingBox
    BBox.minX = Min(mSegments(segmentID).startPt.X, mSegments(segmentID).endPt.X) - radius
    BBox.maxX = Max(mSegments(segmentID).startPt.X, mSegments(segmentID).endPt.X) + radius
    BBox.minY = Min(mSegments(segmentID).startPt.Y, mSegments(segmentID).endPt.Y) - radius
    BBox.maxY = Max(mSegments(segmentID).startPt.Y, mSegments(segmentID).endPt.Y) + radius

    If mKDNodeCount > 0 Then
        SearchKDTreeRange 0, BBox, tempIDs, count
    End If

    If count > 0 Then
        ReDim resultIDs(0 To count - 1)
        Dim i       As Long
        For i = 0 To count - 1
            resultIDs(i) = tempIDs(i)
        Next i
    Else
        ReDim resultIDs(0 To -1)
    End If

    SearchNearbySegments = count
End Function

Private Sub SearchKDTreeRange(nodeID As Long, BBox As BoundingBox, ByRef resultIDs() As Long, ByRef count As Long)
    If nodeID < 0 Or nodeID >= mKDNodeCount Then Exit Sub

    Dim segID       As Long
    segID = mKDTree(nodeID).segmentID

    If mSegments(segID).startPt.X >= BBox.minX And mSegments(segID).startPt.X <= BBox.maxX And _
            mSegments(segID).startPt.Y >= BBox.minY And mSegments(segID).startPt.Y <= BBox.maxY Then
        resultIDs(count) = segID
        count = count + 1
    End If

    Dim splitDim As Integer, splitValue As Double
    splitDim = mKDTree(nodeID).splitDim
    splitValue = mKDTree(nodeID).splitValue

    If splitDim = 0 Then
        If BBox.minX <= splitValue Then SearchKDTreeRange mKDTree(nodeID).LeftChild, BBox, resultIDs, count
        If BBox.maxX >= splitValue Then SearchKDTreeRange mKDTree(nodeID).RightChild, BBox, resultIDs, count
    Else
        If BBox.minY <= splitValue Then SearchKDTreeRange mKDTree(nodeID).LeftChild, BBox, resultIDs, count
        If BBox.maxY >= splitValue Then SearchKDTreeRange mKDTree(nodeID).RightChild, BBox, resultIDs, count
    End If
End Sub

Private Function AreSegmentsParallel(id1 As Long, id2 As Long) As Boolean
    Dim angleDiff   As Double
    angleDiff = Abs(mSegments(id1).angle - mSegments(id2).angle)
    If angleDiff > PI Then angleDiff = 2 * PI - angleDiff
    AreSegmentsParallel = (angleDiff <= (mAngleTolerance * DEG_TO_RAD))
End Function

Private Function CalculateSegmentDistance(id1 As Long, id2 As Long) As Double
    Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double, x3 As Double, y3 As Double
    x1 = mSegments(id1).startPt.X: y1 = mSegments(id1).startPt.Y
    x2 = mSegments(id2).startPt.X: y2 = mSegments(id2).startPt.Y
    x3 = mSegments(id2).endPt.X: y3 = mSegments(id2).endPt.Y

    Dim dx As Double, dy As Double, dlen As Double
    dx = x3 - x2: dy = y3 - y2
    dlen = Sqr(dx * dx + dy * dy)

    If dlen < EPSILON Then
        CalculateSegmentDistance = PointDistance2D(mSegments(id1).startPt, mSegments(id2).startPt)
    Else
        CalculateSegmentDistance = Abs(dy * (x1 - x2) - dx * (y1 - y2)) / dlen
    End If
End Function

' ==================== GENERATE CENTERLINES ====================
Private Sub GenerateCenterLinesFromPairs()
    ReDim mCenterLines(0 To mPairCount - 1)
    mCenterCount = 0

    Dim i           As Long
    For i = 0 To mPairCount - 1
        If Not mWallPairs(i).IsValid Then GoTo NextPair

        Dim seg1 As Long, seg2 As Long
        seg1 = mWallPairs(i).Seg1ID
        seg2 = mWallPairs(i).Seg2ID

        mCenterLines(mCenterCount).startPt.X = (mSegments(seg1).startPt.X + mSegments(seg2).startPt.X) / 2
        mCenterLines(mCenterCount).startPt.Y = (mSegments(seg1).startPt.Y + mSegments(seg2).startPt.Y) / 2
        mCenterLines(mCenterCount).endPt.X = (mSegments(seg1).endPt.X + mSegments(seg2).endPt.X) / 2
        mCenterLines(mCenterCount).endPt.Y = (mSegments(seg1).endPt.Y + mSegments(seg2).endPt.Y) / 2

        Dim dx As Double, dy As Double
        dx = mCenterLines(mCenterCount).endPt.X - mCenterLines(mCenterCount).startPt.X
        dy = mCenterLines(mCenterCount).endPt.Y - mCenterLines(mCenterCount).startPt.Y
        mCenterLines(mCenterCount).Length = Sqr(dx * dx + dy * dy)
        mCenterLines(mCenterCount).angle = Atan2(dy, dx)
        mCenterLines(mCenterCount).SourcePairID = i
        mCenterLines(mCenterCount).VectorID = -1
        mCenterLines(mCenterCount).IsActive = True
        mCenterLines(mCenterCount).uniqueID = GenerateUniqueID(mCenterLines(mCenterCount))

        ' Transfer thickness and wall type from detected pair
        mCenterLines(mCenterCount).Thickness = mWallPairs(i).Thickness
        If mCenterLines(mCenterCount).Thickness > 0 Then
            mCenterLines(mCenterCount).WallType = "W" & CStr(CInt(mCenterLines(mCenterCount).Thickness))
        Else
            mCenterLines(mCenterCount).WallType = ""
        End If

        mCenterCount = mCenterCount + 1

NextPair:
    Next i

    If mCenterCount > 0 Then ReDim Preserve mCenterLines(0 To mCenterCount - 1)
End Sub

Private Function AddSingleLinesToCenterLines() As Long
    AddSingleLinesToCenterLines = 0

    Dim singleCount As Long
    singleCount = 0

    ' Count single lines
    Dim i           As Long
    For i = 0 To mSegmentCount - 1
        If mSegments(i).IsSingleLine Then
            singleCount = singleCount + 1
        End If
    Next i

    If singleCount = 0 Then Exit Function

    ' Expand centerlines array
    Dim oldCount    As Long
    oldCount = mCenterCount

    ReDim Preserve mCenterLines(0 To mCenterCount + singleCount - 1)

    ' Add single lines
    For i = 0 To mSegmentCount - 1
        If mSegments(i).IsSingleLine Then
            mCenterLines(mCenterCount).startPt = mSegments(i).startPt
            mCenterLines(mCenterCount).endPt = mSegments(i).endPt
            mCenterLines(mCenterCount).Length = mSegments(i).Length
            mCenterLines(mCenterCount).angle = mSegments(i).angle
            mCenterLines(mCenterCount).SourcePairID = -1  ' No pair source
            mCenterLines(mCenterCount).VectorID = -1
            mCenterLines(mCenterCount).IsActive = True
            mCenterLines(mCenterCount).uniqueID = GenerateUniqueID(mCenterLines(mCenterCount))

            ' Transfer thickness from segment XData if present
            If mSegments(i).Thickness > 0 Then
                mCenterLines(mCenterCount).Thickness = mSegments(i).Thickness
                mCenterLines(mCenterCount).WallType = "W" & CStr(CInt(mSegments(i).Thickness))
            Else
                ' default/fallback (use smallest known thickness if available)
                If mWallThicknessCount > 0 Then
                    mCenterLines(mCenterCount).Thickness = mWallThicknesses(0)
                    mCenterLines(mCenterCount).WallType = "W" & CStr(CInt(mWallThicknesses(0)))
                Else
                    mCenterLines(mCenterCount).Thickness = 0
                    mCenterLines(mCenterCount).WallType = ""
                End If
            End If

            mCenterCount = mCenterCount + 1
            AddSingleLinesToCenterLines = AddSingleLinesToCenterLines + 1
        End If
    Next i
End Function


' ==================== GENERATE UNIQUE ID FOR CENTERLINE ====================
Private Function GenerateUniqueID(cl As CenterLine) As String
    Dim sx As String, sy As String, ex As String, ey As String
    Dim thick       As String

    sx = Format(cl.startPt.X, "0.000")
    sy = Format(cl.startPt.Y, "0.000")
    ex = Format(cl.endPt.X, "0.000")
    ey = Format(cl.endPt.Y, "0.000")
    thick = Format(cl.Thickness, "0")   ' mm, integer

    ' Add thickness + angle into key to avoid duplication after extend
    GenerateUniqueID = sx & "_" & sy & "_" & ex & "_" & ey & _
            "_T" & thick & "_A" & Format(cl.angle, "0.000")
End Function
' ==================== GLOBAL GAP RECOVERY (NO VECTOR GROUPING) ====================
Private Function RecoverGapsGlobal() As Long
    RecoverGapsGlobal = 0

    If mCenterCount <= 1 Then Exit Function
    If mDoorWidthCount = 0 And mColumnWidthCount = 0 Then Exit Function

Debug.Print "  Starting global gap recovery (geometry-based)..."

    ' Collect all break widths
    Dim allBreakWidths() As Double
    Dim breakCount  As Long
    breakCount = mDoorWidthCount + mColumnWidthCount

    If breakCount = 0 Then Exit Function

    ReDim allBreakWidths(0 To breakCount - 1)
    Dim idx As Long, k As Long
    idx = 0

    For k = 0 To mDoorWidthCount - 1
        allBreakWidths(idx) = mDoorWidths(k)
        idx = idx + 1
    Next k

    For k = 0 To mColumnWidthCount - 1
        allBreakWidths(idx) = mColumnWidths(k)
        idx = idx + 1
    Next k

    QuickSortDoubles allBreakWidths, 0, breakCount - 1

    ' ITERATE EXHAUSTIVELY (no vector grouping)
    Dim maxIterations As Long
    maxIterations = 10
    Dim iteration   As Long
    iteration = 0

    Do While iteration < maxIterations
        iteration = iteration + 1

        Dim joined  As Long
        joined = 0

        ' BRUTE FORCE: Check all pairs
        Dim i As Long, j As Long
        For i = 0 To mCenterCount - 1
            If Not mCenterLines(i).IsActive Then GoTo NextGlobalI

            For j = i + 1 To mCenterCount - 1
                If Not mCenterLines(j).IsActive Then GoTo NextGlobalJ

                ' GEOMETRY CHECK (ignore vector grouping)
                If Not AreGeometricallyCollinear(i, j, 1, 50) Then GoTo NextGlobalJ

                ' Calculate gap distance
                Dim gapDist As Double
                gapDist = CalculateGapDistance(i, j)

                If gapDist < 0 Then GoTo NextGlobalJ  ' Overlapping (handle later)

                ' Check if gap matches any break width
                Dim w As Long
                For w = 0 To breakCount - 1
                    Dim tolerance As Double
                    tolerance = Max(mDistanceTolerance * 3, allBreakWidths(w) * 0.05)

                    If Abs(gapDist - allBreakWidths(w)) <= tolerance Then
                        ' MATCH FOUND - Validate safety before joining
                        If IsConnectionSafe(i, j, gapDist) Then
                            JoinTwoCenterLines i, j
                            mCenterLines(j).IsActive = False
                            mCenterLines(i).uniqueID = GenerateUniqueID(mCenterLines(i))

                            joined = joined + 1
                            RecoverGapsGlobal = RecoverGapsGlobal + 1
                        Else
Debug.Print "    WARNING: Rejected unsafe connection between lines " & i & " and " & j
                        End If

                        Exit For  ' Move to next j after finding match
                    End If
                Next w

NextGlobalJ:
            Next j
NextGlobalI:
        Next i

Debug.Print "    Global gap recovery iteration " & iteration & ": joined " & joined

        If joined = 0 Then Exit Do  ' Converged
    Loop

    ' Compact array
    If RecoverGapsGlobal > 0 Then
        CompactCenterLineArray
    End If
End Function
' ==================== GLOBAL MERGE (GEOMETRY-BASED) ====================
Private Function MergeOverlappingCenterLinesGlobal() As Long
    MergeOverlappingCenterLinesGlobal = 0

    If mCenterCount <= 1 Then Exit Function

Debug.Print "  Starting global merge (geometry-based)..."

    Dim maxIterations As Long
    maxIterations = 5
    Dim iteration   As Long
    iteration = 0

    Do While iteration < maxIterations
        iteration = iteration + 1

        Dim merged  As Long
        merged = 0

        ' BRUTE FORCE: Check all pairs
        Dim i As Long, j As Long
        For i = 0 To mCenterCount - 1
            If Not mCenterLines(i).IsActive Then GoTo NextGlobalMergeI

            For j = i + 1 To mCenterCount - 1
                If Not mCenterLines(j).IsActive Then GoTo NextGlobalMergeJ

                ' GEOMETRY CHECK
                If Not AreGeometricallyCollinear(i, j, 1, 50) Then GoTo NextGlobalMergeJ

                ' Calculate overlap using projection
                Dim overlapResult As overlapResult
                overlapResult = CalculateOverlapScalar(i, j)

                If overlapResult.HasOverlap Then
                    ' MERGE CONDITION 1: Line ng?n n?m tr?n trong line dài
                    '    (overlap >= 80% c?a line ng?n)
                    If overlapResult.OverlapPercent >= 0.8 Then
                        If mCenterLines(i).Length >= mCenterLines(j).Length Then
                            ' i is longer, absorb j
                            mCenterLines(j).IsActive = False
                        Else
                            ' j is longer, absorb i
                            mCenterLines(i).IsActive = False
                        End If

                        merged = merged + 1
                        MergeOverlappingCenterLinesGlobal = MergeOverlappingCenterLinesGlobal + 1
                        GoTo NextGlobalMergeJ
                    End If

                    ' MERGE CONDITION 2: Partial overlap >= 30%
                    If overlapResult.OverlapPercent >= 0.3 Then
                        JoinTwoCenterLines i, j
                        mCenterLines(j).IsActive = False
                        mCenterLines(i).uniqueID = GenerateUniqueID(mCenterLines(i))

                        merged = merged + 1
                        MergeOverlappingCenterLinesGlobal = MergeOverlappingCenterLinesGlobal + 1
                    End If
                End If

NextGlobalMergeJ:
            Next j
NextGlobalMergeI:
        Next i

Debug.Print "    Global merge iteration " & iteration & ": merged " & merged

        If merged = 0 Then Exit Do
    Loop

    If MergeOverlappingCenterLinesGlobal > 0 Then
        CompactCenterLineArray
    End If
End Function
' ==================== STEP 3: ENHANCED EXHAUSTIVE EXACT RECOVERY ====================
Private Function RecoverWallLengthExactExhaustive() As Long
    RecoverWallLengthExactExhaustive = 0

    If mCenterCount <= 1 Then Exit Function
    If mDoorWidthCount = 0 And mColumnWidthCount = 0 Then Exit Function

    ' Collect all break widths
    Dim allBreakWidths() As Double
    Dim breakCount  As Long
    breakCount = mDoorWidthCount + mColumnWidthCount

    If breakCount = 0 Then Exit Function

    ReDim allBreakWidths(0 To breakCount - 1)
    Dim idx As Long, k As Long
    idx = 0

    For k = 0 To mDoorWidthCount - 1
        allBreakWidths(idx) = mDoorWidths(k)
        idx = idx + 1
    Next k

    For k = 0 To mColumnWidthCount - 1
        allBreakWidths(idx) = mColumnWidths(k)
        idx = idx + 1
    Next k

    QuickSortDoubles allBreakWidths, 0, breakCount - 1

    ' EXHAUSTIVE: Keep iterating until no more matches found
    Dim totalJoined As Long
    totalJoined = 0
    Dim maxIterations As Long
    maxIterations = 10
    Dim iteration   As Long
    iteration = 0

    Do While iteration < maxIterations
        iteration = iteration + 1

        ' Regroup centerlines each iteration
        GroupCenterLinesByVector

        Dim joined  As Long
        joined = 0

        Dim v       As Long
        For v = 0 To mVectorCount - 1
            Dim vectorJoined As Long
            vectorJoined = JoinCenterLinesInVectorExactExhaustive(v, allBreakWidths, breakCount)
            joined = joined + vectorJoined
        Next v

        If joined = 0 Then Exit Do  ' No more matches

        totalJoined = totalJoined + joined
Debug.Print "    Exact recovery iteration " & iteration & ": joined " & joined & " segments"
    Loop

    RecoverWallLengthExactExhaustive = totalJoined
End Function

Private Function JoinCenterLinesInVectorExactExhaustive(VectorID As Long, breakWidths() As Double, breakCount As Long) As Long
    JoinCenterLinesInVectorExactExhaustive = 0

    Dim lineCount   As Long
    lineCount = mVectorLines(VectorID).segmentCount

    If lineCount <= 1 Then Exit Function

    ' Project and sort
    Dim projections() As ProjectionData
    ReDim projections(0 To lineCount - 1)

    Dim i As Long, lineID As Long
    For i = 0 To lineCount - 1
        lineID = mVectorLines(VectorID).SegmentIDs(i)

        If Not mCenterLines(lineID).IsActive Then GoTo NextProj

        projections(i).segmentID = lineID
        projections(i).StartProj = ProjectPointOnVectorCL(mCenterLines(lineID).startPt, VectorID)
        projections(i).endProj = ProjectPointOnVectorCL(mCenterLines(lineID).endPt, VectorID)

        If projections(i).StartProj > projections(i).endProj Then
            Dim tmp As Double
            tmp = projections(i).StartProj
            projections(i).StartProj = projections(i).endProj
            projections(i).endProj = tmp
        End If

NextProj:
    Next i

    QuickSortProjections projections, 0, lineCount - 1

    ' Try matching with 2% tolerance
    Dim j As Long, w As Long
    i = 0

    Do While i < lineCount - 1
        lineID = projections(i).segmentID
        If Not mCenterLines(lineID).IsActive Then
            i = i + 1
            GoTo NextJoinLoop
        End If

        Dim matched As Boolean
        matched = False

        For j = i + 1 To lineCount - 1
            Dim otherID As Long
            otherID = projections(j).segmentID

            If Not mCenterLines(otherID).IsActive Then GoTo NextOther

            Dim gap As Double
            gap = projections(j).StartProj - projections(i).endProj

            If gap < -mDistanceTolerance Then GoTo NextOther

            ' Allow 2% tolerance for break widths
            For w = 0 To breakCount - 1
                Dim tolerance As Double
                tolerance = Max(mDistanceTolerance * 2, breakWidths(w) * 0.02)

                If Abs(gap - breakWidths(w)) <= tolerance Then
                    ' Match found - join and update
                    JoinTwoCenterLines lineID, otherID
                    mCenterLines(otherID).IsActive = False

                    projections(i).endProj = projections(j).endProj
                    mCenterLines(lineID).uniqueID = GenerateUniqueID(mCenterLines(lineID))

                    JoinCenterLinesInVectorExactExhaustive = JoinCenterLinesInVectorExactExhaustive + 1
                    matched = True

                    ' Restart from beginning to catch cascading joins
                    i = -1
                    Exit For
                End If
            Next w

            If matched Then Exit For

NextOther:
        Next j

        i = i + 1
NextJoinLoop:
    Loop
End Function

' ==================== STEP 4: ENHANCED EXHAUSTIVE GAP RECOVERY ====================
Private Function RecoverWallLengthByGapExhaustive() As Long
    RecoverWallLengthByGapExhaustive = 0

    If mCenterCount <= 1 Or mAutoJoinGapDistance <= 0 Then Exit Function

    Dim totalJoined As Long
    totalJoined = 0
    Dim maxIterations As Long
    maxIterations = 10
    Dim iteration   As Long
    iteration = 0

    Do While iteration < maxIterations
        iteration = iteration + 1

        GroupCenterLinesByVector

        Dim joined  As Long
        joined = 0

        Dim v       As Long
        For v = 0 To mVectorCount - 1
            Dim vectorJoined As Long
            vectorJoined = JoinCenterLinesInVectorByGapExhaustive(v)
            joined = joined + vectorJoined
        Next v

        If joined = 0 Then Exit Do

        totalJoined = totalJoined + joined
Debug.Print "    Gap join iteration " & iteration & ": joined " & joined & " segments"
    Loop

    RecoverWallLengthByGapExhaustive = totalJoined
End Function

Private Function JoinCenterLinesInVectorByGapExhaustive(VectorID As Long) As Long
    JoinCenterLinesInVectorByGapExhaustive = 0

    Dim lineCount   As Long
    lineCount = mVectorLines(VectorID).segmentCount

    If lineCount <= 1 Then Exit Function

    Dim projections() As ProjectionData
    ReDim projections(0 To lineCount - 1)

    Dim i As Long, lineID As Long
    For i = 0 To lineCount - 1
        lineID = mVectorLines(VectorID).SegmentIDs(i)

        If Not mCenterLines(lineID).IsActive Then GoTo NextProjGap

        projections(i).segmentID = lineID
        projections(i).StartProj = ProjectPointOnVectorCL(mCenterLines(lineID).startPt, VectorID)
        projections(i).endProj = ProjectPointOnVectorCL(mCenterLines(lineID).endPt, VectorID)

        If projections(i).StartProj > projections(i).endProj Then
            Dim tmp As Double
            tmp = projections(i).StartProj
            projections(i).StartProj = projections(i).endProj
            projections(i).endProj = tmp
        End If

NextProjGap:
    Next i

    QuickSortProjections projections, 0, lineCount - 1

    Dim j           As Long
    i = 0

    Do While i < lineCount - 1
        lineID = projections(i).segmentID
        If Not mCenterLines(lineID).IsActive Then
            i = i + 1
            GoTo NextGapLoop
        End If

        For j = i + 1 To lineCount - 1
            Dim otherID As Long
            otherID = projections(j).segmentID

            If Not mCenterLines(otherID).IsActive Then GoTo NextOtherGap

            Dim gap As Double
            gap = projections(j).StartProj - projections(i).endProj

            If gap < -mDistanceTolerance Then GoTo NextOtherGap

            If gap > mAutoJoinGapDistance Then Exit For

            ' Join them
            JoinTwoCenterLines lineID, otherID
            mCenterLines(otherID).IsActive = False

            projections(i).endProj = projections(j).endProj
            mCenterLines(lineID).uniqueID = GenerateUniqueID(mCenterLines(lineID))

            JoinCenterLinesInVectorByGapExhaustive = JoinCenterLinesInVectorByGapExhaustive + 1

            ' Restart to catch cascading
            i = -1
            Exit For

NextOtherGap:
        Next j

        i = i + 1
NextGapLoop:
    Loop
End Function
' ==================== BUILD CORNER NODE GRAPH ====================
Private Sub BuildCornerNodeGraph()
    On Error Resume Next

    If mCenterCount = 0 Then Exit Sub

    ' Reset
    ReDim mCornerNodes(0 To mCenterCount * 2 - 1)
    mCornerNodeCount = 0

    Dim snapTolerance As Double
    snapTolerance = 100  ' 100mm tolerance for corner detection

Debug.Print "  Building corner node graph..."

    ' Collect all endpoints
    Dim i           As Long
    For i = 0 To mCenterCount - 1
        If Not mCenterLines(i).IsActive Then GoTo NextEndpoint

        ' Add start point
        AddOrMergeCornerNode mCenterLines(i).startPt.X, mCenterLines(i).startPt.Y, i, snapTolerance

        ' Add end point
        AddOrMergeCornerNode mCenterLines(i).endPt.X, mCenterLines(i).endPt.Y, i, snapTolerance

NextEndpoint:
    Next i

    If mCornerNodeCount > 0 Then
        ReDim Preserve mCornerNodes(0 To mCornerNodeCount - 1)
    End If

Debug.Print "  Found " & mCornerNodeCount & " corner nodes"

    On Error GoTo 0
End Sub

' ==================== ADD OR MERGE CORNER NODE ====================
Private Sub AddOrMergeCornerNode(X As Double, Y As Double, lineID As Long, tolerance As Double)
    ' Check if this point already exists
    Dim i           As Long
    For i = 0 To mCornerNodeCount - 1
        Dim dist    As Double
        dist = Sqr((mCornerNodes(i).PointX - X) ^ 2 + (mCornerNodes(i).PointY - Y) ^ 2)

        If dist <= tolerance Then
            ' Merge into existing node
            Dim idx As Long
            idx = mCornerNodes(i).ConnectedCount

            ' Expand array if needed
            If idx >= UBound(mCornerNodes(i).ConnectedLineIDs) Then
                ReDim Preserve mCornerNodes(i).ConnectedLineIDs(0 To idx + 10)
            End If

            mCornerNodes(i).ConnectedLineIDs(idx) = lineID
            mCornerNodes(i).ConnectedCount = idx + 1
            Exit Sub
        End If
    Next i

    ' Create new node
    mCornerNodes(mCornerNodeCount).PointX = X
    mCornerNodes(mCornerNodeCount).PointY = Y
    ReDim mCornerNodes(mCornerNodeCount).ConnectedLineIDs(0 To 10)
    mCornerNodes(mCornerNodeCount).ConnectedLineIDs(0) = lineID
    mCornerNodes(mCornerNodeCount).ConnectedCount = 1
    mCornerNodes(mCornerNodeCount).IsProcessed = False

    mCornerNodeCount = mCornerNodeCount + 1
End Sub
' ==================== RECOVER CORNER GAPS (STRICT ORTHOGONAL SHIFT & EXTEND) ====================
Private Function RecoverCornerGaps() As Long
    RecoverCornerGaps = 0
    If mCornerNodeCount = 0 Then Exit Function

    ' B1: Ép ph?ng toàn b? line tru?c khi x? lý góc d? d?m b?o d? li?u s?ch
    FlattenAllLinesStrict

Debug.Print "  Starting corner gap recovery (Manhattan Logic)..."

    Dim i           As Long
    For i = 0 To mCornerNodeCount - 1
        If mCornerNodes(i).ConnectedCount <> 2 Then GoTo NextCorner

        Dim id1 As Long, id2 As Long
        id1 = mCornerNodes(i).ConnectedLineIDs(0)
        id2 = mCornerNodes(i).ConnectedLineIDs(1)

        ' Validate
        If id1 < 0 Or id1 >= mCenterCount Or id2 < 0 Or id2 >= mCenterCount Then GoTo NextCorner
        If Not mCenterLines(id1).IsActive Or Not mCenterLines(id2).IsActive Then GoTo NextCorner

        ' Ki?m tra vuông góc
        If Not AreLinesPerpendicular(id1, id2) Then GoTo NextCorner

        ' --- XÁC Ð?NH HU?NG (Horizontal / Vertical) ---
        Dim hID As Long, vID As Long
        Dim is1Horz As Boolean, is2Horz As Boolean

        ' Ki?m tra l?i hu?ng sau khi dã Flatten
        is1Horz = (Abs(mCenterLines(id1).startPt.Y - mCenterLines(id1).endPt.Y) < EPSILON)
        is2Horz = (Abs(mCenterLines(id2).startPt.Y - mCenterLines(id2).endPt.Y) < EPSILON)

        If is1Horz And Not is2Horz Then
            hID = id1: vID = id2
        ElseIf Not is1Horz And is2Horz Then
            hID = id2: vID = id1
        Else
            ' Cùng hu?ng (song song) => B? qua
            GoTo NextCorner
        End If

        ' --- TÍNH TOÁN T?A Ð? GIAO C?T LÝ TU?NG (MANHATTAN INTERSECTION) ---
        ' Giao di?m c?a 2 du?ng vuông góc vô h?n chính là (X c?a du?ng D?c, Y c?a du?ng Ngang)
        Dim targetX As Double, targetY As Double
        targetX = mCenterLines(vID).startPt.X    ' Vì du?ng d?c dã Flatten nên X start = X end
        targetY = mCenterLines(hID).startPt.Y    ' Vì du?ng ngang dã Flatten nên Y start = Y end

        ' --- KI?M TRA KHO?NG CÁCH AN TOÀN ---
        ' Tìm d?u mút g?n nh?t c?a m?i du?ng d?n giao di?m (targetX, targetY)
        Dim distH As Double, distV As Double
        distH = Min(Abs(mCenterLines(hID).startPt.X - targetX), Abs(mCenterLines(hID).endPt.X - targetX))
        distV = Min(Abs(mCenterLines(vID).startPt.Y - targetY), Abs(mCenterLines(vID).endPt.Y - targetY))

        ' N?u giao di?m quá xa (> 500mm) => Có th? là 2 tu?ng không liên quan => B? qua
        If distH > 500 Or distV > 500 Then GoTo NextCorner

        ' --- QUY T?C VÀNG: CH? KÉO DÀI D?C TR?C (EXTEND LONGITUDINALLY) ---
        ' 1. X? lý du?ng Ngang (hID): Ch? thay d?i t?a d? X d? ch?m targetX. GI? NGUYÊN Y.
        ExtendHorizontalLineToX hID, targetX

        ' 2. X? lý du?ng D?c (vID): Ch? thay d?i t?a d? Y d? ch?m targetY. GI? NGUYÊN X.
        ExtendVerticalLineToY vID, targetY

        ' Luu ý: N?u tu?ng b? l?ch (Offset 100mm), code này s? KHÔNG b? cong tu?ng d? n?i.
        ' Nó s? t?o ra hình ch? L vuông góc nhung các d?u mút có th? không ch?m nhau t?i 1 di?m pixel
        ' (chúng s? ch?m nhau v? m?t toán h?c trên lu?i tr?c).
        ' Ðây là cách duy nh?t d? tránh du?ng chéo.

        mCenterLines(id1).uniqueID = GenerateUniqueID(mCenterLines(id1))
        mCenterLines(id2).uniqueID = GenerateUniqueID(mCenterLines(id2))

        RecoverCornerGaps = RecoverCornerGaps + 1

NextCorner:
    Next i

Debug.Print "    Recovered " & RecoverCornerGaps & " corners (Manhattan Strict)."
End Function

' ==================== HELPER: EXTEND STRICT ====================

Private Sub ExtendHorizontalLineToX(lineID As Long, xVal As Double)
    ' Tìm d?u mút nào g?n xVal nh?t thì kéo d?u dó
    Dim dStart As Double, dEnd As Double
    dStart = Abs(mCenterLines(lineID).startPt.X - xVal)
    dEnd = Abs(mCenterLines(lineID).endPt.X - xVal)

    If dStart < dEnd Then
        mCenterLines(lineID).startPt.X = xVal
    Else
        mCenterLines(lineID).endPt.X = xVal
    End If

    ' Recalc Length only
    mCenterLines(lineID).Length = Abs(mCenterLines(lineID).endPt.X - mCenterLines(lineID).startPt.X)
    ' Angle is implicitly 0 or PI because we didn't touch Y
End Sub

Private Sub ExtendVerticalLineToY(lineID As Long, yVal As Double)
    ' Tìm d?u mút nào g?n yVal nh?t thì kéo d?u dó
    Dim dStart As Double, dEnd As Double
    dStart = Abs(mCenterLines(lineID).startPt.Y - yVal)
    dEnd = Abs(mCenterLines(lineID).endPt.Y - yVal)

    If dStart < dEnd Then
        mCenterLines(lineID).startPt.Y = yVal
    Else
        mCenterLines(lineID).endPt.Y = yVal
    End If

    ' Recalc Length only
    mCenterLines(lineID).Length = Abs(mCenterLines(lineID).endPt.Y - mCenterLines(lineID).startPt.Y)
    ' Angle is implicitly PI/2 or 3PI/2 because we didn't touch X
End Sub

' ==================== HELPER: STRICT EXTEND FUNCTIONS ====================

' Extend/Trim a horizontal line to a specific X coordinate. Y is UNTOUCHED.
Private Sub ExtendLineToX(lineID As Long, targetX As Double)
    Dim dStart As Double, dEnd As Double
    dStart = Abs(mCenterLines(lineID).startPt.X - targetX)
    dEnd = Abs(mCenterLines(lineID).endPt.X - targetX)

    If dStart < dEnd Then
        mCenterLines(lineID).startPt.X = targetX
    Else
        mCenterLines(lineID).endPt.X = targetX
    End If

    ' Recalculate Length
    mCenterLines(lineID).Length = Abs(mCenterLines(lineID).endPt.X - mCenterLines(lineID).startPt.X)
    ' Force Angle to 0 or PI
    If mCenterLines(lineID).endPt.X >= mCenterLines(lineID).startPt.X Then
        mCenterLines(lineID).angle = 0
    Else
        mCenterLines(lineID).angle = 3.14159265358979
    End If
End Sub

' Extend/Trim a vertical line to a specific Y coordinate. X is UNTOUCHED.
Private Sub ExtendLineToY(lineID As Long, targetY As Double)
    Dim dStart As Double, dEnd As Double
    dStart = Abs(mCenterLines(lineID).startPt.Y - targetY)
    dEnd = Abs(mCenterLines(lineID).endPt.Y - targetY)

    If dStart < dEnd Then
        mCenterLines(lineID).startPt.Y = targetY
    Else
        mCenterLines(lineID).endPt.Y = targetY
    End If

    ' Recalculate Length
    mCenterLines(lineID).Length = Abs(mCenterLines(lineID).endPt.Y - mCenterLines(lineID).startPt.Y)
    ' Force Angle to PI/2 or 3PI/2
    If mCenterLines(lineID).endPt.Y >= mCenterLines(lineID).startPt.Y Then
        mCenterLines(lineID).angle = 1.5707963267949
    Else
        mCenterLines(lineID).angle = 4.71238898038469
    End If
End Sub

' Check if line is strictly horizontal (with small tolerance for input data)
Private Function IsLineHorizontal(lineID As Long) As Boolean
    Dim dx As Double, dy As Double
    dx = Abs(mCenterLines(lineID).endPt.X - mCenterLines(lineID).startPt.X)
    dy = Abs(mCenterLines(lineID).endPt.Y - mCenterLines(lineID).startPt.Y)
    IsLineHorizontal = (dx >= dy)
End Function

Private Function GetNearestEndpoint(lineID As Long, refPt As Point2D) As Point2D
    Dim dStart As Double, dEnd As Double
    dStart = (mCenterLines(lineID).startPt.X - refPt.X) ^ 2 + (mCenterLines(lineID).startPt.Y - refPt.Y) ^ 2
    dEnd = (mCenterLines(lineID).endPt.X - refPt.X) ^ 2 + (mCenterLines(lineID).endPt.Y - refPt.Y) ^ 2

    If dStart < dEnd Then
        GetNearestEndpoint = mCenterLines(lineID).startPt
    Else
        GetNearestEndpoint = mCenterLines(lineID).endPt
    End If
End Function

' ==================== HELPER: CALCULATE INFINITE LINE INTERSECTION ====================
' Calculates the intersection point of two lines treated as infinite vectors.
' Uses the general line equation Ax + By = C to avoid division by zero on vertical lines.
Private Function GetIntersectionInfinite(line1 As CenterLine, line2 As CenterLine, ByRef result As Point2D) As Boolean
    Dim A1 As Double, B1 As Double, C1 As Double
    Dim A2 As Double, B2 As Double, C2 As Double

    ' Line 1 equation: A1x + B1y = C1
    A1 = line1.endPt.Y - line1.startPt.Y
    B1 = line1.startPt.X - line1.endPt.X
    C1 = A1 * line1.startPt.X + B1 * line1.startPt.Y

    ' Line 2 equation: A2x + B2y = C2
    A2 = line2.endPt.Y - line2.startPt.Y
    B2 = line2.startPt.X - line2.endPt.X
    C2 = A2 * line2.startPt.X + B2 * line2.startPt.Y

    Dim det         As Double
    det = A1 * B2 - A2 * B1

    If Abs(det) < 0.0000001 Then
        GetIntersectionInfinite = False    ' Lines are parallel
    Else
        result.X = (B2 * C1 - B1 * C2) / det
        result.Y = (A1 * C2 - A2 * C1) / det
        GetIntersectionInfinite = True
    End If
End Function

' ==================== HELPER: UPDATE ENDPOINT & ENFORCE ORTHOGONALITY ====================
' Moves the nearest endpoint of a line to a specific point.
' CRITICAL: Forces the line to remain strictly vertical or horizontal to prevent skewing.
Private Sub UpdateLineEndpointToPoint(lineID As Long, pt As Point2D)
    Dim dStart      As Double
    Dim dEnd        As Double

    ' Determine which endpoint is closer to the target point
    dStart = (mCenterLines(lineID).startPt.X - pt.X) ^ 2 + (mCenterLines(lineID).startPt.Y - pt.Y) ^ 2
    dEnd = (mCenterLines(lineID).endPt.X - pt.X) ^ 2 + (mCenterLines(lineID).endPt.Y - pt.Y) ^ 2

    If dStart < dEnd Then
        mCenterLines(lineID).startPt = pt
    Else
        mCenterLines(lineID).endPt = pt
    End If

    ' Recalculate Length and Angle based on new endpoints
    Dim dx As Double, dy As Double
    dx = mCenterLines(lineID).endPt.X - mCenterLines(lineID).startPt.X
    dy = mCenterLines(lineID).endPt.Y - mCenterLines(lineID).startPt.Y

    ' STRICT ORTHOGONALITY CHECK
    ' This block fixes the "26 degree skew" issue by forcing the line to snap
    ' to the nearest axis (X or Y) even if the intersection point had a micro-deviation.
    If Abs(dx) > Abs(dy) Then
        ' Line is dominant Horizontal
        mCenterLines(lineID).startPt.Y = pt.Y    ' Force Y to match target (flatten Y)
        mCenterLines(lineID).endPt.Y = pt.Y

        ' Recalc Angle strictly
        If dx > 0 Then
            mCenterLines(lineID).angle = 0
        Else
            mCenterLines(lineID).angle = 3.14159265358979    ' PI
        End If

        ' Recalc Length using X only
        mCenterLines(lineID).Length = Abs(mCenterLines(lineID).endPt.X - mCenterLines(lineID).startPt.X)
    Else
        ' Line is dominant Vertical
        mCenterLines(lineID).startPt.X = pt.X    ' Force X to match target (flatten X)
        mCenterLines(lineID).endPt.X = pt.X

        ' Recalc Angle strictly
        If dy > 0 Then
            mCenterLines(lineID).angle = 1.5707963267949    ' HALF_PI
        Else
            mCenterLines(lineID).angle = 4.71238898038469    ' 3 * HALF_PI
        End If

        ' Recalc Length using Y only
        mCenterLines(lineID).Length = Abs(mCenterLines(lineID).endPt.Y - mCenterLines(lineID).startPt.Y)
    End If
End Sub
' ==================== CALCULATE CORNER GAP DISTANCE ====================
Private Function CalculateCornerGapDistance(line1ID As Long, line2ID As Long, _
        corner As Point2D, line1AtStart As Boolean, line2AtStart As Boolean) As Double

    ' Get the endpoints at corner
    Dim p1 As Point2D, p2 As Point2D
    If line1AtStart Then
        p1 = mCenterLines(line1ID).startPt
    Else
        p1 = mCenterLines(line1ID).endPt
    End If

    If line2AtStart Then
        p2 = mCenterLines(line2ID).startPt
    Else
        p2 = mCenterLines(line2ID).endPt
    End If

    ' Calculate distance from each line endpoint to corner
    Dim dist1 As Double, dist2 As Double
    dist1 = PointDistance2D(p1, corner)
    dist2 = PointDistance2D(p2, corner)

    ' Gap distance is the sum of both distances
    ' (representing the "missing" material at the corner)
    CalculateCornerGapDistance = dist1 + dist2
End Function

' ==================== EXTEND LINES TO CORNER ====================
Private Function ExtendLinesToCorner(line1ID As Long, line2ID As Long, _
        corner As Point2D, line1AtStart As Boolean, line2AtStart As Boolean) As Boolean

    ExtendLinesToCorner = False

    On Error Resume Next

    ' Extend line1 to corner
    If line1AtStart Then
        mCenterLines(line1ID).startPt = corner
    Else
        mCenterLines(line1ID).endPt = corner
    End If

    ' Extend line2 to corner
    If line2AtStart Then
        mCenterLines(line2ID).startPt = corner
    Else
        mCenterLines(line2ID).endPt = corner
    End If

    ' Recalculate lengths and angles
    Dim dx As Double, dy As Double

    dx = mCenterLines(line1ID).endPt.X - mCenterLines(line1ID).startPt.X
    dy = mCenterLines(line1ID).endPt.Y - mCenterLines(line1ID).startPt.Y
    mCenterLines(line1ID).Length = Sqr(dx * dx + dy * dy)
    mCenterLines(line1ID).angle = Atan2(dy, dx)
    mCenterLines(line1ID).uniqueID = GenerateUniqueID(mCenterLines(line1ID))

    dx = mCenterLines(line2ID).endPt.X - mCenterLines(line2ID).startPt.X
    dy = mCenterLines(line2ID).endPt.Y - mCenterLines(line2ID).startPt.Y
    mCenterLines(line2ID).Length = Sqr(dx * dx + dy * dy)
    mCenterLines(line2ID).angle = Atan2(dy, dx)
    mCenterLines(line2ID).uniqueID = GenerateUniqueID(mCenterLines(line2ID))

    ExtendLinesToCorner = True

    On Error GoTo 0
End Function
Private Sub FlattenAllLinesStrict()
    Dim i           As Long
    Dim dx As Double, dy As Double
    Dim avg         As Double

    For i = 0 To mCenterCount - 1
        If mCenterLines(i).IsActive Then
            dx = Abs(mCenterLines(i).endPt.X - mCenterLines(i).startPt.X)
            dy = Abs(mCenterLines(i).endPt.Y - mCenterLines(i).startPt.Y)

            ' N?u dx > dy => Tu?ng ngang. N?u dy > dx => Tu?ng d?c.
            If dx >= dy Then
                ' Force Horizontal: Gán Y c?a c? 2 d?u b?ng trung bình c?ng Y
                avg = (mCenterLines(i).startPt.Y + mCenterLines(i).endPt.Y) / 2
                mCenterLines(i).startPt.Y = avg
                mCenterLines(i).endPt.Y = avg
                mCenterLines(i).angle = IIf(mCenterLines(i).endPt.X >= mCenterLines(i).startPt.X, 0, 3.14159265358979)
            Else
                ' Force Vertical: Gán X c?a c? 2 d?u b?ng trung bình c?ng X
                avg = (mCenterLines(i).startPt.X + mCenterLines(i).endPt.X) / 2
                mCenterLines(i).startPt.X = avg
                mCenterLines(i).endPt.X = avg
                mCenterLines(i).angle = IIf(mCenterLines(i).endPt.Y >= mCenterLines(i).startPt.Y, 1.5707963267949, 4.71238898038469)
            End If

            ' Recalc Length
            dx = mCenterLines(i).endPt.X - mCenterLines(i).startPt.X
            dy = mCenterLines(i).endPt.Y - mCenterLines(i).startPt.Y
            mCenterLines(i).Length = Sqr(dx * dx + dy * dy)
        End If
    Next i
End Sub

' ==================== STEP 5: ENHANCED EXHAUSTIVE AUTO-EXTEND ====================
Private Function AutoExtendCenterLinesExhaustive() As Long
    AutoExtendCenterLinesExhaustive = 0

    If mCenterCount <= 1 Then Exit Function

    Dim avgThickness As Double
    avgThickness = 200

    If mWallThicknessCount > 0 Then
        Dim sum     As Double
        sum = 0
        Dim k       As Long
        For k = 0 To mWallThicknessCount - 1
            sum = sum + mWallThicknesses(k)
        Next k
        avgThickness = sum / mWallThicknessCount
    End If

    Dim maxExtend   As Double
    maxExtend = avgThickness * mExtendMultiplier

    Dim totalExtended As Long
    totalExtended = 0
    Dim maxIterations As Long
    maxIterations = 5
    Dim iteration   As Long
    iteration = 0

    Do While iteration < maxIterations
        iteration = iteration + 1

        Dim extended As Long
        extended = 0

        Dim i As Long, j As Long
        For i = 0 To mCenterCount - 1
            If Not mCenterLines(i).IsActive Then GoTo NextExtendI

            For j = 0 To mCenterCount - 1
                If i = j Or Not mCenterLines(j).IsActive Then GoTo NextExtendJ

                If AreLinesPerpendicular(i, j) Then
                    If TryExtendCenterLineToLine(i, j, maxExtend) Then
                        mCenterLines(i).uniqueID = GenerateUniqueID(mCenterLines(i))
                        extended = extended + 1
                    End If
                End If

NextExtendJ:
            Next j
NextExtendI:
        Next i

        If extended = 0 Then Exit Do

        totalExtended = totalExtended + extended
Debug.Print "    Auto-extend iteration " & iteration & ": extended " & extended & " lines"
    Loop

    AutoExtendCenterLinesExhaustive = totalExtended
End Function

' ==================== CENTERLINE HELPER FUNCTIONS ====================
Private Sub GroupCenterLinesByVector()
    ReDim mVectorLines(0 To mCenterCount - 1)
    mVectorCount = 0

    Dim i           As Long
    For i = 0 To mCenterCount - 1
        If Not mCenterLines(i).IsActive Then GoTo NextCenter
        If mCenterLines(i).VectorID >= 0 Then GoTo NextCenter

        mVectorLines(mVectorCount).angle = mCenterLines(i).angle
        mVectorLines(mVectorCount).BasePoint = mCenterLines(i).startPt
        ReDim mVectorLines(mVectorCount).SegmentIDs(0 To mCenterCount - 1)
        mVectorLines(mVectorCount).segmentCount = 0

        mCenterLines(i).VectorID = mVectorCount
        mVectorLines(mVectorCount).SegmentIDs(0) = i
        mVectorLines(mVectorCount).segmentCount = 1

        Dim j       As Long
        For j = i + 1 To mCenterCount - 1
            If Not mCenterLines(j).IsActive Then GoTo NextJ2
            If mCenterLines(j).VectorID >= 0 Then GoTo NextJ2

            If AreCenterLinesCollinear(i, j) Then
                mCenterLines(j).VectorID = mVectorCount
                mVectorLines(mVectorCount).SegmentIDs(mVectorLines(mVectorCount).segmentCount) = j
                mVectorLines(mVectorCount).segmentCount = mVectorLines(mVectorCount).segmentCount + 1
            End If
NextJ2:
        Next j

        If mVectorLines(mVectorCount).segmentCount > 0 Then
            ReDim Preserve mVectorLines(mVectorCount).SegmentIDs(0 To mVectorLines(mVectorCount).segmentCount - 1)
            mVectorCount = mVectorCount + 1
        End If

NextCenter:
    Next i

    If mVectorCount > 0 Then ReDim Preserve mVectorLines(0 To mVectorCount - 1)

    ' Reset vectorID for next grouping
    For i = 0 To mCenterCount - 1
        mCenterLines(i).VectorID = -1
    Next i

    ' Reassign vectorID
    For i = 0 To mVectorCount - 1
        For j = 0 To mVectorLines(i).segmentCount - 1
            mCenterLines(mVectorLines(i).SegmentIDs(j)).VectorID = i
        Next j
    Next i
End Sub

Private Function AreCenterLinesCollinear(id1 As Long, id2 As Long) As Boolean
    Dim angleDiff   As Double
    angleDiff = Abs(mCenterLines(id1).angle - mCenterLines(id2).angle)
    If angleDiff > PI Then angleDiff = 2 * PI - angleDiff

    If angleDiff > (mAngleTolerance * DEG_TO_RAD) Then
        AreCenterLinesCollinear = False
        Exit Function
    End If

    Dim dist1 As Double, dist2 As Double
    dist1 = PointToLineDistanceInfiniteCL(mCenterLines(id2).startPt, id1)
    dist2 = PointToLineDistanceInfiniteCL(mCenterLines(id2).endPt, id1)

    AreCenterLinesCollinear = (dist1 <= mDistanceTolerance And dist2 <= mDistanceTolerance)
End Function

Private Function PointToLineDistanceInfiniteCL(pt As Point2D, lineID As Long) As Double
    Dim dx As Double, dy As Double
    dx = mCenterLines(lineID).endPt.X - mCenterLines(lineID).startPt.X
    dy = mCenterLines(lineID).endPt.Y - mCenterLines(lineID).startPt.Y

    Dim dlen        As Double
    dlen = Sqr(dx * dx + dy * dy)

    If dlen < EPSILON Then
        PointToLineDistanceInfiniteCL = PointDistance2D(pt, mCenterLines(lineID).startPt)
    Else
        PointToLineDistanceInfiniteCL = Abs(dy * (pt.X - mCenterLines(lineID).startPt.X) - dx * (pt.Y - mCenterLines(lineID).startPt.Y)) / dlen
    End If
End Function

Private Function ProjectPointOnVectorCL(pt As Point2D, VectorID As Long) As Double
    Dim basePt      As Point2D
    basePt = mVectorLines(VectorID).BasePoint

    Dim angle       As Double
    angle = mVectorLines(VectorID).angle

    Dim dx As Double, dy As Double
    dx = pt.X - basePt.X
    dy = pt.Y - basePt.Y

    ProjectPointOnVectorCL = dx * Cos(angle) + dy * Sin(angle)
End Function

Private Sub JoinTwoCenterLines(targetID As Long, sourceID As Long)
    ' FIX: Use projection logic to enforce orthogonality

    ' 1. Identify the dominant line (longer one determines the angle/axis)
    Dim dominantID  As Long
    If mCenterLines(targetID).Length >= mCenterLines(sourceID).Length Then
        dominantID = targetID
    Else
        dominantID = sourceID
    End If

    Dim baseAngle   As Double
    Dim refPt       As Point2D
    baseAngle = mCenterLines(dominantID).angle
    refPt = mCenterLines(dominantID).startPt

    Dim cosA As Double, sinA As Double
    cosA = Cos(baseAngle)
    sinA = Sin(baseAngle)

    ' 2. Collect all 4 points
    Dim points(0 To 3) As Point2D
    points(0) = mCenterLines(targetID).startPt
    points(1) = mCenterLines(targetID).endPt
    points(2) = mCenterLines(sourceID).startPt
    points(3) = mCenterLines(sourceID).endPt

    ' 3. Project points onto the dominant axis
    Dim minProj As Double, maxProj As Double
    minProj = 1E+300: maxProj = -1E+300

    Dim i           As Long
    For i = 0 To 3
        Dim dx As Double, dy As Double
        dx = points(i).X - refPt.X
        dy = points(i).Y - refPt.Y

        Dim proj    As Double
        proj = dx * cosA + dy * sinA

        If proj < minProj Then minProj = proj
        If proj > maxProj Then maxProj = proj
    Next i

    ' 4. Update geometry of targetID to match the full extent on the dominant axis
    mCenterLines(targetID).startPt.X = refPt.X + minProj * cosA
    mCenterLines(targetID).startPt.Y = refPt.Y + minProj * sinA
    mCenterLines(targetID).endPt.X = refPt.X + maxProj * cosA
    mCenterLines(targetID).endPt.Y = refPt.Y + maxProj * sinA

    mCenterLines(targetID).Length = maxProj - minProj
    mCenterLines(targetID).angle = baseAngle    ' Enforce orthogonality

    ' Preserve / merge properties: thickness and wallType
    On Error Resume Next
    Dim tTh As Double, sTh As Double
    tTh = mCenterLines(targetID).Thickness
    sTh = mCenterLines(sourceID).Thickness

    ' choose larger thickness if available
    If sTh > tTh Then
        mCenterLines(targetID).Thickness = sTh
    Else
        mCenterLines(targetID).Thickness = tTh
    End If

    ' wallType selection:
    Dim tgtType As String, srcType As String
    tgtType = Trim$(mCenterLines(targetID).WallType)
    srcType = Trim$(mCenterLines(sourceID).WallType)

    If Len(tgtType) = 0 And Len(srcType) > 0 Then
        mCenterLines(targetID).WallType = srcType
    ElseIf Len(tgtType) > 0 And Len(srcType) > 0 Then
        ' if both present but thickness indicates preference, pick the one with larger thickness
        If sTh > tTh Then
            mCenterLines(targetID).WallType = srcType
        Else
            ' keep tgtType
            mCenterLines(targetID).WallType = tgtType
        End If
    ElseIf Len(tgtType) = 0 And mCenterLines(targetID).Thickness > 0 Then
        ' fallback to building type from thickness
        mCenterLines(targetID).WallType = "W" & CStr(CInt(mCenterLines(targetID).Thickness))
    End If

    ' Avoid W0: if thickness <=0, clear wallType
    If mCenterLines(targetID).Thickness <= 0 Then
        If Left$(Trim$(mCenterLines(targetID).WallType), 1) = "W" Then
            ' if it's of form W<number> but thickness 0 => clear
            Dim numPart As String
            numPart = mid$(mCenterLines(targetID).WallType, 2)
            If Len(Trim$(numPart)) = 0 Or val(numPart) = 0 Then
                mCenterLines(targetID).WallType = ""
            End If
        End If
    End If

    ' Preserve SourcePairID if target lacks it
    If mCenterLines(targetID).SourcePairID < 0 And mCenterLines(sourceID).SourcePairID >= 0 Then
        mCenterLines(targetID).SourcePairID = mCenterLines(sourceID).SourcePairID
    End If

    ' Update unique ID after geometry/property changes
    mCenterLines(targetID).uniqueID = GenerateUniqueID(mCenterLines(targetID))

    On Error GoTo 0
End Sub
Private Function AreLinesPerpendicular(line1ID As Long, line2ID As Long) As Boolean
    Dim angleDiff   As Double
    angleDiff = Abs(mCenterLines(line1ID).angle - mCenterLines(line2ID).angle)

    Do While angleDiff > PI
        angleDiff = angleDiff - PI
    Loop

    AreLinesPerpendicular = (Abs(angleDiff - HALF_PI) <= (mAngleTolerance * DEG_TO_RAD * 2))
End Function

Private Function TryExtendCenterLineToLine(lineID As Long, targetID As Long, maxDist As Double) As Boolean
    TryExtendCenterLineToLine = False

    Dim distStart   As Double
    distStart = PointToLineDistanceCL(mCenterLines(lineID).startPt, targetID)

    If distStart > 0 And distStart <= maxDist Then
        Dim projStart As Point2D
        If ProjectPointToLineCL(mCenterLines(lineID).startPt, targetID, projStart) Then
            mCenterLines(lineID).startPt = projStart
            TryExtendCenterLineToLine = True
        End If
    End If

    Dim distEnd     As Double
    distEnd = PointToLineDistanceCL(mCenterLines(lineID).endPt, targetID)

    If distEnd > 0 And distEnd <= maxDist Then
        Dim projEnd As Point2D
        If ProjectPointToLineCL(mCenterLines(lineID).endPt, targetID, projEnd) Then
            mCenterLines(lineID).endPt = projEnd
            TryExtendCenterLineToLine = True
        End If
    End If

    If TryExtendCenterLineToLine Then
        Dim dx As Double, dy As Double
        dx = mCenterLines(lineID).endPt.X - mCenterLines(lineID).startPt.X
        dy = mCenterLines(lineID).endPt.Y - mCenterLines(lineID).startPt.Y
        mCenterLines(lineID).Length = Sqr(dx * dx + dy * dy)
        mCenterLines(lineID).angle = Atan2(dy, dx)
    End If
End Function

Private Function PointToLineDistanceCL(pt As Point2D, lineID As Long) As Double
    Dim dx As Double, dy As Double
    dx = mCenterLines(lineID).endPt.X - mCenterLines(lineID).startPt.X
    dy = mCenterLines(lineID).endPt.Y - mCenterLines(lineID).startPt.Y

    Dim dlen        As Double
    dlen = Sqr(dx * dx + dy * dy)

    If dlen < EPSILON Then
        PointToLineDistanceCL = PointDistance2D(pt, mCenterLines(lineID).startPt)
    Else
        PointToLineDistanceCL = Abs(dy * (pt.X - mCenterLines(lineID).startPt.X) - dx * (pt.Y - mCenterLines(lineID).startPt.Y)) / dlen
    End If
End Function

Private Function ProjectPointToLineCL(pt As Point2D, lineID As Long, ByRef result As Point2D) As Boolean
    Dim dx As Double, dy As Double
    dx = mCenterLines(lineID).endPt.X - mCenterLines(lineID).startPt.X
    dy = mCenterLines(lineID).endPt.Y - mCenterLines(lineID).startPt.Y

    Dim len2        As Double
    len2 = dx * dx + dy * dy

    If len2 < EPSILON Then
        ProjectPointToLineCL = False
        Exit Function
    End If

    Dim t           As Double
    t = ((pt.X - mCenterLines(lineID).startPt.X) * dx + (pt.Y - mCenterLines(lineID).startPt.Y) * dy) / len2

    If t < -0.1 Or t > 1.1 Then
        ProjectPointToLineCL = False
        Exit Function
    End If

    result.X = mCenterLines(lineID).startPt.X + t * dx
    result.Y = mCenterLines(lineID).startPt.Y + t * dy
    ProjectPointToLineCL = True
End Function

' ==================== SNAP TO AXES (from original) ====================
Private Function SnapCenterLinesToAxes() As Long
    SnapCenterLinesToAxes = 0

    If mAxesCount = 0 Or mCenterCount = 0 Or mAxisSnapDistance <= 0 Then Exit Function

    Dim i As Long, j As Long
    For i = 0 To mCenterCount - 1
        If Not mCenterLines(i).IsActive Then GoTo NextSnap

        For j = 0 To mAxesCount - 1
            If AreCenterLineAndAxisParallel(i, j) Then
                Dim dist As Double
                dist = CalculateCenterLineToAxisDistance(i, j)

                If dist > 0 And dist <= mAxisSnapDistance Then
                    SnapCenterLineToAxis i, j
                    mCenterLines(i).uniqueID = GenerateUniqueID(mCenterLines(i))
                    SnapCenterLinesToAxes = SnapCenterLinesToAxes + 1
                    Exit For
                End If
            Else
                If SnapCenterLineEndpointsToAxis(i, j) Then
                    mCenterLines(i).uniqueID = GenerateUniqueID(mCenterLines(i))
                    SnapCenterLinesToAxes = SnapCenterLinesToAxes + 1
                End If
            End If
        Next j

NextSnap:
    Next i
End Function

Private Function AreCenterLineAndAxisParallel(lineID As Long, axisID As Long) As Boolean
    Dim angleDiff   As Double
    angleDiff = Abs(mCenterLines(lineID).angle - mAxes(axisID).angle)
    If angleDiff > PI Then angleDiff = 2 * PI - angleDiff

    AreCenterLineAndAxisParallel = (angleDiff <= (mAngleTolerance * DEG_TO_RAD))
End Function

Private Function CalculateCenterLineToAxisDistance(lineID As Long, axisID As Long) As Double
    Dim midX As Double, midY As Double
    midX = (mCenterLines(lineID).startPt.X + mCenterLines(lineID).endPt.X) / 2
    midY = (mCenterLines(lineID).startPt.Y + mCenterLines(lineID).endPt.Y) / 2

    Dim midPt       As Point2D
    midPt.X = midX
    midPt.Y = midY

    CalculateCenterLineToAxisDistance = PointToAxisDistance(midPt, axisID)
End Function

Private Function PointToAxisDistance(pt As Point2D, axisID As Long) As Double
    Dim dx As Double, dy As Double
    dx = mAxes(axisID).endPt.X - mAxes(axisID).startPt.X
    dy = mAxes(axisID).endPt.Y - mAxes(axisID).startPt.Y

    Dim dlen        As Double
    dlen = Sqr(dx * dx + dy * dy)

    If dlen < EPSILON Then
        PointToAxisDistance = PointDistance2D(pt, mAxes(axisID).startPt)
    Else
        PointToAxisDistance = Abs(dy * (pt.X - mAxes(axisID).startPt.X) - dx * (pt.Y - mAxes(axisID).startPt.Y)) / dlen
    End If
End Function

Private Sub SnapCenterLineToAxis(lineID As Long, axisID As Long)
    Dim newStart As Point2D, newEnd As Point2D

    ProjectPointToAxis mCenterLines(lineID).startPt, axisID, newStart
    ProjectPointToAxis mCenterLines(lineID).endPt, axisID, newEnd

    mCenterLines(lineID).startPt = newStart
    mCenterLines(lineID).endPt = newEnd

    Dim dx As Double, dy As Double
    dx = mCenterLines(lineID).endPt.X - mCenterLines(lineID).startPt.X
    dy = mCenterLines(lineID).endPt.Y - mCenterLines(lineID).startPt.Y
    mCenterLines(lineID).Length = Sqr(dx * dx + dy * dy)
    mCenterLines(lineID).angle = Atan2(dy, dx)
End Sub

Private Sub ProjectPointToAxis(pt As Point2D, axisID As Long, ByRef result As Point2D)
    Dim dx As Double, dy As Double
    dx = mAxes(axisID).endPt.X - mAxes(axisID).startPt.X
    dy = mAxes(axisID).endPt.Y - mAxes(axisID).startPt.Y

    Dim len2        As Double
    len2 = dx * dx + dy * dy

    If len2 < EPSILON Then
        result = mAxes(axisID).startPt
        Exit Sub
    End If

    Dim t           As Double
    t = ((pt.X - mAxes(axisID).startPt.X) * dx + (pt.Y - mAxes(axisID).startPt.Y) * dy) / len2

    result.X = mAxes(axisID).startPt.X + t * dx
    result.Y = mAxes(axisID).startPt.Y + t * dy
End Sub

Private Function SnapCenterLineEndpointsToAxis(lineID As Long, axisID As Long) As Boolean
    SnapCenterLineEndpointsToAxis = False

    Dim snappedStart As Boolean, snappedEnd As Boolean
    snappedStart = False
    snappedEnd = False

    Dim distStart   As Double
    distStart = PointToAxisDistance(mCenterLines(lineID).startPt, axisID)

    If distStart > 0 And distStart <= mAxisSnapDistance Then
        Dim projStart As Point2D
        ProjectPointToAxis mCenterLines(lineID).startPt, axisID, projStart
        mCenterLines(lineID).startPt = projStart
        snappedStart = True
    End If

    Dim distEnd     As Double
    distEnd = PointToAxisDistance(mCenterLines(lineID).endPt, axisID)

    If distEnd > 0 And distEnd <= mAxisSnapDistance Then
        Dim projEnd As Point2D
        ProjectPointToAxis mCenterLines(lineID).endPt, axisID, projEnd
        mCenterLines(lineID).endPt = projEnd
        snappedEnd = True
    End If

    If snappedStart Or snappedEnd Then
        Dim dx As Double, dy As Double
        dx = mCenterLines(lineID).endPt.X - mCenterLines(lineID).startPt.X
        dy = mCenterLines(lineID).endPt.Y - mCenterLines(lineID).startPt.Y
        mCenterLines(lineID).Length = Sqr(dx * dx + dy * dy)
        mCenterLines(lineID).angle = Atan2(dy, dx)
        SnapCenterLineEndpointsToAxis = True
    End If
End Function

' ==================== STEP 6A: MERGE PARALLEL CLOSE CENTERLINES (ENHANCED) ====================
Private Function MergeParallelCloseCenterLines() As Long
    MergeParallelCloseCenterLines = 0

    If mCenterCount <= 1 Then Exit Function

    ' Calculate max distance threshold: 2 * largest wall thickness
    Dim maxMergeDistance As Double
    maxMergeDistance = 200  ' Default minimum

    If mWallThicknessCount > 0 Then
        maxMergeDistance = mWallThicknesses(mWallThicknessCount - 1) * 2
    End If

    ' ENHANCED: Group centerlines into perpendicular distance bands
    Dim bands()     As ParallelBand
    Dim bandCount   As Long
    bandCount = 0
    ReDim bands(0 To mCenterCount - 1)

    Dim i As Long, j As Long

    ' STEP 1: Group centerlines into perpendicular bands
    For i = 0 To mCenterCount - 1
        If Not mCenterLines(i).IsActive Then GoTo NextBandGroup

        ' FIX: Validate centerline has valid geometry
        If mCenterLines(i).Length <= EPSILON Then GoTo NextBandGroup

        Dim foundBand As Boolean
        foundBand = False

        ' Try to add to existing band
        For j = 0 To bandCount - 1
            ' FIX: Validate band representative ID
            If bands(j).RepresentativeID < 0 Or bands(j).RepresentativeID >= mCenterCount Then GoTo NextBandCheck
            If Not mCenterLines(bands(j).RepresentativeID).IsActive Then GoTo NextBandCheck

            Dim perpDist As Double
            perpDist = CalculatePerpendicularDistanceCL(i, bands(j).RepresentativeID)

            ' Check angle similarity
            Dim angleDiff As Double
            angleDiff = Abs(mCenterLines(i).angle - mCenterLines(bands(j).RepresentativeID).angle)
            If angleDiff > PI Then angleDiff = 2 * PI - angleDiff

            If angleDiff <= (mAngleTolerance * DEG_TO_RAD) And perpDist <= (maxMergeDistance * 1.5) Then
                ' FIX: Check array bounds before adding
                If bands(j).count < mCenterCount Then
                    bands(j).CenterLineIDs(bands(j).count) = i
                    bands(j).count = bands(j).count + 1
                    foundBand = True
                    Exit For
                End If
            End If
NextBandCheck:
        Next j

        ' Create new band if not found
        If Not foundBand Then
            ' FIX: Check we haven't exceeded max bands
            If bandCount < mCenterCount Then
                bands(bandCount).RepresentativeID = i
                ReDim bands(bandCount).CenterLineIDs(0 To mCenterCount - 1)
                bands(bandCount).CenterLineIDs(0) = i
                bands(bandCount).count = 1
                bandCount = bandCount + 1
            End If
        End If

NextBandGroup:
    Next i

Debug.Print "  Grouped into " & bandCount & " perpendicular bands"

    ' STEP 2: Merge within each band using average axis projection
    Dim totalMerged As Long
    totalMerged = 0

    For i = 0 To bandCount - 1
        If bands(i).count <= 1 Then GoTo NextBandMerge

        Dim mergedInBand As Long
        mergedInBand = MergeCenterLinesInBand(bands(i), maxMergeDistance)
        totalMerged = totalMerged + mergedInBand

        If mergedInBand > 0 Then
Debug.Print "  Band " & i & ": merged " & mergedInBand & " centerlines"
        End If

NextBandMerge:
    Next i

    MergeParallelCloseCenterLines = totalMerged

    ' Compact array
    If totalMerged > 0 Then
        CompactCenterLineArray
    End If
End Function
' ==================== HELPER: Calculate Perpendicular Distance (SAFE VERSION) ====================
Private Function CalculatePerpendicularDistanceCL(line1ID As Long, line2ID As Long) As Double
    On Error Resume Next

    ' FIX: Validate IDs
    If line1ID < 0 Or line1ID >= mCenterCount Then
        CalculatePerpendicularDistanceCL = 999999
        Exit Function
    End If

    If line2ID < 0 Or line2ID >= mCenterCount Then
        CalculatePerpendicularDistanceCL = 999999
        Exit Function
    End If

    If Not mCenterLines(line1ID).IsActive Or Not mCenterLines(line2ID).IsActive Then
        CalculatePerpendicularDistanceCL = 999999
        Exit Function
    End If

    ' Calculate midpoints
    Dim mid1X As Double, mid1Y As Double
    Dim mid2X As Double, mid2Y As Double

    mid1X = (mCenterLines(line1ID).startPt.X + mCenterLines(line1ID).endPt.X) / 2
    mid1Y = (mCenterLines(line1ID).startPt.Y + mCenterLines(line1ID).endPt.Y) / 2

    mid2X = (mCenterLines(line2ID).startPt.X + mCenterLines(line2ID).endPt.X) / 2
    mid2Y = (mCenterLines(line2ID).startPt.Y + mCenterLines(line2ID).endPt.Y) / 2

    Dim dx As Double, dy As Double
    dx = mCenterLines(line2ID).endPt.X - mCenterLines(line2ID).startPt.X
    dy = mCenterLines(line2ID).endPt.Y - mCenterLines(line2ID).startPt.Y

    Dim dlen        As Double
    dlen = Sqr(dx * dx + dy * dy)

    ' FIX: Guard against zero-length line
    If dlen < EPSILON Then
        CalculatePerpendicularDistanceCL = Sqr((mid1X - mid2X) ^ 2 + (mid1Y - mid2Y) ^ 2)
    Else
        CalculatePerpendicularDistanceCL = Abs(dy * (mid1X - mCenterLines(line2ID).startPt.X) - _
                dx * (mid1Y - mCenterLines(line2ID).startPt.Y)) / dlen
    End If

    On Error GoTo 0
End Function

' ==================== HELPER: Merge Centerlines Within Band (FIXED) ====================
Private Function MergeCenterLinesInBand(band As ParallelBand, maxDist As Double) As Long
    On Error GoTo ErrorHandler
    MergeCenterLinesInBand = 0

    If band.count <= 1 Then Exit Function

    Dim activeCount As Long
    activeCount = 0
    Dim i As Long, lineID As Long

    ' Count active lines
    For i = 0 To band.count - 1
        lineID = band.CenterLineIDs(i)
        If lineID >= 0 And lineID < mCenterCount Then
            If mCenterLines(lineID).IsActive Then activeCount = activeCount + 1
        End If
    Next i

    If activeCount <= 1 Then Exit Function

    ' --- FIX: DO NOT AVERAGE ANGLES. PICK THE DOMINANT ANGLE AND SNAP IT. ---
    ' Averaging angles causes rotation around the center.
    ' Instead, we find the angle of the Longest Line in the band and snap it.

    Dim bestAngle   As Double
    Dim maxLength   As Double
    maxLength = -1

    For i = 0 To band.count - 1
        lineID = band.CenterLineIDs(i)
        If lineID >= 0 And lineID < mCenterCount Then
            If mCenterLines(lineID).IsActive Then
                If mCenterLines(lineID).Length > maxLength Then
                    maxLength = mCenterLines(lineID).Length
                    bestAngle = mCenterLines(lineID).angle
                End If
            End If
        End If
    Next i

    ' Snap the best angle to 0, 90, 180, 270
    Dim avgAngle    As Double
    avgAngle = SnapToCardinalAngle(bestAngle)

    ' Calculate reference point (average of all midpoints) -> This keeps the line Centered
    Dim refX As Double, refY As Double
    Dim validCount  As Long
    refX = 0: refY = 0
    validCount = 0

    For i = 0 To band.count - 1
        lineID = band.CenterLineIDs(i)
        If lineID >= 0 And lineID < mCenterCount Then
            If mCenterLines(lineID).IsActive Then
                refX = refX + (mCenterLines(lineID).startPt.X + mCenterLines(lineID).endPt.X) / 2
                refY = refY + (mCenterLines(lineID).startPt.Y + mCenterLines(lineID).endPt.Y) / 2
                validCount = validCount + 1
            End If
        End If
    Next i

    If validCount = 0 Then Exit Function
    refX = refX / validCount
    refY = refY / validCount

    ' Prepare projections
    Dim projections() As ProjectionData
    ReDim projections(0 To validCount - 1)

    Dim cosA As Double, sinA As Double
    cosA = Cos(avgAngle)
    sinA = Sin(avgAngle)

    Dim projCount   As Long
    projCount = 0

    For i = 0 To band.count - 1
        lineID = band.CenterLineIDs(i)
        If lineID >= 0 And lineID < mCenterCount Then
            If mCenterLines(lineID).IsActive Then
                Dim dx1 As Double, dy1 As Double, dx2 As Double, dy2 As Double
                dx1 = mCenterLines(lineID).startPt.X - refX
                dy1 = mCenterLines(lineID).startPt.Y - refY
                dx2 = mCenterLines(lineID).endPt.X - refX
                dy2 = mCenterLines(lineID).endPt.Y - refY

                projections(projCount).segmentID = lineID
                projections(projCount).StartProj = dx1 * cosA + dy1 * sinA
                projections(projCount).endProj = dx2 * cosA + dy2 * sinA

                If projections(projCount).StartProj > projections(projCount).endProj Then
                    Dim tmp As Double
                    tmp = projections(projCount).StartProj
                    projections(projCount).StartProj = projections(projCount).endProj
                    projections(projCount).endProj = tmp
                End If
                projCount = projCount + 1
            End If
        End If
    Next i

    If projCount <= 1 Then Exit Function

    ' Sort and Merge
    QuickSortProjections projections, 0, projCount - 1

    Dim j           As Long
    For i = 0 To projCount - 1
        lineID = projections(i).segmentID
        If mCenterLines(lineID).IsActive Then
            For j = i + 1 To projCount - 1
                Dim otherID As Long
                otherID = projections(j).segmentID

                If mCenterLines(otherID).IsActive Then
                    Dim overlapStart As Double, overlapEnd As Double
                    overlapStart = Max(projections(i).StartProj, projections(j).StartProj)
                    overlapEnd = Min(projections(i).endProj, projections(j).endProj)

                    Dim OverlapLength As Double
                    OverlapLength = overlapEnd - overlapStart
                    Dim minLength As Double
                    minLength = Min(projections(i).endProj - projections(i).StartProj, projections(j).endProj - projections(j).StartProj)

                    If minLength > 0.000001 Then
                        If (OverlapLength / minLength) >= 0.3 Then
                            Dim perpDistCheck As Double
                            perpDistCheck = CalculatePerpendicularDistanceCL(lineID, otherID)

                            If perpDistCheck <= maxDist Then
                                ' Merge using the SNAPPED axis
                                MergeTwoParallelCenterLinesOnAxis lineID, otherID, avgAngle, refX, refY
                                mCenterLines(otherID).IsActive = False

                                projections(i).StartProj = Min(projections(i).StartProj, projections(j).StartProj)
                                projections(i).endProj = Max(projections(i).endProj, projections(j).endProj)
                                MergeCenterLinesInBand = MergeCenterLinesInBand + 1
                            End If
                        End If
                    End If
                End If
            Next j
        End If
    Next i
    Exit Function
ErrorHandler:
    MergeCenterLinesInBand = 0
End Function
' ==================== HELPER: SNAP TO CARDINAL ANGLE ====================
Private Function SnapToCardinalAngle(angleRad As Double) As Double
    Dim normalized  As Double
    normalized = angleRad

    ' Normalize to 0-2PI
    Do While normalized < 0: normalized = normalized + 6.28318530717959: Loop
    Do While normalized >= 6.28318530717959: normalized = normalized - 6.28318530717959: Loop

    ' Snap to nearest 90 degrees (PI/2)
    Dim piHalf      As Double
    piHalf = 1.5707963267949

    Dim multiple    As Long
    multiple = Round(normalized / piHalf)

    SnapToCardinalAngle = CDbl(multiple) * piHalf
End Function

' ==================== HELPER: Merge Two Parallel Centerlines On Average Axis ====================
Private Sub MergeTwoParallelCenterLinesOnAxis(targetID As Long, sourceID As Long, _
        avgAngle As Double, refX As Double, refY As Double)
    ' Project all 4 endpoints onto average axis
    Dim points(0 To 3) As Point2D
    points(0) = mCenterLines(targetID).startPt
    points(1) = mCenterLines(targetID).endPt
    points(2) = mCenterLines(sourceID).startPt
    points(3) = mCenterLines(sourceID).endPt

    Dim projections(0 To 3) As Double
    Dim cosA As Double, sinA As Double
    cosA = Cos(avgAngle)
    sinA = Sin(avgAngle)

    Dim i           As Long
    For i = 0 To 3
        Dim dx As Double, dy As Double
        dx = points(i).X - refX
        dy = points(i).Y - refY
        projections(i) = dx * cosA + dy * sinA
    Next i

    ' Find min and max projections
    Dim minProj As Double, maxProj As Double
    minProj = projections(0)
    maxProj = projections(0)

    For i = 1 To 3
        If projections(i) < minProj Then minProj = projections(i)
        If projections(i) > maxProj Then maxProj = projections(i)
    Next i

    ' Reconstruct line on average axis
    mCenterLines(targetID).startPt.X = refX + minProj * cosA
    mCenterLines(targetID).startPt.Y = refY + minProj * sinA
    mCenterLines(targetID).endPt.X = refX + maxProj * cosA
    mCenterLines(targetID).endPt.Y = refY + maxProj * sinA

    ' Update geometry
    Dim finalDx As Double, finalDy As Double
    finalDx = mCenterLines(targetID).endPt.X - mCenterLines(targetID).startPt.X
    finalDy = mCenterLines(targetID).endPt.Y - mCenterLines(targetID).startPt.Y
    mCenterLines(targetID).Length = Sqr(finalDx * finalDx + finalDy * finalDy)
    mCenterLines(targetID).angle = Atan2(finalDy, finalDx)

    ' Merge properties (thickness, wallType)
    On Error Resume Next
    Dim tTh As Double, sTh As Double
    tTh = mCenterLines(targetID).Thickness
    sTh = mCenterLines(sourceID).Thickness

    ' Choose larger thickness
    If sTh > tTh Then
        mCenterLines(targetID).Thickness = sTh
        If Len(Trim$(mCenterLines(sourceID).WallType)) > 0 Then
            mCenterLines(targetID).WallType = mCenterLines(sourceID).WallType
        End If
    End If

    ' Update unique ID
    mCenterLines(targetID).uniqueID = GenerateUniqueID(mCenterLines(targetID))
    On Error GoTo 0
End Sub

Private Function CalculateDistanceBetweenCenterLines(id1 As Long, id2 As Long) As Double
    ' Calculate perpendicular distance between two parallel centerlines
    ' Use average of distances from endpoints of line2 to line1

    Dim dist1 As Double, dist2 As Double
    dist1 = PointToLineDistanceCL(mCenterLines(id2).startPt, id1)
    dist2 = PointToLineDistanceCL(mCenterLines(id2).endPt, id1)

    CalculateDistanceBetweenCenterLines = (dist1 + dist2) / 2
End Function

Private Function CalculateParallelOverlapPercentage(id1 As Long, id2 As Long) As Double
    ' Calculate how much two parallel lines overlap when projected onto same axis
    ' Returns value 0.0 to 1.0 (0% to 100% of shorter line)

    CalculateParallelOverlapPercentage = 0

    ' Use line1's direction as reference
    Dim angle       As Double
    angle = mCenterLines(id1).angle

    Dim cosA As Double, sinA As Double
    cosA = Cos(angle)
    sinA = Sin(angle)

    ' Project all 4 endpoints onto line1's direction vector
    Dim baseX As Double, baseY As Double
    baseX = mCenterLines(id1).startPt.X
    baseY = mCenterLines(id1).startPt.Y

    Dim proj1Start As Double, proj1End As Double
    Dim proj2Start As Double, proj2End As Double

    proj1Start = 0  ' Reference point
    proj1End = (mCenterLines(id1).endPt.X - baseX) * cosA + (mCenterLines(id1).endPt.Y - baseY) * sinA

    proj2Start = (mCenterLines(id2).startPt.X - baseX) * cosA + (mCenterLines(id2).startPt.Y - baseY) * sinA
    proj2End = (mCenterLines(id2).endPt.X - baseX) * cosA + (mCenterLines(id2).endPt.Y - baseY) * sinA

    ' Normalize so start < end
    If proj1Start > proj1End Then
        Dim tmp     As Double
        tmp = proj1Start
        proj1Start = proj1End
        proj1End = tmp
    End If

    If proj2Start > proj2End Then
        tmp = proj2Start
        proj2Start = proj2End
        proj2End = tmp
    End If

    ' Calculate overlap
    Dim overlapStart As Double, overlapEnd As Double
    overlapStart = Max(proj1Start, proj2Start)
    overlapEnd = Min(proj1End, proj2End)

    Dim OverlapLength As Double
    OverlapLength = overlapEnd - overlapStart

    If OverlapLength <= 0 Then
        CalculateParallelOverlapPercentage = 0
        Exit Function
    End If

    ' Calculate percentage relative to shorter line
    Dim len1 As Double, len2 As Double
    len1 = Abs(proj1End - proj1Start)
    len2 = Abs(proj2End - proj2Start)

    Dim shorterLength As Double
    shorterLength = Min(len1, len2)

    If shorterLength > EPSILON Then
        CalculateParallelOverlapPercentage = OverlapLength / shorterLength
    Else
        CalculateParallelOverlapPercentage = 0
    End If
End Function

Private Function AreParallelLinesAtEquivalentPosition(id1 As Long, id2 As Long, maxDist As Double) As Boolean
    ' Verify that parallel lines are at equivalent position by checking if
    ' endpoints are reasonably close to the opposite line (perpendicular projection)

    AreParallelLinesAtEquivalentPosition = False

    ' Check if start/end of line2 projects within line1's bounds (with tolerance)
    Dim proj2StartOnLine1 As Point2D, proj2EndOnLine1 As Point2D
    Dim validStart As Boolean, validEnd As Boolean

    validStart = ProjectPointToLineCLWithBounds(mCenterLines(id2).startPt, id1, proj2StartOnLine1)
    validEnd = ProjectPointToLineCLWithBounds(mCenterLines(id2).endPt, id1, proj2EndOnLine1)

    ' At least one endpoint of line2 should project onto line1
    If Not validStart And Not validEnd Then
        ' Check reverse: does line1 project onto line2?
        Dim proj1StartOnLine2 As Point2D, proj1EndOnLine2 As Point2D
        Dim validStart2 As Boolean, validEnd2 As Boolean

        validStart2 = ProjectPointToLineCLWithBounds(mCenterLines(id1).startPt, id2, proj1StartOnLine2)
        validEnd2 = ProjectPointToLineCLWithBounds(mCenterLines(id1).endPt, id2, proj1EndOnLine2)

        If Not validStart2 And Not validEnd2 Then
            Exit Function  ' No overlap in position
        End If
    End If

    ' Additional check: perpendicular distances should be consistent
    Dim distStart1 As Double, distEnd1 As Double
    Dim distStart2 As Double, distEnd2 As Double

    distStart1 = PointToLineDistanceCL(mCenterLines(id1).startPt, id2)
    distEnd1 = PointToLineDistanceCL(mCenterLines(id1).endPt, id2)
    distStart2 = PointToLineDistanceCL(mCenterLines(id2).startPt, id1)
    distEnd2 = PointToLineDistanceCL(mCenterLines(id2).endPt, id1)

    ' All distances should be similar (within tolerance)
    Dim avgDist     As Double
    avgDist = (distStart1 + distEnd1 + distStart2 + distEnd2) / 4

    Dim maxDeviation As Double
    maxDeviation = maxDist * 0.5  ' Allow 50% deviation from average

    If Abs(distStart1 - avgDist) > maxDeviation Then Exit Function
    If Abs(distEnd1 - avgDist) > maxDeviation Then Exit Function
    If Abs(distStart2 - avgDist) > maxDeviation Then Exit Function
    If Abs(distEnd2 - avgDist) > maxDeviation Then Exit Function

    AreParallelLinesAtEquivalentPosition = True
End Function

Private Function ProjectPointToLineCLWithBounds(pt As Point2D, lineID As Long, ByRef result As Point2D) As Boolean
    ' Project point onto line and check if projection is within line bounds (with small tolerance)
    ProjectPointToLineCLWithBounds = False

    Dim dx As Double, dy As Double
    dx = mCenterLines(lineID).endPt.X - mCenterLines(lineID).startPt.X
    dy = mCenterLines(lineID).endPt.Y - mCenterLines(lineID).startPt.Y

    Dim len2        As Double
    len2 = dx * dx + dy * dy

    If len2 < EPSILON Then Exit Function

    Dim t           As Double
    t = ((pt.X - mCenterLines(lineID).startPt.X) * dx + (pt.Y - mCenterLines(lineID).startPt.Y) * dy) / len2

    ' Check if t is within bounds (with 20% tolerance on each side)
    If t < -0.2 Or t > 1.2 Then Exit Function

    result.X = mCenterLines(lineID).startPt.X + t * dx
    result.Y = mCenterLines(lineID).startPt.Y + t * dy
    ProjectPointToLineCLWithBounds = True
End Function
' ==================== GEOMETRY-BASED COLLINEARITY CHECK (STRICT VERSION) ====================
Private Function AreGeometricallyCollinear(id1 As Long, id2 As Long, _
        Optional maxAngleDeg As Double = 1, _
        Optional maxPerpDistMM As Double = 30) As Boolean

    AreGeometricallyCollinear = False

    ' Validate IDs
    If id1 < 0 Or id1 >= mCenterCount Then Exit Function
    If id2 < 0 Or id2 >= mCenterCount Then Exit Function
    If Not mCenterLines(id1).IsActive Then Exit Function
    If Not mCenterLines(id2).IsActive Then Exit Function

    ' CRITICAL CHECK 1: Angle difference (VERY STRICT)
    Dim angleDiff   As Double
    angleDiff = Abs(mCenterLines(id1).angle - mCenterLines(id2).angle)
    If angleDiff > PI Then angleDiff = 2 * PI - angleDiff

    ' Reject if angle difference > 1 degree (prevent diagonal connections)
    If angleDiff > (maxAngleDeg * DEG_TO_RAD) Then Exit Function

    ' CRITICAL CHECK 2: Perpendicular distance (STRICT)
    Dim dist1 As Double, dist2 As Double, dist3 As Double, dist4 As Double
    dist1 = PointToLineDistanceInfiniteCL(mCenterLines(id2).startPt, id1)
    dist2 = PointToLineDistanceInfiniteCL(mCenterLines(id2).endPt, id1)
    dist3 = PointToLineDistanceInfiniteCL(mCenterLines(id1).startPt, id2)
    dist4 = PointToLineDistanceInfiniteCL(mCenterLines(id1).endPt, id2)

    ' All points must be VERY close (< 30mm by default)
    If dist1 > maxPerpDistMM Then Exit Function
    If dist2 > maxPerpDistMM Then Exit Function
    If dist3 > maxPerpDistMM Then Exit Function
    If dist4 > maxPerpDistMM Then Exit Function

    ' CRITICAL CHECK 3: Consistency check
    ' All 4 distances should be similar (not a diagonal case)
    Dim avgDist     As Double
    avgDist = (dist1 + dist2 + dist3 + dist4) / 4

    Dim maxDeviation As Double
    maxDeviation = maxPerpDistMM * 0.5

    If Abs(dist1 - avgDist) > maxDeviation Then Exit Function
    If Abs(dist2 - avgDist) > maxDeviation Then Exit Function
    If Abs(dist3 - avgDist) > maxDeviation Then Exit Function
    If Abs(dist4 - avgDist) > maxDeviation Then Exit Function

    ' CRITICAL CHECK 4: Gap/Overlap validation
    ' Lines should either have gap or overlap, not be far apart
    Dim gapDist     As Double
    gapDist = CalculateGapDistance(id1, id2)

    ' If gap > 5000mm (5 meters), likely not same wall segment
    If gapDist > 5000 Then Exit Function

    AreGeometricallyCollinear = True
End Function
' ==================== CALCULATE GAP DISTANCE ====================
Private Function CalculateGapDistance(id1 As Long, id2 As Long) As Double
    ' Project all 4 endpoints onto line1's direction
    Dim angle       As Double
    angle = mCenterLines(id1).angle

    Dim cosA As Double, sinA As Double
    cosA = Cos(angle)
    sinA = Sin(angle)

    Dim baseX As Double, baseY As Double
    baseX = mCenterLines(id1).startPt.X
    baseY = mCenterLines(id1).startPt.Y

    Dim proj1Start As Double, proj1End As Double
    Dim proj2Start As Double, proj2End As Double

    proj1Start = 0
    proj1End = (mCenterLines(id1).endPt.X - baseX) * cosA + _
            (mCenterLines(id1).endPt.Y - baseY) * sinA

    proj2Start = (mCenterLines(id2).startPt.X - baseX) * cosA + _
            (mCenterLines(id2).startPt.Y - baseY) * sinA
    proj2End = (mCenterLines(id2).endPt.X - baseX) * cosA + _
            (mCenterLines(id2).endPt.Y - baseY) * sinA

    ' Normalize
    If proj1Start > proj1End Then
        Dim tmp     As Double
        tmp = proj1Start
        proj1Start = proj1End
        proj1End = tmp
    End If

    If proj2Start > proj2End Then
        tmp = proj2Start
        proj2Start = proj2End
        proj2End = tmp
    End If

    ' Calculate gap
    If proj2Start >= proj1End Then
        CalculateGapDistance = proj2Start - proj1End
    ElseIf proj1Start >= proj2End Then
        CalculateGapDistance = proj1Start - proj2End
    Else
        CalculateGapDistance = -1  ' Overlapping
    End If
End Function
' ==================== CALCULATE OVERLAP SCALAR ====================
Private Function CalculateOverlapScalar(id1 As Long, id2 As Long) As overlapResult
    Dim result      As overlapResult
    result.HasOverlap = False
    result.OverlapPercent = 0

    ' Project onto line1's direction
    Dim angle       As Double
    angle = mCenterLines(id1).angle

    Dim cosA As Double, sinA As Double
    cosA = Cos(angle)
    sinA = Sin(angle)

    Dim baseX As Double, baseY As Double
    baseX = mCenterLines(id1).startPt.X
    baseY = mCenterLines(id1).startPt.Y

    Dim proj1Start As Double, proj1End As Double
    Dim proj2Start As Double, proj2End As Double

    proj1Start = 0
    proj1End = (mCenterLines(id1).endPt.X - baseX) * cosA + _
            (mCenterLines(id1).endPt.Y - baseY) * sinA

    proj2Start = (mCenterLines(id2).startPt.X - baseX) * cosA + _
            (mCenterLines(id2).startPt.Y - baseY) * sinA
    proj2End = (mCenterLines(id2).endPt.X - baseX) * cosA + _
            (mCenterLines(id2).endPt.Y - baseY) * sinA

    ' Normalize
    If proj1Start > proj1End Then
        Dim tmp     As Double
        tmp = proj1Start
        proj1Start = proj1End
        proj1End = tmp
    End If

    If proj2Start > proj2End Then
        tmp = proj2Start
        proj2Start = proj2End
        proj2End = tmp
    End If

    ' Calculate overlap
    Dim overlapStart As Double, overlapEnd As Double
    overlapStart = Max(proj1Start, proj2Start)
    overlapEnd = Min(proj1End, proj2End)

    Dim OverlapLength As Double
    OverlapLength = overlapEnd - overlapStart

    If OverlapLength > EPSILON Then
        result.HasOverlap = True

        ' Calculate percentage relative to SHORTER line
        Dim len1 As Double, len2 As Double
        len1 = Abs(proj1End - proj1Start)
        len2 = Abs(proj2End - proj2Start)

        Dim shorterLength As Double
        shorterLength = Min(len1, len2)

        If shorterLength > EPSILON Then
            result.OverlapPercent = OverlapLength / shorterLength
        End If
    End If

    CalculateOverlapScalar = result
End Function

Private Sub CompactCenterLineArray()
    Dim writeIdx    As Long
    writeIdx = 0

    Dim i           As Long
    For i = 0 To mCenterCount - 1
        If mCenterLines(i).IsActive Then
            If writeIdx <> i Then
                mCenterLines(writeIdx) = mCenterLines(i)
            End If
            writeIdx = writeIdx + 1
        End If
    Next i

    mCenterCount = writeIdx
    If mCenterCount > 0 Then
        ReDim Preserve mCenterLines(0 To mCenterCount - 1)
    End If
End Sub


' ==================== FINAL: REMOVE DUPLICATE CENTERLINES ====================
Private Function RemoveDuplicateCenterLines() As Long
    RemoveDuplicateCenterLines = 0

    If mCenterCount <= 1 Then Exit Function

    ' Use dictionary to track unique lines
    Dim uniqueDict  As Object
    Set uniqueDict = CreateObject("Scripting.Dictionary")

    Dim i           As Long
    For i = 0 To mCenterCount - 1
        If Not mCenterLines(i).IsActive Then GoTo NextDupCheck

        Dim uniqueID As String
        uniqueID = mCenterLines(i).uniqueID

        If uniqueDict.exists(uniqueID) Then
            ' Duplicate found - keep the longer one
            Dim existingIdx As Long
            existingIdx = uniqueDict(uniqueID)

            If mCenterLines(i).Length > mCenterLines(existingIdx).Length Then
                ' Current is longer, deactivate existing
                mCenterLines(existingIdx).IsActive = False
                uniqueDict(uniqueID) = i
            Else
                ' Existing is longer, deactivate current
                mCenterLines(i).IsActive = False
            End If

            RemoveDuplicateCenterLines = RemoveDuplicateCenterLines + 1
        Else
            uniqueDict.Add uniqueID, i
        End If

NextDupCheck:
    Next i

    ' Compact array
    CompactCenterLineArray

    Set uniqueDict = Nothing
End Function

' ==================== BREAK AT GRID INTERSECTIONS (UPDATED) ====================
Private Function BreakCenterLinesAtGridIntersections() As Long
    BreakCenterLinesAtGridIntersections = 0

    If mAxesCount = 0 Or mCenterCount = 0 Then Exit Function

    Dim newCenterLines() As CenterLine
    Dim newCount    As Long
    newCount = 0
    ReDim newCenterLines(0 To mCenterCount * mAxesCount)

    Dim i As Long, j As Long
    For i = 0 To mCenterCount - 1
        If Not mCenterLines(i).IsActive Then GoTo NextBreakLine

        Dim intersections() As Double
        Dim intCount As Long
        intCount = 0
        ReDim intersections(0 To mAxesCount - 1)

        For j = 0 To mAxesCount - 1
            Dim intPt As Point2D
            If GetLineIntersection2D(mCenterLines(i), mAxes(j), intPt) Then
                intersections(intCount) = ProjectPointOnCenterLine(intPt, i)
                intCount = intCount + 1
            End If
        Next j

        If intCount = 0 Then GoTo NextBreakLine

        If intCount > 1 Then
            ReDim Preserve intersections(0 To intCount - 1)
            QuickSortDoubles intersections, 0, intCount - 1
        End If

        Dim segments() As Point2D
        ReDim segments(0 To intCount + 1)

        segments(0) = mCenterLines(i).startPt

        Dim k       As Long
        For k = 0 To intCount - 1
            segments(k + 1) = UnprojectPointOnCenterLine(intersections(k), i)
        Next k

        segments(intCount + 1) = mCenterLines(i).endPt

        ' CRITICAL: Preserve original properties before breaking
        Dim originalThickness As Double
        Dim originalWallType As String
        Dim originalSourcePairID As Long

        originalThickness = mCenterLines(i).Thickness
        originalWallType = Trim$(mCenterLines(i).WallType)
        originalSourcePairID = mCenterLines(i).SourcePairID

        ' Ensure WallType exists
        If Len(originalWallType) = 0 And originalThickness > 0 Then
            originalWallType = "W" & CStr(CInt(originalThickness))
        End If

        For k = 0 To intCount
            Dim seg As CenterLine
            seg.startPt = segments(k)
            seg.endPt = segments(k + 1)

            Dim dx As Double, dy As Double
            dx = seg.endPt.X - seg.startPt.X
            dy = seg.endPt.Y - seg.startPt.Y
            seg.Length = Sqr(dx * dx + dy * dy)

            If seg.Length >= MIN_SEGMENT_LENGTH Then
                seg.angle = Atan2(dy, dx)
                seg.VectorID = -1
                seg.IsActive = True

                ' INHERIT ALL PROPERTIES from parent
                seg.SourcePairID = originalSourcePairID
                seg.Thickness = originalThickness
                seg.WallType = originalWallType

                ' Generate unique ID AFTER setting all properties
                seg.uniqueID = GenerateUniqueID(seg)

                newCenterLines(newCount) = seg
                newCount = newCount + 1
            End If
        Next k

        mCenterLines(i).IsActive = False
        BreakCenterLinesAtGridIntersections = BreakCenterLinesAtGridIntersections + 1

NextBreakLine:
    Next i

    If newCount > 0 Then
        Dim oldCount As Long
        oldCount = mCenterCount

        ReDim Preserve mCenterLines(0 To mCenterCount + newCount - 1)

        For i = 0 To newCount - 1
            mCenterLines(oldCount + i) = newCenterLines(i)
        Next i

        mCenterCount = mCenterCount + newCount

        CompactCenterLineArray
    End If
End Function
Private Function GetLineIntersection2D(line1 As CenterLine, line2 As AxisLine, ByRef intPt As Point2D) As Boolean
    GetLineIntersection2D = False

    Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
    Dim x3 As Double, y3 As Double, x4 As Double, y4 As Double

    x1 = line1.startPt.X: y1 = line1.startPt.Y
    x2 = line1.endPt.X: y2 = line1.endPt.Y
    x3 = line2.startPt.X: y3 = line2.startPt.Y
    x4 = line2.endPt.X: y4 = line2.endPt.Y

    Dim denom       As Double
    denom = (x1 - x2) * (y3 - y4) - (y1 - y2) * (x3 - x4)

    If Abs(denom) < EPSILON Then Exit Function

    Dim t As Double, u As Double
    t = ((x1 - x3) * (y3 - y4) - (y1 - y3) * (x3 - x4)) / denom
    u = -((x1 - x2) * (y1 - y3) - (y1 - y2) * (x1 - x3)) / denom

    If t >= 0 And t <= 1 And u >= 0 And u <= 1 Then
        intPt.X = x1 + t * (x2 - x1)
        intPt.Y = y1 + t * (y2 - y1)
        GetLineIntersection2D = True
    End If
End Function

Private Function ProjectPointOnCenterLine(pt As Point2D, lineID As Long) As Double
    Dim dx As Double, dy As Double
    dx = mCenterLines(lineID).endPt.X - mCenterLines(lineID).startPt.X
    dy = mCenterLines(lineID).endPt.Y - mCenterLines(lineID).startPt.Y

    Dim len2        As Double
    len2 = dx * dx + dy * dy

    If len2 < EPSILON Then
        ProjectPointOnCenterLine = 0
        Exit Function
    End If

    ProjectPointOnCenterLine = ((pt.X - mCenterLines(lineID).startPt.X) * dx + (pt.Y - mCenterLines(lineID).startPt.Y) * dy) / len2
End Function

Private Function UnprojectPointOnCenterLine(t As Double, lineID As Long) As Point2D
    Dim result      As Point2D
    result.X = mCenterLines(lineID).startPt.X + t * (mCenterLines(lineID).endPt.X - mCenterLines(lineID).startPt.X)
    result.Y = mCenterLines(lineID).startPt.Y + t * (mCenterLines(lineID).endPt.Y - mCenterLines(lineID).startPt.Y)
    UnprojectPointOnCenterLine = result
End Function

' ==================== APPLY OFFSETS ====================
Private Sub ApplyOffsetToCenterLines(offsetX As Double, offsetY As Double)
    Dim i           As Long
    For i = 0 To mCenterCount - 1
        mCenterLines(i).startPt.X = mCenterLines(i).startPt.X + offsetX
        mCenterLines(i).startPt.Y = mCenterLines(i).startPt.Y + offsetY
        mCenterLines(i).endPt.X = mCenterLines(i).endPt.X + offsetX
        mCenterLines(i).endPt.Y = mCenterLines(i).endPt.Y + offsetY
    Next i
End Sub

Sub ApplyOffsetToAxes(offsetX As Double, offsetY As Double)
    Dim i           As Long
    For i = 0 To mAxesCount - 1
        mAxes(i).startPt.X = mAxes(i).startPt.X + offsetX
        mAxes(i).startPt.Y = mAxes(i).startPt.Y + offsetY
        mAxes(i).endPt.X = mAxes(i).endPt.X + offsetX
        mAxes(i).endPt.Y = mAxes(i).endPt.Y + offsetY
    Next i
End Sub

' ==================== DRAW RESULTS (UPDATED - GUARANTEED LABEL & XDATA) ====================
Private Function DrawCenterLines() As Long
    On Error GoTo ErrHandler
    DrawCenterLines = 0

    Dim ms          As Object

    ' Ensure AutoCAD connection
    On Error Resume Next
    If mAcadApp Is Nothing Then
        Set mAcadApp = GetObject(, "AutoCAD.Application")
    End If
    If mAcadDoc Is Nothing Then
        If Not mAcadApp Is Nothing Then Set mAcadDoc = mAcadApp.ActiveDocument
    End If
    On Error GoTo ErrHandler

    ' Set drawing units to millimeters
    On Error Resume Next
    If Not mAcadDoc Is Nothing Then
        mAcadDoc.SetVariable "INSUNITS", 4
    End If
    On Error GoTo ErrHandler

    Set ms = mAcadDoc.ModelSpace

    ' Calculate minimum length threshold
    Dim minWallThickness As Double
    minWallThickness = 100

    If mWallThicknessCount > 0 Then
        minWallThickness = mWallThicknesses(0)
        Dim k       As Long
        For k = 1 To mWallThicknessCount - 1
            If mWallThicknesses(k) < minWallThickness Then
                minWallThickness = mWallThicknesses(k)
            End If
        Next k
    End If

    Dim minLengthThreshold As Double
    minLengthThreshold = Max(minWallThickness * 1.5, 150)

Debug.Print "Drawing threshold: min wall thickness=" & Format(minWallThickness, "0.0") & _
        "mm, min centerline length=" & Format(minLengthThreshold, "0.0") & "mm"

    ' Ensure label layer exists
    EnsureLayerExists "dts_frame_label", 7

    Dim i           As Long
    For i = 0 To mCenterCount - 1
        If Not mCenterLines(i).IsActive Then GoTo NextDraw
        If mCenterLines(i).Length < MIN_SEGMENT_LENGTH Then GoTo NextDraw

        ' Determine required length
        Dim requiredLength As Double
        requiredLength = minLengthThreshold

        If mCenterLines(i).Thickness > 0 Then
            requiredLength = mCenterLines(i).Thickness * 2
        ElseIf mCenterLines(i).SourcePairID >= 0 And mCenterLines(i).SourcePairID < mPairCount Then
            requiredLength = mWallPairs(mCenterLines(i).SourcePairID).Thickness * 2
        End If

        ' Skip if too short
        Dim lengthTolerance As Double
        lengthTolerance = 0.1

        If mCenterLines(i).Length < (requiredLength - lengthTolerance) Then
Debug.Print "Skipping short centerline: Length=" & Format(mCenterLines(i).Length, "0.00") & _
        " < Required=" & Format(requiredLength, "0.00")
            GoTo NextDraw
        End If

        Dim sp(0 To 2) As Double, ep(0 To 2) As Double
        sp(0) = mCenterLines(i).startPt.X
        sp(1) = mCenterLines(i).startPt.Y
        sp(2) = 0
        ep(0) = mCenterLines(i).endPt.X
        ep(1) = mCenterLines(i).endPt.Y
        ep(2) = 0

        Dim lineObj As Object
        Set lineObj = ms.AddLine(sp, ep)

        If Not lineObj Is Nothing Then
            On Error Resume Next
            lineObj.layer = "DTS_WALL_DIAGRAM"
            lineObj.color = 2
            On Error GoTo ErrHandler
            DrawCenterLines = DrawCenterLines + 1

            ' DETERMINE WALL PROPERTIES (with fallback chain)
            Dim wallThickness As Double
            Dim wallTypeStr As String

            wallThickness = mCenterLines(i).Thickness
            wallTypeStr = Trim$(mCenterLines(i).WallType)

            ' Fallback 1: Try source pair
            If wallThickness <= 0 Then
                If mCenterLines(i).SourcePairID >= 0 And mCenterLines(i).SourcePairID < mPairCount Then
                    wallThickness = mWallPairs(mCenterLines(i).SourcePairID).Thickness
                End If
            End If

            ' Fallback 2: Use smallest known thickness
            If wallThickness <= 0 And mWallThicknessCount > 0 Then
                wallThickness = mWallThicknesses(0)
            End If

            ' Fallback 3: Default to 200mm
            If wallThickness <= 0 Then
                wallThickness = 200
            End If

            ' Build WallType from thickness if empty
            If Len(wallTypeStr) = 0 Then
                wallTypeStr = "W" & CStr(CInt(wallThickness))
            End If

            ' ALWAYS ATTACH XDATA (never skip)
            AttachWallThicknessXData lineObj, wallThickness, wallTypeStr

            ' ALWAYS DRAW LABEL (guaranteed fallback)
            Dim labelDrawn As Boolean
            labelDrawn = False

            ' Try Core_CAD_Plotter first
            On Error Resume Next
            Core_CAD_Plotter.PlotFrameLabel mAcadDoc, sp(0), sp(1), sp(2), ep(0), ep(1), ep(2), wallTypeStr, 80
            If err.number = 0 Then
                labelDrawn = True
            End If
            err.Clear
            On Error GoTo ErrHandler

            ' Fallback: Draw label locally
            If Not labelDrawn Then
                Dim midX As Double, midY As Double
                midX = (sp(0) + ep(0)) / 2
                midY = (sp(1) + ep(1)) / 2

                Dim txtPt(0 To 2) As Double
                txtPt(0) = midX
                txtPt(1) = midY
                txtPt(2) = 0

                On Error Resume Next
                Dim txtObj As Object
                Set txtObj = ms.AddText(CStr(wallTypeStr), txtPt, 80)
                If Not txtObj Is Nothing Then
                    txtObj.layer = "dts_frame_label"
                    txtObj.color = 7
                End If
                err.Clear
                On Error GoTo ErrHandler
            End If
        End If

NextDraw:
    Next i

    Exit Function
ErrHandler:
Debug.Print "Error drawing centerline: " & err.description
    Resume Next
End Function

Private Function DrawAxes() As Long
    On Error GoTo ErrHandler
    DrawAxes = 0

    If mAxesCount = 0 Then Exit Function

    ' Ensure AutoCAD connection object available
    On Error Resume Next
    If mAcadApp Is Nothing Then
        Set mAcadApp = GetObject(, "AutoCAD.Application")
    End If
    If mAcadDoc Is Nothing Then
        If Not mAcadApp Is Nothing Then Set mAcadDoc = mAcadApp.ActiveDocument
    End If
    On Error GoTo ErrHandler

    ' Attempt to set drawing units to millimeters (INSUNITS = 4)
    On Error Resume Next
    If Not mAcadDoc Is Nothing Then
        ' Set the drawing insertion units to Millimeters (4)
        ' Note: some AutoCAD versions may restrict setting INSUNITS, handle errors gracefully
        mAcadDoc.SetVariable "INSUNITS", 4
    End If
    On Error GoTo ErrHandler

    Dim ms          As Object
    Set ms = mAcadDoc.ModelSpace

    Dim i           As Long
    For i = 0 To mAxesCount - 1
        Dim sp(0 To 2) As Double, ep(0 To 2) As Double
        sp(0) = mAxes(i).startPt.X
        sp(1) = mAxes(i).startPt.Y
        sp(2) = 0
        ep(0) = mAxes(i).endPt.X
        ep(1) = mAxes(i).endPt.Y
        ep(2) = 0

        Dim lineObj As Object
        Set lineObj = ms.AddLine(sp, ep)

        If Not lineObj Is Nothing Then
            On Error Resume Next
            lineObj.layer = "DTS_AXIS_LINE"
            lineObj.color = 5
            On Error GoTo ErrHandler
            DrawAxes = DrawAxes + 1
        End If
    Next i

    Exit Function
ErrHandler:
Debug.Print "Error drawing axis: " & err.description
    Resume Next
End Function
Private Function GetInsertionPoint() As Variant
    On Error GoTo ErrHandler

    On Error Resume Next
    Set mAcadApp = GetObject(, "AutoCAD.Application")
    If Not mAcadApp Is Nothing Then
        Set mAcadDoc = mAcadApp.ActiveDocument
    End If
    On Error GoTo ErrHandler

    On Error Resume Next
    AppActivate mAcadApp.Caption
    DoEvents
    On Error GoTo ErrHandler

    Dim pt          As Variant
    pt = mAcadApp.ActiveDocument.Utility.GetPoint(, vbCrLf & "Pick insertion point for diagram: ")
    GetInsertionPoint = pt

    Exit Function
ErrHandler:
    GetInsertionPoint = Empty
End Function

' ==================== UTILITY FUNCTIONS ====================
Private Function PointDistance2D(p1 As Point2D, p2 As Point2D) As Double
    Dim dx As Double, dy As Double
    dx = p2.X - p1.X
    dy = p2.Y - p1.Y
    PointDistance2D = Sqr(dx * dx + dy * dy)
End Function

Private Function Atan2(Y As Double, X As Double) As Double
    If X > 0 Then
        Atan2 = Atn(Y / X)
    ElseIf X < 0 And Y >= 0 Then
        Atan2 = Atn(Y / X) + PI
    ElseIf X < 0 And Y < 0 Then
        Atan2 = Atn(Y / X) - PI
    ElseIf X = 0 And Y > 0 Then
        Atan2 = HALF_PI
    ElseIf X = 0 And Y < 0 Then
        Atan2 = -HALF_PI
    Else
        Atan2 = 0
    End If
End Function

Private Sub QuickSortDoubles(ByRef arr() As Double, Left As Long, Right As Long)
    If Left >= Right Then Exit Sub

    Dim pivot       As Double
    pivot = arr((Left + Right) \ 2)

    Dim i As Long, j As Long
    i = Left: j = Right

    Do While i <= j
        Do While arr(i) < pivot
            i = i + 1
        Loop
        Do While arr(j) > pivot
            j = j - 1
        Loop
        If i <= j Then
            Dim tmp As Double
            tmp = arr(i)
            arr(i) = arr(j)
            arr(j) = tmp
            i = i + 1
            j = j - 1
        End If
    Loop

    If Left < j Then QuickSortDoubles arr, Left, j
    If i < Right Then QuickSortDoubles arr, i, Right
End Sub

Private Function Min(a As Double, b As Double) As Double
    If a < b Then Min = a Else Min = b
End Function

Private Function Max(a As Double, b As Double) As Double
    If a > b Then Max = a Else Max = b
End Function

Private Sub EnsureLayerExists(layerName As String, colorIndex As Long)
    On Error Resume Next

    Dim lay         As Object
    Set lay = mAcadDoc.layers.item(layerName)

    If err.number <> 0 Then
        err.Clear
        Set lay = mAcadDoc.layers.Add(layerName)
        lay.color = colorIndex
    End If

    On Error GoTo 0
End Sub
' ==================== VALIDATE CONNECTION SAFETY ====================
Private Function IsConnectionSafe(id1 As Long, id2 As Long, gapDist As Double) As Boolean
    IsConnectionSafe = False

    ' Rule 1: Gap must be reasonable (< 3 meters for typical door/window)
    If gapDist < 0 Or gapDist > 3000 Then Exit Function

    ' Rule 2: Lines must have similar lengths (prevent connecting stub to main wall)
    Dim len1 As Double, len2 As Double
    len1 = mCenterLines(id1).Length
    len2 = mCenterLines(id2).Length

    Dim lengthRatio As Double
    If len1 > len2 Then
        lengthRatio = len2 / len1
    Else
        lengthRatio = len1 / len2
    End If

    ' If one line is > 10x longer than the other, likely different segments
    If lengthRatio < 0.1 Then Exit Function

    ' Rule 3: Check endpoint proximity
    ' At least one pair of endpoints should be close
    Dim d1 As Double, d2 As Double, d3 As Double, d4 As Double
    d1 = PointDistance2D(mCenterLines(id1).startPt, mCenterLines(id2).startPt)
    d2 = PointDistance2D(mCenterLines(id1).startPt, mCenterLines(id2).endPt)
    d3 = PointDistance2D(mCenterLines(id1).endPt, mCenterLines(id2).startPt)
    d4 = PointDistance2D(mCenterLines(id1).endPt, mCenterLines(id2).endPt)

    Dim minEndpointDist As Double
    minEndpointDist = d1
    If d2 < minEndpointDist Then minEndpointDist = d2
    If d3 < minEndpointDist Then minEndpointDist = d3
    If d4 < minEndpointDist Then minEndpointDist = d4

    ' Closest endpoints should be within gap distance + tolerance
    If minEndpointDist > (gapDist + 200) Then Exit Function

    IsConnectionSafe = True
End Function
' ==================== ATTACH WALL XDATA (UPDATED STANDARDIZED) ====================
Private Sub AttachWallThicknessXData(lineObj As Object, Thickness As Double, Optional WallType As String = "")
    On Error GoTo ErrHandler

    ' Register application
    On Error Resume Next
    lineObj.Application.ActiveDocument.RegisteredApplications.Add "DTS_APP"
    On Error GoTo ErrHandler

    ' Standard Schema:
    ' 0: AppName (1001)
    ' 1: Thickness (1040)
    ' 2: WallType (1000)
    ' 3: LoadPattern (1000) -> Default "DL"
    ' 4: LoadValue (1040)   -> Default 0.0

    Dim xdType(0 To 4) As Integer
    Dim xdVal(0 To 4) As Variant

    ' Header
    xdType(0) = 1001: xdVal(0) = "DTS_APP"

    ' Thickness
    xdType(1) = 1040: xdVal(1) = Thickness

    ' WallType
    xdType(2) = 1000
    If Len(Trim$(WallType)) > 0 Then
        xdVal(2) = CStr(WallType)
    Else
        xdVal(2) = ""
    End If

    ' LoadPattern (Default)
    xdType(3) = 1000: xdVal(3) = "DL"

    ' LoadValue (Default)
    xdType(4) = 1040: xdVal(4) = 0#

    lineObj.SetXData xdType, xdVal

    Exit Sub
ErrHandler:
Debug.Print "ERROR in AttachWallThicknessXData: " & err.description
End Sub


Private Function AlignCenterLinesToSmartGrid(tolerance As Double) As Long
    AlignCenterLinesToSmartGrid = 0
    If mCenterCount <= 1 Then Exit Function

    Dim i           As Long
    Dim horizontalCount As Long, verticalCount As Long

    ' Arrays to hold indices
    Dim hIndices() As Long, vIndices() As Long
    ReDim hIndices(0 To mCenterCount - 1)
    ReDim vIndices(0 To mCenterCount - 1)

    ' 1. Separate Horizontal and Vertical lines
    For i = 0 To mCenterCount - 1
        If Not mCenterLines(i).IsActive Then GoTo NextSep

        ' Check strict orthogonality (already normalized in Step 0A)
        Dim isHorz As Boolean, isVert As Boolean
        isHorz = (Abs(mCenterLines(i).startPt.Y - mCenterLines(i).endPt.Y) < EPSILON)
        isVert = (Abs(mCenterLines(i).startPt.X - mCenterLines(i).endPt.X) < EPSILON)

        If isHorz Then
            hIndices(horizontalCount) = i
            horizontalCount = horizontalCount + 1
        ElseIf isVert Then
            vIndices(verticalCount) = i
            verticalCount = verticalCount + 1
        End If
NextSep:
    Next i

    ' 2. Process Horizontal Lines (Align Y coordinates)
    If horizontalCount > 1 Then
        AlignGroup hIndices, horizontalCount, tolerance, True
    End If

    ' 3. Process Vertical Lines (Align X coordinates)
    If verticalCount > 1 Then
        AlignGroup vIndices, verticalCount, tolerance, False
    End If

    ' Return total processed (approximation)
    AlignCenterLinesToSmartGrid = horizontalCount + verticalCount

    ' Important: Re-generate unique IDs after modifying geometry
    For i = 0 To mCenterCount - 1
        If mCenterLines(i).IsActive Then
            mCenterLines(i).uniqueID = GenerateUniqueID(mCenterLines(i))
        End If
    Next i
End Function

Private Sub AlignGroup(indices() As Long, count As Long, tolerance As Double, isHorizontal As Boolean)
    ' Sort indices based on coordinate (Y for horizontal, X for vertical)
    QuickSortIndicesByCoord indices, 0, count - 1, isHorizontal

    Dim clusters()  As AxisCluster
    Dim clusterCount As Long
    clusterCount = 0
    ReDim clusters(0 To count - 1)

    Dim i           As Long
    Dim currentCoord As Double
    Dim lineID      As Long

    ' Initialize first cluster
    lineID = indices(0)
    clusters(0).Coordinate = IIf(isHorizontal, mCenterLines(lineID).startPt.Y, mCenterLines(lineID).startPt.X)
    clusters(0).TotalLength = mCenterLines(lineID).Length
    ReDim clusters(0).lineIndices(0 To count - 1)
    clusters(0).lineIndices(0) = lineID
    clusters(0).count = 1
    clusterCount = 1

    ' Cluster remaining lines
    For i = 1 To count - 1
        lineID = indices(i)
        currentCoord = IIf(isHorizontal, mCenterLines(lineID).startPt.Y, mCenterLines(lineID).startPt.X)

        ' Check distance to the Weighted Average of the current cluster
        Dim prevClusterIdx As Long
        prevClusterIdx = clusterCount - 1

        ' Calculate current weighted average of the cluster
        Dim clusterAvg As Double
        clusterAvg = clusters(prevClusterIdx).Coordinate

        Dim prevID  As Long
        prevID = indices(i - 1)
        Dim prevCoord As Double
        prevCoord = IIf(isHorizontal, mCenterLines(prevID).startPt.Y, mCenterLines(prevID).startPt.X)

        If Abs(currentCoord - prevCoord) <= tolerance Then
            ' Add to current cluster
            Dim idx As Long
            idx = clusters(prevClusterIdx).count
            clusters(prevClusterIdx).lineIndices(idx) = lineID
            clusters(prevClusterIdx).count = idx + 1
        Else
            ' Start new cluster
            clusters(clusterCount).Coordinate = currentCoord
            clusters(clusterCount).TotalLength = mCenterLines(lineID).Length
            ReDim clusters(clusterCount).lineIndices(0 To count - 1)
            clusters(clusterCount).lineIndices(0) = lineID
            clusters(clusterCount).count = 1
            clusterCount = clusterCount + 1
        End If
    Next i

    ' Apply Weighted Average Alignment
    Dim c As Long, k As Long
    For c = 0 To clusterCount - 1
        If clusters(c).count > 0 Then
            Dim sumCoordLen As Double
            Dim sumLen As Double
            sumCoordLen = 0
            sumLen = 0

            ' Calculate Weighted Average Coordinate
            For k = 0 To clusters(c).count - 1
                Dim lid As Long
                lid = clusters(c).lineIndices(k)
                Dim lLen As Double
                lLen = mCenterLines(lid).Length
                Dim lCoord As Double
                lCoord = IIf(isHorizontal, mCenterLines(lid).startPt.Y, mCenterLines(lid).startPt.X)

                sumCoordLen = sumCoordLen + (lCoord * lLen)
                sumLen = sumLen + lLen
            Next k

            Dim targetCoord As Double
            If sumLen > EPSILON Then
                targetCoord = sumCoordLen / sumLen
            Else
                ' Fallback for zero length lines
                targetCoord = IIf(isHorizontal, mCenterLines(clusters(c).lineIndices(0)).startPt.Y, mCenterLines(clusters(c).lineIndices(0)).startPt.X)
            End If

            ' Snap all lines in cluster to targetCoord
            For k = 0 To clusters(c).count - 1
                lid = clusters(c).lineIndices(k)
                If isHorizontal Then
                    mCenterLines(lid).startPt.Y = targetCoord
                    mCenterLines(lid).endPt.Y = targetCoord
                Else
                    mCenterLines(lid).startPt.X = targetCoord
                    mCenterLines(lid).endPt.X = targetCoord
                End If
            Next k
        End If
    Next c
End Sub

Private Sub QuickSortIndicesByCoord(indices() As Long, Left As Long, Right As Long, isHorizontal As Boolean)
    If Left >= Right Then Exit Sub

    Dim pivot       As Double
    Dim midID       As Long
    midID = indices((Left + Right) \ 2)
    pivot = IIf(isHorizontal, mCenterLines(midID).startPt.Y, mCenterLines(midID).startPt.X)

    Dim i As Long, j As Long
    i = Left: j = Right

    Do While i <= j
        Dim valI As Double, valJ As Double
        valI = IIf(isHorizontal, mCenterLines(indices(i)).startPt.Y, mCenterLines(indices(i)).startPt.X)
        valJ = IIf(isHorizontal, mCenterLines(indices(j)).startPt.Y, mCenterLines(indices(j)).startPt.X)

        Do While valI < pivot
            i = i + 1
            valI = IIf(isHorizontal, mCenterLines(indices(i)).startPt.Y, mCenterLines(indices(i)).startPt.X)
        Loop
        Do While valJ > pivot
            j = j - 1
            valJ = IIf(isHorizontal, mCenterLines(indices(j)).startPt.Y, mCenterLines(indices(j)).startPt.X)
        Loop

        If i <= j Then
            Dim tmp As Long
            tmp = indices(i): indices(i) = indices(j): indices(j) = tmp
            i = i + 1: j = j - 1
        End If
    Loop

    If Left < j Then QuickSortIndicesByCoord indices, Left, j, isHorizontal
    If i < Right Then QuickSortIndicesByCoord indices, i, Right, isHorizontal
End Sub
