Attribute VB_Name = "M12A_PunchingShea"
Option Explicit
'===============================================================
' Module: M12A_PunchingShear (IMPROVED - SEGMENT-BASED CLIPPING)
' Purpose: Draw punching shear perimeter with proper corner/edge detection
'
' Main changes vs previous version:
' - Keep polygon vertex order (no global angle sort for final perimeter)
' - Perform segment-based clipping of offset polygon against attached slabs
'   using intersection computation instead of only "nearest projection"
' - Avoid artificially "closing" perimeter through regions without slab
' - Same API and module name, works for both AREA and LINE output
'===============================================================

Private Const COORD_TOLERANCE As Double = 0.001
Private Const PI As Double = 3.14159265358979
Private Const MIN_EFFECTIVE_DEPTH As Double = 20#
Private Const EPSILON As Double = 0.000000001
Private Const VERTEX_MATCH_TOL As Double = 1#  ' mm - tolerance for vertex matching
Private Const EDGE_MATCH_TOL As Double = 2#    ' mm - tolerance for edge matching
Private Const SLAB_ELEV_TOL As Double = 50#    ' mm - elevation tolerance

' Slab cache
Private Type SlabAreaInfo
    Name As String
    elevation As Double
    Thickness As Double
    numPoints As Long
    PointsX() As Double
    PointsY() As Double
End Type

Private g_SlabAreas() As SlabAreaInfo
Private g_SlabAreasCount As Long
Private g_SlabAreasCached As Boolean
Private g_UniqueCounter As Long

'===============================================================
' SEGMENT-BASED CLIPPING WITH ATTACHED SLABS
'===============================================================
' Internal storage for last clipped polygon to reuse after function call
Private g_LastClipPoly() As Double
Private g_LastClipCount As Long

'===============================================================
' INITIALIZATION
'===============================================================
Public Sub M12A_InitializeSlabCache()
    On Error GoTo ErrorHandler
    g_SlabAreasCached = False
    g_SlabAreasCount = 0
    g_UniqueCounter = 0

    Dim numAreas As Long
    Dim areaNames() As String
    If SapModel.AreaObj.GetNameList(numAreas, areaNames) <> 0 Then Exit Sub
    If numAreas = 0 Then Exit Sub

    Dim columnPrefixes As Object
    Set columnPrefixes = GetColumnPrefixesFromSettings()

    ReDim g_SlabAreas(0 To numAreas - 1)
    Dim idx As Long: idx = 0
    Dim skippedCount As Long: skippedCount = 0

    Dim i As Long
    For i = 0 To numAreas - 1
        Dim areaName As String: areaName = areaNames(i)

        If IsColumnCrossSectionArea(areaName, columnPrefixes) Then
            skippedCount = skippedCount + 1
            GoTo NextArea
        End If

        Dim propName As String
        Dim Thickness As Double: Thickness = 0
        Dim acceptArea As Boolean: acceptArea = False
        
        If SapModel.AreaObj.GetProperty(areaName, propName) = 0 Then
            If Trim$(propName) = "" Or UCase$(Trim$(propName)) = "NONE" Then
                ' Heuristic check for "None" properties
                Dim numPts2 As Long
                Dim ptNames2() As String
                If SapModel.AreaObj.GetPoints(areaName, numPts2, ptNames2) = 0 Then
                    Dim minX As Double, maxX As Double, minY As Double, maxY As Double
                    minX = 1E+30: maxX = -1E+30: minY = 1E+30: maxY = -1E+30
                    
                    Dim j As Long
                    For j = 0 To numPts2 - 1
                        Dim px As Double, py As Double, pz As Double
                        If SapModel.pointObj.GetCoordCartesian(ptNames2(j), px, py, pz) = 0 Then
                            If px < minX Then minX = px
                            If px > maxX Then maxX = px
                            If py < minY Then minY = py
                            If py > maxY Then maxY = py
                        End If
                    Next j
                    
                    Dim bboxW As Double, bboxH As Double
                    bboxW = maxX - minX: bboxH = maxY - minY
                    
                    ' Large polygon -> likely slab
                    If bboxW > 100 Or bboxH > 100 Or numPts2 > 8 Then
                        acceptArea = True
                    Else
                        skippedCount = skippedCount + 1
                        GoTo NextArea
                    End If
                End If
            Else
                Thickness = GetAreaSectionThickness(propName)
                If Thickness > 0 Then acceptArea = True
            End If
        End If

        If acceptArea Then
            If Thickness <= 0 And Trim$(propName) <> "" And UCase$(Trim$(propName)) <> "NONE" Then
                Thickness = GetAreaSectionThickness(propName)
            End If

            Dim numPts As Long
            Dim ptNames() As String

            If SapModel.AreaObj.GetPoints(areaName, numPts, ptNames) = 0 And numPts >= 3 Then
                With g_SlabAreas(idx)
                    .Name = areaName
                    .Thickness = Thickness
                    .numPoints = numPts
                    ReDim .PointsX(0 To numPts - 1)
                    ReDim .PointsY(0 To numPts - 1)

                    Dim sumZ As Double: sumZ = 0
                    For j = 0 To numPts - 1
                        If SapModel.pointObj.GetCoordCartesian(ptNames(j), px, py, pz) = 0 Then
                            .PointsX(j) = px
                            .PointsY(j) = py
                            sumZ = sumZ + pz
                        End If
                    Next j
                    .elevation = sumZ / numPts
                End With
                idx = idx + 1
            Else
                skippedCount = skippedCount + 1
            End If
        End If

NextArea:
    Next i

    g_SlabAreasCount = idx
    If idx > 0 And idx < numAreas Then ReDim Preserve g_SlabAreas(0 To idx - 1)
    g_SlabAreasCached = True
    LogMsg "Cached " & g_SlabAreasCount & " slab areas (skipped " & skippedCount & " areas)"
    Exit Sub

ErrorHandler:
    LogMsg "ERROR in M12A_InitializeSlabCache: " & err.description
End Sub

'===============================================================
' COLUMN AREA FILTERING
'===============================================================
Private Function GetColumnPrefixesFromSettings() As Object
    On Error Resume Next
    Set GetColumnPrefixesFromSettings = CreateObject("Scripting.Dictionary")

    If M12_DrawColumnAreas.g_Settings Is Nothing Then Exit Function
    If Not M12_DrawColumnAreas.g_Settings.exists("SectionPrefixes") Then Exit Function

    Dim prefixDict As Object
    Set prefixDict = M12_DrawColumnAreas.g_Settings("SectionPrefixes")
    If prefixDict Is Nothing Then Exit Function

    Dim key As Variant
    For Each key In prefixDict.keys
        On Error Resume Next
        GetColumnPrefixesFromSettings.Add CStr(key), CStr(prefixDict(key))
        On Error GoTo 0
    Next key

    If GetColumnPrefixesFromSettings.count = 0 Then
        GetColumnPrefixesFromSettings.Add "RC", "RC_COL_"
        GetColumnPrefixesFromSettings.Add "STEEL_I", "STEEL_I_"
        GetColumnPrefixesFromSettings.Add "STEEL_H", "STEEL_H_"
        GetColumnPrefixesFromSettings.Add "STEEL_PIPE", "STEEL_PIPE_"
        GetColumnPrefixesFromSettings.Add "STEEL_BOX", "STEEL_BOX_"
        GetColumnPrefixesFromSettings.Add "STEEL_CH", "STEEL_CH_"
        GetColumnPrefixesFromSettings.Add "DEFAULT", "ColSec_"
        GetColumnPrefixesFromSettings.Add "PUNCH", "Punch_"
    End If
End Function

Private Function IsColumnCrossSectionArea(areaName As String, columnPrefixes As Object) As Boolean
    On Error Resume Next
    IsColumnCrossSectionArea = False
    If columnPrefixes Is Nothing Then Exit Function

    Dim areaNameUpper As String
    areaNameUpper = UCase$(Trim$(areaName))

    Dim key As Variant
    For Each key In columnPrefixes.keys
        Dim prefix As String
        prefix = UCase$(Trim$(CStr(columnPrefixes(key))))
        If prefix <> "" Then
            If Left$(areaNameUpper, Len(prefix)) = prefix Then
                IsColumnCrossSectionArea = True
                Exit Function
            End If
        End If
    Next key

    Dim numGroups As Long
    Dim groupNames() As String
    If SapModel.AreaObj.GetGroupAssign(areaName, numGroups, groupNames) = 0 Then
        If numGroups > 0 Then
            Dim i As Long
            For i = 0 To numGroups - 1
                If UCase$(Trim$(groupNames(i))) = "COLUMNCROSSSECTIONS" Then
                    IsColumnCrossSectionArea = True
                    Exit Function
                End If
            Next i
        End If
    End If

    If InStr(areaNameUpper, "COLSEC_") > 0 Or _
       InStr(areaNameUpper, "COLXSEC_") > 0 Or _
       InStr(areaNameUpper, "PUNCH_") > 0 Or _
       Left$(areaNameUpper, 5) = "COLUM" Then
        IsColumnCrossSectionArea = True
    End If
End Function

'===============================================================
' MAIN API - SEGMENT CLIPPING BASED
'===============================================================
Public Sub M12A_DrawPunchingPerimeter(frameName As String, centerX As Double, centerY As Double, _
        elevation As Double, colWidth As Double, colDepth As Double, _
        offsetT2 As Double, offsetT3 As Double, _
        t2AngleRad As Double, t3AngleRad As Double, _
        shapeType As String, storyIndex As Long, _
        sectionName As String, columnAreaName As String)
    On Error GoTo ErrorHandler

    If Not g_SlabAreasCached Then Call M12A_InitializeSlabCache

    ' Calculate actual column center with offset
    Dim cos2 As Double, sin2 As Double, cos3 As Double, sin3 As Double
    cos2 = Cos(t2AngleRad): sin2 = Sin(t2AngleRad)
    cos3 = Cos(t3AngleRad): sin3 = Sin(t3AngleRad)

    Dim rotatedOffsetX As Double, rotatedOffsetY As Double
    rotatedOffsetX = offsetT2 * cos2 + offsetT3 * cos3
    rotatedOffsetY = offsetT2 * sin2 + offsetT3 * sin3

    Dim actualX As Double, actualY As Double
    actualX = centerX + rotatedOffsetX
    actualY = centerY + rotatedOffsetY

    ' Get actual column dimensions from drawn area
    Dim actualColW As Double, actualColD As Double
    If Trim$(columnAreaName) <> "" Then
        If TryGetColumnAreaProjectedExtents(columnAreaName, cos2, sin2, cos3, sin3, actualX, actualY, actualColW, actualColD) Then
            colWidth = actualColW
            colDepth = actualColD
        End If
    End If

    ' Detect attached slabs
    Dim attachedSlabs As Object
    Set attachedSlabs = DetectAttachedSlabsByVertex(actualX, actualY, elevation)
    
    If attachedSlabs.count = 0 Then
        LogMsg "WARNING: No slab attached to column " & frameName & " at (" & _
               Format(actualX, "0.0") & "," & Format(actualY, "0.0") & ") - skipping"
        Exit Sub
    End If
    
    LogMsg "Column " & frameName & " attached to " & attachedSlabs.count & " slab(s)"

    ' Get effective depth from attached slabs
    Dim slabThickness As Double
    slabThickness = GetThicknessFromAttachedSlabs(attachedSlabs)
    
    If slabThickness <= 0 Then
        LogMsg "WARNING: Cannot determine slab thickness for column " & frameName
        Exit Sub
    End If

    ' Get cover and calculate effective depth
    Dim topCover As Double, bottomCover As Double
    Call GetCoverForLocation(storyIndex, sectionName, elevation, actualX, actualY, topCover, bottomCover)

    Dim effectiveDepth As Double
    effectiveDepth = slabThickness - topCover
    
    If effectiveDepth < MIN_EFFECTIVE_DEPTH Then
        LogMsg "WARNING: d_eff=" & Format(effectiveDepth, "0.0") & "mm < " & _
               MIN_EFFECTIVE_DEPTH & "mm for " & frameName
        Exit Sub
    End If

    Dim punchDist As Double
    punchDist = effectiveDepth / 2#

    ' Get column section outline
    Dim sectionPoints() As Double
    Dim sectionCount As Long
    sectionCount = GetColumnSectionOutline(columnAreaName, actualX, actualY, elevation, sectionPoints)

    If sectionCount < 3 Then
        LogMsg "Using rectangular outline for " & frameName
        sectionCount = 4
        ReDim sectionPoints(0 To 3, 0 To 1)
        Call CalculateRectangleCorners(actualX, actualY, colWidth, colDepth, cos2, sin2, cos3, sin3, sectionPoints)
    End If

    ' Offset polygon outward by d/2
    Dim offsetPoints() As Double
    Dim offsetCount As Long
    offsetCount = OffsetPolygonOutward(sectionPoints, sectionCount, punchDist, offsetPoints)

    If offsetCount < 3 Then
        LogMsg "WARNING: Offset failed for " & frameName
        Exit Sub
    End If

    ' Build attached slab index array
    Dim slabIndices() As Long
    Dim k As Long
    Dim key As Variant
    ReDim slabIndices(0 To attachedSlabs.count - 1)
    k = 0
    For Each key In attachedSlabs.keys
        slabIndices(k) = CLng(key)
        k = k + 1
    Next key

    ' Clip offset polygon with attached slabs (segment-based clipping)
    Dim finalPoints() As Double
    Dim finalCount As Long
    finalCount = ClipOffsetPolygonWithSlabs(offsetPoints, offsetCount, slabIndices, attachedSlabs.count)

    If finalCount < 2 Then
        LogMsg "WARNING: Clipping failed for " & frameName & ", using raw offset outline"
        finalCount = offsetCount
        ReDim finalPoints(0 To offsetCount - 1, 0 To 1)
        Dim i As Long
        For i = 0 To offsetCount - 1
            finalPoints(i, 0) = offsetPoints(i, 0)
            finalPoints(i, 1) = offsetPoints(i, 1)
        Next i
    Else
        ' Copy points from internal result
        ReDim finalPoints(0 To finalCount - 1, 0 To 1)
        Dim tmp() As Double
        ReDim tmp(0 To finalCount - 1, 0 To 1)
        Call GetLastClippedPolygon(tmp, finalCount)
        For i = 0 To finalCount - 1
            finalPoints(i, 0) = tmp(i, 0)
            finalPoints(i, 1) = tmp(i, 1)
        Next i
    End If

    LogMsg "Column " & frameName & ": Perimeter=" & finalCount & " pts, d_eff=" & _
           Format(effectiveDepth, "0.0") & "mm, slabs=" & attachedSlabs.count

    ' Draw perimeter
    Dim punchingName As String
    punchingName = "Punch_" & columnAreaName

    Dim outType As String: outType = "AREA"
    On Error Resume Next
    If Not M12_DrawColumnAreas.g_Settings Is Nothing Then
        If M12_DrawColumnAreas.g_Settings.exists("PunchingOutputType") Then
            outType = CStr(M12_DrawColumnAreas.g_Settings("PunchingOutputType"))
        End If
    End If
    On Error GoTo ErrorHandler

    If UCase$(outType) = "AREA" Then
        ' AREA requires at least 3 points
        If finalCount >= 3 Then
            Call DrawPerimeterAsArea(punchingName, finalPoints, finalCount, elevation)
        Else
            LogMsg "WARNING: Not enough points to create AREA for " & frameName & ", drawing LINE instead"
            If finalCount >= 2 Then
                Call DrawPerimeterAsLines(punchingName, finalPoints, finalCount, elevation)
            End If
        End If
    Else
        ' LINE mode can work with open or closed polygon
        If finalCount >= 2 Then
            Call DrawPerimeterAsLines(punchingName, finalPoints, finalCount, elevation)
        Else
            LogMsg "WARNING: Not enough points to create LINE for " & frameName
        End If
    End If

    Exit Sub
ErrorHandler:
    LogMsg "ERROR in M12A_DrawPunchingPerimeter: " & err.description
End Sub

'===============================================================
' DETECT SLABS BY VERTEX / EDGE / INSIDE
'===============================================================
Private Function DetectAttachedSlabsByVertex(colX As Double, colY As Double, elevation As Double) As Object
    ' Returns Dictionary with keys = slab indices, values = slab names
    On Error GoTo ErrorHandler
    
    Set DetectAttachedSlabsByVertex = CreateObject("Scripting.Dictionary")
    
    If g_SlabAreasCount = 0 Then Exit Function
    
    Dim i As Long, j As Long
    For i = 0 To g_SlabAreasCount - 1
        With g_SlabAreas(i)
            ' Check elevation match
            If Abs(.elevation - elevation) > SLAB_ELEV_TOL Then GoTo NextSlab
            
            ' Check if column point is a vertex of slab polygon
            Dim isVertex As Boolean: isVertex = False
            For j = 0 To .numPoints - 1
                Dim dx As Double, dy As Double
                dx = .PointsX(j) - colX
                dy = .PointsY(j) - colY
                Dim dist As Double
                dist = Sqr(dx * dx + dy * dy)
                
                If dist <= VERTEX_MATCH_TOL Then
                    isVertex = True
                    Exit For
                End If
            Next j
            
            If isVertex Then
                DetectAttachedSlabsByVertex.Add i, .Name
                LogMsg "  -> Found attached slab: " & .Name & " (thickness=" & .Thickness & "mm)"
                GoTo NextSlab
            End If
            
            ' Also check if point is on edge
            For j = 0 To .numPoints - 1
                Dim jNext As Long: jNext = (j + 1) Mod .numPoints
                Dim edgeDist As Double
                edgeDist = PointToSegmentDistance(colX, colY, _
                    .PointsX(j), .PointsY(j), _
                    .PointsX(jNext), .PointsY(jNext))
                
                If edgeDist <= EDGE_MATCH_TOL Then
                    DetectAttachedSlabsByVertex.Add i, .Name
                    LogMsg "  -> Found attached slab (edge): " & .Name
                    Exit For
                End If
            Next j
            
            ' Fallback: check inside polygon (internal column)
            If Not DetectAttachedSlabsByVertex.exists(i) Then
                If IsPointInPolygon(colX, colY, .PointsX, .PointsY, .numPoints) Then
                    DetectAttachedSlabsByVertex.Add i, .Name
                    LogMsg "  -> Found attached slab (inside): " & .Name
                End If
            End If
        End With
NextSlab:
    Next i
    
    Exit Function
ErrorHandler:
    LogMsg "ERROR in DetectAttachedSlabsByVertex: " & err.description
    Set DetectAttachedSlabsByVertex = CreateObject("Scripting.Dictionary")
End Function

'===============================================================
' GET THICKNESS FROM ATTACHED SLABS
'===============================================================
Private Function GetThicknessFromAttachedSlabs(attachedSlabs As Object) As Double
    On Error Resume Next
    GetThicknessFromAttachedSlabs = 0
    
    If attachedSlabs.count = 0 Then Exit Function
    
    ' Return minimum thickness from attached slabs
    Dim minThick As Double: minThick = 1E+30
    Dim key As Variant
    
    For Each key In attachedSlabs.keys
        Dim idx As Long: idx = CLng(key)
        If g_SlabAreas(idx).Thickness > 0 Then
            If g_SlabAreas(idx).Thickness < minThick Then
                minThick = g_SlabAreas(idx).Thickness
            End If
        End If
    Next key
    
    If minThick < 1E+30 Then GetThicknessFromAttachedSlabs = minThick
End Function


Private Function ClipOffsetPolygonWithSlabs(offsetPoly() As Double, offsetCount As Long, _
        slabIndices() As Long, slabCount As Long) As Long
    On Error GoTo ErrH
    ClipOffsetPolygonWithSlabs = 0
    g_LastClipCount = 0

    If offsetCount < 2 Or slabCount <= 0 Then Exit Function

    ' Copy initial polygon
    Dim curPoly() As Double
    Dim curCount As Long
    ReDim curPoly(0 To offsetCount - 1, 0 To 1)
    Dim i As Long
    For i = 0 To offsetCount - 1
        curPoly(i, 0) = offsetPoly(i, 0)
        curPoly(i, 1) = offsetPoly(i, 1)
    Next i
    curCount = offsetCount

    ' Clip sequentially with each attached slab polygon (intersection)
    Dim sIndex As Long
    For sIndex = 0 To slabCount - 1
        Dim idx As Long
        idx = slabIndices(sIndex)
        With g_SlabAreas(idx)
            Dim outPoly() As Double
            Dim outCount As Long
            outCount = PolygonIntersectionConvexApprox(curPoly, curCount, .PointsX, .PointsY, .numPoints, outPoly)
            If outCount <= 0 Then
                ' No intersection with this slab, continue with next slab.
                ' We do not reset curPoly to keep intersection with other slabs if already exists.
            Else
                ' Replace current polygon with intersection result
                ReDim curPoly(0 To outCount - 1, 0 To 1)
                For i = 0 To outCount - 1
                    curPoly(i, 0) = outPoly(i, 0)
                    curPoly(i, 1) = outPoly(i, 1)
                Next i
                curCount = outCount
            End If
        End With
    Next sIndex

    ' As fallback, if no clipping produced result, try to keep those vertices inside any slab
    If curCount <= 0 Then
        Dim tmpPoly() As Double
        Dim keptCount As Long
        keptCount = KeepVerticesInsideSlabs(offsetPoly, offsetCount, slabIndices, slabCount, tmpPoly)
        If keptCount <= 0 Then
            ClipOffsetPolygonWithSlabs = 0
            Exit Function
        End If
        ReDim curPoly(0 To keptCount - 1, 0 To 1)
        For i = 0 To keptCount - 1
            curPoly(i, 0) = tmpPoly(i, 0)
            curPoly(i, 1) = tmpPoly(i, 1)
        Next i
        curCount = keptCount
    End If

    ' Remove degenerate segments and duplicate points
    Dim cleanPoly() As Double
    Dim cleanCount As Long
    cleanCount = CleanupPolygon(curPoly, curCount, cleanPoly)

    If cleanCount <= 0 Then
        ClipOffsetPolygonWithSlabs = 0
        Exit Function
    End If

    ' Store to global buffer for later retrieval
    g_LastClipCount = cleanCount
    ReDim g_LastClipPoly(0 To cleanCount - 1, 0 To 1)
    For i = 0 To cleanCount - 1
        g_LastClipPoly(i, 0) = cleanPoly(i, 0)
        g_LastClipPoly(i, 1) = cleanPoly(i, 1)
    Next i

    ClipOffsetPolygonWithSlabs = cleanCount
    Exit Function
ErrH:
    LogMsg "ERROR in ClipOffsetPolygonWithSlabs: " & err.description
    ClipOffsetPolygonWithSlabs = 0
End Function

Private Sub GetLastClippedPolygon(ByRef outPoly() As Double, ByRef outCount As Long)
    outCount = g_LastClipCount
    If outCount <= 0 Then Exit Sub
    Dim i As Long
    For i = 0 To outCount - 1
        outPoly(i, 0) = g_LastClipPoly(i, 0)
        outPoly(i, 1) = g_LastClipPoly(i, 1)
    Next i
End Sub

' Keep only vertices of offsetPoly that are inside or on edges of attached slabs
Private Function KeepVerticesInsideSlabs(offsetPoly() As Double, offsetCount As Long, _
        slabIndices() As Long, slabCount As Long, _
        ByRef outPoly() As Double) As Long
    On Error GoTo ErrH
    KeepVerticesInsideSlabs = 0

    Dim tmp() As Double
    ReDim tmp(0 To offsetCount - 1, 0 To 1)
    Dim count As Long: count = 0
    Dim i As Long

    For i = 0 To offsetCount - 1
        Dim X As Double, Y As Double
        X = offsetPoly(i, 0)
        Y = offsetPoly(i, 1)

        If IsPointOnAnySlabOrInside(X, Y, slabIndices, slabCount) Then
            tmp(count, 0) = X
            tmp(count, 1) = Y
            count = count + 1
        End If
    Next i

    If count <= 0 Then Exit Function

    ReDim outPoly(0 To count - 1, 0 To 1)
    For i = 0 To count - 1
        outPoly(i, 0) = tmp(i, 0)
        outPoly(i, 1) = tmp(i, 1)
    Next i

    KeepVerticesInsideSlabs = count
    Exit Function
ErrH:
    KeepVerticesInsideSlabs = 0
End Function

Private Function IsPointOnAnySlabOrInside(X As Double, Y As Double, _
        slabIndices() As Long, slabCount As Long) As Boolean
    Dim si As Long
    For si = 0 To slabCount - 1
        With g_SlabAreas(slabIndices(si))
            If IsPointInPolygon(X, Y, .PointsX, .PointsY, .numPoints) Then
                IsPointOnAnySlabOrInside = True
                Exit Function
            End If
            If IsPointOnPolygonEdge(X, Y, .PointsX, .PointsY, .numPoints, COORD_TOLERANCE) Then
                IsPointOnAnySlabOrInside = True
                Exit Function
            End If
        End With
    Next si
    IsPointOnAnySlabOrInside = False
End Function

' Approximate polygon intersection by clipping offset polygon against convex hull-like edges of slab polygon
Private Function PolygonIntersectionConvexApprox(polyIn() As Double, inCount As Long, _
        slabX() As Double, slabY() As Double, slabCount As Long, _
        ByRef polyOut() As Double) As Long
    On Error GoTo ErrH
    PolygonIntersectionConvexApprox = 0

    If inCount <= 0 Or slabCount < 3 Then Exit Function

    ' We will apply Sutherland–Hodgman style clipping using each slab edge
    Dim workPoly1() As Double, workPoly2() As Double
    Dim workCount1 As Long, workCount2 As Long
    Dim i As Long

    ReDim workPoly1(0 To inCount - 1, 0 To 1)
    For i = 0 To inCount - 1
        workPoly1(i, 0) = polyIn(i, 0)
        workPoly1(i, 1) = polyIn(i, 1)
    Next i
    workCount1 = inCount

    Dim e As Long
    For e = 0 To slabCount - 1
        Dim ex1 As Double, ey1 As Double, eX2 As Double, eY2 As Double
        ex1 = slabX(e)
        ey1 = slabY(e)
        eX2 = slabX((e + 1) Mod slabCount)
        eY2 = slabY((e + 1) Mod slabCount)

        ' Left side of edge is considered inside
        Dim nx As Double, ny As Double
        nx = eY2 - ey1
        ny = -(eX2 - ex1)

        ' If work polygon empty, stop
        If workCount1 <= 0 Then Exit For

        ' Clip current polygon against this edge
        workCount2 = 0
        ReDim workPoly2(0 To workCount1 * 2 + slabCount, 0 To 1)

        Dim p0x As Double, p0y As Double, p1x As Double, p1y As Double
        Dim j As Long
        For j = 0 To workCount1 - 1
            Dim k As Long
            k = (j + 1) Mod workCount1

            p0x = workPoly1(j, 0): p0y = workPoly1(j, 1)
            p1x = workPoly1(k, 0): p1y = workPoly1(k, 1)

            Dim d0 As Double, d1 As Double
            d0 = (p0x - ex1) * nx + (p0y - ey1) * ny
            d1 = (p1x - ex1) * nx + (p1y - ey1) * ny

            Dim inside0 As Boolean, inside1 As Boolean
            inside0 = (d0 >= -COORD_TOLERANCE)
            inside1 = (d1 >= -COORD_TOLERANCE)

            If inside0 And inside1 Then
                ' Both inside: keep second point
                workPoly2(workCount2, 0) = p1x
                workPoly2(workCount2, 1) = p1y
                workCount2 = workCount2 + 1
            ElseIf inside0 And Not inside1 Then
                ' Leaving: keep intersection
                Dim ix As Double, iy As Double
                If ClipLineToEdge(p0x, p0y, p1x, p1y, ex1, ey1, eX2, eY2, ix, iy) Then
                    workPoly2(workCount2, 0) = ix
                    workPoly2(workCount2, 1) = iy
                    workCount2 = workCount2 + 1
                End If
            ElseIf (Not inside0) And inside1 Then
                ' Entering: keep intersection and second
                Dim ix2 As Double, iy2 As Double
                If ClipLineToEdge(p0x, p0y, p1x, p1y, ex1, ey1, eX2, eY2, ix2, iy2) Then
                    workPoly2(workCount2, 0) = ix2
                    workPoly2(workCount2, 1) = iy2
                    workCount2 = workCount2 + 1
                End If
                workPoly2(workCount2, 0) = p1x
                workPoly2(workCount2, 1) = p1y
                workCount2 = workCount2 + 1
            Else
                ' Both outside: keep nothing
            End If
        Next j

        ' Prepare for next edge
        If workCount2 <= 0 Then
            workCount1 = 0
            Exit For
        End If

        ReDim workPoly1(0 To workCount2 - 1, 0 To 1)
        For i = 0 To workCount2 - 1
            workPoly1(i, 0) = workPoly2(i, 0)
            workPoly1(i, 1) = workPoly2(i, 1)
        Next i
        workCount1 = workCount2
    Next e

    If workCount1 <= 0 Then
        PolygonIntersectionConvexApprox = 0
        Exit Function
    End If

    ' Output
    ReDim polyOut(0 To workCount1 - 1, 0 To 1)
    For i = 0 To workCount1 - 1
        polyOut(i, 0) = workPoly1(i, 0)
        polyOut(i, 1) = workPoly1(i, 1)
    Next i
    PolygonIntersectionConvexApprox = workCount1
    Exit Function
ErrH:
    PolygonIntersectionConvexApprox = 0
End Function

' Compute intersection between segment P0-P1 and infinite line defined by edge E0-E1
Private Function ClipLineToEdge(p0x As Double, p0y As Double, p1x As Double, p1y As Double, _
        ex1 As Double, ey1 As Double, eX2 As Double, eY2 As Double, _
        ByRef ix As Double, ByRef iy As Double) As Boolean
    On Error GoTo ErrH
    ClipLineToEdge = False

    Dim dxp As Double, dyp As Double
    dxp = p1x - p0x
    dyp = p1y - p0y

    Dim dxE As Double, dyE As Double
    dxE = eX2 - ex1
    dyE = eY2 - ey1

    Dim denom As Double
    denom = dxp * dyE - dyp * dxE
    If Abs(denom) < EPSILON Then Exit Function

    Dim t As Double
    t = ((ex1 - p0x) * dyE - (ey1 - p0y) * dxE) / denom

    If t < -0.001 Or t > 1.001 Then Exit Function

    ix = p0x + t * dxp
    iy = p0y + t * dyp
    ClipLineToEdge = True
    Exit Function
ErrH:
    ClipLineToEdge = False
End Function

' Remove duplicate points and nearly zero-length segments
Private Function CleanupPolygon(polyIn() As Double, inCount As Long, _
        ByRef polyOut() As Double) As Long
    On Error GoTo ErrH
    CleanupPolygon = 0
    If inCount <= 0 Then Exit Function

    Dim tmp() As Double
    ReDim tmp(0 To inCount - 1, 0 To 1)

    Dim used() As Boolean
    ReDim used(0 To inCount - 1)

    Dim i As Long
    Dim count As Long: count = 0

    ' Remove consecutive duplicates
    For i = 0 To inCount - 1
        Dim j As Long
        j = (i + 1) Mod inCount
        Dim dx As Double, dy As Double
        dx = polyIn(j, 0) - polyIn(i, 0)
        dy = polyIn(j, 1) - polyIn(i, 1)
        If (dx * dx + dy * dy) > COORD_TOLERANCE * COORD_TOLERANCE Then
            tmp(count, 0) = polyIn(i, 0)
            tmp(count, 1) = polyIn(i, 1)
            count = count + 1
        End If
    Next i

    If count <= 1 Then
        CleanupPolygon = count
        If count > 0 Then
            ReDim polyOut(0 To 0, 0 To 1)
            polyOut(0, 0) = tmp(0, 0)
            polyOut(0, 1) = tmp(0, 1)
        End If
        Exit Function
    End If

    ' Optional: remove last if equal to first
    Dim dxLast As Double, dyLast As Double
    dxLast = tmp(count - 1, 0) - tmp(0, 0)
    dyLast = tmp(count - 1, 1) - tmp(0, 1)
    If (dxLast * dxLast + dyLast * dyLast) <= (COORD_TOLERANCE * COORD_TOLERANCE) Then
        count = count - 1
    End If

    If count <= 1 Then
        CleanupPolygon = count
        If count > 0 Then
            ReDim polyOut(0 To 0, 0 To 1)
            polyOut(0, 0) = tmp(0, 0)
            polyOut(0, 1) = tmp(0, 1)
        End If
        Exit Function
    End If

    ReDim polyOut(0 To count - 1, 0 To 1)
    For i = 0 To count - 1
        polyOut(i, 0) = tmp(i, 0)
        polyOut(i, 1) = tmp(i, 1)
    Next i

    CleanupPolygon = count
    Exit Function
ErrH:
    CleanupPolygon = 0
End Function

'===============================================================
' HELPER FUNCTIONS
'===============================================================
Private Function GetColumnSectionOutline(areaName As String, centerX As Double, centerY As Double, _
        elevation As Double, ByRef outline() As Double) As Long
    On Error GoTo ErrorHandler
    GetColumnSectionOutline = 0
    If Trim$(areaName) = "" Then Exit Function

    Dim numPts As Long
    Dim ptNames() As String
    If SapModel.AreaObj.GetPoints(areaName, numPts, ptNames) <> 0 Then Exit Function
    If numPts < 3 Then Exit Function

    ReDim outline(0 To numPts - 1, 0 To 1)

    Dim i As Long
    Dim px As Double, py As Double, pz As Double
    For i = 0 To numPts - 1
        If SapModel.pointObj.GetCoordCartesian(ptNames(i), px, py, pz) = 0 Then
            outline(i, 0) = px
            outline(i, 1) = py
        End If
    Next i

    GetColumnSectionOutline = numPts
    Exit Function
ErrorHandler:
    GetColumnSectionOutline = 0
End Function

Private Function OffsetPolygonOutward(inputPoly() As Double, inputCount As Long, _
        offsetDist As Double, ByRef outputPoly() As Double) As Long
    On Error GoTo ErrorHandler
    OffsetPolygonOutward = 0
    If inputCount < 3 Or offsetDist <= 0 Then Exit Function

    ReDim outputPoly(0 To inputCount - 1, 0 To 1)

    Dim i As Long
    For i = 0 To inputCount - 1
        Dim prevIdx As Long, nextIdx As Long
        prevIdx = IIf(i = 0, inputCount - 1, i - 1)
        nextIdx = IIf(i = inputCount - 1, 0, i + 1)

        Dim v1x As Double, v1y As Double
        Dim v2x As Double, v2y As Double

        v1x = inputPoly(i, 0) - inputPoly(prevIdx, 0)
        v1y = inputPoly(i, 1) - inputPoly(prevIdx, 1)
        v2x = inputPoly(nextIdx, 0) - inputPoly(i, 0)
        v2y = inputPoly(nextIdx, 1) - inputPoly(i, 1)

        Dim len1 As Double, len2 As Double
        len1 = Sqr(v1x * v1x + v1y * v1y)
        len2 = Sqr(v2x * v2x + v2y * v2y)

        If len1 > EPSILON Then
            v1x = v1x / len1
            v1y = v1y / len1
        End If
        If len2 > EPSILON Then
            v2x = v2x / len2
            v2y = v2y / len2
        End If

        Dim n1x As Double, n1y As Double
        Dim n2x As Double, n2y As Double
        n1x = -v1y: n1y = v1x
        n2x = -v2y: n2y = v2x

        Dim nx As Double, ny As Double
        nx = (n1x + n2x) / 2#
        ny = (n1y + n2y) / 2#

        Dim nlen As Double
        nlen = Sqr(nx * nx + ny * ny)
        If nlen > EPSILON Then
            nx = nx / nlen
            ny = ny / nlen
        End If

        outputPoly(i, 0) = inputPoly(i, 0) + nx * offsetDist
        outputPoly(i, 1) = inputPoly(i, 1) + ny * offsetDist
    Next i

    OffsetPolygonOutward = inputCount
    Exit Function
ErrorHandler:
    OffsetPolygonOutward = 0
End Function

Private Function PointToSegmentDistance(ptX As Double, ptY As Double, _
        x1 As Double, y1 As Double, x2 As Double, y2 As Double) As Double
    Dim cx As Double, cy As Double
    Call ClosestPointOnSegment(ptX, ptY, x1, y1, x2, y2, cx, cy)
    PointToSegmentDistance = Sqr((cx - ptX) ^ 2 + (cy - ptY) ^ 2)
End Function

Private Sub ClosestPointOnSegment(ptX As Double, ptY As Double, _
        x1 As Double, y1 As Double, x2 As Double, y2 As Double, _
        ByRef closestX As Double, ByRef closestY As Double)
    Dim dx As Double, dy As Double
    dx = x2 - x1: dy = y2 - y1
    
    Dim len2 As Double
    len2 = dx * dx + dy * dy
    
    If len2 < EPSILON Then
        closestX = x1: closestY = y1
        Exit Sub
    End If
    
    Dim t As Double
    t = ((ptX - x1) * dx + (ptY - y1) * dy) / len2
    
    If t < 0 Then t = 0
    If t > 1 Then t = 1
    
    closestX = x1 + t * dx
    closestY = y1 + t * dy
End Sub

Private Function IsPointOnPolygonEdge(X As Double, Y As Double, _
        ptsX() As Double, ptsY() As Double, numPts As Long, _
        tolerance As Double) As Boolean
    IsPointOnPolygonEdge = False
    If numPts < 2 Then Exit Function
    
    Dim i As Long
    For i = 0 To numPts - 1
        Dim jNext As Long: jNext = (i + 1) Mod numPts
        If PointToSegmentDistance(X, Y, ptsX(i), ptsY(i), ptsX(jNext), ptsY(jNext)) <= tolerance Then
            IsPointOnPolygonEdge = True
            Exit Function
        End If
    Next i
End Function

Private Function IsPointInPolygon(X As Double, Y As Double, _
        ptsX() As Double, ptsY() As Double, numPts As Long) As Boolean
    IsPointInPolygon = False
    If numPts < 3 Then Exit Function
    
    Dim inside As Boolean: inside = False
    Dim j As Long: j = numPts - 1
    Dim i As Long
    
    For i = 0 To numPts - 1
        Dim xi As Double, yi As Double, xj As Double, yj As Double
        xi = ptsX(i): yi = ptsY(i)
        xj = ptsX(j): yj = ptsY(j)
        
        If ((yi > Y) <> (yj > Y)) Then
            Dim denom As Double
            denom = (yj - yi)
            If Abs(denom) > EPSILON Then
                Dim xint As Double
                xint = (xj - xi) * (Y - yi) / denom + xi
                If X < xint Then inside = Not inside
            End If
        End If
        j = i
    Next i
    
    IsPointInPolygon = inside
End Function

Private Sub CalculateRectangleCorners(centerX As Double, centerY As Double, _
        Width As Double, Depth As Double, _
        cos2 As Double, sin2 As Double, cos3 As Double, sin3 As Double, _
        ByRef corners() As Double)
    Dim localX(0 To 3) As Double, localY(0 To 3) As Double
    
    localX(0) = -Width / 2#: localY(0) = -Depth / 2#
    localX(1) = Width / 2#:  localY(1) = -Depth / 2#
    localX(2) = Width / 2#:  localY(2) = Depth / 2#
    localX(3) = -Width / 2#: localY(3) = Depth / 2#
    
    Dim i As Long
    For i = 0 To 3
        corners(i, 0) = centerX + localX(i) * cos2 + localY(i) * cos3
        corners(i, 1) = centerY + localX(i) * sin2 + localY(i) * sin3
    Next i
End Sub

Private Function TryGetColumnAreaProjectedExtents(areaName As String, _
        cos2 As Double, sin2 As Double, cos3 As Double, sin3 As Double, _
        centerX As Double, centerY As Double, _
        ByRef outWidth As Double, ByRef outDepth As Double) As Boolean
    On Error GoTo ErrHandler
    TryGetColumnAreaProjectedExtents = False
    outWidth = 0: outDepth = 0
    
    Dim numPts As Long
    Dim ptNames() As String
    If SapModel.AreaObj.GetPoints(areaName, numPts, ptNames) <> 0 Then Exit Function
    If numPts < 3 Then Exit Function
    
    Dim i As Long
    Dim px As Double, py As Double, pz As Double
    Dim proj2 As Double, proj3 As Double
    Dim min2 As Double, max2 As Double, min3 As Double, max3 As Double
    min2 = 1E+30: max2 = -1E+30
    min3 = 1E+30: max3 = -1E+30
    
    For i = 0 To numPts - 1
        If SapModel.pointObj.GetCoordCartesian(ptNames(i), px, py, pz) = 0 Then
            proj2 = (px - centerX) * cos2 + (py - centerY) * sin2
            proj3 = (px - centerX) * cos3 + (py - centerY) * sin3
            If proj2 < min2 Then min2 = proj2
            If proj2 > max2 Then max2 = proj2
            If proj3 < min3 Then min3 = proj3
            If proj3 > max3 Then max3 = proj3
        End If
    Next i
    
    If min2 > max2 Or min3 > max3 Then Exit Function
    
    outWidth = max2 - min2
    outDepth = max3 - min3
    TryGetColumnAreaProjectedExtents = True
    Exit Function
ErrHandler:
    TryGetColumnAreaProjectedExtents = False
End Function

Private Function Atan2Func(Y As Double, X As Double) As Double
    If Abs(X) < EPSILON Then
        If Y > 0 Then
            Atan2Func = PI / 2#
        ElseIf Y < 0 Then
            Atan2Func = -PI / 2#
        Else
            Atan2Func = 0#
        End If
        Exit Function
    End If
    
    Atan2Func = Atn(Y / X)
    If X < 0 Then
        If Y >= 0 Then
            Atan2Func = Atan2Func + PI
        Else
            Atan2Func = Atan2Func - PI
        End If
    End If
End Function

Private Sub DrawPerimeterAsArea(punchingName As String, points() As Double, _
        numPoints As Long, elevation As Double)
    On Error GoTo ErrorHandler
    
    Dim pointNames() As String
    ReDim pointNames(0 To numPoints - 1)
    
    Dim i As Long
    For i = 0 To numPoints - 1
        g_UniqueCounter = g_UniqueCounter + 1
        Dim ptName As String
        ptName = punchingName & "_P" & CStr(g_UniqueCounter)
        
        On Error Resume Next
        SapModel.pointObj.DeleteSpecialPoint ptName
        On Error GoTo ErrorHandler
        
        Dim actualPtName As String
        actualPtName = ptName
        Dim ret As Long
        ret = SapModel.pointObj.AddCartesian(points(i, 0), points(i, 1), elevation, actualPtName, ptName)
        
        If ret = 0 Then
            pointNames(i) = actualPtName
        Else
            LogMsg "ERROR: Failed to create point " & ptName
            Exit Sub
        End If
    Next i
    
    Dim actualAreaName As String
    actualAreaName = punchingName
    Dim retA As Long
    retA = SapModel.AreaObj.AddByPoint(numPoints, pointNames, actualAreaName, "None", punchingName)
    
    If retA = 0 Then
        LogMsg "SUCCESS: Created punching area " & actualAreaName & " (" & numPoints & " points)"
        On Error Resume Next
        SapModel.AreaObj.SetProperty actualAreaName, "None"
        SapModel.GroupDef.SetGroup "PunchingPerimeters"
        SapModel.AreaObj.SetGroupAssign actualAreaName, "PunchingPerimeters", False
    Else
        LogMsg "ERROR: Failed to create area " & punchingName & " (ret=" & retA & ")"
    End If
    Exit Sub
    
ErrorHandler:
    LogMsg "ERROR in DrawPerimeterAsArea: " & err.description
End Sub

Private Sub DrawPerimeterAsLines(punchingName As String, points() As Double, _
        numPoints As Long, elevation As Double)
    On Error GoTo ErrorHandler
    
    Dim i As Long
    Dim lineCount As Long: lineCount = 0
    
    ' Draw as an open polyline (do not auto-close to avoid fake edges through missing slab regions)
    For i = 0 To numPoints - 2
        Dim j As Long
        j = i + 1
        
        Dim lineName As String
        lineName = punchingName & "_L" & CStr(i + 1)
        
        Dim actualName As String
        actualName = lineName
        
        On Error Resume Next
        SapModel.frameObj.Delete lineName
        On Error GoTo ErrorHandler
        
        Dim ret As Long
        ret = SapModel.frameObj.AddByCoord(points(i, 0), points(i, 1), elevation, _
                points(j, 0), points(j, 1), elevation, _
                actualName, "None", lineName)
        If ret = 0 Then
            lineCount = lineCount + 1
            On Error Resume Next
            SapModel.GroupDef.SetGroup "PunchingPerimeters"
            SapModel.frameObj.SetGroupAssign actualName, "PunchingPerimeters", False
            On Error GoTo ErrorHandler
        End If
    Next i
    
    LogMsg "SUCCESS: Created " & lineCount & " lines for " & punchingName
    Exit Sub
    
ErrorHandler:
    LogMsg "ERROR in DrawPerimeterAsLines: " & err.description
End Sub

Private Sub GetCoverForLocation(storyIndex As Long, sectionName As String, _
        elevation As Double, X As Double, Y As Double, _
        ByRef topCover As Double, ByRef bottomCover As Double)
    On Error Resume Next
    topCover = 40: bottomCover = 40
    
    Dim coverDict As Object: Set coverDict = Nothing
    If Not M12_DrawColumnAreas.g_Settings Is Nothing Then
        If M12_DrawColumnAreas.g_Settings.exists("CoverSettings") Then
            Set coverDict = M12_DrawColumnAreas.g_Settings("CoverSettings")
        End If
    End If
    If coverDict Is Nothing Then Exit Sub
    
    Dim secKey As String: secKey = "SECTION_" & CStr(sectionName)
    If coverDict.exists(secKey) Then
        Dim secCover As Variant: secCover = coverDict(secKey)
        topCover = secCover(0): bottomCover = secCover(1)
        Exit Sub
    End If
    
    Dim storyKey As String: storyKey = "STORY_" & CStr(storyIndex)
    If coverDict.exists(storyKey) Then
        Dim storyCover As Variant: storyCover = coverDict(storyKey)
        topCover = storyCover(0): bottomCover = storyCover(1)
        Exit Sub
    End If
    
    If coverDict.exists("AUTO") Then
        Dim autoCover As Variant: autoCover = coverDict("AUTO")
        topCover = autoCover(0): bottomCover = autoCover(1)
    End If
End Sub

Private Function GetAreaSectionThickness(propName As String) As Double
    On Error Resume Next
    GetAreaSectionThickness = 0
    
    Dim ret As Long
    Dim shellType1 As Long, includeDrillingDOF As Boolean
    Dim matProp1 As String, matAngle1 As Double, thickness1 As Double
    Dim bending1 As Double, color1 As Long, notes1 As String, GUID1 As String
    
    ret = SapModel.PropArea.GetShell_1(propName, shellType1, includeDrillingDOF, matProp1, _
            matAngle1, thickness1, bending1, color1, notes1, GUID1)
    If ret = 0 And thickness1 > 0 Then
        GetAreaSectionThickness = thickness1
        Exit Function
    End If
    
    Dim shellType0 As Long, matProp0 As String, matAngle0 As Double
    Dim thickness0 As Double, bending0 As Double, color0 As Long, notes0 As String, GUID0 As String
    
    ret = SapModel.PropArea.GetShell(propName, shellType0, matProp0, matAngle0, _
            thickness0, bending0, color0, notes0, GUID0)
    If ret = 0 And thickness0 > 0 Then
        GetAreaSectionThickness = thickness0
        Exit Function
    End If
    
    Dim slabType As Long, slabDepth As Double
    Dim shearStudDiam As Double, shearStudHt As Double, shearStudFu As Double
    
    ret = SapModel.PropArea.GetSlab(propName, slabType, shellType0, matProp0, slabDepth, _
            shearStudDiam, shearStudHt, shearStudFu, color0, notes0, GUID0)
    If ret = 0 And slabDepth > 0 Then GetAreaSectionThickness = slabDepth
End Function

Private Sub LogMsg(msg As String)
    Debug.Print "M12A: " & msg
End Sub


