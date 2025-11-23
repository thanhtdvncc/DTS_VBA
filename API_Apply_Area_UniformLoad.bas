Attribute VB_Name = "API_Apply_Area_UniformLoad"
Sub AssignUniform_GroupLoad()
    Dim ws As Worksheet
    Dim i As Long
    Dim ret As Long

    ' Connect to SAP2000 and make sure SapModel is set
    If Not ConnectSAP2000 Then
        MsgBox "Could not connect to SAP2000.", vbCritical
        Exit Sub
    End If

    Set ws = ActiveSheet

    ' Set units to Ton-force, meter, Celsius
    ret = SapModel.SetPresentUnits(12)

    i = 2 ' Start from row 2 (assuming row 1 is headers)
    Do Until ws.Cells(i, 1).Value2 = ""
        ' Read parameters from worksheet
        Dim groupName As String
        Dim LoadPattern As String
        Dim coorSys As String
        Dim direction As Long
        Dim Value As Double

        groupName = ws.Cells(i, 1).Value2        ' Column A: Group Name
        LoadPattern = ws.Cells(i, 2).Value2      ' Column B: Load Pattern
        coorSys = ws.Cells(i, 3).Value2          ' Column C: Coordinate System
        direction = ws.Cells(i, 4).Value2        ' Column D: Direction
        Value = ws.Cells(i, 6).Value2            ' Column F: Value (tonf/m2)

        ' Assign uniform load to the group in one line
        ret = SapModel.AreaObj.SetLoadUniform(groupName, LoadPattern, Value, direction, True, coorSys, 1) ' Assign uniform load to group

        ' Write result into column G (column 7)
        If ret = 0 Then
            ws.Cells(i, 7).Value = "OK"
        Else
            ws.Cells(i, 7).Value = "Err"
        End If

        ' Write action time into column H (column 8)
        ws.Cells(i, 8).Value = Now

        i = i + 1
    Loop

    ' Disconnect from SAP2000
    Call DisconnectSAP2000
End Sub
Sub AssignUniform_WallHeightLoad()
    ' Assign uniform loads to wall areas across multiple input bands without overwriting contributions.
    ' Added normalization/redistribution step so total assigned force per pattern equals total desired force.
    '
    ' Input columns on ActiveSheet:
    '   A: Group Name
    '   B: Load Pattern
    '   C: CoorSys
    '   D: Direction
    '   E: Value (tonf/m2)
    '   F: Z1 (m)
    '   G: Z2 (m)
    ' Output:
    '   H: Status text (e.g., "OK(Method 1)" or "OK(PerRowMethod3)")
    '   I: Timestamp
    '   J: Assigned area list as text (area1,area2,...), cell formatted as Text

    Dim ws As Worksheet
    Dim ret As Long
    Dim lastRow As Long
    Dim r As Long

    If Not ConnectSAP2000 Then
        MsgBox "Could not connect to SAP2000.", vbCritical
        Exit Sub
    End If

    Set ws = ActiveSheet

    ' Ensure SAP model units are meters
    ret = SapModel.SetPresentUnits(12)

    ' PARAMETERS (tune if needed)
    Dim eps As Double: eps = 0.001 ' 1 mm tolerance
    Dim smallPartialThreshold As Double: smallPartialThreshold = 0.2 ' 20%
    Dim minAssignFraction As Double: minAssignFraction = 0.05 ' 5% minimum fraction to actually assign
    Dim fullTol As Double: fullTol = 0.001 ' tolerance for considering "full"
    Dim minPanelHeight As Double: minPanelHeight = 0.01 ' 1 cm
    Dim tinyE As Double: tinyE = 0.000001 ' tiny epsilon for making upper-exclusive
    Dim forceTol As Double: forceTol = 0.000000001 ' tolerance for force differences

    ' Ask user for redistribution method
    ' 1 = by area (area-weight)
    ' 2 = by current assigned force fraction
    ' 3 = equal among assigned areas
    Dim redisStr As String
    Dim redisMethod As Long
    redisStr = InputBox("Choose redistribution method to conserve total force per pattern:" & vbCrLf & _
                        "1 = by area (area-weight)" & vbCrLf & _
                        "2 = by current assigned force fraction" & vbCrLf & _
                        "3 = equal among assigned areas", "Redistribution Method", "1")
    If redisStr = "" Then
        MsgBox "Operation cancelled by user.", vbInformation
        Call DisconnectSAP2000
        Exit Sub
    End If
    On Error Resume Next
    redisMethod = CLng(redisStr)
    On Error GoTo 0
    If redisMethod < 1 Or redisMethod > 3 Then redisMethod = 1

    ' Read all input rows first
    lastRow = ws.Cells(ws.rows.count, "A").End(xlUp).row
    If lastRow < 2 Then
        MsgBox "No input rows found.", vbExclamation
        Call DisconnectSAP2000
        Exit Sub
    End If

    ' Input arrays (1-based for convenience)
    Dim inGroup() As String, inPattern() As String, inCSys() As String
    Dim inDir() As Long, inValue() As Double, inZ1() As Double, inZ2() As Double
    Dim inRowIndex() As Long
    Dim inCount As Long: inCount = 0

    ReDim inGroup(1 To lastRow)
    ReDim inPattern(1 To lastRow)
    ReDim inCSys(1 To lastRow)
    ReDim inDir(1 To lastRow)
    ReDim inValue(1 To lastRow)
    ReDim inZ1(1 To lastRow)
    ReDim inZ2(1 To lastRow)
    ReDim inRowIndex(1 To lastRow)

    For r = 2 To lastRow
        If Trim(CStr(ws.Cells(r, "A").Value2)) <> "" Then
            inCount = inCount + 1
            inGroup(inCount) = Trim(CStr(ws.Cells(r, "A").Value2))
            inPattern(inCount) = Trim(CStr(ws.Cells(r, "B").Value2))
            inCSys(inCount) = Trim(CStr(ws.Cells(r, "C").Value2))
            inDir(inCount) = CLng(ws.Cells(r, "D").Value2)
            inValue(inCount) = CDbl(ws.Cells(r, "E").Value2)
            inZ1(inCount) = CDbl(ws.Cells(r, "F").Value2)
            inZ2(inCount) = CDbl(ws.Cells(r, "G").Value2)
            inRowIndex(inCount) = r
        End If
    Next r

    If inCount = 0 Then
        MsgBox "No valid input rows.", vbExclamation
        Call DisconnectSAP2000
        Exit Sub
    End If

    ' Cache for point coordinates
    Dim pointDict As Object
    Set pointDict = CreateObject("Scripting.Dictionary")

    ' Load AreaData sheet to get area sizes (mm^2 -> convert to m^2)
    Dim areaDataDict As Object
    Set areaDataDict = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    Dim wsArea As Worksheet
    Set wsArea = ThisWorkbook.Worksheets("AreaData")
    On Error GoTo 0
    If Not wsArea Is Nothing Then
        Dim lastAreaRow As Long
        lastAreaRow = wsArea.Cells(wsArea.rows.count, "A").End(xlUp).row
        Dim ar As Long
        For ar = 2 To lastAreaRow
            Dim aName As String
            aName = Trim(CStr(wsArea.Cells(ar, "A").Value))
            If Len(aName) > 0 Then
                On Error Resume Next
                Dim area_mm2 As Double
                area_mm2 = CDbl(wsArea.Cells(ar, "H").Value)
                On Error GoTo 0
                If area_mm2 > 0 Then
                    areaDataDict(aName) = area_mm2 / 1000000# ' convert to m^2
                End If
            End If
        Next ar
    End If

    ' Group rows by GroupName
    Dim groupsDict As Object
    Set groupsDict = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = 1 To inCount
        Dim groupNameKey As String
        groupNameKey = inGroup(i)
        If Not groupsDict.exists(groupNameKey) Then
            groupsDict.Add groupNameKey, CreateObject("System.Collections.ArrayList")
        End If
        groupsDict(groupNameKey).Add i ' store index into input arrays
    Next i

    ' Prepare per-row outputs
    Dim rowStatus() As String, rowAssignedList() As String
    ReDim rowStatus(1 To inCount)
    ReDim rowAssignedList(1 To inCount)

    ' For each group process its rows and compute aggregated per-area assignment values
    Dim grpKey As Variant
    For Each grpKey In groupsDict.keys
        Dim groupRowIndices As Object
        Set groupRowIndices = groupsDict(grpKey) ' ArrayList of input indices

        ' Sort groupRowIndices by inZ1 (ascending) to detect adjacent bands
        Dim nRows As Long: nRows = groupRowIndices.count
        Dim idxArr() As Long
        ReDim idxArr(0 To nRows - 1)
        Dim k As Long
        For k = 0 To nRows - 1
            idxArr(k) = CLng(groupRowIndices(k))
        Next k
        ' Simple bubble sort
        Dim a As Long, b As Long, tmpIdx As Long
        For a = 0 To nRows - 2
            For b = a + 1 To nRows - 1
                If inZ1(idxArr(a)) > inZ1(idxArr(b)) Then
                    tmpIdx = idxArr(a): idxArr(a) = idxArr(b): idxArr(b) = tmpIdx
                End If
            Next b
        Next a

        ' Build adjusted band bounds (lower inclusive, upper exclusive except last)
        Dim bandLower() As Double, bandUpper() As Double
        ReDim bandLower(0 To nRows - 1)
        ReDim bandUpper(0 To nRows - 1)
        For k = 0 To nRows - 1
            bandLower(k) = inZ1(idxArr(k))
            bandUpper(k) = inZ2(idxArr(k))
        Next k
        ' Adjust upper bounds to be exclusive where adjacent bands meet
        For k = 0 To nRows - 2
            Dim nextLower As Double
            nextLower = bandLower(k + 1)
            If Abs(bandUpper(k) - nextLower) <= eps Then
                bandUpper(k) = nextLower - tinyE
            End If
        Next k
        ' last band's upper stays as given (inclusive)

        ' Get group assignments
        Dim NumberItems As Long
        Dim objectType() As Long
        Dim ObjectName() As String
        Dim gotAssignments As Boolean: gotAssignments = False
        On Error Resume Next
        ret = SapModel.GroupDef.GetAssignments(grpKey, NumberItems, objectType, ObjectName)
        If ret = 0 Then gotAssignments = True
        On Error GoTo 0

        If Not gotAssignments Then
            ' mark rows as error
            For k = 0 To nRows - 1
                Dim inputIndexErr As Long: inputIndexErr = idxArr(k)
                rowStatus(inputIndexErr) = "Err(GetAssignmentsFailed)"
                rowAssignedList(inputIndexErr) = ""
                ws.Cells(inRowIndex(inputIndexErr), "H").Value = rowStatus(inputIndexErr)
                ws.Cells(inRowIndex(inputIndexErr), "I").Value = Now
                ws.Cells(inRowIndex(inputIndexErr), "J").NumberFormat = "@"
                ws.Cells(inRowIndex(inputIndexErr), "J").Value = ""
            Next k
            GoTo NextGroup
        End If

        ' collect area names
        Dim areaNames() As String
        Dim numAreas As Long: numAreas = 0
        If IsArray(objectType) Then
            Dim it As Long
            For it = LBound(objectType) To UBound(objectType)
                If objectType(it) = 5 Then
                    numAreas = numAreas + 1
                    ReDim Preserve areaNames(1 To numAreas)
                    areaNames(numAreas) = Trim(CStr(ObjectName(it)))
                End If
            Next it
        End If

        If numAreas = 0 Then
            For k = 0 To nRows - 1
                Dim iiRow As Long: iiRow = idxArr(k)
                rowStatus(iiRow) = "Err(NoAreaMembers)"
                rowAssignedList(iiRow) = ""
                ws.Cells(inRowIndex(iiRow), "H").Value = rowStatus(iiRow)
                ws.Cells(inRowIndex(iiRow), "I").Value = Now
                ws.Cells(inRowIndex(iiRow), "J").NumberFormat = "@"
                ws.Cells(inRowIndex(iiRow), "J").Value = ""
            Next k
            GoTo NextGroup
        End If

        ' Precompute geometry for areas + area sizes from AreaData sheet
        Dim areaZMin() As Double, areaZMax() As Double, areaIsWall() As Boolean, areaSize() As Double
        ReDim areaZMin(1 To numAreas)
        ReDim areaZMax(1 To numAreas)
        ReDim areaIsWall(1 To numAreas)
        ReDim areaSize(1 To numAreas)

        Dim areaIndex As Long
        For areaIndex = 1 To numAreas
            Dim thisAreaName As String
            thisAreaName = areaNames(areaIndex)
            areaZMin(areaIndex) = 1E+30
            areaZMax(areaIndex) = -1E+30
            areaIsWall(areaIndex) = False
            areaSize(areaIndex) = 0#

            ' try get area size from AreaData cache
            If areaDataDict.exists(thisAreaName) Then
                areaSize(areaIndex) = CDbl(areaDataDict(thisAreaName))
            End If

            ' Get points for geometry
            Dim NumberPoints As Long
            Dim pointNames() As String
            Dim gotPoints As Boolean: gotPoints = False
            On Error Resume Next
            ret = SapModel.AreaObj.GetPoints(thisAreaName, NumberPoints, pointNames)
            If ret = 0 And NumberPoints > 0 Then gotPoints = True
            On Error GoTo 0

            If Not gotPoints Then GoTo NextAreaLoop2

            Dim ptsX() As Double, ptsY() As Double, ptsZ() As Double
            ReDim ptsX(0 To NumberPoints - 1)
            ReDim ptsY(0 To NumberPoints - 1)
            ReDim ptsZ(0 To NumberPoints - 1)
            Dim ptsCollected As Long: ptsCollected = 0

            Dim p As Long
            For p = LBound(pointNames) To UBound(pointNames)
                Dim ptName As String
                ptName = Trim(CStr(pointNames(p)))
                If Len(ptName) = 0 Then GoTo NextPointGeom3

                Dim X As Double, Y As Double, Z As Double
                If pointDict.exists(ptName) Then
                    Dim tmpCoords As Variant
                    tmpCoords = pointDict(ptName)
                    X = tmpCoords(0): Y = tmpCoords(1): Z = tmpCoords(2)
                    ret = 0
                Else
                    On Error Resume Next
                    ret = SapModel.pointObj.GetCoordCartesian(ptName, X, Y, Z)
                    On Error GoTo 0
                    If ret = 0 Then pointDict.Add ptName, Array(X, Y, Z)
                End If

                If ret = 0 Then
                    If Z < areaZMin(areaIndex) Then areaZMin(areaIndex) = Z
                    If Z > areaZMax(areaIndex) Then areaZMax(areaIndex) = Z
                    ptsX(ptsCollected) = X
                    ptsY(ptsCollected) = Y
                    ptsZ(ptsCollected) = Z
                    ptsCollected = ptsCollected + 1
                End If
NextPointGeom3:
            Next p

            ' if area size not found from AreaData, try compute polygon area projected onto best-fit plane (approx)
            If areaSize(areaIndex) <= 0# Then
                If ptsCollected >= 3 Then
                    ' compute polygon area by projecting to XY plane (approx for near-vertical walls this gives projected area)
                    ' Better: compute true polygon area in 3D by projecting to dominant plane (choose plane with largest normal component)
                    Dim projArea As Double
                    projArea = ComputePolygonAreaProjectedXY(ptsX, ptsY, ptsCollected)
                    If projArea > 0 Then
                        areaSize(areaIndex) = projArea ' in model units (m^2)
                    End If
                End If
            End If

            If ptsCollected < 3 Then GoTo NextAreaLoop2

            ' find robust normal
            Dim normalFound As Boolean: normalFound = False
            Dim A1 As Long, B1 As Long, C1 As Long
            Dim nx As Double, ny As Double, nz As Double
            For A1 = 0 To ptsCollected - 3
                For B1 = A1 + 1 To ptsCollected - 2
                    For C1 = B1 + 1 To ptsCollected - 1
                        Dim v1x As Double, v1y As Double, v1z As Double
                        Dim v2x As Double, v2y As Double, v2z As Double
                        v1x = ptsX(B1) - ptsX(A1)
                        v1y = ptsY(B1) - ptsY(A1)
                        v1z = ptsZ(B1) - ptsZ(A1)
                        v2x = ptsX(C1) - ptsX(A1)
                        v2y = ptsY(C1) - ptsY(A1)
                        v2z = ptsZ(C1) - ptsZ(A1)
                        nx = v1y * v2z - v1z * v2y
                        ny = v1z * v2x - v1x * v2z
                        nz = v1x * v2y - v1y * v2x
                        Dim nlen As Double
                        nlen = Sqr(nx * nx + ny * ny + nz * nz)
                        If nlen > 0.000000001 Then
                            nx = nx / nlen: ny = ny / nlen: nz = nz / nlen
                            normalFound = True
                            Exit For
                        End If
                    Next C1
                    If normalFound Then Exit For
                Next B1
                If normalFound Then Exit For
            Next A1

            If Not normalFound Then GoTo NextAreaLoop2

            If (Abs(nz) < 0.1) And ((Abs(nx) > 0.9) Or (Abs(ny) > 0.9)) Then
                areaIsWall(areaIndex) = True
            End If

NextAreaLoop2:
        Next areaIndex

        ' aggregationMap: key = pattern|csys|dir => areaName -> assignedValue (F/L2) (only contributions >= minAssignFraction)
        Dim aggregationMap As Object
        Set aggregationMap = CreateObject("Scripting.Dictionary")
        ' desiredForceMap: key = patternKey => areaName -> desiredForce (N or tonf? Using tonf units consistent with value*area in tonf)
        Dim desiredForceMap As Object
        Set desiredForceMap = CreateObject("Scripting.Dictionary")

        ' For each band (sorted), compute fractions for each area using adjusted band bounds
        For k = 0 To nRows - 1
            Dim inputIndex As Long
            inputIndex = idxArr(k)
            rowAssignedList(inputIndex) = ""
            rowStatus(inputIndex) = ""

            Dim anyMatched As Boolean: anyMatched = False
            Dim allFull As Boolean: allFull = True
            Dim anyPartial As Boolean: anyPartial = False
            Dim anyPartialLarge As Boolean: anyPartialLarge = False

            Dim bandL As Double, bandU As Double
            bandL = bandLower(k)
            bandU = bandUpper(k)

            Dim areaIdx2 As Long
            For areaIdx2 = 1 To numAreas
                If Not areaIsWall(areaIdx2) Then
                    ' skip non-wall areas
                Else
                    Dim zminA As Double, zmaxA As Double
                    zminA = areaZMin(areaIdx2): zmaxA = areaZMax(areaIdx2)

                    Dim overlapH As Double
                    overlapH = Application.Min(zmaxA, bandU) - Application.Max(zminA, bandL)
                    Dim fraction As Double: fraction = 0#
                    If overlapH > 0 Then
                        Dim totalH As Double: totalH = zmaxA - zminA
                        If totalH <= minPanelHeight Then
                            If (zminA >= bandL - eps And zmaxA <= bandU + eps) Then
                                fraction = 1#
                            Else
                                fraction = overlapH / minPanelHeight
                                If fraction > 1 Then fraction = 1#
                            End If
                        Else
                            fraction = overlapH / totalH
                        End If
                    Else
                        fraction = 0#
                    End If

                    ' compute desired force for this area from this row (even if tiny, include for normalization)
                    Dim areaA As Double
                    areaA = areaSize(areaIdx2)
                    If areaA <= 0# Then
                        ' if no area info, fallback: approximate by height*1 (not ideal); skip desired force if unknown
                        ' To be conservative, we will skip desired force calc if area unknown
                    Else
                        Dim desiredForce As Double
                        desiredForce = inValue(inputIndex) * fraction * areaA ' units: (force/area) * area = force (tonf)
                        ' store into desiredForceMap under pattern key
                        Dim keyStr As String
                        keyStr = inPattern(inputIndex) & "|" & inCSys(inputIndex) & "|" & CStr(inDir(inputIndex))
                        If Not desiredForceMap.exists(keyStr) Then
                            Dim dMap As Object
                            Set dMap = CreateObject("Scripting.Dictionary")
                            desiredForceMap.Add keyStr, dMap
                        End If
                        Dim dm As Object
                        Set dm = desiredForceMap(keyStr)
                        Dim prevD As Double
                        If dm.exists(areaNames(areaIdx2)) Then prevD = CDbl(dm(areaNames(areaIdx2))) Else prevD = 0#
                        dm(areaNames(areaIdx2)) = prevD + desiredForce
                    End If

                    ' now, if fraction >= minAssignFraction, include in aggregationMap (assignedValue accumulation)
                    If fraction >= minAssignFraction Then
                        anyMatched = True
                        ' append area name to rowAssignedList
                        If rowAssignedList(inputIndex) = "" Then
                            rowAssignedList(inputIndex) = areaNames(areaIdx2)
                        Else
                            rowAssignedList(inputIndex) = rowAssignedList(inputIndex) & "," & areaNames(areaIdx2)
                        End If

                        If fraction < 1 - fullTol Then
                            allFull = False
                            anyPartial = True
                            If fraction > smallPartialThreshold Then anyPartialLarge = True
                        End If

                        ' aggregate assigned value (F/L2)
                        Dim keyStr2 As String
                        keyStr2 = inPattern(inputIndex) & "|" & inCSys(inputIndex) & "|" & CStr(inDir(inputIndex))
                        If Not aggregationMap.exists(keyStr2) Then
                            Dim aDict As Object
                            Set aDict = CreateObject("Scripting.Dictionary")
                            aggregationMap.Add keyStr2, aDict
                        End If
                        Dim tDict As Object
                        Set tDict = aggregationMap(keyStr2)
                        Dim prevVal As Double
                        If tDict.exists(areaNames(areaIdx2)) Then
                            prevVal = CDbl(tDict(areaNames(areaIdx2)))
                        Else
                            prevVal = 0#
                        End If
                        tDict(areaNames(areaIdx2)) = prevVal + inValue(inputIndex) * fraction
                    End If
                End If
            Next areaIdx2

            ' decide method for this row
            If Not anyMatched Then
                rowStatus(inputIndex) = "Err(NoMatchedAreas)"
            Else
                If allFull Then
                    rowStatus(inputIndex) = "OK(Method 1)"
                ElseIf anyPartial And Not anyPartialLarge Then
                    rowStatus(inputIndex) = "OK(PerRowMethod2)"
                Else
                    rowStatus(inputIndex) = "OK(PerRowMethod3)"
                End If
            End If

            ' provisional write for the row (final assignment happens after aggregation + normalization)
            ws.Cells(inRowIndex(inputIndex), "H").Value = rowStatus(inputIndex)
            ws.Cells(inRowIndex(inputIndex), "I").Value = Now
            ws.Cells(inRowIndex(inputIndex), "J").NumberFormat = "@"
            ws.Cells(inRowIndex(inputIndex), "J").Value = rowAssignedList(inputIndex)
        Next k

        ' Now perform assignments using aggregationMap (accumulated per pattern)
        Dim patternKey As Variant
        For Each patternKey In aggregationMap.keys
            ' compute total desired force and total assigned force for this pattern
            Dim parts() As String
            parts = Split(CStr(patternKey), "|")
            Dim pPattern As String: pPattern = parts(0)
            Dim pCSys As String: pCSys = parts(1)
            Dim pDir As Long: pDir = CLng(parts(2))

            Dim desiredForcesExist As Boolean: desiredForcesExist = False
            Dim totalDesiredForce As Double: totalDesiredForce = 0#
            If desiredForceMap.exists(patternKey) Then
                desiredForcesExist = True
                Dim dmapLocal As Object
                Set dmapLocal = desiredForceMap(patternKey)
                Dim anKey As Variant
                For Each anKey In dmapLocal.keys
                    totalDesiredForce = totalDesiredForce + CDbl(dmapLocal(anKey))
                Next anKey
            End If

            ' compute totalAssignedForce from aggregationMap (value * area)
            Dim areaDict2 As Object
            Set areaDict2 = aggregationMap(patternKey)
            Dim totalAssignedForce As Double: totalAssignedForce = 0#
            Dim areaNameKey As Variant
            For Each areaNameKey In areaDict2.keys
                Dim valF As Double
                valF = CDbl(areaDict2(areaNameKey)) ' F/L2
                Dim aArea As Double
                If areaDataDict.exists(CStr(areaNameKey)) Then
                    aArea = CDbl(areaDataDict(CStr(areaNameKey)))
                Else
                    aArea = 0#
                End If
                totalAssignedForce = totalAssignedForce + valF * aArea
            Next areaNameKey

            ' If desired forces exist compute diff and redistribute if needed
            Dim diffForce As Double
            If desiredForcesExist Then
                diffForce = totalDesiredForce - totalAssignedForce
            Else
                diffForce = 0#
            End If

            ' If diff is significant, redistribute according to redisMethod
            If Abs(diffForce) > forceTol Then
                ' Build distribution weights among assigned areas (areaDict2.Keys)
                Dim weights As Object
                Set weights = CreateObject("Scripting.Dictionary")
                Dim sumWeights As Double: sumWeights = 0#
                ' compute base weights
                For Each areaNameKey In areaDict2.keys
                    Dim w As Double: w = 0#
                    If redisMethod = 1 Then
                        ' by area size
                        If areaDataDict.exists(CStr(areaNameKey)) Then
                            w = CDbl(areaDataDict(CStr(areaNameKey)))
                        Else
                            w = 0#
                        End If
                    ElseIf redisMethod = 2 Then
                        ' by current assigned force (valF * area)
                        Dim curVal As Double
                        curVal = CDbl(areaDict2(areaNameKey))
                        Dim aArea2 As Double
                        If areaDataDict.exists(CStr(areaNameKey)) Then
                            aArea2 = CDbl(areaDataDict(CStr(areaNameKey)))
                        Else
                            aArea2 = 0#
                        End If
                        w = curVal * aArea2
                    Else
                        ' equal
                        w = 1#
                    End If
                    If w < 0 Then w = 0#
                    weights(areaNameKey) = w
                    sumWeights = sumWeights + w
                Next areaNameKey

                If sumWeights <= 0 Then
                    ' fallback: distribute equally
                    For Each areaNameKey In areaDict2.keys
                        weights(areaNameKey) = 1#
                    Next areaNameKey
                    sumWeights = aggregationMap(patternKey).count
                End If

                ' Now compute addForce per area and update areaDict2 (value in F/L2)
                For Each areaNameKey In areaDict2.keys
                    Dim wgt As Double
                    wgt = CDbl(weights(areaNameKey)) / sumWeights
                    Dim addForce As Double
                    addForce = diffForce * wgt ' force units (tonf)
                    Dim area_m2 As Double
                    If areaDataDict.exists(CStr(areaNameKey)) Then
                        area_m2 = CDbl(areaDataDict(CStr(areaNameKey)))
                    Else
                        area_m2 = 0#
                    End If
                    If area_m2 > 0 Then
                        Dim addValue As Double
                        addValue = addForce / area_m2 ' convert force back to F/L2
                        areaDict2(areaNameKey) = CDbl(areaDict2(areaNameKey)) + addValue
                    Else
                        ' if area unknown, try equal distribution in value-space
                        areaDict2(areaNameKey) = CDbl(areaDict2(areaNameKey))
                    End If
                Next areaNameKey
            End If

            ' delete existing loads for group on this pattern to clear previous assignments
            On Error Resume Next
            ret = SapModel.AreaObj.DeleteLoadUniform(grpKey, pPattern, 1)
            On Error GoTo 0

            ' perform final assignment per area using updated areaDict2
            Dim assignedCount As Long: assignedCount = 0
            For Each areaNameKey In areaDict2.keys
                Dim finalVal As Double
                finalVal = CDbl(areaDict2(areaNameKey))
                If Abs(finalVal) < 0.000000000001 Then GoTo NextAssignFinal
                On Error Resume Next
                ret = SapModel.AreaObj.SetLoadUniform(CStr(areaNameKey), pPattern, finalVal, pDir, True, pCSys, 0)
                On Error GoTo 0
                If ret = 0 Then assignedCount = assignedCount + 1
NextAssignFinal:
            Next areaNameKey
            ' no further per-row logging here; rows already updated earlier
        Next patternKey

NextGroup:
    Next grpKey

    ' Disconnect
    Call DisconnectSAP2000
End Sub

' Helper: approximate polygon area projected to XY (signed). pts arrays are 0-based and count points.
Private Function ComputePolygonAreaProjectedXY(ByRef px() As Double, ByRef py() As Double, ByVal countPts As Long) As Double
    Dim i As Long
    Dim area As Double: area = 0#
    If countPts < 3 Then
        ComputePolygonAreaProjectedXY = 0#
        Exit Function
    End If
    For i = 0 To countPts - 1
        Dim j As Long
        j = IIf(i = countPts - 1, 0, i + 1)
        area = area + (px(i) * py(j) - px(j) * py(i))
    Next i
    ComputePolygonAreaProjectedXY = Abs(area) / 2#
End Function
Sub AssignUniform_To_Frame()
    Dim ws As Worksheet
    Dim i As Long
    Dim ret As Long

    ' Connect to SAP2000 and make sure SapModel is set
    If Not ConnectSAP2000 Then
        MsgBox "Could not connect to SAP2000.", vbCritical
        Exit Sub
    End If

    Set ws = ActiveSheet

    ' Set units to Ton-force, meter, Celsius
    ret = SapModel.SetPresentUnits(12)

    i = 2 ' Start from row 2 (assuming row 1 is headers)
    Do Until ws.Cells(i, 1).Value2 = ""
        ' Read parameters from worksheet
        Dim groupName As String
        Dim LoadPattern As String
        Dim coorSys As String
        Dim direction As Long
        Dim distType As Long
        Dim Value As Double

        groupName = ws.Cells(i, 1).Value2        ' Column A: Group Name
        LoadPattern = ws.Cells(i, 2).Value2      ' Column B: Load Pattern
        coorSys = ws.Cells(i, 3).Value2          ' Column C: Coordinate System
        direction = ws.Cells(i, 4).Value2        ' Column D: Direction
        distType = ws.Cells(i, 5).Value2         ' Column E: Distribution Type (1=One-way, 2=Two-way)
        Value = ws.Cells(i, 6).Value2            ' Column F: Value (tonf/m2)

        ' Assign uniform to frame load to the group
        ' SetLoadUniformToFrame(Name, LoadPat, Value, Dir, DistType, Replace, CSys, ItemType)
        ret = SapModel.AreaObj.SetLoadUniformToFrame(groupName, LoadPattern, Value, direction, distType, True, coorSys, 1)

        ' Write result into column G (column 7)
        If ret = 0 Then
            ws.Cells(i, 7).Value = "OK"
        Else
            ws.Cells(i, 7).Value = "Err"
        End If

        ' Write action time into column H (column 8)
        ws.Cells(i, 8).Value = Now

        i = i + 1
    Loop

    ' Disconnect from SAP2000
    Call DisconnectSAP2000
End Sub

Sub AssignSurfacePressure()
    
    SwitchOff True
    'Call SAP2000_Connectv16
    Call SAP2000_Connect
    Set ws = Sheets("AssignSurfacePressure")
    i = 2
    ret = SapModel.SetPresentUnits(12) 'Ton_m_C
    Do Until ws.Range("A" & i).Value2 = ""
        Call AssignJointPatternAndPressure(ws.Range("A" & i).Value2, ws.Range("B" & i).Value2, ws.Range("C" & i).Value2, ws.Range("D" & i).Value2, ws.Range("E" & i).Value2, ws.Range("F" & i).Value2, ws.Range("G" & i).Value2, ws.Range("H" & i).Value2, ws.Range("I" & i).Value2, ws.Range("J" & i).Value2)
        i = i + 1
    Loop

    Call SAP2000_Disconnect
    SwitchOff False

End Sub

Sub AssignJointPatternAndPressure(groupName, LoadPatternName, JointPatternName, z1, p1, z2, p2, Face, Value, Restriction)
    Dim NumberPoints As Long
    Dim Point() As String
    Dim NumberItems As Long
    Dim objectType() As Long
    Dim ObjectName() As String
    Dim X As Double, Y As Double, Z As Double
    '''
    c = (p1 - p2) / (z1 - z2)
    d = (p2 * z1 - p1 * z2) / (z1 - z2)
    If Face = "Top" Then: FaceIndex = -2
    If Face = "Bottom" Then: FaceIndex = -1
    If Restriction = "All" Then: RestrictionIndex = 0
    If Restriction = "Negative" Then: RestrictionIndex = 1
    If Restriction = "Positive" Then: RestrictionIndex = 2
    Z1Z2 = Abs(z1 - z2)
    '''
    ret = SapModel.SelectObj.ClearSelection
    Set areaName = New Collection
    ret = SapModel.GroupDef.GetAssignments(groupName, NumberItems, objectType, ObjectName)
    For i = 0 To NumberItems - 1
        If objectType(i) = 5 Then
            Condition = True
            ret = SapModel.AreaObj.GetPoints(ObjectName(i), NumberPoints, Point)
            For Each item In Point
                ret = SapModel.pointObj.GetCoordCartesian(item, X, Y, Z)
                If Abs(Z1Z2 - (Abs(Z - z1) + Abs(Z - z2))) <= 0.01 Then
                    ret = SapModel.pointObj.SetSelected(item, True)
                Else: Condition = False
                End If
            Next
            If Condition = True Then: areaName.Add ObjectName(i)
        End If
    Next
    ret = SapModel.pointObj.SetPatternByXYZ("", JointPatternName, 0, 0, c, d, 2, RestrictionIndex, True)
    If ret = 1 Then MsgBox ("Warning")
    ret = SapModel.SelectObj.ClearSelection
    For Each item In areaName
        ret = SapModel.AreaObj.SetSelected(item, True, 0)
    Next
    ret = SapModel.AreaObj.SetLoadSurfacePressure("", LoadPatternName, FaceIndex, Value, JointPatternName, True, 2)
    If ret = 1 Then MsgBox ("Warning")
    ret = SapModel.SelectObj.ClearSelection
End Sub

Option Explicit

Public Sub DeleteUniformToShell()
    Dim numSelected As Long
    Dim selTypes() As Long            ' Will be filled by GetSelected
    Dim selNames() As String         ' Will be filled by GetSelected
    Dim resp As VbMsgBoxResult
    Dim numPatterns As Long
    Dim patternNames() As String     ' Will be filled by GetNameList
    Dim i As Long
    Dim deletedCount As Long
    Dim errorCount As Long
    Dim failedList As String
    Dim ret As Long
    Dim hasAreaSelected As Boolean
    Dim idx As Long
    Dim msg As String

    ' Connect to SAP2000 (assumes ConnectSAP2000 sets global SapModel)
    ConnectSAP2000

    ' Optional: set units (example: Ton-force, meter, Celsius)
    ret = SapModel.SetPresentUnits(12)

    ' Get currently selected objects
    ' Correct parameter order: (NumberItems, ObjectType(), ObjectName())
    ret = SapModel.SelectObj.GetSelected(numSelected, selTypes, selNames)
    If ret <> 0 Then
        MsgBox "Error retrieving selected objects. Error code: " & ret, vbCritical, "Error"
        GoTo CleanExit
    End If

    If numSelected = 0 Then
        MsgBox "No objects are selected. Please select area objects and run again.", vbInformation, "No Selection"
        GoTo CleanExit
    End If

    ' Check if at least one selected object is an area object (ObjectType = 5)
    hasAreaSelected = False
    For idx = 0 To numSelected - 1
        If selTypes(idx) = 5 Then
            hasAreaSelected = True
            Exit For
        End If
    Next idx

    If Not hasAreaSelected Then
        MsgBox "No area objects are selected. Please select area objects and run again.", vbInformation, "No Area Selection"
        GoTo CleanExit
    End If

    ' Ask user for confirmation
    resp = MsgBox("Delete ALL uniform load on selection shell?", vbYesNo + vbQuestion, "Confirm Delete Uniform Loads")

    If resp = vbNo Then
        MsgBox "Operation cancelled by user.", vbInformation, "Cancelled"
        GoTo CleanExit
    End If

    ' Get list of load patterns
    ret = SapModel.loadPatterns.GetNameList(numPatterns, patternNames)
    If ret <> 0 Then
        MsgBox "Error retrieving load patterns. Error code: " & ret, vbCritical, "Error"
        GoTo CleanExit
    End If

    deletedCount = 0
    errorCount = 0
    failedList = ""

    ' Loop through all load patterns and delete uniform loads on selected objects
    For i = 0 To numPatterns - 1
        ' ItemType = 2 (SelectedObjects). Name is ignored.
        ret = SapModel.AreaObj.DeleteLoadUniform("", patternNames(i), 2)
        If ret = 0 Then
            deletedCount = deletedCount + 1
        Else
            errorCount = errorCount + 1
            failedList = failedList & patternNames(i) & " (code " & ret & ")" & vbCrLf
        End If
    Next i

    ' Report results
    msg = "Done." & vbCrLf & _
          "Patterns processed: " & numPatterns & vbCrLf & _
          "Deleted (success count): " & deletedCount & vbCrLf & _
          "Errors: " & errorCount
    If errorCount > 0 Then
        msg = msg & vbCrLf & vbCrLf & "Failed patterns:" & vbCrLf & failedList
    End If

    MsgBox msg, vbInformation, "DeleteUniformToShell - Result"

CleanExit:
    On Error Resume Next
    ' Disconnect from SAP2000 (assumes DisconnectSAP2000 exists)
    DisconnectSAP2000
End Sub

