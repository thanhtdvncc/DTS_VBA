Attribute VB_Name = "m05_SAP2000_Write_Sections"
Option Explicit
'===============================================================
' Module: modSAP2000_Sections
' Purpose: Write frame & area section property data
'===============================================================

Public Sub WriteFrameSections()
    If Not ENABLE_FRAME_SECTIONS Then Exit Sub
    If gFrameCount = 0 Or IsArrayEmpty(gFrameProp) Then Exit Sub
    
    Dim uniqProps As Variant
    uniqProps = RemoveDuplicatesInArray(gFrameProp)
    If IsArrayEmpty(uniqProps) Then Exit Sub
    
    Dim ws As Worksheet
    Set ws = SheetOrCreate("Section", True)
    If ws Is Nothing Then
        LogMsg "WriteFrameSections: Cannot access Section sheet."
        Exit Sub
    End If
    
    ws.Range("A3:K" & ws.rows.count).clearContents
    
    Dim i As Long, propName As String, ret As Long
    Dim FileName As String, matProp As String
    Dim t3 As Double, t2 As Double, tf As Double, tw As Double
    Dim t2b As Double, tfb As Double
    Dim color As Long, notes As String, guid As String
    Dim rowPtr As Long: rowPtr = 3
    
    For i = LBound(uniqProps) To UBound(uniqProps)
        propName = CStr(uniqProps(i))
        If Len(propName) > 0 Then
            t3 = 0#: t2 = 0#: tf = 0#: tw = 0#: t2b = 0#: tfb = 0#
            matProp = "": FileName = "": notes = "": guid = ""
            
            ret = SapModel.PropFrame.GetAngle(propName, FileName, matProp, t3, t2, tf, tw, color, notes, guid)
            ret = SapModel.PropFrame.GetChannel(propName, FileName, matProp, t3, t2, tf, tw, color, notes, guid)
            ret = SapModel.PropFrame.GetISection(propName, FileName, matProp, t3, t2, tf, tw, t2b, tfb, color, notes, guid)
            ret = SapModel.PropFrame.GetRectangle(propName, FileName, matProp, t3, t2, color, notes, guid)
            ret = SapModel.PropFrame.GetTee(propName, FileName, matProp, t3, t2, tf, tw, color, notes, guid)
            ret = SapModel.PropFrame.GetTube(propName, FileName, matProp, t3, t2, tf, tw, color, notes, guid)
            ret = SapModel.PropFrame.GetPipe(propName, FileName, matProp, t3, tw, color, notes, guid)
            ret = SapModel.PropFrame.GetCircle(propName, FileName, matProp, t3, color, notes, guid)
            
            ws.Cells(rowPtr, "A").Value = propName
            ws.Cells(rowPtr, "B").Value = matProp
            ws.Cells(rowPtr, "C").Value = t3
            ws.Cells(rowPtr, "D").Value = t2
            ws.Cells(rowPtr, "E").Value = tf
            ws.Cells(rowPtr, "F").Value = tw
            ws.Cells(rowPtr, "G").Value = t2b
            ws.Cells(rowPtr, "H").Value = tfb
            ws.Cells(rowPtr, "I").Value = color
            ws.Cells(rowPtr, "J").Value = notes
            ws.Cells(rowPtr, "K").Value = guid
            rowPtr = rowPtr + 1
        End If
    Next
End Sub

Public Sub WriteAreaSections()
    If Not ENABLE_AREA_SECTIONS Then Exit Sub
    If gAreaCount = 0 Or IsArrayEmpty(gAreaProp) Then Exit Sub
    
    Dim uniqAreaProps As Variant
    uniqAreaProps = RemoveDuplicatesInArray(gAreaProp)
    If IsArrayEmpty(uniqAreaProps) Then Exit Sub
    
    Dim ws As Worksheet
    Set ws = SheetOrCreate("AreaSection", True)
    If ws Is Nothing Then
        LogMsg "WriteAreaSections: Cannot access AreaSection sheet."
        Exit Sub
    End If
    
    ws.Cells.clearContents
    ws.Range("A1:H1").Value = Array("Property", "Material", "Thickness1", "Thickness2", "Notes", "GUID", "Color", "TypeDetected")
    
    Dim i As Long, propName As String, ret As Long
    Dim mat As String, thk As Double, thk1 As Double, thk2 As Double
    Dim color As Long, notes As String, guid As String
    Dim typeDetected As String
    Dim shellType As Long, MatAng As Double
    Dim rowPtr As Long: rowPtr = 2
    
    For i = LBound(uniqAreaProps) To UBound(uniqAreaProps)
        propName = CStr(uniqAreaProps(i))
        If Len(propName) > 0 Then
            mat = "": thk = 0#: thk1 = 0#: thk2 = 0#: color = 0: notes = "": guid = ""
            shellType = 0: MatAng = 0#
            typeDetected = ""
            
            On Error Resume Next
            ret = SapModel.PropArea.GetShell(propName, shellType, mat, MatAng, thk1, thk2, color, notes, guid)
            If err.number = 0 And ret = 0 Then typeDetected = "Shell"
            err.Clear
            
            On Error Resume Next
            If typeDetected = "" Then
                ret = SapModel.PropArea.GetPlate(propName, shellType, mat, MatAng, thk1, thk2, color, notes, guid)
                If err.number = 0 And ret = 0 Then typeDetected = "Plate"
                err.Clear
            End If
            
            If typeDetected = "" Then
                ret = SapModel.PropArea.GetMembrane(propName, mat, thk, color, notes, guid)
                If err.number = 0 And ret = 0 Then
                    typeDetected = "Membrane"
                    thk1 = thk: thk2 = 0#
                End If
                err.Clear
            End If
            
            If typeDetected = "" Then
                Dim nLayer As Long, matNames() As String, thickArr() As Double
                ret = SapModel.PropArea.GetLayered(propName, nLayer, matNames, thickArr, color, notes, guid)
                If err.number = 0 And ret = 0 Then
                    typeDetected = "Layered(" & nLayer & ")"
                    Dim sumT As Double, li As Long
                    For li = 0 To nLayer - 1
                        sumT = sumT + thickArr(li)
                    Next
                    thk1 = sumT
                    thk2 = 0#
                End If
                err.Clear
            End If
            
            If typeDetected = "" Then
                ret = SapModel.PropArea.GetPlaneStress(propName, mat, thk, color, notes, guid)
                If err.number = 0 And ret = 0 Then
                    typeDetected = "PlaneStress"
                    thk1 = thk
                End If
                err.Clear
            End If
            
            If typeDetected = "" Then
                ret = SapModel.PropArea.GetPlaneStrain(propName, mat, thk, color, notes, guid)
                If err.number = 0 And ret = 0 Then
                    typeDetected = "PlaneStrain"
                    thk1 = thk
                End If
                err.Clear
            End If
            
            If typeDetected = "" Then typeDetected = "Unknown"
            
            ws.Cells(rowPtr, "A").Value = propName
            ws.Cells(rowPtr, "B").Value = mat
            ws.Cells(rowPtr, "C").Value = thk1
            ws.Cells(rowPtr, "D").Value = thk2
            ws.Cells(rowPtr, "E").Value = notes
            ws.Cells(rowPtr, "F").Value = guid
            ws.Cells(rowPtr, "G").Value = color
            ws.Cells(rowPtr, "H").Value = typeDetected
            rowPtr = rowPtr + 1
        End If
    Next
    
    ' Generate area summary report
    Call WriteAreaSummaryReport
End Sub

Private Sub WriteAreaSummaryReport()
    ' Get the sheet with area data
    Dim wsData As Worksheet
    Set wsData = Nothing
    
    On Error Resume Next
    Set wsData = ThisWorkbook.Worksheets("AreaData")
    On Error GoTo 0
    
    If wsData Is Nothing Then
        LogMsg "WriteAreaSummaryReport: Cannot find AreaData sheet."
        Exit Sub
    End If
    
    ' Find last row with data
    Dim lastRow As Long
    lastRow = wsData.Cells(wsData.rows.count, "A").End(xlUp).row
    
    If lastRow < 2 Then Exit Sub
    
    ' Dictionary to store summary data
    Dim summaryDict As Object
    Set summaryDict = CreateObject("Scripting.Dictionary")
    
    ' Read and classify data
    Dim i As Long
    Dim areaName As String, propName As String
    Dim cx As Double, cy As Double, cz As Double
    Dim areaVal As Double
    Dim nx As Double, ny As Double, nz As Double
    Dim areaType As String, coordKey As String, keyStr As String
    
    For i = 2 To lastRow
        On Error Resume Next
        areaName = CStr(wsData.Cells(i, "A").Value)
        propName = CStr(wsData.Cells(i, "B").Value)
        
        Dim pointList As String
        pointList = CStr(wsData.Cells(i, "D").Value) ' Column D = PointList
        
        ' Read coordinates and area (convert to meters)
        cx = CDbl(wsData.Cells(i, "E").Value) / 1000 ' Convert mm to m
        cy = CDbl(wsData.Cells(i, "F").Value) / 1000 ' Convert mm to m
        cz = CDbl(wsData.Cells(i, "G").Value) / 1000 ' Convert mm to m
        areaVal = CDbl(wsData.Cells(i, "H").Value) / 1000000 ' Convert mm2 to m2
        
        nx = CDbl(wsData.Cells(i, "I").Value)
        ny = CDbl(wsData.Cells(i, "J").Value)
        nz = CDbl(wsData.Cells(i, "K").Value)
        On Error GoTo 0
        
        If Len(areaName) > 0 And Len(propName) > 0 And Len(pointList) > 0 Then
            ' Skip invalid area names (like "000")
            If areaName <> "000" And areaName <> "0" Then
                ' Classify area type based on normal vector
                areaType = ClassifyAreaType(nx, ny, nz)
                
                ' Get coordinate key from actual node coordinates
                coordKey = GetCoordKeyFromNodes(areaType, pointList, nx, ny)
            
            ' Create unique key: Type|CoordKey|Property
            keyStr = areaType & "|" & coordKey & "|" & propName
            
            ' Add to dictionary
            If summaryDict.exists(keyStr) Then
                Dim existData As Variant
                existData = summaryDict(keyStr)
                existData(0) = existData(0) + areaVal ' Sum area
                existData(1) = CStr(existData(1)) & "," & areaName ' Append area name as string
                summaryDict(keyStr) = existData
            Else
                Dim newData(0 To 5) As Variant
                newData(0) = areaVal ' Total area
                newData(1) = CStr(areaName) ' Area list as string
                newData(2) = areaType ' Type
                newData(3) = propName ' Property
                newData(4) = coordKey ' Coordinate key
                newData(5) = 0 ' Dummy value for sorting
                summaryDict.Add keyStr, newData
            End If
            End If
        End If
    Next i
    
    ' Write summary to columns N onwards
    If summaryDict.count > 0 Then
        Call WriteSummaryToSheet(wsData, summaryDict)
        LogMsg "Area summary report generated: " & summaryDict.count & " groups."
    End If
End Sub

Private Function ClassifyAreaType(nx As Double, ny As Double, nz As Double) As String
    ' Classify based on normal vector from SAP2000
    ' Normal vector shows the direction perpendicular to the area surface
    
    Dim absNx As Double, absNy As Double, absNz As Double
    absNx = Abs(nx)
    absNy = Abs(ny)
    absNz = Abs(nz)
    
    ' Slab (horizontal): Normal parallel to Z axis
    ' NormalZ ˜ ±1, NormalX ˜ 0, NormalY ˜ 0
    If absNz > 0.9 And absNx < 0.1 And absNy < 0.1 Then
        ClassifyAreaType = "Slab"
        
    ' Wall (vertical): Normal parallel to X or Y axis
    ' NormalX ˜ ±1 OR NormalY ˜ ±1, NormalZ ˜ 0
    ElseIf absNz < 0.1 And (absNx > 0.9 Or absNy > 0.9) Then
        ClassifyAreaType = "Wall"
        
    ' Other (inclined): All other cases
    Else
        ClassifyAreaType = "Inclined"
    End If
End Function

Private Function GetCoordKeyFromNodes(areaType As String, pointList As String, nx As Double, ny As Double) As String
    ' Get coordinate key by analyzing actual node coordinates from SAP2000
    ' pointList format: "14,140,139,13" (comma-separated point names)
    
    Dim points() As String
    points = Split(pointList, ",")
    
    If UBound(points) < 0 Then
        GetCoordKeyFromNodes = "Unknown"
        Exit Function
    End If
    
    ' Arrays to store coordinates
    Dim xCoords() As Double, yCoords() As Double, zCoords() As Double
    ReDim xCoords(UBound(points))
    ReDim yCoords(UBound(points))
    ReDim zCoords(UBound(points))
    
    ' Get coordinates for each point
    Dim j As Long, ret As Long
    Dim X As Double, Y As Double, Z As Double
    Dim validCount As Long
    validCount = 0
    
    For j = 0 To UBound(points)
        Dim ptName As String
        ptName = Trim(points(j))
        
        ' Skip invalid point names
        If ptName = "" Or ptName = "0" Or ptName = "000" Then
            GoTo NextPoint
        End If
        
        On Error Resume Next
        ret = SapModel.pointObj.GetCoordCartesian(ptName, X, Y, Z)
        If err.number = 0 And ret = 0 Then
            xCoords(validCount) = X
            yCoords(validCount) = Y
            zCoords(validCount) = Z
            validCount = validCount + 1
        End If
        err.Clear
        On Error GoTo 0
        
NextPoint:
    Next j
    
    ' Check if we have valid coordinates
    If validCount = 0 Then
        GetCoordKeyFromNodes = "NoValidPoints"
        Exit Function
    End If
    
    ' Determine coordinate key based on area type
    ' Note: SAP2000 coordinates are already in mm, keep them in mm
    If areaType = "Slab" Then
        ' For slab: use average Z coordinate (in mm)
        Dim avgZ As Double, sumZ As Double
        Dim i As Long
        For i = 0 To validCount - 1
            sumZ = sumZ + zCoords(i)
        Next i
        avgZ = sumZ / validCount
        
        ' Round to nearest integer (mm)
        Dim zRounded As Long
        zRounded = Round(avgZ)
        GetCoordKeyFromNodes = "Z=" & CStr(zRounded)
        
    ElseIf areaType = "Wall" Then
        ' For wall: use X or Y (whichever is more constant) based on normal
        If Abs(nx) > Abs(ny) Then
            ' Wall parallel to YZ plane (normal in X direction)
            Dim avgX As Double, sumx As Double
            For i = 0 To validCount - 1
                sumx = sumx + xCoords(i)
            Next i
            avgX = sumx / validCount
            
            Dim xRounded As Long
            xRounded = Round(avgX)
            GetCoordKeyFromNodes = "X=" & CStr(xRounded)
        Else
            ' Wall parallel to XZ plane (normal in Y direction)
            Dim avgY As Double, sumy As Double
            For i = 0 To validCount - 1
                sumy = sumy + yCoords(i)
            Next i
            avgY = sumy / validCount
            
            Dim yRounded As Long
            yRounded = Round(avgY)
            GetCoordKeyFromNodes = "Y=" & CStr(yRounded)
        End If
        
    Else
        ' For inclined: use all coordinates (average, in mm)
        Dim ax As Double, ay As Double, az As Double
        Dim sx As Double, sy As Double, sz As Double
        
        For i = 0 To validCount - 1
            sx = sx + xCoords(i)
            sy = sy + yCoords(i)
            sz = sz + zCoords(i)
        Next i
        
        ax = sx / validCount
        ay = sy / validCount
        az = sz / validCount
        
        Dim rx As Long, ry As Long, rz As Long
        rx = Round(ax)
        ry = Round(ay)
        rz = Round(az)
        
        GetCoordKeyFromNodes = "X=" & CStr(rx) & ";Y=" & CStr(ry) & ";Z=" & CStr(rz)
    End If
End Function

Private Function CreateCoordinateKey(areaType As String, cx As Double, cy As Double, cz As Double, nx As Double, ny As Double) As String
    ' This function is now replaced by GetCoordKeyFromNodes
    ' Keeping for backward compatibility only
    Dim roundedZ As Long
    roundedZ = Round(cz / 1000)
    CreateCoordinateKey = "Z=" & roundedZ
End Function

Private Sub WriteSummaryToSheet(ws As Worksheet, summaryDict As Object)
    ' Clear existing summary columns completely
    ws.Range("N:S").Clear
    
    ' Write headers
    ws.Range("N1").Value = "Type"
    ws.Range("O1").Value = "Coordinate (mm)"
    ws.Range("P1").Value = "Property"
    ws.Range("Q1").Value = "Area List"
    ws.Range("R1").Value = "Total Area (m2)"
    ws.Range("S1").Value = "Floor Area (m2)"
    
    ' Sort data by Type, Coordinate (numerically for Z values), Property
    Dim sortedKeys() As Variant
    sortedKeys = SortDictionaryKeysByCoordinate(summaryDict)
    
    ' Calculate floor totals (sum of all Slab areas at each Z coordinate)
    Dim floorTotals As Object
    Set floorTotals = CreateObject("Scripting.Dictionary")
    
    Dim k As Variant
    For Each k In sortedKeys
        Dim data As Variant
        data = summaryDict(CStr(k))
        
        Dim areaType As String, coordKey As String, areaValue As Double
        areaType = CStr(data(2))
        coordKey = CStr(data(4))
        areaValue = CDbl(data(0))
        
        ' Only sum Slab areas for floor totals
        If areaType = "Slab" Then
            If floorTotals.exists(coordKey) Then
                floorTotals(coordKey) = CDbl(floorTotals(coordKey)) + areaValue
            Else
                floorTotals.Add coordKey, areaValue
            End If
        End If
    Next k
    
    ' Write data - hide repeated Type and Coordinate values
    Dim rowPtr As Long
    rowPtr = 2
    
    Dim prevType As String, prevCoord As String
    prevType = ""
    prevCoord = ""
    
    For Each k In sortedKeys
        data = summaryDict(CStr(k))
        
        Dim currType As String, currCoord As String
        currType = CStr(data(2))
        currCoord = CStr(data(4))
        
        ' Write Type only if different from previous row
        If currType <> prevType Then
            ws.Cells(rowPtr, "N").Value = currType
            prevType = currType
            prevCoord = "" ' Reset coordinate when type changes
        Else
            ws.Cells(rowPtr, "N").Value = "" ' Leave blank
        End If
        
        ' Write Coordinate only if different from previous row (within same type)
        If currCoord <> prevCoord Then
            ws.Cells(rowPtr, "O").Value = currCoord
            prevCoord = currCoord
        Else
            ws.Cells(rowPtr, "O").Value = "" ' Leave blank
        End If
        
        ' Always write Property, Area List, and Total Area
        ws.Cells(rowPtr, "P").Value = CStr(data(3)) ' Property
        
        ' Force Area List as text to prevent Excel from converting to number
        ws.Cells(rowPtr, "Q").NumberFormat = "@" ' Text format
        ws.Cells(rowPtr, "Q").Value = CStr(data(1)) ' Area list as string
        
        ws.Cells(rowPtr, "R").Value = data(0) ' Already in m2
        
        ' Write Floor Area total only for Slab type and only on first row of each coordinate
        If currType = "Slab" And ws.Cells(rowPtr, "O").Value <> "" Then
            ' This is the first row for this coordinate
            If floorTotals.exists(currCoord) Then
                ws.Cells(rowPtr, "S").Value = CDbl(floorTotals(currCoord))
            End If
        End If
        
        rowPtr = rowPtr + 1
    Next k
    
    ' Simple number format for area columns only
    ws.Columns("R:S").NumberFormat = "0.00"
End Sub

Private Function SortDictionaryKeys(dict As Object) As Variant
    ' Simple bubble sort for dictionary keys
    Dim keys() As Variant
    ReDim keys(0 To dict.count - 1)
    
    Dim i As Long
    i = 0
    Dim k As Variant
    For Each k In dict.keys
        keys(i) = k
        i = i + 1
    Next k
    
    ' Bubble sort
    Dim j As Long, temp As Variant
    For i = 0 To UBound(keys) - 1
        For j = i + 1 To UBound(keys)
            If CStr(keys(i)) > CStr(keys(j)) Then
                temp = keys(i)
                keys(i) = keys(j)
                keys(j) = temp
            End If
        Next j
    Next i
    
    SortDictionaryKeys = keys
End Function

Private Function SortDictionaryKeysByCoordinate(dict As Object) As Variant
    ' Sort dictionary keys by Type, then by coordinate value (numerically for Z), then by Property
    Dim keys() As Variant
    ReDim keys(0 To dict.count - 1)
    
    Dim i As Long
    i = 0
    Dim k As Variant
    For Each k In dict.keys
        keys(i) = k
        i = i + 1
    Next k
    
    ' Bubble sort with custom comparison
    Dim j As Long, temp As Variant
    For i = 0 To UBound(keys) - 1
        For j = i + 1 To UBound(keys)
            If CompareKeys(CStr(keys(i)), CStr(keys(j))) > 0 Then
                temp = keys(i)
                keys(i) = keys(j)
                keys(j) = temp
            End If
        Next j
    Next i
    
    SortDictionaryKeysByCoordinate = keys
End Function

Private Function CompareKeys(key1 As String, key2 As String) As Integer
    ' Compare two keys: Type|CoordKey|Property
    ' Return: -1 if key1 < key2, 0 if equal, 1 if key1 > key2
    
    Dim parts1() As String, parts2() As String
    parts1 = Split(key1, "|")
    parts2 = Split(key2, "|")
    
    If UBound(parts1) < 2 Or UBound(parts2) < 2 Then
        ' Fallback to string comparison
        If key1 < key2 Then
            CompareKeys = -1
        ElseIf key1 > key2 Then
            CompareKeys = 1
        Else
            CompareKeys = 0
        End If
        Exit Function
    End If
    
    Dim type1 As String, type2 As String
    Dim coord1 As String, coord2 As String
    Dim prop1 As String, prop2 As String
    
    type1 = parts1(0)
    coord1 = parts1(1)
    prop1 = parts1(2)
    
    type2 = parts2(0)
    coord2 = parts2(1)
    prop2 = parts2(2)
    
    ' First compare by Type
    If type1 < type2 Then
        CompareKeys = -1
        Exit Function
    ElseIf type1 > type2 Then
        CompareKeys = 1
        Exit Function
    End If
    
    ' Same type, compare by coordinate (numerically for Z values)
    Dim coordVal1 As Double, coordVal2 As Double
    Dim coordCompare As Integer
    coordCompare = CompareCoordinates(coord1, coord2)
    
    If coordCompare <> 0 Then
        CompareKeys = coordCompare
        Exit Function
    End If
    
    ' Same type and coordinate, compare by Property
    If prop1 < prop2 Then
        CompareKeys = -1
    ElseIf prop1 > prop2 Then
        CompareKeys = 1
    Else
        CompareKeys = 0
    End If
End Function

Private Function CompareCoordinates(coord1 As String, coord2 As String) As Integer
    ' Compare coordinates numerically
    ' Format: "Z=3500" or "X=1000" or "X=100;Y=200;Z=300"
    
    ' Extract numeric value for sorting
    Dim val1 As Double, val2 As Double
    val1 = ExtractPrimaryCoordValue(coord1)
    val2 = ExtractPrimaryCoordValue(coord2)
    
    If val1 < val2 Then
        CompareCoordinates = -1
    ElseIf val1 > val2 Then
        CompareCoordinates = 1
    Else
        ' If equal, use string comparison as fallback
        If coord1 < coord2 Then
            CompareCoordinates = -1
        ElseIf coord1 > coord2 Then
            CompareCoordinates = 1
        Else
            CompareCoordinates = 0
        End If
    End If
End Function

Private Function ExtractPrimaryCoordValue(coord As String) As Double
    ' Extract the primary numeric value from coordinate string
    ' "Z=3500" -> 3500
    ' "X=1000" -> 1000
    ' "X=100;Y=200;Z=300" -> 100 (first value)
    
    On Error Resume Next
    
    Dim parts() As String
    Dim firstPart As String
    
    ' Handle multi-coordinate format (e.g., "X=100;Y=200;Z=300")
    If InStr(coord, ";") > 0 Then
        parts = Split(coord, ";")
        firstPart = parts(0)
    Else
        firstPart = coord
    End If
    
    ' Extract number after "="
    If InStr(firstPart, "=") > 0 Then
        Dim valStr As String
        valStr = mid(firstPart, InStr(firstPart, "=") + 1)
        ExtractPrimaryCoordValue = CDbl(valStr)
    Else
        ExtractPrimaryCoordValue = 0
    End If
    
    If err.number <> 0 Then
        ExtractPrimaryCoordValue = 0
        err.Clear
    End If
    
    On Error GoTo 0
End Function




