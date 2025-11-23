Attribute VB_Name = "m05_SAP2000_Areas"
Option Explicit
'===============================================================
' Module: modSAP2000_Areas
' Purpose: Extract area objects & write AreaData
'===============================================================

Public Sub ExtractAreas()
    If Not ENABLE_AREAS Then Exit Sub
    
    Dim ret As Long
    ret = SapModel.AreaObj.GetNameList(gAreaCount, gAreaNames)
    CheckRet ret, "AreaObj.GetNameList"
    If gAreaCount = 0 Then
        LogMsg "ExtractAreas: No area objects."
        Exit Sub
    End If
    
    ReDim gAreaProp(gAreaCount - 1)
    ReDim gAreaNumPts(gAreaCount - 1)
    ReDim gAreaPointStr(gAreaCount - 1)
    ReDim gAreaCentX(gAreaCount - 1)
    ReDim gAreaCentY(gAreaCount - 1)
    ReDim gAreaCentZ(gAreaCount - 1)
    ReDim gAreaGeomArea(gAreaCount - 1)
    ReDim gAreaNx(gAreaCount - 1)
    ReDim gAreaNy(gAreaCount - 1)
    ReDim gAreaNz(gAreaCount - 1)
    
    Dim a As Long
    For a = 0 To gAreaCount - 1
        Dim nm As String: nm = gAreaNames(a)
        
        ' Property (backward compatibility)
        gAreaProp(a) = ""
        On Error Resume Next
        ret = SapModel.AreaObj.GetSection(nm, gAreaProp(a))
        If err.number <> 0 Then
            err.Clear
            ret = -1
        End If
        On Error GoTo 0
        
        If ret <> 0 Or Len(gAreaProp(a)) = 0 Then
            On Error Resume Next
            ret = SapModel.AreaObj.GetProperty(nm, gAreaProp(a))
            If err.number <> 0 Then
                err.Clear
                ret = -1
            End If
            On Error GoTo 0
        End If
        
        ' Points
        Dim nPts As Long, pts() As String
        On Error Resume Next
        ret = SapModel.AreaObj.GetPoints(nm, nPts, pts)
        If err.number <> 0 Then
            err.Clear
            nPts = 0
        End If
        On Error GoTo 0
        gAreaNumPts(a) = nPts
        If nPts > 0 Then
            gAreaPointStr(a) = Join(pts, ",")
        Else
            gAreaPointStr(a) = ""
        End If
        
        ' Centroid
        Dim cx As Double, cy As Double, cz As Double, centroidOK As Boolean
        centroidOK = False
        If nPts > 0 Then
            On Error Resume Next
            ret = SapModel.AreaObj.GetCentroid(nm, cx, cy, cz)
            If err.number = 0 And ret = 0 Then
                centroidOK = True
            Else
                err.Clear
            End If
            On Error GoTo 0
            
            If centroidOK Then
                gAreaCentX(a) = cx: gAreaCentY(a) = cy: gAreaCentZ(a) = cz
            Else
                Dim sx As Double, sy As Double, sz As Double, PI As Long
                For PI = 0 To nPts - 1
                    sx = sx + gDictX(pts(PI))
                    sy = sy + gDictY(pts(PI))
                    sz = sz + gDictZ(pts(PI))
                Next
                gAreaCentX(a) = sx / nPts
                gAreaCentY(a) = sy / nPts
                gAreaCentZ(a) = sz / nPts
            End If
        End If
        
        ' Area value
        On Error Resume Next
        ret = SapModel.AreaObj.GetArea(nm, gAreaGeomArea(a))
        If err.number <> 0 Or ret <> 0 Then
            err.Clear
            gAreaGeomArea(a) = ApproxAreaFromPointString(gAreaPointStr(a))
        End If
        On Error GoTo 0
        
        ' Normal
        If nPts >= 3 Then
            Dim pa As String, pb As String, pC As String
            pa = pts(0): pb = pts(1): pC = pts(2)
            Dim v1x As Double, v1y As Double, v1z As Double
            Dim v2x As Double, v2y As Double, v2z As Double
            v1x = gDictX(pb) - gDictX(pa)
            v1y = gDictY(pb) - gDictY(pa)
            v1z = gDictZ(pb) - gDictZ(pa)
            v2x = gDictX(pC) - gDictX(pa)
            v2y = gDictY(pC) - gDictY(pa)
            v2z = gDictZ(pC) - gDictZ(pa)
            Dim nx As Double, ny As Double, nz As Double, nlen As Double
            nx = v1y * v2z - v1z * v2y
            ny = v1z * v2x - v1x * v2z
            nz = v1x * v2y - v1y * v2x
            nlen = Sqr(nx * nx + ny * ny + nz * nz)
            If nlen > 0# Then
                gAreaNx(a) = nx / nlen
                gAreaNy(a) = ny / nlen
                gAreaNz(a) = nz / nlen
            End If
        End If
    Next
End Sub

Public Sub WriteAreaData()
    If Not ENABLE_AREAS Then Exit Sub
    If gAreaCount = 0 Or IsArrayEmpty(gAreaNames) Then Exit Sub
    
    Dim ws As Worksheet
    Set ws = SheetOrCreate("AreaData", True)
    If ws Is Nothing Then
        LogMsg "WriteAreaData: Cannot create AreaData."
        Exit Sub
    End If
    
    ws.Cells.clearContents
    ws.Range("A1:K1").Value = Array( _
        "AreaName", "Property", "NumPoints", "PointList", "CentroidX", "CentroidY", "CentroidZ", _
        "AreaValue", "NormalX", "NormalY", "NormalZ")
    
    Dim n As Long: n = gAreaCount
    ws.Range("A2").Resize(n, 1).Value = ToVerticalVariant(gAreaNames)
    ws.Range("B2").Resize(n, 1).Value = ToVerticalVariant(gAreaProp)
    ws.Range("C2").Resize(n, 1).Value = ToVerticalVariant(gAreaNumPts)
    ws.Range("D2").Resize(n, 1).Value = ToVerticalVariant(gAreaPointStr)
    ws.Range("E2").Resize(n, 1).Value = ToVerticalVariant(gAreaCentX)
    ws.Range("F2").Resize(n, 1).Value = ToVerticalVariant(gAreaCentY)
    ws.Range("G2").Resize(n, 1).Value = ToVerticalVariant(gAreaCentZ)
    ws.Range("H2").Resize(n, 1).Value = ToVerticalVariant(gAreaGeomArea)
    ws.Range("I2").Resize(n, 1).Value = ToVerticalVariant(gAreaNx)
    ws.Range("J2").Resize(n, 1).Value = ToVerticalVariant(gAreaNy)
    ws.Range("K2").Resize(n, 1).Value = ToVerticalVariant(gAreaNz)
End Sub

Private Function ApproxAreaFromPointString(ByVal pointStr As String) As Double
    If Len(pointStr) = 0 Then Exit Function
    Dim pts() As String
    pts = Split(pointStr, ",")
    Dim n As Long: n = UBound(pts) - LBound(pts) + 1
    If n < 3 Then Exit Function
    
    Dim p0 As String: p0 = pts(0)
    Dim i As Long, total As Double
    For i = 1 To n - 2
        Dim pa As String, pb As String
        pa = pts(i): pb = pts(i + 1)
        Dim v1x As Double, v1y As Double, v1z As Double
        Dim v2x As Double, v2y As Double, v2z As Double
        v1x = gDictX(pa) - gDictX(p0)
        v1y = gDictY(pa) - gDictY(p0)
        v1z = gDictZ(pa) - gDictZ(p0)
        v2x = gDictX(pb) - gDictX(p0)
        v2y = gDictY(pb) - gDictY(p0)
        v2z = gDictZ(pb) - gDictZ(p0)
        Dim cx As Double, cy As Double, cz As Double
        cx = v1y * v2z - v1z * v2y
        cy = v1z * v2x - v1x * v2z
        cz = v1x * v2y - v1y * v2x
        total = total + 0.5 * Sqr(cx * cx + cy * cy + cz * cz)
    Next
    ApproxAreaFromPointString = total
End Function

