Attribute VB_Name = "Core_CAD_Plotter"
Option Explicit
'===============================================================
' Core Module: Core_CAD_Plotter (FIXED VERSION)
' Purpose: Plot SAP2000 elements to AutoCAD (2D & 3D)
' Dependencies: Core_XData_Manager for data attachment
' FIXES:
'  - Fixed SetXData typo (missing space before parentheses)
'  - Added XData verification after attachment
'  - Improved error handling and logging
'=========================== Const ================================

Private Const APP_NAME As String = "DTS_APP"
Private Const PI As Double = 3.14159265358979

Public Const GLOBAL_TEXT_COLOR As Integer = 254 ' Label Color Global
Private Const NODE_TEXT_HEIGHT As Double = 100

' Layer names
Private Const LAYER_POINT As String = "dts_point"
Private Const LAYER_FRAME As String = "dts_frame"
Private Const LAYER_AREA As String = "dts_area"

' XData type codes
Private Const XD_APPNAME As Integer = 1001
Private Const XD_STRING As Integer = 1000
Private Const XD_REAL As Integer = 1040

' Callback for status updates
Private StatusCallback As Object


'=========================== Type ===============================
Private Type SmartTextGeo
    InsertPt(0 To 2) As Double
    Rotation As Double
    AttachmentPoint As Integer
End Type

' ----------------------------
' Public API
' ----------------------------

Public Sub SetStatusCallback(callback As Object)
    Set StatusCallback = callback
End Sub

' Plot Points with visual markers and text labels
Public Function PlotPoints(acadDoc As Object, nodeDict As Object, showNames As Boolean, LogProc As String) As Long
    On Error Resume Next
    
    LogStatus "Plotting points to AutoCAD..."
    
    Dim ms As Object
    Set ms = acadDoc.ModelSpace
    
    Dim count As Long: count = 0
    Dim nodeName As Variant
    
    For Each nodeName In nodeDict.keys
        Dim nInfo As Object
        Set nInfo = nodeDict(nodeName)
        
        Dim X As Double, Y As Double, Z As Double
        X = nInfo("X")
        Y = nInfo("Y")
        Z = nInfo("Z")
        
        ' Create circle marker
        Dim centerPt(0 To 2) As Double
        centerPt(0) = X: centerPt(1) = Y: centerPt(2) = Z
        
        Dim radius As Double
        radius = NODE_TEXT_HEIGHT / 10
        
        Dim circObj As Object
        Set circObj = ms.AddCircle(centerPt, radius)
        
        If Not circObj Is Nothing Then
            circObj.layer = LAYER_POINT
            circObj.color = 3 ' Green
            
            ' Attach metadata
            Dim springData As String
            If nInfo.exists("Spring") Then
                springData = nInfo("Spring")
            Else
                springData = ""
            End If
            
            AttachPointData circObj, CStr(nodeName), X, Y, Z, springData
            
            count = count + 1
            
            ' Log progress every 50 points
            If count Mod 50 = 0 Then
                LogStatus "Creating Node " & nodeName & " (" & _
                         Format(X, "0.00") & ", " & _
                         Format(Y, "0.00") & ", " & _
                         Format(Z, "0.00") & ")"
            End If
        End If
        
        ' Add text label if requested
        If showNames Then
            Dim textPt(0 To 2) As Double
            textPt(0) = X + radius * 2
            textPt(1) = Y + radius * 2
            textPt(2) = Z
            
            Dim txtObj As Object
            Set txtObj = ms.AddText(CStr(nodeName), textPt, NODE_TEXT_HEIGHT * 0.6)
            If Not txtObj Is Nothing Then
                txtObj.layer = LAYER_POINT
                txtObj.color = 3
            End If
        End If
    Next nodeName
    
    LogStatus "Completed: " & count & " points plotted"
    PlotPoints = count
    
    On Error GoTo 0
End Function

' Plot Frames (lines) with metadata
Public Function PlotFrames(acadDoc As Object, frameDict As Object, nodeDict As Object, showNames As Boolean) As Long
    On Error Resume Next
    
    LogStatus "Plotting frames to AutoCAD..."
    
    Dim ms As Object
    Set ms = acadDoc.ModelSpace
    
    Dim count As Long: count = 0
    Dim frameName As Variant
    
    For Each frameName In frameDict.keys
        Dim finfo As Object
        Set finfo = frameDict(frameName)
        
        Dim p1 As String, p2 As String
        p1 = finfo("P1")
        p2 = finfo("P2")
        
        ' Get coordinates from nodeDict
        If Not nodeDict.exists(p1) Or Not nodeDict.exists(p2) Then
            LogStatus "Skipped frame " & frameName & " - missing node data"
            GoTo NextFrame
        End If
        
        Dim n1 As Object, n2 As Object
        Set n1 = nodeDict(p1)
        Set n2 = nodeDict(p2)
        
        Dim startPt(0 To 2) As Double
        Dim endPt(0 To 2) As Double
        
        startPt(0) = n1("X"): startPt(1) = n1("Y"): startPt(2) = n1("Z")
        endPt(0) = n2("X"): endPt(1) = n2("Y"): endPt(2) = n2("Z")
        
        ' Create line
        Dim lineObj As Object
        Set lineObj = ms.AddLine(startPt, endPt)
        
        If Not lineObj Is Nothing Then
            lineObj.layer = LAYER_FRAME
            lineObj.color = 7 ' White
            
            ' Attach metadata
            Dim Section As String
            Section = finfo("Section")
            AttachFrameData lineObj, CStr(frameName), p1, p2, Section
            
            count = count + 1
            
            ' Log progress every 50 frames
            If count Mod 50 = 0 Then
                LogStatus "Creating Frame " & frameName & " (" & _
                         p1 & " to " & p2 & ") [" & Section & "] (" & _
                         Format(startPt(0), "0.00") & "," & Format(startPt(1), "0.00") & "," & Format(startPt(2), "0.00") & " -> " & _
                         Format(endPt(0), "0.00") & "," & Format(endPt(1), "0.00") & "," & Format(endPt(2), "0.00") & ")"
            End If
        End If
        
        ' Add label at midpoint if requested
        If showNames Then
            Dim midX As Double, midY As Double, midZ As Double
            midX = (startPt(0) + endPt(0)) / 2
            midY = (startPt(1) + endPt(1)) / 2
            midZ = (startPt(2) + endPt(2)) / 2
            
            Dim txtPt(0 To 2) As Double
            txtPt(0) = midX: txtPt(1) = midY: txtPt(2) = midZ
            
            Dim txtObj As Object
            Set txtObj = ms.AddText(CStr(frameName), txtPt, NODE_TEXT_HEIGHT * 0.8)
            If Not txtObj Is Nothing Then
                txtObj.layer = LAYER_FRAME
                txtObj.color = 7
            End If
        End If
        
NextFrame:
    Next frameName
    
    LogStatus "Completed: " & count & " frames plotted"
    PlotFrames = count
    
    On Error GoTo 0
End Function

' Plot Areas (shells/walls) - auto-detect 2D or 3D
Public Function PlotAreas(acadDoc As Object, areaDict As Object, nodeDict As Object, showNames As Boolean) As Long
    On Error Resume Next
    
    LogStatus "Plotting areas to AutoCAD..."
    
    Dim ms As Object
    Set ms = acadDoc.ModelSpace
    
    Dim count As Long: count = 0
    Dim areaName As Variant
    
    For Each areaName In areaDict.keys
        Dim aInfo As Object
        Set aInfo = areaDict(areaName)
        
        Dim pointList As String
        pointList = aInfo("PointList")
        
        If Len(pointList) = 0 Then GoTo NextArea
        
        Dim pts() As String
        pts = Split(pointList, ",")
        
        Dim nPts As Long
        nPts = UBound(pts) + 1
        
        If nPts < 3 Then GoTo NextArea
        
        ' Collect coordinates
        Dim validCount As Long: validCount = 0
        Dim coords3D() As Double
        ReDim coords3D(0 To nPts * 3 - 1)
        
        Dim j As Long
        Dim minZ As Double, maxZ As Double
        Dim firstZ As Boolean: firstZ = True
        
        For j = 0 To nPts - 1
            Dim pName As String
            pName = Trim$(pts(j))
            
            If nodeDict.exists(pName) Then
                Dim nInfo As Object
                Set nInfo = nodeDict(pName)
                
                coords3D(validCount * 3 + 0) = nInfo("X")
                coords3D(validCount * 3 + 1) = nInfo("Y")
                coords3D(validCount * 3 + 2) = nInfo("Z")
                
                ' Track Z range
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
        
        If validCount < 3 Then GoTo NextArea
        
        ReDim Preserve coords3D(0 To validCount * 3 - 1)
        
        ' Determine if horizontal (slab) or vertical/sloped (wall)
        Dim zDiff As Double
        zDiff = Abs(maxZ - minZ)
        
        Dim isHorizontal As Boolean
        isHorizontal = (zDiff < 10) ' Tolerance 10mm for horizontal
        
        Dim polyObj As Object
        
        If isHorizontal Then
            ' Use 2D polyline for horizontal slabs
            Dim coords2D() As Double
            ReDim coords2D(0 To validCount * 2 - 1)
            
            For j = 0 To validCount - 1
                coords2D(j * 2 + 0) = coords3D(j * 3 + 0)
                coords2D(j * 2 + 1) = coords3D(j * 3 + 1)
            Next j
            
            Set polyObj = ms.AddLightWeightPolyline(coords2D)
            
            If Not polyObj Is Nothing Then
                polyObj.elevation = (minZ + maxZ) / 2 ' Average Z
            End If
            
        Else
            ' Use 3D polyline for walls and sloped surfaces
            Set polyObj = ms.Add3DPoly(coords3D)
        End If
        
        If Not polyObj Is Nothing Then
            polyObj.layer = LAYER_AREA
            polyObj.color = 2 ' Yellow
            polyObj.Closed = True
            
            ' Attach metadata
            Dim Section As String
            Section = aInfo("Section")
            AttachAreaData polyObj, CStr(areaName), Section, pointList
            
            count = count + 1
            
            ' Log progress
            Dim shapeType As String
            If isHorizontal Then
                shapeType = "Slab"
            Else
                shapeType = "Wall"
            End If
            
            If count Mod 20 = 0 Then
                LogStatus "Creating " & shapeType & " " & areaName & " [" & Section & "] (" & validCount & " points, Z-diff=" & Format(zDiff, "0.0") & ")"
            End If
        End If
        
        ' Add label at centroid if requested
        If showNames Then
            Dim cx As Double, cy As Double, cz As Double
            cx = 0: cy = 0: cz = 0
            
            For j = 0 To validCount - 1
                cx = cx + coords3D(j * 3 + 0)
                cy = cy + coords3D(j * 3 + 1)
                cz = cz + coords3D(j * 3 + 2)
            Next j
            
            cx = cx / validCount
            cy = cy / validCount
            cz = cz / validCount
            
            Dim txtPt(0 To 2) As Double
            txtPt(0) = cx: txtPt(1) = cy: txtPt(2) = cz
            
            Dim txtObj As Object
            Set txtObj = ms.AddText(CStr(areaName), txtPt, NODE_TEXT_HEIGHT * 0.8)
            If Not txtObj Is Nothing Then
                txtObj.layer = LAYER_AREA
                txtObj.color = 2
            End If
        End If
        
NextArea:
    Next areaName
    
    LogStatus "Completed: " & count & " areas plotted"
    PlotAreas = count
    
    On Error GoTo 0
End Function

' ----------------------------
' Helper Functions - Metadata Attachment (FIXED)
' ----------------------------

Private Sub AttachPointData(entObj As Object, nodeName As String, X As Double, Y As Double, Z As Double, springData As String)
    On Error GoTo ErrHandler
    
    ' Register app first
    On Error Resume Next
    entObj.Application.ActiveDocument.RegisteredApplications.Add APP_NAME
    On Error GoTo ErrHandler
    
    Dim xdType() As Integer
    Dim xdVal() As Variant
    
    If springData = "" Then
        ReDim xdType(0 To 4)
        ReDim xdVal(0 To 4)
        
        xdType(0) = XD_APPNAME: xdVal(0) = APP_NAME
        xdType(1) = XD_STRING: xdVal(1) = nodeName
        xdType(2) = XD_REAL: xdVal(2) = X
        xdType(3) = XD_REAL: xdVal(3) = Y
        xdType(4) = XD_REAL: xdVal(4) = Z
    Else
        ReDim xdType(0 To 5)
        ReDim xdVal(0 To 5)
        
        xdType(0) = XD_APPNAME: xdVal(0) = APP_NAME
        xdType(1) = XD_STRING: xdVal(1) = nodeName
        xdType(2) = XD_REAL: xdVal(2) = X
        xdType(3) = XD_REAL: xdVal(3) = Y
        xdType(4) = XD_REAL: xdVal(4) = Z
        xdType(5) = XD_STRING: xdVal(5) = springData
    End If
    
    ' FIX: Add proper spacing before parentheses
    entObj.SetXData xdType, xdVal
    
    ' NEW: Verify XData was attached successfully
    Dim vType As Variant, vVal As Variant
    On Error Resume Next
    entObj.GetXData APP_NAME, vType, vVal
    If err.number <> 0 Or Not IsArray(vVal) Then
        LogStatus "WARNING: XData verification failed for point '" & nodeName & "' (handle=" & SafeHandle(entObj) & ")"
        err.Clear
    End If
    On Error GoTo ErrHandler
    
    Exit Sub
    
ErrHandler:
    LogStatus "ERROR: AttachPointData failed for '" & nodeName & "': " & err.description
End Sub

Private Sub AttachFrameData(entObj As Object, frameName As String, p1 As String, p2 As String, Section As String)
    On Error GoTo ErrHandler
    
    ' Register app
    On Error Resume Next
    entObj.Application.ActiveDocument.RegisteredApplications.Add APP_NAME
    On Error GoTo ErrHandler
    
    Dim xdType(0 To 4) As Integer
    Dim xdVal(0 To 4) As Variant
    
    xdType(0) = XD_APPNAME: xdVal(0) = APP_NAME
    xdType(1) = XD_STRING: xdVal(1) = frameName
    xdType(2) = XD_STRING: xdVal(2) = p1
    xdType(3) = XD_STRING: xdVal(3) = p2
    xdType(4) = XD_STRING: xdVal(4) = Section
    
    ' FIX: Proper spacing
    entObj.SetXData xdType, xdVal
    
    ' NEW: Verify
    Dim vType As Variant, vVal As Variant
    On Error Resume Next
    entObj.GetXData APP_NAME, vType, vVal
    If err.number <> 0 Or Not IsArray(vVal) Then
        LogStatus "WARNING: XData verification failed for frame '" & frameName & "'"
        err.Clear
    End If
    On Error GoTo ErrHandler
    
    Exit Sub
    
ErrHandler:
    LogStatus "ERROR: AttachFrameData failed for '" & frameName & "': " & err.description
End Sub

Private Sub AttachAreaData(entObj As Object, areaName As String, Section As String, pointList As String)
    On Error GoTo ErrHandler
    
    ' Register app
    On Error Resume Next
    entObj.Application.ActiveDocument.RegisteredApplications.Add APP_NAME
    On Error GoTo ErrHandler
    
    Dim xdType(0 To 3) As Integer
    Dim xdVal(0 To 3) As Variant
    
    xdType(0) = XD_APPNAME: xdVal(0) = APP_NAME
    xdType(1) = XD_STRING: xdVal(1) = areaName
    xdType(2) = XD_STRING: xdVal(2) = Section
    xdType(3) = XD_STRING: xdVal(3) = pointList
    
    ' FIX: Proper spacing
    entObj.SetXData xdType, xdVal
    
    ' NEW: Verify
    Dim vType As Variant, vVal As Variant
    On Error Resume Next
    entObj.GetXData APP_NAME, vType, vVal
    If err.number <> 0 Or Not IsArray(vVal) Then
        LogStatus "WARNING: XData verification failed for area '" & areaName & "'"
        err.Clear
    End If
    On Error GoTo ErrHandler
    
    Exit Sub
    
ErrHandler:
    LogStatus "ERROR: AttachAreaData failed for '" & areaName & "': " & err.description
End Sub

' ==========================================================================================
' CORE: PlotSmartLabel
' Features:
'   1. Calculates midpoint.
'   2. Calculates Angle.
'   3. Normalizes Angle (Readability: Bottom-Up, Left-Right).
'   4. Offsets Text (Prevents text overlapping the line).
'   5. Sets Width Factor (0.8).
' ==========================================================================================

Private Function PlotSmartLabel(acadDoc As Object, x1 As Double, y1 As Double, z1 As Double, _
        x2 As Double, y2 As Double, z2 As Double, ByVal labelText As String, _
        ByVal layerName As String, ByVal textHeight As Double, _
        Optional ByVal colorIndex As Integer = 7) As Object
    
    On Error Resume Next
    Set PlotSmartLabel = Nothing
    
    If acadDoc Is Nothing Then Exit Function
    If Len(Trim$(labelText)) = 0 Then Exit Function
    
    ' 1. Ensure layer exists
    CreateLayerIfNotExist acadDoc, layerName, colorIndex
    
    ' 2. Calculate Midpoint
    Dim midX As Double, midY As Double, midZ As Double
    midX = (x1 + x2) / 2
    midY = (y1 + y2) / 2
    midZ = (z1 + z2) / 2
    
    ' 3. Calculate Angle
    Dim dx As Double, dy As Double
    dx = x2 - x1
    dy = y2 - y1
    
    Dim angle As Double
    angle = 0
    If Abs(dx) > 0.000001 Or Abs(dy) > 0.000001 Then
        angle = Atan2(dy, dx)
    End If
    
    ' 4. Normalize Angle for Readability (Text should face Up or Left)
    ' Rule: If angle is between 90 and 270 degrees (exclusive), rotate 180
    Dim readableAngle As Double
    readableAngle = angle
    
    ' Convert to positive 0-2PI for checking
    Dim checkAng As Double
    checkAng = angle
    If checkAng < 0 Then checkAng = checkAng + 2 * PI
    
    ' Check ranges: > 90 degrees (PI/2) and <= 270 degrees (3*PI/2)
    If checkAng > (PI / 2 + 0.001) And checkAng <= (3 * PI / 2 + 0.001) Then
        readableAngle = readableAngle + PI
    End If
    
    ' 5. Calculate Offset (Move text slightly ABOVE the line relative to text rotation)
    ' Offset distance = half text height + gap
    Dim offsetDist As Double
    offsetDist = textHeight * 0.6 ' 0.5 is half height, + 0.1 gap
    
    ' Vector perpendicular to readableAngle (Rotate 90 deg CCW)
    Dim offX As Double, offY As Double
    offX = -Sin(readableAngle) * offsetDist
    offY = Cos(readableAngle) * offsetDist
    
    ' Apply offset to insertion point
    Dim insPt(0 To 2) As Double
    insPt(0) = midX + offX
    insPt(1) = midY + offY
    insPt(2) = midZ
    
    ' 6. Create Text Object
    Dim textObj As Object
    Set textObj = acadDoc.ModelSpace.AddText(labelText, insPt, textHeight)
    
    If Not textObj Is Nothing Then
        ' Set properties
        textObj.layer = layerName
        textObj.color = colorIndex
        textObj.Rotation = readableAngle
        textObj.Alignment = 10 ' acAlignmentMiddleCenter = 10 (Better for offset control)
        textObj.TextAlignmentPoint = insPt
        
        ' Set Width Factor (0.8 as requested)
        textObj.scaleFactor = 0.8
        
        Set PlotSmartLabel = textObj
    End If
End Function
' ==========================================================================================
' HELPER: Atan2
' ==========================================================================================
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
' Public API: generic label plotting and element-specific wrappers
Public Sub PlotLabel(acadDoc As Object, X As Double, Y As Double, Z As Double, _
        ByVal labelText As String, Optional ByVal layerName As String = "dts_frame_label", _
        Optional ByVal textHeight As Double = 80, Optional ByVal colorIndex As Long = 7)
    ' Draw a text label at the specified coordinates on the specified layer.
    ' This is a generic utility; callers may use the element-specific wrappers below.
    On Error Resume Next
    If acadDoc Is Nothing Then Exit Sub
    If Trim$(CStr(labelText)) = "" Then Exit Sub

    EnsureLayerExistsLocal acadDoc, layerName, colorIndex

    Dim ms As Object
    Set ms = acadDoc.ModelSpace

    Dim txtPt(0 To 2) As Double
    txtPt(0) = X
    txtPt(1) = Y
    txtPt(2) = Z

    Dim txtObj As Object
    Set txtObj = ms.AddText(CStr(labelText), txtPt, CDbl(textHeight))
    If Not txtObj Is Nothing Then
        On Error Resume Next
        txtObj.layer = layerName
        txtObj.color = CLng(colorIndex)
        On Error GoTo 0
    End If
End Sub

Public Sub PlotLabelAtMidpoint(acadDoc As Object, _
        x1 As Double, y1 As Double, z1 As Double, _
        x2 As Double, y2 As Double, z2 As Double, _
        ByVal labelText As String, Optional ByVal layerName As String = "dts_frame_label", _
        Optional ByVal textHeight As Double = 80, Optional ByVal colorIndex As Long = 7)
    ' Convenience: compute midpoint between two points and call PlotLabel
    On Error Resume Next
    Dim mx As Double, my As Double, mz As Double
    mx = (CDbl(x1) + CDbl(x2)) / 2
    my = (CDbl(y1) + CDbl(y2)) / 2
    mz = (CDbl(z1) + CDbl(z2)) / 2

    PlotLabel acadDoc, mx, my, mz, labelText, layerName, textHeight, colorIndex
End Sub

Public Sub PlotNodeLabel(acadDoc As Object, X As Double, Y As Double, Z As Double, _
        ByVal labelText As String, Optional ByVal textHeight As Double = 60)
    ' Wrapper for node labels
    PlotLabel acadDoc, X, Y, Z, labelText, "dts_node_label", textHeight, 3
End Sub

Public Sub PlotFrameLabel(acadDoc As Object, x1 As Double, y1 As Double, z1 As Double, _
        x2 As Double, y2 As Double, z2 As Double, ByVal labelText As String, _
        Optional ByVal textHeight As Double = 80)
    ' Wrapper for frame labels (place at midpoint)
    PlotLabelAtMidpoint acadDoc, x1, y1, z1, x2, y2, z2, labelText, "dts_frame_label", textHeight, 7
End Sub
' ==========================================================================================
' PUBLIC: PlotFrameLabelEx (Create NEW MText)
' ==========================================================================================
Public Function PlotFrameLabelEx(acadDoc As Object, x1 As Double, y1 As Double, z1 As Double, _
        x2 As Double, y2 As Double, z2 As Double, ByVal labelText As String, _
        ByVal textHeight As Double) As Object
    
    On Error Resume Next
    Set PlotFrameLabelEx = Nothing
    
    ' 1. Calculate Geometry
    Dim geo As SmartTextGeo
    geo = CalcSmartGeometry(x1, y1, z1, x2, y2, z2, textHeight)
    
    ' 2. Create MText Object (Supports formatting codes like \C1; for color)
    ' Note: AddMText takes (InsertionPoint, Width, Text)
    Dim mtextObj As Object
    Set mtextObj = acadDoc.ModelSpace.AddMText(geo.InsertPt, 0, labelText)
    
    If Not mtextObj Is Nothing Then
        With mtextObj
            .layer = "dts_frame_label"
            .height = textHeight
            .color = GLOBAL_TEXT_COLOR '
            .Rotation = geo.Rotation
            .AttachmentPoint = geo.AttachmentPoint
            .insertionPoint = geo.InsertPt
            
            ' Optional: Background mask
            ' .BackgroundFill = True
        End With
        
        Set PlotFrameLabelEx = mtextObj
    End If
End Function
' ==========================================================================================
' PRIVATE: Core Geometry Calculation Logic (Updated for MText)
' ==========================================================================================
Private Function CalcSmartGeometry(x1 As Double, y1 As Double, z1 As Double, _
        x2 As Double, y2 As Double, z2 As Double, textHeight As Double) As SmartTextGeo
        
    Dim res As SmartTextGeo
    
    ' 1. Midpoint
    Dim midX As Double, midY As Double, midZ As Double
    midX = (x1 + x2) / 2
    midY = (y1 + y2) / 2
    midZ = (z1 + z2) / 2
    
    ' 2. Angle
    Dim dx As Double, dy As Double
    dx = x2 - x1
    dy = y2 - y1
    
    Dim angle As Double: angle = 0
    If Abs(dx) > 0.000001 Or Abs(dy) > 0.000001 Then
        angle = Atan2(dy, dx)
    End If
    
    ' 3. Normalize Angle (Readable: Bottom-Up, Left-Right)
    Dim readableAngle As Double: readableAngle = angle
    Dim checkAng As Double: checkAng = angle
    If checkAng < 0 Then checkAng = checkAng + 2 * PI
    
    If checkAng > (PI / 2 + 0.001) And checkAng <= (3 * PI / 2 + 0.001) Then
        readableAngle = readableAngle + PI
    End If
    
    ' 4. Calculate Offset (Nh?c cao lên 1 kho?ng b?ng 0.8 chi?u cao ch?)
    ' MText MiddleCenter: InsertPoint is the center of the text box.
    Dim offsetDist As Double
    offsetDist = textHeight * 0.8
    
    Dim offX As Double, offY As Double
    offX = -Sin(readableAngle) * offsetDist
    offY = Cos(readableAngle) * offsetDist
    
    res.InsertPt(0) = midX + offX
    res.InsertPt(1) = midY + offY
    res.InsertPt(2) = midZ
    res.Rotation = readableAngle
    
    ' 5 = acAttachmentPointMiddleCenter (Dành riêng cho MText)
    res.AttachmentPoint = 5
    
    CalcSmartGeometry = res
End Function
' ==========================================================================================
' PUBLIC: UpdateFrameLabelEx (Update EXISTING MText)
' ==========================================================================================
Public Sub UpdateFrameLabelEx(textObj As Object, x1 As Double, y1 As Double, z1 As Double, _
        x2 As Double, y2 As Double, z2 As Double, ByVal labelText As String, _
        ByVal textHeight As Double)
        
    On Error Resume Next
    If textObj Is Nothing Then Exit Sub
    
    ' 1. Calculate Geometry
    Dim geo As SmartTextGeo
    geo = CalcSmartGeometry(x1, y1, z1, x2, y2, z2, textHeight)
    
    ' 2. Update Properties (MText)
    With textObj
        ' Check Object Type to avoid errors if switching from Text to MText manually
        If .ObjectName = "AcDbMText" Then
            If .TextString <> labelText Then .TextString = labelText
            If Abs(.height - textHeight) > 0.001 Then .height = textHeight
            
            ' Reset base color if needed (individual parts formatted by string will override this)
            If .color <> GLOBAL_TEXT_COLOR Then .color = GLOBAL_TEXT_COLOR
            
            If Abs(.Rotation - geo.Rotation) > 0.001 Then .Rotation = geo.Rotation
            
            ' Update Location
            If .AttachmentPoint <> geo.AttachmentPoint Then .AttachmentPoint = geo.AttachmentPoint
            
            Dim curPt As Variant
            curPt = .insertionPoint
            If Abs(curPt(0) - geo.InsertPt(0)) > 0.001 Or _
               Abs(curPt(1) - geo.InsertPt(1)) > 0.001 Then
                .insertionPoint = geo.InsertPt
            End If
        Else
            ' Fallback if user manually exploded text (rare case): Do nothing or delete/recreate
            ' For performance, we ignore non-MText updates or delete/recreate in Refresh logic
        End If
    End With
End Sub

' ==========================================================================================
' HELPER: PlotLabelAtMidpointEx (Function version - Returns text object)
' ==========================================================================================
Private Function PlotLabelAtMidpointEx(acadDoc As Object, x1 As Double, y1 As Double, z1 As Double, _
        x2 As Double, y2 As Double, z2 As Double, ByVal labelText As String, _
        ByVal layerName As String, Optional ByVal textHeight As Double = 80, _
        Optional ByVal colorIndex As Integer = 7) As Object
    On Error Resume Next
    Set PlotLabelAtMidpointEx = Nothing
    
    If acadDoc Is Nothing Then Exit Function
    If Len(Trim$(labelText)) = 0 Then Exit Function
    
    ' Ensure layer exists
    CreateLayerIfNotExist acadDoc, layerName, colorIndex
    
    ' Calculate midpoint
    Dim midX As Double, midY As Double, midZ As Double
    midX = (x1 + x2) / 2
    midY = (y1 + y2) / 2
    midZ = (z1 + z2) / 2
    
    ' Calculate angle
    Dim dx As Double, dy As Double
    dx = x2 - x1
    dy = y2 - y1
    Dim angle As Double
    angle = 0
    If Abs(dx) > 0.001 Or Abs(dy) > 0.001 Then
        angle = VBA.Math.Atn(dy / dx)
        If dx < 0 Then angle = angle + 3.14159265358979
    End If
    
    ' Create insertion point
    Dim insPt(0 To 2) As Double
    insPt(0) = midX
    insPt(1) = midY
    insPt(2) = midZ
    
    ' Create text object
    Dim textObj As Object
    Set textObj = acadDoc.ModelSpace.AddText(labelText, insPt, textHeight)
    
    If Not textObj Is Nothing Then
        ' Set properties
        textObj.layer = layerName
        textObj.Rotation = angle
        textObj.Alignment = 4 ' acAlignmentMiddleCenter
        textObj.TextAlignmentPoint = insPt
        
        ' Return the created object
        Set PlotLabelAtMidpointEx = textObj
    End If
End Function

' ==========================================================================================
' PUBLIC: Ensure Layer Exists (Call this ONCE before loops)
' ==========================================================================================
Public Sub EnsureLabelLayer(acadDoc As Object, Optional colorIndex As Integer = 7)
    On Error Resume Next
    If acadDoc Is Nothing Then Exit Sub
    Dim layerName As String: layerName = "dts_frame_label"
    
    Dim layer As Object
    Set layer = acadDoc.layers.item(layerName)
    
    If layer Is Nothing Or err.number <> 0 Then
        err.Clear
        Set layer = acadDoc.layers.Add(layerName)
        If Not layer Is Nothing Then layer.color = colorIndex
    End If
    On Error GoTo 0
End Sub
' ==========================================================================================
' HELPER: Create Layer If Not Exist
' ==========================================================================================
Private Sub CreateLayerIfNotExist(acadDoc As Object, layerName As String, Optional colorIndex As Integer = 7)
    On Error Resume Next
    
    If acadDoc Is Nothing Then Exit Sub
    If Len(Trim$(layerName)) = 0 Then Exit Sub
    
    ' Try to get existing layer
    Dim layer As Object
    Set layer = Nothing
    Set layer = acadDoc.layers.item(layerName)
    
    ' If not found, create new
    If layer Is Nothing Or err.number <> 0 Then
        err.Clear
        Set layer = acadDoc.layers.Add(layerName)
        
        If Not layer Is Nothing Then
            layer.color = colorIndex
        End If
    End If
    
    On Error GoTo 0
End Sub
Public Sub PlotAreaLabel(acadDoc As Object, X As Double, Y As Double, Z As Double, _
        ByVal labelText As String, Optional ByVal textHeight As Double = 80)
    ' Wrapper for area labels (centroid)
    PlotLabel acadDoc, X, Y, Z, labelText, "dts_area_label", textHeight, 2
End Sub

' Local helper to ensure layer exists in drawing (kept local to avoid conflicts)
Private Sub EnsureLayerExistsLocal(acadDoc As Object, layerName As String, colorIndex As Long)
    On Error Resume Next
    If acadDoc Is Nothing Then Exit Sub

    Dim lay As Object
    Set lay = Nothing
    Set lay = acadDoc.layers.item(layerName)
    If err.number <> 0 Then
        err.Clear
        Set lay = acadDoc.layers.Add(layerName)
        On Error Resume Next
        lay.color = colorIndex
        On Error GoTo 0
    End If
    On Error GoTo 0
End Sub

' Helper to get handle safely
Private Function SafeHandle(obj As Object) As String
    On Error Resume Next
    SafeHandle = CStr(obj.Handle)
    If err.number <> 0 Then
        err.Clear
        SafeHandle = "<unknown>"
    End If
    On Error GoTo 0
End Function

' ----------------------------
' Status Logging
' ----------------------------

Public Sub LogStatus(msg As String)
    On Error Resume Next
    If Not StatusCallback Is Nothing Then
        StatusCallback.SetStatus msg
    End If
    Debug.Print msg
    On Error GoTo 0
End Sub

