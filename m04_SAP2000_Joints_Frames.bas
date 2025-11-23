Attribute VB_Name = "m04_SAP2000_Joints_Frames"
Option Explicit
'===============================================================
' Module: modSAP2000_JointsFrames
' Purpose: Extract joints & frames, write GeometryData
'===============================================================

Public Sub ExtractPoints()
    Dim ret As Long
    ret = SapModel.pointObj.GetNameList(gPointCount, gPointNames)
    CheckRet ret, "PointObj.GetNameList"
    
    If gPointCount > 0 Then
        Dim arrX() As Double, arrY() As Double, arrZ() As Double
        ReDim arrX(gPointCount - 1)
        ReDim arrY(gPointCount - 1)
        ReDim arrZ(gPointCount - 1)
        
        Dim i As Long
        For i = 0 To gPointCount - 1
            ret = SapModel.pointObj.GetCoordCartesian(gPointNames(i), arrX(i), arrY(i), arrZ(i))
            CheckRet ret, "PointObj.GetCoordCartesian(" & gPointNames(i) & ")"
        Next
        
        Set gDictX = Combine2ArrayToDictionary(gPointNames, arrX)
        Set gDictY = Combine2ArrayToDictionary(gPointNames, arrY)
        Set gDictZ = Combine2ArrayToDictionary(gPointNames, arrZ)
    Else
        LogMsg "ExtractPoints: No joints found."
    End If
End Sub

Public Sub ExtractFrames()
    Dim ret As Long
    ret = SapModel.frameObj.GetNameList(gFrameCount, gFrameNames)
    CheckRet ret, "FrameObj.GetNameList"
    If gFrameCount = 0 Then
        LogMsg "ExtractFrames: No frames found."
        Exit Sub
    End If
    
    ReDim gFrameP1(gFrameCount - 1)
    ReDim gFrameP2(gFrameCount - 1)
    ReDim gFrameProp(gFrameCount - 1)
    ReDim gFrameAngle(gFrameCount - 1)
    
    Dim i As Long, sAuto As String, advLocal As Boolean
    For i = 0 To gFrameCount - 1
        ret = SapModel.frameObj.GetPoints(gFrameNames(i), gFrameP1(i), gFrameP2(i))
        CheckRet ret, "FrameObj.GetPoints(" & gFrameNames(i) & ")"
        ret = SapModel.frameObj.GetSection(gFrameNames(i), gFrameProp(i), sAuto)
        CheckRet ret, "FrameObj.GetSection(" & gFrameNames(i) & ")"
        ret = SapModel.frameObj.GetLocalAxes(gFrameNames(i), gFrameAngle(i), advLocal)
        CheckRet ret, "FrameObj.GetLocalAxes(" & gFrameNames(i) & ")"
    Next
End Sub

Public Sub WriteGeometryData()
    Dim ws As Worksheet
    Set ws = SheetOrCreate("GeometryData", True)
    If ws Is Nothing Then
        LogMsg "WriteGeometryData: Cannot access sheet."
        Exit Sub
    End If
    
    ws.Range("A3:L" & ws.rows.count).clearContents
    If gFrameCount = 0 Or IsArrayEmpty(gFrameNames) Then Exit Sub
    
    ws.Range("A3").Resize(gFrameCount, 1).Value = ToVerticalVariant(gFrameNames)
    ws.Range("B3").Resize(gFrameCount, 1).Value = ToVerticalVariant(gFrameP1)
    ws.Range("C3").Resize(gFrameCount, 1).Value = ToVerticalVariant(gFrameP2)
    ws.Range("D3").Resize(gFrameCount, 1).Value = ToVerticalVariant(gFrameProp)
    ws.Range("E3").Resize(gFrameCount, 1).Value = ToVerticalVariant(gFrameAngle)
    
    Dim i As Long, rowPtr As Long, j1 As String, j2 As String
    For i = 0 To gFrameCount - 1
        rowPtr = 3 + i
        j1 = gFrameP1(i): j2 = gFrameP2(i)
        
        ws.Cells(rowPtr, "G").Value = gDictX(j1)
        ws.Cells(rowPtr, "H").Value = gDictY(j1)
        ws.Cells(rowPtr, "I").Value = gDictZ(j1)
        ws.Cells(rowPtr, "J").Value = gDictX(j2)
        ws.Cells(rowPtr, "K").Value = gDictY(j2)
        ws.Cells(rowPtr, "L").Value = gDictZ(j2)
        
        ws.Cells(rowPtr, "F").Value = Sqr( _
            (gDictX(j1) - gDictX(j2)) ^ 2 + _
            (gDictY(j1) - gDictY(j2)) ^ 2 + _
            (gDictZ(j1) - gDictZ(j2)) ^ 2)
    Next
End Sub

