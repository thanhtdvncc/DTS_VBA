Attribute VB_Name = "LibDTS_DriverSAP"
' Module: LibDTS_DriverSAP
Option Explicit

Private SapObject As Object
Private SapModel As Object

' ==========================================
' 1. CONNECTION
' ==========================================

' Connect to active SAP2000 instance or start new
' Ket noi voi SAP2000 dang chay hoac khoi dong moi
Public Function Connect() As Boolean
    On Error Resume Next
    Set SapObject = GetObject(, "CSI.SAP2000.API.SapObject")
    If SapObject Is Nothing Then
        Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")
        SapObject.ApplicationStart
    End If
    
    If Not SapObject Is Nothing Then
        Set SapModel = SapObject.SapModel
        SapModel.InitializeNewModel 5 ' kN_m_C
        Connect = True
    Else
        Connect = False
    End If
End Function

' ==========================================
' 2. MODELING
' ==========================================

' Push a Frame object to SAP
' Day doi tuong Frame sang SAP
Public Function PushFrame(frame As clsDTSFrame) As String
    If SapModel Is Nothing Then Exit Function
    
    Dim ret As Long
    Dim frameName As String
    
    ' 1. Ensure Points exist
    ' 1. Dam bao cac diem ton tai
    Dim p1 As String, p2 As String
    p1 = CreatePoint(frame.StartPoint)
    p2 = CreatePoint(frame.EndPoint)
    
    ' 2. Add Frame
    ' 2. Them Frame
    ret = SapModel.frameObj.AddByPoint(p1, p2, frameName, frame.sectionName)
    
    If ret = 0 Then
        ' 3. Sync GUID (Store DTS GUID in SAP)
        ' 3. Dong bo GUID (Luu DTS GUID vao SAP)
        ' Try to set GUID directly if supported, else use Comment
        ' Thu dat GUID truc tiep neu ho tro, neu khong dung Comment
        On Error Resume Next
        SapModel.frameObj.SetGUID frameName, frame.Base.guid
        If err.number <> 0 Then
             ' Fallback for older SAP versions
             ' Phuong an du phong cho SAP phien ban cu
             ' SapModel.FrameObj.SetComment frameName, "GUID:" & frame.Base.GUID
        End If
        On Error GoTo 0
        
        PushFrame = frameName
    End If
End Function

' Helper: Create SAP Point
' Ham phu: Tao diem SAP
Private Function CreatePoint(pt As clsDTSPoint) As String
    Dim pName As String
    Dim ret As Long
    
    ' Add point (SAP handles duplicates automatically by tolerance usually)
    ' Them diem (SAP thuong tu xu ly trung lap theo sai so)
    ret = SapModel.pointObj.AddCartesian(pt.X, pt.Y, pt.Z, pName)
    CreatePoint = pName
End Function
