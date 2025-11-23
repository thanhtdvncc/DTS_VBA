Attribute VB_Name = "LibDTS_DriverCAD"
' Module: LibDTS_DriverCAD
Option Explicit

' ==========================================
' 1. DRAWING OPERATIONS
' ==========================================

' Draw a Frame (Beam/Column)
' Ve Thanh (Dam/Cot)
Public Function DrawFrame(frameObj As clsDTSFrame, acadDoc As Object) As Object
    Dim ms As Object
    Set ms = acadDoc.ModelSpace
    
    ' Draw Line
    ' Ve duong thang
    Dim p1(0 To 2) As Double, p2(0 To 2) As Double
    p1(0) = frameObj.StartPoint.X: p1(1) = frameObj.StartPoint.Y: p1(2) = frameObj.StartPoint.Z
    p2(0) = frameObj.EndPoint.X: p2(1) = frameObj.EndPoint.Y: p2(2) = frameObj.EndPoint.Z
    
    Dim lineObj As Object
    Set lineObj = ms.AddLine(p1, p2)
    
    ' Set Properties
    ' Thiet lap thuoc tinh
    On Error Resume Next
    lineObj.layer = frameObj.Base.layer
    lineObj.color = frameObj.Base.color
    On Error GoTo 0
    
    ' Save Metadata
    ' Luu du lieu
    SaveXData frameObj, lineObj
    
    Set DrawFrame = lineObj
End Function

' Draw a Tag (Text)
' Ve Nhan (Text)
Public Function DrawTag(tagObj As clsDTSTag, acadDoc As Object) As Object
    Dim ms As Object
    Set ms = acadDoc.ModelSpace
    
    Dim insPt(0 To 2) As Double
    insPt(0) = tagObj.Position.X: insPt(1) = tagObj.Position.Y: insPt(2) = tagObj.Position.Z
    
    Dim textObj As Object
    Set textObj = ms.AddText(tagObj.TextContent, insPt, tagObj.textHeight)
    
    textObj.Rotation = tagObj.Rotation
    textObj.layer = tagObj.Base.layer
    
    SaveXData tagObj, textObj
    
    Set DrawTag = textObj
End Function

' ==========================================
' 2. DATA OPERATIONS (XDATA)
' ==========================================

' Generic Save function for any DTS Object
' Ham luu tong quat cho moi doi tuong DTS
Public Sub SaveXData(dtsObj As Object, ent As Object)
    ' Register App
    ' Dang ky App
    On Error Resume Next
    ent.Document.RegisteredApplications.Add LibDTS_Global.DTS_APP_NAME
    On Error GoTo 0
    
    ' Check Identity (Self-Healing)
    ' Kiem tra dinh danh (Tu chua lanh)
    ' Accessing .Base property requires the object to be a DTS Class
    ' Truy cap thuoc tinh .Base yeu cau doi tuong phai la DTS Class
    dtsObj.Base.ValidateIdentity ent.Handle
    
    ' Serialize content
    ' Dong goi noi dung
    Dim dataStr As String
    dataStr = dtsObj.Serialize()
    
    ' Write to XData
    ' Ghi vao XData
    Dim xType(0 To 1) As Integer
    Dim xVal(0 To 1) As Variant
    
    xType(0) = 1001: xVal(0) = LibDTS_Global.DTS_APP_NAME
    xType(1) = 1000: xVal(1) = dataStr
    
    ent.SetXData xType, xVal
End Sub

' Read Frame from Entity
' Doc Frame tu Entity
Public Function ReadFrame(ent As Object) As clsDTSFrame
    If ent Is Nothing Then Exit Function
    If ent.ObjectName <> "AcDbLine" Then Exit Function
    
    Dim frame As New clsDTSFrame
    
    ' 1. Geometry from CAD
    ' 1. Hinh hoc tu CAD
    frame.StartPoint.Init ent.StartPoint(0), ent.StartPoint(1), ent.StartPoint(2)
    frame.EndPoint.Init ent.EndPoint(0), ent.EndPoint(1), ent.EndPoint(2)
    
    ' 2. Data from XData
    ' 2. Du lieu tu XData
    Dim xDataStr As String
    xDataStr = GetRawXData(ent)
    
    If Len(xDataStr) > 0 Then
        frame.Deserialize xDataStr
        ' Heal identity if needed
        ' Chua lanh dinh danh neu can
        If frame.Base.CheckAndHealIdentity(ent.Handle) Then
            SaveXData frame, ent
        End If
    End If
    
    Set ReadFrame = frame
End Function

' Helper: Get raw string
' Ham phu: Lay chuoi tho
Private Function GetRawXData(ent As Object) As String
    Dim xType As Variant, xVal As Variant
    On Error Resume Next
    ent.GetXData LibDTS_Global.DTS_APP_NAME, xType, xVal
    If err.number = 0 And Not IsEmpty(xVal) Then
        GetRawXData = CStr(xVal(1))
    End If
    On Error GoTo 0
End Function
