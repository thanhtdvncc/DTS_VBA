VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDTSFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class Module: clsDTSFrame
Option Explicit

' --- COMPOSITION ---
Public Base As clsDTSElement

' --- PRIVATE DATA ---
Private p_StartPoint As clsDTSPoint
Private p_EndPoint As clsDTSPoint
Private p_Section As clsDTSSection
Private p_FrameType As DTSFrameType

' --- CONSTRUCTOR ---
Private Sub Class_Initialize()
    Set Base = New clsDTSElement
    Base.ElementType = DTS_ELEM_FRAME
    
    Set p_StartPoint = New clsDTSPoint
    Set p_EndPoint = New clsDTSPoint
    Set p_Section = New clsDTSSection
    
    p_FrameType = DTS_FRM_BEAM ' Default
End Sub

' --- CLEANUP ---
Private Sub Class_Terminate()
    Set Base = Nothing
    Set p_StartPoint = Nothing
    Set p_EndPoint = Nothing
    Set p_Section = Nothing
End Sub

' --- PROPERTIES ---

Public Property Get StartPoint() As clsDTSPoint
    Set StartPoint = p_StartPoint
End Property
Public Property Set StartPoint(Value As clsDTSPoint)
    Set p_StartPoint = Value
    Base.Properties("GeoChanged") = True ' Mark geometry as changed
End Property

Public Property Get EndPoint() As clsDTSPoint
    Set EndPoint = p_EndPoint
End Property
Public Property Set EndPoint(Value As clsDTSPoint)
    Set p_EndPoint = Value
    Base.Properties("GeoChanged") = True
End Property

Public Property Get Section() As clsDTSSection
    Set Section = p_Section
End Property

' --- LOGIC METHODS ---

' Calculated Property: Length
Public Property Get Length() As Double
    Length = p_StartPoint.DistanceTo(p_EndPoint)
End Property

' Logic: Check validity
Public Function IsValid() As Boolean
    IsValid = (Me.Length > LibDTS_Global.DTS_PRECISION)
End Function

' Logic: Get Midpoint
Public Function GetMidPoint() As clsDTSPoint
    Dim midPt As New clsDTSPoint
    midPt.Init (p_StartPoint.X + p_EndPoint.X) / 2, _
               (p_StartPoint.Y + p_EndPoint.Y) / 2, _
               (p_StartPoint.Z + p_EndPoint.Z) / 2
    Set GetMidPoint = midPt
End Function

' --- SERIALIZATION ---

Public Function ToJson() As String
    Dim root As Object
    ' 1. Get Base Data
    Set root = Base.SerializeBase()
    
    ' 2. Add Frame Data
    root.Add "FT", p_FrameType
    
    ' 3. Geometry (Array format for compactness)
    Dim geo As Object
    Set geo = CreateObject("Scripting.Dictionary")
    geo.Add "P1", p_StartPoint.ToArray()
    geo.Add "P2", p_EndPoint.ToArray()
    root.Add "G", geo
    
    ' 4. Section
    root.Add "S", p_Section.ToDictionary()
    
    ' 5. Encrypt & Return
    Dim jsonStr As String
    jsonStr = LibDTS_Base.ToJson(root)
    ToJson = LibDTS_Security.Encrypt(jsonStr)
End Function

Public Sub FromXData(encryptedStr As String)
    ' 1. Decrypt
    Dim jsonStr As String
    jsonStr = LibDTS_Security.Decrypt(encryptedStr)
    If Len(jsonStr) = 0 Then Exit Sub
    
    ' 2. Parse
    Dim root As Object
    Set root = LibDTS_Base.ParseJson(jsonStr)
    
    ' 3. Hydrate Base
    Base.DeserializeBase root
    
    ' 4. Hydrate Frame Properties
    If root.exists("FT") Then p_FrameType = CLng(root("FT"))
    
    ' 5. Hydrate Geometry
    If root.exists("G") Then
        Dim geo As Object
        Set geo = root("G")
        ' Note: JsonConverter returns Collection for Arrays
        p_StartPoint.Init CDbl(geo("P1")(1)), CDbl(geo("P1")(2)), CDbl(geo("P1")(3))
        p_EndPoint.Init CDbl(geo("P2")(1)), CDbl(geo("P2")(2)), CDbl(geo("P2")(3))
    End If
    
    ' 6. Hydrate Section
    If root.exists("S") Then
        p_Section.FromDictionary root("S")
    End If
End Sub

