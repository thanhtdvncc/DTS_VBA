VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDTSRebarSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class Module: clsDTSRebarSet
Option Explicit

' Inherits Base (Because a Set is also an Element on CAD)
Public Base As clsDTSElement

' The definition of the single bar
' Dinh nghia thanh thep don le
Public BarDefinition As clsDTSRebar

' Distribution Geometry
' Hinh hoc phan bo
Public StartPoint As clsDTSPoint
Public EndPoint As clsDTSPoint
Public DistributionType As String ' "UNIFORM", "VARIABLE"

Private Sub Class_Initialize()
    Set Base = New clsDTSElement
    Base.ElementType = "REBAR_SET"
    
    Set BarDefinition = New clsDTSRebar
    Set StartPoint = New clsDTSPoint
    Set EndPoint = New clsDTSPoint
End Sub

' Logic: Auto-calculate quantity based on distance
' Logic: Tu dong tinh so luong dua tren khoang cach
Public Sub RecalculateQuantity()
    Dim dist As Double
    dist = StartPoint.DistanceTo(EndPoint)
    
    If BarDefinition.Spacing > 0 Then
        ' Round up: (L / a) + 1
        BarDefinition.Quantity = Int(dist / BarDefinition.Spacing) + 1
    End If
End Sub

' Serialization (Similar pattern)
Public Function Serialize() As String
    Dim root As Object
    Set root = Base.SerializeBase()
    
    ' Embed the Rebar Definition fully
    ' Nhung toan bo dinh nghia thanh thep
    root.Add "BarDef", BarDefinition.ToJson() ' Note: This returns encrypted string, might need raw dict
    
    ' For cleaner nesting, better to use ToDictionary method on Rebar instead of ToJson
    ' But for simplicity now, we store it as a string property
    
    Dim geo As Object
    Set geo = CreateObject("Scripting.Dictionary")
    geo.Add "P1", StartPoint.ToArray()
    geo.Add "P2", EndPoint.ToArray()
    root.Add "DistGeo", geo
    
    Dim jsonStr As String
    jsonStr = LibDTS_Base.ToJson(root)
    Serialize = LibDTS_Security.Encrypt(jsonStr)
End Function
