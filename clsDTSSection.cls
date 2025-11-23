VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDTSSection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class Module: clsDTSSection
Option Explicit

Private p_Name As String
Private p_Material As String
Private p_Shape As DTSShapeType
Private p_Dims As Object ' Dictionary for W, H, Tf, Tw...

' --- CONSTRUCTOR ---
Private Sub Class_Initialize()
    p_Shape = DTS_SHP_RECTANGLE ' Default
    Set p_Dims = CreateObject("Scripting.Dictionary")
    p_Dims.Add "W", 200#
    p_Dims.Add "H", 300#
End Sub

' --- PROPERTIES ---

Public Property Get Name() As String
    Name = p_Name
End Property
Public Property Let Name(Value As String)
    p_Name = Value
End Property

Public Property Get Width() As Double
    If p_Dims.exists("W") Then Width = p_Dims("W") Else Width = 0
End Property
Public Property Let Width(Value As Double)
    p_Dims("W") = Value
End Property

Public Property Get Depth() As Double
    If p_Dims.exists("H") Then Depth = p_Dims("H") Else Depth = 0
End Property
Public Property Let Depth(Value As Double)
    p_Dims("H") = Value
End Property

' --- METHODS ---

' Serialize to Dictionary
Public Function ToDictionary() As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    dict.Add "N", p_Name
    dict.Add "M", p_Material
    dict.Add "S", p_Shape
    
    ' Flatten dimensions
    Dim k As Variant
    For Each k In p_Dims.keys
        dict.Add k, p_Dims(k)
    Next k
    
    Set ToDictionary = dict
End Function

' Deserialize
Public Sub FromDictionary(dict As Object)
    If dict Is Nothing Then Exit Sub
    
    If dict.exists("N") Then p_Name = dict("N")
    If dict.exists("M") Then p_Material = dict("M")
    If dict.exists("S") Then p_Shape = CLng(dict("S"))
    
    ' Restore dims
    If dict.exists("W") Then p_Dims("W") = CDbl(dict("W"))
    If dict.exists("H") Then p_Dims("H") = CDbl(dict("H"))
    If dict.exists("Tf") Then p_Dims("Tf") = CDbl(dict("Tf"))
End Sub
