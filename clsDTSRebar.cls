VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDTSRebar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class Module: clsDTSRebar
Option Explicit

' --- COMPOSITION ---
Public Base As clsDTSElement

' --- REBAR PROPERTIES ---
Public Diameter As Double
Public Quantity As Long
Public Spacing As Double        ' a200
Public Mark As String           ' "1", "2"
Public RebarClass As String     ' "CB400V", "SD390"
Public HostGUID As String       ' Link to Parent Beam/Column

' --- SUB-OBJECTS ---
Public Shape As clsDTSRebarShape

' --- CONSTRUCTOR ---
Private Sub Class_Initialize()
    Set Base = New clsDTSElement
    Base.ElementType = DTS_ELEM_REBAR
    Set Shape = New clsDTSRebarShape
End Sub

' --- LOGIC METHODS ---

' Calculate Weight (kg)
' Tinh trong luong
Public Function GetTotalWeight() As Double
    Dim lenMM As Double
    lenMM = Shape.CalculateTotalLength(Me.Diameter)
    
    ' Weight = (D^2 / 162) * L(m) * Qty
    GetTotalWeight = (Diameter * Diameter / 162) * (lenMM / 1000) * Quantity
End Function

' --- SERIALIZATION ---

Public Function ToJson() As String
    Dim root As Object
    ' 1. Base
    Set root = Base.SerializeBase()
    
    ' 2. Props
    root.Add "Dia", Diameter
    root.Add "Qty", Quantity
    root.Add "Spc", Spacing
    root.Add "Mrk", Mark
    root.Add "Mat", RebarClass
    root.Add "Host", HostGUID
    
    ' 3. Shape (Embedded Object)
    root.Add "Shp", Shape.ToDictionary()
    
    ' 4. Encrypt
    Dim jsonStr As String
    jsonStr = LibDTS_Base.ToJson(root)
    ToJson = LibDTS_Security.Encrypt(jsonStr)
End Function

Public Sub FromXData(encryptedStr As String)
    Dim jsonStr As String
    jsonStr = LibDTS_Security.Decrypt(encryptedStr)
    If Len(jsonStr) = 0 Then Exit Sub
    
    Dim root As Object
    Set root = LibDTS_Base.ParseJson(jsonStr)
    
    ' Restore Base
    Base.DeserializeBase root
    
    ' Restore Props
    If root.exists("Dia") Then Diameter = CDbl(root("Dia"))
    If root.exists("Qty") Then Quantity = CLng(root("Qty"))
    If root.exists("Spc") Then Spacing = CDbl(root("Spc"))
    If root.exists("Mrk") Then Mark = root("Mrk")
    If root.exists("Mat") Then RebarClass = root("Mat")
    If root.exists("Host") Then HostGUID = root("Host")
    
    ' Restore Shape
    If root.exists("Shp") Then
        Shape.FromDictionary root("Shp")
    End If
End Sub
