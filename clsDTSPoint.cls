VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDTSPoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class Module: clsDTSPoint
Option Explicit

' Private storage
Private p_X As Double
Private p_Y As Double
Private p_Z As Double

' --- PROPERTIES ---

Public Property Get X() As Double
    X = p_X
End Property
Public Property Let X(Value As Double)
    p_X = Value
End Property

Public Property Get Y() As Double
    Y = p_Y
End Property
Public Property Let Y(Value As Double)
    p_Y = Value
End Property

Public Property Get Z() As Double
    Z = p_Z
End Property
Public Property Let Z(Value As Double)
    p_Z = Value
End Property

' --- INITIALIZATION ---

Public Sub Init(xVal As Double, yVal As Double, Optional zVal As Double = 0)
    p_X = xVal
    p_Y = yVal
    p_Z = zVal
End Sub

' --- VECTOR OPERATIONS (Math Logic) ---

' Returns a new point derived from offset
' Tra ve diem moi duoc dich chuyen (Khong thay doi diem goc)
Public Function Offset(dx As Double, dy As Double, dz As Double) As clsDTSPoint
    Dim p As New clsDTSPoint
    p.Init p_X + dx, p_Y + dy, p_Z + dz
    Set Offset = p
End Function

' Calculate Euclidean Distance
' Tinh khoang cach Euclid
Public Function DistanceTo(other As clsDTSPoint) As Double
    If other Is Nothing Then Exit Function
    DistanceTo = Sqr((p_X - other.X) ^ 2 + (p_Y - other.Y) ^ 2 + (p_Z - other.Z) ^ 2)
End Function

' Check equality with tolerance
' Kiem tra trung nhau voi sai so
Public Function IsEqual(other As clsDTSPoint) As Boolean
    If other Is Nothing Then Exit Function
    IsEqual = (DistanceTo(other) < LibDTS_Global.DTS_PRECISION)
End Function

' Create a clone (Deep Copy)
Public Function Clone() As clsDTSPoint
    Dim p As New clsDTSPoint
    p.Init p_X, p_Y, p_Z
    Set Clone = p
End Function

' Convert to simple array (for JSON serialization)
Public Function ToArray() As Variant
    ToArray = Array(p_X, p_Y, p_Z)
End Function

' Format string for Debug
Public Function ToString() As String
    ToString = String(2, " ") & Format(p_X, "0.0") & ", " & Format(p_Y, "0.0") & ", " & Format(p_Z, "0.0")
End Function
