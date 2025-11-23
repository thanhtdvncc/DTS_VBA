VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDTSRebarShape"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class Module: clsDTSRebarShape
Option Explicit

' --- PROPERTIES ---
Private p_ShapeCode As String    ' e.g., "00", "18", "51"
Private p_Dimensions As Object   ' Dictionary(Index -> Value)
Private p_TotalLength As Double
Private p_HookLength As Double

' --- CONSTRUCTOR ---
Private Sub Class_Initialize()
    p_ShapeCode = "00"
    Set p_Dimensions = CreateObject("Scripting.Dictionary")
End Sub

' --- PUBLIC ACCESSORS ---
Public Property Get Code() As String
    Code = p_ShapeCode
End Property
Public Property Let Code(Value As String)
    p_ShapeCode = Value
End Property

' Set dimension by index (L1, L2, L3...)
' Gan kich thuoc theo chi so
Public Sub SetDim(idx As Integer, val As Double)
    If p_Dimensions.exists(idx) Then
        p_Dimensions(idx) = val
    Else
        p_Dimensions.Add idx, val
    End If
End Sub

Public Function GetDim(idx As Integer) As Double
    If p_Dimensions.exists(idx) Then
        GetDim = p_Dimensions(idx)
    Else
        GetDim = 0
    End If
End Function

' --- GEOMETRY LOGIC (THE "BRAIN") ---

' Calculate cut length based on shape rules
' Tinh chieu dai cat dua tren quy tac hinh dang
Public Function CalculateTotalLength(Diameter As Double) As Double
    Dim total As Double: total = 0
    Dim k As Variant
    
    ' Sum all segments
    ' Cong tong cac doan
    For Each k In p_Dimensions.keys
        total = total + p_Dimensions(k)
    Next k
    
    ' Add Hooks based on ShapeCode
    ' Cong them moc dua vao Ma hinh
    Select Case p_ShapeCode
        Case "18" ' U-Shape with hooks
            total = total + (2 * GetHookLength(Diameter, 180))
        Case "02" ' Straight with hooks
            total = total + (2 * GetHookLength(Diameter, 180))
        Case "01" ' Straight no hook
            ' No add
    End Select
    
    ' Subtract bending deduction (Optional advanced logic)
    ' Tru di do gian dai do uon (Logic nang cao tuy chon)
    ' total = total - GetBendingDeduction(diameter)
    
    p_TotalLength = total
    CalculateTotalLength = total
End Function

' Helper: Standard hook length
' Ham phu: Chieu dai moc tieu chuan
Private Function GetHookLength(dia As Double, angle As Double) As Double
    ' Example rule: 6.25d for 180 deg, but min 50mm
    Dim lenCalc As Double
    If angle = 180 Then
        lenCalc = 6.25 * dia
        If lenCalc < 50 Then lenCalc = 50
    ElseIf angle = 90 Then
        lenCalc = 12 * dia
    End If
    GetHookLength = lenCalc
End Function

' Serialize Shape Data
Public Function ToDictionary() As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.Add "C", p_ShapeCode
    
    ' Serialize dims as array for compactness
    ' Dong goi kich thuoc thanh mang de tiet kiem cho
    Dim dimArr() As Variant
    ReDim dimArr(0 To p_Dimensions.count - 1)
    Dim i As Integer: i = 0
    Dim k As Variant
    For Each k In p_Dimensions.keys
        ' Store as "1:200" (Index:Value) string to preserve order
        dimArr(i) = k & ":" & p_Dimensions(k)
        i = i + 1
    Next k
    dict.Add "D", dimArr
    
    Set ToDictionary = dict
End Function

Public Sub FromDictionary(dict As Object)
    If dict Is Nothing Then Exit Sub
    If dict.exists("C") Then p_ShapeCode = dict("C")
    
    If dict.exists("D") Then
        Dim arr As Variant
        arr = dict("D") ' Collection from JSON
        Dim item As Variant
        For Each item In arr
            Dim parts() As String
            parts = Split(CStr(item), ":")
            If UBound(parts) = 1 Then
                Me.SetDim CInt(parts(0)), CDbl(parts(1))
            End If
        Next item
    End If
End Sub
