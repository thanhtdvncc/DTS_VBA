VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDTSElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class Module: clsDTSElement
Option Explicit

' Core Data
Private p_GUID As String
Private p_OwnerHandle As String
Private p_ElementType As DTSElementType
Private p_Layer As String
Private p_IsDirty As Boolean ' Track if object needs saving

' --- CONSTRUCTOR ---
Private Sub Class_Initialize()
    p_GUID = LibDTS_Base.GenerateGUID()
    p_ElementType = DTS_ELEM_UNKNOWN
    p_Layer = "0"
    p_IsDirty = True ' New objects are always dirty
End Sub

' --- PROPERTIES ---

Public Property Get guid() As String
    guid = p_GUID
End Property
' GUID is Read-Only externally, except via specific Load methods

Public Property Get OwnerHandle() As String
    OwnerHandle = p_OwnerHandle
End Property
Public Property Let OwnerHandle(Value As String)
    p_OwnerHandle = Value
End Property

Public Property Get ElementType() As DTSElementType
    ElementType = p_ElementType
End Property
Public Property Let ElementType(Value As DTSElementType)
    p_ElementType = Value
End Property

Public Property Get IsDirty() As Boolean
    IsDirty = p_IsDirty
End Property
Public Sub MarkAsClean()
    p_IsDirty = False
End Sub

' --- CORE LOGIC ---

' Self-Healing Identity Check
' Kiem tra va tu sua chua dinh danh
Public Function ValidateIdentity(cadHandle As String) As Boolean
    If p_OwnerHandle = "" Then
        p_OwnerHandle = cadHandle
        ValidateIdentity = True
        Exit Function
    End If
    
    If p_OwnerHandle <> cadHandle Then
        ' Detected Copy -> Generate New GUID
        p_GUID = LibDTS_Base.GenerateGUID()
        p_OwnerHandle = cadHandle
        p_IsDirty = True
        ValidateIdentity = True ' Identity changed
    Else
        ValidateIdentity = False
    End If
End Function

' Base Serialize - Returns Dictionary
Public Function SerializeBase() As Object
    Dim root As Object
    Set root = CreateObject("Scripting.Dictionary")
    root.Add "ID", p_GUID
    root.Add "Hdl", p_OwnerHandle
    root.Add "Type", p_ElementType
    root.Add "Lay", p_Layer
    Set SerializeBase = root
End Function

' Base Deserialize
Public Sub DeserializeBase(dict As Object)
    If dict Is Nothing Then Exit Sub
    If dict.exists("ID") Then p_GUID = dict("ID")
    If dict.exists("Hdl") Then p_OwnerHandle = dict("Hdl")
    If dict.exists("Type") Then p_ElementType = CLng(dict("Type"))
    If dict.exists("Lay") Then p_Layer = dict("Lay")
    p_IsDirty = False
End Sub
