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
Private p_Properties As Object ' Dictionary for custom properties
Private p_IsDirty As Boolean ' Track if object needs saving

' --- CONSTRUCTOR ---
Private Sub Class_Initialize()
    p_GUID = LibDTS_Base.GenerateGUID()
    p_ElementType = DTS_ELEM_UNKNOWN
    p_Layer = "0"
    Set p_Properties = CreateObject("Scripting.Dictionary")
    p_IsDirty = True ' New objects are always dirty
End Sub

' --- DESTRUCTOR ---
Private Sub Class_Terminate()
    Set p_Properties = Nothing
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

Public Property Get Layer() As String
    Layer = p_Layer
End Property
Public Property Let Layer(Value As String)
    p_Layer = Value
    p_IsDirty = True
End Property

Public Property Get Properties() As Object
    Set Properties = p_Properties
End Property

Public Property Get IsDirty() As Boolean
    IsDirty = p_IsDirty
End Property
Public Sub MarkAsClean()
    p_IsDirty = False
End Sub
Public Sub MarkAsDirty()
    p_IsDirty = True
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
    
    ' Include custom properties if any exist
    If Not p_Properties Is Nothing Then
        If p_Properties.Count > 0 Then
            root.Add "Props", p_Properties
        End If
    End If
    
    Set SerializeBase = root
End Function

' Base Deserialize
Public Sub DeserializeBase(dict As Object)
    If dict Is Nothing Then Exit Sub
    If dict.exists("ID") Then p_GUID = dict("ID")
    If dict.exists("Hdl") Then p_OwnerHandle = dict("Hdl")
    If dict.exists("Type") Then p_ElementType = CLng(dict("Type"))
    If dict.exists("Lay") Then p_Layer = dict("Lay")
    
    ' Restore custom properties
    If dict.exists("Props") Then
        Dim propKey As Variant
        For Each propKey In dict("Props").Keys
            p_Properties(propKey) = dict("Props")(propKey)
        Next propKey
    End If
    
    p_IsDirty = False
End Sub
