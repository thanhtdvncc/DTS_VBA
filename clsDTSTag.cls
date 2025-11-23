VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDTSTag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class Module: clsDTSTag
' Purpose: Annotation/Tag element that can be attached to structural elements
' Inherits: clsDTSElement (via composition)
Option Explicit

' --- COMPOSITION ---
Public Base As clsDTSElement

' --- PRIVATE DATA ---
Private p_HostGUID As String        ' GUID of parent element (Beam, Column, etc.)
Private p_TextContent As String     ' Annotation text
Private p_Position As clsDTSPoint   ' Position in 3D space
Private p_Rotation As Double        ' Text rotation in radians
Private p_Height As Double          ' Text height
Private p_Style As String           ' Text style name

' --- CONSTRUCTOR ---
Private Sub Class_Initialize()
    Set Base = New clsDTSElement
    Base.ElementType = DTS_ELEM_ANNOTATION
    
    Set p_Position = New clsDTSPoint
    p_Position.Init 0, 0, 0
    
    p_TextContent = ""
    p_HostGUID = ""
    p_Rotation = 0#
    p_Height = 2.5 ' Default text height
    p_Style = "Standard"
End Sub

' --- DESTRUCTOR ---
Private Sub Class_Terminate()
    Set Base = Nothing
    Set p_Position = Nothing
End Sub

' --- PROPERTIES ---

Public Property Get HostGUID() As String
    HostGUID = p_HostGUID
End Property
Public Property Let HostGUID(Value As String)
    p_HostGUID = Value
    Base.MarkAsDirty
End Property

Public Property Get TextContent() As String
    TextContent = p_TextContent
End Property
Public Property Let TextContent(Value As String)
    p_TextContent = Value
    Base.MarkAsDirty
End Property

Public Property Get Position() As clsDTSPoint
    Set Position = p_Position
End Property
Public Property Set Position(Value As clsDTSPoint)
    Set p_Position = Value
    Base.MarkAsDirty
End Property

Public Property Get Rotation() As Double
    Rotation = p_Rotation
End Property
Public Property Let Rotation(Value As Double)
    p_Rotation = Value
    Base.MarkAsDirty
End Property

Public Property Get Height() As Double
    Height = p_Height
End Property
Public Property Let Height(Value As Double)
    p_Height = Value
    Base.MarkAsDirty
End Property

Public Property Get Style() As String
    Style = p_Style
End Property
Public Property Let Style(Value As String)
    p_Style = Value
    Base.MarkAsDirty
End Property

' --- VALIDATION ---

' Check if tag is valid
Public Function IsValid() As Boolean
    IsValid = (Len(Trim$(p_TextContent)) > 0 And Not p_Position Is Nothing)
End Function

' Check if tag is attached to a host
Public Function HasHost() As Boolean
    HasHost = (Len(Trim$(p_HostGUID)) > 0)
End Function

' --- SERIALIZATION ---

' Serialize to JSON (encrypted)
Public Function ToJson() As String
    On Error GoTo ErrHandler
    
    Dim root As Object
    ' 1. Get Base Data
    Set root = Base.SerializeBase()
    
    ' 2. Add Tag-specific Data
    root.Add "Host", p_HostGUID
    root.Add "Text", p_TextContent
    root.Add "Rot", p_Rotation
    root.Add "Hgt", p_Height
    root.Add "Sty", p_Style
    
    ' 3. Add Position
    root.Add "Pos", p_Position.ToArray()
    
    ' 4. Convert to JSON and encrypt
    Dim jsonStr As String
    jsonStr = LibDTS_Base.ToJson(root)
    ToJson = LibDTS_Security.Encrypt(jsonStr)
    Exit Function
    
ErrHandler:
    LibDTS_Logger.Log "clsDTSTag.ToJson: Error - " & Err.Description, DTS_ERROR
    ToJson = ""
End Function

' Deserialize from encrypted JSON
Public Sub FromXData(encryptedStr As String)
    On Error GoTo ErrHandler
    
    ' 1. Decrypt
    Dim jsonStr As String
    jsonStr = LibDTS_Security.Decrypt(encryptedStr)
    If Len(jsonStr) = 0 Then Exit Sub
    
    ' 2. Parse JSON
    Dim root As Object
    Set root = LibDTS_Base.ParseJson(jsonStr)
    
    ' 3. Hydrate Base
    Base.DeserializeBase root
    
    ' 4. Hydrate Tag Properties
    If root.exists("Host") Then p_HostGUID = root("Host")
    If root.exists("Text") Then p_TextContent = root("Text")
    If root.exists("Rot") Then p_Rotation = CDbl(root("Rot"))
    If root.exists("Hgt") Then p_Height = CDbl(root("Hgt"))
    If root.exists("Sty") Then p_Style = root("Sty")
    
    ' 5. Hydrate Position
    If root.exists("Pos") Then
        Dim posArr As Variant
        posArr = root("Pos")
        If IsArray(posArr) Then
            p_Position.Init CDbl(posArr(0)), CDbl(posArr(1)), CDbl(posArr(2))
        End If
    End If
    
    Exit Sub
    
ErrHandler:
    LibDTS_Logger.Log "clsDTSTag.FromXData: Error - " & Err.Description, DTS_ERROR
End Sub

' --- HELPER METHODS ---

' Get display text (can be overridden for formatted output)
Public Function GetDisplayText() As String
    GetDisplayText = p_TextContent
End Function

' Clone this tag (creates deep copy with new GUID)
Public Function Clone() As clsDTSTag
    Dim newTag As New clsDTSTag
    
    ' Copy properties (new GUID generated automatically)
    newTag.HostGUID = p_HostGUID
    newTag.TextContent = p_TextContent
    newTag.Rotation = p_Rotation
    newTag.Height = p_Height
    newTag.Style = p_Style
    Set newTag.Position = p_Position.Clone()
    
    ' Copy layer
    newTag.Base.Layer = Base.Layer
    
    Set Clone = newTag
End Function
