VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDTSArea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class Module: clsDTSArea
' Purpose: Represents 2D area elements (Slabs, Walls, Foundations)
' Inherits: clsDTSElement (via composition)
Option Explicit

' --- COMPOSITION ---
Public Base As clsDTSElement

' --- PRIVATE DATA ---
Private p_BoundaryPoints As Collection  ' Collection of clsDTSPoint
Private p_Thickness As Double           ' Element thickness
Private p_Material As String            ' Material name (e.g., "C30")
Private p_AreaType As String            ' "SLAB", "WALL", "FOUNDATION"
Private p_LoadBearing As Boolean        ' Is it load bearing?
Private p_LocalZ As Double              ' Local Z-axis rotation (for walls)
Private p_OpeningGUIDs As Collection    ' Collection of GUID strings for openings

' --- CONSTRUCTOR ---
Private Sub Class_Initialize()
    Set Base = New clsDTSElement
    Base.ElementType = DTS_ELEM_AREA
    
    Set p_BoundaryPoints = New Collection
    Set p_OpeningGUIDs = New Collection
    
    p_Thickness = 200#  ' Default 200mm
    p_Material = "C30"
    p_AreaType = "SLAB"
    p_LoadBearing = True
    p_LocalZ = 0#
End Sub

' --- DESTRUCTOR ---
Private Sub Class_Terminate()
    Set Base = Nothing
    Set p_BoundaryPoints = Nothing
    Set p_OpeningGUIDs = Nothing
End Sub

' --- PROPERTIES ---

Public Property Get BoundaryPoints() As Collection
    Set BoundaryPoints = p_BoundaryPoints
End Property

Public Property Get Thickness() As Double
    Thickness = p_Thickness
End Property
Public Property Let Thickness(Value As Double)
    p_Thickness = Value
    Base.MarkAsDirty
End Property

Public Property Get Material() As String
    Material = p_Material
End Property
Public Property Let Material(Value As String)
    p_Material = Value
    Base.MarkAsDirty
End Property

Public Property Get AreaType() As String
    AreaType = p_AreaType
End Property
Public Property Let AreaType(Value As String)
    p_AreaType = Value
    Base.MarkAsDirty
End Property

Public Property Get LoadBearing() As Boolean
    LoadBearing = p_LoadBearing
End Property
Public Property Let LoadBearing(Value As Boolean)
    p_LoadBearing = Value
    Base.MarkAsDirty
End Property

Public Property Get LocalZ() As Double
    LocalZ = p_LocalZ
End Property
Public Property Let LocalZ(Value As Double)
    p_LocalZ = Value
    Base.MarkAsDirty
End Property

Public Property Get OpeningGUIDs() As Collection
    Set OpeningGUIDs = p_OpeningGUIDs
End Property

' --- BOUNDARY MANAGEMENT ---

' Add a point to the boundary
Public Sub AddBoundaryPoint(pt As clsDTSPoint)
    If Not pt Is Nothing Then
        p_BoundaryPoints.Add pt
        Base.MarkAsDirty
    End If
End Sub

' Clear all boundary points
Public Sub ClearBoundary()
    Set p_BoundaryPoints = New Collection
    Base.MarkAsDirty
End Sub

' Get boundary point by index (1-based)
Public Function GetBoundaryPoint(index As Long) As clsDTSPoint
    If index > 0 And index <= p_BoundaryPoints.Count Then
        Set GetBoundaryPoint = p_BoundaryPoints(index)
    Else
        Set GetBoundaryPoint = Nothing
    End If
End Function

' Get number of boundary points
Public Function BoundaryPointCount() As Long
    BoundaryPointCount = p_BoundaryPoints.Count
End Function

' --- VALIDATION ---

' Check if area is valid (minimum 3 points for a closed boundary)
Public Function IsValid() As Boolean
    IsValid = (p_BoundaryPoints.Count >= 3 And p_Thickness > 0)
End Function

' Check if boundary is closed (first point equals last point)
Public Function IsClosed() As Boolean
    If p_BoundaryPoints.Count < 2 Then
        IsClosed = False
        Exit Function
    End If
    
    Dim firstPt As clsDTSPoint
    Dim lastPt As clsDTSPoint
    Set firstPt = p_BoundaryPoints(1)
    Set lastPt = p_BoundaryPoints(p_BoundaryPoints.Count)
    
    IsClosed = firstPt.IsEqual(lastPt)
End Function

' --- GEOMETRIC CALCULATIONS ---

' Calculate area (using Shoelace formula for 2D polygon)
Public Function CalculateArea() As Double
    On Error GoTo ErrHandler
    
    If p_BoundaryPoints.Count < 3 Then
        CalculateArea = 0#
        Exit Function
    End If
    
    Dim area As Double
    Dim i As Long
    Dim pt1 As clsDTSPoint
    Dim pt2 As clsDTSPoint
    
    area = 0#
    
    ' Shoelace formula: Area = 0.5 * |sum((xi*yi+1) - (xi+1*yi))|
    ' Convert collection to variables for O(1) access
    Dim pt1 As clsDTSPoint
    Dim pt2 As clsDTSPoint
    Dim count As Long
    count = p_BoundaryPoints.count
    
    For i = 1 To count - 1
        Set pt1 = p_BoundaryPoints(i)
        Set pt2 = p_BoundaryPoints(i + 1)
        area = area + (pt1.X * pt2.Y - pt2.X * pt1.Y)
    Next i
    
    ' Close the polygon
    Set pt1 = p_BoundaryPoints(count)
    Set pt2 = p_BoundaryPoints(1)
    area = area + (pt1.X * pt2.Y - pt2.X * pt1.Y)
    
    CalculateArea = Abs(area) / 2#
    Exit Function
    
ErrHandler:
    LibDTS_Logger.Log "clsDTSArea.CalculateArea: Error - " & Err.Description, DTS_ERROR
    CalculateArea = 0#
End Function

' Calculate perimeter
Public Function CalculatePerimeter() As Double
    On Error GoTo ErrHandler
    
    If p_BoundaryPoints.Count < 2 Then
        CalculatePerimeter = 0#
        Exit Function
    End If
    
    Dim perimeter As Double
    Dim i As Long
    Dim pt1 As clsDTSPoint
    Dim pt2 As clsDTSPoint
    
    perimeter = 0#
    
    For i = 1 To p_BoundaryPoints.Count - 1
        Set pt1 = p_BoundaryPoints(i)
        Set pt2 = p_BoundaryPoints(i + 1)
        perimeter = perimeter + pt1.DistanceTo(pt2)
    Next i
    
    ' Close the polygon if needed
    If Not IsClosed() Then
        Set pt1 = p_BoundaryPoints(p_BoundaryPoints.Count)
        Set pt2 = p_BoundaryPoints(1)
        perimeter = perimeter + pt1.DistanceTo(pt2)
    End If
    
    CalculatePerimeter = perimeter
    Exit Function
    
ErrHandler:
    LibDTS_Logger.Log "clsDTSArea.CalculatePerimeter: Error - " & Err.Description, DTS_ERROR
    CalculatePerimeter = 0#
End Function

' Calculate centroid (geometric center)
Public Function GetCentroid() As clsDTSPoint
    On Error GoTo ErrHandler
    
    If p_BoundaryPoints.Count = 0 Then
        Set GetCentroid = Nothing
        Exit Function
    End If
    
    Dim sumX As Double, sumY As Double, sumZ As Double
    Dim pt As clsDTSPoint
    Dim count As Long
    
    sumX = 0#: sumY = 0#: sumZ = 0#
    count = p_BoundaryPoints.Count
    
    For Each pt In p_BoundaryPoints
        sumX = sumX + pt.X
        sumY = sumY + pt.Y
        sumZ = sumZ + pt.Z
    Next pt
    
    Dim centroid As New clsDTSPoint
    centroid.Init sumX / count, sumY / count, sumZ / count
    Set GetCentroid = centroid
    Exit Function
    
ErrHandler:
    LibDTS_Logger.Log "clsDTSArea.GetCentroid: Error - " & Err.Description, DTS_ERROR
    Set GetCentroid = Nothing
End Function

' --- SERIALIZATION ---

' Serialize to JSON (encrypted)
Public Function ToJson() As String
    On Error GoTo ErrHandler
    
    Dim root As Object
    ' 1. Get Base Data
    Set root = Base.SerializeBase()
    
    ' 2. Add Area-specific Data
    root.Add "Thk", p_Thickness
    root.Add "Mat", p_Material
    root.Add "AType", p_AreaType
    root.Add "LdBr", p_LoadBearing
    root.Add "LocZ", p_LocalZ
    
    ' 3. Serialize Boundary Points
    Dim boundaryArr As Object
    Set boundaryArr = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    Dim pt As clsDTSPoint
    i = 0
    For Each pt In p_BoundaryPoints
        boundaryArr.Add "P" & i, pt.ToArray()
        i = i + 1
    Next pt
    root.Add "Bnd", boundaryArr
    
    ' 4. Serialize Opening GUIDs
    If p_OpeningGUIDs.Count > 0 Then
        Dim openingArr As Object
        Set openingArr = CreateObject("Scripting.Dictionary")
        
        Dim j As Long
        Dim guid As Variant
        j = 0
        For Each guid In p_OpeningGUIDs
            openingArr.Add "O" & j, CStr(guid)
            j = j + 1
        Next guid
        root.Add "Opn", openingArr
    End If
    
    ' 5. Convert to JSON and encrypt
    Dim jsonStr As String
    jsonStr = LibDTS_Base.ToJson(root)
    ToJson = LibDTS_Security.Encrypt(jsonStr)
    Exit Function
    
ErrHandler:
    LibDTS_Logger.Log "clsDTSArea.ToJson: Error - " & Err.Description, DTS_ERROR
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
    
    ' 4. Hydrate Area Properties
    If root.exists("Thk") Then p_Thickness = CDbl(root("Thk"))
    If root.exists("Mat") Then p_Material = root("Mat")
    If root.exists("AType") Then p_AreaType = root("AType")
    If root.exists("LdBr") Then p_LoadBearing = CBool(root("LdBr"))
    If root.exists("LocZ") Then p_LocalZ = CDbl(root("LocZ"))
    
    ' 5. Hydrate Boundary Points
    If root.exists("Bnd") Then
        Dim boundaryDict As Object
        Set boundaryDict = root("Bnd")
        
        Set p_BoundaryPoints = New Collection
        
        Dim key As Variant
        For Each key In boundaryDict.Keys
            Dim ptArr As Variant
            ptArr = boundaryDict(key)
            
            Dim pt As New clsDTSPoint
            If IsArray(ptArr) Then
                pt.Init CDbl(ptArr(0)), CDbl(ptArr(1)), CDbl(ptArr(2))
                p_BoundaryPoints.Add pt
            End If
        Next key
    End If
    
    ' 6. Hydrate Opening GUIDs
    If root.exists("Opn") Then
        Dim openingDict As Object
        Set openingDict = root("Opn")
        
        Set p_OpeningGUIDs = New Collection
        
        Dim oKey As Variant
        For Each oKey In openingDict.Keys
            p_OpeningGUIDs.Add openingDict(oKey)
        Next oKey
    End If
    
    Exit Sub
    
ErrHandler:
    LibDTS_Logger.Log "clsDTSArea.FromXData: Error - " & Err.Description, DTS_ERROR
End Sub

' --- HELPER METHODS ---

' Clone this area (creates deep copy with new GUID)
Public Function Clone() As clsDTSArea
    Dim newArea As New clsDTSArea
    
    ' Copy properties
    newArea.Thickness = p_Thickness
    newArea.Material = p_Material
    newArea.AreaType = p_AreaType
    newArea.LoadBearing = p_LoadBearing
    newArea.LocalZ = p_LocalZ
    
    ' Copy boundary points
    Dim pt As clsDTSPoint
    For Each pt In p_BoundaryPoints
        newArea.AddBoundaryPoint pt.Clone()
    Next pt
    
    ' Copy layer
    newArea.Base.Layer = Base.Layer
    
    Set Clone = newArea
End Function

' Add opening by GUID
Public Sub AddOpening(openingGUID As String)
    If Len(Trim$(openingGUID)) > 0 Then
        p_OpeningGUIDs.Add openingGUID
        Base.MarkAsDirty
    End If
End Sub
