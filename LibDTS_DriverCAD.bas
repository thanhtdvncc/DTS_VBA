Attribute VB_Name = "LibDTS_DriverCAD"
' Module: LibDTS_DriverCAD
' Purpose: Canonical driver for AutoCAD integration
' Features: Drawing, reading, XData operations, GUID mapping, dry-run support
' Version: 2.0 (Enhanced with defensive programming and full API surface)
Option Explicit

' ==========================================
' MODULE-LEVEL VARIABLES
' ==========================================
Private m_LastError As String        ' Last error message
Private m_EntityCache As Object      ' Dictionary: GUID -> Entity Handle
Private m_AppName As String          ' XData application name

' ==========================================
' PUBLIC CONSTANTS
' ==========================================
Public Const DRIVER_NAME As String = "LibDTS_DriverCAD"
Public Const DEFAULT_APP_NAME As String = "DTS_CORE"

' ==========================================
' MODULE INITIALIZATION
' ==========================================
Private Sub Class_Initialize()
    Set m_EntityCache = CreateObject("Scripting.Dictionary")
    m_AppName = DEFAULT_APP_NAME
End Sub

' ==========================================
' 1. DRAWING OPERATIONS (with Dry-Run Support)
' ==========================================

' Draw a point (node/joint) in AutoCAD
' Parameters:
'   pt: clsDTSPoint object with coordinates
'   acadDoc: AutoCAD Document object
'   dryRun: False (commit), True (validate only)
' Returns: AutoCAD Circle object or Nothing on failure
Public Function DrawPoint(pt As clsDTSPoint, _
                         acadDoc As Object, _
                         Optional dryRun As Boolean = False) As Object
    On Error GoTo ErrHandler
    
    ' Validate inputs
    If pt Is Nothing Then
        m_LastError = "Point object is Nothing"
        LibDTS_Logger.Log DRIVER_NAME & ".DrawPoint: " & m_LastError, DTS_ERROR
        Set DrawPoint = Nothing
        Exit Function
    End If
    
    If acadDoc Is Nothing Then
        m_LastError = "AutoCAD Document object is Nothing"
        LibDTS_Logger.Log DRIVER_NAME & ".DrawPoint: " & m_LastError, DTS_ERROR
        Set DrawPoint = Nothing
        Exit Function
    End If
    
    ' Ensure GUID exists (Points don't have Base, so skip this)
    
    ' Dry-run mode
    If dryRun Then
        LibDTS_Logger.Log DRIVER_NAME & ".DrawPoint [DRY-RUN]: Would draw point at (" & pt.X & ", " & pt.Y & ", " & pt.Z & ")", DTS_INFO
        Set DrawPoint = Nothing
        Exit Function
    End If
    
    ' Get ModelSpace
    Dim ms As Object
    Set ms = acadDoc.ModelSpace
    
    ' Create point representation (circle with small radius)
    Dim center(0 To 2) As Double
    center(0) = pt.X
    center(1) = pt.Y
    center(2) = pt.Z
    
    Dim circleObj As Object
    Set circleObj = ms.AddCircle(center, 10#) ' 10mm radius
    
    ' Set properties (no Base for Point)
    On Error Resume Next
    ' Points are simple geometry objects without Base
    On Error GoTo ErrHandler
    
    ' Note: SaveXData for Point would need special handling
    ' SaveXData pt, circleObj
    
    ' Cache mapping (Points don't have GUID, use handle)
    ' Points are basic geometry - no caching needed
    
    LibDTS_Logger.Log DRIVER_NAME & ".DrawPoint: Drew point at (" & pt.X & ", " & pt.Y & ", " & pt.Z & ")", DTS_INFO
    Set DrawPoint = circleObj
    Exit Function
    
ErrHandler:
    m_LastError = "DrawPoint error: " & err.description
    LibDTS_Logger.Log DRIVER_NAME & ".DrawPoint: " & m_LastError, DTS_ERROR
    Set DrawPoint = Nothing
End Function

' Draw a frame (beam/column) in AutoCAD
' Parameters:
'   frame: clsDTSFrame object with geometry
'   acadDoc: AutoCAD Document object
'   dryRun: False (commit), True (validate only)
' Returns: AutoCAD Line object or Nothing on failure
Public Function DrawFrame(frame As clsDTSFrame, _
                         acadDoc As Object, _
                         Optional dryRun As Boolean = False) As Object
    On Error GoTo ErrHandler
    
    ' Validate inputs
    If frame Is Nothing Then
        m_LastError = "Frame object is Nothing"
        LibDTS_Logger.Log DRIVER_NAME & ".DrawFrame: " & m_LastError, DTS_ERROR
        Set DrawFrame = Nothing
        Exit Function
    End If
    
    If acadDoc Is Nothing Then
        m_LastError = "AutoCAD Document object is Nothing"
        LibDTS_Logger.Log DRIVER_NAME & ".DrawFrame: " & m_LastError, DTS_ERROR
        Set DrawFrame = Nothing
        Exit Function
    End If
    
    ' Ensure GUID exists
    If Len(Trim$(frame.Base.guid)) = 0 Then
        frame.Base.guid = LibDTS_Base.GenerateGUID()
    End If
    
    ' Dry-run mode
    If dryRun Then
        LibDTS_Logger.Log DRIVER_NAME & ".DrawFrame [DRY-RUN]: Would draw frame from (" & frame.StartPoint.X & "," & frame.StartPoint.Y & "," & frame.StartPoint.Z & ") to (" & frame.EndPoint.X & "," & frame.EndPoint.Y & "," & frame.EndPoint.Z & ")", DTS_INFO
        Set DrawFrame = Nothing
        Exit Function
    End If
    
    ' Get ModelSpace
    Dim ms As Object
    Set ms = acadDoc.ModelSpace
    
    ' Create line
    Dim p1(0 To 2) As Double, p2(0 To 2) As Double
    p1(0) = frame.StartPoint.X: p1(1) = frame.StartPoint.Y: p1(2) = frame.StartPoint.Z
    p2(0) = frame.EndPoint.X: p2(1) = frame.EndPoint.Y: p2(2) = frame.EndPoint.Z
    
    Dim lineObj As Object
    Set lineObj = ms.AddLine(p1, p2)
    
    ' Set properties
    On Error Resume Next
    If Len(frame.Base.layer) > 0 Then lineObj.layer = frame.Base.layer
    If frame.Base.color <> 0 Then lineObj.color = frame.Base.color
    On Error GoTo ErrHandler
    
    ' Save XData
    SaveXData frame, lineObj
    
    ' Cache mapping
    If m_EntityCache.Exists(frame.Base.guid) Then
        m_EntityCache(frame.Base.guid) = lineObj.Handle
    Else
        m_EntityCache.Add frame.Base.guid, lineObj.Handle
    End If
    
    LibDTS_Logger.Log DRIVER_NAME & ".DrawFrame: Drew frame " & frame.Base.guid, DTS_INFO
    Set DrawFrame = lineObj
    Exit Function
    
ErrHandler:
    m_LastError = "DrawFrame error: " & err.description
    LibDTS_Logger.Log DRIVER_NAME & ".DrawFrame: " & m_LastError, DTS_ERROR
    Set DrawFrame = Nothing
End Function

' Draw an area (slab/wall) in AutoCAD
' Parameters:
'   area: clsDTSArea object with boundary points
'   acadDoc: AutoCAD Document object
'   dryRun: False (commit), True (validate only)
' Returns: AutoCAD Polyline object or Nothing on failure
Public Function DrawArea(area As clsDTSArea, _
                        acadDoc As Object, _
                        Optional dryRun As Boolean = False) As Object
    On Error GoTo ErrHandler
    
    ' Validate inputs
    If area Is Nothing Then
        m_LastError = "Area object is Nothing"
        LibDTS_Logger.Log DRIVER_NAME & ".DrawArea: " & m_LastError, DTS_ERROR
        Set DrawArea = Nothing
        Exit Function
    End If
    
    If acadDoc Is Nothing Then
        m_LastError = "AutoCAD Document object is Nothing"
        LibDTS_Logger.Log DRIVER_NAME & ".DrawArea: " & m_LastError, DTS_ERROR
        Set DrawArea = Nothing
        Exit Function
    End If
    
    If area.BoundaryPoints.count < 3 Then
        m_LastError = "Area must have at least 3 boundary points"
        LibDTS_Logger.Log DRIVER_NAME & ".DrawArea: " & m_LastError, DTS_ERROR
        Set DrawArea = Nothing
        Exit Function
    End If
    
    ' Ensure GUID exists
    If Len(Trim$(area.Base.guid)) = 0 Then
        area.Base.guid = LibDTS_Base.GenerateGUID()
    End If
    
    ' Dry-run mode
    If dryRun Then
        LibDTS_Logger.Log DRIVER_NAME & ".DrawArea [DRY-RUN]: Would draw area with " & area.BoundaryPoints.count & " points", DTS_INFO
        Set DrawArea = Nothing
        Exit Function
    End If
    
    ' Get ModelSpace
    Dim ms As Object
    Set ms = acadDoc.ModelSpace
    
    ' Build point array
    Dim numPoints As Long
    numPoints = area.BoundaryPoints.count
    
    Dim points() As Double
    ReDim points(0 To (numPoints * 3) - 1) ' 3 coordinates per point
    
    Dim i As Long
    Dim pt As clsDTSPoint
    For i = 1 To numPoints
        Set pt = area.BoundaryPoints(i)
        points((i - 1) * 3 + 0) = pt.X
        points((i - 1) * 3 + 1) = pt.Y
        points((i - 1) * 3 + 2) = pt.Z
    Next i
    
    ' Create 3D polyline
    Dim polyObj As Object
    Set polyObj = ms.Add3DPoly(points)
    polyObj.Closed = True
    
    ' Set properties
    On Error Resume Next
    If Len(area.Base.layer) > 0 Then polyObj.layer = area.Base.layer
    If area.Base.color <> 0 Then polyObj.color = area.Base.color
    On Error GoTo ErrHandler
    
    ' Save XData
    SaveXData area, polyObj
    
    ' Cache mapping
    If m_EntityCache.Exists(area.Base.guid) Then
        m_EntityCache(area.Base.guid) = polyObj.Handle
    Else
        m_EntityCache.Add area.Base.guid, polyObj.Handle
    End If
    
    LibDTS_Logger.Log DRIVER_NAME & ".DrawArea: Drew area " & area.Base.guid & " with " & numPoints & " points", DTS_INFO
    Set DrawArea = polyObj
    Exit Function
    
ErrHandler:
    m_LastError = "DrawArea error: " & err.description
    LibDTS_Logger.Log DRIVER_NAME & ".DrawArea: " & m_LastError, DTS_ERROR
    Set DrawArea = Nothing
End Function

' Draw a tag (text annotation) in AutoCAD
' Parameters:
'   tag: clsDTSTag object
'   acadDoc: AutoCAD Document object
'   dryRun: False (commit), True (validate only)
' Returns: AutoCAD Text object or Nothing on failure
Public Function DrawTag(tag As clsDTSTag, _
                       acadDoc As Object, _
                       Optional dryRun As Boolean = False) As Object
    On Error GoTo ErrHandler
    
    ' Validate inputs
    If tag Is Nothing Then
        m_LastError = "Tag object is Nothing"
        LibDTS_Logger.Log DRIVER_NAME & ".DrawTag: " & m_LastError, DTS_ERROR
        Set DrawTag = Nothing
        Exit Function
    End If
    
    If acadDoc Is Nothing Then
        m_LastError = "AutoCAD Document object is Nothing"
        LibDTS_Logger.Log DRIVER_NAME & ".DrawTag: " & m_LastError, DTS_ERROR
        Set DrawTag = Nothing
        Exit Function
    End If
    
    ' Ensure GUID exists
    If Len(Trim$(tag.Base.guid)) = 0 Then
        tag.Base.guid = LibDTS_Base.GenerateGUID()
    End If
    
    ' Dry-run mode
    If dryRun Then
        LibDTS_Logger.Log DRIVER_NAME & ".DrawTag [DRY-RUN]: Would draw text '" & tag.TextContent & "' at (" & tag.Position.X & "," & tag.Position.Y & "," & tag.Position.Z & ")", DTS_INFO
        Set DrawTag = Nothing
        Exit Function
    End If
    
    ' Validate text height
    If tag.Height <= 0 Then
        m_LastError = "Tag height must be positive (got: " & tag.Height & ")"
        LibDTS_Logger.Log DRIVER_NAME & ".DrawTag: " & m_LastError, DTS_ERROR
        Set DrawTag = Nothing
        Exit Function
    End If
    
    ' Get ModelSpace
    Dim ms As Object
    Set ms = acadDoc.ModelSpace
    
    ' Create text
    Dim insPt(0 To 2) As Double
    insPt(0) = tag.Position.X
    insPt(1) = tag.Position.Y
    insPt(2) = tag.Position.Z
    
    Dim textObj As Object
    Set textObj = ms.AddText(tag.TextContent, insPt, tag.Height)
    
    ' Set properties
    On Error Resume Next
    textObj.Rotation = tag.Rotation
    If Len(tag.Base.layer) > 0 Then textObj.layer = tag.Base.layer
    On Error GoTo ErrHandler
    
    ' Save XData
    SaveXData tag, textObj
    
    ' Cache mapping
    If m_EntityCache.Exists(tag.Base.guid) Then
        m_EntityCache(tag.Base.guid) = textObj.Handle
    Else
        m_EntityCache.Add tag.Base.guid, textObj.Handle
    End If
    
    LibDTS_Logger.Log DRIVER_NAME & ".DrawTag: Drew tag " & tag.Base.guid, DTS_INFO
    Set DrawTag = textObj
    Exit Function
    
ErrHandler:
    m_LastError = "DrawTag error: " & err.description
    LibDTS_Logger.Log DRIVER_NAME & ".DrawTag: " & m_LastError, DTS_ERROR
    Set DrawTag = Nothing
End Function

' ==========================================
' 2. READING OPERATIONS (Single Entity)
' ==========================================

' Read a point from CAD entity
' Parameters:
'   ent: AutoCAD Circle entity
' Returns: clsDTSPoint object or Nothing on failure
Public Function ReadPoint(ent As Object) As clsDTSPoint
    On Error GoTo ErrHandler
    
    ' Validate input
    If ent Is Nothing Then
        m_LastError = "Entity is Nothing"
        LibDTS_Logger.Log DRIVER_NAME & ".ReadPoint: " & m_LastError, DTS_ERROR
        Set ReadPoint = Nothing
        Exit Function
    End If
    
    ' Check entity type
    Dim entType As String
    entType = LCase$(TypeName(ent))
    If InStr(entType, "circle") = 0 Then
        m_LastError = "Entity is not a circle (got: " & TypeName(ent) & ")"
        LibDTS_Logger.Log DRIVER_NAME & ".ReadPoint: " & m_LastError, DTS_WARNING
        Set ReadPoint = Nothing
        Exit Function
    End If
    
    ' Create point object
    Dim pt As New clsDTSPoint
    
    ' Get geometry from CAD
    Dim center As Variant
    center = ent.center
    pt.Init center(0), center(1), center(2)
    
    ' Note: Points don't have XData/Base in this implementation
    ' They are simple geometry primitives
    
    LibDTS_Logger.Log DRIVER_NAME & ".ReadPoint: Read point at (" & pt.X & ", " & pt.Y & ", " & pt.Z & ")", DTS_INFO
    Set ReadPoint = pt
    Exit Function
    
ErrHandler:
    m_LastError = "ReadPoint error: " & err.description
    LibDTS_Logger.Log DRIVER_NAME & ".ReadPoint: " & m_LastError, DTS_ERROR
    Set ReadPoint = Nothing
End Function

' Read a frame from CAD entity
' Parameters:
'   ent: AutoCAD Line entity
' Returns: clsDTSFrame object or Nothing on failure
Public Function ReadFrame(ent As Object) As clsDTSFrame
    On Error GoTo ErrHandler
    
    ' Validate input
    If ent Is Nothing Then
        m_LastError = "Entity is Nothing"
        LibDTS_Logger.Log DRIVER_NAME & ".ReadFrame: " & m_LastError, DTS_ERROR
        Set ReadFrame = Nothing
        Exit Function
    End If
    
    ' Check entity type
    If LCase$(TypeName(ent)) <> "acdbline" And InStr(LCase$(TypeName(ent)), "line") = 0 Then
        m_LastError = "Entity is not a line (got: " & TypeName(ent) & ")"
        LibDTS_Logger.Log DRIVER_NAME & ".ReadFrame: " & m_LastError, DTS_WARNING
        Set ReadFrame = Nothing
        Exit Function
    End If
    
    ' Create frame object
    Dim frame As New clsDTSFrame
    
    ' Get geometry from CAD
    Dim startPt As New clsDTSPoint
    Dim endPt As New clsDTSPoint
    startPt.Init ent.StartPoint(0), ent.StartPoint(1), ent.StartPoint(2)
    endPt.Init ent.EndPoint(0), ent.EndPoint(1), ent.EndPoint(2)
    Set frame.StartPoint = startPt
    Set frame.EndPoint = endPt
    
    ' Get XData
    Dim xDataStr As String
    xDataStr = GetRawXData(ent)
    
    If Len(xDataStr) > 0 Then
        frame.Deserialize xDataStr
        
        ' Self-healing: update handle if changed
        If frame.Base.CheckAndHealIdentity(ent.Handle) Then
            SaveXData frame, ent
        End If
    End If
    
    LibDTS_Logger.Log DRIVER_NAME & ".ReadFrame: Read frame " & frame.Base.guid, DTS_INFO
    Set ReadFrame = frame
    Exit Function
    
ErrHandler:
    m_LastError = "ReadFrame error: " & err.description
    LibDTS_Logger.Log DRIVER_NAME & ".ReadFrame: " & m_LastError, DTS_ERROR
    Set ReadFrame = Nothing
End Function

' Read an area from CAD entity
' Parameters:
'   ent: AutoCAD Polyline entity
' Returns: clsDTSArea object or Nothing on failure
Public Function ReadArea(ent As Object) As clsDTSArea
    On Error GoTo ErrHandler
    
    ' Validate input
    If ent Is Nothing Then
        m_LastError = "Entity is Nothing"
        LibDTS_Logger.Log DRIVER_NAME & ".ReadArea: " & m_LastError, DTS_ERROR
        Set ReadArea = Nothing
        Exit Function
    End If
    
    ' Check entity type
    Dim entType As String
    entType = LCase$(TypeName(ent))
    If InStr(entType, "polyline") = 0 Then
        m_LastError = "Entity is not a polyline (got: " & TypeName(ent) & ")"
        LibDTS_Logger.Log DRIVER_NAME & ".ReadArea: " & m_LastError, DTS_WARNING
        Set ReadArea = Nothing
        Exit Function
    End If
    
    ' Create area object
    Dim area As New clsDTSArea
    
    ' Get boundary points from polyline (this would need implementation in clsDTSArea)
    ' For now, note that this is placeholder
    
    ' Get XData
    Dim xDataStr As String
    xDataStr = GetRawXData(ent)
    
    If Len(xDataStr) > 0 Then
        area.Deserialize xDataStr
        
        ' Self-healing
        If area.Base.CheckAndHealIdentity(ent.Handle) Then
            SaveXData area, ent
        End If
    End If
    
    LibDTS_Logger.Log DRIVER_NAME & ".ReadArea: Read area " & area.Base.guid, DTS_INFO
    Set ReadArea = area
    Exit Function
    
ErrHandler:
    m_LastError = "ReadArea error: " & err.description
    LibDTS_Logger.Log DRIVER_NAME & ".ReadArea: " & m_LastError, DTS_ERROR
    Set ReadArea = Nothing
End Function

' ==========================================
' 3. BATCH READING OPERATIONS
' ==========================================

' Read all points from AutoCAD drawing
' Parameters:
'   acadDoc: AutoCAD Document object
' Returns: Collection of clsDTSPoint objects
Public Function ReadAllPoints(acadDoc As Object) As Collection
    On Error GoTo ErrHandler
    
    Dim points As New Collection
    
    If acadDoc Is Nothing Then
        m_LastError = "AutoCAD Document is Nothing"
        LibDTS_Logger.Log DRIVER_NAME & ".ReadAllPoints: " & m_LastError, DTS_ERROR
        Set ReadAllPoints = points
        Exit Function
    End If
    
    Dim ms As Object
    Set ms = acadDoc.ModelSpace
    
    Dim ent As Object
    Dim count As Long
    count = 0
    
    For Each ent In ms
        Dim entType As String
        entType = LCase$(TypeName(ent))
        
        If InStr(entType, "circle") > 0 Then
            Dim pt As clsDTSPoint
            Set pt = ReadPoint(ent)
            If Not pt Is Nothing Then
                points.Add pt
                count = count + 1
            End If
        End If
    Next ent
    
    LibDTS_Logger.Log DRIVER_NAME & ".ReadAllPoints: Read " & count & " points", DTS_INFO
    Set ReadAllPoints = points
    Exit Function
    
ErrHandler:
    m_LastError = "ReadAllPoints error: " & err.description
    LibDTS_Logger.Log DRIVER_NAME & ".ReadAllPoints: " & m_LastError, DTS_ERROR
    Set ReadAllPoints = New Collection
End Function

' Read all frames from AutoCAD drawing
' Parameters:
'   acadDoc: AutoCAD Document object
' Returns: Collection of clsDTSFrame objects
Public Function ReadAllFrames(acadDoc As Object) As Collection
    On Error GoTo ErrHandler
    
    Dim frames As New Collection
    
    If acadDoc Is Nothing Then
        m_LastError = "AutoCAD Document is Nothing"
        LibDTS_Logger.Log DRIVER_NAME & ".ReadAllFrames: " & m_LastError, DTS_ERROR
        Set ReadAllFrames = frames
        Exit Function
    End If
    
    Dim ms As Object
    Set ms = acadDoc.ModelSpace
    
    Dim ent As Object
    Dim count As Long
    count = 0
    
    For Each ent In ms
        Dim entType As String
        entType = LCase$(TypeName(ent))
        
        If InStr(entType, "line") > 0 Then
            Dim frame As clsDTSFrame
            Set frame = ReadFrame(ent)
            If Not frame Is Nothing Then
                frames.Add frame
                count = count + 1
            End If
        End If
    Next ent
    
    LibDTS_Logger.Log DRIVER_NAME & ".ReadAllFrames: Read " & count & " frames", DTS_INFO
    Set ReadAllFrames = frames
    Exit Function
    
ErrHandler:
    m_LastError = "ReadAllFrames error: " & err.description
    LibDTS_Logger.Log DRIVER_NAME & ".ReadAllFrames: " & m_LastError, DTS_ERROR
    Set ReadAllFrames = New Collection
End Function

' Read all areas from AutoCAD drawing
' Parameters:
'   acadDoc: AutoCAD Document object
' Returns: Collection of clsDTSArea objects
Public Function ReadAllAreas(acadDoc As Object) As Collection
    On Error GoTo ErrHandler
    
    Dim areas As New Collection
    
    If acadDoc Is Nothing Then
        m_LastError = "AutoCAD Document is Nothing"
        LibDTS_Logger.Log DRIVER_NAME & ".ReadAllAreas: " & m_LastError, DTS_ERROR
        Set ReadAllAreas = areas
        Exit Function
    End If
    
    Dim ms As Object
    Set ms = acadDoc.ModelSpace
    
    Dim ent As Object
    Dim count As Long
    count = 0
    
    For Each ent In ms
        Dim entType As String
        entType = LCase$(TypeName(ent))
        
        If InStr(entType, "polyline") > 0 Then
            Dim area As clsDTSArea
            Set area = ReadArea(ent)
            If Not area Is Nothing Then
                areas.Add area
                count = count + 1
            End If
        End If
    Next ent
    
    LibDTS_Logger.Log DRIVER_NAME & ".ReadAllAreas: Read " & count & " areas", DTS_INFO
    Set ReadAllAreas = areas
    Exit Function
    
ErrHandler:
    m_LastError = "ReadAllAreas error: " & err.description
    LibDTS_Logger.Log DRIVER_NAME & ".ReadAllAreas: " & m_LastError, DTS_ERROR
    Set ReadAllAreas = New Collection
End Function

' ==========================================
' 4. XDATA OPERATIONS
' ==========================================

' Save XData to entity
' Parameters:
'   dtsObj: DTS object with Serialize method
'   ent: AutoCAD entity
Public Sub SaveXData(dtsObj As Object, ent As Object)
    On Error GoTo ErrHandler
    
    If dtsObj Is Nothing Or ent Is Nothing Then Exit Sub
    
    ' Register application
    On Error Resume Next
    ent.Document.RegisteredApplications.Add m_AppName
    On Error GoTo ErrHandler
    
    ' Validate identity (self-healing)
    dtsObj.Base.ValidateIdentity ent.Handle
    
    ' Serialize content
    Dim dataStr As String
    dataStr = dtsObj.Serialize()
    
    ' Write to XData
    Dim xType(0 To 1) As Integer
    Dim xVal(0 To 1) As Variant
    
    xType(0) = 1001: xVal(0) = m_AppName
    xType(1) = 1000: xVal(1) = dataStr
    
    ent.SetXData xType, xVal
    
    LibDTS_Logger.Log DRIVER_NAME & ".SaveXData: Saved XData for entity " & ent.Handle, DTS_INFO
    Exit Sub
    
ErrHandler:
    m_LastError = "SaveXData error: " & err.description
    LibDTS_Logger.Log DRIVER_NAME & ".SaveXData: " & m_LastError, DTS_ERROR
End Sub

' Read raw XData string from entity
' Parameters:
'   ent: AutoCAD entity
'   appName: Application name (optional, uses default if not specified)
' Returns: XData string or "" if not found
Public Function ReadXData(ent As Object, Optional appName As String = "") As String
    If Len(appName) = 0 Then appName = m_AppName
    ReadXData = GetRawXData(ent, appName)
End Function

' Check if entity has XData
' Parameters:
'   ent: AutoCAD entity
'   appName: Application name (optional, uses default if not specified)
' Returns: True if entity has XData for specified app
Public Function HasXData(ent As Object, Optional appName As String = "") As Boolean
    On Error Resume Next
    If Len(appName) = 0 Then appName = m_AppName
    HasXData = (Len(GetRawXData(ent, appName)) > 0)
End Function

' ==========================================
' 5. GUID MAPPING OPERATIONS
' ==========================================

' Find entity by GUID in AutoCAD drawing
' Parameters:
'   acadDoc: AutoCAD Document object
'   guid: GUID string to find
' Returns: AutoCAD entity object or Nothing if not found
Public Function FindEntityByGUID(acadDoc As Object, guid As String) As Object
    On Error GoTo ErrHandler
    
    ' Check cache first
    If m_EntityCache.Exists(guid) Then
        Dim handle As String
        handle = m_EntityCache(guid)
        
        ' Try to get entity by handle
        On Error Resume Next
        Set FindEntityByGUID = acadDoc.HandleToObject(handle)
        If Not FindEntityByGUID Is Nothing Then Exit Function
        On Error GoTo ErrHandler
    End If
    
    ' Scan drawing (expensive)
    Dim ms As Object
    Set ms = acadDoc.ModelSpace
    
    Dim ent As Object
    For Each ent In ms
        If HasXData(ent) Then
            Dim xDataStr As String
            xDataStr = GetRawXData(ent)
            
            ' Check if GUID matches (simple string search)
            If InStr(xDataStr, guid) > 0 Then
                ' Cache and return
                If m_EntityCache.Exists(guid) Then
                    m_EntityCache(guid) = ent.Handle
                Else
                    m_EntityCache.Add guid, ent.Handle
                End If
                
                Set FindEntityByGUID = ent
                LibDTS_Logger.Log DRIVER_NAME & ".FindEntityByGUID: Found entity " & ent.Handle & " for GUID " & guid, DTS_INFO
                Exit Function
            End If
        End If
    Next ent
    
    ' Not found
    LibDTS_Logger.Log DRIVER_NAME & ".FindEntityByGUID: GUID not found: " & guid, DTS_WARNING
    Set FindEntityByGUID = Nothing
    Exit Function
    
ErrHandler:
    m_LastError = "FindEntityByGUID error: " & err.description
    LibDTS_Logger.Log DRIVER_NAME & ".FindEntityByGUID: " & m_LastError, DTS_ERROR
    Set FindEntityByGUID = Nothing
End Function

' Map GUID to entity handle
' Parameters:
'   guid: GUID string
'   handle: AutoCAD entity handle
' Returns: True if mapping created/updated
Public Function MapGUIDToHandle(guid As String, handle As String) As Boolean
    On Error GoTo ErrHandler
    
    If m_EntityCache.Exists(guid) Then
        m_EntityCache(guid) = handle
    Else
        m_EntityCache.Add guid, handle
    End If
    
    LibDTS_Logger.Log DRIVER_NAME & ".MapGUIDToHandle: Mapped GUID " & guid & " to handle " & handle, DTS_INFO
    MapGUIDToHandle = True
    Exit Function
    
ErrHandler:
    m_LastError = "MapGUIDToHandle error: " & err.description
    LibDTS_Logger.Log DRIVER_NAME & ".MapGUIDToHandle: " & m_LastError, DTS_ERROR
    MapGUIDToHandle = False
End Function

' ==========================================
' 6. UTILITY FUNCTIONS
' ==========================================

' Get last error message
' Returns: String describing last error
Public Function GetLastError() As String
    GetLastError = m_LastError
    m_LastError = "" ' Clear after reading
End Function

' Clear entity cache
Public Sub ClearCache()
    On Error Resume Next
    Set m_EntityCache = CreateObject("Scripting.Dictionary")
    LibDTS_Logger.Log DRIVER_NAME & ".ClearCache: Cache cleared", DTS_INFO
End Sub

' Set application name for XData
' Parameters:
'   appName: Application name to use
Public Sub SetAppName(appName As String)
    If Len(Trim$(appName)) > 0 Then
        m_AppName = appName
        LibDTS_Logger.Log DRIVER_NAME & ".SetAppName: App name set to " & appName, DTS_INFO
    End If
End Sub

' ==========================================
' PRIVATE HELPER FUNCTIONS
' ==========================================

' Get raw XData string from entity
Private Function GetRawXData(ent As Object, Optional appName As String = "") As String
    On Error Resume Next
    
    If Len(appName) = 0 Then appName = m_AppName
    
    Dim xType As Variant, xVal As Variant
    ent.GetXData appName, xType, xVal
    
    If err.number = 0 And Not IsEmpty(xVal) Then
        If IsArray(xVal) Then
            If UBound(xVal) >= 1 Then
                GetRawXData = CStr(xVal(1))
            End If
        End If
    Else
        GetRawXData = ""
    End If
    
    On Error GoTo 0
End Function
