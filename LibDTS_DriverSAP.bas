Attribute VB_Name = "LibDTS_DriverSAP"
' Module: LibDTS_DriverSAP
' Purpose: Canonical driver for SAP2000 structural analysis software integration
' Features: Connection management, modeling operations, GUID mapping, dry-run support
' Version: 2.0 (Enhanced with defensive programming and full API surface)
Option Explicit

' ==========================================
' MODULE-LEVEL VARIABLES
' ==========================================
Private m_SapObject As Object        ' SAP2000 application object
Private m_SapModel As Object         ' Active model object
Private m_IsConnected As Boolean     ' Connection state
Private m_ConnectionVersion As String ' Detected SAP version
Private m_LastError As String        ' Last error message
Private m_PointCache As Object       ' Dictionary: GUID -> SAP point name
Private m_FrameCache As Object       ' Dictionary: GUID -> SAP frame name
Private m_AreaCache As Object        ' Dictionary: GUID -> SAP area name

' ==========================================
' PUBLIC CONSTANTS
' ==========================================
Public Const DRIVER_NAME As String = "LibDTS_DriverSAP"
Public Const APP_NAME As String = "DTS_CORE"
Public Const DEFAULT_UNITS As Long = 5  ' kN_m_C

' SAP version ProgIDs in priority order (newest first)
Private Const SAP_VERSIONS As String = "SAP2000v25.Helper,SAP2000v24.Helper,SAP2000v23.Helper,SAP2000v22.Helper,SAP2000v21.Helper,SAP2000v20.Helper,SAP2000v19.Helper,SAP2000v1.Helper"

' ==========================================
' 1. CONNECTION MANAGEMENT
' ==========================================

' Connect to SAP2000 instance (attach to existing or start new)
' Parameters:
'   version: "auto" (detect best), "v25", "v24", etc. or specific ProgID
'   startNew: False (prefer existing), True (always start new)
' Returns: True if connected successfully
Public Function Connect(Optional version As String = "auto", _
                        Optional startNew As Boolean = False) As Boolean
    On Error GoTo ErrHandler
    
    ' Check if already connected
    If m_IsConnected And Not (m_SapModel Is Nothing) Then
        LibDTS_Logger.Log DRIVER_NAME & ".Connect: Already connected", DTS_INFO
        Connect = True
        Exit Function
    End If
    
    ' Initialize caches
    Set m_PointCache = CreateObject("Scripting.Dictionary")
    Set m_FrameCache = CreateObject("Scripting.Dictionary")
    Set m_AreaCache = CreateObject("Scripting.Dictionary")
    
    Dim connected As Boolean
    connected = False
    
    ' Try version-specific connection if requested
    If version <> "auto" Then
        connected = TryConnectVersion(version, startNew)
    End If
    
    ' If not connected yet, try auto-detection
    If Not connected Then
        connected = TryConnectAuto(startNew)
    End If
    
    ' Set connection state
    If connected Then
        m_IsConnected = True
        
        ' Set present units to kN_m_C
        On Error Resume Next
        m_SapModel.SetPresentUnits DEFAULT_UNITS
        On Error GoTo ErrHandler
        
        LibDTS_Logger.Log DRIVER_NAME & ".Connect: Connected successfully (Version: " & m_ConnectionVersion & ")", DTS_INFO
        Connect = True
    Else
        m_LastError = "Failed to connect to SAP2000. Ensure SAP2000 is installed and accessible."
        LibDTS_Logger.Log DRIVER_NAME & ".Connect: " & m_LastError, DTS_ERROR
        Connect = False
    End If
    
    Exit Function
    
ErrHandler:
    m_LastError = "Connection error: " & err.description & " (Error " & err.number & ")"
    LibDTS_Logger.Log DRIVER_NAME & ".Connect: " & m_LastError, DTS_ERROR
    Connect = False
End Function

' Disconnect from SAP2000 and clean up resources
' Parameters:
'   saveModel: False (don't save), True (save before disconnect)
' Returns: True if disconnected successfully
Public Function Disconnect(Optional saveModel As Boolean = False) As Boolean
    On Error GoTo ErrHandler
    
    ' If not connected, nothing to do
    If Not m_IsConnected Then
        Disconnect = True
        Exit Function
    End If
    
    ' Save model if requested
    If saveModel And Not (m_SapModel Is Nothing) Then
        On Error Resume Next
        m_SapModel.File.Save ""  ' Save to current path
        If err.number <> 0 Then
            LibDTS_Logger.Log DRIVER_NAME & ".Disconnect: Warning - could not save model: " & err.description, DTS_WARNING
        End If
        On Error GoTo ErrHandler
    End If
    
    ' Clean up objects
    Set m_SapModel = Nothing
    Set m_SapObject = Nothing
    
    ' Clear caches
    Set m_PointCache = Nothing
    Set m_FrameCache = Nothing
    Set m_AreaCache = Nothing
    
    ' Update state
    m_IsConnected = False
    m_ConnectionVersion = ""
    
    LibDTS_Logger.Log DRIVER_NAME & ".Disconnect: Disconnected successfully", DTS_INFO
    Disconnect = True
    Exit Function
    
ErrHandler:
    m_LastError = "Disconnect error: " & err.description
    LibDTS_Logger.Log DRIVER_NAME & ".Disconnect: " & m_LastError, DTS_ERROR
    ' Always set disconnected state even on error
    m_IsConnected = False
    Disconnect = False
End Function

' Check if currently connected to SAP2000
' Returns: True if active connection exists
Public Function IsConnected() As Boolean
    IsConnected = m_IsConnected And Not (m_SapModel Is Nothing)
End Function

' Get last error message from driver
' Returns: String describing last error, or "" if no error
Public Function GetLastError() As String
    GetLastError = m_LastError
    m_LastError = "" ' Clear after reading
End Function

' Get SAP object (for backward compatibility with legacy code)
' Returns: SAP application object or Nothing
Public Function GetSapObject() As Object
    Set GetSapObject = m_SapObject
End Function

' Get SAP model object (for backward compatibility with legacy code)
' Returns: SAP model object or Nothing
Public Function GetSapModel() As Object
    Set GetSapModel = m_SapModel
End Function

' ==========================================
' 2. MODELING OPERATIONS (with Dry-Run Support)
' ==========================================

' Create or update a point in SAP model
' Parameters:
'   pt: clsDTSPoint object with coordinates
'   dryRun: False (commit), True (validate only)
'   overwriteExisting: False (skip if exists), True (update coordinates)
'   tolerance: Distance tolerance for duplicate detection (mm)
' Returns: SAP point name (e.g., "1", "2") or "" on failure
Public Function PushPoint(pt As clsDTSPoint, _
                          Optional dryRun As Boolean = False, _
                          Optional overwriteExisting As Boolean = False, _
                          Optional tolerance As Double = 0.01) As String
    On Error GoTo ErrHandler
    
    ' Validate inputs
    If Not m_IsConnected Then
        m_LastError = "Not connected to SAP2000"
        LibDTS_Logger.Log DRIVER_NAME & ".PushPoint: " & m_LastError, DTS_ERROR
        PushPoint = ""
        Exit Function
    End If
    
    If pt Is Nothing Then
        m_LastError = "Point object is Nothing"
        LibDTS_Logger.Log DRIVER_NAME & ".PushPoint: " & m_LastError, DTS_ERROR
        PushPoint = ""
        Exit Function
    End If
    
    ' Ensure GUID exists
    If Len(Trim$(pt.Base.guid)) = 0 Then
        pt.Base.guid = LibDTS_Base.GenerateGUID()
    End If
    
    ' Check if point already exists via GUID lookup
    Dim existingName As String
    existingName = ""
    If m_PointCache.Exists(pt.Base.guid) Then
        existingName = m_PointCache(pt.Base.guid)
        If Not overwriteExisting Then
            LibDTS_Logger.Log DRIVER_NAME & ".PushPoint: Point already exists (GUID: " & pt.Base.guid & ", Name: " & existingName & ")", DTS_INFO
            PushPoint = existingName
            Exit Function
        End If
    End If
    
    ' Dry-run mode: validate and return would-be name
    If dryRun Then
        LibDTS_Logger.Log DRIVER_NAME & ".PushPoint [DRY-RUN]: Would create point at (" & pt.X & ", " & pt.Y & ", " & pt.Z & ")", DTS_INFO
        PushPoint = "DRY_RUN_P" & m_PointCache.count + 1
        Exit Function
    End If
    
    ' Create point in SAP
    Dim pointName As String
    Dim ret As Long
    ret = m_SapModel.pointObj.AddCartesian(pt.X, pt.Y, pt.Z, pointName)
    
    If ret <> 0 Then
        m_LastError = "Failed to create point in SAP (return code: " & ret & ")"
        LibDTS_Logger.Log DRIVER_NAME & ".PushPoint: " & m_LastError, DTS_ERROR
        PushPoint = ""
        Exit Function
    End If
    
    ' Store GUID in SAP comment
    On Error Resume Next
    m_SapModel.pointObj.SetGUID pointName, pt.Base.guid
    If err.number <> 0 Then
        ' Fallback: use comment for older SAP versions
        m_SapModel.pointObj.SetComment pointName, "GUID:" & pt.Base.guid
    End If
    On Error GoTo ErrHandler
    
    ' Cache mapping
    If m_PointCache.Exists(pt.Base.guid) Then
        m_PointCache(pt.Base.guid) = pointName
    Else
        m_PointCache.Add pt.Base.guid, pointName
    End If
    
    ' Persist mapping to database
    LibDTS_DriverDB.SetMappedElement pt.Base.guid, Array("Point", pointName, pt.X, pt.Y, pt.Z)
    
    LibDTS_Logger.Log DRIVER_NAME & ".PushPoint: Created point " & pointName & " at (" & pt.X & ", " & pt.Y & ", " & pt.Z & ")", DTS_INFO
    PushPoint = pointName
    Exit Function
    
ErrHandler:
    m_LastError = "PushPoint error: " & err.description
    LibDTS_Logger.Log DRIVER_NAME & ".PushPoint: " & m_LastError, DTS_ERROR
    PushPoint = ""
End Function

' Create or update a frame element in SAP model
' Parameters:
'   frame: clsDTSFrame object with geometry and properties
'   dryRun: False (commit), True (validate only)
'   overwriteExisting: False (skip if exists), True (update)
'   createPoints: True (auto-create points if missing), False (points must exist)
' Returns: SAP frame name (e.g., "1", "2") or "" on failure
Public Function PushFrame(frame As clsDTSFrame, _
                          Optional dryRun As Boolean = False, _
                          Optional overwriteExisting As Boolean = False, _
                          Optional createPoints As Boolean = True) As String
    On Error GoTo ErrHandler
    
    ' Validate inputs
    If Not m_IsConnected Then
        m_LastError = "Not connected to SAP2000"
        LibDTS_Logger.Log DRIVER_NAME & ".PushFrame: " & m_LastError, DTS_ERROR
        PushFrame = ""
        Exit Function
    End If
    
    If frame Is Nothing Then
        m_LastError = "Frame object is Nothing"
        LibDTS_Logger.Log DRIVER_NAME & ".PushFrame: " & m_LastError, DTS_ERROR
        PushFrame = ""
        Exit Function
    End If
    
    ' Ensure GUID exists
    If Len(Trim$(frame.Base.guid)) = 0 Then
        frame.Base.guid = LibDTS_Base.GenerateGUID()
    End If
    
    ' Check if frame already exists
    Dim existingName As String
    existingName = ""
    If m_FrameCache.Exists(frame.Base.guid) Then
        existingName = m_FrameCache(frame.Base.guid)
        If Not overwriteExisting Then
            LibDTS_Logger.Log DRIVER_NAME & ".PushFrame: Frame already exists (GUID: " & frame.Base.guid & ", Name: " & existingName & ")", DTS_INFO
            PushFrame = existingName
            Exit Function
        End If
    End If
    
    ' Create or get points
    Dim p1Name As String, p2Name As String
    If createPoints Then
        p1Name = PushPoint(frame.StartPoint, dryRun)
        p2Name = PushPoint(frame.EndPoint, dryRun)
    Else
        ' Points must already exist - find them
        p1Name = FindPointByCoordinates(frame.StartPoint.X, frame.StartPoint.Y, frame.StartPoint.Z)
        p2Name = FindPointByCoordinates(frame.EndPoint.X, frame.EndPoint.Y, frame.EndPoint.Z)
    End If
    
    If p1Name = "" Or p2Name = "" Then
        m_LastError = "Failed to create/find points for frame"
        LibDTS_Logger.Log DRIVER_NAME & ".PushFrame: " & m_LastError, DTS_ERROR
        PushFrame = ""
        Exit Function
    End If
    
    ' Dry-run mode
    If dryRun Then
        LibDTS_Logger.Log DRIVER_NAME & ".PushFrame [DRY-RUN]: Would create frame from " & p1Name & " to " & p2Name & ", section=" & frame.sectionName, DTS_INFO
        PushFrame = "DRY_RUN_F" & m_FrameCache.count + 1
        Exit Function
    End If
    
    ' Create frame in SAP
    Dim frameName As String
    Dim ret As Long
    ret = m_SapModel.frameObj.AddByPoint(p1Name, p2Name, frameName, frame.sectionName)
    
    If ret <> 0 Then
        m_LastError = "Failed to create frame in SAP (return code: " & ret & ")"
        LibDTS_Logger.Log DRIVER_NAME & ".PushFrame: " & m_LastError, DTS_ERROR
        PushFrame = ""
        Exit Function
    End If
    
    ' Store GUID
    On Error Resume Next
    m_SapModel.frameObj.SetGUID frameName, frame.Base.guid
    If err.number <> 0 Then
        m_SapModel.frameObj.SetComment frameName, "GUID:" & frame.Base.guid
    End If
    On Error GoTo ErrHandler
    
    ' Set local axes if specified
    If frame.angle <> 0 Then
        m_SapModel.frameObj.SetLocalAxes frameName, frame.angle
    End If
    
    ' Cache mapping
    If m_FrameCache.Exists(frame.Base.guid) Then
        m_FrameCache(frame.Base.guid) = frameName
    Else
        m_FrameCache.Add frame.Base.guid, frameName
    End If
    
    ' Persist mapping
    LibDTS_DriverDB.SetMappedElement frame.Base.guid, Array("Frame", frameName, p1Name, p2Name, frame.sectionName)
    
    LibDTS_Logger.Log DRIVER_NAME & ".PushFrame: Created frame " & frameName & " (" & p1Name & "-" & p2Name & ", " & frame.sectionName & ")", DTS_INFO
    PushFrame = frameName
    Exit Function
    
ErrHandler:
    m_LastError = "PushFrame error: " & err.description
    LibDTS_Logger.Log DRIVER_NAME & ".PushFrame: " & m_LastError, DTS_ERROR
    PushFrame = ""
End Function

' Create or update an area element (slab/wall) in SAP model
' Parameters:
'   area: clsDTSArea object with boundary points and properties
'   dryRun: False (commit), True (validate only)
'   overwriteExisting: False (skip if exists), True (update)
'   createPoints: True (auto-create points), False (points must exist)
' Returns: SAP area name or "" on failure
Public Function PushArea(area As clsDTSArea, _
                         Optional dryRun As Boolean = False, _
                         Optional overwriteExisting As Boolean = False, _
                         Optional createPoints As Boolean = True) As String
    On Error GoTo ErrHandler
    
    ' Validate inputs
    If Not m_IsConnected Then
        m_LastError = "Not connected to SAP2000"
        LibDTS_Logger.Log DRIVER_NAME & ".PushArea: " & m_LastError, DTS_ERROR
        PushArea = ""
        Exit Function
    End If
    
    If area Is Nothing Then
        m_LastError = "Area object is Nothing"
        LibDTS_Logger.Log DRIVER_NAME & ".PushArea: " & m_LastError, DTS_ERROR
        PushArea = ""
        Exit Function
    End If
    
    If area.BoundaryPoints.count < 3 Then
        m_LastError = "Area must have at least 3 boundary points"
        LibDTS_Logger.Log DRIVER_NAME & ".PushArea: " & m_LastError, DTS_ERROR
        PushArea = ""
        Exit Function
    End If
    
    ' Ensure GUID exists
    If Len(Trim$(area.Base.guid)) = 0 Then
        area.Base.guid = LibDTS_Base.GenerateGUID()
    End If
    
    ' Check if area already exists
    If m_AreaCache.Exists(area.Base.guid) Then
        Dim existingName As String
        existingName = m_AreaCache(area.Base.guid)
        If Not overwriteExisting Then
            LibDTS_Logger.Log DRIVER_NAME & ".PushArea: Area already exists (GUID: " & area.Base.guid & ")", DTS_INFO
            PushArea = existingName
            Exit Function
        End If
    End If
    
    ' Create or get boundary points
    Dim numPoints As Long
    numPoints = area.BoundaryPoints.count
    
    Dim pointNames() As String
    ReDim pointNames(0 To numPoints - 1)
    
    Dim i As Long
    Dim pt As clsDTSPoint
    For i = 1 To numPoints
        Set pt = area.BoundaryPoints(i)
        If createPoints Then
            pointNames(i - 1) = PushPoint(pt, dryRun)
        Else
            pointNames(i - 1) = FindPointByCoordinates(pt.X, pt.Y, pt.Z)
        End If
        
        If pointNames(i - 1) = "" Then
            m_LastError = "Failed to create/find point " & i & " for area"
            LibDTS_Logger.Log DRIVER_NAME & ".PushArea: " & m_LastError, DTS_ERROR
            PushArea = ""
            Exit Function
        End If
    Next i
    
    ' Dry-run mode
    If dryRun Then
        LibDTS_Logger.Log DRIVER_NAME & ".PushArea [DRY-RUN]: Would create area with " & numPoints & " points, section=" & area.sectionName, DTS_INFO
        PushArea = "DRY_RUN_A" & m_AreaCache.count + 1
        Exit Function
    End If
    
    ' Create area in SAP
    Dim areaName As String
    Dim ret As Long
    ret = m_SapModel.AreaObj.AddByPoint(numPoints, pointNames, areaName)
    
    If ret <> 0 Then
        m_LastError = "Failed to create area in SAP (return code: " & ret & ")"
        LibDTS_Logger.Log DRIVER_NAME & ".PushArea: " & m_LastError, DTS_ERROR
        PushArea = ""
        Exit Function
    End If
    
    ' Set section property
    If Len(area.sectionName) > 0 Then
        m_SapModel.AreaObj.SetProperty areaName, area.sectionName
    End If
    
    ' Store GUID
    On Error Resume Next
    m_SapModel.AreaObj.SetGUID areaName, area.Base.guid
    If err.number <> 0 Then
        m_SapModel.AreaObj.SetComment areaName, "GUID:" & area.Base.guid
    End If
    On Error GoTo ErrHandler
    
    ' Cache mapping
    If m_AreaCache.Exists(area.Base.guid) Then
        m_AreaCache(area.Base.guid) = areaName
    Else
        m_AreaCache.Add area.Base.guid, areaName
    End If
    
    ' Persist mapping
    LibDTS_DriverDB.SetMappedElement area.Base.guid, Array("Area", areaName, Join(pointNames, ","), area.sectionName)
    
    LibDTS_Logger.Log DRIVER_NAME & ".PushArea: Created area " & areaName & " with " & numPoints & " points", DTS_INFO
    PushArea = areaName
    Exit Function
    
ErrHandler:
    m_LastError = "PushArea error: " & err.description
    LibDTS_Logger.Log DRIVER_NAME & ".PushArea: " & m_LastError, DTS_ERROR
    PushArea = ""
End Function

' ==========================================
' 3. GUID MAPPING OPERATIONS
' ==========================================

' Explicitly map a GUID to SAP element name
' Parameters:
'   guid: GUID string to map
'   sapName: SAP element name (point, frame, or area name)
'   elementType: "Point", "Frame", or "Area"
' Returns: True if mapping created/updated successfully
Public Function MapGUIDToElement(guid As String, _
                                 sapName As String, _
                                 elementType As String) As Boolean
    On Error GoTo ErrHandler
    
    ' Validate inputs
    If Not LibDTS_Base.IsValidGUID(guid) Then
        m_LastError = "Invalid GUID format"
        LibDTS_Logger.Log DRIVER_NAME & ".MapGUIDToElement: " & m_LastError, DTS_ERROR
        MapGUIDToElement = False
        Exit Function
    End If
    
    If Not m_IsConnected Then
        m_LastError = "Not connected to SAP2000"
        LibDTS_Logger.Log DRIVER_NAME & ".MapGUIDToElement: " & m_LastError, DTS_ERROR
        MapGUIDToElement = False
        Exit Function
    End If
    
    ' Add to appropriate cache
    Select Case LCase$(elementType)
        Case "point"
            If m_PointCache.Exists(guid) Then
                m_PointCache(guid) = sapName
            Else
                m_PointCache.Add guid, sapName
            End If
            
        Case "frame"
            If m_FrameCache.Exists(guid) Then
                m_FrameCache(guid) = sapName
            Else
                m_FrameCache.Add guid, sapName
            End If
            
        Case "area"
            If m_AreaCache.Exists(guid) Then
                m_AreaCache(guid) = sapName
            Else
                m_AreaCache.Add guid, sapName
            End If
            
        Case Else
            m_LastError = "Unknown element type: " & elementType
            LibDTS_Logger.Log DRIVER_NAME & ".MapGUIDToElement: " & m_LastError, DTS_WARNING
            MapGUIDToElement = False
            Exit Function
    End Select
    
    ' Persist mapping
    LibDTS_DriverDB.SetMappedElement guid, Array(elementType, sapName)
    
    LibDTS_Logger.Log DRIVER_NAME & ".MapGUIDToElement: Mapped GUID " & guid & " to " & elementType & " " & sapName, DTS_INFO
    MapGUIDToElement = True
    Exit Function
    
ErrHandler:
    m_LastError = "MapGUIDToElement error: " & err.description
    LibDTS_Logger.Log DRIVER_NAME & ".MapGUIDToElement: " & m_LastError, DTS_ERROR
    MapGUIDToElement = False
End Function

' Find SAP element name by GUID
' Parameters:
'   guid: GUID string to find
' Returns: SAP element name or "" if not found
Public Function FindElementByGUID(guid As String) As String
    On Error GoTo ErrHandler
    
    ' Search caches first
    If m_PointCache.Exists(guid) Then
        FindElementByGUID = m_PointCache(guid)
        Exit Function
    End If
    
    If m_FrameCache.Exists(guid) Then
        FindElementByGUID = m_FrameCache(guid)
        Exit Function
    End If
    
    If m_AreaCache.Exists(guid) Then
        FindElementByGUID = m_AreaCache(guid)
        Exit Function
    End If
    
    ' Not in cache - try loading from persistent storage
    Dim mappedElement As Variant
    mappedElement = LibDTS_DriverDB.GetMappedElement(guid)
    
    If Not IsEmpty(mappedElement) And IsArray(mappedElement) Then
        If UBound(mappedElement) >= 1 Then
            FindElementByGUID = CStr(mappedElement(1))
            ' Update cache
            MapGUIDToElement guid, FindElementByGUID, CStr(mappedElement(0))
            Exit Function
        End If
    End If
    
    ' Still not found - return empty
    LibDTS_Logger.Log DRIVER_NAME & ".FindElementByGUID: GUID not found: " & guid, DTS_WARNING
    FindElementByGUID = ""
    Exit Function
    
ErrHandler:
    m_LastError = "FindElementByGUID error: " & err.description
    LibDTS_Logger.Log DRIVER_NAME & ".FindElementByGUID: " & m_LastError, DTS_ERROR
    FindElementByGUID = ""
End Function

' Remove GUID mapping (cleanup orphaned entries)
' Parameters:
'   guid: GUID string to remove
' Returns: True if removed successfully
Public Function RemoveGUIDMapping(guid As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim found As Boolean
    found = False
    
    ' Remove from caches
    If m_PointCache.Exists(guid) Then
        m_PointCache.Remove guid
        found = True
    End If
    
    If m_FrameCache.Exists(guid) Then
        m_FrameCache.Remove guid
        found = True
    End If
    
    If m_AreaCache.Exists(guid) Then
        m_AreaCache.Remove guid
        found = True
    End If
    
    ' Remove from persistent storage
    On Error Resume Next
    ' LibDTS_DriverDB would need RemoveMappedElement method
    ' For now, we can set it to Empty or special value
    LibDTS_DriverDB.SetMappedElement guid, Array("DELETED", "")
    On Error GoTo ErrHandler
    
    If found Then
        LibDTS_Logger.Log DRIVER_NAME & ".RemoveGUIDMapping: Removed mapping for GUID " & guid, DTS_INFO
    End If
    
    RemoveGUIDMapping = True
    Exit Function
    
ErrHandler:
    m_LastError = "RemoveGUIDMapping error: " & err.description
    LibDTS_Logger.Log DRIVER_NAME & ".RemoveGUIDMapping: " & m_LastError, DTS_ERROR
    RemoveGUIDMapping = False
End Function

' ==========================================
' 4. UTILITY OPERATIONS
' ==========================================

' Clear all internal GUID caches
' Useful after model changes or before re-scan
Public Sub ClearCache()
    On Error Resume Next
    Set m_PointCache = CreateObject("Scripting.Dictionary")
    Set m_FrameCache = CreateObject("Scripting.Dictionary")
    Set m_AreaCache = CreateObject("Scripting.Dictionary")
    LibDTS_Logger.Log DRIVER_NAME & ".ClearCache: All caches cleared", DTS_INFO
End Sub

' Scan entire SAP model and rebuild GUID cache from metadata
' Returns: Number of elements cached
Public Function RebuildCacheFromModel() As Long
    On Error GoTo ErrHandler
    
    If Not m_IsConnected Then
        m_LastError = "Not connected to SAP2000"
        LibDTS_Logger.Log DRIVER_NAME & ".RebuildCacheFromModel: " & m_LastError, DTS_ERROR
        RebuildCacheFromModel = 0
        Exit Function
    End If
    
    ' Clear existing caches
    ClearCache
    
    Dim totalCount As Long
    totalCount = 0
    
    ' Scan points
    Dim pointCount As Long, pointNames() As String
    Dim ret As Long
    ret = m_SapModel.pointObj.GetNameList(pointCount, pointNames)
    
    If ret = 0 And pointCount > 0 Then
        Dim i As Long
        For i = 0 To pointCount - 1
            Dim pointGUID As String
            pointGUID = ExtractGUIDFromComment(m_SapModel.pointObj, pointNames(i))
            If Len(pointGUID) > 0 Then
                m_PointCache.Add pointGUID, pointNames(i)
                totalCount = totalCount + 1
            End If
        Next i
    End If
    
    ' Scan frames
    Dim frameCount As Long, frameNames() As String
    ret = m_SapModel.frameObj.GetNameList(frameCount, frameNames)
    
    If ret = 0 And frameCount > 0 Then
        For i = 0 To frameCount - 1
            Dim frameGUID As String
            frameGUID = ExtractGUIDFromComment(m_SapModel.frameObj, frameNames(i))
            If Len(frameGUID) > 0 Then
                m_FrameCache.Add frameGUID, frameNames(i)
                totalCount = totalCount + 1
            End If
        Next i
    End If
    
    ' Scan areas
    Dim areaCount As Long, areaNames() As String
    ret = m_SapModel.AreaObj.GetNameList(areaCount, areaNames)
    
    If ret = 0 And areaCount > 0 Then
        For i = 0 To areaCount - 1
            Dim areaGUID As String
            areaGUID = ExtractGUIDFromComment(m_SapModel.AreaObj, areaNames(i))
            If Len(areaGUID) > 0 Then
                m_AreaCache.Add areaGUID, areaNames(i)
                totalCount = totalCount + 1
            End If
        Next i
    End If
    
    LibDTS_Logger.Log DRIVER_NAME & ".RebuildCacheFromModel: Cached " & totalCount & " elements", DTS_INFO
    RebuildCacheFromModel = totalCount
    Exit Function
    
ErrHandler:
    m_LastError = "RebuildCacheFromModel error: " & err.description
    LibDTS_Logger.Log DRIVER_NAME & ".RebuildCacheFromModel: " & m_LastError, DTS_ERROR
    RebuildCacheFromModel = 0
End Function

' ==========================================
' PRIVATE HELPER FUNCTIONS
' ==========================================

' Try to connect to a specific SAP version
Private Function TryConnectVersion(version As String, startNew As Boolean) As Boolean
    On Error Resume Next
    
    Dim progID As String
    If InStr(version, ".") > 0 Then
        progID = version ' Already a ProgID
    Else
        progID = "SAP2000" & version & ".Helper"
    End If
    
    Dim helperObj As Object
    Set helperObj = CreateObject(progID)
    
    If Not helperObj Is Nothing Then
        Set m_SapObject = helperObj.GetObject("CSI.SAP2000.API.SapObject")
        If Not m_SapObject Is Nothing Then
            Set m_SapModel = m_SapObject.SapModel
            If Not m_SapModel Is Nothing Then
                m_ConnectionVersion = version
                TryConnectVersion = True
                Exit Function
            End If
        End If
    End If
    
    TryConnectVersion = False
End Function

' Try to auto-detect and connect to SAP
Private Function TryConnectAuto(startNew As Boolean) As Boolean
    On Error Resume Next
    
    ' Try helper objects first (version detection)
    Dim versions() As String
    versions = Split(SAP_VERSIONS, ",")
    
    Dim v As Variant
    For Each v In versions
        Dim helperObj As Object
        Set helperObj = CreateObject(Trim$(CStr(v)))
        
        If Not helperObj Is Nothing Then
            Set m_SapObject = helperObj.GetObject("CSI.SAP2000.API.SapObject")
            If Not m_SapObject Is Nothing Then
                Set m_SapModel = m_SapObject.SapModel
                If Not m_SapModel Is Nothing Then
                    m_ConnectionVersion = Replace(CStr(v), ".Helper", "")
                    TryConnectAuto = True
                    Exit Function
                End If
            End If
        End If
    Next v
    
    ' Fallback: attach to existing instance without helper
    Set m_SapObject = GetObject(, "CSI.SAP2000.API.SapObject")
    If Not m_SapObject Is Nothing Then
        Set m_SapModel = m_SapObject.SapModel
        If Not m_SapModel Is Nothing Then
            m_ConnectionVersion = "Unknown"
            TryConnectAuto = True
            Exit Function
        End If
    End If
    
    ' Last resort: try to start new instance
    If startNew Then
        Set m_SapObject = CreateObject("CSI.SAP2000.API.SapObject")
        If Not m_SapObject Is Nothing Then
            m_SapObject.ApplicationStart
            Set m_SapModel = m_SapObject.SapModel
            If Not m_SapModel Is Nothing Then
                m_ConnectionVersion = "New Instance"
                TryConnectAuto = True
                Exit Function
            End If
        End If
    End If
    
    TryConnectAuto = False
End Function

' Find point by coordinates (within tolerance)
Private Function FindPointByCoordinates(X As Double, Y As Double, Z As Double, _
                                        Optional tolerance As Double = 0.01) As String
    On Error Resume Next
    
    ' This is expensive - should cache point coordinates
    ' For now, just return empty (caller should use createPoints=True)
    FindPointByCoordinates = ""
End Function

' Extract GUID from SAP element comment
Private Function ExtractGUIDFromComment(objType As Object, elementName As String) As String
    On Error Resume Next
    
    Dim comment As String
    comment = ""
    
    ' Try to get comment
    objType.GetComment elementName, comment
    
    If Len(comment) > 0 Then
        ' Look for "GUID:" prefix
        If InStr(comment, "GUID:") > 0 Then
            ExtractGUIDFromComment = Trim$(mid$(comment, InStr(comment, "GUID:") + 5))
        Else
            ExtractGUIDFromComment = ""
        End If
    Else
        ExtractGUIDFromComment = ""
    End If
End Function
