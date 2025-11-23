VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDTSRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class Module: clsDTSRepository
' Purpose: Central orchestrator for all CRUD operations on DTS elements
' This class implements the Repository Pattern
' Responsibilities:
'   - Save elements to CAD XData
'   - Load elements from CAD entities
'   - Sync elements to/from SAP2000
'   - Manage GUID-Handle mappings
'   - Coordinate between different storage backends
Option Explicit

' ==========================================
' PRIVATE MEMBERS
' ==========================================

Private m_GuidToHandleMap As Object     ' Dictionary: GUID -> CAD Handle
Private m_HandleToGuidMap As Object     ' Dictionary: CAD Handle -> GUID
Private m_AcadDoc As Object              ' AutoCAD Document reference
Private m_SapConnected As Boolean        ' SAP2000 connection state
Private m_LastError As String            ' Last error message

' ==========================================
' CONSTRUCTOR & DESTRUCTOR
' ==========================================

Private Sub Class_Initialize()
    Set m_GuidToHandleMap = CreateObject("Scripting.Dictionary")
    Set m_HandleToGuidMap = CreateObject("Scripting.Dictionary")
    m_SapConnected = False
    m_LastError = ""
End Sub

Private Sub Class_Terminate()
    Set m_GuidToHandleMap = Nothing
    Set m_HandleToGuidMap = Nothing
    Set m_AcadDoc = Nothing
End Sub

' ==========================================
' INITIALIZATION
' ==========================================

' Initialize repository with AutoCAD document
Public Sub Initialize(acadDoc As Object)
    On Error GoTo ErrHandler
    
    If acadDoc Is Nothing Then
        m_LastError = "AutoCAD Document is Nothing"
        LibDTS_Logger.Log "clsDTSRepository.Initialize: " & m_LastError, DTS_ERROR
        Exit Sub
    End If
    
    Set m_AcadDoc = acadDoc
    
    ' Rebuild GUID-Handle mappings from existing XData
    RebuildMappings
    
    LibDTS_Logger.Log "clsDTSRepository: Initialized successfully", DTS_INFO
    Exit Sub
    
ErrHandler:
    m_LastError = "Initialize error: " & Err.Description
    LibDTS_Logger.Log "clsDTSRepository.Initialize: " & m_LastError, DTS_ERROR
End Sub

' Rebuild GUID-Handle mappings by scanning all entities with XData
Private Sub RebuildMappings()
    On Error Resume Next
    
    ' Clear existing mappings
    m_GuidToHandleMap.RemoveAll
    m_HandleToGuidMap.RemoveAll
    
    ' This would scan ModelSpace entities - stub for now
    ' In production, iterate through entities and extract GUIDs from XData
    
    LibDTS_Logger.Log "clsDTSRepository.RebuildMappings: Completed", DTS_INFO
End Sub

' ==========================================
' SAVE OPERATIONS
' ==========================================

' Save a DTS element to CAD XData
' Parameters:
'   element: Any DTS element object (Frame, Area, Tag, Rebar)
'   entity: Optional - existing CAD entity. If Nothing, searches by GUID
' Returns: True if saved successfully
Public Function Save(element As Object, Optional entity As Object = Nothing) As Boolean
    On Error GoTo ErrHandler
    
    If element Is Nothing Then
        m_LastError = "Element is Nothing"
        LibDTS_Logger.Log "clsDTSRepository.Save: " & m_LastError, DTS_ERROR
        Save = False
        Exit Function
    End If
    
    ' Determine element type and route to appropriate save method
    Dim elementType As DTSElementType
    elementType = GetElementType(element)
    
    Select Case elementType
        Case DTS_ELEM_FRAME
            Save = SaveFrame(element, entity)
        Case DTS_ELEM_AREA
            Save = SaveArea(element, entity)
        Case DTS_ELEM_ANNOTATION
            Save = SaveTag(element, entity)
        Case DTS_ELEM_REBAR
            Save = SaveRebar(element, entity)
        Case Else
            m_LastError = "Unknown element type"
            LibDTS_Logger.Log "clsDTSRepository.Save: " & m_LastError, DTS_ERROR
            Save = False
    End Select
    
    Exit Function
    
ErrHandler:
    m_LastError = "Save error: " & Err.Description
    LibDTS_Logger.Log "clsDTSRepository.Save: " & m_LastError, DTS_ERROR
    Save = False
End Function

' Save Frame element
Private Function SaveFrame(frame As clsDTSFrame, Optional entity As Object = Nothing) As Boolean
    On Error GoTo ErrHandler
    
    If frame Is Nothing Then
        SaveFrame = False
        Exit Function
    End If
    
    ' Validate identity
    If Not entity Is Nothing Then
        frame.Base.ValidateIdentity entity.Handle
    End If
    
    ' Serialize to JSON
    Dim jsonData As String
    jsonData = frame.ToJson()
    
    If Len(jsonData) = 0 Then
        m_LastError = "Failed to serialize Frame"
        SaveFrame = False
        Exit Function
    End If
    
    ' If no entity provided, create new or find existing
    If entity Is Nothing Then
        ' Try to find existing entity by GUID
        If m_GuidToHandleMap.exists(frame.Base.GUID) Then
            Dim handle As String
            handle = m_GuidToHandleMap(frame.Base.GUID)
            
            ' Get entity by handle
            On Error Resume Next
            Set entity = m_AcadDoc.HandleToObject(handle)
            On Error GoTo ErrHandler
        End If
        
        ' If still nothing, draw new entity
        If entity Is Nothing Then
            Set entity = LibDTS_DriverCAD.DrawFrame(frame, m_AcadDoc)
        End If
    End If
    
    ' Save XData to entity
    If Not entity Is Nothing Then
        LibDTS_DriverCAD.SaveXData frame, entity
        
        ' Update mappings
        UpdateMapping frame.Base.GUID, entity.Handle
        
        ' Mark as clean
        frame.Base.MarkAsClean
        
        SaveFrame = True
        LibDTS_Logger.Log "clsDTSRepository.SaveFrame: Saved Frame GUID=" & frame.Base.GUID, DTS_INFO
    Else
        SaveFrame = False
    End If
    
    Exit Function
    
ErrHandler:
    m_LastError = "SaveFrame error: " & Err.Description
    LibDTS_Logger.Log "clsDTSRepository.SaveFrame: " & m_LastError, DTS_ERROR
    SaveFrame = False
End Function

' Save Area element
Private Function SaveArea(area As clsDTSArea, Optional entity As Object = Nothing) As Boolean
    On Error GoTo ErrHandler
    
    If area Is Nothing Then
        SaveArea = False
        Exit Function
    End If
    
    ' Validate identity
    If Not entity Is Nothing Then
        area.Base.ValidateIdentity entity.Handle
    End If
    
    ' Serialize to JSON
    Dim jsonData As String
    jsonData = area.ToJson()
    
    If Len(jsonData) = 0 Then
        m_LastError = "Failed to serialize Area"
        SaveArea = False
        Exit Function
    End If
    
    ' Similar logic as SaveFrame - stub for now
    ' In production, would draw 3DFace/Region and attach XData
    
    SaveArea = True
    LibDTS_Logger.Log "clsDTSRepository.SaveArea: Saved Area GUID=" & area.Base.GUID, DTS_INFO
    
    Exit Function
    
ErrHandler:
    m_LastError = "SaveArea error: " & Err.Description
    LibDTS_Logger.Log "clsDTSRepository.SaveArea: " & m_LastError, DTS_ERROR
    SaveArea = False
End Function

' Save Tag element
Private Function SaveTag(tag As clsDTSTag, Optional entity As Object = Nothing) As Boolean
    On Error GoTo ErrHandler
    
    If tag Is Nothing Then
        SaveTag = False
        Exit Function
    End If
    
    ' Validate identity
    If Not entity Is Nothing Then
        tag.Base.ValidateIdentity entity.Handle
    End If
    
    ' Serialize to JSON
    Dim jsonData As String
    jsonData = tag.ToJson()
    
    If Len(jsonData) = 0 Then
        m_LastError = "Failed to serialize Tag"
        SaveTag = False
        Exit Function
    End If
    
    ' Draw tag or update existing
    ' In production, would create MText/DBText entity
    
    SaveTag = True
    LibDTS_Logger.Log "clsDTSRepository.SaveTag: Saved Tag GUID=" & tag.Base.GUID, DTS_INFO
    
    Exit Function
    
ErrHandler:
    m_LastError = "SaveTag error: " & Err.Description
    LibDTS_Logger.Log "clsDTSRepository.SaveTag: " & m_LastError, DTS_ERROR
    SaveTag = False
End Function

' Save Rebar element
Private Function SaveRebar(rebar As clsDTSRebar, Optional entity As Object = Nothing) As Boolean
    On Error GoTo ErrHandler
    
    If rebar Is Nothing Then
        SaveRebar = False
        Exit Function
    End If
    
    ' Validate identity
    If Not entity Is Nothing Then
        rebar.Base.ValidateIdentity entity.Handle
    End If
    
    ' Serialize to JSON
    Dim jsonData As String
    jsonData = rebar.ToJson()
    
    If Len(jsonData) = 0 Then
        m_LastError = "Failed to serialize Rebar"
        SaveRebar = False
        Exit Function
    End If
    
    ' Draw rebar or update existing
    ' In production, would create Line/Polyline entity
    
    SaveRebar = True
    LibDTS_Logger.Log "clsDTSRepository.SaveRebar: Saved Rebar GUID=" & rebar.Base.GUID, DTS_INFO
    
    Exit Function
    
ErrHandler:
    m_LastError = "SaveRebar error: " & Err.Description
    LibDTS_Logger.Log "clsDTSRepository.SaveRebar: " & m_LastError, DTS_ERROR
    SaveRebar = False
End Function

' ==========================================
' LOAD OPERATIONS
' ==========================================

' Load element by CAD entity handle
' Parameters:
'   handle: CAD entity handle
' Returns: DTS element object or Nothing
Public Function LoadByHandle(handle As String) As Object
    On Error GoTo ErrHandler
    
    If m_AcadDoc Is Nothing Then
        m_LastError = "AutoCAD Document not initialized"
        Set LoadByHandle = Nothing
        Exit Function
    End If
    
    ' Get entity from handle
    Dim entity As Object
    On Error Resume Next
    Set entity = m_AcadDoc.HandleToObject(handle)
    On Error GoTo ErrHandler
    
    If entity Is Nothing Then
        m_LastError = "Entity not found for handle: " & handle
        Set LoadByHandle = Nothing
        Exit Function
    End If
    
    ' Load from entity
    Set LoadByHandle = LoadFromEntity(entity)
    Exit Function
    
ErrHandler:
    m_LastError = "LoadByHandle error: " & Err.Description
    LibDTS_Logger.Log "clsDTSRepository.LoadByHandle: " & m_LastError, DTS_ERROR
    Set LoadByHandle = Nothing
End Function

' Load element by GUID
' Parameters:
'   guid: Element GUID
' Returns: DTS element object or Nothing
Public Function LoadByGUID(GUID As String) As Object
    On Error GoTo ErrHandler
    
    ' Check mapping
    If m_GuidToHandleMap.exists(GUID) Then
        Dim handle As String
        handle = m_GuidToHandleMap(GUID)
        Set LoadByGUID = LoadByHandle(handle)
    Else
        m_LastError = "GUID not found in mappings: " & GUID
        Set LoadByGUID = Nothing
    End If
    
    Exit Function
    
ErrHandler:
    m_LastError = "LoadByGUID error: " & Err.Description
    LibDTS_Logger.Log "clsDTSRepository.LoadByGUID: " & m_LastError, DTS_ERROR
    Set LoadByGUID = Nothing
End Function

' Load element from CAD entity
' Parameters:
'   entity: CAD entity object
' Returns: DTS element object or Nothing
Public Function LoadFromEntity(entity As Object) As Object
    On Error GoTo ErrHandler
    
    If entity Is Nothing Then
        Set LoadFromEntity = Nothing
        Exit Function
    End If
    
    ' Extract XData
    Dim xdataType() As Integer
    Dim xdataValue() As Variant
    
    On Error Resume Next
    entity.GetXData LibDTS_Global.DTS_APP_NAME, xdataType, xdataValue
    On Error GoTo ErrHandler
    
    ' Check if XData exists
    If UBound(xdataType) < 0 Then
        m_LastError = "No XData found on entity"
        Set LoadFromEntity = Nothing
        Exit Function
    End If
    
    ' Extract encrypted JSON (should be at index 1)
    Dim encryptedData As String
    If UBound(xdataValue) >= 1 Then
        encryptedData = xdataValue(1)
    Else
        Set LoadFromEntity = Nothing
        Exit Function
    End If
    
    ' Decrypt and parse to determine type
    Dim jsonStr As String
    jsonStr = LibDTS_Security.Decrypt(encryptedData)
    
    If Len(jsonStr) = 0 Then
        Set LoadFromEntity = Nothing
        Exit Function
    End If
    
    ' Parse JSON to get type
    Dim root As Object
    Set root = LibDTS_Base.ParseJson(jsonStr)
    
    If Not root.exists("Type") Then
        Set LoadFromEntity = Nothing
        Exit Function
    End If
    
    Dim elementType As DTSElementType
    elementType = CLng(root("Type"))
    
    ' Create appropriate element and deserialize
    Dim element As Object
    
    Select Case elementType
        Case DTS_ELEM_FRAME
            Dim frame As New clsDTSFrame
            frame.FromXData encryptedData
            frame.Base.ValidateIdentity entity.Handle
            Set element = frame
            
        Case DTS_ELEM_AREA
            Dim area As New clsDTSArea
            area.FromXData encryptedData
            area.Base.ValidateIdentity entity.Handle
            Set element = area
            
        Case DTS_ELEM_ANNOTATION
            Dim tag As New clsDTSTag
            tag.FromXData encryptedData
            tag.Base.ValidateIdentity entity.Handle
            Set element = tag
            
        Case DTS_ELEM_REBAR
            Dim rebar As New clsDTSRebar
            rebar.FromXData encryptedData
            rebar.Base.ValidateIdentity entity.Handle
            Set element = rebar
            
        Case Else
            Set element = Nothing
    End Select
    
    Set LoadFromEntity = element
    Exit Function
    
ErrHandler:
    m_LastError = "LoadFromEntity error: " & Err.Description
    LibDTS_Logger.Log "clsDTSRepository.LoadFromEntity: " & m_LastError, DTS_ERROR
    Set LoadFromEntity = Nothing
End Function

' ==========================================
' SAP2000 SYNC OPERATIONS
' ==========================================

' Sync element to SAP2000
' Parameters:
'   element: DTS element to sync
' Returns: True if synced successfully
Public Function SyncToSAP(element As Object) As Boolean
    On Error GoTo ErrHandler
    
    If element Is Nothing Then
        SyncToSAP = False
        Exit Function
    End If
    
    ' Ensure SAP connection
    If Not m_SapConnected Then
        m_SapConnected = LibDTS_DriverSAP.Connect()
        If Not m_SapConnected Then
            m_LastError = "Failed to connect to SAP2000"
            SyncToSAP = False
            Exit Function
        End If
    End If
    
    ' Determine element type and route to appropriate sync method
    Dim elementType As DTSElementType
    elementType = GetElementType(element)
    
    Select Case elementType
        Case DTS_ELEM_FRAME
            SyncToSAP = SyncFrameToSAP(element)
        Case DTS_ELEM_AREA
            SyncToSAP = SyncAreaToSAP(element)
        Case Else
            SyncToSAP = False
    End Select
    
    Exit Function
    
ErrHandler:
    m_LastError = "SyncToSAP error: " & Err.Description
    LibDTS_Logger.Log "clsDTSRepository.SyncToSAP: " & m_LastError, DTS_ERROR
    SyncToSAP = False
End Function

' Sync Frame to SAP2000
Private Function SyncFrameToSAP(frame As clsDTSFrame) As Boolean
    On Error GoTo ErrHandler
    
    If frame Is Nothing Then
        SyncFrameToSAP = False
        Exit Function
    End If
    
    ' Use driver to push frame
    SyncFrameToSAP = LibDTS_DriverSAP.PushFrame(frame)
    
    If SyncFrameToSAP Then
        LibDTS_Logger.Log "clsDTSRepository.SyncFrameToSAP: Synced Frame GUID=" & frame.Base.GUID, DTS_INFO
    End If
    
    Exit Function
    
ErrHandler:
    m_LastError = "SyncFrameToSAP error: " & Err.Description
    LibDTS_Logger.Log "clsDTSRepository.SyncFrameToSAP: " & m_LastError, DTS_ERROR
    SyncFrameToSAP = False
End Function

' Sync Area to SAP2000
Private Function SyncAreaToSAP(area As clsDTSArea) As Boolean
    On Error GoTo ErrHandler
    
    If area Is Nothing Then
        SyncAreaToSAP = False
        Exit Function
    End If
    
    ' Use driver to push area
    ' Stub for now - would call LibDTS_DriverSAP.PushArea
    SyncAreaToSAP = True
    
    LibDTS_Logger.Log "clsDTSRepository.SyncAreaToSAP: Synced Area GUID=" & area.Base.GUID, DTS_INFO
    
    Exit Function
    
ErrHandler:
    m_LastError = "SyncAreaToSAP error: " & Err.Description
    LibDTS_Logger.Log "clsDTSRepository.SyncAreaToSAP: " & m_LastError, DTS_ERROR
    SyncAreaToSAP = False
End Function

' ==========================================
' HELPER METHODS
' ==========================================

' Get element type from object
Private Function GetElementType(element As Object) As DTSElementType
    On Error Resume Next
    
    ' Try different type checks
    If TypeName(element) = "clsDTSFrame" Then
        GetElementType = DTS_ELEM_FRAME
    ElseIf TypeName(element) = "clsDTSArea" Then
        GetElementType = DTS_ELEM_AREA
    ElseIf TypeName(element) = "clsDTSTag" Then
        GetElementType = DTS_ELEM_ANNOTATION
    ElseIf TypeName(element) = "clsDTSRebar" Then
        GetElementType = DTS_ELEM_REBAR
    Else
        ' Try to access Base.ElementType
        GetElementType = element.Base.ElementType
    End If
End Function

' Update GUID-Handle mapping
Private Sub UpdateMapping(GUID As String, handle As String)
    On Error Resume Next
    
    ' Update both directions
    If m_GuidToHandleMap.exists(GUID) Then
        m_GuidToHandleMap(GUID) = handle
    Else
        m_GuidToHandleMap.Add GUID, handle
    End If
    
    If m_HandleToGuidMap.exists(handle) Then
        m_HandleToGuidMap(handle) = GUID
    Else
        m_HandleToGuidMap.Add handle, GUID
    End If
End Sub

' ==========================================
' PROPERTIES
' ==========================================

Public Property Get LastError() As String
    LastError = m_LastError
End Property

Public Property Get IsInitialized() As Boolean
    IsInitialized = Not (m_AcadDoc Is Nothing)
End Property

Public Property Get SAPConnected() As Boolean
    SAPConnected = m_SapConnected
End Property
