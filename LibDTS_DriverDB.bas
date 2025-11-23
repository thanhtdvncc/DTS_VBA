Attribute VB_Name = "LibDTS_DriverDB"
' Module: LibDTS_DriverDB
' Purpose: Canonical driver for database and persistent storage operations
' Features: Settings management, GUID mapping persistence, configuration storage
' Version: 2.0 (Enhanced with GUID mapping and validation)
Option Explicit

' ==========================================
' MODULE-LEVEL VARIABLES
' ==========================================
Private m_LastError As String        ' Last error message
Private m_SettingsPath As String     ' Path to settings file
Private m_MappingPath As String      ' Path to GUID mapping file

' ==========================================
' PUBLIC CONSTANTS
' ==========================================
Public Const DRIVER_NAME As String = "LibDTS_DriverDB"
Public Const SETTINGS_FILE As String = "settings.json"
Public Const MAPPING_FILE As String = "guid_mapping.json"

' ==========================================
' 1. SETTINGS MANAGEMENT
' ==========================================

' Load settings from JSON file
' Returns: Dictionary object with settings or empty Dictionary on failure
Public Function LoadSettings() As Object
    On Error GoTo ErrHandler
    
    Dim fso As Object, ts As Object
    Dim jsonStr As String
    Dim path As String
    
    path = GetSettingsPath()
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FileExists(path) Then
        Set ts = fso.OpenTextFile(path, 1) ' ForReading
        jsonStr = ts.ReadAll
        ts.Close
        
        ' Parse JSON
        Set LoadSettings = LibDTS_Base.ParseJson(jsonStr)
        LibDTS_Logger.Log DRIVER_NAME & ".LoadSettings: Loaded settings from " & path, DTS_INFO
    Else
        ' Return empty dictionary if no file
        Set LoadSettings = CreateObject("Scripting.Dictionary")
        LibDTS_Logger.Log DRIVER_NAME & ".LoadSettings: No settings file found, returning empty dictionary", DTS_INFO
    End If
    
    Exit Function
    
ErrHandler:
    m_LastError = "LoadSettings error: " & err.description
    LibDTS_Logger.Log DRIVER_NAME & ".LoadSettings: " & m_LastError, DTS_ERROR
    Set LoadSettings = CreateObject("Scripting.Dictionary")
End Function

' Save settings to JSON file
' Parameters:
'   settingsDict: Dictionary object with settings
Public Sub SaveSettings(settingsDict As Object)
    On Error GoTo ErrHandler
    
    If settingsDict Is Nothing Then
        m_LastError = "Settings dictionary is Nothing"
        LibDTS_Logger.Log DRIVER_NAME & ".SaveSettings: " & m_LastError, DTS_ERROR
        Exit Sub
    End If
    
    Dim fso As Object, ts As Object
    Dim jsonStr As String
    Dim path As String
    
    path = GetSettingsPath()
    jsonStr = LibDTS_Base.ToJson(settingsDict)
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Create folder if not exists
    Dim folder As String
    folder = fso.GetParentFolderName(path)
    If Not fso.FolderExists(folder) Then
        fso.CreateFolder folder
    End If
    
    ' Write file
    Set ts = fso.CreateTextFile(path, True) ' Overwrite
    ts.Write jsonStr
    ts.Close
    
    LibDTS_Logger.Log DRIVER_NAME & ".SaveSettings: Saved settings to " & path, DTS_INFO
    Exit Sub
    
ErrHandler:
    m_LastError = "SaveSettings error: " & err.description
    LibDTS_Logger.Log DRIVER_NAME & ".SaveSettings: " & m_LastError, DTS_ERROR
End Sub

' ==========================================
' 2. GUID MAPPING PERSISTENCE
' ==========================================

' Load GUID mapping from persistent storage
' Returns: Dictionary object with GUID -> Element info mappings
Public Function LoadGUIDMapping() As Object
    On Error GoTo ErrHandler
    
    Dim fso As Object, ts As Object
    Dim jsonStr As String
    Dim path As String
    
    path = GetMappingPath()
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FileExists(path) Then
        Set ts = fso.OpenTextFile(path, 1) ' ForReading
        jsonStr = ts.ReadAll
        ts.Close
        
        ' Parse JSON
        Set LoadGUIDMapping = LibDTS_Base.ParseJson(jsonStr)
        LibDTS_Logger.Log DRIVER_NAME & ".LoadGUIDMapping: Loaded " & LoadGUIDMapping.count & " mappings from " & path, DTS_INFO
    Else
        ' Return empty dictionary if no file
        Set LoadGUIDMapping = CreateObject("Scripting.Dictionary")
        LibDTS_Logger.Log DRIVER_NAME & ".LoadGUIDMapping: No mapping file found, returning empty dictionary", DTS_INFO
    End If
    
    Exit Function
    
ErrHandler:
    m_LastError = "LoadGUIDMapping error: " & err.description
    LibDTS_Logger.Log DRIVER_NAME & ".LoadGUIDMapping: " & m_LastError, DTS_ERROR
    Set LoadGUIDMapping = CreateObject("Scripting.Dictionary")
End Function

' Save GUID mapping to persistent storage
' Parameters:
'   mappingDict: Dictionary object with GUID -> Element info mappings
Public Sub SaveGUIDMapping(mappingDict As Object)
    On Error GoTo ErrHandler
    
    If mappingDict Is Nothing Then
        m_LastError = "Mapping dictionary is Nothing"
        LibDTS_Logger.Log DRIVER_NAME & ".SaveGUIDMapping: " & m_LastError, DTS_ERROR
        Exit Sub
    End If
    
    Dim fso As Object, ts As Object
    Dim jsonStr As String
    Dim path As String
    
    path = GetMappingPath()
    jsonStr = LibDTS_Base.ToJson(mappingDict)
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Create folder if not exists
    Dim folder As String
    folder = fso.GetParentFolderName(path)
    If Not fso.FolderExists(folder) Then
        fso.CreateFolder folder
    End If
    
    ' Write file
    Set ts = fso.CreateTextFile(path, True) ' Overwrite
    ts.Write jsonStr
    ts.Close
    
    LibDTS_Logger.Log DRIVER_NAME & ".SaveGUIDMapping: Saved " & mappingDict.count & " mappings to " & path, DTS_INFO
    Exit Sub
    
ErrHandler:
    m_LastError = "SaveGUIDMapping error: " & err.description
    LibDTS_Logger.Log DRIVER_NAME & ".SaveGUIDMapping: " & m_LastError, DTS_ERROR
End Sub

' Get mapped element information by GUID
' Parameters:
'   guid: GUID string to look up
' Returns: Variant (Array with element info) or Empty if not found
Public Function GetMappedElement(guid As String) As Variant
    On Error GoTo ErrHandler
    
    ' Load current mapping
    Dim mapping As Object
    Set mapping = LoadGUIDMapping()
    
    If mapping.Exists(guid) Then
        GetMappedElement = mapping(guid)
        LibDTS_Logger.Log DRIVER_NAME & ".GetMappedElement: Found mapping for GUID " & guid, DTS_INFO
    Else
        GetMappedElement = Empty
        LibDTS_Logger.Log DRIVER_NAME & ".GetMappedElement: No mapping found for GUID " & guid, DTS_WARNING
    End If
    
    Exit Function
    
ErrHandler:
    m_LastError = "GetMappedElement error: " & err.description
    LibDTS_Logger.Log DRIVER_NAME & ".GetMappedElement: " & m_LastError, DTS_ERROR
    GetMappedElement = Empty
End Function

' Set mapped element information for GUID
' Parameters:
'   guid: GUID string
'   elementInfo: Variant (typically Array with element type, name, etc.)
Public Sub SetMappedElement(guid As String, elementInfo As Variant)
    On Error GoTo ErrHandler
    
    If Not LibDTS_Base.IsValidGUID(guid) Then
        m_LastError = "Invalid GUID format: " & guid
        LibDTS_Logger.Log DRIVER_NAME & ".SetMappedElement: " & m_LastError, DTS_ERROR
        Exit Sub
    End If
    
    ' Load current mapping
    Dim mapping As Object
    Set mapping = LoadGUIDMapping()
    
    ' Add or update mapping
    If mapping.Exists(guid) Then
        mapping(guid) = elementInfo
        LibDTS_Logger.Log DRIVER_NAME & ".SetMappedElement: Updated mapping for GUID " & guid, DTS_INFO
    Else
        mapping.Add guid, elementInfo
        LibDTS_Logger.Log DRIVER_NAME & ".SetMappedElement: Created mapping for GUID " & guid, DTS_INFO
    End If
    
    ' Save back to file
    SaveGUIDMapping mapping
    
    Exit Sub
    
ErrHandler:
    m_LastError = "SetMappedElement error: " & err.description
    LibDTS_Logger.Log DRIVER_NAME & ".SetMappedElement: " & m_LastError, DTS_ERROR
End Sub

' ==========================================
' 3. MAPPING VALIDATION & REPAIR
' ==========================================

' Validate mapping integrity (check for orphaned or invalid entries)
' Returns: Collection of problem GUIDs with descriptions
Public Function ValidateMappingIntegrity() As Collection
    On Error GoTo ErrHandler
    
    Dim problems As New Collection
    
    ' Load mapping
    Dim mapping As Object
    Set mapping = LoadGUIDMapping()
    
    ' Check each mapping
    Dim guid As Variant
    For Each guid In mapping.keys
        ' Check GUID format
        If Not LibDTS_Base.IsValidGUID(CStr(guid)) Then
            problems.Add "Invalid GUID format: " & guid
        End If
        
        ' Check element info
        Dim elementInfo As Variant
        elementInfo = mapping(guid)
        
        If IsEmpty(elementInfo) Then
            problems.Add "Empty element info for GUID: " & guid
        ElseIf Not IsArray(elementInfo) Then
            problems.Add "Element info is not an array for GUID: " & guid
        ElseIf IsArray(elementInfo) Then
            If UBound(elementInfo) < 1 Then
                problems.Add "Insufficient element info for GUID: " & guid
            End If
        End If
    Next guid
    
    If problems.count > 0 Then
        LibDTS_Logger.Log DRIVER_NAME & ".ValidateMappingIntegrity: Found " & problems.count & " problems", DTS_WARNING
    Else
        LibDTS_Logger.Log DRIVER_NAME & ".ValidateMappingIntegrity: No problems found", DTS_INFO
    End If
    
    Set ValidateMappingIntegrity = problems
    Exit Function
    
ErrHandler:
    m_LastError = "ValidateMappingIntegrity error: " & err.description
    LibDTS_Logger.Log DRIVER_NAME & ".ValidateMappingIntegrity: " & m_LastError, DTS_ERROR
    Set ValidateMappingIntegrity = New Collection
End Function

' Repair mapping by removing invalid entries
' Parameters:
'   guidList: Variant (Array of GUIDs to remove)
' Returns: Number of entries removed
Public Function RepairMapping(guidList As Variant) As Long
    On Error GoTo ErrHandler
    
    If Not IsArray(guidList) Then
        m_LastError = "guidList is not an array"
        LibDTS_Logger.Log DRIVER_NAME & ".RepairMapping: " & m_LastError, DTS_ERROR
        RepairMapping = 0
        Exit Function
    End If
    
    ' Load mapping
    Dim mapping As Object
    Set mapping = LoadGUIDMapping()
    
    ' Remove specified GUIDs
    Dim removed As Long
    removed = 0
    
    Dim guid As Variant
    For Each guid In guidList
        If mapping.Exists(CStr(guid)) Then
            mapping.Remove CStr(guid)
            removed = removed + 1
        End If
    Next guid
    
    ' Save repaired mapping
    If removed > 0 Then
        SaveGUIDMapping mapping
        LibDTS_Logger.Log DRIVER_NAME & ".RepairMapping: Removed " & removed & " invalid entries", DTS_INFO
    End If
    
    RepairMapping = removed
    Exit Function
    
ErrHandler:
    m_LastError = "RepairMapping error: " & err.description
    LibDTS_Logger.Log DRIVER_NAME & ".RepairMapping: " & m_LastError, DTS_ERROR
    RepairMapping = 0
End Function

' ==========================================
' 4. UTILITY FUNCTIONS
' ==========================================

' Get last error message
' Returns: String describing last error
Public Function GetLastError() As String
    GetLastError = m_LastError
    m_LastError = "" ' Clear after reading
End Function

' Get full path to settings file
' Returns: Full path string
Private Function GetSettingsPath() As String
    If Len(m_SettingsPath) = 0 Then
        m_SettingsPath = Environ("APPDATA") & "\DTS_Core\" & SETTINGS_FILE
    End If
    GetSettingsPath = m_SettingsPath
End Function

' Get full path to mapping file
' Returns: Full path string
Private Function GetMappingPath() As String
    If Len(m_MappingPath) = 0 Then
        m_MappingPath = Environ("APPDATA") & "\DTS_Core\" & MAPPING_FILE
    End If
    GetMappingPath = m_MappingPath
End Function

' Export mapping to Excel (for debugging/backup)
' Parameters:
'   ws: Excel Worksheet object
Public Sub ExportMappingToExcel(ws As Object)
    On Error GoTo ErrHandler
    
    If ws Is Nothing Then
        m_LastError = "Worksheet is Nothing"
        LibDTS_Logger.Log DRIVER_NAME & ".ExportMappingToExcel: " & m_LastError, DTS_ERROR
        Exit Sub
    End If
    
    ' Load mapping
    Dim mapping As Object
    Set mapping = LoadGUIDMapping()
    
    ' Write headers
    ws.Cells(1, 1).value = "GUID"
    ws.Cells(1, 2).value = "Element Type"
    ws.Cells(1, 3).value = "Element Name"
    ws.Cells(1, 4).value = "Additional Info"
    
    ' Write data
    Dim row As Long
    row = 2
    
    Dim guid As Variant
    For Each guid In mapping.keys
        ws.Cells(row, 1).value = CStr(guid)
        
        Dim elementInfo As Variant
        elementInfo = mapping(guid)
        
        If IsArray(elementInfo) Then
            If UBound(elementInfo) >= 0 Then ws.Cells(row, 2).value = CStr(elementInfo(0))
            If UBound(elementInfo) >= 1 Then ws.Cells(row, 3).value = CStr(elementInfo(1))
            If UBound(elementInfo) >= 2 Then ws.Cells(row, 4).value = Join(elementInfo, ", ")
        End If
        
        row = row + 1
    Next guid
    
    ' Auto-fit columns
    ws.Columns("A:D").AutoFit
    
    LibDTS_Logger.Log DRIVER_NAME & ".ExportMappingToExcel: Exported " & mapping.count & " mappings to Excel", DTS_INFO
    Exit Sub
    
ErrHandler:
    m_LastError = "ExportMappingToExcel error: " & err.description
    LibDTS_Logger.Log DRIVER_NAME & ".ExportMappingToExcel: " & m_LastError, DTS_ERROR
End Sub
