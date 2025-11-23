Attribute VB_Name = "Core_Event_Handler"

Option Explicit
'===============================================================
' Core Module: Core_Event_Handler (UPDATED)
' Purpose: Real-time synchronization event handling (event-driven)
' Notes:
'  - Removed SAP polling. SAP is updated immediately after CAD creation.
'  - Track CAD handle -> SAP(type|name) mapping in-session to support deletes.
'  - When enabling real-time sync, rebuild mapping from the log sheet "CADtoSAP_Map".
'  - Added RegisterCADtoSAPMapping public API for other modules to register mappings.
'===============================================================

Private Const POLL_INTERVAL_SEC As Long = 2 ' 2 seconds (CAD scan loop)

' State management
Private mSyncEnabled As Boolean
Private mAcadDoc As Object
Private mSapModel As Object

' Debounce flag to prevent infinite loop
Private mSyncInProgress As Boolean

' CAD entity tracking
Private mCADEntityHandles As Object ' Dictionary of handles -> True
' Mapping of CAD handle -> "Type|SAPName" (e.g., "Point|N101" or "Frame|F23")
Private mCADEntityToSAPMap As Object

' Log sheet name (must match Core_Sync_Manager)
Private Const LOG_SHEET_NAME As String = "CADtoSAP_Map"

' ----------------------------
' Public API
' ----------------------------

Public Sub EnableRealTimeSync(acadDoc As Object, SapModel As Object)
    On Error Resume Next
    
    Set mAcadDoc = acadDoc
    Set mSapModel = SapModel
    mSyncEnabled = True
    mSyncInProgress = False
    
    ' Initialize tracking
    Set mCADEntityHandles = CreateObject("Scripting.Dictionary")
    Set mCADEntityToSAPMap = CreateObject("Scripting.Dictionary")
    UpdateCADTracking
    
    ' Rebuild mapping from previous session log sheet (if exists)
    RebuildMappingFromLogSheet
    
    ' Enable auto-sync mode in Core_Sync_Manager
    On Error Resume Next
    Core_Sync_Manager.SetAutoSyncMode True
    On Error GoTo 0
    
    Core_Sync_Manager.LogStatus "Real-time sync ENABLED (auto mode)"
    Core_Sync_Manager.LogStatus "Excel is bridge only - all changes auto-synced"
    
    ' Start monitoring loop (CAD only)
    Application.OnTime Now + TimeSerial(0, 0, POLL_INTERVAL_SEC), "Core_Event_Handler.MonitorChanges"
    
    On Error GoTo 0
End Sub

Public Sub DisableRealTimeSync()
    mSyncEnabled = False
    Set mAcadDoc = Nothing
    Set mSapModel = Nothing
    If Not mCADEntityHandles Is Nothing Then Set mCADEntityHandles = Nothing
    If Not mCADEntityToSAPMap Is Nothing Then Set mCADEntityToSAPMap = Nothing
    
    ' Disable auto-sync mode
    On Error Resume Next
    Core_Sync_Manager.SetAutoSyncMode False
    On Error GoTo 0
    
    Core_Sync_Manager.LogStatus "Real-time sync DISABLED"
End Sub

Public Function IsSyncEnabled() As Boolean
    IsSyncEnabled = mSyncEnabled
End Function

' Allow other modules to register mapping (called from Core_Sync_Manager when auto-creating)
' handle: CAD entity handle
' sapType: "Point" / "Frame" / "Area"
' sapName: name in SAP2000
Public Sub RegisterCADtoSAPMapping(Handle As String, sapType As String, sapName As String)
    On Error Resume Next
    If mCADEntityToSAPMap Is Nothing Then
        Set mCADEntityToSAPMap = CreateObject("Scripting.Dictionary")
    End If
    If Len(Trim$(Handle)) = 0 Then Exit Sub
    mCADEntityToSAPMap(CStr(Handle)) = sapType & "|" & sapName
    
    ' Also append to log sheet for persistence across sessions
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(LOG_SHEET_NAME)
    If Not ws Is Nothing Then
        Dim NextRow As Long
        NextRow = ws.Cells(ws.rows.count, "A").End(xlUp).row + 1
        ws.Cells(NextRow, 1).Value = Handle
        ws.Cells(NextRow, 2).Value = "AutoMapping"
        ws.Cells(NextRow, 3).Value = "" ' layer unknown
        ws.Cells(NextRow, 4).Value = sapType
        ws.Cells(NextRow, 5).Value = sapName
        ws.Cells(NextRow, 6).Value = "Registered by RegisterCADtoSAPMapping"
    End If
    On Error GoTo 0
End Sub

' Optional public accessor (read-only) - returns "Type|Name" or empty
Public Function GetSAPMappingForCADHandle(Handle As String) As String
    On Error Resume Next
    GetSAPMappingForCADHandle = ""
    If mCADEntityToSAPMap Is Nothing Then Exit Function
    If mCADEntityToSAPMap.exists(CStr(Handle)) Then
        GetSAPMappingForCADHandle = CStr(mCADEntityToSAPMap(CStr(Handle)))
    End If
    On Error GoTo 0
End Function

' ----------------------------
' Monitoring Loop (CAD only)
' ----------------------------

Public Sub MonitorChanges()
    On Error Resume Next
    
    If Not mSyncEnabled Then Exit Sub
    If mSyncInProgress Then GoTo ScheduleNext
    
    ' Check for CAD changes (new + deleted)
    Dim newHandles As Collection
    Dim deletedHandles As Collection
    Set newHandles = New Collection
    Set deletedHandles = New Collection
    
    mSyncInProgress = True
    If CheckCADChanges(newHandles, deletedHandles) Then
        If newHandles.count > 0 Then
            Core_Sync_Manager.LogStatus "[AUTO-SYNC] CAD additions detected: " & newHandles.count
            SyncCADChangesToSAP newHandles
        End If
        
        If deletedHandles.count > 0 Then
            Core_Sync_Manager.LogStatus "[AUTO-SYNC] CAD deletions detected: " & deletedHandles.count
            SyncDeletionsFromCAD deletedHandles
        End If
        
        ' After processing, update tracking
        UpdateCADTracking
    End If
    mSyncInProgress = False
    
ScheduleNext:
    ' Schedule next check
    If mSyncEnabled Then
        Application.OnTime Now + TimeSerial(0, 0, POLL_INTERVAL_SEC), "Core_Event_Handler.MonitorChanges"
    End If
    
    On Error GoTo 0
End Sub

' ----------------------------
' Change Detection
' ----------------------------

' Returns True if any change (new or deleted) found.
Private Function CheckCADChanges(ByRef newHandles As Collection, ByRef deletedHandles As Collection) As Boolean
    On Error Resume Next
    
    CheckCADChanges = False
    If mAcadDoc Is Nothing Then Exit Function
    
    Dim ms As Object
    Set ms = mAcadDoc.ModelSpace
    
    Dim currentHandles As Object
    Set currentHandles = CreateObject("Scripting.Dictionary")
    
    Dim ent As Object
    For Each ent In ms
        Dim h As String
        h = CStr(ent.Handle)
        If Not currentHandles.exists(h) Then
            currentHandles.Add h, True
        End If
        
        ' If new entity
        If Not mCADEntityHandles.exists(h) Then
            newHandles.Add h
            CheckCADChanges = True
        End If
    Next ent
    
    ' Detect deletions: keys previously present but not in current
    Dim prevKey As Variant
    For Each prevKey In mCADEntityHandles.keys
        If Not currentHandles.exists(prevKey) Then
            deletedHandles.Add prevKey
            CheckCADChanges = True
        End If
    Next prevKey
    
    On Error GoTo 0
End Function

Private Sub UpdateCADTracking()
    On Error Resume Next
    
    If mAcadDoc Is Nothing Then Exit Sub
    
    Set mCADEntityHandles = CreateObject("Scripting.Dictionary")
    
    Dim ms As Object
    Set ms = mAcadDoc.ModelSpace
    
    Dim ent As Object
    For Each ent In ms
        Dim h As String
        h = CStr(ent.Handle)
        If Not mCADEntityHandles.exists(h) Then
            mCADEntityHandles.Add h, True
        End If
    Next ent
    
    On Error GoTo 0
End Sub

' ----------------------------
' Sync Actions (SILENT MODE - NO DIALOGS)
' ----------------------------

Private Sub SyncCADChangesToSAP(newHandles As Collection)
    On Error Resume Next
    
    If newHandles.count = 0 Then Exit Sub
    If mSapModel Is Nothing Then
        Core_Sync_Manager.LogStatus "[AUTO-SYNC] SAP model not available - cannot import CAD entities."
        Exit Sub
    End If
    
    Dim h As Variant
    For Each h In newHandles
        ' Find entity in ModelSpace by handle
        Dim ent As Object
        Set ent = FindCADEntityByHandle(CStr(h))
        
        ' If ent is Nothing then skip processing this handle
        If Not ent Is Nothing Then
            ' Try to extract the intended SAP name/type from XData (if present)
            Dim sapType As String, sapName As String
            sapType = ""
            sapName = ""
            GetSAPNameAndTypeFromCADEntity ent, sapType, sapName
            
            ' Call existing import routine (silent)
            On Error Resume Next
            Core_Sync_Manager.ImportCADEntityToSAP_Silent mAcadDoc, mSapModel, CStr(h), 1#, 1#
            On Error GoTo 0
            
            ' If we were able to read a name from XData, persist mapping
            If Len(sapName) > 0 Then
                On Error Resume Next
                If mCADEntityToSAPMap Is Nothing Then Set mCADEntityToSAPMap = CreateObject("Scripting.Dictionary")
                mCADEntityToSAPMap(CStr(h)) = sapType & "|" & sapName
                On Error GoTo 0
            End If
            
            ' After creating in SAP, refresh SAP view immediately
            On Error Resume Next
            Dim ret As Long
            mSapModel.View.RefreshView
            On Error GoTo 0
        End If
    Next h
    
    Core_Sync_Manager.LogStatus "[AUTO-SYNC] CAD->SAP completed: " & newHandles.count & " entities (attempted)"
    
    On Error GoTo 0
End Sub

Private Sub SyncDeletionsFromCAD(deletedHandles As Collection)
    On Error Resume Next
    
    If deletedHandles.count = 0 Then Exit Sub
    If mSapModel Is Nothing Then
        Core_Sync_Manager.LogStatus "[AUTO-SYNC] SAP model not available - cannot delete SAP entities."
        Exit Sub
    End If
    
    Dim h As Variant
    For Each h In deletedHandles
        Dim key As String
        key = CStr(h)
        If mCADEntityToSAPMap Is Nothing Then
            ' nothing to do
        ElseIf mCADEntityToSAPMap.exists(key) Then
            Dim val As String
            val = CStr(mCADEntityToSAPMap(key))
            
            Dim sepPos As Long
            sepPos = InStr(val, "|")
            Dim sapType As String, sapName As String
            If sepPos > 0 Then
                sapType = Left$(val, sepPos - 1)
                sapName = mid$(val, sepPos + 1)
            Else
                sapType = ""
                sapName = val
            End If
            
            ' Call SAP API delete based on type
            Dim ret As Long
            On Error Resume Next
            Select Case LCase$(sapType)
                Case "point"
                    ret = mSapModel.pointObj.Delete(sapName)
                Case "frame"
                    ret = mSapModel.frameObj.Delete(sapName)
                Case "area"
                    ret = mSapModel.AreaObj.Delete(sapName)
                Case Else
                    ' Try frame delete then point then area (best-effort)
                    ret = mSapModel.frameObj.Delete(sapName)
                    If ret <> 0 Then ret = mSapModel.pointObj.Delete(sapName)
                    If ret <> 0 Then ret = mSapModel.AreaObj.Delete(sapName)
            End Select
            On Error GoTo 0
            
            ' Remove mapping if deletion attempted
            On Error Resume Next
            mCADEntityToSAPMap.Remove key
            On Error GoTo 0
        Else
            ' no mapping stored: cannot determine SAP object to delete
            Core_Sync_Manager.LogStatus "[AUTO-SYNC] No SAP mapping for CAD handle " & key & " - skipping delete"
        End If
    Next h
    
    ' Refresh view after deletions
    On Error Resume Next
    Dim r2 As Long
    r2 = mSapModel.View.RefreshView
    On Error GoTo 0
    
    Core_Sync_Manager.LogStatus "[AUTO-SYNC] CAD->SAP deletes completed: " & deletedHandles.count & " entities (attempted)"
End Sub

' ----------------------------
' Helper functions
' ----------------------------

Private Function FindCADEntityByHandle(Handle As String) As Object
    On Error Resume Next
    Set FindCADEntityByHandle = Nothing

    ' Sanity checks
    If mAcadDoc Is Nothing Then Exit Function
    If Len(Trim$(Handle)) = 0 Then Exit Function

    ' Use AutoCAD's HandleToObject which is direct and reliable
    Dim ent As Object
    Set ent = mAcadDoc.HandleToObject(CStr(Handle))

    If Not ent Is Nothing Then
        Set FindCADEntityByHandle = ent
    Else
        ' If HandleToObject failed (object not found), return Nothing
        Set FindCADEntityByHandle = Nothing
    End If

    On Error GoTo 0
End Function

' Extract SAP element name and type from CAD entity's XData (if possible)
' sapType: "Point" | "Frame" | "Area" or "" if unknown
' sapName: the name string used for SAP element (if found)
Private Sub GetSAPNameAndTypeFromCADEntity(ent As Object, ByRef sapType As String, ByRef sapName As String)
    On Error Resume Next
    sapType = ""
    sapName = ""
    Dim xdType As Variant, xdVal As Variant
    ent.GetXData "DTS_SAP2000", xdType, xdVal
    If Not IsEmpty(xdVal) And IsArray(xdVal) Then
        Dim entType As String
        entType = LCase$(TypeName(ent))
        If InStr(entType, "circle") > 0 Then
            ' point: expected format [APP, nodeName, X, Y, Z, ...]
            If UBound(xdVal) >= 1 Then
                sapName = CStr(xdVal(1))
                sapType = "Point"
            End If
        ElseIf InStr(entType, "line") > 0 Or InStr(entType, "polyline") > 0 Then
            ' frame: expected [APP, frameName, Point1, Point2, Section]
            If UBound(xdVal) >= 1 Then
                sapName = CStr(xdVal(1))
                sapType = "Frame"
            End If
        ElseIf InStr(entType, "lwpolyline") > 0 Or InStr(entType, "hatch") > 0 Then
            ' area: expected [APP, areaName, sectionName, pointList]
            If UBound(xdVal) >= 1 Then
                sapName = CStr(xdVal(1))
                sapType = "Area"
            End If
        Else
            ' fallback: try first element
            If UBound(xdVal) >= 1 Then
                sapName = CStr(xdVal(1))
            End If
        End If
    End If
    On Error GoTo 0
End Sub

' Rebuild mapping dictionary from log sheet "CADtoSAP_Map" when enabling sync.
' This allows deletes in CAD to find SAP names to delete.
Private Sub RebuildMappingFromLogSheet()
    On Error Resume Next
    If mCADEntityToSAPMap Is Nothing Then Set mCADEntityToSAPMap = CreateObject("Scripting.Dictionary")
    Dim ws As Worksheet
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(LOG_SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.rows.count, "A").End(xlUp).row
    If lastRow < 2 Then Exit Sub
    
    Dim r As Long
    For r = 2 To lastRow
        Dim h As String
        Dim sapType As String
        Dim sapName As String
        h = CStr(ws.Cells(r, 1).Value)
        sapType = CStr(ws.Cells(r, 4).Value)
        sapName = CStr(ws.Cells(r, 5).Value)
        If Len(Trim$(h)) > 0 And Len(Trim$(sapName)) > 0 Then
            On Error Resume Next
            mCADEntityToSAPMap(h) = sapType & "|" & sapName
            On Error GoTo 0
        End If
    Next r
    Core_Sync_Manager.LogStatus "Rebuilt CAD->SAP mapping from log sheet (" & mCADEntityToSAPMap.count & " entries)"
    On Error GoTo 0
End Sub

' Optional helper: allow external modules to query mapping (alias kept for compatibility)
Public Function GetSAPMappingForHandle(Handle As String) As String
    GetSAPMappingForHandle = GetSAPMappingForCADHandle(Handle)
End Function

' ----------------------------
' End of module
' ----------------------------
