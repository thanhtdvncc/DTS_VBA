VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsViewModel_WallSync"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ==============================================================================
' Class Module: clsViewModel_WallSync
' Purpose: ViewModel for Wall-to-SAP synchronization
' Architecture: ViewModel Layer (MVVM)
' Responsibilities:
'   - Manage UI state (LayerName, WallThickness, Status)
'   - Orchestrate Logic Layer (LibDTS_Algo_Wall)
'   - Provide data binding for View (frmWallConverter)
' ==============================================================================
Option Explicit

' --- CONSTANTS ---
Private Const PROP_SYNCED_TO_SAP As String = "SyncedToSAP"
Private Const PROP_SYNC_TIME As String = "SyncTime"

' --- PRIVATE MEMBERS ---
Private p_LayerName As String
Private p_WallThickness As Double
Private p_Status As String
Private p_LastError As String

' --- CONSTRUCTOR ---
Private Sub Class_Initialize()
    p_LayerName = "DTS_WALL_DIAGRAM"    ' Default
    p_WallThickness = 200                 ' Default 200mm
    p_Status = "Ready"
    p_LastError = ""
End Sub

' --- DESTRUCTOR ---
Private Sub Class_Terminate()
    ' Cleanup if needed
End Sub

' ==============================================================================
' PROPERTIES (Data Binding)
' ==============================================================================

Public Property Get LayerName() As String
    LayerName = p_LayerName
End Property

Public Property Let LayerName(Value As String)
    p_LayerName = Value
End Property

Public Property Get WallThickness() As Double
    WallThickness = p_WallThickness
End Property

Public Property Let WallThickness(Value As Double)
    If Value > 0 Then
        p_WallThickness = Value
    End If
End Property

Public Property Get Status() As String
    Status = p_Status
End Property

' Read-only property (controlled by ViewModel logic)
Public Property Get LastError() As String
    LastError = p_LastError
End Property

' ==============================================================================
' METHOD: Initialize
' Purpose: Load default settings from global configuration
' ==============================================================================
Public Sub Initialize()
    On Error GoTo ErrHandler
    
    ' Load settings from LibDTS_Global.Config
    Dim config As clsDTSConfig
    Set config = LibDTS_Global.Config
    
    ' Load layer name (if configured)
    Dim layerCfg As String
    layerCfg = config.GetVal("Layer_Wall", "DTS_WALL_DIAGRAM")
    If Len(Trim$(layerCfg)) > 0 Then
        p_LayerName = layerCfg
    End If
    
    ' Load default wall thickness (if configured)
    Dim thicknessCfg As Variant
    thicknessCfg = config.GetVal("WallThickness_Default", 200)
    If IsNumeric(thicknessCfg) Then
        If CDbl(thicknessCfg) > 0 Then
            p_WallThickness = CDbl(thicknessCfg)
        End If
    End If
    
    p_Status = "Initialized"
    LibDTS_Logger.Log "clsViewModel_WallSync: Initialized with Layer=" & p_LayerName & ", Thickness=" & p_WallThickness, DTS_INFO
    
    Exit Sub
    
ErrHandler:
    p_LastError = "Initialize error: " & Err.Description
    p_Status = "Error"
    LibDTS_Logger.Log "clsViewModel_WallSync.Initialize: " & p_LastError, DTS_ERROR
End Sub

' ==============================================================================
' METHOD: RunSyncProcess
' Purpose: Execute the complete wall-to-SAP synchronization workflow
' Returns: True if successful, False otherwise
' ==============================================================================
Public Function RunSyncProcess() As Boolean
    On Error GoTo ErrHandler
    
    RunSyncProcess = False
    p_Status = "Processing..."
    p_LastError = ""
    
    ' Log start
    LibDTS_Logger.Log "clsViewModel_WallSync.RunSyncProcess: Starting wall sync process", DTS_INFO
    
    ' ===== STEP 1: GET SELECTION AS FRAMES =====
    p_Status = "Step 1/3: Getting selection..."
    
    Dim acadDoc As Object
    On Error Resume Next
    Set acadDoc = GetObject(, "AutoCAD.Application").ActiveDocument
    On Error GoTo ErrHandler
    
    If acadDoc Is Nothing Then
        p_LastError = "Cannot connect to AutoCAD"
        p_Status = "Error: " & p_LastError
        LibDTS_Logger.Log "clsViewModel_WallSync.RunSyncProcess: " & p_LastError, DTS_ERROR
        Exit Function
    End If
    
    Dim frames As Collection
    Set frames = LibDTS_Algo_Wall.GetSelectionAsFrames(acadDoc)
    
    If frames Is Nothing Or frames.count = 0 Then
        p_LastError = "No valid wall frames selected"
        p_Status = "Warning: " & p_LastError
        LibDTS_Logger.Log "clsViewModel_WallSync.RunSyncProcess: " & p_LastError, DTS_WARNING
        ' Not an error - user might have cancelled selection
        RunSyncProcess = True
        Exit Function
    End If
    
    LibDTS_Logger.Log "clsViewModel_WallSync.RunSyncProcess: Retrieved " & frames.count & " frames", DTS_INFO
    
    ' ===== STEP 2: ANALYZE WALL CONNECTIONS =====
    p_Status = "Step 2/3: Analyzing connections..."
    
    If Not LibDTS_Algo_Wall.AnalyzeWallConnections(frames) Then
        p_LastError = "Failed to analyze wall connections: " & LibDTS_Algo_Wall.LastError
        p_Status = "Error: " & p_LastError
        LibDTS_Logger.Log "clsViewModel_WallSync.RunSyncProcess: " & p_LastError, DTS_ERROR
        Exit Function
    End If
    
    LibDTS_Logger.Log "clsViewModel_WallSync.RunSyncProcess: Analysis complete", DTS_INFO
    
    ' ===== STEP 3: SYNC FRAMES TO SAP =====
    p_Status = "Step 3/3: Syncing to SAP2000..."
    
    On Error Resume Next
    LibDTS_Algo_Wall.SyncFramesToSAP frames
    
    ' Check for errors during sync
    If Err.Number <> 0 Then
        p_LastError = "Sync failed: " & Err.Description
        p_Status = "Error: " & p_LastError
        LibDTS_Logger.Log "clsViewModel_WallSync.RunSyncProcess: " & p_LastError, DTS_ERROR
        Exit Function
    End If
    On Error GoTo ErrHandler
    
    ' Check if all frames were synced successfully
    Dim syncCount As Long
    Dim failCount As Long
    syncCount = 0
    failCount = 0
    
    Dim frame As Variant
    For Each frame In frames
        If frame.Base.Properties.Exists(PROP_SYNCED_TO_SAP) Then
            If frame.Base.Properties(PROP_SYNCED_TO_SAP) = True Then
                syncCount = syncCount + 1
            Else
                failCount = failCount + 1
            End If
        Else
            failCount = failCount + 1
        End If
    Next frame
    
    ' Update status
    If failCount = 0 Then
        p_Status = "Success: Synced " & syncCount & " frames"
        RunSyncProcess = True
    Else
        p_Status = "Partial: Synced " & syncCount & " frames, Failed " & failCount
        RunSyncProcess = True ' Partial success
    End If
    
    LibDTS_Logger.Log "clsViewModel_WallSync.RunSyncProcess: " & p_Status, DTS_INFO
    Exit Function
    
ErrHandler:
    p_LastError = "RunSyncProcess error: " & Err.Description
    p_Status = "Error: " & p_LastError
    LibDTS_Logger.Log "clsViewModel_WallSync.RunSyncProcess: " & p_LastError, DTS_ERROR
    RunSyncProcess = False
End Function

' ==============================================================================
' METHOD: ValidateSettings
' Purpose: Validate ViewModel settings before processing
' Returns: True if valid, False otherwise
' ==============================================================================
Public Function ValidateSettings() As Boolean
    On Error Resume Next
    
    ValidateSettings = False
    p_LastError = ""
    
    ' Validate layer name
    If Len(Trim$(p_LayerName)) = 0 Then
        p_LastError = "Layer name cannot be empty"
        Exit Function
    End If
    
    ' Validate wall thickness
    If p_WallThickness <= 0 Then
        p_LastError = "Wall thickness must be positive"
        Exit Function
    End If
    
    ValidateSettings = True
End Function

' ==============================================================================
' METHOD: UpdateStatus
' Purpose: Update status message (for UI notifications)
' ==============================================================================
Public Sub UpdateStatus(newStatus As String)
    p_Status = newStatus
    LibDTS_Logger.Log "clsViewModel_WallSync.Status: " & newStatus, DTS_INFO
End Sub
