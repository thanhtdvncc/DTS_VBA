VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmWallConverter 
   Caption         =   "SAP2000 Wall Tool - MVVM Edition"
   ClientHeight    =   3000
   ClientLeft      =   120
   ClientTop       =   675
   ClientWidth     =   4500
   OleObjectBlob   =   "frmWallConverter.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmWallConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ==============================================================================
' UserForm: frmWallConverter (Refactored to MVVM)
' Architecture: View Layer
' Responsibilities:
'   - Display UI (no business logic)
'   - Forward events to ViewModel
'   - Data binding with ViewModel properties
' ==============================================================================
Option Explicit

' --- VIEWMODEL INSTANCE ---
Private m_ViewModel As clsViewModel_WallSync

' ==============================================================================
' EVENT: UserForm_Initialize
' Purpose: Set up ViewModel and bind initial values
' ==============================================================================
Private Sub UserForm_Initialize()
    On Error GoTo ErrHandler
    
    ' Create ViewModel instance
    Set m_ViewModel = New clsViewModel_WallSync
    
    ' Initialize ViewModel (loads settings from Config)
    m_ViewModel.Initialize
    
    ' Bind ViewModel properties to UI controls (if they exist)
    On Error Resume Next
    
    ' If we have a LayerName textbox
    If HasControl("txtLayerName") Then
        Me.Controls("txtLayerName").Text = m_ViewModel.LayerName
    End If
    
    ' If we have a WallThickness textbox
    If HasControl("txtThickness") Then
        Me.Controls("txtThickness").Text = CStr(m_ViewModel.WallThickness)
    End If
    
    ' If we have a Status label
    If HasControl("lblStatus") Then
        Me.Controls("lblStatus").Caption = m_ViewModel.Status
    End If
    
    On Error GoTo ErrHandler
    
    LibDTS_Logger.Log "frmWallConverter: Initialized with ViewModel", DTS_INFO
    Exit Sub
    
ErrHandler:
    MsgBox "Error initializing form: " & Err.Description, vbCritical, "Initialization Error"
End Sub

' ==============================================================================
' EVENT: btnCombineWithSAP_Click
' Purpose: Execute wall-to-SAP synchronization via ViewModel
' ==============================================================================
Private Sub btnCombineWithSAP_Click()
    On Error GoTo ErrHandler
    
    ' Validate ViewModel settings
    If Not m_ViewModel.ValidateSettings() Then
        MsgBox "Invalid settings: " & m_ViewModel.LastError, vbExclamation, "Validation Error"
        Exit Sub
    End If
    
    ' Update status label before processing
    UpdateStatusLabel "Processing..."
    
    ' Call ViewModel to run the sync process
    Dim success As Boolean
    success = m_ViewModel.RunSyncProcess()
    
    ' Update status label after processing
    UpdateStatusLabel m_ViewModel.Status
    
    ' Show result to user
    If success Then
        MsgBox m_ViewModel.Status, vbInformation, "Sync Complete"
    Else
        MsgBox "Sync failed: " & m_ViewModel.LastError, vbCritical, "Sync Error"
    End If
    
    LibDTS_Logger.Log "frmWallConverter: Sync process completed - " & m_ViewModel.Status, DTS_INFO
    Exit Sub
    
ErrHandler:
    MsgBox "Error during sync: " & Err.Description, vbCritical, "Sync Error"
    UpdateStatusLabel "Error: " & Err.Description
End Sub

' ==============================================================================
' EVENT: btnCancel_Click
' Purpose: Close the form
' ==============================================================================
Private Sub btnCancel_Click()
    Unload Me
End Sub

' ==============================================================================
' HELPER: UpdateStatusLabel
' Purpose: Update status label if it exists
' ==============================================================================
Private Sub UpdateStatusLabel(statusText As String)
    On Error Resume Next
    If HasControl("lblStatus") Then
        Me.Controls("lblStatus").Caption = statusText
    End If
    On Error GoTo 0
End Sub

' ==============================================================================
' HELPER: HasControl
' Purpose: Check if a control exists on the form
' ==============================================================================
Private Function HasControl(ctrlName As String) As Boolean
    On Error Resume Next
    Dim ctrl As Object
    Set ctrl = Me.Controls(ctrlName)
    HasControl = Not ctrl Is Nothing
    On Error GoTo 0
End Function

' ==============================================================================
' EVENT: UserForm_QueryClose
' Purpose: Cleanup when form is closed
' ==============================================================================
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    On Error Resume Next
    
    ' Cleanup ViewModel
    Set m_ViewModel = Nothing
    
    LibDTS_Logger.Log "frmWallConverter: Form closed", DTS_INFO
End Sub

' ==============================================================================
' OPTIONAL: Property Get/Set for data binding (if needed by other code)
' ==============================================================================

' Get ViewModel instance (for external access if needed)
Public Property Get ViewModel() As clsViewModel_WallSync
    Set ViewModel = m_ViewModel
End Property
