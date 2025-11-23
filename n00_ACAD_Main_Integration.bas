Attribute VB_Name = "n00_ACAD_Main_Integration"
Option Explicit
'===============================================================
' Module: mod_Main_Integration
' Purpose: Main integration module - connects all cores
' Dependencies: All Core modules, existing SAP reader modules
'===============================================================

' ----------------------------
' Public API - Called from UserForm
' ----------------------------

' Plot SAP2000 model to AutoCAD
Public Sub PlotModelToNewDrawing(Optional ByVal framesOnly As Boolean = True, _
                                 Optional ByVal showNodeNames As Boolean = True, _
                                 Optional ByVal showFrameNames As Boolean = True, _
                                 Optional ByVal showShellNames As Boolean = True)
    On Error GoTo ErrHandler
    
    SetStatusToForm "========== Starting Plot to CAD =========="
    
    ' Get or create AutoCAD
    Dim acadApp As Object
    Set acadApp = Core_Utils.GetOrCreateAutoCAD()
    
    If acadApp Is Nothing Then
        MsgBox "AutoCAD not available.", vbExclamation, "Plot to CAD"
        Exit Sub
    End If
    
    ' Get active document or create new
    Dim acadDoc As Object
    If acadApp.Documents.count = 0 Then
        Set acadDoc = acadApp.Documents.Add
    Else
        Set acadDoc = acadApp.ActiveDocument
    End If
    
    Core_Utils.ShowExcelOnTop
    
    ' Ensure SAP connection
    ConnectSAP2000
    If Not IsObject(SapModel) Then
        SetStatusToForm "SAP2000 not connected. Attempting connection..."
        If Not ConnectSAP2000() Then
            MsgBox "SAP2000 connection required for plotting.", vbExclamation, "Plot to CAD"
            Exit Sub
        End If
    End If
    
    
    ' Set AutoCAD units to mm
    Core_Utils.SetInsUnits acadDoc, 4
    
     ' Set SAP2000 units to kN-mm-C
    SapModel.SetPresentUnits (5)
    
    ' Set callback for status updates
    Dim callback As New StatusCallbackHandler
    Core_CAD_Plotter.SetStatusCallback callback
    Core_Sync_Manager.SetSyncStatusCallback callback
    
    ' Execute sync
    Core_Sync_Manager.SyncSAPToCAD acadDoc, SapModel, framesOnly, showNodeNames, showFrameNames, showShellNames
    
    ' Finalize view
    Core_Utils.ZoomAll acadDoc
    Core_Utils.EnableViewCube acadApp, acadDoc
    
    SetStatusToForm "========== Plot to CAD Complete =========="
    
    Exit Sub
    
ErrHandler:
    SetStatusToForm "ERROR: " & err.description
    MsgBox "Error in PlotModelToNewDrawing: " & err.description, vbCritical, "Error"
End Sub
' Add this wrapper function to Main_CAD_SAP module

Public Sub PlotModelToNewDrawingWithFilters(Optional framesOnly As Boolean = True, _
        Optional showNodeNames As Boolean = True, _
        Optional showFrameNames As Boolean = False, _
        Optional showShellNames As Boolean = False, _
        Optional sapOnlyMode As Boolean = False, _
        Optional frameFilterRules As String = "", _
        Optional areaFilterRules As String = "")
    
    On Error GoTo ErrHandler
    
    ' Get or create AutoCAD
    Dim acadApp As Object
    Set acadApp = Core_Utils.GetOrCreateAutoCAD()
    
    If acadApp Is Nothing Then
        MsgBox "Cannot start AutoCAD.", vbCritical, "Plot to CAD"
        Exit Sub
    End If
    
    ' Create new drawing
    Dim acadDoc As Object
    Set acadDoc = acadApp.Documents.Add
    
    acadApp.Visible = True
    
    ' Connect to SAP2000
    If Not IsObject(SapModel) Then
        If Not ConnectSAP2000() Then
            MsgBox "SAP2000 connection required.", vbExclamation, "Plot to CAD"
            Exit Sub
        End If
    End If
    
    ' Call sync with filters
    Core_Sync_Manager.SyncSAPToCADWithFilters acadDoc, SapModel, framesOnly, _
        showNodeNames, showFrameNames, showShellNames, _
        sapOnlyMode, frameFilterRules, areaFilterRules, False
    
    ' Zoom extents
    On Error Resume Next
    acadApp.ZoomExtents
    On Error GoTo ErrHandler
    
    MsgBox "Model plotted to AutoCAD successfully!", vbInformation, "Plot Complete"
    
    Exit Sub
    
ErrHandler:
    MsgBox "Error plotting to CAD: " & err.description, vbCritical, "Plot Error"
End Sub
' Import edited drawing from AutoCAD to SAP2000
Public Sub ImportEditedDrawingToSAP(Optional ByVal tolerance As Double = 1#, _
                                    Optional ByVal scaleFactor As Double = 1#)
    On Error GoTo ErrHandler
    
    SetStatusToForm "========== Starting Import from CAD =========="
    
    ' Validate inputs
    tolerance = Core_Utils.ValidateTolerance(tolerance)
    scaleFactor = Core_Utils.ValidateScaleFactor(scaleFactor)
    
    ' Get AutoCAD
    Dim acadApp As Object
    Set acadApp = Core_Utils.GetOrCreateAutoCAD()
    
    If acadApp Is Nothing Then
        MsgBox "AutoCAD not available.", vbExclamation, "Import to SAP"
        Exit Sub
    End If
    
    Dim acadDoc As Object
    Set acadDoc = acadApp.ActiveDocument
    
    If acadDoc Is Nothing Then
        MsgBox "No active drawing.", vbExclamation, "Import to SAP"
        Exit Sub
    End If
    
    ' Ensure SAP connection
    If Not IsObject(SapModel) Then
        If Not ConnectSAP2000() Then
            MsgBox "SAP2000 connection failed.", vbExclamation, "Import to SAP"
            Exit Sub
        End If
    End If
    
    ' Set callback for status updates
    Dim callback As New StatusCallbackHandler
    Core_Sync_Manager.SetSyncStatusCallback callback
    
    ' Execute sync
    Core_Sync_Manager.SyncCADToSAP acadDoc, SapModel, tolerance, scaleFactor
    
    SetStatusToForm "========== Import from CAD Complete =========="
    
    Exit Sub
    
ErrHandler:
    SetStatusToForm "ERROR: " & err.description
    MsgBox "Error during import: " & err.description, vbCritical, "Error"
End Sub

' ----------------------------
' Status Callback Handler Class
' ----------------------------

' Create a class module named: StatusCallbackHandler
' Code for StatusCallbackHandler.cls:
'
' Option Explicit
' Public Sub SetStatus(msg As String)
'     Core_Utils.SetStatusToForm msg
' End Sub

