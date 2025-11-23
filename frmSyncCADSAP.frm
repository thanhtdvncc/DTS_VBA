VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSyncCADSAP 
   Caption         =   "Model from AutoCAD (by thanhtdvncc):"
   ClientHeight    =   4860
   ClientLeft      =   120
   ClientTop       =   675
   ClientWidth     =   7470
   OleObjectBlob   =   "frmSyncCADSAP.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSyncCADSAP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'===============================================================
' UserForm: frmSyncCADSAP (UPDATED & FIXED)
' - Single Initialize routine that sets up controls consistently
' - Keeps quick-filter helpers, persistence, and plotting/import handlers
' - Uses ControlExists checks when referencing controls
' - Ensures txtStatus display and chkRealTimeSync reflect actual state
' - Preserves original module/sub names (no renaming)
'===============================================================

' ----------------------------
' Helpers
' ----------------------------
Private Function ControlExists(ctrlName As String) As Boolean
    On Error Resume Next
    Dim c As Object
    Set c = Me.Controls(ctrlName)
    ControlExists = Not c Is Nothing
    On Error GoTo 0
End Function

' Build multiline filter rules from quick inputs for side = "Frame" or "Area"
Private Function BuildFilterRulesFromQuick(ByVal side As String) As String
    On Error GoTo ErrHandler
    Dim result As String
    result = ""
    
    Dim nameBox As String, guidBox As String, sectionBox As String
    nameBox = ""
    guidBox = ""
    sectionBox = ""
    
    If LCase$(side) = "frame" Then
        If ControlExists("txtFrameNameQuick") Then nameBox = Trim$(Me.Controls("txtFrameNameQuick").text)
        If ControlExists("txtFrameGUIDQuick") Then guidBox = Trim$(Me.Controls("txtFrameGUIDQuick").text)
        If ControlExists("txtFrameSectionQuick") Then sectionBox = Trim$(Me.Controls("txtFrameSectionQuick").text)
    Else
        If ControlExists("txtAreaNameQuick") Then nameBox = Trim$(Me.Controls("txtAreaNameQuick").text)
        If ControlExists("txtAreaGUIDQuick") Then guidBox = Trim$(Me.Controls("txtAreaGUIDQuick").text)
        If ControlExists("txtAreaSectionQuick") Then sectionBox = Trim$(Me.Controls("txtAreaSectionQuick").text)
    End If
    
    ' Helper to append tokens into result as "field:token" per line
    Dim tokens() As String
    Dim i As Long
    Dim t As String
    If Len(nameBox) > 0 Then
        tokens = Split(nameBox, ",")
        For i = LBound(tokens) To UBound(tokens)
            t = Trim$(tokens(i))
            If Len(t) > 0 Then
                If result = "" Then
                    result = "name:" & t
                Else
                    result = result & vbCrLf & "name:" & t
                End If
            End If
        Next i
    End If
    
    If Len(guidBox) > 0 Then
        tokens = Split(guidBox, ",")
        For i = LBound(tokens) To UBound(tokens)
            t = Trim$(tokens(i))
            If Len(t) > 0 Then
                If result = "" Then
                    result = "guid:" & t
                Else
                    result = result & vbCrLf & "guid:" & t
                End If
            End If
        Next i
    End If
    
    If Len(sectionBox) > 0 Then
        tokens = Split(sectionBox, ",")
        For i = LBound(tokens) To UBound(tokens)
            t = Trim$(tokens(i))
            If Len(t) > 0 Then
                If result = "" Then
                    result = "section:" & t
                Else
                    result = result & vbCrLf & "section:" & t
                End If
            End If
        Next i
    End If
    
    BuildFilterRulesFromQuick = result
    Exit Function
ErrHandler:
    BuildFilterRulesFromQuick = ""
    Resume Next
End Function

' Update summary label text for side Frame/Area
Private Sub UpdateSummaryLabel(ByVal side As String)
    On Error Resume Next
    Dim nameBox As String, guidBox As String, sectionBox As String
    nameBox = "": guidBox = "": sectionBox = ""
    
    If LCase$(side) = "frame" Then
        If ControlExists("txtFrameNameQuick") Then nameBox = Trim$(Me.Controls("txtFrameNameQuick").text)
        If ControlExists("txtFrameGUIDQuick") Then guidBox = Trim$(Me.Controls("txtFrameGUIDQuick").text)
        If ControlExists("txtFrameSectionQuick") Then sectionBox = Trim$(Me.Controls("txtFrameSectionQuick").text)
        
        Dim summary As String
        summary = "Exclude by"
        If Len(nameBox) > 0 Then summary = summary & " Name=[" & nameBox & "]"
        If Len(guidBox) > 0 Then summary = summary & " GUID=[" & guidBox & "]"
        If Len(sectionBox) > 0 Then summary = summary & " Section=[" & sectionBox & "]"
        If Len(summary) = 7 Then summary = "No Frame filters set"
        
        If ControlExists("lblFrameSummary") Then Me.Controls("lblFrameSummary").Caption = summary
        If ControlExists("lblFrameGuide") Then
            If Len(nameBox) + Len(guidBox) + Len(sectionBox) > 0 Then
                Me.Controls("lblFrameGuide").Visible = True
                Me.Controls("lblFrameGuide").Caption = "Hints: Use '*' for wildcard. Separate multiple values by comma."
            Else
                Me.Controls("lblFrameGuide").Visible = False
            End If
        End If
    Else
        If ControlExists("txtAreaNameQuick") Then nameBox = Trim$(Me.Controls("txtAreaNameQuick").text)
        If ControlExists("txtAreaGUIDQuick") Then guidBox = Trim$(Me.Controls("txtAreaGUIDQuick").text)
        If ControlExists("txtAreaSectionQuick") Then sectionBox = Trim$(Me.Controls("txtAreaSectionQuick").text)
        
        Dim summaryA As String
        summaryA = "Exclude by"
        If Len(nameBox) > 0 Then summaryA = summaryA & " Name=[" & nameBox & "]"
        If Len(guidBox) > 0 Then summaryA = summaryA & " GUID=[" & guidBox & "]"
        If Len(sectionBox) > 0 Then summaryA = summaryA & " Section=[" & sectionBox & "]"
        If Len(summaryA) = 7 Then summaryA = "No Area filters set"
        
        If ControlExists("lblAreaSummary") Then Me.Controls("lblAreaSummary").Caption = summaryA
        If ControlExists("lblAreaGuide") Then
            If Len(nameBox) + Len(guidBox) + Len(sectionBox) > 0 Then
                Me.Controls("lblAreaGuide").Visible = True
                Me.Controls("lblAreaGuide").Caption = "Hints: Use '*' for wildcard. Separate multiple values by comma."
            Else
                Me.Controls("lblAreaGuide").Visible = False
            End If
        End If
    End If
    On Error GoTo 0
End Sub

' ----------------------------
' Persistence: Save/Load quick inputs using Workbook Defined Names
' Names:
'   DTS_FrameNameQuick, DTS_FrameGUIDQuick, DTS_FrameSectionQuick
'   DTS_AreaNameQuick, DTS_AreaGUIDQuick, DTS_AreaSectionQuick
'   DTS_Filters_LastSaved
' ----------------------------
Private Sub SaveQuickFilterSettings()
    On Error GoTo ErrHandler
    Dim wb As Workbook
    Set wb = ThisWorkbook
    
    Dim v1 As String, v2 As String, v3 As String, v4 As String, v5 As String, v6 As String
    v1 = "": v2 = "": v3 = "": v4 = "": v5 = "": v6 = ""
    On Error Resume Next
    If ControlExists("txtFrameNameQuick") Then v1 = Me.Controls("txtFrameNameQuick").text
    If ControlExists("txtFrameGUIDQuick") Then v2 = Me.Controls("txtFrameGUIDQuick").text
    If ControlExists("txtFrameSectionQuick") Then v3 = Me.Controls("txtFrameSectionQuick").text
    If ControlExists("txtAreaNameQuick") Then v4 = Me.Controls("txtAreaNameQuick").text
    If ControlExists("txtAreaGUIDQuick") Then v5 = Me.Controls("txtAreaGUIDQuick").text
    If ControlExists("txtAreaSectionQuick") Then v6 = Me.Controls("txtAreaSectionQuick").text
    On Error GoTo ErrHandler
    
    ' Remove old names
    On Error Resume Next
    wb.names("DTS_FrameNameQuick").Delete
    wb.names("DTS_FrameGUIDQuick").Delete
    wb.names("DTS_FrameSectionQuick").Delete
    wb.names("DTS_AreaNameQuick").Delete
    wb.names("DTS_AreaGUIDQuick").Delete
    wb.names("DTS_AreaSectionQuick").Delete
    wb.names("DTS_Filters_LastSaved").Delete
    On Error GoTo ErrHandler
    
    Dim ref1 As String, ref2 As String, ref3 As String, ref4 As String, ref5 As String, ref6 As String, refTime As String
    ref1 = "=""" & Replace(v1, """", """""") & """"
    ref2 = "=""" & Replace(v2, """", """""") & """"
    ref3 = "=""" & Replace(v3, """", """""") & """"
    ref4 = "=""" & Replace(v4, """", """""") & """"
    ref5 = "=""" & Replace(v5, """", """""") & """"
    ref6 = "=""" & Replace(v6, """", """""") & """"
    refTime = "=""" & Format(Now, "yyyy-mm-dd HH:nn:ss") & """"
    
    On Error Resume Next
    wb.names.Add Name:="DTS_FrameNameQuick", RefersTo:=ref1
    wb.names.Add Name:="DTS_FrameGUIDQuick", RefersTo:=ref2
    wb.names.Add Name:="DTS_FrameSectionQuick", RefersTo:=ref3
    wb.names.Add Name:="DTS_AreaNameQuick", RefersTo:=ref4
    wb.names.Add Name:="DTS_AreaGUIDQuick", RefersTo:=ref5
    wb.names.Add Name:="DTS_AreaSectionQuick", RefersTo:=ref6
    wb.names.Add Name:="DTS_Filters_LastSaved", RefersTo:=refTime
    On Error GoTo ErrHandler
    
    AddLog "[UI] Quick filter settings saved to workbook Names."
    Exit Sub
ErrHandler:
    AddLog "[UI] SaveQuickFilterSettings error: " & err.description
    Resume Next
End Sub

Private Sub LoadQuickFilterSettings()
    On Error GoTo ErrHandler
    Dim wb As Workbook
    Set wb = ThisWorkbook
    
    Dim v1 As String, v2 As String, v3 As String, v4 As String, v5 As String, v6 As String
    v1 = "": v2 = "": v3 = "": v4 = "": v5 = "": v6 = ""
    
    On Error Resume Next
    Dim nm As Name
    Set nm = Nothing
    Set nm = wb.names("DTS_FrameNameQuick")
    If Not nm Is Nothing Then
        v1 = CStr(nm.RefersTo)
        If Left$(v1, 1) = "=" Then v1 = mid$(v1, 2)
        If Len(v1) >= 2 Then
            If Left$(v1, 1) = """" And Right$(v1, 1) = """" Then
                v1 = mid$(v1, 2, Len(v1) - 2)
                v1 = Replace(v1, """""", """")
            End If
        End If
    End If
    
    Set nm = Nothing
    Set nm = wb.names("DTS_FrameGUIDQuick")
    If Not nm Is Nothing Then
        v2 = CStr(nm.RefersTo)
        If Left$(v2, 1) = "=" Then v2 = mid$(v2, 2)
        If Len(v2) >= 2 Then
            If Left$(v2, 1) = """" And Right$(v2, 1) = """" Then
                v2 = mid$(v2, 2, Len(v2) - 2)
                v2 = Replace(v2, """""", """")
            End If
        End If
    End If
    
    Set nm = Nothing
    Set nm = wb.names("DTS_FrameSectionQuick")
    If Not nm Is Nothing Then
        v3 = CStr(nm.RefersTo)
        If Left$(v3, 1) = "=" Then v3 = mid$(v3, 2)
        If Len(v3) >= 2 Then
            If Left$(v3, 1) = """" And Right$(v3, 1) = """" Then
                v3 = mid$(v3, 2, Len(v3) - 2)
                v3 = Replace(v3, """""", """")
            End If
        End If
    End If
    
    Set nm = Nothing
    Set nm = wb.names("DTS_AreaNameQuick")
    If Not nm Is Nothing Then
        v4 = CStr(nm.RefersTo)
        If Left$(v4, 1) = "=" Then v4 = mid$(v4, 2)
        If Len(v4) >= 2 Then
            If Left$(v4, 1) = """" And Right$(v4, 1) = """" Then
                v4 = mid$(v4, 2, Len(v4) - 2)
                v4 = Replace(v4, """""", """")
            End If
        End If
    End If
    
    Set nm = Nothing
    Set nm = wb.names("DTS_AreaGUIDQuick")
    If Not nm Is Nothing Then
        v5 = CStr(nm.RefersTo)
        If Left$(v5, 1) = "=" Then v5 = mid$(v5, 2)
        If Len(v5) >= 2 Then
            If Left$(v5, 1) = """" And Right$(v5, 1) = """" Then
                v5 = mid$(v5, 2, Len(v5) - 2)
                v5 = Replace(v5, """""", """")
            End If
        End If
    End If
    
    Set nm = Nothing
    Set nm = wb.names("DTS_AreaSectionQuick")
    If Not nm Is Nothing Then
        v6 = CStr(nm.RefersTo)
        If Left$(v6, 1) = "=" Then v6 = mid$(v6, 2)
        If Len(v6) >= 2 Then
            If Left$(v6, 1) = """" And Right$(v6, 1) = """" Then
                v6 = mid$(v6, 2, Len(v6) - 2)
                v6 = Replace(v6, """""", """")
            End If
        End If
    End If
    On Error GoTo ErrHandler
    
    On Error Resume Next
    If ControlExists("txtFrameNameQuick") Then Me.Controls("txtFrameNameQuick").text = v1
    If ControlExists("txtFrameGUIDQuick") Then Me.Controls("txtFrameGUIDQuick").text = v2
    If ControlExists("txtFrameSectionQuick") Then Me.Controls("txtFrameSectionQuick").text = v3
    If ControlExists("txtAreaNameQuick") Then Me.Controls("txtAreaNameQuick").text = v4
    If ControlExists("txtAreaGUIDQuick") Then Me.Controls("txtAreaGUIDQuick").text = v5
    If ControlExists("txtAreaSectionQuick") Then Me.Controls("txtAreaSectionQuick").text = v6
    On Error GoTo ErrHandler
    
    AddLog "[UI] Quick filter settings loaded from workbook Names (if any)."
    Exit Sub
ErrHandler:
    AddLog "[UI] LoadQuickFilterSettings error: " & err.description
    Resume Next
End Sub

' ----------------------------
' UI: Update summaries when quick inputs change
' ----------------------------
Private Sub txtFrameNameQuick_Change()
    On Error Resume Next
    UpdateSummaryLabel "Frame"
End Sub
Private Sub txtFrameGUIDQuick_Change()
    On Error Resume Next
    UpdateSummaryLabel "Frame"
End Sub
Private Sub txtFrameSectionQuick_Change()
    On Error Resume Next
    UpdateSummaryLabel "Frame"
End Sub
Private Sub txtAreaNameQuick_Change()
    On Error Resume Next
    UpdateSummaryLabel "Area"
End Sub
Private Sub txtAreaGUIDQuick_Change()
    On Error Resume Next
    UpdateSummaryLabel "Area"
End Sub
Private Sub txtAreaSectionQuick_Change()
    On Error Resume Next
    UpdateSummaryLabel "Area"
End Sub

' ----------------------------
' Initialize / UI lifecycle (single Initialize)
' ----------------------------
Private Sub UserForm_Initialize()
    On Error Resume Next
    ' Set default checkbox states only if controls exist
    If ControlExists("chkFramesOnly") Then Me.chkFramesOnly.Value = True
    If ControlExists("chkShowNodeName") Then Me.chkShowNodeName.Value = True
    If ControlExists("chkShowFrameName") Then Me.chkShowFrameName.Value = False
    If ControlExists("chkShowShellName") Then Me.chkShowShellName.Value = False
    If ControlExists("chkOverwrite") Then Me.chkOverwrite.Value = False
    
    ' Tolerance / Scale defaults
    If ControlExists("txtTolerance") Then
        Me.txtTolerance.text = "1"
    End If
    If ControlExists("txtScaleFactor") Then
        Me.txtScaleFactor.text = "1"
    End If
    
    ' Setup status textbox robustly
    If ControlExists("txtStatus") Then
        With Me.txtStatus
            .Multiline = True
            .Locked = True
            On Error Resume Next
            .ScrollBars = fmScrollBarsVertical
            On Error GoTo 0
            .BackColor = RGB(240, 248, 255)
            .ForeColor = RGB(0, 0, 0)
            .Font.Name = "Consolas"
            .Font.Size = 9
            .text = "Ready to synchronize."
        End With
    End If
    
    ' Setup button colors if present
    On Error Resume Next
    If ControlExists("btnPlotToCAD") Then Me.btnPlotToCAD.BackColor = RGB(144, 238, 144)
    If ControlExists("btnImportToSAP") Then Me.btnImportToSAP.BackColor = RGB(135, 206, 250)
    If ControlExists("btnClearLog") Then Me.btnClearLog.BackColor = RGB(255, 255, 224)
    If ControlExists("btnClose") Then Me.btnClose.BackColor = RGB(255, 182, 193)
    On Error GoTo 0
    
    ' Set tooltips for quick filter boxes
    On Error Resume Next
    If ControlExists("txtFrameNameQuick") Then Me.Controls("txtFrameNameQuick").ControlTipText = "Quick add: name pattern (e.g. BEAM_*). Use comma to separate multiple values."
    If ControlExists("txtFrameGUIDQuick") Then Me.Controls("txtFrameGUIDQuick").ControlTipText = "Quick add: guid (e.g. FRAME_123). Use comma to separate multiple values."
    If ControlExists("txtFrameSectionQuick") Then Me.Controls("txtFrameSectionQuick").ControlTipText = "Quick add: section pattern (e.g. RC*). Use comma to separate multiple values."
    If ControlExists("txtAreaNameQuick") Then Me.Controls("txtAreaNameQuick").ControlTipText = "Quick add: name pattern (e.g. FLOOR_*). Use comma to separate multiple values."
    If ControlExists("txtAreaGUIDQuick") Then Me.Controls("txtAreaGUIDQuick").ControlTipText = "Quick add: guid (e.g. AREA_456). Use comma to separate multiple values."
    If ControlExists("txtAreaSectionQuick") Then Me.Controls("txtAreaSectionQuick").ControlTipText = "Quick add: section pattern (e.g. SLAB*). Use comma to separate multiple values."
    On Error GoTo 0
    
    ' Load saved quick filter values
    LoadQuickFilterSettings
    
    ' Set chkRealTimeSync initial state based on Core_Event_Handler, if control exists
    On Error Resume Next
    If ControlExists("chkRealTimeSync") Then
        Dim syncOn As Boolean
        syncOn = False
        ' Safely call Core_Event_Handler.IsSyncEnabled() under error-handling
        On Error Resume Next
        syncOn = Core_Event_Handler.IsSyncEnabled()
        If err.number <> 0 Then
            err.Clear
            syncOn = False
        End If
        On Error GoTo 0
        Me.chkRealTimeSync.Value = syncOn
    End If
    On Error GoTo 0
    
    ' Update summaries initially
    UpdateSummaryLabel "Frame"
    UpdateSummaryLabel "Area"
    
    ' Update sync status label text/color
    UpdateSyncStatusLabel
End Sub

Private Sub UserForm_Activate()
    On Error Resume Next
    ' Ensure form is visible and repainted
    Me.Show vbModeless
    Me.Repaint
    DoEvents
    On Error GoTo 0
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' CRITICAL: Always disable real-time sync when form closes
    On Error Resume Next
    ' Safely call IsSyncEnabled / DisableRealTimeSync without checking module object
    If err.number <> 0 Then err.Clear
    On Error Resume Next
    If Core_Event_Handler.IsSyncEnabled() Then
        Core_Event_Handler.DisableRealTimeSync
        AddLog "[SYSTEM] Real-time sync auto-disabled (form closed)"
    End If
    If err.number <> 0 Then err.Clear
    On Error GoTo 0
End Sub

Private Sub UserForm_Terminate()
    ' Safety: Ensure sync disabled
    On Error Resume Next
    On Error Resume Next
    If Core_Event_Handler.IsSyncEnabled() Then
        Core_Event_Handler.DisableRealTimeSync
    End If
    If err.number <> 0 Then err.Clear
    On Error GoTo 0
End Sub

' ----------------------------
' Plot to CAD button click (uses quick inputs)
' ----------------------------
Private Sub btnPlotToCAD_Click()
    On Error GoTo ErrHandler
    Dim framesOnly  As Boolean
    Dim showNodeNames As Boolean
    Dim showFrameNames As Boolean
    Dim showShellNames As Boolean
    Dim sapOnlyMode As Boolean
    framesOnly = False
    showNodeNames = False
    showFrameNames = False
    showShellNames = False
    sapOnlyMode = False
    On Error Resume Next
    If ControlExists("chkFramesOnly") Then framesOnly = CBool(Me.chkFramesOnly.Value)
    If ControlExists("chkShowNodeName") Then showNodeNames = CBool(Me.chkShowNodeName.Value)
    If ControlExists("chkShowFrameName") Then showFrameNames = CBool(Me.chkShowFrameName.Value)
    If ControlExists("chkShowShellName") Then showShellNames = CBool(Me.chkShowShellName.Value)
    If ControlExists("chkSAPOnly") Then sapOnlyMode = CBool(Me.chkSAPOnly.Value)
    On Error GoTo ErrHandler
    
    AddLog "========================================="
    AddLog "PLOT TO CAD: Started (using quick filters)"
    AddLog "  Options: Frames Only=" & framesOnly
    AddLog "  SAP Selection Only=" & sapOnlyMode
    AddLog "======================================="
    DoEvents
    If ControlExists("btnPlotToCAD") Then Me.btnPlotToCAD.enabled = False
    If ControlExists("btnImportToSAP") Then Me.btnImportToSAP.enabled = False
    ' Build rules from quick inputs
    Dim frameFilters As String
    Dim areaFilters As String
    frameFilters = BuildFilterRulesFromQuick("Frame")
    areaFilters = BuildFilterRulesFromQuick("Area")
    ' Save quick inputs
    SaveQuickFilterSettings
    ' Call plot with built parameters
    ' Assumes PlotModelToNewDrawingWithFilters exists elsewhere
    PlotModelToNewDrawingWithFilters framesOnly, showNodeNames, showFrameNames, showShellNames, _
        sapOnlyMode, frameFilters, areaFilters
    If ControlExists("btnPlotToCAD") Then Me.btnPlotToCAD.enabled = True
    If ControlExists("btnImportToSAP") Then Me.btnImportToSAP.enabled = True
    AddLog "PLOT TO CAD: Completed successfully"
    Exit Sub
ErrHandler:
    If ControlExists("btnPlotToCAD") Then Me.btnPlotToCAD.enabled = True
    If ControlExists("btnImportToSAP") Then Me.btnImportToSAP.enabled = True
    AddLog "ERROR in Plot to CAD: " & err.description
    MsgBox "Error: " & err.description, vbCritical, "Plot Error"
End Sub

' ----------------------------
' Other existing handlers (safe-check controls)
' ----------------------------
Private Sub btnImportToSAP_Click()
    On Error GoTo ErrHandler
    Dim tolerance   As Double
    Dim scaleFactor As Double
    tolerance = 1
    scaleFactor = 1
    On Error Resume Next
    If ControlExists("txtTolerance") Then tolerance = val(Me.txtTolerance.text)
    If ControlExists("txtScaleFactor") Then scaleFactor = val(Me.txtScaleFactor.text)
    On Error GoTo ErrHandler
    If tolerance <= 0 Then tolerance = 1
    If scaleFactor <= 0 Then scaleFactor = 1
    AddLog "========================================="
    AddLog "IMPORT TO SAP: Started (selection-based)"
    AddLog "  Settings: Tolerance=" & tolerance & " mm, Scale=" & scaleFactor
    AddLog "======================================="
    DoEvents
    If ControlExists("btnPlotToCAD") Then Me.btnPlotToCAD.enabled = False
    If ControlExists("btnImportToSAP") Then Me.btnImportToSAP.enabled = False
    Dim acadApp    As Object
    Dim acadDoc    As Object
    Set acadApp = Core_Utils.GetOrCreateAutoCAD()
    If acadApp Is Nothing Then
        MsgBox "AutoCAD not available.", vbExclamation, "Import to SAP"
        GoTo RestoreButtons
    End If
    If acadApp.Documents.count = 0 Then
        Set acadDoc = acadApp.Documents.Add
    Else
        Set acadDoc = acadApp.ActiveDocument
    End If
    If Not IsObject(SapModel) Then
        If Not ConnectSAP2000() Then
            MsgBox "SAP2000 connection required.", vbExclamation, "Import to SAP"
            GoTo RestoreButtons
        End If
    End If
    Core_Sync_Manager.ImportSelectedEntitiesToSAP acadDoc, SapModel, tolerance, scaleFactor
RestoreButtons:
    If ControlExists("btnPlotToCAD") Then Me.btnPlotToCAD.enabled = True
    If ControlExists("btnImportToSAP") Then Me.btnImportToSAP.enabled = True
    AddLog "IMPORT TO SAP: Completed"
    Exit Sub
ErrHandler:
    If ControlExists("btnPlotToCAD") Then Me.btnPlotToCAD.enabled = True
    If ControlExists("btnImportToSAP") Then Me.btnImportToSAP.enabled = True
    AddLog "ERROR in Import to SAP: " & err.description
    MsgBox "Error: " & err.description, vbCritical, "Import Error"
End Sub

Private Sub btnImportFloorPlan_Click()
    On Error GoTo ErrHandler
    Dim tolerance   As Double
    Dim scaleFactor As Double
    tolerance = 1
    scaleFactor = 1
    On Error Resume Next
    If ControlExists("txtTolerance") Then tolerance = val(Me.txtTolerance.text)
    If ControlExists("txtScaleFactor") Then scaleFactor = val(Me.txtScaleFactor.text)
    On Error GoTo ErrHandler
    If tolerance <= 0 Then tolerance = 1
    If scaleFactor <= 0 Then scaleFactor = 1
    Dim zElevation As String
    zElevation = InputBox("Enter Z elevation (mm) for floor plan placement:", "Floor Plan Elevation", "0")
    If zElevation = "" Then
        AddLog "Floor plan import cancelled by user."
        Exit Sub
    End If
    Dim zValue As Double
    zValue = val(zElevation)
    If ControlExists("btnPlotToCAD") Then Me.btnPlotToCAD.enabled = False
    If ControlExists("btnImportToSAP") Then Me.btnImportToSAP.enabled = False
    If ControlExists("btnImportFloorPlan") Then Me.btnImportFloorPlan.enabled = False
    Dim acadApp As Object, acadDoc As Object
    Set acadApp = Core_Utils.GetOrCreateAutoCAD()
    If acadApp Is Nothing Then
        MsgBox "AutoCAD not available.", vbExclamation, "Import Floor Plan"
        GoTo RestoreButtons2
    End If
    If acadApp.Documents.count = 0 Then
        Set acadDoc = acadApp.Documents.Add
    Else
        Set acadDoc = acadApp.ActiveDocument
    End If
    If Not IsObject(SapModel) Then
        If Not ConnectSAP2000() Then
            MsgBox "SAP2000 connection required.", vbExclamation, "Import Floor Plan"
            GoTo RestoreButtons2
        End If
    End If
    Core_Sync_Manager.ImportFloorPlanFromCAD acadDoc, SapModel, zValue, tolerance, scaleFactor
RestoreButtons2:
    If ControlExists("btnPlotToCAD") Then Me.btnPlotToCAD.enabled = True
    If ControlExists("btnImportToSAP") Then Me.btnImportToSAP.enabled = True
    If ControlExists("btnImportFloorPlan") Then Me.btnImportFloorPlan.enabled = True
    AddLog "IMPORT FLOOR PLAN: Completed"
    Exit Sub
ErrHandler:
    If ControlExists("btnPlotToCAD") Then Me.btnPlotToCAD.enabled = True
    If ControlExists("btnImportToSAP") Then Me.btnImportToSAP.enabled = True
    If ControlExists("btnImportFloorPlan") Then Me.btnImportFloorPlan.enabled = True
    AddLog "ERROR in Import Floor Plan: " & err.description
    MsgBox "Error: " & err.description, vbCritical, "Import Floor Plan Error"
End Sub

Private Sub chkRealTimeSync_Click()
    On Error GoTo ErrHandler
    If Not ControlExists("chkRealTimeSync") Then Exit Sub
    If Me.chkRealTimeSync.Value Then
        Dim acadApp As Object, acadDoc As Object
        Set acadApp = Core_Utils.GetOrCreateAutoCAD()
        If acadApp Is Nothing Then
            MsgBox "AutoCAD not available.", vbExclamation, "Real-Time Sync"
            Me.chkRealTimeSync.Value = False
            Exit Sub
        End If
        If acadApp.Documents.count = 0 Then
            Set acadDoc = acadApp.Documents.Add
        Else
            Set acadDoc = acadApp.ActiveDocument
        End If
        If Not IsObject(SapModel) Then
            If Not ConnectSAP2000() Then
                MsgBox "SAP2000 connection required.", vbExclamation, "Real-Time Sync"
                Me.chkRealTimeSync.Value = False
                Exit Sub
            End If
        End If
        Core_Event_Handler.EnableRealTimeSync acadDoc, SapModel
        AddLog "Real-time AUTO-SYNC ACTIVE"
    Else
        Core_Event_Handler.DisableRealTimeSync
        AddLog "Real-time sync DISABLED"
    End If
    UpdateSyncStatusLabel
    Exit Sub
ErrHandler:
    Me.chkRealTimeSync.Value = False
    AddLog "ERROR enabling real-time sync: " & err.description
    MsgBox "Error: " & err.description, vbCritical, "Real-Time Sync Error"
End Sub

Private Sub btnClearLog_Click()
    On Error Resume Next
    If ControlExists("txtStatus") Then Me.txtStatus.text = "Log cleared." & vbCrLf & "Ready for next operation."
End Sub

Private Sub btnClose_Click()
    ' Disable real-time sync before closing
    On Error Resume Next
    ' Safely call IsSyncEnabled / DisableRealTimeSync
    On Error Resume Next
    If Core_Event_Handler.IsSyncEnabled() Then
        Core_Event_Handler.DisableRealTimeSync
    End If
    If err.number <> 0 Then err.Clear
    On Error GoTo 0
    Unload Me
End Sub

Private Sub UpdateSyncStatusLabel()
    On Error Resume Next
    If Not ControlExists("lblSyncStatus") Then Exit Sub
    Dim syncEnabled As Boolean
    syncEnabled = False
    ' Safely query Core_Event_Handler.IsSyncEnabled()
    On Error Resume Next
    syncEnabled = Core_Event_Handler.IsSyncEnabled()
    If err.number <> 0 Then
        err.Clear
        syncEnabled = False
    End If
    On Error GoTo 0
    If syncEnabled Then
        Me.lblSyncStatus.Caption = "AUTO-SYNC ACTIVE"
        Me.lblSyncStatus.ForeColor = RGB(0, 128, 0)
    Else
        Me.lblSyncStatus.Caption = "Manual sync only"
        Me.lblSyncStatus.ForeColor = RGB(128, 128, 128)
    End If
    On Error GoTo 0
End Sub

Public Sub SetStatus(msg As String)
    AddLog msg
End Sub

' Centralized AddLog that ensures txtStatus exists
Private Sub AddLog(msg As String)
    On Error Resume Next
    Dim timestampMsg As String
    If InStr(msg, "ERROR") > 0 Or InStr(msg, "Started") > 0 Or InStr(msg, "ACTIVE") > 0 Then
        timestampMsg = "[" & Format(Now, "hh:nn:ss") & "] " & msg
    Else
        timestampMsg = msg
    End If
    If ControlExists("txtStatus") Then
        Me.txtStatus.text = timestampMsg & vbCrLf & Me.txtStatus.text
        If Len(Me.txtStatus.text) > 5000 Then
            Me.txtStatus.text = Left(Me.txtStatus.text, 5000) & vbCrLf & "... (log truncated)"
        End If
    End If
    DoEvents
    On Error GoTo 0
End Sub

