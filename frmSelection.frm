VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSelection 
   Caption         =   "SAP2000 Smart Selection Tool"
   ClientHeight    =   4816
   ClientLeft      =   120
   ClientTop       =   675
   ClientWidth     =   7365
   OleObjectBlob   =   "frmSelection.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ===============================================================
' Userform: m02_sap_selection_tool (UserForm code-behind)
' Purpose: Smart selection operations for SAP2000 - Redesigned
' Notes:
' - Uses InputListStore accessor from m02_sap_window_manager for shared storage.
' - Uses EnsureSapModelAvailable before calling SAP API.
' - API calls check return codes (ret) and counts (cnt) before using results.
' - SelectSingleObject uses two-parameter SetSelected calls.
' - UI: txtInputList is MultiLine/WordWrap with vertical scrollbar.
' - lstPropertyValues uses Extended multi-select (Ctrl to add).
' - CleanText normalizes newlines and removes control characters.
' - lblTotalCount is created dynamically and updated after any list change.
' - Node can be toggled independently and combined with ONE non-node type.
' - Non-node types are mutually exclusive (only one can be selected).
' - Apply with both Node+NonNode will expand to nodes (add nodes of elements).
' - btnReplace replaced with btnUnselect (deselect objects in SAP2000).
' ===============================================================
Option Explicit

' Gridline map: axis -> dictionary(coordString -> gridName)
Private gGridlineMap As Object

' Master list of property values currently loaded (unfiltered)
Private gPropertyValuesMaster As Collection

' Temporary store of original txtInputList text when user uses the input-list filter.
' Used so we can restore original content when the filter is cleared.
Private gInputListMasterText As String

' When a filter is applied we also keep the expanded collection of items that matched the filter
' (this represents the "filtered original" state before the user edits the visible filtered list).
Private gInputListFilterMasterCol As Collection

' Flag indicating the user edited the visible (filtered) txtInputList while a filter was active.
Private gInputListFilteredEdited As Boolean

' Store the edited visible filtered text (so we can merge edits back into the master when filter is cleared).
Private gInputListEditedText As String

' Unique button display mode: True = compact (ranges), False = expanded (explicit list)
Private gUniqueDisplayCompact As Boolean



' ---------------- Helper: ensure SapModel is connected ----------------
Private Function EnsureSapModelAvailable() As Boolean
    On Error Resume Next
    If Not (Not SapModel Is Nothing) Then
        ' Try to connect via central ConnectSAP2000
        If Not ConnectSAP2000() Then
            EnsureSapModelAvailable = False
            Exit Function
        End If
    End If
    EnsureSapModelAvailable = Not (SapModel Is Nothing)
End Function

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnReplaceInput_Click()
    If Not AnyPropertySelected() Then Exit Sub
    Dim elements    As Collection
    Set elements = GetElementsBySelectedPropertiesUnion()
    If elements.count = 0 Then MsgBox "No elements found", vbInformation: Exit Sub
    txtInputList.text = CleanText(CollectionToCompactText(elements))
    SaveCurrentInputList
    UpdateInputListCount
End Sub
' Add tooltips to all CommandButton controls on the UserForm.
' This helper sets a meaningful ControlTipText for each known button name,
' and falls back to a generic tooltip for any other CommandButton controls.
Private Sub SetAllButtonTooltips()
    On Error Resume Next
    Dim ctrl As MSForms.Control
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "CommandButton" Then
            Select Case LCase(ctrl.Name)
                Case "btnreplaceinput"
                    ctrl.ControlTipText = "Replace input list using selected property values."
                Case "btnaddtoinput"
                    ctrl.ControlTipText = "Add elements matching selected properties to the input list."
                Case "btnunselect"
                    ctrl.ControlTipText = "Deselect objects in SAP2000 based on the input list."
                Case "btnremovefrominput"
                    ctrl.ControlTipText = "Remove elements matching selected properties from the input list."
                Case "btnintersectinput"
                    ctrl.ControlTipText = "Intersect the input list with elements matching selected properties."
                Case "btnall"
                    ctrl.ControlTipText = "Select all objects in the model."
                Case "btnnone"
                    ctrl.ControlTipText = "Clear the current selection in SAP2000."
                Case "btninversion"
                    ctrl.ControlTipText = "Invert the current selection in SAP2000."
                Case "btnprevious"
                    ctrl.ControlTipText = "Select the previous selection in SAP2000."
                Case "btnapplytosap"
                    ctrl.ControlTipText = "Apply the input list selection to SAP2000 (add to current selection)."
                Case "btngetfromsap"
                    ctrl.ControlTipText = "Load the current SAP2000 selection into the input list."
                Case "btnsort"
                    ctrl.ControlTipText = "Sort and compact the input list."
                Case "btnclear"
                    ctrl.ControlTipText = "Clear the input list."
                Case "btnunique"
                    ctrl.ControlTipText = "Remove duplicate entries and compact the input list."
                Case "btntoggleshrinkexcel"
                    ctrl.ControlTipText = "Toggle Excel window shrink/restore."
                Case "btnclose"
                    ctrl.ControlTipText = "Close the tool and restore Excel windows."
                Case Else
                    ' Generic tooltip for any other command buttons
                    ctrl.ControlTipText = "Click to activate: " & ctrl.Name
            End Select
        End If
    Next ctrl
    On Error GoTo 0
End Sub

Sub UserForm_Initialize()

    ' ApplyWindowManagementWithCaption
    'ApplyWindowManagementWithCaption Me.Caption
'    AddButtons GetFormHandle(Me.Caption), WS_MINIMIZEBOX    ' Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX if needed
'    StartTopmostTimer

    ' Set default object type
    optNode.Value = False
    optFrame.Value = True
    chkIncludeGroup.Value = False

    ' Ensure GroupName separation so optNode can be independent of non-node options
    ' This prevents OptionButton mutual-exclusive behavior across node vs non-node
    On Error Resume Next
    optNode.groupName = "NodeGroup"
    optFrame.groupName = "NonNodeGroup"
    optCable.groupName = "NonNodeGroup"
    optTendon.groupName = "NonNodeGroup"
    optArea.groupName = "NonNodeGroup"
    optSolid.groupName = "NonNodeGroup"
    optLink.groupName = "NonNodeGroup"
    On Error GoTo 0

    ' Ensure central dictionary exists via accessor and create default per-type keys
    Dim tmpStore As Object
    On Error Resume Next
    Set tmpStore = gInputListStore
    On Error GoTo 0

    If tmpStore Is Nothing Then
        ' create store and default keys (7 types x 2 group flags)
        Set gInputListStore = CreateObject("Scripting.Dictionary")
        EnsureDefaultInputKeys
    Else
        ' ensure at least the default single-type keys exist for both group settings
        EnsureDefaultInputKeys
    End If

    On Error Resume Next
    Set tmpStore = gInputListStore
    On Error GoTo 0

    ' Configure txtInputList to wrap and be multiline with vertical scrollbar
    On Error Resume Next
    txtInputList.Multiline = True
    txtInputList.WordWrap = True
    txtInputList.ScrollBars = 2    ' fmScrollBarsVertical
    On Error GoTo 0

    ' Ensure InputList filter control exists (creates txtFilterInputList above/near txtInputList)
    On Error Resume Next
    Call EnsureInputListFilterControl
    On Error GoTo 0

    ' Make property list require Ctrl to add (Extended selection)
    On Error Resume Next
    lstPropertyValues.MultiSelect = 2    ' fmMultiSelectExtended
    On Error GoTo 0

    ' Initialize UI lists
    LoadAttributes

    ' Create or ensure a dynamic label to show total count
    EnsureTotalCountLabel

    ' Create checkbox for coordinate-plane selection if not present
    On Error Resume Next
    Dim chk         As MSForms.CheckBox
    Set chk = Nothing
    Set chk = Me.Controls("chkSelectByCoordinate")
    If chk Is Nothing Then
        Dim cb      As MSForms.CheckBox
        Set cb = Me.Controls.Add("Forms.CheckBox.1", "chkSelectByCoordinate", True)
        cb.Left = txtInputList.Left
        cb.Top = txtInputList.Top + txtInputList.height + 26
        cb.Width = txtInputList.Width
        cb.height = 18
        cb.Caption = "SelectByCoordinate"
        cb.Font.Size = 10
        cb.Value = False
        cb.enabled = False    ' disabled by default; will enable when attribute is coordinate
    Else
        chk.Value = False
        chk.enabled = False
    End If
    On Error GoTo 0

    ' Create filter TextBox for property values (if not present)
    On Error Resume Next
    Dim tb As MSForms.TextBox
    Set tb = Nothing
    Set tb = Me.Controls("txtFilterPropertyValues")
    If tb Is Nothing Then
        Dim lblFilt As MSForms.Label
        ' optional label for filter
        On Error Resume Next
        Set lblFilt = Me.Controls("lblFilterPropertyValues")
        If lblFilt Is Nothing Then
            Set lblFilt = Me.Controls.Add("Forms.Label.1", "lblFilterPropertyValues", True)
            lblFilt.Caption = "Filter:"
            lblFilt.Font.Size = 9
            ' position label above lstPropertyValues
            lblFilt.Left = lstPropertyValues.Left
            lblFilt.Top = lstPropertyValues.Top - 20
            lblFilt.Width = 36
            lblFilt.height = 18
        End If
        On Error GoTo 0
        Dim newTB As MSForms.TextBox
        Set newTB = Me.Controls.Add("Forms.TextBox.1", "txtFilterPropertyValues", True)
        newTB.Left = lstPropertyValues.Left + 40
        newTB.Top = lstPropertyValues.Top - 22
        newTB.Width = lstPropertyValues.Width - 40
        newTB.height = 18
        newTB.text = ""
        newTB.Font.Size = 9
    Else
        tb.text = ""
    End If
    On Error GoTo 0

    ' Load saved input list for this objectType+group combination (if any)
    LoadSavedInputList

    ' Initial update of the count
    UpdateInputListCount

    ' Ensure new Toggle button caption (if design-time control exists, keep as is)
    On Error Resume Next
    Me.Controls("btnToggleShrinkExcel").Caption = "ToggleExcelShrink"
    On Error GoTo 0
    
    ' enable tooltips for all buttons
    SetAllButtonTooltips
    
    ' Connect to SAP2000
    ConnectSAP2000
    SapModel.SetPresentUnits (5)
    
    
End Sub
'Private Sub UserForm_Activate()
'    ' When the form becomes active, minimize Excel so user only sees the form.
'    ' (This is existing behavior; leave it if desired)
'    MinimizeExcelWindowForForm
'End Sub
'Private Sub UserForm_Terminate()
'    SaveCurrentInputList
'    ' Restore Excel window if this form minimized it (minimize path)
'    RestoreExcelWindowAfterForm
'    ' Also restore if shrink path was used
'    RestoreShrunkExcelWindow
'    ' RestoreFormParentToExcel Me  ' optional depending on window manager
'End Sub

' ============ OBJECT TYPE MANAGEMENT ============
' Returns comma-separated list of selected types: "Node", "Frame", "Node,Frame", etc.
Private Function GetSelectedObjectTypes() As String
    Dim types       As String
    types = ""

    If optNode.Value Then
        types = "Node"
    End If

    ' Check non-node types (mutually exclusive)
    If optFrame.Value Then
        If types <> "" Then types = types & ","
        types = types & "Frame"
    End If
    If optCable.Value Then
        If types <> "" Then types = types & ","
        types = types & "Cable"
    End If
    If optTendon.Value Then
        If types <> "" Then types = types & ","
        types = types & "Tendon"
    End If
    If optArea.Value Then
        If types <> "" Then types = types & ","
        types = types & "Area"
    End If
    If optSolid.Value Then
        If types <> "" Then types = types & ","
        types = types & "Solid"
    End If
    If optLink.Value Then
        If types <> "" Then types = types & ","
        types = types & "Link"
    End If

    GetSelectedObjectTypes = types
End Function

' Returns the primary non-node type for attribute loading
Private Function GetPrimaryObjectType() As String
    If optFrame.Value Then GetPrimaryObjectType = "Frame": Exit Function
    If optCable.Value Then GetPrimaryObjectType = "Cable": Exit Function
    If optTendon.Value Then GetPrimaryObjectType = "Tendon": Exit Function
    If optArea.Value Then GetPrimaryObjectType = "Area": Exit Function
    If optSolid.Value Then GetPrimaryObjectType = "Solid": Exit Function
    If optLink.Value Then GetPrimaryObjectType = "Link": Exit Function
    If optNode.Value Then GetPrimaryObjectType = "Node": Exit Function
    GetPrimaryObjectType = "Frame"    ' default
End Function

Private Function IsNodeSelected() As Boolean
    IsNodeSelected = optNode.Value
End Function

Private Function GetNonNodeType() As String
    ' Returns the selected non-node type, or empty if none
    If optFrame.Value Then GetNonNodeType = "Frame": Exit Function
    If optCable.Value Then GetNonNodeType = "Cable": Exit Function
    If optTendon.Value Then GetNonNodeType = "Tendon": Exit Function
    If optArea.Value Then GetNonNodeType = "Area": Exit Function
    If optSolid.Value Then GetNonNodeType = "Solid": Exit Function
    If optLink.Value Then GetNonNodeType = "Link": Exit Function
    GetNonNodeType = ""
End Function

Private Sub SwitchObjectType()
    ' NOTE: We intentionally DO NOT call SaveCurrentInputList here because
    ' callers (MouseDown/overlay handlers) must save the list under the previous key
    ' before they change option values. This prevents overwriting the wrong slot.
    ' Reload attributes for the newly selected type and load saved list for the new key.
    LoadAttributes

    ' Ensure property list uses Ctrl-based multi-select (Extended)
    On Error Resume Next
    lstPropertyValues.MultiSelect = 2    ' fmMultiSelectExtended
    On Error GoTo 0

    ' Load saved input list for the new object type / combination (if any)
    LoadSavedInputList

    ' Update count label and UI
    UpdateInputListCount
End Sub

' Node can toggle independently - use MouseDown to allow toggling off
Private Sub optNode_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    On Error Resume Next
    ' Save current list for current key before any change
    SaveCurrentInputListForKey GetStoreKey()

    ' Toggle Node state
    If optNode.Value = True Then
        optNode.Value = False
    Else
        optNode.Value = True
    End If
    On Error GoTo 0

    SwitchObjectType
End Sub

' Non-node options: toggle behavior via MouseDown - clicking an already selected toggle it off
Private Sub ToggleNonNodeOption(ByVal selName As String)
    On Error Resume Next
    ' Save current list for current key before any change
    SaveCurrentInputListForKey GetStoreKey()

    Dim ctrl As MSForms.OptionButton
    Set ctrl = Me.Controls("opt" & selName)
    If ctrl Is Nothing Then Exit Sub
    If ctrl.Value = True Then
        ' currently selected => toggle off
        ctrl.Value = False
    Else
        ' select this and unselect others
        optFrame.Value = False: optCable.Value = False: optTendon.Value = False
        optArea.Value = False: optSolid.Value = False: optLink.Value = False
        ctrl.Value = True
    End If
    On Error GoTo 0

    SwitchObjectType
End Sub

Private Sub optFrame_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ToggleNonNodeOption "Frame"
End Sub
Private Sub optCable_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ToggleNonNodeOption "Cable"
End Sub
Private Sub optTendon_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ToggleNonNodeOption "Tendon"
End Sub
Private Sub optArea_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ToggleNonNodeOption "Area"
End Sub
Private Sub optSolid_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ToggleNonNodeOption "Solid"
End Sub
Private Sub optLink_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ToggleNonNodeOption "Link"
End Sub

' Keep Click events for compatibility with other flows (they will call EnforceNonNodeExclusivity and SwitchObjectType)
Private Sub optNode_Click()
    SwitchObjectType
End Sub
Private Sub optFrame_Click()
    EnforceNonNodeExclusivity "Frame"
    SwitchObjectType
End Sub
Private Sub optCable_Click()
    EnforceNonNodeExclusivity "Cable"
    SwitchObjectType
End Sub
Private Sub optTendon_Click()
    EnforceNonNodeExclusivity "Tendon"
    SwitchObjectType
End Sub
Private Sub optArea_Click()
    EnforceNonNodeExclusivity "Area"
    SwitchObjectType
End Sub
Private Sub optSolid_Click()
    EnforceNonNodeExclusivity "Solid"
    SwitchObjectType
End Sub
Private Sub optLink_Click()
    EnforceNonNodeExclusivity "Link"
    SwitchObjectType
End Sub

Private Sub chkIncludeGroup_Click()
    SwitchObjectType
End Sub

' Ensure only one non-node type is selected at a time
Private Sub EnforceNonNodeExclusivity(ByVal selectedType As String)
    ' Turn off all other non-node types
    If selectedType <> "Frame" Then optFrame.Value = False
    If selectedType <> "Cable" Then optCable.Value = False
    If selectedType <> "Tendon" Then optTendon.Value = False
    If selectedType <> "Area" Then optArea.Value = False
    If selectedType <> "Solid" Then optSolid.Value = False
    If selectedType <> "Link" Then optLink.Value = False
End Sub

' ============ LOAD ATTRIBUTES LIST ============
Private Sub LoadAttributes()
    Dim objType     As String
    objType = GetPrimaryObjectType()

    lstAttributes.Clear
    lstPropertyValues.Clear

    lstAttributes.AddItem "Name"
    If chkIncludeGroup.Value Then lstAttributes.AddItem "Group"

    Select Case objType
        Case "Node"
            lstAttributes.AddItem "Constraint"
            lstAttributes.AddItem "Support"
            lstAttributes.AddItem "Spring"            ' Node can have joint springs
            lstAttributes.AddItem "Coordinate-X"
            lstAttributes.AddItem "Coordinate-Y"
            lstAttributes.AddItem "Coordinate-Z"

        Case "Frame"
            lstAttributes.AddItem "Section"
            lstAttributes.AddItem "Material"
            lstAttributes.AddItem "Design-Type"
            lstAttributes.AddItem "AutoSection"
            lstAttributes.AddItem "Spring"            ' frame springs (line/frame)
            lstAttributes.AddItem "Coordinate-X"
            lstAttributes.AddItem "Coordinate-Y"
            lstAttributes.AddItem "Coordinate-Z"

        Case "Cable"
            lstAttributes.AddItem "Section"
            lstAttributes.AddItem "Material"
            lstAttributes.AddItem "Spring"            ' cable/tendon are line-like -> line springs
            lstAttributes.AddItem "Coordinate-X"
            lstAttributes.AddItem "Coordinate-Y"
            lstAttributes.AddItem "Coordinate-Z"

        Case "Tendon"
            lstAttributes.AddItem "Section"
            lstAttributes.AddItem "Material"
            lstAttributes.AddItem "Spring"
            lstAttributes.AddItem "Coordinate-X"
            lstAttributes.AddItem "Coordinate-Y"
            lstAttributes.AddItem "Coordinate-Z"

        Case "Area"
            lstAttributes.AddItem "Section"
            lstAttributes.AddItem "Material"
            lstAttributes.AddItem "Area-Type"
            lstAttributes.AddItem "Spring"            ' area springs
            lstAttributes.AddItem "Coordinate-X"
            lstAttributes.AddItem "Coordinate-Y"
            lstAttributes.AddItem "Coordinate-Z"

        Case "Solid"
            lstAttributes.AddItem "Section"
            lstAttributes.AddItem "Material"
            lstAttributes.AddItem "Spring"            ' solid springs
            lstAttributes.AddItem "Coordinate-X"
            lstAttributes.AddItem "Coordinate-Y"
            lstAttributes.AddItem "Coordinate-Z"

        Case "Link"
            lstAttributes.AddItem "Section"
            lstAttributes.AddItem "Property"
            lstAttributes.AddItem "Spring"            ' link-specific springs
            lstAttributes.AddItem "Coordinate-X"
            lstAttributes.AddItem "Coordinate-Y"
            lstAttributes.AddItem "Coordinate-Z"

    End Select
End Sub

' ============ ATTRIBUTE SELECTION ============
Private Sub lstAttributes_Click()
    If lstAttributes.ListIndex < 0 Then Exit Sub
    Dim selectedAttr As String
    Dim objType     As String
    selectedAttr = lstAttributes.text
    objType = GetPrimaryObjectType()
    LoadPropertyValues selectedAttr, objType

    ' Enable the coordinate checkbox only when attribute is coordinate type
    On Error Resume Next
    Dim cb          As MSForms.CheckBox
    Set cb = Me.Controls("chkSelectByCoordinate")
    If Not cb Is Nothing Then
        If InStr(UCase(selectedAttr), "COORDINATE-") > 0 Then
            cb.enabled = True
        Else
            cb.Value = False
            cb.enabled = False
        End If
    End If
    On Error GoTo 0
End Sub

Private Sub LoadPropertyValues(ByVal attrName As String, ByVal objType As String)
    ' Ensure SAP model exists
    If Not EnsureSapModelAvailable() Then
        lstPropertyValues.Clear
        ' For coordinate attributes do NOT add "Any" (prevent user selecting invalid "Any")
        If InStr(UCase(attrName), "COORDINATE-") = 0 Then
            lstPropertyValues.AddItem "Any"
        End If
        ' store master and apply sorting/filter
        StoreMasterPropertyValues
        ApplyFilterAndSortPropertyValues
        Exit Sub
    End If

    Dim ret As Long, cnt As Long
    Dim names()     As String
    Dim i           As Long

    lstPropertyValues.Clear
    ' Add "Any" only for non-coordinate attributes
    If InStr(UCase(attrName), "COORDINATE-") = 0 Then
        lstPropertyValues.AddItem "Any"
    End If

    Select Case attrName
        Case "Name"
            LoadObjectNames objType
        Case "Group"
            ret = SapModel.GroupDef.GetNameList(cnt, names)
            If ret = 0 And cnt > 0 Then
                For i = 0 To cnt - 1
                    If Trim$(names(i)) <> "" Then lstPropertyValues.AddItem names(i)
                Next i
            End If
        Case "Section"
            ' Primary population from property lists (API calls) with fallback.
            Select Case objType
                Case "Frame"
                    ret = SapModel.PropFrame.GetNameList(cnt, names)
                    If Not (ret = 0 And cnt > 0) Then
                        On Error Resume Next
                        ret = SapModel.PropFrame.GetAutoSelectList(cnt, names)
                        On Error GoTo 0
                    End If
                Case "Cable"
                    ret = SapModel.PropCable.GetNameList(cnt, names)
                Case "Tendon"
                    ret = SapModel.PropTendon.GetNameList(cnt, names)
                Case "Area"
                    ret = SapModel.PropArea.GetNameList(cnt, names)
                Case "Solid"
                    ret = SapModel.PropSolid.GetNameList(cnt, names)
                Case "Link"
                    ret = SapModel.PropLink.GetNameList(cnt, names)
            End Select

            If ret = 0 And cnt > 0 Then
                For i = 0 To cnt - 1
                    If Trim$(names(i)) <> "" Then lstPropertyValues.AddItem names(i)
                Next i
            Else
                Dim fallbackColl As Collection
                Set fallbackColl = GetDistinctPropertyValuesBySection(objType)
                Dim it As Variant
                For Each it In fallbackColl
                    lstPropertyValues.AddItem it
                Next it
            End If

            ' Force-add "None" for Frame / Area (shell) / Solid (so it always appears like "Any")
            If UCase(objType) = "FRAME" Or UCase(objType) = "AREA" Or UCase(objType) = "SOLID" Then
                Dim alreadyNone As Boolean
                alreadyNone = False
                For i = 0 To lstPropertyValues.ListCount - 1
                    If StrComp(Trim$(lstPropertyValues.List(i)), "None", vbTextCompare) = 0 Then
                        alreadyNone = True
                        Exit For
                    End If
                Next i
                If Not alreadyNone Then lstPropertyValues.AddItem "None"
            End If

        Case "Material"
            ret = SapModel.PropMaterial.GetNameList(cnt, names)
            If ret = 0 And cnt > 0 Then
                For i = 0 To cnt - 1
                    If Trim$(names(i)) <> "" Then lstPropertyValues.AddItem names(i)
                Next i
            End If
        Case "Constraint"
            ret = SapModel.ConstraintDef.GetNameList(cnt, names)
            If ret = 0 And cnt > 0 Then
                For i = 0 To cnt - 1
                    If Trim$(names(i)) <> "" Then lstPropertyValues.AddItem names(i)
                Next i
            End If
        Case "Support"
            If objType <> "Node" Then
                ' Support (restraint) applies only to nodes; for non-node types show only Any (handled above)
                lstPropertyValues.Clear
                If InStr(UCase(attrName), "COORDINATE-") = 0 Then lstPropertyValues.AddItem "Any"
            Else
                PopulateSupportPropertyValuesForNodes
            End If
        Case "Spring"
            LoadSpringTypes objType
        Case "Design-Type"
            lstPropertyValues.AddItem "Program Determined"
            lstPropertyValues.AddItem "Column"
            lstPropertyValues.AddItem "Beam"
            lstPropertyValues.AddItem "Brace"
        Case "AutoSection"
            On Error Resume Next
            ret = SapModel.PropFrame.GetAutoSelectList(cnt, names)
            On Error GoTo 0
            If ret = 0 And cnt > 0 Then
                For i = 0 To cnt - 1
                    If Trim$(names(i)) <> "" Then lstPropertyValues.AddItem names(i)
                Next i
            Else
                ret = SapModel.PropFrame.GetNameList(cnt, names)
                If ret = 0 And cnt > 0 Then
                    For i = 0 To cnt - 1
                        If Trim$(names(i)) <> "" Then lstPropertyValues.AddItem names(i)
                    Next i
                End If
            End If
        Case "Area-Type"
            lstPropertyValues.AddItem "Shell"
            lstPropertyValues.AddItem "Plane"
            lstPropertyValues.AddItem "Asolid"
        Case "Property"
            ret = SapModel.PropLink.GetNameList(cnt, names)
            If ret = 0 And cnt > 0 Then
                For i = 0 To cnt - 1
                    If Trim$(names(i)) <> "" Then lstPropertyValues.AddItem names(i)
                Next i
            End If
        Case "Coordinate-X"
            If EnsureSapModelAvailable() Then Call AddCoordinateValuesToPropertyList("X", objType)
        Case "Coordinate-Y"
            If EnsureSapModelAvailable() Then Call AddCoordinateValuesToPropertyList("Y", objType)
        Case "Coordinate-Z"
            If EnsureSapModelAvailable() Then Call AddCoordinateValuesToPropertyList("Z", objType)
    End Select

    ' After populating lstPropertyValues, store master list and apply filter+sort so the UI shows sorted (A->Z) and filtered view.
    StoreMasterPropertyValues
    ApplyFilterAndSortPropertyValues
End Sub
Private Sub AddCoordinateValuesToPropertyList(ByVal axis As String, Optional ByVal objType As String = "Node")
    Dim cnt         As Long
    Dim names()     As String
    Dim ret         As Long
    Dim i           As Long
    Dim X As Double, Y As Double, Z As Double
    Dim dict        As Object
    Set dict = CreateObject("Scripting.Dictionary")

    ' Ensure gridline map is loaded
    EnsureGridlineMapInitialized

    ' Collect from nodes first
    On Error Resume Next
    ret = SapModel.pointObj.GetNameList(cnt, names)
    On Error GoTo 0
    If ret = 0 And cnt > 0 Then
        For i = 0 To cnt - 1
            If Trim$(names(i)) <> "" Then
                On Error Resume Next
                ret = SapModel.pointObj.GetCoordCartesian(names(i), X, Y, Z)
                On Error GoTo 0
                If ret = 0 Then
                    Dim val As Double, key As String
                    Select Case axis
                        Case "X": val = X
                        Case "Y": val = Y
                        Case "Z": val = Z
                    End Select
                    key = Format$(Round(val, 6), "0.######")
                    If Not dict.exists(key) Then dict.Add key, CDbl(key)
                End If
            End If
        Next i
    End If

    ' If requested, also collect representative coordinates from frames/areas/solids/cables/tendons/links
    Dim allElems    As Collection
    Set allElems = GetAllElements(objType)
    Dim elem        As Variant
    For Each elem In allElems
        Dim px As Double, py As Double, pz As Double
        px = 0: py = 0: pz = 0
        Dim success As Boolean: success = False

        If objType = "Frame" Or objType = "Cable" Or objType = "Tendon" Then
            ' try to get start/end point names for frame-like objects
            Dim spt As String, ept As String
            On Error Resume Next
            Select Case objType
                Case "Frame": ret = SapModel.frameObj.GetPoints(CStr(elem), spt, ept)
                Case "Cable": ret = SapModel.CableObj.GetPoints(CStr(elem), spt, ept)
                Case "Tendon": ret = SapModel.TendonObj.GetPoints(CStr(elem), spt, ept)
                Case Else: ret = -1
            End Select
            On Error GoTo 0
            If ret = 0 And Trim$(spt) <> "" Then
                Dim sx As Double, sy As Double, sz As Double
                On Error Resume Next
                ret = SapModel.pointObj.GetCoordCartesian(spt, sx, sy, sz)
                On Error GoTo 0
                If ret = 0 Then
                    If Trim$(ept) <> "" Then
                        Dim ex As Double, ey As Double, ez As Double
                        On Error Resume Next
                        ret = SapModel.pointObj.GetCoordCartesian(ept, ex, ey, ez)
                        On Error GoTo 0
                        If ret = 0 Then
                            px = (sx + ex) / 2: py = (sy + ey) / 2: pz = (sz + ez) / 2
                            success = True
                        Else
                            px = sx: py = sy: pz = sz: success = True
                        End If
                    Else
                        px = sx: py = sy: pz = sz: success = True
                    End If
                End If
            End If
        ElseIf objType = "Area" Then
            ' try to get point list for area and compute centroid
            Dim numPts As Long
            Dim ptNames() As String
            On Error Resume Next
            ret = SapModel.AreaObj.GetPoints(CStr(elem), numPts, ptNames)
            On Error GoTo 0
            If ret = 0 And numPts > 0 Then
                Dim k As Long, cntPts As Long
                cntPts = 0
                px = 0: py = 0: pz = 0
                For k = 0 To numPts - 1
                    If Trim$(ptNames(k)) <> "" Then
                        Dim tx As Double, ty As Double, tz As Double
                        On Error Resume Next
                        ret = SapModel.pointObj.GetCoordCartesian(ptNames(k), tx, ty, tz)
                        On Error GoTo 0
                        If ret = 0 Then
                            px = px + tx: py = py + ty: pz = pz + tz
                            cntPts = cntPts + 1
                        End If
                    End If
                Next k
                If cntPts > 0 Then
                    px = px / cntPts: py = py / cntPts: pz = pz / cntPts
                    success = True
                End If
            End If
        ElseIf objType = "Solid" Then
            ' Solid: get point list for solid and compute centroid (use Variant for returned array)
            Dim sPtNames() As String
            On Error Resume Next
            ret = SapModel.SolidObj.GetPoints(CStr(elem), sPtNames)
            On Error GoTo 0
            If ret = 0 Then
                If IsArray(sPtNames) Then
                    Dim sIdx As Long, scnt As Long
                    scnt = 0
                    px = 0: py = 0: pz = 0
                    For sIdx = LBound(sPtNames) To UBound(sPtNames)
                        If Trim$(CStr(sPtNames(sIdx))) <> "" Then
                            Dim tx2 As Double, ty2 As Double, tz2 As Double
                            On Error Resume Next
                            ret = SapModel.pointObj.GetCoordCartesian(sPtNames(sIdx), tx2, ty2, tz2)
                            On Error GoTo 0
                            If ret = 0 Then
                                px = px + tx2: py = py + ty2: pz = pz + tz2
                                scnt = scnt + 1
                            End If
                        End If
                    Next sIdx
                    If scnt > 0 Then
                        px = px / scnt: py = py / scnt: pz = pz / scnt
                        success = True
                    End If
                End If
            End If
        ElseIf objType = "Link" Then
            ' attempt to use link points (similar to frame)
            Dim l1 As String, l2 As String
            On Error Resume Next
            ret = SapModel.LinkObj.GetPoints(CStr(elem), l1, l2)
            On Error GoTo 0
            If ret = 0 And Trim$(l1) <> "" Then
                Dim lx As Double, ly As Double, lz As Double
                On Error Resume Next
                ret = SapModel.pointObj.GetCoordCartesian(l1, lx, ly, lz)
                On Error GoTo 0
                If ret = 0 Then
                    If Trim$(l2) <> "" Then
                        Dim lx2 As Double, ly2 As Double, lz2 As Double
                        On Error Resume Next
                        ret = SapModel.pointObj.GetCoordCartesian(l2, lx2, ly2, lz2)
                        On Error GoTo 0
                        If ret = 0 Then
                            px = (lx + lx2) / 2: py = (ly + ly2) / 2: pz = (lz + lz2) / 2
                            success = True
                        Else
                            px = lx: py = ly: pz = lz: success = True
                        End If
                    Else
                        px = lx: py = ly: pz = lz: success = True
                    End If
                End If
            End If
        End If

        If success Then
            Dim cval As Double, ckey As String
            Select Case axis
                Case "X": cval = px
                Case "Y": cval = py
                Case "Z": cval = pz
                Case Else: cval = px
            End Select
            ckey = Format$(Round(cval, 6), "0.######")
            If Not dict.exists(ckey) Then dict.Add ckey, CDbl(ckey)
        End If
    Next elem

    ' sort numeric keys
    If dict.count > 0 Then
        Dim arr()   As Double
        ReDim arr(1 To dict.count)
        Dim idx     As Long
        idx = 1
        Dim k2      As Variant
        For Each k2 In dict.keys
            arr(idx) = CDbl(dict(k2))
            idx = idx + 1
        Next k2
        Dim j As Long, tmp As Double
        For i = 1 To UBound(arr) - 1
            For j = i + 1 To UBound(arr)
                If arr(i) > arr(j) Then
                    tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
                End If
            Next j
        Next i
        For i = 1 To UBound(arr)
            ' Format numeric display
            Dim displayVal As String
            displayVal = Format$(arr(i), "0.######")
            ' Append gridline name if available for this axis + coordinate
            If Not gGridlineMap Is Nothing Then
                If gGridlineMap.exists(UCase(axis)) Then
                    Dim axisDict As Object
                    Set axisDict = gGridlineMap(UCase(axis))
                    If axisDict.exists(displayVal) Then
                        lstPropertyValues.AddItem displayVal & " (" & CStr(axisDict(displayVal)) & ")"
                    Else
                        lstPropertyValues.AddItem displayVal
                    End If
                Else
                    lstPropertyValues.AddItem displayVal
                End If
            Else
                lstPropertyValues.AddItem displayVal
            End If
        Next i
    End If
End Sub

Private Sub LoadObjectNames(ByVal objType As String)
    If Not EnsureSapModelAvailable() Then Exit Sub
    Dim ret As Long, cnt As Long, names() As String, i As Long
    Select Case objType
        Case "Node": ret = SapModel.pointObj.GetNameList(cnt, names)
        Case "Frame": ret = SapModel.frameObj.GetNameList(cnt, names)
        Case "Cable": ret = SapModel.CableObj.GetNameList(cnt, names)
        Case "Tendon": ret = SapModel.TendonObj.GetNameList(cnt, names)
        Case "Area": ret = SapModel.AreaObj.GetNameList(cnt, names)
        Case "Solid": ret = SapModel.SolidObj.GetNameList(cnt, names)
        Case "Link": ret = SapModel.LinkObj.GetNameList(cnt, names)
        Case Else: ret = -1
    End Select
    If ret = 0 And cnt > 0 Then
        For i = 0 To cnt - 1
            If Trim$(names(i)) <> "" Then lstPropertyValues.AddItem names(i)
        Next i
    End If
End Sub

Private Sub LoadSupportTypes(ByVal objType As String)
    ' Caller should clear lstPropertyValues before calling or we clear here
    lstPropertyValues.Clear
    lstPropertyValues.AddItem "Any"

    Select Case objType
        Case "Node"
            lstPropertyValues.AddItem "Fixed"
            lstPropertyValues.AddItem "Pinned"
            lstPropertyValues.AddItem "Roller"
        Case "Frame", "Cable", "Tendon"
            ' Frame-like objects: support selection usually maps to attached nodes
            lstPropertyValues.AddItem "Fixed"
            lstPropertyValues.AddItem "Pinned"
            lstPropertyValues.AddItem "Roller"
        Case "Area", "Solid", "Link"
            lstPropertyValues.AddItem "Fixed"
            lstPropertyValues.AddItem "Pinned"
            lstPropertyValues.AddItem "Roller"
        Case Else
            lstPropertyValues.AddItem "Fixed"
            lstPropertyValues.AddItem "Pinned"
            lstPropertyValues.AddItem "Roller"
    End Select
End Sub

' Populate spring options depending on object type
Private Sub LoadSpringTypes(ByVal objType As String)
    lstPropertyValues.Clear
    lstPropertyValues.AddItem "Any"

    Select Case objType
        Case "Node"
            ' Joint springs attached to points
            lstPropertyValues.AddItem "Joint Spring"
            lstPropertyValues.AddItem "Any Spring"
        Case "Frame", "Cable", "Tendon"
            ' Line-like springs (frame/tendon/cable)
            lstPropertyValues.AddItem "Line Spring"
            lstPropertyValues.AddItem "Frame Spring"
            lstPropertyValues.AddItem "Any Spring"
        Case "Area"
            lstPropertyValues.AddItem "Area Spring"
            lstPropertyValues.AddItem "Any Spring"
        Case "Solid"
            lstPropertyValues.AddItem "Solid Spring"
            lstPropertyValues.AddItem "Any Spring"
        Case "Link"
            lstPropertyValues.AddItem "Link Spring"
            lstPropertyValues.AddItem "Any Spring"
        Case Else
            lstPropertyValues.AddItem "Any Spring"
    End Select
End Sub

' Helper to collect distinct section/property by iterating objects
Private Function GetDistinctPropertyValuesBySection(ByVal objType As String) As Collection
    Dim result      As New Collection
    Dim allItems    As Collection
    Dim it          As Variant
    Dim propName As String, sAuto As String
    Dim added       As Object
    Set added = CreateObject("Scripting.Dictionary")
    Set allItems = GetAllElements(objType)
    For Each it In allItems
        If objType = "Frame" Then
            If SapModel.frameObj.GetSection(CStr(it), propName, sAuto) = 0 Then
                If Trim$(propName) <> "" Then
                    If Not added.exists(propName) Then
                        result.Add propName
                        added.Add propName, True
                    End If
                End If
            End If
        ElseIf objType = "Area" Then
            If SapModel.AreaObj.GetProperty(CStr(it), propName) = 0 Then
                If Trim$(propName) <> "" Then
                    If Not added.exists(propName) Then
                        result.Add propName
                        added.Add propName, True
                    End If
                End If
            End If
        End If
    Next it
    Set GetDistinctPropertyValuesBySection = result
End Function

' ============ TRANSFER BUTTONS ============
Private Sub btnAddToInput_Click()
    If Not AnyPropertySelected() Then Exit Sub
    Dim elements    As Collection
    Set elements = GetElementsBySelectedPropertiesUnion()
    If elements.count = 0 Then MsgBox "No elements found", vbInformation: Exit Sub
    Dim currentList As Collection
    Set currentList = ParseObjectList(txtInputList.text)
    Dim item        As Variant
    For Each item In elements
        If Not IsInCollection(currentList, CStr(item)) Then currentList.Add item
    Next item
    txtInputList.text = CleanText(CollectionToCompactText(currentList))
    SaveCurrentInputList
    UpdateInputListCount
End Sub

' ============ UNSELECT / DESELECT FROM SAP2000 ============
Private Sub btnUnselect_Click()
    ' Deselect objects in the input list from SAP2000 current selection
    If Not EnsureSapModelAvailable() Then MsgBox "SAP model not available", vbExclamation: Exit Sub

    Dim objList     As Collection
    Set objList = ParseObjectList(txtInputList.text)
    If objList.count = 0 Then Exit Sub

    Dim objName     As Variant
    Dim ret         As Long
    Dim successCount As Long
    successCount = 0

    ' Determine object types to deselect
    Dim nodeSelected As Boolean
    Dim nonNodeType As String
    nodeSelected = IsNodeSelected()
    nonNodeType = GetNonNodeType()

    ' If both Node and NonNode selected, first deselect non-node elements, then deselect nodes that belong to them
    If nodeSelected And nonNodeType <> "" Then
        Dim selectedElements As New Collection
        Dim elem    As Variant

        ' First pass: deselect non-node elements (based on names provided)
        For Each objName In objList
            ret = SelectSingleObject(nonNodeType, CStr(objName), False)
            selectedElements.Add CStr(objName)
            If ret = 0 Then successCount = successCount + 1
        Next objName

        ' Second pass: for each such element, get its nodes and deselect them
        Dim nodesDict As Object
        Set nodesDict = CreateObject("Scripting.Dictionary")
        For Each elem In selectedElements
            Dim nodesFromElement As Collection
            Set nodesFromElement = GetNodesFromElement(CStr(elem), nonNodeType)
            Dim n   As Variant
            For Each n In nodesFromElement
                If Not nodesDict.exists(CStr(n)) Then
                    nodesDict.Add CStr(n), True
                End If
            Next n
        Next elem

        Dim nodeName As Variant
        For Each nodeName In nodesDict.keys
            ret = SelectSingleObject("Node", CStr(nodeName), False)
            If ret = 0 Then successCount = successCount + 1
        Next nodeName

    ElseIf nonNodeType <> "" Then
        ' Only non-node type selected, deselect those elements
        For Each objName In objList
            ret = SelectSingleObject(nonNodeType, CStr(objName), False)
            If ret = 0 Then successCount = successCount + 1
        Next objName

    ElseIf nodeSelected Then
        ' Only Node selected, deselect nodes
        For Each objName In objList
            ret = SelectSingleObject("Node", CStr(objName), False)
            If ret = 0 Then successCount = successCount + 1
        Next objName
    End If

    SapModel.View.RefreshWindow

End Sub

Private Sub btnRemoveFromInput_Click()
    If Not AnyPropertySelected() Then Exit Sub
    Dim elementsToRemove As Collection
    Set elementsToRemove = GetElementsBySelectedPropertiesUnion()
    Dim currentList As Collection
    Set currentList = ParseObjectList(txtInputList.text)
    Dim newList     As New Collection
    Dim item        As Variant
    For Each item In currentList
        If Not IsInCollection(elementsToRemove, CStr(item)) Then newList.Add item
    Next item
    txtInputList.text = CleanText(CollectionToCompactText(newList))
    SaveCurrentInputList
    UpdateInputListCount
End Sub

Private Sub btnIntersectInput_Click()
    If Not AnyPropertySelected() Then Exit Sub
    Dim intersection As Collection
    Set intersection = GetElementsBySelectedPropertiesIntersection()
    Dim currentList As Collection
    Set currentList = ParseObjectList(txtInputList.text)
    Dim newList     As New Collection
    Dim item        As Variant
    For Each item In currentList
        If IsInCollection(intersection, CStr(item)) Then newList.Add item
    Next item
    txtInputList.text = CleanText(CollectionToCompactText(newList))
    SaveCurrentInputList
    UpdateInputListCount
End Sub

Private Function GetElementsByPropertyValue(ByVal attrName As String, ByVal propValue As String, ByVal objType As String) As Collection
    ' Returns a collection of element names that match the given attribute value for the specified object type.
    Dim result      As New Collection
    Dim ret         As Long
    Dim cnt         As Long
    Dim names()     As String
    Dim i           As Long

    ' Normalize propValue for comparisons
    Dim pvNorm As String
    pvNorm = Trim$(CStr(propValue))

    ' "Any" means return all elements of the object type
    ' But for attributes that are special (Support / Spring), handle in their case blocks.
    If StrComp(LCase$(pvNorm), "any", vbTextCompare) = 0 Then
        If LCase$(attrName) <> "support" And LCase$(attrName) <> "spring" Then
            Set result = GetAllElements(objType)
            Set GetElementsByPropertyValue = result
            Exit Function
        End If
    End If

    Select Case attrName
        Case "Name"
            result.Add propValue

        Case "Group"
            Dim savedCount As Long
            Dim savedObjTypeArr() As Long
            Dim savedObjNameArr() As String
            savedCount = 0
            SaveSelectionState savedCount, savedObjTypeArr, savedObjNameArr

            ret = SapModel.SelectObj.ClearSelection()
            On Error Resume Next
            ret = SapModel.SelectObj.Group(propValue, False)
            On Error GoTo 0
            Set result = GetSelectedElements(objType)

            RestoreSelectionState savedCount, savedObjTypeArr, savedObjNameArr

        Case "Section"
            ' If user selected "None" or an empty string, return elements that have no section/property assigned.
            If pvNorm = "" Or StrComp(LCase$(pvNorm), "none", vbTextCompare) = 0 Then
                Dim allElems As Collection
                Set allElems = GetAllElements(objType)
                Dim e As Variant
                Dim pName As String, pAuto As String

                For Each e In allElems
                    On Error Resume Next
                    If UCase(objType) = "FRAME" Then
                        ret = SapModel.frameObj.GetSection(CStr(e), pName, pAuto)
                        On Error GoTo 0
                        If ret = 0 Then
                            If Trim$(pName) = "" Or StrComp(LCase$(Trim$(pName)), "none", vbTextCompare) = 0 Then
                                result.Add CStr(e)
                            End If
                        End If
                    ElseIf UCase(objType) = "AREA" Then
                        ret = SapModel.AreaObj.GetProperty(CStr(e), pName)
                        On Error GoTo 0
                        If ret = 0 Then
                            If Trim$(pName) = "" Or StrComp(LCase$(Trim$(pName)), "none", vbTextCompare) = 0 Then
                                result.Add CStr(e)
                            End If
                        End If
                    ElseIf UCase(objType) = "SOLID" Then
                        ' Some SAP API versions may not expose SolidObj.GetProperty; handle defensively.
                        ret = -1
                        On Error Resume Next
                        ret = SapModel.SolidObj.GetProperty(CStr(e), pName)
                        On Error GoTo 0
                        If ret = 0 Then
                            If Trim$(pName) = "" Or StrComp(LCase$(Trim$(pName)), "none", vbTextCompare) = 0 Then
                                result.Add CStr(e)
                            End If
                        End If
                    ElseIf UCase(objType) = "CABLE" Or UCase(objType) = "TENDON" Or UCase(objType) = "LINK" Then
                        ' Try appropriate GetProperty/GetSection if available; fall back to skip.
                        On Error Resume Next
                        Select Case UCase(objType)
                            Case "CABLE"
                                ret = SapModel.CableObj.GetProperty(CStr(e), pName)
                            Case "TENDON"
                                ret = SapModel.TendonObj.GetProperty(CStr(e), pName)
                            Case "LINK"
                                ret = SapModel.LinkObj.GetProperty(CStr(e), pName)
                            Case Else
                                ret = -1
                        End Select
                        On Error GoTo 0
                        If ret = 0 Then
                            If Trim$(pName) = "" Or StrComp(LCase$(Trim$(pName)), "none", vbTextCompare) = 0 Then
                                result.Add CStr(e)
                            End If
                        End If
                    End If
                Next e

                Set GetElementsByPropertyValue = result
                Exit Function
            End If

            ' Otherwise try to use selection-by-property API (fallback to enumerating if selection returns nothing)
            Dim sSavedCnt As Long
            Dim sSavedTypes() As Long
            Dim sSavedNames() As String
            sSavedCnt = 0
            SaveSelectionState sSavedCnt, sSavedTypes, sSavedNames

            ret = SapModel.SelectObj.ClearSelection()
            Select Case objType
                Case "Frame": On Error Resume Next: ret = SapModel.SelectObj.PropertyFrame(propValue, False): On Error GoTo 0
                Case "Cable": On Error Resume Next: ret = SapModel.SelectObj.PropertyCable(propValue, False): On Error GoTo 0
                Case "Tendon": On Error Resume Next: ret = SapModel.SelectObj.PropertyTendon(propValue, False): On Error GoTo 0
                Case "Area": On Error Resume Next: ret = SapModel.SelectObj.PropertyArea(propValue, False): On Error GoTo 0
                Case "Solid": On Error Resume Next: ret = SapModel.SelectObj.PropertySolid(propValue, False): On Error GoTo 0
                Case "Link": On Error Resume Next: ret = SapModel.SelectObj.PropertyLink(propValue, False): On Error GoTo 0
            End Select
            Set result = GetSelectedElements(objType)
            If result.count = 0 Then
                ' Fallbacks for Frame/Area: iterate elements to find matches by property if API selection returned nothing
                If objType = "Frame" Then
                    Set result = GetFramesWithSection(propValue)
                ElseIf objType = "Area" Then
                    Set result = GetAreasWithProperty(propValue)
                End If
            End If

            RestoreSelectionState sSavedCnt, sSavedTypes, sSavedNames

        Case "Material"
            Dim matSavedCnt As Long
            Dim matSavedTypes() As Long
            Dim matSavedNames() As String
            matSavedCnt = 0
            SaveSelectionState matSavedCnt, matSavedTypes, matSavedNames

            On Error Resume Next
            ret = SapModel.SelectObj.ClearSelection()
            ret = SapModel.SelectObj.PropertyMaterial(propValue, False)
            On Error GoTo 0
            Set result = GetSelectedElements(objType)

            RestoreSelectionState matSavedCnt, matSavedTypes, matSavedNames

        Case "Constraint"
            Dim cSavedCnt As Long
            Dim cSavedTypes() As Long
            Dim cSavedNames() As String
            cSavedCnt = 0
            SaveSelectionState cSavedCnt, cSavedTypes, cSavedNames

            On Error Resume Next
            ret = SapModel.SelectObj.ClearSelection()
            ret = SapModel.SelectObj.Constraint(propValue, False)
            On Error GoTo 0
            Set result = GetSelectedElements("Node")
            RestoreSelectionState cSavedCnt, cSavedTypes, cSavedNames

            If result.count = 0 Then
                Dim numItems As Long, pointNames() As String, constraintNames() As String, j As Long
                ret = SapModel.pointObj.GetConstraint("ALL", numItems, pointNames, constraintNames, 1)
                If ret = 0 And numItems > 0 Then
                    For j = 0 To numItems - 1
                        If Trim$(constraintNames(j)) = propValue Then result.Add pointNames(j)
                    Next j
                End If
            End If

        Case "Support"
            ' Support (restraint) applies only to nodes.
            Dim spSavedCnt As Long
            Dim spSavedTypes() As Long
            Dim spSavedNames() As String

            ' If objType is not Node, support (restraint) does not apply -> return only "Any" handling above.
            If UCase(objType) <> "NODE" Then
                Set GetElementsByPropertyValue = New Collection
                Exit Function
            End If

            ' Use SupportedPoints to quickly get candidate nodes that have any restraint (DOF all True)
            Dim DOF() As Boolean, si As Long
            ReDim DOF(5)
            For si = 0 To 5: DOF(si) = True: Next si

            spSavedCnt = 0
            SaveSelectionState spSavedCnt, spSavedTypes, spSavedNames

            On Error Resume Next
            SapModel.SelectObj.ClearSelection
            ' SelectRestraints = True, spring flags = False
            ret = SapModel.SelectObj.SupportedPoints(DOF, "Local", False, True, False, False, False, False, False)
            On Error GoTo 0

            Dim cand As Collection
            Set cand = GetSelectedElements("Node")

            RestoreSelectionState spSavedCnt, spSavedTypes, spSavedNames

            ' Now filter candidates by exact restraint patterns via GetRestraint
            Dim node As Variant
            For Each node In cand
                Dim restraints() As Boolean
                Dim rret As Long
                On Error Resume Next
                rret = SapModel.pointObj.GetRestraint(CStr(node), restraints)
                On Error GoTo 0
                If rret = 0 And IsArray(restraints) Then
                    Dim b(5) As Boolean
                    Dim lb As Long, ub As Long
                    lb = LBound(restraints): ub = UBound(restraints)
                    For i = 0 To 5
                        If i >= lb And i <= ub Then
                            b(i) = CBool(restraints(i))
                        Else
                            b(i) = False
                        End If
                    Next i

                    Select Case LCase(propValue)
                        Case "any"
                            If (b(0) Or b(1) Or b(2) Or b(3) Or b(4) Or b(5)) Then
                                result.Add CStr(node)
                            End If
                        Case "fixed"
                            If b(0) And b(1) And b(2) And b(3) And b(4) And b(5) Then
                                result.Add CStr(node)
                            End If
                        Case "pinned"
                            If (b(0) And b(1) And b(2)) And Not (b(3) Or b(4) Or b(5)) Then
                                result.Add CStr(node)
                            End If
                        Case "roller"
                            Dim transCount As Long
                            transCount = 0
                            For i = 0 To 2
                                If b(i) Then transCount = transCount + 1
                            Next i
                            If transCount = 1 And Not (b(3) Or b(4) Or b(5)) Then
                                result.Add CStr(node)
                            End If
                        Case "other restraint"
                            Dim hasAny As Boolean
                            hasAny = (b(0) Or b(1) Or b(2) Or b(3) Or b(4) Or b(5))
                            Dim isFixed As Boolean, isPinned As Boolean, isRoller As Boolean
                            isFixed = (b(0) And b(1) And b(2) And b(3) And b(4) And b(5))
                            isPinned = ((b(0) And b(1) And b(2)) And Not (b(3) Or b(4) Or b(5)))
                            transCount = 0
                            For i = 0 To 2
                                If b(i) Then transCount = transCount + 1
                            Next i
                            isRoller = (transCount = 1 And Not (b(3) Or b(4) Or b(5)))
                            If hasAny And Not (isFixed Or isPinned Or isRoller) Then
                                result.Add CStr(node)
                            End If
                        Case Else
                            ' Unknown subtype -> do nothing
                    End Select
                End If
            Next node

        Case "Design-Type"
            ' Handle frame design type: "Program Determined", "Column", "Beam", "Brace"
            ' Only meaningful for Frame object type; return empty for others.
            If objType <> "Frame" Then
                Set GetElementsByPropertyValue = result
                Exit Function
            End If

            Dim allFrames As Collection
            Dim fr As Variant
            Dim sectionName As String, sectionAuto As String
            Dim typeRebar As Long
            Dim mappedType As String

            Set allFrames = GetAllElements("Frame")

            For Each fr In allFrames
                sectionName = "": sectionAuto = ""
                On Error Resume Next
                ret = SapModel.frameObj.GetSection(CStr(fr), sectionName, sectionAuto)
                On Error GoTo 0

                If ret <> 0 Then
                    mappedType = "Program Determined"
                    If StrComp(propValue, mappedType, vbTextCompare) = 0 Then result.Add CStr(fr)
                    GoTo NextFrame_DT
                End If

                If Trim$(sectionName) = "" Then
                    mappedType = "Program Determined"
                    If StrComp(propValue, mappedType, vbTextCompare) = 0 Then result.Add CStr(fr)
                    GoTo NextFrame_DT
                End If

                typeRebar = -1
                On Error Resume Next
                ret = SapModel.PropFrame.GetTypeRebar(sectionName, typeRebar)
                On Error GoTo 0

                If ret <> 0 And Trim$(sectionAuto) <> "" Then
                    typeRebar = -1
                    On Error Resume Next
                    ret = SapModel.PropFrame.GetTypeRebar(sectionAuto, typeRebar)
                    On Error GoTo 0
                End If

                If ret = 0 Then
                    Select Case typeRebar
                        Case 1: mappedType = "Column"
                        Case 2: mappedType = "Beam"
                        Case Else: mappedType = "Program Determined"
                    End Select
                Else
                    mappedType = "Program Determined"
                End If

                If StrComp(propValue, mappedType, vbTextCompare) = 0 Or (StrComp(propValue, "Brace", vbTextCompare) = 0 And StrComp(mappedType, "Beam", vbTextCompare) = 0) Then
                    result.Add CStr(fr)
                End If

NextFrame_DT:
                sectionName = "": sectionAuto = ""
            Next fr

        Case "AutoSection"
            Dim allFrs As Collection
            Dim fItem As Variant
            Set allFrs = GetAllElements("Frame")
            For Each fItem In allFrs
                Dim secName As String, secAuto As String
                On Error Resume Next
                ret = SapModel.frameObj.GetSection(CStr(fItem), secName, secAuto)
                On Error GoTo 0
                If ret = 0 Then
                    If Trim$(secName) = propValue Then
                        result.Add CStr(fItem)
                    ElseIf Trim$(secAuto) = propValue Then
                        result.Add CStr(fItem)
                    End If
                End If
            Next fItem

        Case "Area-Type"
            Dim allAreas As Collection
            Dim a   As Variant
            Set allAreas = GetAllElements("Area")
            For Each a In allAreas
                result.Add CStr(a)
            Next a

        Case "Property"
            Dim allLinks As Collection
            Dim l   As Variant
            Set allLinks = GetAllElements("Link")
            For Each l In allLinks
                Dim propName As String
                On Error Resume Next
                ret = SapModel.LinkObj.GetProperty(CStr(l), propName)
                On Error GoTo 0
                If ret = 0 Then
                    If Trim$(propName) = propValue Then result.Add CStr(l)
                End If
            Next l

        Case "Coordinate-X"
            If EnsureSapModelAvailable() Then
                Set result = GetElementsByCoordinateAxis(objType, "X", propValue)
            End If

        Case "Coordinate-Y"
            If EnsureSapModelAvailable() Then
                Set result = GetElementsByCoordinateAxis(objType, "Y", propValue)
            End If

        Case "Coordinate-Z"
            If EnsureSapModelAvailable() Then
                Set result = GetElementsByCoordinateAxis(objType, "Z", propValue)
            End If

        Case Else
            ' Unknown attribute: return empty collection
    End Select

    Set GetElementsByPropertyValue = result
End Function
Private Function GetElementsByProperty() As Collection
    Dim result      As New Collection
    If lstAttributes.ListIndex < 0 Then Set GetElementsByProperty = result: Exit Function
    Dim attrName    As String: attrName = lstAttributes.text
    Dim selProps    As Collection: Set selProps = GetSelectedPropertyValues()
    If selProps.count = 0 Then Set GetElementsByProperty = result: Exit Function
    Dim pv As Variant, unionColl As New Collection
    For Each pv In selProps
        Dim tmp     As Collection
        Set tmp = GetElementsByPropertyValue(attrName, CStr(pv), GetPrimaryObjectType())
        Dim it      As Variant
        For Each it In tmp
            If Not IsInCollection(unionColl, CStr(it)) Then unionColl.Add it
        Next it
    Next pv
    Set GetElementsByProperty = unionColl
End Function

Private Function GetElementsBySelectedPropertiesUnion() As Collection
    Dim attrName    As String: attrName = lstAttributes.text
    Dim selProps    As Collection: Set selProps = GetSelectedPropertyValues()
    Dim unionColl   As New Collection
    Dim pv          As Variant
    For Each pv In selProps
        Dim tmp     As Collection
        Set tmp = GetElementsByPropertyValue(attrName, CStr(pv), GetPrimaryObjectType())
        Dim it      As Variant
        For Each it In tmp
            If Not IsInCollection(unionColl, CStr(it)) Then unionColl.Add it
        Next it
    Next pv
    Set GetElementsBySelectedPropertiesUnion = unionColl
End Function

Private Function GetElementsBySelectedPropertiesIntersection() As Collection
    Dim selProps    As Collection: Set selProps = GetSelectedPropertyValues()
    Dim first       As Boolean: first = True
    Dim currentSet  As Collection
    Dim pv          As Variant
    For Each pv In selProps
        Dim tmp     As Collection
        Set tmp = GetElementsByPropertyValue(lstAttributes.text, CStr(pv), GetPrimaryObjectType())
        If first Then
            Set currentSet = New Collection
            Dim it  As Variant
            For Each it In tmp: currentSet.Add it: Next it
            first = False
        Else
            Dim newSet As New Collection, item As Variant
            For Each item In currentSet
                If IsInCollection(tmp, CStr(item)) Then newSet.Add item
            Next item
            Set currentSet = newSet
        End If
    Next pv
    If first Then Set GetElementsBySelectedPropertiesIntersection = New Collection Else Set GetElementsBySelectedPropertiesIntersection = currentSet
End Function

Private Function GetAllElements(ByVal objType As String) As Collection
    Dim result      As New Collection
    If Not EnsureSapModelAvailable() Then Set GetAllElements = result: Exit Function
    Dim ret As Long, cnt As Long, names() As String, i As Long
    Select Case objType
        Case "Node": ret = SapModel.pointObj.GetNameList(cnt, names)
        Case "Frame": ret = SapModel.frameObj.GetNameList(cnt, names)
        Case "Cable": ret = SapModel.CableObj.GetNameList(cnt, names)
        Case "Tendon": ret = SapModel.TendonObj.GetNameList(cnt, names)
        Case "Area": ret = SapModel.AreaObj.GetNameList(cnt, names)
        Case "Solid": ret = SapModel.SolidObj.GetNameList(cnt, names)
        Case "Link": ret = SapModel.LinkObj.GetNameList(cnt, names)
        Case Else: ret = -1
    End Select
    If ret = 0 And cnt > 0 Then
        For i = 0 To cnt - 1
            If Trim$(names(i)) <> "" Then result.Add names(i)
        Next i
    End If
    Set GetAllElements = result
End Function

Private Function GetSelectedElements(ByVal objType As String) As Collection
    Dim result      As New Collection
    Dim numItems As Long, objTypeArr() As Long, objNameArr() As String, ret As Long, i As Long
    If SapModel Is Nothing Then Set GetSelectedElements = result: Exit Function
    ret = SapModel.SelectObj.GetSelected(numItems, objTypeArr, objNameArr)
    If ret = 0 And numItems > 0 Then
        For i = 0 To numItems - 1
            If objTypeArr(i) = GetObjectTypeCode(objType) Then
                If Trim$(objNameArr(i)) <> "" Then result.Add objNameArr(i)
            End If
        Next i
    End If
    Set GetSelectedElements = result
End Function

Private Function GetObjectTypeCode(ByVal objType As String) As Long
    Select Case objType
        Case "Node": GetObjectTypeCode = 1
        Case "Frame": GetObjectTypeCode = 2
        Case "Cable": GetObjectTypeCode = 3
        Case "Tendon": GetObjectTypeCode = 4
        Case "Area": GetObjectTypeCode = 5
        Case "Solid": GetObjectTypeCode = 6
        Case "Link": GetObjectTypeCode = 7
        Case Else: GetObjectTypeCode = 0
    End Select
End Function

' Convert object type code to name string
Private Function ObjectTypeCodeToName(ByVal Code As Long) As String
    Select Case Code
        Case 1: ObjectTypeCodeToName = "Node"
        Case 2: ObjectTypeCodeToName = "Frame"
        Case 3: ObjectTypeCodeToName = "Cable"
        Case 4: ObjectTypeCodeToName = "Tendon"
        Case 5: ObjectTypeCodeToName = "Area"
        Case 6: ObjectTypeCodeToName = "Solid"
        Case 7: ObjectTypeCodeToName = "Link"
        Case Else: ObjectTypeCodeToName = ""
    End Select
End Function

' Save current selection arrays (types and names) into variants to restore later
Private Sub SaveSelectionState(ByRef outCount As Long, ByRef outTypes() As Long, ByRef outNames() As String)
    ' Populate outCount, outTypes(), outNames() using SapModel.SelectObj.GetSelected
    outCount = 0
    On Error Resume Next
    If SapModel Is Nothing Then Exit Sub
    Dim ret         As Long
    ret = SapModel.SelectObj.GetSelected(outCount, outTypes, outNames)
    If ret <> 0 Or outCount <= 0 Then
        ' No selection retrieved - clear arrays
        outCount = 0
        On Error Resume Next
        Erase outTypes
        Erase outNames
        On Error GoTo 0
    End If
    On Error GoTo 0
End Sub

' Restore previously saved selection
Private Sub RestoreSelectionState(ByVal inCount As Long, ByRef inTypes() As Long, ByRef inNames() As String)
    On Error Resume Next
    If SapModel Is Nothing Then Exit Sub
    ' Clear current selection and re-select saved ones
    SapModel.SelectObj.ClearSelection
    Dim i           As Long
    For i = 0 To inCount - 1
        Dim tcode   As Long
        Dim nm      As String
        tcode = inTypes(i)
        nm = inNames(i)
        Dim tName   As String
        tName = ObjectTypeCodeToName(tcode)
        If tName <> "" Then
            Call SelectSingleObject(tName, CStr(nm), True)
        End If
    Next i
    On Error GoTo 0
End Sub

Private Function GetElementsByGroup(ByVal groupName As String, ByVal objType As String) As Collection
    Dim result As New Collection, ret As Long
    ret = SapModel.SelectObj.ClearSelection()
    ret = SapModel.SelectObj.Group(groupName, False)
    Set result = GetSelectedElements(objType)
    Set GetElementsByGroup = result
End Function

Private Function GetElementsBySection(ByVal sectionName As String, ByVal objType As String) As Collection
    Set GetElementsBySection = GetElementsByPropertyValue("Section", sectionName, objType)
End Function

Private Function GetElementsByMaterial(ByVal materialName As String, ByVal objType As String) As Collection
    Set GetElementsByMaterial = GetElementsByPropertyValue("Material", materialName, objType)
End Function

Private Function GetElementsByConstraint(ByVal constraintName As String) As Collection
    Set GetElementsByConstraint = GetElementsByPropertyValue("Constraint", constraintName, "Node")
End Function

' =========== Map support subtype into elements or nodes ===========
' If objType = "Node" -> return nodes
' Else -> return elements that either contain matching nodes OR have springs assigned directly
Private Function GetElementsBySupport(ByVal supportValue As String, Optional ByVal objType As String = "Node") As Collection
    Dim result      As New Collection
    If Not EnsureSapModelAvailable() Then
        Set GetElementsBySupport = result
        Exit Function
    End If

    ' supportValue expected to be like "Support - Fixed" or "Spring - Joint" or "Fixed" etc.
    Dim subtype     As String
    subtype = Trim$(supportValue)

    ' If user selected "Any" -> return all elements of type
    If LCase(subtype) = "any" Then
        Set result = GetAllElements(objType)
        Set GetElementsBySupport = result
        Exit Function
    End If

    ' If objType is Node -> simply return nodes matching subtype (checks above)
    If UCase(objType) = "NODE" Then
        Set result = GetNodesMatchingSupport(subtype)
        Set GetElementsBySupport = result
        Exit Function
    End If

    ' For non-node types: get nodes matching subtype, then map to elements.
    Dim matchingNodes As Collection
    Set matchingNodes = GetNodesMatchingSupport(subtype)

    Dim elems       As Collection
    Set elems = GetAllElements(objType)

    Dim addedElems  As Object
    Set addedElems = CreateObject("Scripting.Dictionary")

    Dim e           As Variant
    Dim n           As Variant

    For Each e In elems
        Dim elemNodes As Collection
        Set elemNodes = GetNodesFromElement(CStr(e), objType)
        Dim foundMatch As Boolean: foundMatch = False
        For Each n In elemNodes
            If IsInCollection(matchingNodes, CStr(n)) Then
                foundMatch = True: Exit For
            End If
        Next n
        If foundMatch Then
            If Not addedElems.exists(CStr(e)) Then
                addedElems.Add CStr(e), True
                result.Add CStr(e)
            End If
        End If
    Next e

    ' Also include elements that have springs directly assigned (if subtype is spring-related or "Spring - Any")
    If InStr(LCase(subtype), "spring") > 0 Then
        ' check frames/areas/solids/links depending on objType
        If UCase(objType) = "FRAME" Or UCase(objType) = "CABLE" Or UCase(objType) = "TENDON" Then
            For Each e In elems
                Dim NumberSpringsF As Long
                Dim MyTypeF() As Long
                Dim sF() As Double
                Dim SimpleSpringTypeF() As Long
                Dim LinkPropF() As String
                Dim SpringLocalOneTypeF() As Long
                Dim DirF() As Long
                Dim Plane23AngleF() As Double
                Dim VecXF() As Double, VecYF() As Double, VecZF() As Double
                Dim CSysF() As String, AngF() As Double
                On Error Resume Next
                ret = SapModel.frameObj.GetSpring(CStr(e), NumberSpringsF, MyTypeF, sF, SimpleSpringTypeF, LinkPropF, SpringLocalOneTypeF, DirF, Plane23AngleF, VecXF, VecYF, VecZF, CSysF, AngF)
                On Error GoTo 0
                If ret = 0 And NumberSpringsF > 0 Then
                    If Not addedElems.exists(CStr(e)) Then
                        addedElems.Add CStr(e), True
                        result.Add CStr(e)
                    End If
                End If
            Next e
        End If

        If UCase(objType) = "AREA" Then
            For Each e In elems
                Dim NumberSpringsA As Long
                Dim MyTypeA() As Long
                Dim sA() As Double
                Dim SimpleSpringTypeA() As Long
                Dim LinkPropA() As String
                Dim FaceA() As Long
                Dim SpringLocalOneTypeA() As Long
                Dim DirA() As Long
                Dim OutwardA() As Boolean
                Dim VecXA() As Double, VecYA() As Double, VecZA() As Double
                Dim CSysA() As String, AngA() As Double
                On Error Resume Next
                ret = SapModel.AreaObj.GetSpring(CStr(e), NumberSpringsA, MyTypeA, sA, SimpleSpringTypeA, LinkPropA, FaceA, SpringLocalOneTypeA, DirA, OutwardA, VecXA, VecYA, VecZA, CSysA, AngA)
                On Error GoTo 0
                If ret = 0 And NumberSpringsA > 0 Then
                    If Not addedElems.exists(CStr(e)) Then
                        addedElems.Add CStr(e), True
                        result.Add CStr(e)
                    End If
                End If
            Next e
        End If

        If UCase(objType) = "SOLID" Then
            For Each e In elems
                Dim NumberSpringsS As Long
                Dim MyTypeS() As Long
                Dim ss() As Double
                Dim SimpleSpringTypeS() As Long
                Dim LinkPropS() As String
                Dim FaceS() As Long
                Dim SpringLocalOneTypeS() As Long
                Dim DirS() As Long
                Dim OutwardS() As Boolean
                Dim VecXS() As Double, VecYS() As Double, VecZS() As Double
                Dim CSysS() As String, AngS() As Double
                On Error Resume Next
                ret = SapModel.SolidObj.GetSpring(CStr(e), NumberSpringsS, MyTypeS, ss, SimpleSpringTypeS, LinkPropS, FaceS, SpringLocalOneTypeS, DirS, OutwardS, VecXS, VecYS, VecZS, CSysS, AngS)
                On Error GoTo 0
                If ret = 0 And NumberSpringsS > 0 Then
                    If Not addedElems.exists(CStr(e)) Then
                        addedElems.Add CStr(e), True
                        result.Add CStr(e)
                    End If
                End If
            Next e
        End If

        If UCase(objType) = "LINK" Then
            For Each e In elems
                Dim lpProp As String
                On Error Resume Next
                ret = SapModel.LinkObj.GetProperty(CStr(e), lpProp)
                On Error GoTo 0
                If ret = 0 And Trim$(lpProp) <> "" Then
                    If Not addedElems.exists(CStr(e)) Then
                        addedElems.Add CStr(e), True
                        result.Add CStr(e)
                    End If
                End If
            Next e
        End If
    End If

    Set GetElementsBySupport = result
End Function


' =========== Node matching using SupportedPoints (fast) + strict DOF verification ===========
' Determine nodes matching a support/spring subtype. Uses SupportedPoints for speed but
' then verifies exact DOF pattern for Fixed/Pinned/Roller and checks springs via GetSpring.
Private Function GetNodesMatchingSupport(ByVal supportSubtype As String) As Collection
    Dim nodesResult As New Collection
    If Not EnsureSapModelAvailable() Then
        Set GetNodesMatchingSupport = nodesResult
        Exit Function
    End If

    Dim s           As String
    s = LCase(Trim$(supportSubtype))

    ' Build DOF mask and flags
    Dim reqDOF(5)   As Boolean
    Dim checkExact  As Boolean
    Dim includeRestraints As Boolean: includeRestraints = False
    Dim includeSprings As Boolean: includeSprings = False
    Dim i           As Long
    For i = 0 To 5: reqDOF(i) = False: Next i
    checkExact = False

    Select Case True
        Case InStr(s, "support - fixed") > 0 Or s = "fixed"
            ' all DOFs restrained
            For i = 0 To 5: reqDOF(i) = True: Next i
            includeRestraints = True
            checkExact = True
        Case InStr(s, "support - pinned") > 0 Or s = "pinned"
            ' translations restrained, rotations free => strict
            reqDOF(0) = True: reqDOF(1) = True: reqDOF(2) = True
            includeRestraints = True
            checkExact = True
        Case InStr(s, "support - roller") > 0 Or s = "roller"
            ' roller interpretation: restrain U2 only (strict)
            reqDOF(1) = True
            includeRestraints = True
            checkExact = True
        Case InStr(s, "spring -") > 0 Or InStr(s, "spring") > 0
            ' any spring related option -> consider springs
            includeSprings = True
        Case Else
            ' allow direct subtype names without prefix
            If s = "fixed" Or s = "pinned" Or s = "roller" Then
                ' handled above
            Else
                ' If user provided custom list like "u1,u2" or numeric indices
                If InStr(s, ",") > 0 Or InStr(s, "u") > 0 Or InStr(s, "r") > 0 Or IsNumeric(s) Then
                    Dim parts() As String, p As Variant
                    parts = Split(s, ",")
                    For Each p In parts
                        Dim token As String
                        token = Trim$(p)
                        If token = "" Then GoTo NextTok
                        If Left$(token, 1) = "u" Or Left$(token, 1) = "r" Then
                            Dim idxVal As Long
                            idxVal = val(mid$(token, 2))
                            If idxVal >= 1 And idxVal <= 3 Then
                                If Left$(token, 1) = "u" Then
                                    reqDOF(idxVal - 1) = True
                                Else
                                    reqDOF(2 + idxVal) = True
                                End If
                                includeRestraints = True
                            End If
                        ElseIf IsNumeric(token) Then
                            Dim n As Long
                            n = CLng(token)
                            If n >= 0 And n <= 5 Then
                                reqDOF(n) = True
                                includeRestraints = True
                            End If
                        End If
NextTok:
                    Next p
                Else
                    ' unknown -> return empty
                    Set GetNodesMatchingSupport = nodesResult
                    Exit Function
                End If
            End If
    End Select

    ' Build DOF array for SupportedPoints call
    Dim DOF()       As Boolean
    ReDim DOF(5)
    For i = 0 To 5: DOF(i) = reqDOF(i): Next i

    ' dictionary for uniqueness
    Dim added       As Object
    Set added = CreateObject("Scripting.Dictionary")

    On Error GoTo FinishNodes

    ' Use SupportedPoints to quickly get candidates (wrap save/restore so UI not affected)
    If includeRestraints Then
        Dim savedCnt As Long
        Dim savedTypes() As Long
        Dim savedNames() As String
        savedCnt = 0
        SaveSelectionState savedCnt, savedTypes, savedNames

        On Error Resume Next
        SapModel.SelectObj.ClearSelection
        SapModel.SelectObj.SupportedPoints DOF, "Local", False, True, False, False, False, False, False
        On Error GoTo 0

        Dim candidateRes As Collection
        Set candidateRes = GetSelectedElements("Node")

        ' restore selection immediately
        RestoreSelectionState savedCnt, savedTypes, savedNames

        ' verify strict DOF pattern if requested, otherwise accept candidates
        Dim nod     As Variant
        For Each nod In candidateRes
            Dim restraints() As Boolean
            Dim ret As Long
            On Error Resume Next
            ret = SapModel.pointObj.GetRestraint(CStr(nod), restraints)
            On Error GoTo 0
            If ret = 0 And IsArray(restraints) Then
                Dim lb As Long, ub As Long
                lb = LBound(restraints): ub = UBound(restraints)
                Dim ok As Boolean: ok = True
                For i = 0 To 5
                    If reqDOF(i) Then
                        ' required DOF must be True
                        If i < lb Or i > ub Then ok = False: Exit For
                        If restraints(i) = False Then ok = False: Exit For
                    ElseIf checkExact Then
                        ' for exact match we require other DOFs to be False
                        If i >= lb And i <= ub Then
                            If restraints(i) = True Then ok = False: Exit For
                        End If
                    End If
                Next i
                If ok Then
                    If Not added.exists(CStr(nod)) Then
                        added.Add CStr(nod), True
                        nodesResult.Add CStr(nod)
                    End If
                End If
            End If
        Next nod
    End If

    ' Springs: use SupportedPoints with spring flags -> then verify actual spring presence via GetSpring/get element springs
    If includeSprings Then
        Dim sSavedCnt As Long
        Dim sSavedTypes() As Long
        Dim sSavedNames() As String
        sSavedCnt = 0
        SaveSelectionState sSavedCnt, sSavedTypes, sSavedNames

        On Error Resume Next
        SapModel.SelectObj.ClearSelection
        SapModel.SelectObj.SupportedPoints DOF, "Local", False, False, True, True, True, True, True
        On Error GoTo 0

        Dim candidateSpr As Collection
        Set candidateSpr = GetSelectedElements("Node")

        RestoreSelectionState sSavedCnt, sSavedTypes, sSavedNames

        ' Verify candidate spring nodes using pointObj.GetSpring (direct)
        Dim nk      As Variant
        For Each nk In candidateSpr
            If Not added.exists(CStr(nk)) Then
                Dim k() As Double
                Dim ret2 As Long
                On Error Resume Next
                ReDim k(5)
                ret2 = SapModel.pointObj.GetSpring(CStr(nk), k)
                On Error GoTo 0
                Dim found As Boolean: found = False
                If ret2 = 0 And IsArray(k) Then
                    For i = 0 To 5
                        If k(i) <> 0# Then
                            found = True: Exit For
                        End If
                    Next i
                End If
                If found Then
                    added.Add CStr(nk), True
                    nodesResult.Add CStr(nk)
                Else
                    ' fallback: node might receive spring via element assignment - we will check elements only when needed in GetElementsBySupport
                    ' here we don't add; element-level checks done when mapping to elements
                End If
            End If
        Next nk
    End If

FinishNodes:
    On Error GoTo 0
    Set GetNodesMatchingSupport = nodesResult
End Function
' ---------- New helper: build lstPropertyValues for Support when objType = "Node" ----------
' Scans all nodes, reads GetRestraint for each node, classifies into Fixed/Pinned/Roller/Other,
' then populates lstPropertyValues with only the support types actually present.
Private Sub PopulateSupportPropertyValuesForNodes()
    lstPropertyValues.Clear
    lstPropertyValues.AddItem "Any"

    If Not EnsureSapModelAvailable() Then Exit Sub

    Dim allNodes    As Collection
    Set allNodes = GetAllElements("Node")

    Dim foundFixed As Boolean, foundPinned As Boolean, foundRoller As Boolean, foundOther As Boolean
    foundFixed = False: foundPinned = False: foundRoller = False: foundOther = False

    Dim n           As Variant
    Dim ret         As Long
    For Each n In allNodes
        Dim restraints() As Boolean
        On Error Resume Next
        ret = SapModel.pointObj.GetRestraint(CStr(n), restraints)
        On Error GoTo 0
        If ret = 0 And IsArray(restraints) Then
            ' Normalize to indices 0..5 into b(0..5)
            Dim b(5) As Boolean
            Dim lb As Long, ub As Long
            lb = LBound(restraints): ub = UBound(restraints)
            Dim i   As Long
            For i = 0 To 5
                If i >= lb And i <= ub Then
                    b(i) = CBool(restraints(i))
                Else
                    b(i) = False
                End If
            Next i

            ' Classification rules:
            ' Fixed: all 6 DOFs true
            If b(0) And b(1) And b(2) And b(3) And b(4) And b(5) Then
                foundFixed = True
                GoTo NextNode
            End If

            ' Pinned: U1,U2,U3 true AND R1,R2,R3 false
            If (b(0) And b(1) And b(2)) And Not (b(3) Or b(4) Or b(5)) Then
                foundPinned = True
                GoTo NextNode
            End If

            ' Roller: exactly one translational DOF true AND all other 5 DOFs false
            Dim transCount As Long
            transCount = 0
            For i = 0 To 2
                If b(i) Then transCount = transCount + 1
            Next i
            If transCount = 1 Then
                ' ensure all rotations false and the other two trans false (transCount=1 ensures other trans false)
                If Not (b(3) Or b(4) Or b(5)) Then
                    foundRoller = True
                    GoTo NextNode
                End If
            End If

            ' Otherwise if any restraint exists mark as Other
            If transCount > 0 Or b(3) Or b(4) Or b(5) Then
                foundOther = True
            End If
        End If
NextNode:
    Next n

    ' Add entries in order, only if present
    If foundFixed Then lstPropertyValues.AddItem "Fixed"
    If foundPinned Then lstPropertyValues.AddItem "Pinned"
    If foundRoller Then lstPropertyValues.AddItem "Roller"
    If foundOther Then lstPropertyValues.AddItem "Other Restraint"

    ' After populating, ensure master list and apply sort+filter
    StoreMasterPropertyValues
    ApplyFilterAndSortPropertyValues
End Sub

' ============ NEW FUNCTION: Get nodes from element ============
Private Function GetNodesFromElement(ByVal elemName As String, ByVal objType As String) As Collection
    ' Returns collection of node names that belong to the given element
    Dim result      As New Collection
    Dim ret         As Long

    On Error Resume Next

    Select Case objType
        Case "Frame"
            Dim pt1 As String, pt2 As String
            ret = SapModel.frameObj.GetPoints(elemName, pt1, pt2)
            If ret = 0 Then
                If Trim$(pt1) <> "" Then result.Add pt1
                If Trim$(pt2) <> "" And pt2 <> pt1 Then result.Add pt2
            End If

        Case "Cable"
            Dim cpt1 As String, cpt2 As String
            ret = SapModel.CableObj.GetPoints(elemName, cpt1, cpt2)
            If ret = 0 Then
                If Trim$(cpt1) <> "" Then result.Add cpt1
                If Trim$(cpt2) <> "" And cpt2 <> cpt1 Then result.Add cpt2
            End If

        Case "Tendon"
            Dim tpt1 As String, tpt2 As String
            ret = SapModel.TendonObj.GetPoints(elemName, tpt1, tpt2)
            If ret = 0 Then
                If Trim$(tpt1) <> "" Then result.Add tpt1
                If Trim$(tpt2) <> "" And tpt2 <> tpt1 Then result.Add tpt2
            End If

        Case "Area"
            Dim numPts As Long
            Dim ptNames() As String
            ret = SapModel.AreaObj.GetPoints(elemName, numPts, ptNames)
            If ret = 0 And numPts > 0 Then
                Dim i As Long
                For i = 0 To numPts - 1
                    If Trim$(ptNames(i)) <> "" Then
                        If Not IsInCollection(result, ptNames(i)) Then
                            result.Add ptNames(i)
                        End If
                    End If
                Next i
            End If

        Case "Solid"
            Dim solidPts() As String
            On Error Resume Next
            ret = SapModel.SolidObj.GetPoints(elemName, solidPts)
            On Error GoTo 0
            If ret = 0 Then
                If IsArray(solidPts) Then
                    Dim j As Long
                    For j = LBound(solidPts) To UBound(solidPts)
                        If Trim$(CStr(solidPts(j))) <> "" Then
                            If Not IsInCollection(result, CStr(solidPts(j))) Then
                                result.Add CStr(solidPts(j))
                            End If
                        End If
                    Next j
                End If
            End If

        Case "Link"
            Dim lpt1 As String, lpt2 As String
            ret = SapModel.LinkObj.GetPoints(elemName, lpt1, lpt2)
            If ret = 0 Then
                If Trim$(lpt1) <> "" Then result.Add lpt1
                If Trim$(lpt2) <> "" And lpt2 <> lpt1 Then result.Add lpt2
            End If
    End Select

    On Error GoTo 0
    Set GetNodesFromElement = result
End Function

' ============ QUICK ACTIONS ============
Private Sub btnAll_Click()
    If Not EnsureSapModelAvailable() Then Exit Sub
    SapModel.SelectObj.All False
    btnGetFromSAP_Click
End Sub

Private Sub btnNone_Click()
    If Not EnsureSapModelAvailable() Then Exit Sub
    SapModel.SelectObj.ClearSelection
End Sub

Private Sub btnInversion_Click()
    If Not EnsureSapModelAvailable() Then Exit Sub
    SapModel.SelectObj.InvertSelection
    btnGetFromSAP_Click
End Sub

Private Sub btnPrevious_Click()
    If Not EnsureSapModelAvailable() Then Exit Sub
    SapModel.SelectObj.PreviousSelection
    btnGetFromSAP_Click
End Sub

' ============ APPLY TO SAP2000 ============
Private Function SelectSingleObject(ByVal objType As String, ByVal objName As String, ByVal selectFlag As Boolean) As Long
    Dim ret         As Long
    If Not EnsureSapModelAvailable() Then SelectSingleObject = -1: Exit Function
    On Error Resume Next
    Select Case objType
        Case "Node": ret = SapModel.pointObj.SetSelected(objName, selectFlag)
        Case "Frame": ret = SapModel.frameObj.SetSelected(objName, selectFlag)
        Case "Cable": ret = SapModel.CableObj.SetSelected(objName, selectFlag)
        Case "Tendon": ret = SapModel.TendonObj.SetSelected(objName, selectFlag)
        Case "Area": ret = SapModel.AreaObj.SetSelected(objName, selectFlag)
        Case "Solid": ret = SapModel.SolidObj.SetSelected(objName, selectFlag)
        Case "Link": ret = SapModel.LinkObj.SetSelected(objName, selectFlag)
        Case Else: ret = -1
    End Select
    On Error GoTo 0
    SelectSingleObject = ret
End Function

Private Sub btnApplyToSAP_Click()
    ' Apply selection to SAP2000 with new logic:
    ' - If chkSelectByCoordinate is checked and attribute is Coordinate-X/Y/Z:
    '     -> For each selected property value, call SelectPlaneForCoordinateValue(axis, displayVal, True)
    '     -> Refresh view and exit (do not continue with normal btnApplyToSAP element-by-element flow)
    ' - Otherwise: keep original behavior (select items from txtInputList)
    If Not EnsureSapModelAvailable() Then MsgBox "SAP model not available", vbExclamation: Exit Sub

    On Error GoTo ErrHandler

    ' If coordinate-mode is active, perform selection by plane for each selected property value
    Dim attr As String
    attr = ""
    If lstAttributes.ListIndex >= 0 Then attr = lstAttributes.text

    Dim cb As MSForms.CheckBox
    On Error Resume Next
    Set cb = Me.Controls("chkSelectByCoordinate")
    On Error GoTo ErrHandler

    If Not cb Is Nothing Then
        If cb.Value = True Then
            ' Only handle when attribute is a Coordinate axis
            Dim axisChar As String
            axisChar = ""
            If InStr(1, attr, "Coordinate-X", vbTextCompare) > 0 Then axisChar = "X"
            If InStr(1, attr, "Coordinate-Y", vbTextCompare) > 0 Then axisChar = "Y"
            If InStr(1, attr, "Coordinate-Z", vbTextCompare) > 0 Then axisChar = "Z"

            If axisChar <> "" Then
                ' Get selected display values from lstPropertyValues
                Dim selVals As Collection
                Set selVals = GetSelectedPropertyValues()
                If selVals.count = 0 Then
                    MsgBox "Please select one or more coordinate values first.", vbExclamation
                    Exit Sub
                End If

                Dim val As Variant
                For Each val In selVals
                    ' Call SelectPlaneForCoordinateValue to select (doSelect = True)
                    ' This function will call the API to select the plane based on a representative point or coordinate range
                    Call SelectPlaneForCoordinateValue(axisChar, CStr(val), True)
                Next val

                ' Refresh SAP view and return (do not proceed with normal apply behavior)
                On Error Resume Next
                SapModel.View.RefreshWindow
                On Error GoTo ErrHandler
                Exit Sub
            End If
        End If
    End If

    ' If not coordinate-mode, proceed with original btnApplyToSAP behavior:
    ' - If Node + NonNode: select non-node elements first, then select nodes of those elements
    ' - If only NonNode: select those elements (add to current selection)
    ' - If only Node: select nodes (add to current selection)
    ' - DO NOT clear selection first - always ADD to existing selection

    Dim objList As Collection
    Set objList = ParseObjectList(txtInputList.text)
    If objList.count = 0 Then Exit Sub

    Dim objName As Variant
    Dim ret As Long
    Dim successCount As Long
    successCount = 0

    ' Determine what's selected
    Dim nodeSelected As Boolean
    Dim nonNodeType As String
    nodeSelected = IsNodeSelected()
    nonNodeType = GetNonNodeType()

    ' Case 1: Both Node and NonNode selected - select elements first, then nodes belonging to them
    If nodeSelected And nonNodeType <> "" Then
        Dim selectedElements As New Collection
        ' First pass: select non-node elements
        For Each objName In objList
            ret = SelectSingleObject(nonNodeType, CStr(objName), True)
            selectedElements.Add CStr(objName)
            If ret = 0 Then successCount = successCount + 1
        Next objName

        ' Second pass: gather nodes from those elements and select them
        Dim nodesDict As Object
        Set nodesDict = CreateObject("Scripting.Dictionary")
        Dim elem As Variant
        For Each elem In selectedElements
            Dim nodesFromElement As Collection
            Set nodesFromElement = GetNodesFromElement(CStr(elem), nonNodeType)
            Dim n As Variant
            For Each n In nodesFromElement
                If Not nodesDict.exists(CStr(n)) Then
                    nodesDict.Add CStr(n), True
                End If
            Next n
        Next elem

        Dim nodeName As Variant
        For Each nodeName In nodesDict.keys
            ret = SelectSingleObject("Node", CStr(nodeName), True)
            If ret = 0 Then successCount = successCount + 1
        Next nodeName

    ElseIf nonNodeType <> "" Then
        ' Only non-node type selected, select those elements (add to existing selection)
        For Each objName In objList
            ret = SelectSingleObject(nonNodeType, CStr(objName), True)
            If ret = 0 Then successCount = successCount + 1
        Next objName

    ElseIf nodeSelected Then
        ' Only Node selected, select nodes (add to existing selection)
        For Each objName In objList
            ret = SelectSingleObject("Node", CStr(objName), True)
            If ret = 0 Then successCount = successCount + 1
        Next objName
    End If

    SapModel.View.RefreshWindow

    Exit Sub

ErrHandler:
    MsgBox "Error in Apply: " & err.number & " - " & err.description, vbExclamation
End Sub

Private Sub btnGetFromSAP_Click()
    If Not EnsureSapModelAvailable() Then txtInputList.text = "": UpdateInputListCount: Exit Sub
    Dim numItems As Long, objTypeArr() As Long, objNameArr() As String, ret As Long, i As Long
    Dim objType As String, targetTypeCode As Long, elements As New Collection
    objType = GetPrimaryObjectType()
    targetTypeCode = GetObjectTypeCode(objType)
    ret = SapModel.SelectObj.GetSelected(numItems, objTypeArr, objNameArr)
    If ret = 0 And numItems > 0 Then
        For i = 0 To numItems - 1
            If objTypeArr(i) = targetTypeCode Then
                If Trim$(objNameArr(i)) <> "" Then elements.Add objNameArr(i)
            End If
        Next i
        txtInputList.text = CleanText(CollectionToCompactText(elements))
    Else
        txtInputList.text = ""
    End If
    SaveCurrentInputList
    UpdateInputListCount
End Sub

' ============ TOOLS (sorting/unique/parse) ============
Private Sub btnSort_Click()
    Dim col         As Collection: Set col = ParseObjectList(txtInputList.text)
    Dim sorted      As Collection: Set sorted = SortCollection(col)
    txtInputList.text = CleanText(CollectionToCompactText(sorted))
    SaveCurrentInputList
    UpdateInputListCount
End Sub

Private Sub btnClear_Click()
    txtInputList.text = ""
    SaveCurrentInputList
    UpdateInputListCount
End Sub

Private Sub btnUnique_Click()
    ' Toggle display mode between compact (ranges) and expanded (explicit list) while removing duplicates.
    On Error Resume Next

    ' Parse input and build unique set
    Dim col As Collection
    Set col = ParseObjectList(txtInputList.text)

    ' Use dictionary to build unique list preserving first instance
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Dim item As Variant
    For Each item In col
        If Not dict.exists(CStr(item)) Then
            dict.Add CStr(item), True
        End If
    Next item

    ' Convert back to Collection
    Dim uniqCol As New Collection
    Dim key As Variant
    For Each key In dict.keys
        uniqCol.Add CStr(key)
    Next key

    ' Sort the unique collection to produce deterministic order
    Dim sortedCol As Collection
    Set sortedCol = SortCollection(uniqCol)

    ' Toggle display mode (flip on each click). Initial default will flip from False->True on first click,
    ' which preserves previous behavior (compact on first press).
    gUniqueDisplayCompact = Not gUniqueDisplayCompact

    Dim outText As String
    If gUniqueDisplayCompact Then
        ' compact (range) representation (existing helper)
        outText = CollectionToCompactText(sortedCol)
    Else
        ' expanded explicit list representation
        outText = CollectionToExpandedText(sortedCol)
    End If

    ' Write back, save and update UI
    txtInputList.text = CleanText(outText)
    SaveCurrentInputList
    UpdateInputListCount

    On Error GoTo 0
End Sub

' ============ HELPER FUNCTIONS (parse/compact/sort/is-in-collection) ============
Private Function ParseObjectList(ByVal inputText As String) As Collection
    Dim result As New Collection, temp As String, parts() As String, i As Long, part As String
    Dim rangeStart As Long, rangeEnd As Long, j As Long
    temp = Trim(inputText)
    If temp = "" Then Set ParseObjectList = result: Exit Function
    temp = Replace(temp, vbCrLf, ",")
    temp = Replace(temp, vbTab, ",")
    temp = Replace(temp, ";", ",")
    temp = Replace(temp, " ", ",")
    Do While InStr(temp, ",,") > 0: temp = Replace(temp, ",,", ","): Loop
    parts = Split(temp, ",")
    For i = LBound(parts) To UBound(parts)
        part = Trim(parts(i))
        If part <> "" Then
            If InStr(LCase(part), "to") > 0 Then
                rangeStart = val(Left(part, InStr(LCase(part), "to") - 1))
                rangeEnd = val(mid(part, InStr(LCase(part), "to") + 2))
                For j = rangeStart To rangeEnd: On Error Resume Next: result.Add CStr(j): On Error GoTo 0: Next j
            ElseIf InStr(part, "-") > 0 And IsNumeric(Left(part, 1)) Then
                Dim dashParts() As String: dashParts = Split(part, "-")
                If UBound(dashParts) = 1 And IsNumeric(dashParts(0)) And IsNumeric(dashParts(1)) Then
                    rangeStart = val(dashParts(0)): rangeEnd = val(dashParts(1))
                    For j = rangeStart To rangeEnd: On Error Resume Next: result.Add CStr(j): On Error GoTo 0: Next j
                Else
                    On Error Resume Next: result.Add part: On Error GoTo 0
                End If
            Else
                On Error Resume Next: result.Add part: On Error GoTo 0
            End If
        End If
    Next i
    Set ParseObjectList = result
End Function

Private Function CollectionToCompactText(ByVal col As Collection) As String
    If col.count = 0 Then
        CollectionToCompactText = ""
        Exit Function
    End If
    Dim sorted      As Collection: Set sorted = SortCollection(col)
    Dim result As String, numbers() As Long, i As Long, j As Long, isAllNumeric As Boolean, item As Variant
    Dim sep         As String: sep = vbCrLf
    isAllNumeric = True
    ReDim numbers(1 To sorted.count)
    i = 1
    For Each item In sorted
        If IsNumeric(item) Then numbers(i) = CLng(item): i = i + 1 Else isAllNumeric = False: Exit For
    Next item
    If Not isAllNumeric Then
        result = ""
        For Each item In sorted
            If result <> "" Then result = result & sep
            result = result & CStr(item)
        Next item
        CollectionToCompactText = result: Exit Function
    End If
    result = "": i = 1
    Do While i <= sorted.count
        Dim rangeStart As Long, rangeEnd As Long
        rangeStart = numbers(i): rangeEnd = rangeStart
        j = i + 1
        Do While j <= sorted.count
            If numbers(j) = rangeEnd + 1 Then rangeEnd = numbers(j): j = j + 1 Else Exit Do
        Loop
        If result <> "" Then result = result & sep
        If rangeEnd - rangeStart >= 2 Then
            result = result & rangeStart & "to" & rangeEnd
        ElseIf rangeEnd = rangeStart + 1 Then
            result = result & rangeStart & ", " & rangeEnd
        Else
            result = result & rangeStart
        End If
        i = j
    Loop
    CollectionToCompactText = result
End Function

Private Function SortCollection(ByVal col As Collection) As Collection
    Dim result As New Collection, arr() As String, i As Long, j As Long, temp As String, item As Variant
    If col.count = 0 Then Set SortCollection = result: Exit Function
    ReDim arr(1 To col.count)
    i = 1
    For Each item In col: arr(i) = CStr(item): i = i + 1: Next item
    For i = 1 To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If IsNumeric(arr(i)) And IsNumeric(arr(j)) Then
                If CLng(arr(i)) > CLng(arr(j)) Then temp = arr(i): arr(i) = arr(j): arr(j) = temp
            Else
                If StrComp(arr(i), arr(j), vbTextCompare) > 0 Then temp = arr(i): arr(i) = arr(j): arr(j) = temp
            End If
        Next j
    Next i
    For i = 1 To UBound(arr): result.Add arr(i): Next i
    Set SortCollection = result
End Function

Private Function IsInCollection(ByVal col As Collection, ByVal key As String) As Boolean
    Dim item        As Variant
    IsInCollection = False
    For Each item In col
        If CStr(item) = key Then IsInCollection = True: Exit Function
    Next item
End Function

' ============ Selected properties helpers ============
Private Function AnyPropertySelected() As Boolean
    Dim i           As Long
    AnyPropertySelected = False
    For i = 0 To lstPropertyValues.ListCount - 1
        If lstPropertyValues.Selected(i) Then AnyPropertySelected = True: Exit Function
    Next i
End Function

Private Function GetSelectedPropertyValues() As Collection
    Dim result As New Collection, i As Long
    For i = 0 To lstPropertyValues.ListCount - 1
        If lstPropertyValues.Selected(i) Then result.Add lstPropertyValues.List(i)
    Next i
    Set GetSelectedPropertyValues = result
End Function

' ============ TEXT CLEANUP & COUNT UI ============
Private Function CleanText(ByVal raw As String) As String
    If raw = "" Then CleanText = "": Exit Function
    Dim s           As String
    s = raw
    ' Normalize CR/LF combinations: convert CRLF and CR and LF to CRLF
    s = Replace(s, vbCrLf, vbLf)    ' unify to LF
    s = Replace(s, vbCr, vbLf)
    s = Replace(s, vbLf, vbCrLf)
    ' Remove control characters except tab(9), LF(10), CR(13)
    Dim i           As Long
    For i = 0 To 31
        If i <> 9 And i <> 10 And i <> 13 Then
            s = Replace(s, Chr$(i), "")
        End If
    Next i
    CleanText = Trim$(s)
End Function

Private Sub EnsureTotalCountLabel()
    On Error Resume Next
    Dim ctrl        As MSForms.Control
    Set ctrl = Nothing
    Set ctrl = Me.Controls("lblTotalCount")
    If ctrl Is Nothing Then
        Dim lbl     As MSForms.Label
        Set lbl = Me.Controls.Add("Forms.Label.1", "lblTotalCount", True)
        lbl.Left = txtInputList.Left
        lbl.Top = txtInputList.Top + txtInputList.height + 6
        lbl.Width = txtInputList.Width
        lbl.height = 18
        lbl.Caption = "Total: 0"
        lbl.Font.Size = 10
    End If
    On Error GoTo 0
End Sub

Private Sub UpdateInputListCount()
    Dim col         As Collection
    Set col = ParseObjectList(txtInputList.text)
    On Error Resume Next
    If Not Me.Controls("lblTotalCount") Is Nothing Then
        Me.Controls("lblTotalCount").Caption = "Total: " & CStr(col.count)
    End If
    On Error GoTo 0
End Sub

Private Sub txtInputList_Change()
    UpdateInputListCount
    On Error Resume Next
    ' If we currently have an active input-list filter (we stored original master), detect user edits
    If gInputListMasterText <> "" Then
        ' Mark that the user edited the visible (filtered) list and store the edited text for later merge
        gInputListFilteredEdited = True
        gInputListEditedText = txtInputList.text
    End If
    On Error GoTo 0
End Sub

' Canonical GetStoreKey: build key in fixed order to avoid ordering differences
Private Function GetStoreKey() As String
    Dim parts As Collection
    Set parts = New Collection

    ' Canonical order: Node first, then this fixed sequence
    On Error Resume Next
    If optNode.Value Then parts.Add "Node"
    If optFrame.Value Then parts.Add "Frame"
    If optCable.Value Then parts.Add "Cable"
    If optTendon.Value Then parts.Add "Tendon"
    If optArea.Value Then parts.Add "Area"
    If optSolid.Value Then parts.Add "Solid"
    If optLink.Value Then parts.Add "Link"
    On Error GoTo 0

    Dim s As String
    s = ""
    Dim it As Variant
    For Each it In parts
        If s <> "" Then s = s & ","
        s = s & CStr(it)
    Next it

    If s = "" Then
        ' fallback to primary type if nothing selected (shouldn't normally happen)
        s = GetPrimaryObjectType()
    End If

    GetStoreKey = s & "|Group=" & IIf(chkIncludeGroup.Value, "1", "0")
End Function
' EnsureDefaultInputKeys
' Create default empty entries for each primary object type for both Group flags (0 and 1).
Private Sub EnsureDefaultInputKeys()
    On Error Resume Next
    Dim store As Object
    Set store = gInputListStore
    If store Is Nothing Then
        Set store = CreateObject("Scripting.Dictionary")
        Set gInputListStore = store
    End If

    Dim typesArr As Variant
    typesArr = Array("Node", "Frame", "Cable", "Tendon", "Area", "Solid", "Link")

    Dim grp As Long
    For grp = 0 To 1
        Dim t As Variant
        For Each t In typesArr
            Dim key As String
            key = CStr(t) & "|Group=" & IIf(grp = 1, "1", "0")
            If Not store.exists(key) Then
                store.Add key, ""    ' initialize as empty string
            End If
        Next t
    Next grp
End Sub

' Replace SaveCurrentInputList with robust version that ensures store exists
Private Sub SaveCurrentInputList()
    Dim store As Object
    On Error Resume Next
    Set store = gInputListStore
    On Error GoTo 0
    If store Is Nothing Then
        Set store = CreateObject("Scripting.Dictionary")
        Set gInputListStore = store
    End If

    Dim key As String
    key = GetStoreKey()

    ' Save cleaned text for this key
    If store.exists(key) Then
        store(key) = CleanText(txtInputList.text)
    Else
        store.Add key, CleanText(txtInputList.text)
    End If
End Sub

' Robust LoadSavedInputList: load exactly current key; if missing -> blank
Private Sub LoadSavedInputList()
    Dim store As Object
    On Error Resume Next
    Set store = gInputListStore
    On Error GoTo 0

    Dim key As String
    key = GetStoreKey()

    If Not store Is Nothing Then
        If store.exists(key) Then
            txtInputList.text = CleanText(CStr(store(key)))
        Else
            ' Key not found -> initialize blank to avoid mixing names from other types
            txtInputList.text = ""
            ' ensure default slot exists for future saves
            store.Add key, ""
        End If
    Else
        ' No store at all -> create store & defaults, leave blank
        Set store = CreateObject("Scripting.Dictionary")
        Set gInputListStore = store
        EnsureDefaultInputKeys
        txtInputList.text = ""
    End If

    UpdateInputListCount
End Sub
' -----------------------------
' SaveCurrentInputListForKey: save txtInputList under a specific store key
' -----------------------------
Private Sub SaveCurrentInputListForKey(ByVal key As String)
    Dim store As Object
    On Error Resume Next
    Set store = gInputListStore
    On Error GoTo 0
    If store Is Nothing Then
        Set store = CreateObject("Scripting.Dictionary")
        Set gInputListStore = store
    End If

    If Trim$(CStr(key)) = "" Then
        key = GetStoreKey()
    End If

    If store.exists(key) Then
        store(key) = CleanText(txtInputList.text)
    Else
        store.Add key, CleanText(txtInputList.text)
    End If
End Sub

' ============ Missing helpers used for fallback (implementations) ============
Private Function GetFramesWithSection(ByVal sectionName As String) As Collection
    Dim result As New Collection, allFrames As Collection, f As Variant, ret As Long
    Dim propName As String, sAuto As String
    Set allFrames = GetAllElements("Frame")
    For Each f In allFrames
        ret = SapModel.frameObj.GetSection(CStr(f), propName, sAuto)
        If ret = 0 Then
            If Trim$(propName) = sectionName Then result.Add CStr(f)
        End If
    Next f
    Set GetFramesWithSection = result
End Function

Private Function GetAreasWithProperty(ByVal propNameSearch As String) As Collection
    Dim result As New Collection, allAreas As Collection, a As Variant, ret As Long, propName As String
    Set allAreas = GetAllElements("Area")
    For Each a In allAreas
        ret = SapModel.AreaObj.GetProperty(CStr(a), propName)
        If ret = 0 Then
            If Trim$(propName) = propNameSearch Then result.Add CStr(a)
        End If
    Next a
    Set GetAreasWithProperty = result
End Function

'' ============ Close handler =================
'Private Sub btnClose_Click()
'    SaveCurrentInputList
'    ' Restore Excel if needed (both minimize and shrink)
'    RestoreExcelWindowAfterForm
'    RestoreShrunkExcelWindow
'    ' RestoreFormParentToExcel Me
'    Unload Me
'End Sub


' Return nodes whose coordinate on given axis matches propValue (string numeric)
Private Function GetNodesByCoordinateAxis(ByVal axis As String, ByVal propValue As String) As Collection
    Dim result      As New Collection
    If Not EnsureSapModelAvailable() Then
        Set GetNodesByCoordinateAxis = result
        Exit Function
    End If

    Dim cnt         As Long
    Dim names()     As String
    Dim ret         As Long
    Dim i           As Long
    Dim X As Double, Y As Double, Z As Double
    Dim target      As Double

    On Error GoTo CleanFail
    ' strip appended grid names like "500 (1F)" -> "500"
    Dim numericStr  As String
    numericStr = NumericStringFromDisplay(propValue)
    target = CDbl(numericStr)
    ret = SapModel.pointObj.GetNameList(cnt, names)
    If ret <> 0 Or cnt <= 0 Then
        Set GetNodesByCoordinateAxis = result
        Exit Function
    End If

    For i = 0 To cnt - 1
        If Trim$(names(i)) <> "" Then
            ret = SapModel.pointObj.GetCoordCartesian(names(i), X, Y, Z)
            If ret = 0 Then
                Dim val As Double
                Select Case UCase(axis)
                    Case "X": val = X
                    Case "Y": val = Y
                    Case "Z": val = Z
                    Case Else: val = X
                End Select
                ' compare rounded to 6 decimals
                If Abs(Round(val, 6) - Round(target, 6)) < 0.0000005 Then
                    result.Add names(i)
                End If
            End If
        End If
    Next i

CleanFail:
    On Error GoTo 0
    Set GetNodesByCoordinateAxis = result
End Function

' Return elements of various types whose representative coordinate on given axis matches propValue
Private Function GetElementsByCoordinateAxis(ByVal objType As String, ByVal axis As String, ByVal propValue As String) As Collection
    Dim result      As New Collection
    If Not EnsureSapModelAvailable() Then
        Set GetElementsByCoordinateAxis = result
        Exit Function
    End If

    Dim target      As Double
    On Error GoTo CleanExit
    ' strip appended grid names like "500 (1F)" -> "500"
    Dim numericStr2 As String
    numericStr2 = NumericStringFromDisplay(propValue)
    target = CDbl(numericStr2)

    Dim elems       As Collection
    Set elems = GetAllElements(objType)
    Dim e           As Variant
    Dim ret         As Long
    Dim px As Double, py As Double, pz As Double

    For Each e In elems
        px = 0: py = 0: pz = 0
        Dim found   As Boolean: found = False

        If objType = "Node" Then
            Dim nx As Double, ny As Double, nz As Double
            On Error Resume Next
            ret = SapModel.pointObj.GetCoordCartesian(CStr(e), nx, ny, nz)
            On Error GoTo 0
            If ret = 0 Then
                px = nx: py = ny: pz = nz: found = True
            End If

        ElseIf objType = "Frame" Or objType = "Cable" Or objType = "Tendon" Then
            Dim sp As String, ep As String
            On Error Resume Next
            Select Case objType
                Case "Frame": ret = SapModel.frameObj.GetPoints(CStr(e), sp, ep)
                Case "Cable": ret = SapModel.CableObj.GetPoints(CStr(e), sp, ep)
                Case "Tendon": ret = SapModel.TendonObj.GetPoints(CStr(e), sp, ep)
                Case Else: ret = -1
            End Select
            On Error GoTo 0
            If ret = 0 And Trim$(sp) <> "" Then
                Dim sx As Double, sy As Double, sz As Double
                On Error Resume Next
                ret = SapModel.pointObj.GetCoordCartesian(sp, sx, sy, sz)
                On Error GoTo 0
                If ret = 0 Then
                    If Trim$(ep) <> "" Then
                        Dim ex As Double, ey As Double, ez As Double
                        On Error Resume Next
                        ret = SapModel.pointObj.GetCoordCartesian(ep, ex, ey, ez)
                        On Error GoTo 0
                        If ret = 0 Then
                            px = (sx + ex) / 2: py = (sy + ey) / 2: pz = (sz + ez) / 2
                            found = True
                        Else
                            px = sx: py = sy: pz = sz: found = True
                        End If
                    Else
                        px = sx: py = sy: pz = sz: found = True
                    End If
                End If
            End If

        ElseIf objType = "Area" Then
            Dim numPts As Long
            Dim ptNames() As String
            On Error Resume Next
            ret = SapModel.AreaObj.GetPoints(CStr(e), numPts, ptNames)
            On Error GoTo 0
            If ret = 0 And numPts > 0 Then
                Dim sumx As Double, sumy As Double, sumZ As Double, cntPts As Long, k As Long
                sumx = 0: sumy = 0: sumZ = 0: cntPts = 0
                For k = 0 To numPts - 1
                    If Trim$(ptNames(k)) <> "" Then
                        Dim tx As Double, ty As Double, tz As Double
                        On Error Resume Next
                        ret = SapModel.pointObj.GetCoordCartesian(ptNames(k), tx, ty, tz)
                        On Error GoTo 0
                        If ret = 0 Then
                            sumx = sumx + tx: sumy = sumy + ty: sumZ = sumZ + tz
                            cntPts = cntPts + 1
                        End If
                    End If
                Next k
                If cntPts > 0 Then
                    px = sumx / cntPts: py = sumy / cntPts: pz = sumZ / cntPts
                    found = True
                End If
            End If

        ElseIf objType = "Solid" Then
            ' Solid: handle returned points array as Variant
            Dim sPts() As String
            On Error Resume Next
            ret = SapModel.SolidObj.GetPoints(CStr(e), sPts)
            On Error GoTo 0
            If ret = 0 Then
                If IsArray(sPts) Then
                    Dim kk As Long, scnt As Long
                    scnt = 0
                    Dim sx2 As Double, sy2 As Double, sz2 As Double
                    sx2 = 0: sy2 = 0: sz2 = 0
                    For kk = LBound(sPts) To UBound(sPts)
                        If Trim$(CStr(sPts(kk))) <> "" Then
                            Dim tx2 As Double, ty2 As Double, tz2 As Double
                            On Error Resume Next
                            ret = SapModel.pointObj.GetCoordCartesian(sPts(kk), tx2, ty2, tz2)
                            On Error GoTo 0
                            If ret = 0 Then
                                sx2 = sx2 + tx2: sy2 = sy2 + ty2: sz2 = sz2 + tz2
                                scnt = scnt + 1
                            End If
                        End If
                    Next kk
                    If scnt > 0 Then
                        px = sx2 / scnt: py = sy2 / scnt: pz = sz2 / scnt
                        found = True
                    End If
                End If
            End If

        ElseIf objType = "Link" Then
            Dim lp1 As String, lp2 As String
            On Error Resume Next
            ret = SapModel.LinkObj.GetPoints(CStr(e), lp1, lp2)
            On Error GoTo 0
            If ret = 0 And Trim$(lp1) <> "" Then
                Dim lx As Double, ly As Double, lz As Double
                On Error Resume Next
                ret = SapModel.pointObj.GetCoordCartesian(lp1, lx, ly, lz)
                On Error GoTo 0
                If ret = 0 Then
                    If Trim$(lp2) <> "" Then
                        Dim lx2 As Double, ly2 As Double, lz2 As Double
                        On Error Resume Next
                        ret = SapModel.pointObj.GetCoordCartesian(lp2, lx2, ly2, lz2)
                        On Error GoTo 0
                        If ret = 0 Then
                            px = (lx + lx2) / 2: py = (ly + ly2) / 2: pz = (lz + lz2) / 2
                            found = True
                        Else
                            px = lx: py = ly: pz = lz: found = True
                        End If
                    Else
                        px = lx: py = ly: pz = lz: found = True
                    End If
                End If
            End If
        End If

        If found Then
            Dim val As Double
            Select Case UCase(axis)
                Case "X": val = px
                Case "Y": val = py
                Case "Z": val = pz
                Case Else: val = px
            End Select
            If Abs(Round(val, 6) - Round(target, 6)) < 0.0000005 Then
                result.Add CStr(e)
            End If
        End If
    Next e

CleanExit:
    On Error GoTo 0
    Set GetElementsByCoordinateAxis = result
End Function

' ============ New Toggle Shrink Button (UI) =================
' Note: You should add a CommandButton named "btnToggleShrinkExcel" to the userform.
' If you add it at design time set its Name property = btnToggleShrinkExcel and Caption = "ToggleExcelShrink"
' If created at runtime, make sure to name it exactly "btnToggleShrinkExcel".

Private Sub btnToggleShrinkExcel_Click()
    ' Toggle shrink (not minimize) of Excel window
    ToggleShrinkExcelWindow
    ' Update button caption to reflect state
    On Error Resume Next
    If gExcelShrunkByForm Then
        Me.Controls("btnToggleShrinkExcel").Caption = "RestoreExcelWindow"
    Else
        Me.Controls("btnToggleShrinkExcel").Caption = "ToggleExcelShrink"
    End If
    On Error GoTo 0
End Sub

' ============ Gridline sheet helpers & export ============
Public Sub ExportGridLinesToGirdlineSheet()
    ' Read gridlines from sheet "Girdline" and populate gGridlineMap
    Dim ws          As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Girdline")
    On Error GoTo 0
    If ws Is Nothing Then
        ' Sheet not found - clear map and exit
        Set gGridlineMap = Nothing
        Exit Sub
    End If

    Dim lastRow     As Long
    lastRow = ws.Cells(ws.rows.count, "B").End(xlUp).row
    If lastRow < 5 Then
        ' no data rows
        Set gGridlineMap = Nothing
        Exit Sub
    End If

    Dim map         As Object
    Set map = CreateObject("Scripting.Dictionary")

    Dim r           As Long
    For r = 5 To lastRow
        Dim axis    As String
        axis = Trim$(CStr(ws.Cells(r, "B").Value))
        If axis = "" Then GoTo NextRowGL
        axis = UCase(axis)
        Dim gid     As String
        gid = Trim$(CStr(ws.Cells(r, "C").Value))
        Dim coordVal As Variant
        coordVal = ws.Cells(r, "D").Value
        If IsNumeric(coordVal) Then
            Dim key As String
            key = Format$(Round(CDbl(coordVal), 6), "0.######")
            If Not map.exists(axis) Then
                Dim inner As Object
                Set inner = CreateObject("Scripting.Dictionary")
                map.Add axis, inner
            End If
            If gid = "" Then
                ' if name empty use GridID column C fallback - already gid
                gid = ""
            End If
            If gid <> "" Then
                map(axis)(key) = gid
            End If
        End If
NextRowGL:
    Next r

    Set gGridlineMap = map
End Sub

Private Sub EnsureGridlineMapInitialized()
    ' Initialize gridline map if not present
    If gGridlineMap Is Nothing Then
        On Error Resume Next
        ExportGridLinesToGirdlineSheet
        On Error GoTo 0
    End If
End Sub

Private Function NumericStringFromDisplay(ByVal displayValue As String) As String
    ' Extracts numeric substring from a display value like "500 (1F)" -> "500"
    Dim s           As String
    s = Trim$(displayValue)
    If s = "" Then NumericStringFromDisplay = "0": Exit Function
    Dim pos         As Long
    pos = InStr(s, "(")
    If pos > 0 Then
        s = Trim$(Left$(s, pos - 1))
    Else
        ' also handle space separated e.g. "500 1F" or "500 Label"
        pos = InStr(s, " ")
        If pos > 0 Then s = Trim$(Left$(s, pos - 1))
    End If
    NumericStringFromDisplay = s
End Function

Private Sub chkSelectByCoordinate_Click()
    ' Do NOT perform selection here.
    ' This checkbox now only toggles "coordinate selection mode".
    ' Actual selection/deselection will be performed when the user clicks btnApplyToSAP.
    On Error Resume Next
    ' Ensure checkbox remains enabled/disabled correctly (UI only)
    Dim cb As MSForms.CheckBox
    Set cb = Me.Controls("chkSelectByCoordinate")
    If cb Is Nothing Then Exit Sub
    ' No automatic selection/deselection here.
    ' Optionally update any UI hint if needed.
End Sub

' Helper to select/deselect plane given axis and displayed coordinate value
Private Sub SelectPlaneForCoordinateValue(ByVal axis As String, ByVal displayVal As String, ByVal doSelect As Boolean)
    On Error GoTo ErrHandler
    If Not EnsureSapModelAvailable() Then Exit Sub

    ' Try to find a representative point with that coordinate
    Dim nodes       As Collection
    Set nodes = GetNodesByCoordinateAxis(axis, displayVal)
    Dim ret         As Long

    If nodes.count > 0 Then
        Dim ptName  As String
        ptName = CStr(nodes(1))
        Dim deSelectFlag As Boolean
        deSelectFlag = Not doSelect    ' API expects DeSelect parameter: False to select, True to deselect
        Select Case UCase(axis)
            Case "X"
                ' PlaneYZ selects same YZ plane => constant X
                On Error Resume Next
                ret = SapModel.SelectObj.PlaneYZ(ptName, deSelectFlag)
                On Error GoTo 0
            Case "Y"
                ' PlaneXZ selects same XZ plane => constant Y
                On Error Resume Next
                ret = SapModel.SelectObj.PlaneXZ(ptName, deSelectFlag)
                On Error GoTo 0
            Case "Z"
                ' PlaneXY selects same XY plane => constant Z
                On Error Resume Next
                ret = SapModel.SelectObj.PlaneXY(ptName, deSelectFlag)
                On Error GoTo 0
        End Select
        ' optionally refresh view
        SapModel.View.RefreshWindow
    Else
        ' No point representative found: fallback to CoordinateRange with small epsilon
        Dim numericStr As String
        numericStr = NumericStringFromDisplay(displayVal)
        If Not IsNumeric(numericStr) Then Exit Sub
        Dim target  As Double
        target = CDbl(numericStr)
        Dim eps     As Double
        eps = 0.0001
        Dim DeSel   As Boolean
        DeSel = Not doSelect
        If UCase(axis) = "X" Then
            On Error Resume Next
            ret = SapModel.SelectObj.CoordinateRange(target - eps, target + eps, -1000000000#, 1000000000#, -1000000000#, 1000000000#, DeSel, "Global", True, True, True, True, True, True)
            On Error GoTo 0
        ElseIf UCase(axis) = "Y" Then
            On Error Resume Next
            ret = SapModel.SelectObj.CoordinateRange(-1000000000#, 1000000000#, target - eps, target + eps, -1000000000#, 1000000000#, DeSel, "Global", True, True, True, True, True, True)
            On Error GoTo 0
        ElseIf UCase(axis) = "Z" Then
            On Error Resume Next
            ret = SapModel.SelectObj.CoordinateRange(-1000000000#, 1000000000#, -1000000000#, 1000000000#, target - eps, target + eps, DeSel, "Global", True, True, True, True, True, True)
            On Error GoTo 0
        End If
        SapModel.View.RefreshWindow
    End If

    Exit Sub
ErrHandler:
Debug.Print "SelectPlaneForCoordinateValue error: " & err.number & " - " & err.description
    On Error GoTo 0
End Sub

Private Sub lstPropertyValues_Click()
    ' Previously this sub auto-selected plane when chkSelectByCoordinate was checked.
    ' Remove that behavior to require explicit btnApplyToSAP click.
    On Error GoTo ExitSilent
    If lstAttributes.ListIndex < 0 Then Exit Sub

    ' Keep existing behavior for attribute selection UI (no auto selection)
    ' Optionally you can update UI elements or enable/disable controls here.

ExitSilent:
    On Error GoTo 0
End Sub

' ============ Gridline sheet helpers & export (end) ============
' --- START: Overlay labels for OptionButton Click-only toggle ---
Private Sub CreateOptionOverlays()
    Dim names       As Variant
    names = Array("Node", "Frame", "Cable", "Tendon", "Area", "Solid", "Link")
    Dim i           As Long
    Dim optCtrl     As MSForms.Control
    Dim lblName     As String
    Dim lbl         As MSForms.Label

    On Error Resume Next
    ' Remove any existing overlays first (safe to call repeatedly)
    For i = LBound(names) To UBound(names)
        lblName = "lblOverlay" & CStr(names(i))
        If Not Me.Controls Is Nothing Then
            If Not (Me.Controls(lblName) Is Nothing) Then
                Me.Controls.Remove lblName
            End If
        End If
    Next i

    ' Create overlays
    For i = LBound(names) To UBound(names)
        Set optCtrl = Nothing
        Set optCtrl = Me.Controls("opt" & CStr(names(i)))
        If Not optCtrl Is Nothing Then
            lblName = "lblOverlay" & CStr(names(i))
            ' Add label overlay
            Set lbl = Me.Controls.Add("Forms.Label.1", lblName, True)
            With lbl
                .Caption = ""                  ' no visible text
                .Left = optCtrl.Left
                .Top = optCtrl.Top
                .Width = optCtrl.Width
                .height = optCtrl.height
                .BackStyle = 0                 ' transparent
                .TakeFocusOnClick = False
                .TabStop = False
                ' bring overlay to front so it receives clicks
                .ZOrder 0
            End With
        End If
    Next i
    On Error GoTo 0
End Sub

' Overlay click handlers - toggle underlying OptionButtons and preserve exclusivity logic
Private Sub lblOverlayNode_Click()
    On Error Resume Next
    SaveCurrentInputListForKey GetStoreKey()

    ' Toggle Node
    If optNode.Value = True Then
        optNode.Value = False
    Else
        optNode.Value = True
    End If
    On Error GoTo 0

    SwitchObjectType
End Sub
Private Sub lblOverlayFrame_Click()
    On Error Resume Next
    SaveCurrentInputListForKey GetStoreKey()

    If optFrame.Value = True Then
        optFrame.Value = False
    Else
        optFrame.Value = True
        optCable.Value = False: optTendon.Value = False
        optArea.Value = False: optSolid.Value = False: optLink.Value = False
    End If
    On Error GoTo 0
    SwitchObjectType
End Sub

Private Sub lblOverlayCable_Click()
    On Error Resume Next
    SaveCurrentInputListForKey GetStoreKey()

    If optCable.Value = True Then
        optCable.Value = False
    Else
        optFrame.Value = False
        optCable.Value = True
        optTendon.Value = False: optArea.Value = False
        optSolid.Value = False: optLink.Value = False
    End If
    On Error GoTo 0
    SwitchObjectType
End Sub

Private Sub lblOverlayTendon_Click()
    On Error Resume Next
    SaveCurrentInputListForKey GetStoreKey()

    If optTendon.Value = True Then
        optTendon.Value = False
    Else
        optFrame.Value = False: optCable.Value = False
        optTendon.Value = True
        optArea.Value = False: optSolid.Value = False: optLink.Value = False
    End If
    On Error GoTo 0
    SwitchObjectType
End Sub

Private Sub lblOverlayArea_Click()
    On Error Resume Next
    SaveCurrentInputListForKey GetStoreKey()

    If optArea.Value = True Then
        optArea.Value = False
    Else
        optFrame.Value = False: optCable.Value = False: optTendon.Value = False
        optArea.Value = True
        optSolid.Value = False: optLink.Value = False
    End If
    On Error GoTo 0
    SwitchObjectType
End Sub

Private Sub lblOverlaySolid_Click()
    On Error Resume Next
    SaveCurrentInputListForKey GetStoreKey()

    If optSolid.Value = True Then
        optSolid.Value = False
    Else
        optFrame.Value = False: optCable.Value = False: optTendon.Value = False
        optArea.Value = False
        optSolid.Value = True
        optLink.Value = False
    End If
    On Error GoTo 0
    SwitchObjectType
End Sub

Private Sub lblOverlayLink_Click()
    On Error Resume Next
    SaveCurrentInputListForKey GetStoreKey()

    If optLink.Value = True Then
        optLink.Value = False
    Else
        optFrame.Value = False: optCable.Value = False: optTendon.Value = False
        optArea.Value = False: optSolid.Value = False
        optLink.Value = True
    End If
    On Error GoTo 0
    SwitchObjectType
End Sub


' -----------------------------
' chkIncludeGroup_MouseDown: save current list BEFORE group checkbox changes
' -----------------------------
Private Sub chkIncludeGroup_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    On Error Resume Next
    SaveCurrentInputListForKey GetStoreKey()
    On Error GoTo 0
End Sub

' --- END: Overlay labels for OptionButton Click-only toggle ---

' -----------------------------
' Property Values: master store, filter & sort helpers
' -----------------------------
Private Sub StoreMasterPropertyValues()
    ' Store current items from lstPropertyValues into gPropertyValuesMaster
    On Error Resume Next
    Dim coll As New Collection
    Dim i As Long
    For i = 0 To lstPropertyValues.ListCount - 1
        coll.Add CStr(lstPropertyValues.List(i))
    Next i
    Set gPropertyValuesMaster = Nothing
    Set gPropertyValuesMaster = coll
    On Error GoTo 0
End Sub

Private Sub ApplyFilterAndSortPropertyValues()
    ' Apply case-insensitive "contains" filter from txtFilterPropertyValues and sort alphabetically (A->Z)
    ' Preserve "Any" at the top if present in the master list.
    ' Numeric-aware sort: if both items are numeric (after stripping appended labels like " (1F)"),
    ' sort by numeric value. Numeric items come before non-numeric items.
    On Error Resume Next
    If gPropertyValuesMaster Is Nothing Then
        ' nothing to apply
        Exit Sub
    End If

    Dim filterText As String
    filterText = ""
    On Error Resume Next
    If Not Me.Controls("txtFilterPropertyValues") Is Nothing Then
        filterText = Trim$(CStr(Me.Controls("txtFilterPropertyValues").text))
    End If
    On Error GoTo 0
    filterText = LCase$(filterText)

    Dim temp As New Collection
    Dim itm As Variant
    Dim hasAny As Boolean: hasAny = False

    For Each itm In gPropertyValuesMaster
        Dim s As String
        s = CStr(itm)
        If LCase$(s) = "any" Then
            hasAny = True
        Else
            If filterText = "" Or InStr(LCase$(s), filterText) > 0 Then
                temp.Add s
            End If
        End If
    Next itm

    ' Prepare arrays for numeric-aware sorting
    Dim n As Long
    n = temp.count
    Dim orig() As String
    Dim isNum() As Boolean
    Dim numVal() As Double
    Dim i As Long, j As Long

    If n > 0 Then
        ReDim orig(1 To n)
        ReDim isNum(1 To n)
        ReDim numVal(1 To n)

        For i = 1 To n
            orig(i) = CStr(temp(i))
            Dim numericStr As String
            numericStr = NumericStringFromDisplay(orig(i))
            If IsNumeric(numericStr) Then
                isNum(i) = True
                numVal(i) = CDbl(numericStr)
            Else
                isNum(i) = False
                numVal(i) = 0#
            End If
        Next i

        ' Simple stable-ish sort: compare and swap rows (orig/isNum/numVal)
        Dim swapNeeded As Boolean
        For i = 1 To n - 1
            For j = i + 1 To n
                swapNeeded = False
                If isNum(i) And isNum(j) Then
                    ' Both numeric -> compare numeric values
                    If numVal(i) > numVal(j) Then swapNeeded = True
                ElseIf isNum(i) And Not isNum(j) Then
                    ' numeric should come before non-numeric -> no swap
                    swapNeeded = False
                ElseIf Not isNum(i) And isNum(j) Then
                    ' non-numeric currently before numeric -> swap to put numeric first
                    swapNeeded = True
                Else
                    ' Both non-numeric -> case-insensitive string compare
                    If StrComp(orig(i), orig(j), vbTextCompare) > 0 Then swapNeeded = True
                End If

                If swapNeeded Then
                    Dim ts As String: ts = orig(i): orig(i) = orig(j): orig(j) = ts
                    Dim tb As Boolean: tb = isNum(i): isNum(i) = isNum(j): isNum(j) = tb
                    Dim tD As Double: tD = numVal(i): numVal(i) = numVal(j): numVal(j) = tD
                End If
            Next j
        Next i
    End If

    ' Re-populate lstPropertyValues
    lstPropertyValues.Clear
    If hasAny Then lstPropertyValues.AddItem "Any"
    If n > 0 Then
        For i = 1 To n
            lstPropertyValues.AddItem orig(i)
        Next i
    End If
End Sub
Private Sub txtFilterPropertyValues_Change()
    ' Called when user types into filter textbox: update visible list immediately
    On Error Resume Next
    ApplyFilterAndSortPropertyValues
    On Error GoTo 0
End Sub
' Ensure a filter TextBox for txtInputList exists (creates lblFilterInputList + txtFilterInputList)
Private Sub EnsureInputListFilterControl()
    On Error Resume Next
    Dim tb As MSForms.TextBox
    Set tb = Nothing
    Set tb = Me.Controls("txtFilterInputList")
    If tb Is Nothing Then
        Dim lblFilt As MSForms.Label
        Set lblFilt = Nothing
        Set lblFilt = Me.Controls("lblFilterInputList")
        If lblFilt Is Nothing Then
            Set lblFilt = Me.Controls.Add("Forms.Label.1", "lblFilterInputList", True)
            lblFilt.Caption = "Filter Input:"
            lblFilt.Font.Size = 9
            ' position label above txtInputList (left-aligned with txtInputList)
            lblFilt.Left = txtInputList.Left
            lblFilt.Top = txtInputList.Top - 20
            lblFilt.Width = 70
            lblFilt.height = 18
        End If

        Dim newTB As MSForms.TextBox
        Set newTB = Me.Controls.Add("Forms.TextBox.1", "txtFilterInputList", True)
        ' position to the right of label, align with txtInputList width
        newTB.Left = txtInputList.Left + 74
        newTB.Top = txtInputList.Top - 22
        newTB.Width = txtInputList.Width - 74
        newTB.height = 18
        newTB.text = ""
        newTB.Font.Size = 9
    Else
        tb.text = ""
    End If
    On Error GoTo 0
End Sub

' Called when user types into txtFilterInputList: filter txtInputList content by "contains" using master text
Private Sub txtFilterInputList_Change()
    On Error Resume Next

    Dim ctrl As MSForms.TextBox
    Set ctrl = Nothing
    Set ctrl = Me.Controls("txtFilterInputList")
    If ctrl Is Nothing Then Exit Sub

    Dim ftext As String
    ftext = Trim$(CStr(ctrl.text))
    Dim lf As String
    lf = LCase$(ftext)

    ' If filter is empty -> restore original master (if stored), then clear master state
    If lf = "" Then
        If gInputListMasterText <> "" Then
            ' If user did NOT edit the visible filtered list, simply restore original master text
            If Not gInputListFilteredEdited Then
                txtInputList.text = gInputListMasterText
            Else
                ' User edited the filtered view -> merge edits back into master
                Dim baseCol As Collection
                Dim filteredOrig As Collection
                Dim currentCol As Collection
                Dim mergedDict As Object
                Dim item As Variant

                ' Expand original master and current edited text
                Set baseCol = ParseObjectList(gInputListMasterText)
                ' filtered original should have been stored when filter was first applied
                If Not gInputListFilterMasterCol Is Nothing Then
                    Set filteredOrig = gInputListFilterMasterCol
                Else
                    ' Fallback: if we don't have filtered original, infer as intersection of baseCol and current visible
                    Set filteredOrig = New Collection
                    Dim tmpBase As Collection
                    Set tmpBase = baseCol
                    Dim visCol As Collection
                    Set visCol = ParseObjectList(txtInputList.text)
                    For Each item In tmpBase
                        If IsInCollection(visCol, CStr(item)) Then filteredOrig.Add CStr(item)
                    Next item
                End If

                Set currentCol = ParseObjectList(gInputListEditedText)

                ' Build merged master: start with baseCol items, then remove items user deleted from filteredOrig,
                ' and add new items user added in currentCol.
                Set mergedDict = CreateObject("Scripting.Dictionary")

                ' Add all base items initially
                For Each item In baseCol
                    mergedDict(CStr(item)) = True
                Next item

                ' For each item that was in filteredOrig but now not in currentCol -> remove from mergedDict
                For Each item In filteredOrig
                    If Not IsInCollection(currentCol, CStr(item)) Then
                        If mergedDict.exists(CStr(item)) Then mergedDict.Remove CStr(item)
                    End If
                Next item

                ' For each item in currentCol that is not in baseCol -> add (user added new items)
                For Each item In currentCol
                    If Not mergedDict.exists(CStr(item)) Then
                        mergedDict(CStr(item)) = True
                    End If
                Next item

                ' Convert mergedDict back to a sorted collection
                Dim mergedCol As New Collection
                Dim k As Variant
                For Each k In mergedDict.keys
                    mergedCol.Add CStr(k)
                Next k
                Dim sortedMerged As Collection
                Set sortedMerged = SortCollection(mergedCol)

                ' Save merged as the new master text (compact) and show it
                txtInputList.text = CleanText(CollectionToCompactText(sortedMerged))
                gInputListMasterText = txtInputList.text
            End If

            ' Clear filter state
            gInputListMasterText = ""
            Set gInputListFilterMasterCol = Nothing
            gInputListFilteredEdited = False
            gInputListEditedText = ""
            UpdateInputListCount
        End If
        Exit Sub
    End If

    ' On first use of filter, store current txtInputList content and compute the filtered result collection
    If gInputListMasterText = "" Then
        gInputListMasterText = txtInputList.text
    End If

    ' Parse master into collection of items (ParseObjectList will expand ranges)
    Dim baseCol2 As Collection
    Set baseCol2 = ParseObjectList(gInputListMasterText)

    Dim resultCol As New Collection
    Dim it As Variant
    For Each it In baseCol2
        If InStr(LCase$(CStr(it)), lf) > 0 Then
            resultCol.Add CStr(it)
        End If
    Next it

    ' Store the filtered-original collection for later merge if user edits visible list
    Set gInputListFilterMasterCol = Nothing
    Set gInputListFilterMasterCol = New Collection
    For Each it In resultCol
        gInputListFilterMasterCol.Add CStr(it)
    Next it

    ' Reset edited flag/state
    gInputListFilteredEdited = False
    gInputListEditedText = ""

    ' Present filtered content as explicit expanded list (comma-separated) to make contents clear to the user
    Dim outText As String
    outText = CollectionToExpandedText(resultCol)
    txtInputList.text = CleanText(outText)
    UpdateInputListCount
    On Error GoTo 0
End Sub

' Convert collection to explicit comma-separated text
Private Function CollectionToExpandedText(ByVal col As Collection) As String
    If col Is Nothing Then
        CollectionToExpandedText = ""
        Exit Function
    End If
    If col.count = 0 Then
        CollectionToExpandedText = ""
        Exit Function
    End If

    Dim parts() As String
    ReDim parts(1 To col.count)
    Dim i As Long
    Dim it As Variant
    i = 1
    For Each it In col
        parts(i) = CStr(it)
        i = i + 1
    Next it

    CollectionToExpandedText = Join(parts, ",")
End Function

