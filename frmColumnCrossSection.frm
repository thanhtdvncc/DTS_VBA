VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmColumnCrossSection 
   Caption         =   "ColumnCrossSectionSettings"
   ClientHeight    =   6210
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6315
   OleObjectBlob   =   "frmColumnCrossSection.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmColumnCrossSection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'===============================================================
' UserForm: frmColumnCrossSection
' Purpose: Settings for column cross-section drawing and punching shear perimeter
' FIXED VERSION:
' 1. Separated OptionButton groups (Output Type vs Rebar Cover Definition)
' 2. Changed lstCoverSettup to 4 columns: No | Key | TopCover | BottomCover
' 3. Auto-adjust column widths based on ListBox width (10-60-15-15)
' 4. Fixed cell editing to correctly target columns 2 (Top) and 3 (Bottom)
' 5. Auto-delete PunchingCoverData sheet on form close/unload
' 6. Fixed story name mapping to match txtStoryList
' 7. Added section prefix settings for different shapes
' 8. Fixed dialog integration with M12_DrawColumnAreas
'===============================================================

Private m_StoryCount As Long
Private m_StoryElevations As Object
Private m_SlabSections As Object
Private m_UserClickedOK As Boolean

' Temp mouse/cell indices for editing
Private m_LastMouseX As Single
Private m_LastMouseY As Single
Private m_LastRowIndex As Long
Private m_LastColIndex As Long

' Preserved user-entered cover values
Private m_PreservedCoverValues As Object

' Section prefix settings
Private m_SectionPrefixes As Object  ' Dictionary: shape type -> prefix


'==========================
' Initialization
'==========================
Private Sub UserForm_Initialize()
    On Error Resume Next
    m_UserClickedOK = False
    Set m_StoryElevations = Nothing
    Set m_SlabSections = Nothing
    Set m_PreservedCoverValues = CreateObject("Scripting.Dictionary")
    Set m_SectionPrefixes = CreateObject("Scripting.Dictionary")
    m_LastMouseX = 0: m_LastMouseY = 0
    m_LastRowIndex = -1: m_LastColIndex = -1

    ' Initialize default section prefixes
    InitializeDefaultPrefixes

    ' Load story data
    Set m_StoryElevations = M12_DrawColumnAreas.GetGridlineElevations()
    If Not m_StoryElevations Is Nothing Then
        m_StoryCount = m_StoryElevations.count
    Else
        m_StoryCount = 0
    End If

    ' Set MultiPage to first page
    Dim mp          As Object
    Set mp = GetMultiPage()
    If Not mp Is Nothing Then mp.Value = 0

    ' General defaults
    If HasControl("txtStoryRange") Then GetControl("txtStoryRange").text = ""
    If HasControl("txtStoryList") Then
        If m_StoryCount > 0 Then
            GetControl("txtStoryList").text = M12_DrawColumnAreas.GetStoryListText(m_StoryElevations)
        Else
            GetControl("txtStoryList").text = "No stories detected. Please open a SAP2000 model first."
        End If
    End If

    ' FIXED: Load section prefixes into textboxes
    LoadSectionPrefixesToForm

    If HasControl("chkAddFloorSuffix") Then GetControl("chkAddFloorSuffix").Value = True

    ' Punching defaults
    If HasControl("chkEnablePunching") Then GetControl("chkEnablePunching").Value = False

    ' FIXED: Separate option groups by using GroupName property
    ' Output Type group
    If HasControl("optOutputArea") Then
        With GetControl("optOutputArea")
            .groupName = "OutputType"
            .Value = True
        End With
    End If
    If HasControl("optOutputLines") Then
        With GetControl("optOutputLines")
            .groupName = "OutputType"
            .Value = False
        End With
    End If

    ' Rebar Cover Definition Mode group
    If HasControl("optModeAuto") Then
        With GetControl("optModeAuto")
            .groupName = "CoverMode"
            .Value = True
        End With
    End If
    If HasControl("optModeByStory") Then
        With GetControl("optModeByStory")
            .groupName = "CoverMode"
            .Value = False
        End With
    End If
    If HasControl("optModeBySection") Then
        With GetControl("optModeBySection")
            .groupName = "CoverMode"
            .Value = False
        End With
    End If

    ' Combine checkbox
    If HasControl("chkCombineModes") Then
        With GetControl("chkCombineModes")
            .Value = True
            .enabled = True
            .Visible = True
        End With
    End If

    ' Auto covers defaults
    If HasControl("txtAutoTopCover") Then GetControl("txtAutoTopCover").text = "50"
    If HasControl("txtAutoBottomCover") Then GetControl("txtAutoBottomCover").text = "50"

    ' FIXED: Setup listbox with 4 columns
    If HasControl("lstCoverSettup") Then
        With GetControl("lstCoverSettup")
            .Clear
            .ColumnCount = 4
            ' Initial widths will be auto-adjusted later
            .ColumnWidths = "18;110;28;28"
            .ListStyle = fmListStylePlain
        End With
        ' Auto-adjust column widths
        AdjustListBoxColumnWidths
    End If

    ' Hide optional overlay
    If HasControl("txtEditOverlay") Then
        With GetControl("txtEditOverlay")
            On Error Resume Next
            .Visible = False
            .BorderStyle = fmBorderStyleSingle
            On Error GoTo 0
        End With
    End If

    UpdatePunchingUI
    On Error GoTo 0
End Sub

'==========================
' Section Prefix Management
'==========================
Private Sub InitializeDefaultPrefixes()
    On Error Resume Next
    Set m_SectionPrefixes = CreateObject("Scripting.Dictionary")

    ' Default prefixes for each shape type
    m_SectionPrefixes.Add "RC", "RC_COL_"
    m_SectionPrefixes.Add "I", "STEEL_I_"
    m_SectionPrefixes.Add "H", "STEEL_H_"
    m_SectionPrefixes.Add "PIPE", "STEEL_PIPE_"
    m_SectionPrefixes.Add "BOX", "STEEL_BOX_"
    m_SectionPrefixes.Add "CHANNEL", "STEEL_CH_"
    m_SectionPrefixes.Add "TEE", "STEEL_TEE_"
    m_SectionPrefixes.Add "ANGLE", "STEEL_ANG_"
    m_SectionPrefixes.Add "CIRCLE", "STEEL_CIR_"
    m_SectionPrefixes.Add "DEFAULT", "ColSec_"
End Sub

Private Sub LoadSectionPrefixesToForm()
    On Error Resume Next
    ' Load prefixes into textboxes (you need to add these textboxes to your form)
    If HasControl("txtPrefixRC") Then GetControl("txtPrefixRC").text = m_SectionPrefixes("RC")
    If HasControl("txtPrefixI") Then GetControl("txtPrefixI").text = m_SectionPrefixes("I")
    If HasControl("txtPrefixH") Then GetControl("txtPrefixH").text = m_SectionPrefixes("H")
    If HasControl("txtPrefixPipe") Then GetControl("txtPrefixPipe").text = m_SectionPrefixes("PIPE")
    If HasControl("txtPrefixBox") Then GetControl("txtPrefixBox").text = m_SectionPrefixes("BOX")
    If HasControl("txtPrefixChannel") Then GetControl("txtPrefixChannel").text = m_SectionPrefixes("CHANNEL")
    If HasControl("txtPrefixTee") Then GetControl("txtPrefixTee").text = m_SectionPrefixes("TEE")
    If HasControl("txtPrefixAngle") Then GetControl("txtPrefixAngle").text = m_SectionPrefixes("ANGLE")
    If HasControl("txtPrefixCircle") Then GetControl("txtPrefixCircle").text = m_SectionPrefixes("CIRCLE")
    If HasControl("txtPrefixDefault") Then GetControl("txtPrefixDefault").text = m_SectionPrefixes("DEFAULT")
End Sub

Private Sub SaveSectionPrefixesFromForm()
    On Error Resume Next
    ' Save prefixes from textboxes
    If HasControl("txtPrefixRC") Then m_SectionPrefixes("RC") = GetControl("txtPrefixRC").text
    If HasControl("txtPrefixI") Then m_SectionPrefixes("I") = GetControl("txtPrefixI").text
    If HasControl("txtPrefixH") Then m_SectionPrefixes("H") = GetControl("txtPrefixH").text
    If HasControl("txtPrefixPipe") Then m_SectionPrefixes("PIPE") = GetControl("txtPrefixPipe").text
    If HasControl("txtPrefixBox") Then m_SectionPrefixes("BOX") = GetControl("txtPrefixBox").text
    If HasControl("txtPrefixChannel") Then m_SectionPrefixes("CHANNEL") = GetControl("txtPrefixChannel").text
    If HasControl("txtPrefixTee") Then m_SectionPrefixes("TEE") = GetControl("txtPrefixTee").text
    If HasControl("txtPrefixAngle") Then m_SectionPrefixes("ANGLE") = GetControl("txtPrefixAngle").text
    If HasControl("txtPrefixCircle") Then m_SectionPrefixes("CIRCLE") = GetControl("txtPrefixCircle").text
    If HasControl("txtPrefixDefault") Then m_SectionPrefixes("DEFAULT") = GetControl("txtPrefixDefault").text
End Sub

'==========================
' FIXED: Auto-adjust ListBox column widths
'==========================
Private Sub AdjustListBoxColumnWidths()
    On Error Resume Next
    If Not HasControl("lstCoverSettup") Then Exit Sub

    Dim lst         As Object: Set lst = GetControl("lstCoverSettup")
    Dim totalWidth  As Single
    totalWidth = lst.Width

    ' FIXED: Column distribution: No(10%), Key(60%), Top(15%), Bottom(15%)
    Dim col1 As Long, col2 As Long, col3 As Long, col4 As Long
    col1 = Int(totalWidth * 0.1)
    col2 = Int(totalWidth * 0.5)
    col3 = Int(totalWidth * 0.2)
    col4 = Int(totalWidth * 0.2) - 4  ' Small margin for scrollbar

    ' Ensure minimum widths
    If col1 < 15 Then col1 = 15
    If col2 < 60 Then col2 = 60
    If col3 < 25 Then col3 = 25
    If col4 < 25 Then col4 = 25

    lst.ColumnWidths = col1 & ";" & col2 & ";" & col3 & ";" & col4
End Sub

'==========================
' Form Resize Event (if you want dynamic adjustment)
'==========================
Private Sub UserForm_Resize()
    On Error Resume Next
    AdjustListBoxColumnWidths
End Sub

'==========================
' Helpers: safe control access
'==========================
Private Function GetMultiPage() As Object
    On Error Resume Next
    Dim c           As Object
    Set c = Nothing
    On Error Resume Next
    Set c = Me.Controls("tabMain")
    If err.number = 0 Then
        If TypeName(c) = "MultiPage" Then
            Set GetMultiPage = c
            Exit Function
        End If
    End If
    err.Clear
    Dim ctrl        As Object
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "MultiPage" Then
            Set GetMultiPage = ctrl
            Exit Function
        End If
    Next ctrl
    Set GetMultiPage = Nothing
End Function

Private Function HasControl(ctrlName As String) As Boolean
    On Error Resume Next
    Dim c           As Object
    Set c = Nothing
    Set c = Me.Controls(ctrlName)
    If err.number <> 0 Then
        err.Clear
        HasControl = False
    Else
        HasControl = Not c Is Nothing
    End If
End Function

Private Function GetControl(ctrlName As String) As Object
    On Error Resume Next
    Dim c           As Object
    Set c = Nothing
    Set c = Me.Controls(ctrlName)
    If err.number <> 0 Then
        err.Clear
        Set GetControl = Nothing
    Else
        Set GetControl = c
    End If
End Function

'==========================
' Update UI
'==========================
Private Sub UpdatePunchingUI()
    On Error Resume Next
    Dim enabled As Boolean
    If HasControl("chkEnablePunching") Then
        enabled = CBool(GetControl("chkEnablePunching").Value)
    Else
        enabled = False
    End If

    Dim c As Object
    ' Frames that should be enabled/disabled with punching option
    If HasControl("frameOutput") Then
        Set c = GetControl("frameOutput")
        On Error Resume Next: c.enabled = enabled: On Error GoTo 0
    End If
    If HasControl("frameModes") Then
        Set c = GetControl("frameModes")
        On Error Resume Next: c.enabled = enabled: On Error GoTo 0
    End If
    If HasControl("frameAutoMode") Then
        Set c = GetControl("frameAutoMode")
        On Error Resume Next: c.enabled = enabled: On Error GoTo 0
    End If
    If HasControl("frameStoryMode") Then
        Set c = GetControl("frameStoryMode")
        On Error Resume Next: c.enabled = enabled: On Error GoTo 0
    End If
    If HasControl("frameSectionMode") Then
        Set c = GetControl("frameSectionMode")
        On Error Resume Next: c.enabled = enabled: On Error GoTo 0
    End If

    ' Buttons: keep LoadStory and LoadSections enabled always so user can populate lists
    If HasControl("btnPopulateAuto") Then
        Set c = GetControl("btnPopulateAuto")
        On Error Resume Next: c.enabled = enabled: On Error GoTo 0
    End If
    If HasControl("btnLoadStory") Then
        Set c = GetControl("btnLoadStory")
        On Error Resume Next
        c.enabled = True   ' Always enabled
        On Error GoTo 0
    End If
    If HasControl("btnLoadSections") Then
        Set c = GetControl("btnLoadSections")
        On Error Resume Next
        c.enabled = True   ' Always enabled
        On Error GoTo 0
    End If

    ' Ensure Export and Import are enabled so user can export/import PunchingCoverData regardless of punching toggle
    If HasControl("btnExport") Then
        Set c = GetControl("btnExport")
        On Error Resume Next
        c.enabled = True   ' Always enabled
        On Error GoTo 0
    End If
    If HasControl("btnImport") Then
        Set c = GetControl("btnImport")
        On Error Resume Next
        c.enabled = True   ' Always enabled
        On Error GoTo 0
    End If

    ' Show/hide mode-specific frames depending on selected mode (visibility only)
    Dim autoVal As Boolean, byStoryVal As Boolean, bySectionVal As Boolean
    autoVal = False: byStoryVal = False: bySectionVal = False
    If HasControl("optModeAuto") Then autoVal = CBool(GetControl("optModeAuto").Value)
    If HasControl("optModeByStory") Then byStoryVal = CBool(GetControl("optModeByStory").Value)
    If HasControl("optModeBySection") Then bySectionVal = CBool(GetControl("optModeBySection").Value)

    If HasControl("frameAutoMode") Then
        Set c = GetControl("frameAutoMode")
        On Error Resume Next: c.Visible = autoVal: On Error GoTo 0
    End If
    If HasControl("frameStoryMode") Then
        Set c = GetControl("frameStoryMode")
        On Error Resume Next: c.Visible = byStoryVal: On Error GoTo 0
    End If
    If HasControl("frameSectionMode") Then
        Set c = GetControl("frameSectionMode")
        On Error Resume Next: c.Visible = bySectionVal: On Error GoTo 0
    End If

    ' Ensure combine checkbox state/visibility unchanged
    If HasControl("chkCombineModes") Then
        Set c = GetControl("chkCombineModes")
        On Error Resume Next
        c.Visible = True
        c.enabled = True
        On Error GoTo 0
    End If

    On Error GoTo 0
End Sub

'==========================
' Preserve functions
'==========================
Private Sub EnsurePreserved()
    On Error Resume Next
    If m_PreservedCoverValues Is Nothing Then Set m_PreservedCoverValues = CreateObject("Scripting.Dictionary")
End Sub

Private Function StripAPrefixIfNumeric(val As String) As String
    On Error Resume Next
    val = Trim$(CStr(val))
    If val = "" Then
        StripAPrefixIfNumeric = ""
        Exit Function
    End If
    If Left$(LCase$(val), 1) = "a" Then
        StripAPrefixIfNumeric = mid$(val, 2)
    Else
        StripAPrefixIfNumeric = val
    End If
End Function

Private Function MakeDisplayValue(rawVal As String) As String
    On Error Resume Next
    rawVal = Trim$(CStr(rawVal))
    If rawVal = "" Then
        MakeDisplayValue = ""
        Exit Function
    End If
    If InStr(rawVal, "/") > 0 Then
        MakeDisplayValue = rawVal
        Exit Function
    End If
    If Left$(LCase$(rawVal), 1) = "a" Then
        MakeDisplayValue = rawVal
        Exit Function
    End If
    If IsNumeric(rawVal) Then
        MakeDisplayValue = "a" & rawVal
        Exit Function
    End If
    MakeDisplayValue = rawVal
End Function

Private Sub SaveCurrentListToPreserve()
    On Error Resume Next
    If Not HasControl("lstCoverSettup") Then Exit Sub
    EnsurePreserved
    Dim lst         As Object: Set lst = GetControl("lstCoverSettup")
    Dim i           As Long
    For i = 0 To lst.ListCount - 1
        Dim key     As String: key = ""
        Dim topRaw  As String: topRaw = ""
        Dim botRaw  As String: botRaw = ""
        ' FIXED: Access columns by index, not vbTab split
        ' Column 1 contains the story name, we need to map it back to story index for preservation
        On Error Resume Next
        Dim storyName As String: storyName = Trim$(CStr(lst.List(i, 1)))  ' Column 1 = Story Name or Section
        topRaw = StripAPrefixIfNumeric(CStr(lst.List(i, 2)))  ' Column 2 = Top
        botRaw = StripAPrefixIfNumeric(CStr(lst.List(i, 3)))  ' Column 3 = Bottom
        On Error GoTo 0

        ' Use story name as key for preservation
        key = storyName

        If key <> "" Then
            If m_PreservedCoverValues.exists(key) Then
                m_PreservedCoverValues(key) = Array(topRaw, botRaw)
            Else
                m_PreservedCoverValues.Add key, Array(topRaw, botRaw)
            End If
        End If
    Next i
End Sub

Private Sub UpdatePreservedForKey(key As String, topRaw As String, botRaw As String)
    On Error Resume Next
    If Trim$(key) = "" Then Exit Sub
    EnsurePreserved
    If m_PreservedCoverValues.exists(key) Then
        m_PreservedCoverValues(key) = Array(StripAPrefixIfNumeric(topRaw), StripAPrefixIfNumeric(botRaw))
    Else
        m_PreservedCoverValues.Add key, Array(StripAPrefixIfNumeric(topRaw), StripAPrefixIfNumeric(botRaw))
    End If
End Sub

Private Function GetPreservedForKey(key As String) As Variant
    On Error Resume Next
    EnsurePreserved

    ' First try exact match
    If m_PreservedCoverValues.exists(key) Then
        GetPreservedForKey = m_PreservedCoverValues(key)
        Exit Function
    End If

    ' If key is numeric (story index), try to find by story name
    If IsNumeric(key) Then
        Dim storyName As String
        storyName = GetStoryNameByIndex(CLng(key))
        If m_PreservedCoverValues.exists(storyName) Then
            GetPreservedForKey = m_PreservedCoverValues(storyName)
            Exit Function
        End If
    End If

    GetPreservedForKey = Empty
End Function

'==========================
' FIXED: Populate/Load with 4 columns (No | Key | Top | Bottom)
'==========================
Private Sub btnPopulateAuto_Click()
    On Error Resume Next
    If Not HasControl("lstCoverSettup") Then Exit Sub
    Dim lst         As Object: Set lst = GetControl("lstCoverSettup")

    SaveCurrentListToPreserve
    lst.Clear

    If m_StoryCount = 0 Then
        MsgBox "No stories available. Please open a SAP2000 model.", vbExclamation, "No Data"
        Exit Sub
    End If

    Dim topDef As String, botDef As String
    If HasControl("txtAutoTopCover") Then topDef = Trim$(GetControl("txtAutoTopCover").text) Else topDef = "50"
    If HasControl("txtAutoBottomCover") Then botDef = Trim$(GetControl("txtAutoBottomCover").text) Else botDef = "50"
    If topDef = "" Then topDef = "50"
    If botDef = "" Then botDef = "50"

    Dim i           As Long
    For i = 1 To m_StoryCount
        Dim storyIndex As Long: storyIndex = m_StoryCount - i + 1  ' Reverse order
        Dim key     As String: key = CStr(storyIndex)

        ' FIXED: Get story name from elevations dictionary
        Dim storyName As String
        storyName = GetStoryNameByIndex(storyIndex)

        Dim preserved As Variant: preserved = GetPreservedForKey(key)
        Dim displayTop As String, displayBot As String
        If Not IsEmpty(preserved) Then
            displayTop = MakeDisplayValue(CStr(preserved(0)))
            displayBot = MakeDisplayValue(CStr(preserved(1)))
        Else
            displayTop = MakeDisplayValue(topDef)
            displayBot = MakeDisplayValue(botDef)
        End If
        ' FIXED: Add item then set columns individually with story name
        lst.AddItem
        lst.List(lst.ListCount - 1, 0) = CStr(i)
        lst.List(lst.ListCount - 1, 1) = storyName  ' Use story name instead of index
        lst.List(lst.ListCount - 1, 2) = displayTop
        lst.List(lst.ListCount - 1, 3) = displayBot
    Next i
End Sub

Private Sub btnLoadStory_Click()
    btnPopulateAuto_Click
End Sub

' FIXED: Helper function to get story name by index
Private Function GetStoryNameByIndex(storyIndex As Long) As String
    On Error Resume Next
    GetStoryNameByIndex = CStr(storyIndex)  ' Default to index

    If m_StoryElevations Is Nothing Then Exit Function
    If m_StoryCount = 0 Then Exit Function

    ' m_StoryElevations is a Dictionary with keys as story names
    ' We need to get the story name at the given index (1-based, top to bottom)
    Dim keys        As Variant
    keys = m_StoryElevations.keys

    Dim elevations() As Double
    ReDim elevations(0 To UBound(keys))

    Dim i           As Long
    For i = 0 To UBound(keys)
        elevations(i) = CDbl(m_StoryElevations(keys(i)))
    Next i

    ' Sort by elevation (descending - highest first)
    Dim j As Long, k As Long
    For j = 0 To UBound(keys) - 1
        For k = j + 1 To UBound(keys)
            If elevations(j) < elevations(k) Then
                ' Swap elevations
                Dim tempElev As Double: tempElev = elevations(j)
                elevations(j) = elevations(k)
                elevations(k) = tempElev
                ' Swap keys
                Dim tempKey As Variant: tempKey = keys(j)
                keys(j) = keys(k)
                keys(k) = tempKey
            End If
        Next k
    Next j

    ' Get story name at index (1-based)
    If storyIndex >= 1 And storyIndex <= UBound(keys) + 1 Then
        GetStoryNameByIndex = CStr(keys(storyIndex - 1))
    End If
End Function

Private Sub btnLoadSections_Click()
    On Error Resume Next
    If Not HasControl("lstCoverSettup") Then Exit Sub
    Dim lst         As Object: Set lst = GetControl("lstCoverSettup")

    SaveCurrentListToPreserve
    lst.Clear

    Set m_SlabSections = DetectSlabSectionsFromModel()
    If m_SlabSections Is Nothing Or m_SlabSections.count = 0 Then
        MsgBox "No slab sections detected in model.", vbInformation, "No Sections"
        Exit Sub
    End If

    Dim topDef As String, botDef As String
    If HasControl("txtAutoTopCover") Then topDef = Trim$(GetControl("txtAutoTopCover").text) Else topDef = "50"
    If HasControl("txtAutoBottomCover") Then botDef = Trim$(GetControl("txtAutoBottomCover").text) Else botDef = "50"
    If topDef = "" Then topDef = "50"
    If botDef = "" Then botDef = "50"

    Dim sec As Variant, idx As Long
    idx = 1
    For Each sec In m_SlabSections.keys
        Dim key     As String: key = CStr(sec)
        Dim preserved As Variant: preserved = GetPreservedForKey(key)
        Dim displayTop As String, displayBot As String
        If Not IsEmpty(preserved) Then
            displayTop = MakeDisplayValue(CStr(preserved(0)))
            displayBot = MakeDisplayValue(CStr(preserved(1)))
        Else
            displayTop = MakeDisplayValue(topDef)
            displayBot = MakeDisplayValue(botDef)
        End If
        ' FIXED: Add item then set columns individually
        lst.AddItem
        lst.List(lst.ListCount - 1, 0) = CStr(idx)
        lst.List(lst.ListCount - 1, 1) = key
        lst.List(lst.ListCount - 1, 2) = displayTop
        lst.List(lst.ListCount - 1, 3) = displayBot
        idx = idx + 1
    Next sec
End Sub

'==========================
' Detect slab sections
'==========================
Private Function DetectSlabSectionsFromModel() As Object
    On Error Resume Next
    Set DetectSlabSectionsFromModel = CreateObject("Scripting.Dictionary")
    Dim numAreas    As Long
    Dim areaNames() As String
    If SapModel.AreaObj.GetNameList(numAreas, areaNames) = 0 Then
        Dim i       As Long
        For i = 0 To numAreas - 1
            Dim propName As String
            If SapModel.AreaObj.GetProperty(areaNames(i), propName) = 0 Then
                If Trim$(propName) <> "" And propName <> "None" Then
                    If Not DetectSlabSectionsFromModel.exists(propName) Then DetectSlabSectionsFromModel.Add propName, True
                End If
            End If
        Next i
    End If
End Function

'==========================
' FIXED: List editing with correct column indices
'==========================
Private Sub lstCoverSettup_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    On Error Resume Next
    If Not HasControl("lstCoverSettup") Then Exit Sub
    Dim lst         As Object: Set lst = GetControl("lstCoverSettup")
    Dim idx         As Long: idx = lst.ListIndex
    If idx < 0 Then Exit Sub

    Dim colIdx      As Long: colIdx = m_LastColIndex
    If colIdx < 0 Then
        colIdx = GetColumnIndexFromX(lst, m_LastMouseX)
    End If

    ' FIXED: Column mapping: 0=No, 1=Key, 2=Top, 3=Bottom
    ' Only allow editing columns 2 (Top) and 3 (Bottom)
    If colIdx = 2 Or colIdx = 3 Then
        EditCellWithInputBox idx, colIdx
    ElseIf colIdx <= 1 Then
        ' If clicked on No or Key, edit both Top & Bottom
        EditRowWithInputBoxes idx
    End If
End Sub

Private Sub lstCoverSettup_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    On Error Resume Next
    If Not HasControl("lstCoverSettup") Then Exit Sub
    m_LastMouseX = X: m_LastMouseY = Y
    Dim lst         As Object: Set lst = GetControl("lstCoverSettup")
    Dim idx         As Long: idx = lst.ListIndex
    If idx < 0 Then
        Dim rowHeight As Long: rowHeight = lst.Font.Size * 15
        idx = Int(Y / rowHeight)
        If idx < 0 Or idx > lst.ListCount - 1 Then
            m_LastRowIndex = -1
            m_LastColIndex = -1
            Exit Sub
        End If
        lst.ListIndex = idx
    End If
    m_LastRowIndex = lst.ListIndex
    m_LastColIndex = GetColumnIndexFromX(lst, X)
End Sub

Private Function GetColumnIndexFromX(lst As Object, X As Single) As Long
    On Error Resume Next
    Dim widths()    As String: widths = Split(lst.ColumnWidths, ";")
    Dim i As Long, total As Double: total = 0
    For i = LBound(widths) To UBound(widths)
        total = total + CDbl(Trim$(widths(i)))
    Next i
    If total <= 0 Then
        GetColumnIndexFromX = Int((X / lst.Width) * lst.ColumnCount)
        If GetColumnIndexFromX < 0 Then GetColumnIndexFromX = 0
        If GetColumnIndexFromX > lst.ColumnCount - 1 Then GetColumnIndexFromX = lst.ColumnCount - 1
        Exit Function
    End If
    Dim acc         As Double: acc = 0
    For i = LBound(widths) To UBound(widths)
        acc = acc + CDbl(Trim$(widths(i)))
        If X <= (acc / total) * lst.Width Then
            GetColumnIndexFromX = i
            Exit Function
        End If
    Next i
    GetColumnIndexFromX = UBound(widths)
    If GetColumnIndexFromX < 0 Then GetColumnIndexFromX = 0
End Function

' FIXED: Edit both Top and Bottom
Private Sub EditRowWithInputBoxes(rowIndex As Long)
    On Error Resume Next
    Dim lst         As Object: Set lst = GetControl("lstCoverSettup")
    If lst Is Nothing Then Exit Sub
    If rowIndex < 0 Or rowIndex > lst.ListCount - 1 Then Exit Sub

    Dim no As String, key As String, curTop As String, curBot As String
    On Error Resume Next
    no = CStr(lst.List(rowIndex, 0))
    key = CStr(lst.List(rowIndex, 1))
    curTop = CStr(lst.List(rowIndex, 2))
    curBot = CStr(lst.List(rowIndex, 3))
    On Error GoTo 0

    Dim newTop      As String: newTop = InputBox("Top cover (a50 or 1/15 or number):", "Edit Top", curTop)
    If StrPtr(newTop) = 0 Then Exit Sub
    Dim newBot      As String: newBot = InputBox("Bottom cover (a50 or 1/15 or number):", "Edit Bottom", curBot)
    If StrPtr(newBot) = 0 Then Exit Sub

    ' Update the row
    lst.List(rowIndex, 2) = newTop
    lst.List(rowIndex, 3) = newBot

    UpdatePreservedForKey key, StripAPrefixIfNumeric(newTop), StripAPrefixIfNumeric(newBot)
End Sub

' FIXED: Edit single cell (Top or Bottom only)
Private Sub EditCellWithInputBox(rowIndex As Long, colIndex As Long)
    On Error Resume Next
    Dim lst         As Object: Set lst = GetControl("lstCoverSettup")
    If lst Is Nothing Then Exit Sub
    If rowIndex < 0 Or rowIndex > lst.ListCount - 1 Then Exit Sub

    ' FIXED: Only allow editing columns 2 (Top) and 3 (Bottom)
    If colIndex <> 2 And colIndex <> 3 Then Exit Sub

    Dim curVal      As String
    On Error Resume Next
    curVal = CStr(lst.List(rowIndex, colIndex))
    On Error GoTo 0

    Dim colName     As String
    If colIndex = 2 Then colName = "Top" Else colName = "Bottom"

    Dim newVal      As String: newVal = InputBox("Edit " & colName & " cover (a50 or 1/15 or number):", "Edit " & colName, curVal)
    If StrPtr(newVal) = 0 Then Exit Sub

    ' Update the cell
    lst.List(rowIndex, colIndex) = newVal

    ' Update preserved
    Dim key As String, topVal As String, botVal As String
    On Error Resume Next
    key = CStr(lst.List(rowIndex, 1))
    topVal = CStr(lst.List(rowIndex, 2))
    botVal = CStr(lst.List(rowIndex, 3))
    On Error GoTo 0

    If key <> "" Then UpdatePreservedForKey key, topVal, botVal
End Sub

Public Sub CommitOverlayEditGeneric()
    On Error Resume Next
    Dim overlay     As Object: Set overlay = GetControl("txtEditOverlay")
    If overlay Is Nothing Then Exit Sub
    If Not overlay.Visible Then Exit Sub
    Dim tag         As String: tag = CStr(overlay.tag)
    If InStr(tag, ":") = 0 Then overlay.Visible = False: Exit Sub
    Dim t()         As String: t = Split(tag, ":")
    Dim rowIdx      As Long: rowIdx = CLng(t(0))
    Dim colIdx      As Long: colIdx = CLng(t(1))

    ' FIXED: Only allow editing columns 2 and 3
    If colIdx <> 2 And colIdx <> 3 Then overlay.Visible = False: Exit Sub

    Dim lst         As Object: Set lst = GetControl("lstCoverSettup")
    If lst Is Nothing Then overlay.Visible = False: Exit Sub
    If rowIdx < 0 Or rowIdx > lst.ListCount - 1 Then overlay.Visible = False: Exit Sub

    ' Update the cell directly
    lst.List(rowIdx, colIdx) = CStr(overlay.text)

    ' Update preserved
    Dim key As String, topVal As String, botVal As String
    On Error Resume Next
    key = CStr(lst.List(rowIdx, 1))
    topVal = CStr(lst.List(rowIdx, 2))
    botVal = CStr(lst.List(rowIdx, 3))
    On Error GoTo 0

    If key <> "" Then UpdatePreservedForKey key, topVal, botVal

    overlay.Visible = False
End Sub

'==========================
' FIXED: Export with 4 columns
'==========================
Private Sub btnExport_Click()
    On Error GoTo ErrHandler
    Dim ws As Worksheet, wb As Workbook
    Set wb = ThisWorkbook
    Set ws = Nothing
    On Error Resume Next
    Set ws = wb.Worksheets("PunchingCoverData")
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = "PunchingCoverData"
    End If
    On Error GoTo ErrHandler

    ws.Cells.Clear
    ws.Range("A1").Value = "No"
    ws.Range("B1").Value = "Key"
    ws.Range("C1").Value = "TopCover"
    ws.Range("D1").Value = "BottomCover"

    If Not HasControl("lstCoverSettup") Then Exit Sub
    Dim lst         As Object: Set lst = GetControl("lstCoverSettup")
    Dim i As Long, r As Long: r = 2
    For i = 0 To lst.ListCount - 1
        ' FIXED: Access columns by index
        On Error Resume Next
        ws.Cells(r, 1).Value = lst.List(i, 0)  ' No
        ws.Cells(r, 2).Value = lst.List(i, 1)  ' Key
        ws.Cells(r, 3).Value = lst.List(i, 2)  ' Top
        ws.Cells(r, 4).Value = lst.List(i, 3)  ' Bottom
        On Error GoTo ErrHandler
        r = r + 1
    Next i

    'MsgBox "Exported " & (r - 2) & " rows to sheet 'PunchingCoverData'.", vbInformation, "Export Complete"
    Exit Sub
ErrHandler:
    MsgBox "Error during export: " & err.description, vbExclamation, "Export Error"
End Sub

Private Sub btnImport_Click()
    On Error GoTo ErrHandler
    Dim ws As Worksheet, wb As Workbook
    Set wb = ThisWorkbook
    Set ws = Nothing
    On Error Resume Next
    Set ws = wb.Worksheets("PunchingCoverData")
    On Error GoTo ErrHandler
    If ws Is Nothing Then MsgBox "Sheet 'PunchingCoverData' not found.", vbExclamation: Exit Sub

    If Not HasControl("lstCoverSettup") Then Exit Sub
    Dim lst         As Object: Set lst = GetControl("lstCoverSettup")
    lst.Clear
    Dim r           As Long: r = 2
    Do While Trim$(CStr(ws.Cells(r, 2).Value)) <> ""  ' Check Key column (B)
        Dim no As String, k As String, t As String, b As String
        no = CStr(ws.Cells(r, 1).Value)
        k = CStr(ws.Cells(r, 2).Value)
        t = CStr(ws.Cells(r, 3).Value)
        b = CStr(ws.Cells(r, 4).Value)
        If t = "" And HasControl("txtAutoTopCover") Then t = MakeDisplayValue(GetControl("txtAutoTopCover").text)
        If b = "" And HasControl("txtAutoBottomCover") Then b = MakeDisplayValue(GetControl("txtAutoBottomCover").text)
        If no = "" Then no = CStr(r - 1)

        ' FIXED: Add using List property
        lst.AddItem
        lst.List(lst.ListCount - 1, 0) = no
        lst.List(lst.ListCount - 1, 1) = k
        lst.List(lst.ListCount - 1, 2) = t
        lst.List(lst.ListCount - 1, 3) = b
        r = r + 1
    Loop

    SaveCurrentListToPreserve

    'MsgBox "Imported rows from 'PunchingCoverData': " & lst.ListCount, vbInformation, "Import Complete"
    Exit Sub
ErrHandler:
    MsgBox "Error during import: " & err.description, vbExclamation, "Import Error"
End Sub

'==========================
' FIXED: Auto-delete PunchingCoverData sheet on form close
'==========================
Private Sub DeletePunchingCoverDataSheet()
    On Error Resume Next
    Dim ws          As Worksheet
    Dim wb          As Workbook
    Set wb = ThisWorkbook
    Set ws = wb.Worksheets("PunchingCoverData")
    If Not ws Is Nothing Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' This event fires when user clicks X or form is unloaded
    DeletePunchingCoverDataSheet
End Sub

Private Sub UserForm_Terminate()
    ' Additional cleanup when form object is destroyed
    DeletePunchingCoverDataSheet
End Sub

'==========================
' Buttons OK / Cancel / Apply
'==========================
Private Sub btnOK_Click()
    On Error GoTo ErrorHandler
Debug.Print "btnOK_Click: Start"

    ' Validate settings first
    If Not ValidateAllSettings() Then
Debug.Print "btnOK_Click: Validation failed"
        Exit Sub
    End If

    ' Save current state
    SaveCurrentListToPreserve
    SaveSectionPrefixesFromForm

    ' Set flag to True
    m_UserClickedOK = True
Debug.Print "btnOK_Click: m_UserClickedOK set to True"

    ' Get settings
    Dim settingsDict As Object
    Set settingsDict = GetSettings()

    If settingsDict Is Nothing Then
        MsgBox "Failed to generate settings.", vbExclamation, "Error"
        Exit Sub
    End If

Debug.Print "btnOK_Click: Settings generated, count = " & settingsDict.count

    ' Hide form before calling drawing function
    Me.Hide

    ' Call the main drawing function with settings
Debug.Print "btnOK_Click: Calling DrawColumnCrossSectionsWithSettings"
    Call M12_DrawColumnAreas.DrawColumnCrossSectionsWithSettings(settingsDict)

    ' Close and cleanup
    DeletePunchingCoverDataSheet
    Unload Me

Debug.Print "btnOK_Click: Complete"
    Exit Sub

ErrorHandler:
    MsgBox "ERROR in btnOK_Click: " & err.description, vbCritical, "Error"
Debug.Print "btnOK_Click: ERROR - " & err.description
End Sub
' CRITICAL: Public function to check if user clicked OK
Public Function UserClickedOK() As Boolean
Debug.Print "UserClickedOK() called: returning " & m_UserClickedOK
    UserClickedOK = m_UserClickedOK
End Function

Private Sub btnCancel_Click()
Debug.Print "btnCancel_Click: Start"

    ' Keep as False when cancelled
    m_UserClickedOK = False

Debug.Print "btnCancel_Click: m_UserClickedOK = False"

    DeletePunchingCoverDataSheet
    Unload Me

Debug.Print "btnCancel_Click: Form unloaded"
End Sub

Private Sub btnApply_Click()
Debug.Print "btnApply_Click: Start"

    If ValidateAllSettings() Then
        SaveCurrentListToPreserve
        SaveSectionPrefixesFromForm
        MsgBox "Settings validated and saved successfully!" & vbCrLf & _
                "Click OK to draw column sections.", vbInformation, "Validation Success"
    End If
End Sub

Private Sub btnApplyOther_Click()
    On Error Resume Next
    If Not HasControl("lstCoverSettup") Then Exit Sub
    Dim lst         As Object: Set lst = GetControl("lstCoverSettup")
    Dim selIdx      As Long: selIdx = lst.ListIndex
    If selIdx < 0 Then
        MsgBox "Please select a source row first.", vbExclamation, "No Selection"
        Exit Sub
    End If

    ' FIXED: Access columns by index
    Dim srcTop As String, srcBot As String
    On Error Resume Next
    srcTop = CStr(lst.List(selIdx, 2))  ' Column 2 = Top
    srcBot = CStr(lst.List(selIdx, 3))  ' Column 3 = Bottom
    On Error GoTo 0

    Dim i           As Long
    For i = selIdx + 1 To lst.ListCount - 1
        ' Update only Top and Bottom columns
        lst.List(i, 2) = srcTop
        lst.List(i, 3) = srcBot

        ' Update preserved
        Dim tgtKey  As String
        On Error Resume Next
        tgtKey = CStr(lst.List(i, 1))
        On Error GoTo 0
        If tgtKey <> "" Then UpdatePreservedForKey tgtKey, srcTop, srcBot
    Next i
End Sub


'==========================
' Validation & GetSettings
'==========================
Private Function ValidateAllSettings() As Boolean
    ValidateAllSettings = False
    If HasControl("txtStoryRange") Then
        If Trim$(GetControl("txtStoryRange").text) <> "" Then
            If Not IsValidStoryRange(GetControl("txtStoryRange").text) Then
                MsgBox "Invalid story range format." & vbCrLf & "Valid examples: 1,2-5,7 or 1,2to5,7", vbExclamation, "Validation Error"
                Dim mp As Object: Set mp = GetMultiPage()
                If Not mp Is Nothing Then mp.Value = 0
                GetControl("txtStoryRange").SetFocus
                Exit Function
            End If
        End If
    End If

    If HasControl("chkEnablePunching") Then
        If CBool(GetControl("chkEnablePunching").Value) Then
            If HasControl("txtAutoTopCover") Then
                If Not IsValidCoverValue(GetControl("txtAutoTopCover").text) Then
                    MsgBox "Invalid auto top cover format.", vbExclamation, "Validation Error"
                    Set mp = GetMultiPage(): If Not mp Is Nothing Then mp.Value = 1
                    GetControl("txtAutoTopCover").SetFocus
                    Exit Function
                End If
            End If
            If HasControl("txtAutoBottomCover") Then
                If Not IsValidCoverValue(GetControl("txtAutoBottomCover").text) Then
                    MsgBox "Invalid auto bottom cover format.", vbExclamation, "Validation Error"
                    Set mp = GetMultiPage(): If Not mp Is Nothing Then mp.Value = 1
                    GetControl("txtAutoBottomCover").SetFocus
                    Exit Function
                End If
            End If
            If HasControl("lstCoverSettup") Then
                Dim lst As Object: Set lst = GetControl("lstCoverSettup")
                Dim i As Long
                For i = 0 To lst.ListCount - 1
                    Dim rowText As String: rowText = lst.List(i)
                    Dim p() As String: p = Split(rowText, vbTab)
                    If UBound(p) >= 3 Then
                        If Not IsValidCoverValue(Trim$(p(2))) Then MsgBox "Invalid Top Cover in row " & (i + 1): Exit Function
                        If Not IsValidCoverValue(Trim$(p(3))) Then MsgBox "Invalid Bottom Cover in row " & (i + 1): Exit Function
                    End If
                Next i
            End If
        End If
    End If

    ValidateAllSettings = True
End Function

Private Function IsValidStoryRange(rangeStr As String) As Boolean
    On Error Resume Next
    IsValidStoryRange = True
    Dim i           As Long
    For i = 1 To Len(rangeStr)
        Dim ch      As String: ch = mid$(rangeStr, i, 1)
        If Not (IsNumeric(ch) Or ch = "," Or ch = "-" Or ch = " " Or LCase$(ch) = "t" Or LCase$(ch) = "o") Then
            IsValidStoryRange = False: Exit Function
        End If
    Next i
End Function

Private Function IsValidCoverValue(coverStr As String) As Boolean
    On Error Resume Next
    IsValidCoverValue = False
    coverStr = Trim$(coverStr)
    If coverStr = "" Then Exit Function
    If InStr(coverStr, "/") > 0 Then
        Dim parts() As String: parts = Split(coverStr, "/")
        If UBound(parts) = 1 Then If IsNumeric(parts(0)) And IsNumeric(parts(1)) Then IsValidCoverValue = True: Exit Function
    End If
    If Left$(LCase$(coverStr), 1) = "a" Then If IsNumeric(mid$(coverStr, 2)) Then IsValidCoverValue = True: Exit Function
    If IsNumeric(coverStr) Then IsValidCoverValue = True
End Function

Public Function GetSettings() As Object
    Dim settings    As Object
    Set settings = CreateObject("Scripting.Dictionary")

    On Error Resume Next
    If HasControl("txtStoryRange") Then
        settings.Add "storyRange", CStr(GetControl("txtStoryRange").text)
    Else
        settings.Add "storyRange", ""
    End If
    If HasControl("txtAreaPrefix") Then
        settings.Add "AreaPrefix", CStr(GetControl("txtAreaPrefix").text)
    Else
        settings.Add "AreaPrefix", ""
    End If
    If HasControl("chkAddFloorSuffix") Then
        settings.Add "AddFloorSuffix", CBool(GetControl("chkAddFloorSuffix").Value)
    Else
        settings.Add "AddFloorSuffix", False
    End If

    If HasControl("chkEnablePunching") Then
        settings.Add "DrawPunching", CBool(GetControl("chkEnablePunching").Value)
    Else
        settings.Add "DrawPunching", False
    End If

    Dim outType     As String
    outType = "AREA"
    If HasControl("optOutputArea") Then
        If CBool(GetControl("optOutputArea").Value) Then outType = "AREA" Else outType = "LINES"
    End If
    settings.Add "PunchingOutputType", outType

    ' --- NEW: include SectionPrefixes built from m_SectionPrefixes so M12_DrawColumnAreas can use correct prefixes
    Dim prefixDict  As Object
    Set prefixDict = CreateObject("Scripting.Dictionary")
    If Not m_SectionPrefixes Is Nothing Then
        Dim spk     As Variant
        For Each spk In m_SectionPrefixes.keys
            On Error Resume Next
            prefixDict.Add spk, CStr(m_SectionPrefixes(spk))
            On Error GoTo 0
        Next spk
    End If
    settings.Add "SectionPrefixes", prefixDict
    ' --- end new

    Dim coverDict   As Object
    Set coverDict = CreateObject("Scripting.Dictionary")
    Dim slabRef     As Double: slabRef = 300
    Dim autoTop As Double, autoBottom As Double
    If HasControl("txtAutoTopCover") Then
        autoTop = ParseCoverValue(CStr(GetControl("txtAutoTopCover").text), slabRef)
    Else
        autoTop = 50
    End If
    If HasControl("txtAutoBottomCover") Then
        autoBottom = ParseCoverValue(CStr(GetControl("txtAutoBottomCover").text), slabRef)
    Else
        autoBottom = 50
    End If
    Dim arrAuto(0 To 1) As Double: arrAuto(0) = autoTop: arrAuto(1) = autoBottom
    coverDict.Add "AUTO", arrAuto

    If HasControl("lstCoverSettup") Then
        Dim lst     As Object: Set lst = GetControl("lstCoverSettup")
        Dim i       As Long
        For i = 0 To lst.ListCount - 1
            ' FIXED: Access columns by index
            ' Column 1 now contains story name (not index), need to map back
            Dim storyName As String, topVal As Double, botVal As Double
            On Error Resume Next
            storyName = Trim$(CStr(lst.List(i, 1)))  ' Column 1 = Story Name or Section
            topVal = ParseCoverValue(CStr(lst.List(i, 2)), slabRef)  ' Column 2 = Top
            botVal = ParseCoverValue(CStr(lst.List(i, 3)), slabRef)  ' Column 3 = Bottom
            On Error GoTo 0

            If storyName <> "" Then
                ' Try to get story index from name
                Dim storyIdx As Long: storyIdx = GetStoryIndexByName(storyName)

                If storyIdx > 0 Then
                    ' It's a story
                    Dim sk As String: sk = "STORY_" & CStr(storyIdx)
                    If Not coverDict.exists(sk) Then
                        Dim sarr(0 To 1) As Double: sarr(0) = topVal: sarr(1) = botVal
                        coverDict.Add sk, sarr
                    End If
                Else
                    ' It's a section name
                    Dim seck As String: seck = "SECTION_" & storyName
                    If Not coverDict.exists(seck) Then
                        Dim carr(0 To 1) As Double: carr(0) = topVal: carr(1) = botVal
                        coverDict.Add seck, carr
                    End If
                End If
            End If
        Next i
    End If

    settings.Add "CoverSettings", coverDict
    Set GetSettings = settings
End Function

' FIXED: Helper function to get story index from name
Private Function GetStoryIndexByName(storyName As String) As Long
    On Error Resume Next
    GetStoryIndexByName = 0  ' Default: not found

    If m_StoryElevations Is Nothing Then Exit Function
    If m_StoryCount = 0 Then Exit Function

    ' Get sorted story names (same logic as GetStoryNameByIndex)
    Dim keys        As Variant
    keys = m_StoryElevations.keys

    Dim elevations() As Double
    ReDim elevations(0 To UBound(keys))

    Dim i           As Long
    For i = 0 To UBound(keys)
        elevations(i) = CDbl(m_StoryElevations(keys(i)))
    Next i

    ' Sort by elevation (descending)
    Dim j As Long, k As Long
    For j = 0 To UBound(keys) - 1
        For k = j + 1 To UBound(keys)
            If elevations(j) < elevations(k) Then
                Dim tempElev As Double: tempElev = elevations(j)
                elevations(j) = elevations(k)
                elevations(k) = tempElev
                Dim tempKey As Variant: tempKey = keys(j)
                keys(j) = keys(k)
                keys(k) = tempKey
            End If
        Next k
    Next j

    ' Find index of story name
    For i = 0 To UBound(keys)
        If CStr(keys(i)) = storyName Then
            GetStoryIndexByName = i + 1  ' 1-based index
            Exit Function
        End If
    Next i
End Function

Private Function ParseCoverValue(coverStr As String, slabThickness As Double) As Double
    On Error Resume Next
    ParseCoverValue = 50
    If Trim$(coverStr) = "" Then Exit Function
    If InStr(coverStr, "/") > 0 Then
        Dim p()     As String: p = Split(coverStr, "/")
        If UBound(p) = 1 Then ParseCoverValue = slabThickness * CDbl(p(0)) / CDbl(p(1)): Exit Function
    End If
    If Left$(LCase$(coverStr), 1) = "a" Then ParseCoverValue = CDbl(mid$(coverStr, 2)): Exit Function
    If IsNumeric(coverStr) Then ParseCoverValue = CDbl(coverStr)
End Function

'==========================
' Auto-format textboxes
'==========================
Private Sub txtAutoTopCover_AfterUpdate()
    On Error Resume Next
    If Not HasControl("txtAutoTopCover") Then Exit Sub
    Dim s           As String: s = Trim$(GetControl("txtAutoTopCover").text)
    If s = "" Then Exit Sub
    If Left$(LCase$(s), 1) = "a" Then Exit Sub
    If InStr(s, "/") > 0 Then Exit Sub
    If IsNumeric(s) Then GetControl("txtAutoTopCover").text = "a" & s
End Sub

Private Sub txtAutoBottomCover_AfterUpdate()
    On Error Resume Next
    If Not HasControl("txtAutoBottomCover") Then Exit Sub
    Dim s           As String: s = Trim$(GetControl("txtAutoBottomCover").text)
    If s = "" Then Exit Sub
    If Left$(LCase$(s), 1) = "a" Then Exit Sub
    If InStr(s, "/") > 0 Then Exit Sub
    If IsNumeric(s) Then GetControl("txtAutoBottomCover").text = "a" & s
End Sub

' End of UserForm module

