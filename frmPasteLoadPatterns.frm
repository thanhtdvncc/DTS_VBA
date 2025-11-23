VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPasteLoadPatterns 
   Caption         =   "Load Pattern Data Input (by thanhtdvncc):"
   ClientHeight    =   4396
   ClientLeft      =   120
   ClientTop       =   675
   ClientWidth     =   7230
   OleObjectBlob   =   "frmPasteLoadPatterns.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPasteLoadPatterns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'===============================================================
' UserForm: frmPasteLoadPatterns
' Purpose: Preview and edit load patterns before creating in SAP2000
' Version: updated - checkboxes shown but disabled at startup
'===============================================================

Private previewedPatterns As Object
Private isPreviewMode As Boolean
Private originalInputText As String

Private Sub UserForm_Initialize()
    ' Set example text (NO comments - directly usable)
    
    txtPatternData.BackColor = RGB(255, 255, 155)      ' yellow - indicates preview mode
    txtPatternData.text = _
    "# Paste or type your load patterns below" & vbCrLf & _
    "# Format: PatternName : Description" & vbCrLf & _
    "# Example patterns:" & vbCrLf & vbCrLf & _
    "DL : Dead Load" & vbCrLf & _
    "SDL : Super Dead Load" & vbCrLf & _
    "LL : Live Load" & vbCrLf & _
    "RL : Roof Live Load" & vbCrLf & _
    "WLX : Wind Load in X Direction" & vbCrLf & _
    "EQX : Earthquake Load in X Direction" & vbCrLf & vbCrLf & _
    "Tip: You can copy a range from Excel and paste it here directly as text."
    
    txtPatternData.SetFocus
    txtPatternData.SelStart = 0
    'txtPatternData.SelLength = Len(txtPatternData.Text)
    
    isPreviewMode = False
    btnOK.Caption = "Preview >>>"
    
    ' Initialize original input holder and Cancel caption
    originalInputText = ""
    btnCancel.Caption = "Cancel"
    
    ' Ensure the two checkboxes are visible but disabled at startup
    On Error Resume Next
    With chkOverwritePatterns
        .Visible = True
        .enabled = False    ' disabled initially until parsing/preview
        .Caption = "Overwrite with new LoadPattern"
        .Value = False
    End With
    
    With chkDeleteLoadcases
        .Visible = True
        .enabled = False    ' disabled until overwrite is enabled and checked
        .Caption = "Delete old Loadcase (except Modal)"
        .Value = False
    End With
    On Error GoTo 0
End Sub

' When user toggles Overwrite (only active after preview), enable/disable delete option
Private Sub chkOverwritePatterns_Click()
    On Error Resume Next
    If chkOverwritePatterns.Value = True Then
        chkDeleteLoadcases.enabled = True
    Else
        chkDeleteLoadcases.enabled = False
        chkDeleteLoadcases.Value = False
    End If
    On Error GoTo 0
End Sub

Private Sub btnOK_Click()
    If Not isPreviewMode Then
        '-------------------------------------------------------
        ' FIRST CLICK - Show Preview
        '-------------------------------------------------------
        Dim originalText As String
        originalText = txtPatternData.text
        
        If Trim(originalText) = "" Then
            MsgBox "Please paste your load pattern data.", vbExclamation
            txtPatternData.SetFocus
            Exit Sub
        End If
        
        ' Parse using MODULE function (shared logic)
        Set previewedPatterns = ParseLoadPatternText(originalText)

        
        If previewedPatterns Is Nothing Or previewedPatterns.count = 0 Then
            MsgBox "No valid load patterns found in the data." & vbCrLf & vbCrLf & _
                   "Please check the format:" & vbCrLf & _
                   "PatternName : Description" & vbCrLf & vbCrLf & _
                   "Example:" & vbCrLf & _
                   "BL : Dead Load" & vbCrLf & _
                   "LL : Live Load", vbExclamation, "No Valid Patterns"
            Exit Sub
        End If
        
        ' Store original text so user can go back to edit
        originalInputText = originalText
        
        ' Show preview in TextBox
        ShowPreviewInTextBox previewedPatterns
        
        ' Update UI
        isPreviewMode = True
        btnOK.Caption = "Create in SAP2000"
        btnCancel.Caption = "<<< Previous"
        Me.Caption = "Confirm Load Patterns (" & previewedPatterns.count & " found)"
        txtPatternData.BackColor = &HFFFFC0  ' Light yellow - indicates preview mode
        
        ' Enable overwrite option now that parsing succeeded
        On Error Resume Next
        chkOverwritePatterns.enabled = True
        ' keep delete disabled until overwrite is checked
        If chkOverwritePatterns.Value = True Then
            chkDeleteLoadcases.enabled = True
        Else
            chkDeleteLoadcases.enabled = False
            chkDeleteLoadcases.Value = False
        End If
        On Error GoTo 0
        
    Else
        '-------------------------------------------------------
        ' SECOND CLICK - Create Patterns in SAP2000
        '-------------------------------------------------------
        Dim finalText As String
        finalText = txtPatternData.text
        
        ' Check if it's table format (from preview) or original format
        If InStr(finalText, "PREVIEW") > 0 And InStr(finalText, "Pattern") > 0 Then
            ' Parse TABLE format (user may have edited the table)
            Set previewedPatterns = ParseTableFormat(finalText)
        Else
            ' Parse original format (user replaced everything)
            Set previewedPatterns = ParseLoadPatternText(finalText)
        End If
        
        If previewedPatterns Is Nothing Or previewedPatterns.count = 0 Then
            MsgBox "No valid load patterns found. Please check the data.", vbExclamation
            Exit Sub
        End If
        
        ' Read checkbox options (use default False if controls not present)
        Dim overwriteOpt As Boolean
        Dim deleteCasesOpt As Boolean
        On Error Resume Next
        overwriteOpt = CBool(chkOverwritePatterns.Value)
        deleteCasesOpt = CBool(chkDeleteLoadcases.Value)
        On Error GoTo 0
        
        ' Hide form
        Me.Hide
        
        ' Create patterns in SAP2000 using MODULE function
        Dim result As String
        result = CreatePatternsInSAP(previewedPatterns, overwriteOpt, deleteCasesOpt)
        
        ' Show result
        MsgBox result, vbInformation, "Load Pattern Creation Complete"
        
        ' Close form
        isFormOpen = False
        Unload Me
    End If
End Sub

Private Sub btnCancel_Click()
    If isPreviewMode Then
        ' Acts as "Previous" - restore original input for editing
        If Len(originalInputText) > 0 Then
            txtPatternData.text = originalInputText
        Else
            ' If for some reason original input is empty, just clear preview and allow editing
            txtPatternData.text = ""
        End If
        
        ' Restore UI to initial state
        isPreviewMode = False
        btnOK.Caption = "Preview >>>"
        btnCancel.Caption = "Cancel"
        Me.Caption = "Paste Load Patterns"
        txtPatternData.BackColor = vbWhite
        txtPatternData.SetFocus
        txtPatternData.SelStart = 0
        txtPatternData.BackColor = RGB(255, 255, 155)      ' yellow - indicates preview mode
        
        ' After returning to edit mode, disable overwrite/delete again until next successful parse
        On Error Resume Next
        chkOverwritePatterns.enabled = False
        chkOverwritePatterns.Value = False
        chkDeleteLoadcases.enabled = False
        chkDeleteLoadcases.Value = False
        On Error GoTo 0
    Else
        ' Normal Cancel behavior: close form
        isFormOpen = False
        Unload Me
    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' Handle X button click
    If CloseMode = vbFormControlMenu Then
        isFormOpen = False
    End If
End Sub

Private Sub ShowPreviewInTextBox(patterns As Object)
    ' Display formatted preview table in TextBox
    
    Dim previewText As String
    previewText = "-----------------------------------------------------------" & vbCrLf
    previewText = previewText & "  PREVIEW - " & patterns.count & " LOAD PATTERNS WILL BE CREATED" & vbCrLf
    previewText = previewText & "-----------------------------------------------------------" & vbCrLf & vbCrLf
    
    ' Table Header
    previewText = previewText & PadRight("Pattern", 15) & " " & _
                                 PadRight("Type", 25) & " " & _
                                 PadRight("SW Mult", 10) & vbCrLf
    previewText = previewText & "-----------------------------------------------------------" & vbCrLf
    
    ' Pattern data rows
    Dim key As Variant
    For Each key In patterns.keys
        Dim info As Variant
        info = patterns(key)
        
        Dim patternName As String
        Dim typeStr As String
        Dim swMult As String
        
        patternName = CStr(info(0))
        typeStr = CStr(info(2))
        swMult = Format(info(3), "0.0")
        
        ' Format row with fixed-width columns
        previewText = previewText & PadRight(patternName, 15) & " " & _
                                     PadRight(typeStr, 25) & " " & _
                                     PadRight(swMult, 10) & vbCrLf
    Next key
    
    previewText = previewText & "-----------------------------------------------------------" & vbCrLf & vbCrLf
    previewText = previewText & "You can edit Type or SW Mult. above if needed." & vbCrLf
    previewText = previewText & "Keep the table format when editing." & vbCrLf & vbCrLf
    previewText = previewText & "Options (enabled after Preview):" & vbCrLf
    previewText = previewText & " - Overwrite with new LoadPattern (will replace existing patterns with this set)" & vbCrLf
    previewText = previewText & " - Delete old Loadcase (except Modal) (only applies when Overwrite is enabled and checked)" & vbCrLf & vbCrLf
    previewText = previewText & "Click 'Create in SAP2000' to proceed or 'Previous' to go back and edit."
    
    ' Display in TextBox
    txtPatternData.text = previewText
    txtPatternData.SelStart = 0
End Sub

Private Function PadRight(text As String, Length As Integer) As String
    ' Pad string to fixed length for table formatting
    If Len(text) >= Length Then
        PadRight = Left(text, Length)
    Else
        PadRight = text & Space(Length - Len(text))
    End If
End Function

