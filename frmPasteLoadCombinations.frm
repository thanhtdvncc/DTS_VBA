VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPasteLoadCombinations 
   Caption         =   "Load Combo Data Input (by thanhtdvncc):"
   ClientHeight    =   4396
   ClientLeft      =   120
   ClientTop       =   675
   ClientWidth     =   7230
   OleObjectBlob   =   "frmPasteLoadCombinations.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPasteLoadCombinations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'===============================================================
' UserForm: frmPasteLoadCombinations
' Purpose: Preview and edit load combinations before creating in SAP2000
'===============================================================

Private previewedCombos As Object
Private isPreviewMode As Boolean
Private originalInputText As String

Private Sub UserForm_Initialize()
    ' Set example text
    txtComboData.BackColor = RGB(255, 255, 155)  ' yellow
    txtComboData.text = _
    "# Paste or type your load combinations below" & vbCrLf & _
    "# Format: ComboName   Formula" & vbCrLf & _
    "# Example combinations:" & vbCrLf & vbCrLf & _
    "LC1001   1.4BL+1.4DL+1.4CL+1.4EE+1.4PE" & vbCrLf & _
    "LC1002   1.2BL+1.2DL+1.2CL+1.2EE+1.2PE+1.6LL+0.5Lr" & vbCrLf & _
    "LC1003   1.2BL+1.2DL+1.2CL+1.2EE+1.2PE+1.0LL+1.6Lr" & vbCrLf & _
    "LC1004   1.2BL+1.2DL+1.2CL+1.2EE+1.2PE+1.6Lr+0.8WXP" & vbCrLf & _
    "LC1005   1.2BL+1.2DL+1.2CL+1.2EE+1.2PE+1.0LL+1.0SXS+0.3SZS" & vbCrLf & vbCrLf & _
    "Tip: Copy your combo list from Excel or Word and paste it here directly."

    txtComboData.SetFocus
    txtComboData.SelStart = 0
    'txtComboData.SelLength = Len(txtComboData.Text)

    isPreviewMode = False
    btnOK.Caption = "Preview >>>"

    originalInputText = ""
    btnCancel.Caption = "Cancel"

    ' Initialize overwrite checkbox
    On Error Resume Next
    With chkOverwriteCombos
        .Visible = True
        .enabled = False
        .Caption = "Overwrite existing Load Combinations"
        .Value = False
    End With
    On Error GoTo 0
End Sub

Private Sub chkOverwriteCombos_Click()
    ' Optional: Add warning if overwrite is checked
    On Error Resume Next
    If chkOverwriteCombos.Value = True Then
        ' Could show warning message
    End If
    On Error GoTo 0
End Sub

Private Sub btnOK_Click()
    If Not isPreviewMode Then
        ' FIRST CLICK - Show Preview
        Dim originalText As String
        originalText = txtComboData.text

        If Trim(originalText) = "" Then
            MsgBox "Please paste your load combination data.", vbExclamation
            txtComboData.SetFocus
            Exit Sub
        End If

        ' Parse using MODULE function
        Set previewedCombos = ParseLoadComboText(originalText)

        If previewedCombos Is Nothing Or previewedCombos.count = 0 Then
            MsgBox "No valid load combinations found in the data." & vbCrLf & vbCrLf & _
                   "Please check the format:" & vbCrLf & _
                   "ComboName   Formula" & vbCrLf & vbCrLf & _
                   "Example:" & vbCrLf & _
                   "LC1001   1.4BL+1.4DL+1.6LL" & vbCrLf & _
                   "LC1002   1.2BL+1.2DL+1.0LL+1.0SX", vbExclamation, "No Valid Combinations"
            Exit Sub
        End If

        ' Store original text
        originalInputText = originalText

        ' Show preview
        ShowPreviewInTextBox previewedCombos

        ' Update UI
        isPreviewMode = True
        btnOK.Caption = "Create in SAP2000"
        btnCancel.Caption = "<<< Previous"
        Me.Caption = "Confirm Load Combinations (" & previewedCombos.count & " found)"
        txtComboData.BackColor = &HFFFFC0  ' Light yellow

        ' Enable overwrite option
        On Error Resume Next
        chkOverwriteCombos.enabled = True
        On Error GoTo 0

    Else
        ' SECOND CLICK - Create Combos in SAP2000
        Dim finalText As String
        finalText = txtComboData.text

        ' Check format
        If InStr(finalText, "PREVIEW") > 0 And InStr(finalText, "Combo") > 0 Then
            ' Parse TABLE format
            Set previewedCombos = ParseComboTableFormat(finalText)
        Else
            ' Parse original format
            Set previewedCombos = ParseLoadComboText(finalText)
        End If

        If previewedCombos Is Nothing Or previewedCombos.count = 0 Then
            MsgBox "No valid load combinations found. Please check the data.", vbExclamation
            Exit Sub
        End If

        ' Read checkbox option
        Dim overwriteOpt As Boolean
        On Error Resume Next
        overwriteOpt = CBool(chkOverwriteCombos.Value)
        On Error GoTo 0

        ' Hide form
        Me.Hide

        ' Create combos in SAP2000
        Dim result As String
        result = CreateCombosInSAP(previewedCombos, overwriteOpt)

        ' Show result
        MsgBox result, vbInformation, "Load Combination Creation Complete"

        ' Close form
        isFormOpen_Combos = False
        Unload Me
    End If
End Sub

Private Sub btnCancel_Click()
    If isPreviewMode Then
        ' Acts as "Previous" - restore original input
        If Len(originalInputText) > 0 Then
            txtComboData.text = originalInputText
        Else
            txtComboData.text = ""
        End If

        ' Restore UI to initial state
        isPreviewMode = False
        btnOK.Caption = "Preview >>>"
        btnCancel.Caption = "Cancel"
        Me.Caption = "Paste Load Combinations"
        txtComboData.BackColor = vbWhite
        txtComboData.SetFocus
        txtComboData.SelStart = 0
        txtComboData.BackColor = RGB(255, 255, 155)  ' yellow

        ' Disable overwrite checkbox
        On Error Resume Next
        chkOverwriteCombos.enabled = False
        chkOverwriteCombos.Value = False
        On Error GoTo 0
    Else
        ' Normal Cancel: close form
        isFormOpen_Combos = False
        Unload Me
    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' Handle X button click
    If CloseMode = vbFormControlMenu Then
        isFormOpen_Combos = False
    End If
End Sub

Private Sub ShowPreviewInTextBox(combos As Object)
    ' Display formatted preview table

    Dim previewText As String
    previewText = "-----------------------------------------------------------" & vbCrLf
    previewText = previewText & "  PREVIEW - " & combos.count & " LOAD COMBINATIONS WILL BE CREATED" & vbCrLf
    previewText = previewText & "-----------------------------------------------------------" & vbCrLf & vbCrLf

    ' Table Header
    previewText = previewText & PadRight("Combo Name", 20) & " " & _
                                 PadRight("Type", 20) & " " & _
                                 "Formula" & vbCrLf
    previewText = previewText & "-----------------------------------------------------------" & vbCrLf

    ' Combo data rows
    Dim key As Variant
    For Each key In combos.keys
        Dim info As Variant
        info = combos(key)

        Dim comboName As String
        Dim typeStr As String
        Dim formulaStr As String

        comboName = CStr(info(0))
        formulaStr = CStr(info(1))
        typeStr = GetComboTypeString(CLng(info(2)))

        ' Format row
        previewText = previewText & PadRight(comboName, 20) & " " & _
                                     PadRight(typeStr, 20) & " " & _
                                     formulaStr & vbCrLf
    Next key

    previewText = previewText & "-----------------------------------------------------------" & vbCrLf & vbCrLf
    previewText = previewText & "You can edit Type or Formula above if needed." & vbCrLf
    previewText = previewText & "Keep the table format when editing." & vbCrLf & vbCrLf
    previewText = previewText & "Options (enabled after Preview):" & vbCrLf
    previewText = previewText & " - Overwrite existing Load Combinations (will delete all existing combos first)" & vbCrLf & vbCrLf
    previewText = previewText & "Click 'Create in SAP2000' to proceed or 'Previous' to go back and edit."

    ' Display
    txtComboData.text = previewText
    txtComboData.SelStart = 0
End Sub

Private Function PadRight(text As String, Length As Integer) As String
    ' Pad string to fixed length for table formatting
    If Len(text) >= Length Then
        PadRight = Left(text, Length)
    Else
        PadRight = text & Space(Length - Len(text))
    End If
End Function


