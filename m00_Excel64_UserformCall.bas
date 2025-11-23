Attribute VB_Name = "m00_Excel64_UserformCall"
'ShowFormWithWindowManagement myForm, topmost(0-1), String FormCaption, minimize excel (0-1)
'by thanhtdvncc

Sub show_frmWallConverter()
    Dim myForm As frmWallConverter
    Set myForm = New frmWallConverter
    ShowFormWithWindowManagement myForm, 0, "SAP2000 Wall Tool (by thanhtdvncc)", 0
End Sub


Sub show_frmSelection()
    Dim myForm As frmSelection
    Set myForm = New frmSelection
    ShowFormWithWindowManagement myForm, Range("FAOT").Value, "SAP2000 Selection Tool (by thanhtdvncc)", 1
End Sub


Sub show_frmColumnSection()
    ConnectSAP2000
    Dim myForm As frmColumnCrossSection
    Set myForm = New frmColumnCrossSection
    ShowFormWithWindowManagement myForm, 0, "SAP2000 Plan Colunm Tool (by thanhtdvncc)", 0
End Sub

Public Sub show_frmSyncCADSAP()

    Dim frm As frmSyncCADSAP
    Set frm = New frmSyncCADSAP

    On Error Resume Next
    frm.txtStatus.text = "Initializing..." & vbCrLf
    frm.txtStatus.BackColor = RGB(173, 216, 230)
    DoEvents
    On Error GoTo 0
    
    ' Check AutoCAD
    Dim acadApp As Object
    On Error Resume Next
    Set acadApp = GetObject(, "AutoCAD.Application")
    On Error GoTo 0
    
    If acadApp Is Nothing Then
        Dim startACAD As VbMsgBoxResult
        startACAD = MsgBox("AutoCAD is not running. Do you want to start AutoCAD now?", _
                           vbYesNo + vbQuestion, "Start AutoCAD?")
        If startACAD = vbYes Then
            On Error Resume Next
            Set acadApp = CreateObject("AutoCAD.Application")
            If Not acadApp Is Nothing Then
                acadApp.Visible = True
                frm.txtStatus.text = "AutoCAD started." & vbCrLf & frm.txtStatus.text
            Else
                frm.txtStatus.text = "Failed to start AutoCAD." & vbCrLf & frm.txtStatus.text
            End If
            On Error GoTo 0
        Else
            frm.txtStatus.text = "AutoCAD not running. Some operations disabled." & vbCrLf & frm.txtStatus.text
        End If
    Else
        frm.txtStatus.text = "AutoCAD running." & vbCrLf & frm.txtStatus.text
    End If
    DoEvents
    
    ' Connect SAP2000
    Dim sapOk As Boolean
    On Error Resume Next
    sapOk = ConnectSAP2000()
    If err.number <> 0 Then
        frm.txtStatus.text = "ConnectSAP2000 error: " & err.number & " - " & err.description & vbCrLf & frm.txtStatus.text
        err.Clear
    Else
        If sapOk Then
            frm.txtStatus.text = "SAP2000 connection: SUCCESS." & vbCrLf & frm.txtStatus.text
        Else
            frm.txtStatus.text = "SAP2000 connection: FAILED." & vbCrLf & frm.txtStatus.text
        End If
    End If
    On Error GoTo 0
    
    frm.txtStatus.text = "Form ready." & vbCrLf & frm.txtStatus.text
    
    ShowManagedForm frm, "SAP2000 Model from AutoCAD (by thanhtdvncc)", 1, 1

End Sub

