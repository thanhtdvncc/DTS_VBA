Attribute VB_Name = "n05_ACAD_Openfrm"
Option Explicit

' Launcher for frmSyncCADSAP
' Shows the sync UserForm (modeless), checks AutoCAD and SAP2000 connections, writes status to txtStatus.

Public Sub ShowSyncForm()
    Dim f As Object, uf As Object, found As Boolean: found = False

    On Error Resume Next
    For Each uf In VBA.UserForms
        If uf.Name = "frmSyncCADSAP" Then
            Set f = uf
            found = True
            Exit For
        End If
    Next uf
    On Error GoTo 0

    If Not found Then
        Dim frm As frmSyncCADSAP
        Set frm = New frmSyncCADSAP
        Set f = frm
        f.Show vbModeless
        DoEvents
    Else
        On Error Resume Next
        If Not f.Visible Then f.Show vbModeless
        On Error GoTo 0
    End If

    On Error Resume Next
    f.txtStatus.text = "Initializing..." & vbCrLf & f.txtStatus.text
    f.txtStatus.BackColor = RGB(173, 216, 230)
    DoEvents
    On Error GoTo 0

    ' Check AutoCAD
    Dim acadApp As Object
    On Error Resume Next
    Set acadApp = GetObject(, "AutoCAD.Application")
    On Error GoTo 0

    If acadApp Is Nothing Then
        Dim startACAD As VbMsgBoxResult
        startACAD = MsgBox("AutoCAD is not running. Do you want to start AutoCAD now?", vbYesNo + vbQuestion, "Start AutoCAD?")
        If startACAD = vbYes Then
            On Error Resume Next
            Set acadApp = CreateObject("AutoCAD.Application")
            If Not acadApp Is Nothing Then
                acadApp.Visible = True
                f.txtStatus.text = "AutoCAD started." & vbCrLf & f.txtStatus.text
            Else
                f.txtStatus.text = "Failed to start AutoCAD." & vbCrLf & f.txtStatus.text
            End If
            On Error GoTo 0
        Else
            f.txtStatus.text = "AutoCAD not running. Some operations disabled." & vbCrLf & f.txtStatus.text
        End If
    Else
        f.txtStatus.text = "AutoCAD running." & vbCrLf & f.txtStatus.text
    End If
    DoEvents

    ' Connect SAP2000 via ConnectSAP2000 (module modSAP2000_Connection)
    Dim sapOk As Boolean: sapOk = False
    On Error Resume Next
    sapOk = ConnectSAP2000()
    If err.number <> 0 Then
        f.txtStatus.text = "ConnectSAP2000 error: " & err.number & " - " & err.description & vbCrLf & f.txtStatus.text
        err.Clear
    Else
        If sapOk Then
            f.txtStatus.text = "SAP2000 connection: SUCCESS." & vbCrLf & f.txtStatus.text
        Else
            f.txtStatus.text = "SAP2000 connection: FAILED." & vbCrLf & f.txtStatus.text
        End If
    End If
    On Error GoTo 0

    f.txtStatus.text = "Form ready." & vbCrLf & f.txtStatus.text
    f.Repaint
End Sub
