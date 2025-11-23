VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StatusCallbackHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'===============================================================
' Class: StatusCallbackHandler
' Purpose: Callback handler for status updates from Core modules
'===============================================================
Public Sub SetStatus(msg As String)
    On Error Resume Next
    
    ' Try to update UserForm status
    Dim uf As Object
    For Each uf In VBA.UserForms
        If uf.Name = "frmSyncCADSAP" Then
            uf.txtStatus.text = msg & vbCrLf & uf.txtStatus.text
            DoEvents
            Exit Sub
        End If
    Next uf
    
    ' Fallback to debug window if form not found
    Debug.Print Format(Now, "hh:nn:ss") & " | " & msg
    
    On Error GoTo 0
End Sub

