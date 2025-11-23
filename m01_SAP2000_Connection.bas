Attribute VB_Name = "m01_SAP2000_Connection"
Option Explicit
'===============================================================
' Module: modSAP2000_Connection
' Purpose: Connect / Disconnect to SAP2000 (late binding)
'===============================================================

Public Function ConnectSAP2000() As Boolean
    On Error GoTo FAIL
    
    Dim helperProgIDs As Variant
    helperProgIDs = Array( _
        "SAP2000v25.Helper", "SAP2000v24.Helper", "SAP2000v23.Helper", _
        "SAP2000v22.Helper", "SAP2000v21.Helper", "SAP2000v20.Helper", _
        "SAP2000v19.Helper", "SAP2000v1.Helper")
    
    Dim helperObj As Object, candidate As Object, pid As Variant, ret As Long
    
    ' Try helper objects first
    For Each pid In helperProgIDs
        On Error Resume Next
        Set helperObj = CreateObject(CStr(pid))
        On Error GoTo 0
        If Not helperObj Is Nothing Then
            On Error Resume Next
            Set candidate = helperObj.GetObject("CSI.SAP2000.API.SapObject")
            On Error GoTo 0
            If Not candidate Is Nothing Then
                Set SapApp = candidate
                Exit For
            End If
        End If
    Next pid
    
    ' Fallback: attach to existing instance
    If SapApp Is Nothing Then
        On Error Resume Next
        Set SapApp = GetObject(, "CSI.SAP2000.API.SapObject")
        On Error GoTo 0
    End If
    
    If SapApp Is Nothing Then
        LogMsg "ConnectSAP2000: Could not get a SapObject instance."
        ConnectSAP2000 = False
        Exit Function
    End If
    
    ret = SapApp.SetAsActiveObject
    If ret <> 0 Then LogMsg "SetAsActiveObject returned " & ret
    
    Set SapModel = SapApp.SapModel
    If SapModel Is Nothing Then
        LogMsg "ConnectSAP2000: SapModel is Nothing."
        ConnectSAP2000 = False
        Exit Function
    End If
    
    ret = SapModel.SetPresentUnits(UNIT_KN_MM_C)
    If ret <> 0 Then LogMsg "SetPresentUnits returned " & ret
    
    ConnectSAP2000 = True
    Exit Function
    
FAIL:
    LogMsg "ConnectSAP2000 failed: Err=" & err.number & " " & err.description
    ConnectSAP2000 = False
End Function

Public Sub DisconnectSAP2000(Optional showMsg As Boolean = False)
    On Error Resume Next
    Set SapModel = Nothing
    Set SapApp = Nothing
    If showMsg Then MsgBox "Disconnected from SAP2000.", vbInformation
End Sub

