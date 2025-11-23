Attribute VB_Name = "m02_SAP2000_Environment"
Option Explicit
'===============================================================
' Module: modSAP2000_Environment
' Purpose: Manage Excel environment + basic logging
'===============================================================

Private oldCalc As XlCalculation
Private oldScreenUpdating As Boolean
Private oldEnableEvents As Boolean

Public Sub PrepareExcelEnvironment(ByVal Restore As Boolean)
    If Restore Then
        On Error Resume Next
        Application.ScreenUpdating = oldScreenUpdating
        Application.Calculation = oldCalc
        Application.EnableEvents = oldEnableEvents
    Else
        oldScreenUpdating = Application.ScreenUpdating
        oldCalc = Application.Calculation
        oldEnableEvents = Application.EnableEvents
        
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        Application.EnableEvents = False
    End If
End Sub

Public Sub LogMsg(ByVal msg As String)
    Debug.Print Format(Now, "hh:nn:ss") & " | " & msg
End Sub
