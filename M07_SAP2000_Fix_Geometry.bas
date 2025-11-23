Attribute VB_Name = "M07_SAP2000_Fix_Geometry"
Option Explicit
'===============================================================
' Module: modSAP2000_FixNodes
' Purpose: Force fix all floating point errors in node coordinates
' Note: Works in model's current units without changing them
'===============================================================

Public Sub FixNodeCoordinates()
    ' Connect to SAP2000
    If Not ConnectSAP2000() Then
        MsgBox "Could not connect to SAP2000. Please make sure SAP2000 is running with a model open.", vbExclamation
        Exit Sub
    End If
    
    ' Get current model units (READ ONLY - DO NOT CHANGE THEM)
    Dim modelUnits As Long
    modelUnits = SapModel.GetPresentUnits()
    
    ' Determine unit scale factor to meters for user input conversion
    Dim unitToMeter As Double
    Dim unitName As String
    Select Case modelUnits
        Case 1: unitToMeter = 0.0254: unitName = "lb_in_F"      ' inches
        Case 2: unitToMeter = 0.3048: unitName = "lb_ft_F"      ' feet
        Case 3: unitToMeter = 0.0254: unitName = "kip_in_F"     ' inches
        Case 4: unitToMeter = 0.3048: unitName = "kip_ft_F"     ' feet
        Case 5: unitToMeter = 0.001: unitName = "kN_mm_C"       ' millimeters
        Case 6: unitToMeter = 1#: unitName = "kN_m_C"           ' meters
        Case 7: unitToMeter = 0.001: unitName = "kgf_mm_C"      ' millimeters
        Case 8: unitToMeter = 1#: unitName = "kgf_m_C"          ' meters
        Case 9: unitToMeter = 0.001: unitName = "N_mm_C"        ' millimeters
        Case 10: unitToMeter = 1#: unitName = "N_m_C"           ' meters
        Case 11: unitToMeter = 0.001: unitName = "Ton_mm_C"     ' millimeters
        Case 12: unitToMeter = 1#: unitName = "Ton_m_C"         ' meters
        Case 13: unitToMeter = 0.01: unitName = "kN_cm_C"       ' centimeters
        Case 14: unitToMeter = 0.01: unitName = "kgf_cm_C"      ' centimeters
        Case 15: unitToMeter = 0.01: unitName = "N_cm_C"        ' centimeters
        Case 16: unitToMeter = 0.01: unitName = "Ton_cm_C"      ' centimeters
        Case Else
            MsgBox "Unknown unit system: " & modelUnits, vbExclamation
            DisconnectSAP2000
            Exit Sub
    End Select
    
    ' Ask user for rounding tolerance (in meters, will convert to model units)
    Dim tolerance As String
    tolerance = InputBox("Enter rounding tolerance in METERS:" & vbCrLf & _
                        "Default: 0.01 (1cm)" & vbCrLf & vbCrLf & _
                        "Examples:" & vbCrLf & _
                        "  0.001 = 1mm" & vbCrLf & _
                        "  0.01  = 1cm" & vbCrLf & _
                        "  0.1   = 10cm" & vbCrLf & vbCrLf & _
                        "Model units: " & unitName, _
                        "Fix Node Coordinates", "0.01")
    
    ' Validate input
    If tolerance = "" Then
        DisconnectSAP2000
        Exit Sub ' User cancelled
    End If
    
    Dim roundTolMeters As Double
    On Error Resume Next
    roundTolMeters = CDbl(tolerance)
    On Error GoTo 0
    
    If roundTolMeters <= 0 Or roundTolMeters > 1 Then
        MsgBox "Invalid tolerance value. Must be between 0 and 1 meters.", vbExclamation
        DisconnectSAP2000
        Exit Sub
    End If
    
    ' Convert tolerance to model units
    Dim roundTolModel As Double
    roundTolModel = roundTolMeters / unitToMeter
    
    ' Confirm action
    Dim confirmMsg As String
    confirmMsg = "This will FORCE ROUND ALL node coordinates." & vbCrLf & vbCrLf & _
                 "Model units: " & unitName & vbCrLf & _
                 "Tolerance: " & roundTolMeters & " m = " & roundTolModel & " model units" & vbCrLf & vbCrLf & _
                 "This will eliminate floating point errors (e-12, etc.)" & vbCrLf & _
                 "NO unit conversion will occur - working in current units only." & vbCrLf & vbCrLf & _
                 "Make sure you have saved your model before proceeding." & vbCrLf & vbCrLf & _
                 "Do you want to continue?"
    
    If MsgBox(confirmMsg, vbYesNo + vbQuestion, "Confirm Fix Node Coordinates") <> vbYes Then
        DisconnectSAP2000
        Exit Sub
    End If
    
    ' Get all point objects
    Dim ret As Long
    Dim numPoints As Long
    Dim pointNames() As String
    
    ret = SapModel.pointObj.GetNameList(numPoints, pointNames)
    
    If ret <> 0 Or numPoints = 0 Then
        MsgBox "Could not get point list from SAP2000." & vbCrLf & _
               "Return code: " & ret, vbExclamation
        DisconnectSAP2000
        Exit Sub
    End If
    
    ' Progress tracking
    Dim i As Long
    Dim X As Double, Y As Double, Z As Double
    Dim xNew As Double, yNew As Double, zNew As Double
    Dim changedCount As Long
    Dim precision As Double
    Dim startTime As Double
    
    startTime = Timer
    
    ' Calculate precision for rounding in MODEL UNITS
    precision = 1# / roundTolModel
    
    ' Disable screen updating for performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    LogMsg "FixNodeCoordinates: Starting - " & numPoints & " nodes, tolerance=" & roundTolModel & " " & unitName & " (FORCE ALL)"
    
    ' Process each point - FORCE round ALL coordinates IN MODEL UNITS
    For i = 0 To numPoints - 1
        ' Get current coordinates (in model's current units - NO conversion)
        ret = SapModel.pointObj.GetCoordCartesian(pointNames(i), X, Y, Z)
        
        If ret = 0 Then
            ' FORCE round ALL coordinates in MODEL UNITS to eliminate floating point errors
            xNew = Round(X * precision) / precision
            yNew = Round(Y * precision) / precision
            zNew = Round(Z * precision) / precision
            
            ' Update ALL nodes - write in SAME MODEL UNITS (no conversion occurs)
            ret = SapModel.EditPoint.ChangeCoordinates_1(pointNames(i), xNew, yNew, zNew, True)
            
            If ret = 0 Then
                changedCount = changedCount + 1
            Else
                LogMsg "  Warning: Could not update node " & pointNames(i) & " (ret=" & ret & ")"
            End If
        Else
            LogMsg "  Warning: Could not read coordinates for node " & pointNames(i) & " (ret=" & ret & ")"
        End If
        
        ' Show progress every 100 points
        If i Mod 100 = 0 Then
            Application.StatusBar = "Processing node " & (i + 1) & " of " & numPoints & "..."
        End If
    Next i
    
    ' Refresh model view once at the end
    ret = SapModel.View.RefreshView(0, False)
    
    ' Restore Excel settings
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
    
    ' Calculate elapsed time
    Dim elapsedTime As Double
    elapsedTime = Timer - startTime
    
    ' Show result
    Dim resultMsg As String
    resultMsg = "Force rounded ALL " & changedCount & " nodes (total: " & numPoints & ")." & vbCrLf & vbCrLf & _
                "Model units: " & unitName & vbCrLf & _
                "Tolerance: " & roundTolMeters & " m = " & roundTolModel & " model units" & vbCrLf & _
                "Eliminated floating point errors (e-12, etc.)" & vbCrLf & _
                "NO scaling occurred - worked in current units only." & vbCrLf & _
                "Time elapsed: " & Format(elapsedTime, "0.0") & " seconds"
    
    MsgBox resultMsg, vbInformation, "Fix Node Coordinates Complete"
    
    LogMsg "FixNodeCoordinates: Complete - Force rounded " & changedCount & " nodes in " & Format(elapsedTime, "0.0") & "s (" & unitName & ")"
    
    ' Disconnect
    DisconnectSAP2000
End Sub

