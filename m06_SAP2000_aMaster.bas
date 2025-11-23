Attribute VB_Name = "m06_SAP2000_aMaster"
Option Explicit
'===============================================================
' Module: modSAP2000_Master
' Purpose: Orchestrate extraction using feature flags
'===============================================================

' Add this constant at the top of your module with other ENABLE_* flags:
' Public Const ENABLE_LOADCASES As Boolean = True

Public Sub SAP2000_GetData()
    On Error GoTo ERR_HANDLER
    
    Dim t0 As Double: t0 = Timer
    LogMsg "SAP2000_GetData: Starting..."
    PrepareExcelEnvironment False
    
    If Not ConnectSAP2000() Then
        MsgBox "Failed to connect to SAP2000.", vbCritical
        GoTo CleanUp
    End If
    
    ' Points
    LogMsg "Extracting joints..."
    ExtractPoints
    
    ' Frames
    LogMsg "Extracting frames..."
    ExtractFrames
    WriteGeometryData
    
    If ENABLE_FRAME_SECTIONS Then
        LogMsg "Writing frame sections..."
        WriteFrameSections
    Else
        LogMsg "Frame section extraction disabled."
    End If
    
    ' Areas
    If ENABLE_AREAS Then
        LogMsg "Extracting areas..."
        ExtractAreas
        WriteAreaData
        If ENABLE_AREA_SECTIONS Then
            LogMsg "Writing area sections..."
            WriteAreaSections
        Else
            LogMsg "Area section extraction disabled."
        End If
    Else
        LogMsg "Area extraction disabled."
    End If
    
    ' Groups
    If ENABLE_GROUPS Then
        LogMsg "Extracting groups..."
        WriteGroups
        LogMsg "Building group validation..."
        BuildGroupValidation
    Else
        LogMsg "Group extraction disabled."
    End If
    
    ' Load Cases & Combinations
    If ENABLE_LOADCASES Then
        LogMsg "Extracting load cases and combinations..."
        WriteLoadCases
        ' NEW: After writing load cases, build the pattern validation list automatically
        On Error Resume Next
        LogMsg "Building pattern validation list..."
        BuildPatternValidation
        If err.number <> 0 Then
            LogMsg "BuildPatternValidation error: " & err.number & " " & err.description
            err.Clear
        End If
        On Error GoTo 0
    Else
        LogMsg "Load case extraction disabled."
    End If
    
    
    ExportGridLinesToGirdlineSheet
    
    
    DisconnectSAP2000 False
    
    LogMsg "Finished in " & Format(Timer - t0, "0.0") & " s"
    MsgBox "SAP2000 data extraction completed in " & Format(Timer - t0, "0.0") & " s", vbInformation
    
CleanUp:
    PrepareExcelEnvironment True
    Exit Sub
    
ERR_HANDLER:
    LogMsg "ERROR [" & err.number & "] " & err.description
    MsgBox "Error [" & err.number & "] " & err.description, vbCritical, "SAP2000_GetData"
    Resume CleanUp
End Sub
