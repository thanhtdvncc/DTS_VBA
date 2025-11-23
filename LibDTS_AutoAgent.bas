Attribute VB_Name = "LibDTS_AutoAgent"
' ===============================================================
' Module: LibDTS_AutoAgent
' Purpose: Automated engineering agent for DTS_VBA repository consolidation
' Features: CLI-style commands, token budget tracking, dry-run support
' Version: 1.0
' ===============================================================
Option Explicit

' ==========================================
' MODULE-LEVEL STATE
' ==========================================
Private m_TokenBudget As Long           ' Total tokens permitted
Private m_TokensUsed As Long            ' Tokens consumed so far
Private m_DryRun As Boolean             ' Validation-only mode
Private m_UserApprovalRequired As Boolean ' Require confirmation before writes
Private m_LastError As String           ' Last error message
Private m_InventoryCache As Object      ' Cached inventory data

' ==========================================
' PUBLIC CONSTANTS
' ==========================================
Public Const AGENT_NAME As String = "LibDTS_AutoAgent"
Public Const DEFAULT_TOKEN_BUDGET As Long = 20000
Public Const REPO_NAME As String = "thanhtdvncc/DTS_VBA"

' Status codes
Public Enum AgentStatus
    STATUS_READY = 0
    STATUS_RUNNING = 1
    STATUS_COMPLETED = 2
    STATUS_ERROR = 3
    STATUS_TOKEN_EXHAUSTED = 4
End Enum

' ==========================================
' INITIALIZATION
' ==========================================

' Initialize agent with startup parameters
' Parameters:
'   tokenBudget: Total tokens allowed (default: 20000)
'   dryRun: Validation-only mode (default: True)
'   userApprovalRequired: Require confirmation (default: True)
Public Sub Initialize(Optional ByVal tokenBudget As Long = 0, _
                     Optional ByVal dryRun As Boolean = True, _
                     Optional ByVal userApprovalRequired As Boolean = True)
    On Error GoTo ErrHandler
    
    ' Set parameters
    If tokenBudget <= 0 Then
        m_TokenBudget = DEFAULT_TOKEN_BUDGET
    Else
        m_TokenBudget = tokenBudget
    End If
    
    m_DryRun = dryRun
    m_UserApprovalRequired = userApprovalRequired
    m_TokensUsed = 0
    
    ' Initialize cache
    Set m_InventoryCache = CreateObject("Scripting.Dictionary")
    
    ' Log initialization
    LibDTS_Logger.Log AGENT_NAME & ": Initialized with tokenBudget=" & m_TokenBudget & _
                      ", dryRun=" & m_DryRun & ", userApprovalRequired=" & m_UserApprovalRequired, DTS_INFO
    
    ' Report budget
    ReportTokens
    Exit Sub
    
ErrHandler:
    m_LastError = "Initialize error: " & err.description
    LibDTS_Logger.Log AGENT_NAME & ": " & m_LastError, DTS_ERROR
End Sub

' Report remaining token budget
Public Sub ReportTokens()
    Dim remaining As Long
    remaining = m_TokenBudget - m_TokensUsed
    
    Debug.Print "Tokens remaining: " & remaining
    
    If remaining <= 0 Then
        Debug.Print "token exhausted"
        LibDTS_Logger.Log AGENT_NAME & ": Token budget exhausted", DTS_WARNING
    End If
End Sub

' Check if sufficient tokens available for operation
' Parameters:
'   estimatedCost: Estimated tokens needed
' Returns: True if sufficient budget
Private Function CheckTokenBudget(estimatedCost As Long) As Boolean
    Dim remaining As Long
    remaining = m_TokenBudget - m_TokensUsed
    
    If estimatedCost > remaining Then
        Debug.Print "Insufficient tokens. Required: " & estimatedCost & ", Available: " & remaining
        Debug.Print "Propose splitting into sub-tasks. Confirm to proceed? (Y/N)"
        CheckTokenBudget = False
    Else
        CheckTokenBudget = True
    End If
End Function

' Consume tokens and update budget
Private Sub ConsumeTokens(cost As Long)
    m_TokensUsed = m_TokensUsed + cost
    ReportTokens
End Sub

' ==========================================
' VERIFICATION
' ==========================================

' Verify JsonConverter.bas exists (critical dependency)
' Returns: True if found
Public Function VerifyJsonConverter() As Boolean
    On Error GoTo ErrHandler
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim repoPath As String
    repoPath = GetRepoPath()
    
    Dim jsonPath As String
    jsonPath = repoPath & "\JsonConverter.bas"
    
    If fso.FileExists(jsonPath) Then
        LibDTS_Logger.Log AGENT_NAME & ": JsonConverter.bas found at " & jsonPath, DTS_INFO
        Debug.Print "[OK] JsonConverter.bas exists"
        VerifyJsonConverter = True
    Else
        m_LastError = "CRITICAL: JsonConverter.bas not found. Cannot proceed."
        LibDTS_Logger.Log AGENT_NAME & ": " & m_LastError, DTS_ERROR
        Debug.Print "[ERROR] " & m_LastError
        VerifyJsonConverter = False
    End If
    
    ConsumeTokens 50
    Exit Function
    
ErrHandler:
    m_LastError = "VerifyJsonConverter error: " & err.description
    LibDTS_Logger.Log AGENT_NAME & ": " & m_LastError, DTS_ERROR
    VerifyJsonConverter = False
End Function

' Scan for Core_Utils.bas and identify SAP helper APIs
' Returns: Dictionary with helper function names and info
Public Function ScanCoreUtils() As Object
    On Error GoTo ErrHandler
    
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    
    Dim fso As Object, ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim repoPath As String
    repoPath = GetRepoPath()
    
    Dim coreUtilsPath As String
    coreUtilsPath = repoPath & "\Core_Utils.bas"
    
    If Not fso.FileExists(coreUtilsPath) Then
        LibDTS_Logger.Log AGENT_NAME & ".ScanCoreUtils: Core_Utils.bas not found", DTS_WARNING
        Set ScanCoreUtils = result
        ConsumeTokens 50
        Exit Function
    End If
    
    ' Read file and parse for public functions
    Set ts = fso.OpenTextFile(coreUtilsPath, 1)
    Dim content As String
    content = ts.ReadAll
    ts.Close
    
    ' Simple parsing for public functions/subs
    Dim lines() As String
    lines = Split(content, vbCrLf)
    
    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        Dim line As String
        line = Trim$(lines(i))
        
        ' Look for Public Function/Sub declarations
        If InStr(1, line, "Public Function ", vbTextCompare) > 0 Or _
           InStr(1, line, "Public Sub ", vbTextCompare) > 0 Then
            
            ' Extract function name
            Dim funcName As String
            funcName = ExtractFunctionName(line)
            
            If Len(funcName) > 0 Then
                result.Add funcName, "Line " & (i + 1) & ": " & Left$(line, 80)
            End If
        End If
    Next i
    
    LibDTS_Logger.Log AGENT_NAME & ".ScanCoreUtils: Found " & result.Count & " public APIs", DTS_INFO
    Debug.Print "[INFO] Core_Utils.bas contains " & result.Count & " public APIs"
    
    Set ScanCoreUtils = result
    ConsumeTokens 200
    Exit Function
    
ErrHandler:
    m_LastError = "ScanCoreUtils error: " & err.description
    LibDTS_Logger.Log AGENT_NAME & ": " & m_LastError, DTS_ERROR
    Set ScanCoreUtils = CreateObject("Scripting.Dictionary")
End Function

' Extract function name from declaration line
Private Function ExtractFunctionName(line As String) As String
    On Error Resume Next
    
    Dim startPos As Long
    Dim endPos As Long
    
    ' Find "Function" or "Sub"
    If InStr(1, line, "Function ", vbTextCompare) > 0 Then
        startPos = InStr(1, line, "Function ", vbTextCompare) + 9
    ElseIf InStr(1, line, "Sub ", vbTextCompare) > 0 Then
        startPos = InStr(1, line, "Sub ", vbTextCompare) + 4
    Else
        ExtractFunctionName = ""
        Exit Function
    End If
    
    ' Find opening parenthesis
    endPos = InStr(startPos, line, "(")
    If endPos = 0 Then
        ' No parameters, look for "As" or end of line
        endPos = InStr(startPos, line, " As ")
        If endPos = 0 Then endPos = Len(line) + 1
    End If
    
    ExtractFunctionName = Trim$(Mid$(line, startPos, endPos - startPos))
End Function

' ==========================================
' INVENTORY OPERATIONS
' ==========================================

' Scan repository and generate inventory
' Parameters:
'   outputFormat: "CSV", "MD" (Markdown), or "DICT" (Dictionary object)
' Returns: Inventory data in requested format
Public Function ScanRepository(Optional outputFormat As String = "MD") As Variant
    On Error GoTo ErrHandler
    
    If Not CheckTokenBudget(2000) Then
        ScanRepository = ""
        Exit Function
    End If
    
    LibDTS_Logger.Log AGENT_NAME & ".ScanRepository: Starting inventory scan", DTS_INFO
    Debug.Print "[SCAN] Starting repository inventory..."
    
    ' Get repository path
    Dim repoPath As String
    repoPath = GetRepoPath()
    
    ' Scan for VBA files
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim files As Object
    Set files = CreateObject("Scripting.Dictionary")
    
    ' Collect all .bas and .cls files
    CollectVBAFiles repoPath, files, fso
    
    LibDTS_Logger.Log AGENT_NAME & ".ScanRepository: Found " & files.Count & " VBA files", DTS_INFO
    Debug.Print "[SCAN] Found " & files.Count & " VBA files"
    
    ' Parse each file
    Dim inventory As Object
    Set inventory = CreateObject("Scripting.Dictionary")
    
    Dim filePath As Variant
    For Each filePath In files.Keys
        Dim fileInfo As Object
        Set fileInfo = ParseVBAFile(CStr(filePath), fso)
        inventory.Add filePath, fileInfo
    Next filePath
    
    ' Cache for later use
    Set m_InventoryCache = inventory
    
    ' Format output
    Select Case UCase$(outputFormat)
        Case "CSV"
            ScanRepository = FormatInventoryCSV(inventory)
        Case "MD"
            ScanRepository = FormatInventoryMD(inventory)
        Case "DICT"
            Set ScanRepository = inventory
        Case Else
            ScanRepository = FormatInventoryMD(inventory)
    End Select
    
    ConsumeTokens 2000
    
    LibDTS_Logger.Log AGENT_NAME & ".ScanRepository: Inventory complete", DTS_INFO
    Debug.Print "[SCAN] Inventory complete"
    
    Exit Function
    
ErrHandler:
    m_LastError = "ScanRepository error: " & err.description
    LibDTS_Logger.Log AGENT_NAME & ": " & m_LastError, DTS_ERROR
    ScanRepository = ""
End Function

' Recursively collect VBA files
Private Sub CollectVBAFiles(folderPath As String, files As Object, fso As Object)
    On Error Resume Next
    
    Dim folder As Object
    Set folder = fso.GetFolder(folderPath)
    
    If folder Is Nothing Then Exit Sub
    
    ' Process files in current folder
    Dim file As Object
    For Each file In folder.files
        If LCase$(fso.GetExtensionName(file.Path)) = "bas" Or _
           LCase$(fso.GetExtensionName(file.Path)) = "cls" Then
            files.Add file.Path, file.Name
        End If
    Next file
    
    ' Recurse into subfolders (skip .git)
    Dim subfolder As Object
    For Each subfolder In folder.SubFolders
        If subfolder.Name <> ".git" Then
            CollectVBAFiles subfolder.Path, files, fso
        End If
    Next subfolder
End Sub

' Parse VBA file and extract public APIs
Private Function ParseVBAFile(filePath As String, fso As Object) As Object
    On Error GoTo ErrHandler
    
    Dim info As Object
    Set info = CreateObject("Scripting.Dictionary")
    info.Add "path", filePath
    info.Add "name", fso.GetFileName(filePath)
    
    Dim publicAPIs As Object
    Set publicAPIs = CreateObject("Scripting.Dictionary")
    
    ' Read file
    Dim ts As Object
    Set ts = fso.OpenTextFile(filePath, 1)
    Dim lines() As String
    lines = Split(ts.ReadAll, vbCrLf)
    ts.Close
    
    ' Parse for public APIs
    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        Dim line As String
        line = Trim$(lines(i))
        
        If InStr(1, line, "Public Function ", vbTextCompare) > 0 Or _
           InStr(1, line, "Public Sub ", vbTextCompare) > 0 Then
            
            Dim apiName As String
            apiName = ExtractFunctionName(line)
            
            If Len(apiName) > 0 And Not publicAPIs.Exists(apiName) Then
                publicAPIs.Add apiName, i + 1 ' Line number
            End If
        End If
    Next i
    
    info.Add "publicAPIs", publicAPIs
    info.Add "apiCount", publicAPIs.Count
    
    Set ParseVBAFile = info
    Exit Function
    
ErrHandler:
    Set ParseVBAFile = CreateObject("Scripting.Dictionary")
End Function

' Format inventory as Markdown
Private Function FormatInventoryMD(inventory As Object) As String
    Dim output As String
    output = "# DTS_VBA Repository Inventory" & vbCrLf & vbCrLf
    output = output & "**Generated:** " & Now & vbCrLf & vbCrLf
    output = output & "## Summary" & vbCrLf & vbCrLf
    output = output & "- Total Files: " & inventory.Count & vbCrLf
    
    Dim totalAPIs As Long
    totalAPIs = 0
    
    Dim key As Variant
    For Each key In inventory.Keys
        Dim fileInfo As Object
        Set fileInfo = inventory(key)
        totalAPIs = totalAPIs + fileInfo("apiCount")
    Next key
    
    output = output & "- Total Public APIs: " & totalAPIs & vbCrLf & vbCrLf
    output = output & "## Files" & vbCrLf & vbCrLf
    
    ' List each file
    For Each key In inventory.Keys
        Set fileInfo = inventory(key)
        output = output & "### " & fileInfo("name") & vbCrLf & vbCrLf
        output = output & "**Path:** " & fileInfo("path") & vbCrLf
        output = output & "**Public APIs:** " & fileInfo("apiCount") & vbCrLf & vbCrLf
        
        If fileInfo("apiCount") > 0 Then
            Dim apis As Object
            Set apis = fileInfo("publicAPIs")
            
            Dim apiName As Variant
            For Each apiName In apis.Keys
                output = output & "- `" & apiName & "` (Line " & apis(apiName) & ")" & vbCrLf
            Next apiName
            output = output & vbCrLf
        End If
    Next key
    
    FormatInventoryMD = output
End Function

' Format inventory as CSV
Private Function FormatInventoryCSV(inventory As Object) As String
    Dim output As String
    output = "File,Path,API_Name,Line_Number" & vbCrLf
    
    Dim key As Variant
    For Each key In inventory.Keys
        Dim fileInfo As Object
        Set fileInfo = inventory(key)
        
        Dim fileName As String
        fileName = fileInfo("name")
        
        Dim filePath As String
        filePath = fileInfo("path")
        
        If fileInfo("apiCount") > 0 Then
            Dim apis As Object
            Set apis = fileInfo("publicAPIs")
            
            Dim apiName As Variant
            For Each apiName In apis.Keys
                output = output & fileName & "," & filePath & "," & apiName & "," & apis(apiName) & vbCrLf
            Next apiName
        Else
            output = output & fileName & "," & filePath & ",(none),0" & vbCrLf
        End If
    Next key
    
    FormatInventoryCSV = output
End Function

' ==========================================
' DRIVER INTEGRATION ANALYSIS
' ==========================================

' Check if LibDTS_DriverSAP reuses Core_Utils helpers
' Returns: Dictionary with analysis results
Public Function AnalyzeDriverSAPIntegration() As Object
    On Error GoTo ErrHandler
    
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    
    ' Scan Core_Utils for SAP helpers
    Dim coreHelpers As Object
    Set coreHelpers = ScanCoreUtils()
    
    ' Read LibDTS_DriverSAP
    Dim fso As Object, ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim driverPath As String
    driverPath = GetRepoPath() & "\LibDTS_DriverSAP.bas"
    
    If Not fso.FileExists(driverPath) Then
        result.Add "error", "LibDTS_DriverSAP.bas not found"
        Set AnalyzeDriverSAPIntegration = result
        Exit Function
    End If
    
    Set ts = fso.OpenTextFile(driverPath, 1)
    Dim driverContent As String
    driverContent = ts.ReadAll
    ts.Close
    
    ' Check for references to Core_Utils functions
    Dim reusedFunctions As Object
    Set reusedFunctions = CreateObject("Scripting.Dictionary")
    
    Dim missingFunctions As Object
    Set missingFunctions = CreateObject("Scripting.Dictionary")
    
    Dim helperName As Variant
    For Each helperName In coreHelpers.Keys
        If InStr(1, driverContent, CStr(helperName), vbTextCompare) > 0 Then
            reusedFunctions.Add helperName, "Referenced"
        Else
            ' Check if this is a SAP-related helper
            If IsSAPRelatedHelper(CStr(helperName)) Then
                missingFunctions.Add helperName, "Not referenced"
            End If
        End If
    Next helperName
    
    result.Add "coreHelpersCount", coreHelpers.Count
    result.Add "reusedCount", reusedFunctions.Count
    result.Add "missingCount", missingFunctions.Count
    result.Add "reusedFunctions", reusedFunctions
    result.Add "missingFunctions", missingFunctions
    
    ' Generate report
    Debug.Print "[ANALYSIS] Driver-Helper Integration:"
    Debug.Print "  Core_Utils helpers: " & coreHelpers.Count
    Debug.Print "  Reused in DriverSAP: " & reusedFunctions.Count
    Debug.Print "  Missing SAP helpers: " & missingFunctions.Count
    
    Set AnalyzeDriverSAPIntegration = result
    ConsumeTokens 300
    Exit Function
    
ErrHandler:
    m_LastError = "AnalyzeDriverSAPIntegration error: " & err.description
    LibDTS_Logger.Log AGENT_NAME & ": " & m_LastError, DTS_ERROR
    Set AnalyzeDriverSAPIntegration = CreateObject("Scripting.Dictionary")
End Function

' Check if function name is SAP-related
Private Function IsSAPRelatedHelper(funcName As String) As Boolean
    Dim lowerName As String
    lowerName = LCase$(funcName)
    
    IsSAPRelatedHelper = (InStr(lowerName, "sap") > 0) Or _
                        (InStr(lowerName, "connect") > 0) Or _
                        (InStr(lowerName, "model") > 0) Or _
                        (InStr(lowerName, "element") > 0)
End Function

' ==========================================
' UTILITY FUNCTIONS
' ==========================================

' Get repository root path
Private Function GetRepoPath() As String
    ' This should be the repository root
    ' In actual VBA environment, this would need to be configured
    GetRepoPath = ThisWorkbook.Path
    
    ' Fallback to current directory
    If Len(GetRepoPath) = 0 Then
        GetRepoPath = CurDir$
    End If
End Function

' Get last error message
Public Function GetLastError() As String
    GetLastError = m_LastError
    m_LastError = ""
End Function

' ==========================================
' DELIVERABLE GENERATION
' ==========================================

' Generate adapter mapping table (legacy -> driver)
' Returns: CSV string with mappings
Public Function GenerateAdapterMappingCSV() As String
    On Error GoTo ErrHandler
    
    If Not CheckTokenBudget(500) Then
        GenerateAdapterMappingCSV = ""
        Exit Function
    End If
    
    Dim output As String
    output = "Legacy_Module,Legacy_Function,Driver_Module,Driver_Function,Notes" & vbCrLf
    
    ' SAP Connection mappings
    output = output & "m01_SAP2000_Connection,ConnectSAP2000,LibDTS_DriverSAP,Connect,Enhanced version detection" & vbCrLf
    output = output & "m01_SAP2000_Connection,DisconnectSAP2000,LibDTS_DriverSAP,Disconnect,Simplified" & vbCrLf
    output = output & "Core_Utils,IsSAPConnected,LibDTS_DriverSAP,IsConnected,Direct replacement" & vbCrLf
    
    ' CAD Read mappings
    output = output & "Core_XData_Reader,ReadPointsFromCAD,LibDTS_DriverCAD,ReadAllPoints,Batch operation" & vbCrLf
    output = output & "Core_XData_Reader,ReadFramesFromCAD,LibDTS_DriverCAD,ReadAllFrames,Batch operation" & vbCrLf
    output = output & "Core_XData_Reader,ReadAreasFromCAD,LibDTS_DriverCAD,ReadAllAreas,Batch operation" & vbCrLf
    
    ' SAP Push mappings
    output = output & "m04_SAP2000_Joints_Frames,ExtractPoints,LibDTS_DriverSAP,ReadPoint,Individual read" & vbCrLf
    output = output & "m04_SAP2000_Joints_Frames,ExtractFrames,LibDTS_DriverSAP,ReadFrame,Individual read" & vbCrLf
    output = output & "m05_SAP2000_Areas,ExtractAreas,LibDTS_DriverSAP,ReadArea,Individual read" & vbCrLf
    
    GenerateAdapterMappingCSV = output
    ConsumeTokens 500
    Exit Function
    
ErrHandler:
    m_LastError = "GenerateAdapterMappingCSV error: " & err.description
    GenerateAdapterMappingCSV = ""
End Function

' ==========================================
' COMMAND INTERFACE
' ==========================================

' Execute agent command
' Parameters:
'   command: Command string (e.g., "scan repo and inventory")
' Returns: Result string or object
Public Function ExecuteCommand(command As String) As Variant
    On Error GoTo ErrHandler
    
    Dim cmd As String
    cmd = LCase$(Trim$(command))
    
    LibDTS_Logger.Log AGENT_NAME & ".ExecuteCommand: " & command, DTS_INFO
    Debug.Print "[CMD] Executing: " & command
    
    ' Parse and execute command
    Select Case True
        Case InStr(cmd, "scan repo") > 0 Or InStr(cmd, "inventory") > 0
            ExecuteCommand = ScanRepository("MD")
            
        Case InStr(cmd, "verify json") > 0
            ExecuteCommand = VerifyJsonConverter()
            
        Case InStr(cmd, "scan core_utils") > 0
            Set ExecuteCommand = ScanCoreUtils()
            
        Case InStr(cmd, "analyze driver") > 0
            Set ExecuteCommand = AnalyzeDriverSAPIntegration()
            
        Case InStr(cmd, "generate adapter") > 0 Or InStr(cmd, "mapping csv") > 0
            ExecuteCommand = GenerateAdapterMappingCSV()
            
        Case InStr(cmd, "help") > 0
            ExecuteCommand = GetHelpText()
            
        Case Else
            ExecuteCommand = "Unknown command: " & command & vbCrLf & "Type 'help' for available commands."
    End Select
    
    Exit Function
    
ErrHandler:
    m_LastError = "ExecuteCommand error: " & err.description
    LibDTS_Logger.Log AGENT_NAME & ": " & m_LastError, DTS_ERROR
    ExecuteCommand = "Error: " & m_LastError
End Function

' Get help text
Private Function GetHelpText() As String
    Dim help As String
    help = "DTS_VBA Automated Engineering Agent - Available Commands:" & vbCrLf & vbCrLf
    help = help & "scan repo and inventory - Generate repository inventory" & vbCrLf
    help = help & "verify json - Check if JsonConverter.bas exists" & vbCrLf
    help = help & "scan core_utils - List public APIs in Core_Utils.bas" & vbCrLf
    help = help & "analyze driver - Check DriverSAP integration with Core_Utils" & vbCrLf
    help = help & "generate adapter mapping csv - Create legacy->driver mapping table" & vbCrLf
    help = help & "help - Show this help message" & vbCrLf & vbCrLf
    help = help & "Token Budget: " & (m_TokenBudget - m_TokensUsed) & " remaining" & vbCrLf
    GetHelpText = help
End Function
