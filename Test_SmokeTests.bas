' ==========================================
' SMOKE TEST SCRIPTS FOR DTS DRIVERS
' ==========================================
' Purpose: Quick validation tests for driver functionality
' Run these manually in Excel VBA immediate window
' ==========================================

Option Explicit

' ==========================================
' TEST 1: LibDTS_Base Utilities
' ==========================================
Public Sub Test_LibDTS_Base()
    Debug.Print "========== Testing LibDTS_Base =========="
    
    ' Test GUID generation
    Dim guid1 As String, guid2 As String
    guid1 = LibDTS_Base.GenerateGUID()
    guid2 = LibDTS_Base.GenerateGUID()
    
    Debug.Print "Generated GUID 1: " & guid1
    Debug.Print "Generated GUID 2: " & guid2
    Debug.Print "GUIDs are different: " & (guid1 <> guid2)
    Debug.Print "GUID1 is valid: " & LibDTS_Base.IsValidGUID(guid1)
    
    ' Test JSON operations
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.Add "name", "Test Frame"
    dict.Add "section", "W12X26"
    dict.Add "length", 5000
    
    Dim jsonStr As String
    jsonStr = LibDTS_Base.ToJson(dict)
    Debug.Print "JSON String: " & jsonStr
    
    Dim dict2 As Object
    Set dict2 = LibDTS_Base.ParseJson(jsonStr)
    Debug.Print "Parsed back - name: " & dict2("name")
    Debug.Print "Parsed back - section: " & dict2("section")
    
    Debug.Print "========== LibDTS_Base Tests Complete =========="
End Sub

' ==========================================
' TEST 2: LibDTS_Security
' ==========================================
Public Sub Test_LibDTS_Security()
    Debug.Print "========== Testing LibDTS_Security =========="
    
    Dim originalText As String
    originalText = "This is sensitive data with special chars: ñ, é, ü"
    
    Debug.Print "Original: " & originalText
    
    Dim encrypted As String
    encrypted = LibDTS_Security.Encrypt(originalText)
    Debug.Print "Encrypted: " & encrypted
    
    Dim decrypted As String
    decrypted = LibDTS_Security.Decrypt(encrypted)
    Debug.Print "Decrypted: " & decrypted
    
    Debug.Print "Round-trip successful: " & (originalText = decrypted)
    
    Debug.Print "========== LibDTS_Security Tests Complete =========="
End Sub

' ==========================================
' TEST 3: LibDTS_DriverSAP Connection
' ==========================================
Public Sub Test_DriverSAP_Connection()
    Debug.Print "========== Testing LibDTS_DriverSAP Connection =========="
    
    ' Test connection
    Dim connected As Boolean
    connected = LibDTS_DriverSAP.Connect()
    
    Debug.Print "Connection successful: " & connected
    
    If connected Then
        Debug.Print "Is connected: " & LibDTS_DriverSAP.IsConnected()
        
        ' Test disconnect
        Dim disconnected As Boolean
        disconnected = LibDTS_DriverSAP.Disconnect()
        Debug.Print "Disconnect successful: " & disconnected
        Debug.Print "Is connected after disconnect: " & LibDTS_DriverSAP.IsConnected()
    Else
        Debug.Print "Error: " & LibDTS_DriverSAP.GetLastError()
    End If
    
    Debug.Print "========== LibDTS_DriverSAP Connection Tests Complete =========="
End Sub

' ==========================================
' TEST 4: LibDTS_DriverSAP Dry-Run
' ==========================================
Public Sub Test_DriverSAP_DryRun()
    Debug.Print "========== Testing LibDTS_DriverSAP Dry-Run =========="
    
    ' Connect to SAP
    If Not LibDTS_DriverSAP.Connect() Then
        Debug.Print "Could not connect to SAP. Test aborted."
        Exit Sub
    End If
    
    ' Create test point
    Dim pt As New clsDTSPoint
    pt.Init 0, 0, 0
    
    ' Test dry-run
    Dim resultDryRun As String
    resultDryRun = LibDTS_DriverSAP.PushPoint(pt, dryRun:=True)
    Debug.Print "Dry-run result: " & resultDryRun
    Debug.Print "Dry-run successful (no actual creation): " & (InStr(resultDryRun, "DRY_RUN") > 0)
    
    ' Test actual creation
    Dim resultReal As String
    resultReal = LibDTS_DriverSAP.PushPoint(pt, dryRun:=False)
    Debug.Print "Real creation result: " & resultReal
    Debug.Print "Point created successfully: " & (Len(resultReal) > 0 And InStr(resultReal, "DRY_RUN") = 0)
    
    ' Clean up
    LibDTS_DriverSAP.Disconnect
    
    Debug.Print "========== LibDTS_DriverSAP Dry-Run Tests Complete =========="
End Sub

' ==========================================
' TEST 5: LibDTS_DriverCAD (Requires AutoCAD)
' ==========================================
Public Sub Test_DriverCAD_Drawing()
    Debug.Print "========== Testing LibDTS_DriverCAD Drawing =========="
    
    ' Get AutoCAD application
    Dim acadApp As Object
    On Error Resume Next
    Set acadApp = GetObject(, "AutoCAD.Application")
    On Error GoTo 0
    
    If acadApp Is Nothing Then
        Debug.Print "AutoCAD not running. Test skipped."
        Exit Sub
    End If
    
    Dim acadDoc As Object
    Set acadDoc = acadApp.ActiveDocument
    
    ' Create test frame
    Dim frame As New clsDTSFrame
    frame.StartPoint.Init 0, 0, 0
    frame.EndPoint.Init 1000, 0, 0
    frame.sectionName = "W12X26"
    
    ' Test dry-run
    Dim objDryRun As Object
    Set objDryRun = LibDTS_DriverCAD.DrawFrame(frame, acadDoc, dryRun:=True)
    Debug.Print "Dry-run (should be Nothing): " & (objDryRun Is Nothing)
    
    ' Test actual drawing
    Dim obj As Object
    Set obj = LibDTS_DriverCAD.DrawFrame(frame, acadDoc, dryRun:=False)
    Debug.Print "Frame drawn: " & Not (obj Is Nothing)
    
    If Not obj Is Nothing Then
        Debug.Print "Frame handle: " & obj.Handle
        
        ' Test reading back
        Dim frameRead As clsDTSFrame
        Set frameRead = LibDTS_DriverCAD.ReadFrame(obj)
        Debug.Print "Frame read back: " & Not (frameRead Is Nothing)
        
        If Not frameRead Is Nothing Then
            Debug.Print "GUID matches: " & (frameRead.Base.guid = frame.Base.guid)
        End If
    End If
    
    Debug.Print "========== LibDTS_DriverCAD Drawing Tests Complete =========="
End Sub

' ==========================================
' TEST 6: LibDTS_DriverDB Mapping
' ==========================================
Public Sub Test_DriverDB_Mapping()
    Debug.Print "========== Testing LibDTS_DriverDB Mapping =========="
    
    ' Generate test GUID
    Dim testGUID As String
    testGUID = LibDTS_Base.GenerateGUID()
    
    Debug.Print "Test GUID: " & testGUID
    
    ' Set mapping
    Dim elementInfo As Variant
    elementInfo = Array("Frame", "F123", "P1", "P2", "W12X26")
    
    LibDTS_DriverDB.SetMappedElement testGUID, elementInfo
    Debug.Print "Mapping set"
    
    ' Get mapping back
    Dim retrieved As Variant
    retrieved = LibDTS_DriverDB.GetMappedElement(testGUID)
    
    If Not IsEmpty(retrieved) Then
        Debug.Print "Mapping retrieved successfully"
        If IsArray(retrieved) Then
            Debug.Print "Element Type: " & retrieved(0)
            Debug.Print "Element Name: " & retrieved(1)
        End If
    Else
        Debug.Print "Error: Mapping not found"
    End If
    
    ' Test validation
    Dim problems As Collection
    Set problems = LibDTS_DriverDB.ValidateMappingIntegrity()
    Debug.Print "Validation problems found: " & problems.count
    
    Debug.Print "========== LibDTS_DriverDB Mapping Tests Complete =========="
End Sub

' ==========================================
' RUN ALL TESTS
' ==========================================
Public Sub RunAllSmokeTests()
    Debug.Print ""
    Debug.Print "########## STARTING ALL SMOKE TESTS ##########"
    Debug.Print ""
    
    Test_LibDTS_Base
    Debug.Print ""
    
    Test_LibDTS_Security
    Debug.Print ""
    
    Test_DriverSAP_Connection
    Debug.Print ""
    
    ' Test_DriverSAP_DryRun  ' Uncomment if SAP is available
    ' Debug.Print ""
    
    ' Test_DriverCAD_Drawing  ' Uncomment if AutoCAD is available
    ' Debug.Print ""
    
    Test_DriverDB_Mapping
    Debug.Print ""
    
    Debug.Print "########## ALL SMOKE TESTS COMPLETE ##########"
End Sub

' ==========================================
' TEST 7: LibDTS_AutoAgent Initialization
' ==========================================
Public Sub Test_AutoAgent_Init()
    Debug.Print "========== Testing LibDTS_AutoAgent Initialization =========="
    
    ' Test with default parameters
    LibDTS_AutoAgent.Initialize
    Debug.Print "Initialized with defaults"
    
    ' Test with custom parameters
    LibDTS_AutoAgent.Initialize tokenBudget:=50000, dryRun:=True, userApprovalRequired:=True
    Debug.Print "Initialized with custom parameters"
    
    Debug.Print "========== LibDTS_AutoAgent Initialization Tests Complete =========="
End Sub

' ==========================================
' TEST 8: LibDTS_AutoAgent Verification
' ==========================================
Public Sub Test_AutoAgent_Verification()
    Debug.Print "========== Testing LibDTS_AutoAgent Verification =========="
    
    LibDTS_AutoAgent.Initialize tokenBudget:=10000, dryRun:=True
    
    ' Test JsonConverter verification
    Dim jsonFound As Boolean
    jsonFound = LibDTS_AutoAgent.VerifyJsonConverter()
    Debug.Print "JsonConverter.bas found: " & jsonFound
    
    ' Test Core_Utils scanning
    Dim coreHelpers As Object
    Set coreHelpers = LibDTS_AutoAgent.ScanCoreUtils()
    Debug.Print "Core_Utils public APIs found: " & coreHelpers.Count
    
    If coreHelpers.Count > 0 Then
        Debug.Print "Sample APIs in Core_Utils:"
        Dim helperName As Variant
        Dim counter As Long
        counter = 0
        For Each helperName In coreHelpers.Keys
            Debug.Print "  - " & helperName
            counter = counter + 1
            If counter >= 5 Then Exit For
        Next helperName
    End If
    
    Debug.Print "========== LibDTS_AutoAgent Verification Tests Complete =========="
End Sub

' ==========================================
' TEST 9: LibDTS_AutoAgent Integration Analysis
' ==========================================
Public Sub Test_AutoAgent_Analysis()
    Debug.Print "========== Testing LibDTS_AutoAgent Integration Analysis =========="
    
    LibDTS_AutoAgent.Initialize tokenBudget:=10000, dryRun:=True
    
    ' Test driver integration analysis
    Dim analysis As Object
    Set analysis = LibDTS_AutoAgent.AnalyzeDriverSAPIntegration()
    
    If Not analysis Is Nothing Then
        Debug.Print "Analysis results:"
        Debug.Print "  Core_Utils helpers: " & analysis("coreHelpersCount")
        Debug.Print "  Reused in DriverSAP: " & analysis("reusedCount")
        Debug.Print "  Missing SAP helpers: " & analysis("missingCount")
        
        If analysis("missingCount") > 0 Then
            Debug.Print "  Missing functions:"
            Dim missingFuncs As Object
            Set missingFuncs = analysis("missingFunctions")
            
            Dim funcName As Variant
            For Each funcName In missingFuncs.Keys
                Debug.Print "    - " & funcName
            Next funcName
        End If
    End If
    
    Debug.Print "========== LibDTS_AutoAgent Integration Analysis Tests Complete =========="
End Sub

' ==========================================
' TEST 10: LibDTS_AutoAgent Command Interface
' ==========================================
Public Sub Test_AutoAgent_Commands()
    Debug.Print "========== Testing LibDTS_AutoAgent Commands =========="
    
    LibDTS_AutoAgent.Initialize tokenBudget:=15000, dryRun:=True
    
    ' Test help command
    Dim helpText As String
    helpText = LibDTS_AutoAgent.ExecuteCommand("help")
    Debug.Print "Help command executed, output length: " & Len(helpText)
    
    ' Test verify command
    Dim verifyResult As Boolean
    verifyResult = LibDTS_AutoAgent.ExecuteCommand("verify json")
    Debug.Print "Verify JSON command result: " & verifyResult
    
    ' Test scan command
    Dim scanCoreResult As Object
    Set scanCoreResult = LibDTS_AutoAgent.ExecuteCommand("scan core_utils")
    Debug.Print "Scan Core_Utils command result count: " & scanCoreResult.Count
    
    ' Test adapter mapping generation
    Dim mappingCSV As String
    mappingCSV = LibDTS_AutoAgent.ExecuteCommand("generate adapter mapping csv")
    Debug.Print "Adapter mapping CSV generated, length: " & Len(mappingCSV)
    Debug.Print "Sample lines:"
    Dim lines() As String
    lines = Split(mappingCSV, vbCrLf)
    Dim i As Long
    For i = 0 To UBound(lines)
        If i < 5 Then Debug.Print "  " & lines(i)
    Next i
    
    Debug.Print "========== LibDTS_AutoAgent Commands Tests Complete =========="
End Sub

' ==========================================
' TEST 11: LibDTS_AutoAgent Full Scan (Optional - takes longer)
' ==========================================
Public Sub Test_AutoAgent_FullScan()
    Debug.Print "========== Testing LibDTS_AutoAgent Full Repository Scan =========="
    
    LibDTS_AutoAgent.Initialize tokenBudget:=25000, dryRun:=True
    
    ' Test full repository scan
    Dim scanResult As String
    scanResult = LibDTS_AutoAgent.ExecuteCommand("scan repo and inventory")
    
    Debug.Print "Scan complete, output length: " & Len(scanResult)
    
    ' Display first few lines
    Dim lines() As String
    lines = Split(scanResult, vbCrLf)
    Debug.Print "First 10 lines of inventory:"
    Dim i As Long
    For i = 0 To UBound(lines)
        If i < 10 Then Debug.Print lines(i)
    Next i
    
    Debug.Print "========== LibDTS_AutoAgent Full Scan Tests Complete =========="
End Sub

' ==========================================
' RUN ALL AGENT TESTS
' ==========================================
Public Sub RunAllAgentTests()
    Debug.Print ""
    Debug.Print "########## STARTING AUTO AGENT TESTS ##########"
    Debug.Print ""
    
    Test_AutoAgent_Init
    Debug.Print ""
    
    Test_AutoAgent_Verification
    Debug.Print ""
    
    Test_AutoAgent_Analysis
    Debug.Print ""
    
    Test_AutoAgent_Commands
    Debug.Print ""
    
    ' Test_AutoAgent_FullScan  ' Uncomment for full scan (takes longer)
    
    Debug.Print "########## AUTO AGENT TESTS COMPLETE ##########"
End Sub
