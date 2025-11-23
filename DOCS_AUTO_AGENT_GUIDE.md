# DTS_VBA Automated Engineering Agent - User Guide

## Overview

The **LibDTS_AutoAgent** module provides a CLI-style automated engineering agent for managing the DTS_VBA repository consolidation project. It implements token budget tracking, dry-run support, and various analysis commands to facilitate API consolidation and driver enhancement.

---

## Quick Start

### 1. Initialize the Agent

```vba
' In VBA Immediate Window or a test procedure
LibDTS_AutoAgent.Initialize tokenBudget:=20000, dryRun:=True, userApprovalRequired:=True
```

**Parameters:**
- `tokenBudget` (Optional, default: 20000): Total tokens permitted for operations
- `dryRun` (Optional, default: True): Validation-only mode
- `userApprovalRequired` (Optional, default: True): Require confirmation before writes

**Output:**
```
Tokens remaining: 20000
```

### 2. Verify Critical Dependencies

```vba
' Check if JsonConverter.bas exists
Dim jsonFound As Boolean
jsonFound = LibDTS_AutoAgent.VerifyJsonConverter()
Debug.Print "JsonConverter found: " & jsonFound
```

### 3. Execute Commands

```vba
' Get help on available commands
Dim helpText As String
helpText = LibDTS_AutoAgent.ExecuteCommand("help")
Debug.Print helpText

' Scan Core_Utils for SAP helpers
Dim coreHelpers As Object
Set coreHelpers = LibDTS_AutoAgent.ExecuteCommand("scan core_utils")
Debug.Print "Core_Utils APIs: " & coreHelpers.Count

' Generate full repository inventory
Dim inventory As String
inventory = LibDTS_AutoAgent.ExecuteCommand("scan repo and inventory")
' Save to file or display in form
```

---

## Available Commands

### Command: `help`

**Description:** Display available commands and current token budget

**Usage:**
```vba
result = LibDTS_AutoAgent.ExecuteCommand("help")
```

**Output:**
```
DTS_VBA Automated Engineering Agent - Available Commands:

scan repo and inventory - Generate repository inventory
verify json - Check if JsonConverter.bas exists
scan core_utils - List public APIs in Core_Utils.bas
analyze driver - Check DriverSAP integration with Core_Utils
generate adapter mapping csv - Create legacy->driver mapping table
help - Show this help message

Token Budget: 19950 remaining
```

---

### Command: `verify json`

**Description:** Verify that JsonConverter.bas exists in the repository (critical dependency)

**Usage:**
```vba
Dim found As Boolean
found = LibDTS_AutoAgent.ExecuteCommand("verify json")
```

**Returns:** Boolean (True if found, False otherwise)

**Token Cost:** ~50 tokens

---

### Command: `scan core_utils`

**Description:** Scan Core_Utils.bas and list all public functions/subs

**Usage:**
```vba
Dim helpers As Object
Set helpers = LibDTS_AutoAgent.ExecuteCommand("scan core_utils")

' Iterate through results
Dim helperName As Variant
For Each helperName In helpers.Keys
    Debug.Print helperName & " - " & helpers(helperName)
Next helperName
```

**Returns:** Dictionary object with function names as keys and line info as values

**Token Cost:** ~200 tokens

**Example Output:**
```
DTS_SAP2000_Getlist - Line 27: Public Function DTS_SAP2000_Getlist(...)
DTS_SAP2000_Getlist2node - Line 96: Public Function DTS_SAP2000_Getlist2node(...)
EnsureSapModelAvailable - Line 500: Public Function EnsureSapModelAvailable()
```

---

### Command: `analyze driver`

**Description:** Analyze LibDTS_DriverSAP integration with Core_Utils helpers

**Usage:**
```vba
Dim analysis As Object
Set analysis = LibDTS_AutoAgent.ExecuteCommand("analyze driver")

Debug.Print "Core helpers: " & analysis("coreHelpersCount")
Debug.Print "Reused in DriverSAP: " & analysis("reusedCount")
Debug.Print "Missing SAP helpers: " & analysis("missingCount")

' List missing functions
Dim missingFuncs As Object
Set missingFuncs = analysis("missingFunctions")
Dim funcName As Variant
For Each funcName In missingFuncs.Keys
    Debug.Print "  Missing: " & funcName
Next funcName
```

**Returns:** Dictionary with analysis results

**Token Cost:** ~300 tokens

**Output Fields:**
- `coreHelpersCount`: Total public APIs in Core_Utils.bas
- `reusedCount`: Number referenced in LibDTS_DriverSAP
- `missingCount`: Number of SAP-related helpers not referenced
- `reusedFunctions`: Dictionary of reused functions
- `missingFunctions`: Dictionary of missing SAP helpers

---

### Command: `scan repo and inventory`

**Description:** Generate comprehensive repository inventory of all VBA files and public APIs

**Usage:**
```vba
' Generate Markdown format (default)
Dim inventoryMD As String
inventoryMD = LibDTS_AutoAgent.ExecuteCommand("scan repo and inventory")

' Save to file
Dim fso As Object, ts As Object
Set fso = CreateObject("Scripting.FileSystemObject")
Set ts = fso.CreateTextFile("C:\Temp\DTS_Inventory.md", True)
ts.Write inventoryMD
ts.Close
```

**Returns:** String in Markdown format

**Token Cost:** ~2000 tokens (high cost - ensure sufficient budget)

**Output Format:**
```markdown
# DTS_VBA Repository Inventory

**Generated:** 2025-11-23 03:50:00

## Summary

- Total Files: 56
- Total Public APIs: 234

## Files

### LibDTS_DriverSAP.bas

**Path:** C:\Project\DTS_VBA\LibDTS_DriverSAP.bas
**Public APIs:** 15

- `Connect` (Line 39)
- `Disconnect` (Line 97)
- `IsConnected` (Line 143)
...
```

**Alternative Formats:**

```vba
' Get as Dictionary object for programmatic access
Dim inventoryDict As Object
Set inventoryDict = LibDTS_AutoAgent.ScanRepository("DICT")

' Iterate through files
Dim filePath As Variant
For Each filePath In inventoryDict.Keys
    Dim fileInfo As Object
    Set fileInfo = inventoryDict(filePath)
    Debug.Print fileInfo("name") & " has " & fileInfo("apiCount") & " APIs"
Next filePath
```

---

### Command: `generate adapter mapping csv`

**Description:** Generate CSV mapping table for legacy function migration to driver APIs

**Usage:**
```vba
Dim mappingCSV As String
mappingCSV = LibDTS_AutoAgent.ExecuteCommand("generate adapter mapping csv")

' Save to file
Dim fso As Object, ts As Object
Set fso = CreateObject("Scripting.FileSystemObject")
Set ts = fso.CreateTextFile("C:\Temp\Legacy_Mapping.csv", True)
ts.Write mappingCSV
ts.Close
```

**Returns:** String in CSV format

**Token Cost:** ~500 tokens

**Output Format:**
```csv
Legacy_Module,Legacy_Function,Driver_Module,Driver_Function,Notes
m01_SAP2000_Connection,ConnectSAP2000,LibDTS_DriverSAP,Connect,Enhanced version detection
m01_SAP2000_Connection,DisconnectSAP2000,LibDTS_DriverSAP,Disconnect,Simplified
Core_Utils,IsSAPConnected,LibDTS_DriverSAP,IsConnected,Direct replacement
Core_XData_Reader,ReadPointsFromCAD,LibDTS_DriverCAD,ReadAllPoints,Batch operation
Core_XData_Reader,ReadFramesFromCAD,LibDTS_DriverCAD,ReadAllFrames,Batch operation
...
```

---

## Token Budget Management

### Understanding Token Budget

The agent tracks token consumption to prevent resource exhaustion and provide cost transparency. Each operation estimates and consumes tokens based on complexity.

**Token Costs by Operation:**
- `verify json`: ~50 tokens
- `scan core_utils`: ~200 tokens
- `analyze driver`: ~300 tokens
- `generate adapter mapping csv`: ~500 tokens
- `scan repo and inventory`: ~2000 tokens

### Checking Remaining Budget

```vba
' After any operation
LibDTS_AutoAgent.ReportTokens()
```

**Output:**
```
Tokens remaining: 17450
```

### When Budget Exhausted

If an operation exceeds remaining budget:
```
Insufficient tokens. Required: 2000, Available: 500
Propose splitting into sub-tasks. Confirm to proceed? (Y/N)
token exhausted
```

**Solution:** Re-initialize with larger budget
```vba
LibDTS_AutoAgent.Initialize tokenBudget:=50000, dryRun:=True
```

---

## Testing

### Run All Agent Tests

```vba
' Run comprehensive test suite
RunAllAgentTests()
```

**Tests Include:**
1. **Test_AutoAgent_Init**: Initialization with various parameters
2. **Test_AutoAgent_Verification**: JsonConverter and Core_Utils scanning
3. **Test_AutoAgent_Analysis**: Driver integration analysis
4. **Test_AutoAgent_Commands**: CLI command interface
5. **Test_AutoAgent_FullScan**: Full repository scan (optional, high cost)

**Expected Output:**
```
########## STARTING AUTO AGENT TESTS ##########

========== Testing LibDTS_AutoAgent Initialization ==========
Initialized with defaults
Initialized with custom parameters
========== LibDTS_AutoAgent Initialization Tests Complete ==========

========== Testing LibDTS_AutoAgent Verification ==========
[OK] JsonConverter.bas exists
Core_Utils public APIs found: 42
Sample APIs in Core_Utils:
  - DTS_SAP2000_Getlist
  - DTS_SAP2000_Getlist2node
  - EnsureSapModelAvailable
  - GetElementsByPropertyValue_Core
  - FilterByCoordinateRange_Core
========== LibDTS_AutoAgent Verification Tests Complete ==========

...

########## AUTO AGENT TESTS COMPLETE ##########
```

---

## Integration with Existing Workflow

### Step 1: Initial Assessment

```vba
' Initialize agent
LibDTS_AutoAgent.Initialize tokenBudget:=25000, dryRun:=True

' Verify dependencies
If Not LibDTS_AutoAgent.VerifyJsonConverter() Then
    MsgBox "Critical error: JsonConverter.bas not found", vbCritical
    Exit Sub
End If

' Scan Core_Utils
Dim coreHelpers As Object
Set coreHelpers = LibDTS_AutoAgent.ExecuteCommand("scan core_utils")
Debug.Print "Found " & coreHelpers.Count & " helper APIs in Core_Utils"
```

### Step 2: Analyze Integration

```vba
' Check driver integration
Dim analysis As Object
Set analysis = LibDTS_AutoAgent.ExecuteCommand("analyze driver")

If analysis("missingCount") > 0 Then
    Debug.Print "WARNING: " & analysis("missingCount") & " SAP helpers not integrated"
    
    ' List missing functions for manual review
    Dim missingFuncs As Object
    Set missingFuncs = analysis("missingFunctions")
    
    Dim funcName As Variant
    For Each funcName In missingFuncs.Keys
        Debug.Print "  TODO: Integrate " & funcName
    Next funcName
End If
```

### Step 3: Generate Documentation

```vba
' Generate inventory
Dim inventoryMD As String
inventoryMD = LibDTS_AutoAgent.ExecuteCommand("scan repo and inventory")

' Save to project docs folder
Dim docsPath As String
docsPath = ThisWorkbook.Path & "\DOCS_INVENTORY_GENERATED.md"

Dim fso As Object, ts As Object
Set fso = CreateObject("Scripting.FileSystemObject")
Set ts = fso.CreateTextFile(docsPath, True)
ts.Write inventoryMD
ts.Close

Debug.Print "Inventory saved to: " & docsPath
```

### Step 4: Generate Migration Mapping

```vba
' Generate adapter mapping
Dim mappingCSV As String
mappingCSV = LibDTS_AutoAgent.ExecuteCommand("generate adapter mapping csv")

' Save for reference
Dim csvPath As String
csvPath = ThisWorkbook.Path & "\LEGACY_MIGRATION_MAP.csv"

Set ts = fso.CreateTextFile(csvPath, True)
ts.Write mappingCSV
ts.Close

Debug.Print "Migration map saved to: " & csvPath
```

---

## Advanced Usage

### Custom Repository Path

By default, the agent uses `ThisWorkbook.Path` as the repository root. To override:

```vba
' Edit LibDTS_AutoAgent.bas - GetRepoPath() function
Private Function GetRepoPath() As String
    ' Custom path
    GetRepoPath = "C:\Projects\DTS_VBA"
End Function
```

### Extending Command Set

Add new commands by editing `ExecuteCommand()`:

```vba
' In LibDTS_AutoAgent.bas
Public Function ExecuteCommand(command As String) As Variant
    Dim cmd As String
    cmd = LCase$(Trim$(command))
    
    Select Case True
        ' ... existing commands ...
        
        Case InStr(cmd, "your custom command") > 0
            ExecuteCommand = YourCustomFunction()
        
        Case Else
            ExecuteCommand = "Unknown command: " & command
    End Select
End Function
```

### Error Handling

```vba
On Error GoTo ErrHandler

Dim result As Variant
result = LibDTS_AutoAgent.ExecuteCommand("some command")

' Process result...
Exit Sub

ErrHandler:
    Dim lastError As String
    lastError = LibDTS_AutoAgent.GetLastError()
    Debug.Print "Agent error: " & lastError
```

---

## Troubleshooting

### Issue: "JsonConverter.bas not found"

**Solution:** Ensure JsonConverter.bas is in the repository root or correct the path in `VerifyJsonConverter()`.

### Issue: "token exhausted" message appears

**Solution:** Re-initialize with larger budget:
```vba
LibDTS_AutoAgent.Initialize tokenBudget:=100000, dryRun:=True
```

### Issue: Core_Utils scan returns 0 APIs

**Solution:** 
1. Check that Core_Utils.bas exists in repository
2. Verify file encoding is readable
3. Check parsing logic in `ScanCoreUtils()`

### Issue: Commands return empty results

**Solution:**
1. Ensure agent is initialized first
2. Check that file paths in `GetRepoPath()` are correct
3. Review error log via `GetLastError()`

---

## Best Practices

### 1. Always Initialize Before Use

```vba
' At start of session
LibDTS_AutoAgent.Initialize tokenBudget:=50000, dryRun:=True
```

### 2. Check Token Budget Regularly

```vba
' After expensive operations
LibDTS_AutoAgent.ReportTokens()
```

### 3. Save High-Cost Results

Don't re-run expensive scans unnecessarily:
```vba
' Run once and save
Dim inventory As String
inventory = LibDTS_AutoAgent.ExecuteCommand("scan repo and inventory")

' Save to file for reuse
' ... file operations ...
```

### 4. Use dryRun for Validation

```vba
' Always start with dryRun=True for safety
LibDTS_AutoAgent.Initialize tokenBudget:=20000, dryRun:=True

' Switch to False only when ready to commit changes
LibDTS_AutoAgent.Initialize tokenBudget:=20000, dryRun:=False, userApprovalRequired:=True
```

### 5. Test with Small Operations First

```vba
' Start with low-cost commands
result = LibDTS_AutoAgent.ExecuteCommand("verify json")
result = LibDTS_AutoAgent.ExecuteCommand("scan core_utils")

' Then move to expensive operations
result = LibDTS_AutoAgent.ExecuteCommand("scan repo and inventory")
```

---

## Appendix: Command Reference Table

| Command | Description | Returns | Token Cost | Notes |
|---------|-------------|---------|------------|-------|
| `help` | Show available commands | String | ~10 | Always available |
| `verify json` | Check JsonConverter.bas exists | Boolean | ~50 | Critical check |
| `scan core_utils` | List Core_Utils APIs | Dictionary | ~200 | Analysis tool |
| `analyze driver` | Check DriverSAP integration | Dictionary | ~300 | Integration check |
| `generate adapter mapping csv` | Create migration map | String (CSV) | ~500 | Migration aid |
| `scan repo and inventory` | Full repository scan | String (MD) | ~2000 | High cost |

---

## Support and Feedback

For issues, suggestions, or contributions:
- Review code: `LibDTS_AutoAgent.bas`
- Check logs: `%TEMP%\DTS_System.log`
- Run tests: `RunAllAgentTests()`

---

**Document Version:** 1.0  
**Last Updated:** 2025-11-23  
**Module Version:** LibDTS_AutoAgent v1.0
