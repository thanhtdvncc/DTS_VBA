# LibDTS_AutoAgent - Automated Engineering Agent

## Quick Start

```vba
' 1. Initialize the agent
LibDTS_AutoAgent.Initialize tokenBudget:=20000, dryRun:=True, userApprovalRequired:=True

' 2. Verify critical dependencies
If Not LibDTS_AutoAgent.VerifyJsonConverter() Then
    MsgBox "ERROR: JsonConverter.bas not found!", vbCritical
    Exit Sub
End If

' 3. Execute commands
Dim result As Variant
result = LibDTS_AutoAgent.ExecuteCommand("help")
Debug.Print result
```

## Available Commands

| Command | Description | Token Cost |
|---------|-------------|------------|
| `help` | Show available commands | ~10 |
| `verify json` | Check JsonConverter.bas exists | ~50 |
| `scan core_utils` | List Core_Utils APIs | ~200 |
| `analyze driver` | Check DriverSAP integration | ~300 |
| `generate adapter mapping csv` | Create migration map | ~500 |
| `scan repo and inventory` | Full repository scan | ~2000 |

## Configuration

### Repository Path

Set environment variable (recommended):
```
SET DTS_REPO_PATH=C:\Projects\DTS_VBA
```

Or modify `GetRepoPath()` function in LibDTS_AutoAgent.bas:
```vba
Private Function GetRepoPath() As String
    GetRepoPath = "C:\Projects\DTS_VBA"
End Function
```

## Usage Examples

### Example 1: Basic Analysis

```vba
Sub Example_BasicAnalysis()
    ' Initialize
    LibDTS_AutoAgent.Initialize tokenBudget:=10000, dryRun:=True
    
    ' Verify dependencies
    LibDTS_AutoAgent.VerifyJsonConverter()
    
    ' Scan Core_Utils
    Dim helpers As Object
    Set helpers = LibDTS_AutoAgent.ExecuteCommand("scan core_utils")
    Debug.Print "Found " & helpers.Count & " APIs"
    
    ' Check budget
    LibDTS_AutoAgent.ReportTokens()
End Sub
```

### Example 2: Generate Documentation

```vba
Sub Example_GenerateDocs()
    ' Initialize with higher budget
    LibDTS_AutoAgent.Initialize tokenBudget:=30000, dryRun:=True
    
    ' Generate inventory
    Dim inventory As String
    inventory = LibDTS_AutoAgent.ExecuteCommand("scan repo and inventory")
    
    ' Save to file
    Dim fso As Object, ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.CreateTextFile(ThisWorkbook.Path & "\INVENTORY.md", True)
    ts.Write inventory
    ts.Close
    
    Debug.Print "Inventory saved"
End Sub
```

### Example 3: Integration Analysis

```vba
Sub Example_IntegrationAnalysis()
    ' Initialize
    LibDTS_AutoAgent.Initialize tokenBudget:=15000, dryRun:=True
    
    ' Analyze driver integration
    Dim analysis As Object
    Set analysis = LibDTS_AutoAgent.ExecuteCommand("analyze driver")
    
    ' Report
    Debug.Print "Core_Utils APIs: " & analysis("coreHelpersCount")
    Debug.Print "Reused: " & analysis("reusedCount")
    Debug.Print "Missing: " & analysis("missingCount")
End Sub
```

## Testing

Run all tests:
```vba
RunAllAgentTests()
```

Run specific test:
```vba
Test_AutoAgent_Init()
Test_AutoAgent_Verification()
Test_AutoAgent_Commands()
```

## Documentation

For complete documentation, see:
- **DOCS_AUTO_AGENT_GUIDE.md** - Full user guide
- **DOCS_XDATA_FORMAT.md** - XData format specification
- **DOCS_IMPLEMENTATION_SUMMARY.md** - Implementation details

## Token Budget Management

```vba
' Check remaining budget
LibDTS_AutoAgent.ReportTokens()
' Output: Tokens remaining: 18750

' Re-initialize with larger budget if needed
LibDTS_AutoAgent.Initialize tokenBudget:=50000, dryRun:=True
```

## Error Handling

```vba
On Error GoTo ErrHandler

Dim result As Variant
result = LibDTS_AutoAgent.ExecuteCommand("some command")

Exit Sub

ErrHandler:
    Dim lastError As String
    lastError = LibDTS_AutoAgent.GetLastError()
    Debug.Print "Error: " & lastError
```

## Features

✅ CLI-style command interface  
✅ Token budget tracking  
✅ Dry-run mode for safety  
✅ User approval control  
✅ Repository scanning and analysis  
✅ Integration analysis tools  
✅ Migration mapping generation  
✅ Comprehensive error handling  
✅ Automated testing  

## Requirements

- **JsonConverter.bas** - Required (verified automatically)
- **Core_Utils.bas** - Optional (for analysis features)
- **LibDTS_Base.bas** - Required (GUID & JSON utilities)
- **LibDTS_Logger.bas** - Required (logging)

## Version

**Current Version:** 1.0  
**Last Updated:** 2025-11-23  
**Status:** Production Ready  

## Support

For issues or questions:
1. Check **DOCS_AUTO_AGENT_GUIDE.md** for detailed help
2. Review test examples in **Test_SmokeTests.bas**
3. Check logs at `%TEMP%\DTS_System.log`

## License

Part of DTS_VBA project. See repository LICENSE file.
