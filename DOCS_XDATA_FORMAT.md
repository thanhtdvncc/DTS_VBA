# DTS_VBA XData Flexible Format Strategy

## Overview

This document specifies the **XData Flexible Format** strategy for DTS_VBA, supporting both legacy index-based XData and modern key-value JSON format with backward compatibility and future extensibility.

---

## Design Principles

### 1. Key-Value + JSON Hybrid Approach

**Strategy:** Use DXF tags in a key-value pattern with JSON support for complex data

**Benefits:**
- Simple properties → Key-Value pairs (efficient, readable)
- Complex properties → JSON under reserved key (flexible, extensible)
- Backward compatible with legacy format
- Forward compatible with future schema changes

### 2. Configurable RegApp Name

**Default:** `DTS_CORE_DATA` (from LibDTS_Global.DTS_APP_NAME)
**Configurable via:** clsDTSConfig settings

### 3. Schema Versioning

**Version Field:** `SCHEMA_VER` key in XData
**Current Version:** `2.0`
**Legacy Version:** `1.0` (index-based)

---

## XData Format Specification

### Modern Format (v2.0) - Key-Value + JSON

#### Structure

```
RegApp: "DTS_CORE_DATA"
1001: "DTS_CORE_DATA"          ' RegApp registration
1000: "SCHEMA_VER"              ' Schema version key
1000: "2.0"                     ' Version value
1000: "DTS_GUID"                ' GUID key
1000: "{12345678-1234-...}"     ' GUID value
1000: "ELEM_TYPE"               ' Element type key
1071: 1                         ' Integer: Frame type (1=FRAME)
1000: "LAYER"                   ' Layer key
1000: "DTS_FRAMES"              ' Layer value
1000: "SECTION"                 ' Section name key
1000: "W12X26"                  ' Section value
1000: "PROPS_JSON"              ' JSON properties key
1000: "{\"material\":\"A992\",...}" ' JSON value (complex props)
```

#### DXF Tag Types

| DXF Code | Data Type | Max Length | Usage |
|----------|-----------|------------|-------|
| 1000 | String | 255 chars | Keys and string values |
| 1001 | RegApp Name | - | Application registration |
| 1040 | Double | - | Numeric values (coordinates, lengths) |
| 1070 | Short Integer | - | Enums, flags |
| 1071 | Long Integer | - | Large integers, IDs |

#### Reserved Keys

| Key | Type | Description |
|-----|------|-------------|
| `SCHEMA_VER` | String | XData schema version ("2.0") |
| `DTS_GUID` | String | Element unique identifier |
| `ELEM_TYPE` | Integer | DTSElementType enum value |
| `LAYER` | String | AutoCAD layer name |
| `COLOR` | Integer | AutoCAD color index |
| `PROPS_JSON` | String | JSON for complex properties |
| `LEGACY_FLAG` | Integer | 1 if migrated from legacy format |

---

### Legacy Format (v1.0) - Index-Based

#### Structure

```
RegApp: "DTS_CORE_DATA"
1001: "DTS_CORE_DATA"
1000: "{12345678-1234-...}"  ' Index 0: GUID
1070: 1                       ' Index 1: Element Type
1000: "DTS_FRAMES"            ' Index 2: Layer
1000: "W12X26"                ' Index 3: Section
...
```

**Issue:** Fixed indices, not self-describing, hard to extend

---

## Implementation

### LibDTS_DriverCAD Enhancement

#### 1. SaveXData (Write)

```vba
' Enhanced SaveXData with key-value format
' Parameters:
'   dtsObj: Core object (clsDTSFrame, clsDTSPoint, etc.)
'   ent: AutoCAD entity object
'   useJSON: True to store complex props as JSON (default: True)
Public Sub SaveXData(dtsObj As Object, _
                     ent As Object, _
                     Optional useJSON As Boolean = True)
    On Error GoTo ErrHandler
    
    ' Validate inputs
    If dtsObj Is Nothing Or ent Is Nothing Then Exit Sub
    
    ' Get RegApp name from config
    Dim appName As String
    appName = DTS_APP_NAME
    
    ' Register application if needed
    RegisterXDataApp ent.Document.Application, appName
    
    ' Build XData arrays
    Dim xdataType() As Integer
    Dim xdataVal() As Variant
    
    Dim idx As Long
    idx = 0
    
    ' Header: RegApp
    ReDim Preserve xdataType(idx)
    ReDim Preserve xdataVal(idx)
    xdataType(idx) = 1001
    xdataVal(idx) = appName
    idx = idx + 1
    
    ' Schema version
    AddXDataPair xdataType, xdataVal, idx, "SCHEMA_VER", "2.0"
    
    ' GUID
    Dim guid As String
    guid = dtsObj.Base.guid
    If Len(guid) = 0 Then guid = LibDTS_Base.GenerateGUID()
    AddXDataPair xdataType, xdataVal, idx, "DTS_GUID", guid
    
    ' Element type
    AddXDataIntPair xdataType, xdataVal, idx, "ELEM_TYPE", dtsObj.Base.elementType
    
    ' Layer
    If Len(dtsObj.Base.layer) > 0 Then
        AddXDataPair xdataType, xdataVal, idx, "LAYER", dtsObj.Base.layer
    End If
    
    ' Color
    If dtsObj.Base.color <> 0 Then
        AddXDataIntPair xdataType, xdataVal, idx, "COLOR", dtsObj.Base.color
    End If
    
    ' Element-specific properties
    If TypeName(dtsObj) = "clsDTSFrame" Then
        ' Simple properties as key-value
        AddXDataPair xdataType, xdataVal, idx, "SECTION", dtsObj.sectionName
        AddXDataDoublePair xdataType, xdataVal, idx, "START_X", dtsObj.StartPoint.X
        AddXDataDoublePair xdataType, xdataVal, idx, "START_Y", dtsObj.StartPoint.Y
        AddXDataDoublePair xdataType, xdataVal, idx, "START_Z", dtsObj.StartPoint.Z
        AddXDataDoublePair xdataType, xdataVal, idx, "END_X", dtsObj.EndPoint.X
        AddXDataDoublePair xdataType, xdataVal, idx, "END_Y", dtsObj.EndPoint.Y
        AddXDataDoublePair xdataType, xdataVal, idx, "END_Z", dtsObj.EndPoint.Z
        
        ' Complex properties as JSON (if useJSON enabled)
        If useJSON And dtsObj.Base.Properties.Count > 0 Then
            Dim jsonProps As String
            jsonProps = LibDTS_Base.ToJson(dtsObj.Base.Properties)
            AddXDataPair xdataType, xdataVal, idx, "PROPS_JSON", jsonProps
        End If
    End If
    
    ' Attach XData to entity
    ent.SetXData xdataType, xdataVal
    
    LibDTS_Logger.Log DRIVER_NAME & ".SaveXData: Saved XData (v2.0) for " & guid, DTS_INFO
    Exit Sub
    
ErrHandler:
    m_LastError = "SaveXData error: " & err.description
    LibDTS_Logger.Log DRIVER_NAME & ".SaveXData: " & m_LastError, DTS_ERROR
End Sub

' Helper: Add string key-value pair
Private Sub AddXDataPair(ByRef typeArr() As Integer, _
                         ByRef valArr() As Variant, _
                         ByRef idx As Long, _
                         key As String, _
                         value As String)
    ' Add key
    ReDim Preserve typeArr(idx)
    ReDim Preserve valArr(idx)
    typeArr(idx) = 1000
    valArr(idx) = key
    idx = idx + 1
    
    ' Add value
    ReDim Preserve typeArr(idx)
    ReDim Preserve valArr(idx)
    typeArr(idx) = 1000
    valArr(idx) = value
    idx = idx + 1
End Sub

' Helper: Add integer key-value pair
Private Sub AddXDataIntPair(ByRef typeArr() As Integer, _
                            ByRef valArr() As Variant, _
                            ByRef idx As Long, _
                            key As String, _
                            value As Long)
    ' Add key
    ReDim Preserve typeArr(idx)
    ReDim Preserve valArr(idx)
    typeArr(idx) = 1000
    valArr(idx) = key
    idx = idx + 1
    
    ' Add value
    ReDim Preserve typeArr(idx)
    ReDim Preserve valArr(idx)
    typeArr(idx) = 1071
    valArr(idx) = value
    idx = idx + 1
End Sub

' Helper: Add double key-value pair
Private Sub AddXDataDoublePair(ByRef typeArr() As Integer, _
                               ByRef valArr() As Variant, _
                               ByRef idx As Long, _
                               key As String, _
                               value As Double)
    ' Add key
    ReDim Preserve typeArr(idx)
    ReDim Preserve valArr(idx)
    typeArr(idx) = 1000
    valArr(idx) = key
    idx = idx + 1
    
    ' Add value
    ReDim Preserve typeArr(idx)
    ReDim Preserve valArr(idx)
    typeArr(idx) = 1040
    valArr(idx) = value
    idx = idx + 1
End Sub
```

#### 2. ReadXData (Read)

```vba
' Enhanced ReadXData with backward compatibility
' Returns: Dictionary with key-value pairs
'   Special keys: "SCHEMA_VER", "LEGACY_FLAG"
Public Function ReadXData(ent As Object, _
                         Optional appName As String = "") As Object
    On Error GoTo ErrHandler
    
    ' Default app name
    If Len(appName) = 0 Then appName = DTS_APP_NAME
    
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    
    ' Get XData
    Dim xdataType As Variant
    Dim xdataVal As Variant
    ent.GetXData appName, xdataType, xdataVal
    
    ' Check if XData exists
    If IsEmpty(xdataType) Or IsEmpty(xdataVal) Then
        LibDTS_Logger.Log DRIVER_NAME & ".ReadXData: No XData found", DTS_WARNING
        Set ReadXData = result
        Exit Function
    End If
    
    ' Detect schema version
    Dim schemaVer As String
    schemaVer = DetectSchemaVersion(xdataType, xdataVal)
    result.Add "SCHEMA_VER", schemaVer
    
    ' Parse based on version
    If schemaVer = "2.0" Then
        ' Modern key-value format
        ParseKeyValueXData xdataType, xdataVal, result
    Else
        ' Legacy index-based format
        ParseLegacyXData xdataType, xdataVal, result
        result.Add "LEGACY_FLAG", 1
    End If
    
    Set ReadXData = result
    Exit Function
    
ErrHandler:
    m_LastError = "ReadXData error: " & err.description
    LibDTS_Logger.Log DRIVER_NAME & ".ReadXData: " & m_LastError, DTS_ERROR
    Set ReadXData = CreateObject("Scripting.Dictionary")
End Function

' Detect XData schema version
Private Function DetectSchemaVersion(xdataType As Variant, _
                                    xdataVal As Variant) As String
    ' Look for SCHEMA_VER key
    Dim i As Long
    For i = LBound(xdataVal) To UBound(xdataVal)
        If xdataType(i) = 1000 And xdataVal(i) = "SCHEMA_VER" Then
            ' Next value is version
            If i + 1 <= UBound(xdataVal) Then
                DetectSchemaVersion = CStr(xdataVal(i + 1))
                Exit Function
            End If
        End If
    Next i
    
    ' No version key found = legacy format
    DetectSchemaVersion = "1.0"
End Function

' Parse modern key-value XData
Private Sub ParseKeyValueXData(xdataType As Variant, _
                               xdataVal As Variant, _
                               result As Object)
    Dim i As Long
    i = 1 ' Skip RegApp at index 0
    
    Do While i <= UBound(xdataVal)
        ' Expect key (1000) followed by value (1000/1040/1071)
        If xdataType(i) = 1000 Then
            Dim key As String
            key = CStr(xdataVal(i))
            
            ' Skip if already processed (e.g., RegApp)
            If key <> DTS_APP_NAME And Not result.Exists(key) Then
                ' Get next value
                If i + 1 <= UBound(xdataVal) Then
                    Dim value As Variant
                    
                    Select Case xdataType(i + 1)
                        Case 1000  ' String
                            value = CStr(xdataVal(i + 1))
                        Case 1040  ' Double
                            value = CDbl(xdataVal(i + 1))
                        Case 1070, 1071  ' Integer
                            value = CLng(xdataVal(i + 1))
                        Case Else
                            value = xdataVal(i + 1)
                    End Select
                    
                    result.Add key, value
                    i = i + 2  ' Skip value
                Else
                    i = i + 1
                End If
            Else
                i = i + 1
            End If
        Else
            i = i + 1
        End If
    Loop
End Sub

' Parse legacy index-based XData and convert to key-value
Private Sub ParseLegacyXData(xdataType As Variant, _
                            xdataVal As Variant, _
                            result As Object)
    ' Legacy format indices (after RegApp):
    ' 0: GUID (1000)
    ' 1: Element Type (1070)
    ' 2: Layer (1000)
    ' 3: Section or other property (1000)
    
    Dim idx As Long
    idx = 1 ' Skip RegApp at 0
    
    ' GUID
    If idx <= UBound(xdataVal) And xdataType(idx) = 1000 Then
        result.Add "DTS_GUID", CStr(xdataVal(idx))
        idx = idx + 1
    End If
    
    ' Element Type
    If idx <= UBound(xdataVal) And (xdataType(idx) = 1070 Or xdataType(idx) = 1071) Then
        result.Add "ELEM_TYPE", CLng(xdataVal(idx))
        idx = idx + 1
    End If
    
    ' Layer
    If idx <= UBound(xdataVal) And xdataType(idx) = 1000 Then
        result.Add "LAYER", CStr(xdataVal(idx))
        idx = idx + 1
    End If
    
    ' Additional properties (generic)
    Dim propIdx As Long
    propIdx = 0
    Do While idx <= UBound(xdataVal)
        result.Add "PROP_" & propIdx, xdataVal(idx)
        propIdx = propIdx + 1
        idx = idx + 1
    Loop
    
    LibDTS_Logger.Log DRIVER_NAME & ".ParseLegacyXData: Converted legacy format to key-value", DTS_INFO
End Sub
```

---

## Migration Strategy

### Phase 1: Dual-Format Support (Current)

**Goals:**
- ✅ Read both legacy and modern formats
- ✅ Write modern format by default
- ✅ Flag legacy data for upgrade

**Implementation:**
- ReadXData detects schema version automatically
- SaveXData always writes v2.0 format
- Migration flag tracks converted entities

### Phase 2: Gradual Migration

**Process:**
1. **Scan**: Identify all entities with legacy XData
2. **Report**: Generate migration report
3. **Convert**: Batch convert legacy XData to v2.0
4. **Validate**: Verify data integrity after conversion

**Script:**
```vba
Public Sub MigrateXDataToV2()
    ' Get AutoCAD document
    Dim acadDoc As Object
    Set acadDoc = GetObject(, "AutoCAD.Application").ActiveDocument
    
    ' Scan all entities
    Dim ent As Object
    Dim legacyCount As Long
    Dim migratedCount As Long
    
    For Each ent In acadDoc.ModelSpace
        ' Check if has legacy XData
        Dim xdata As Object
        Set xdata = LibDTS_DriverCAD.ReadXData(ent)
        
        If xdata.Count > 0 Then
            If xdata.Exists("LEGACY_FLAG") Then
                ' Has legacy XData - convert it
                Dim dtsObj As Object
                Set dtsObj = LibDTS_DriverCAD.ReadFrame(ent) ' or ReadPoint, ReadArea
                
                If Not dtsObj Is Nothing Then
                    ' Re-save with modern format
                    LibDTS_DriverCAD.SaveXData dtsObj, ent, useJSON:=True
                    migratedCount = migratedCount + 1
                End If
                
                legacyCount = legacyCount + 1
            End If
        End If
    Next ent
    
    Debug.Print "Migration complete: " & migratedCount & " of " & legacyCount & " entities upgraded"
End Sub
```

### Phase 3: Legacy Support Deprecation (Future)

**Timeline:** After 6 months of dual-format usage

**Actions:**
- Remove legacy parsing code
- Update documentation
- Notify users of format requirement

---

## Best Practices

### 1. Always Use Key-Value Format for New Data

```vba
' Good
LibDTS_DriverCAD.SaveXData frameObj, acadEntity, useJSON:=True

' Avoid
' Manual XData array building without keys
```

### 2. Check Schema Version When Reading

```vba
Dim xdata As Object
Set xdata = LibDTS_DriverCAD.ReadXData(entity)

If xdata("SCHEMA_VER") = "2.0" Then
    ' Modern format - direct access
    Dim guid As String
    guid = xdata("DTS_GUID")
Else
    ' Legacy format - may need special handling
    If xdata.Exists("LEGACY_FLAG") Then
        ' Consider upgrading
    End If
End If
```

### 3. Use JSON for Complex Properties

```vba
' Store complex properties in JSON
Dim frame As New clsDTSFrame
frame.Base.Properties.Add "material", "A992"
frame.Base.Properties.Add "fireRating", "2HR"
frame.Base.Properties.Add "custom_data", "value"

' Will be stored under PROPS_JSON key
LibDTS_DriverCAD.SaveXData frame, entity, useJSON:=True
```

### 4. Validate XData After Write

```vba
' Write
LibDTS_DriverCAD.SaveXData frameObj, entity

' Read back and validate
Dim xdataCheck As Object
Set xdataCheck = LibDTS_DriverCAD.ReadXData(entity)

If xdataCheck.Exists("DTS_GUID") And _
   xdataCheck("DTS_GUID") = frameObj.Base.guid Then
    Debug.Print "XData validation: OK"
Else
    Debug.Print "XData validation: FAILED"
End If
```

---

## Appendix: Complete Example

### Write Frame with XData

```vba
Public Sub Example_WriteFrameWithXData()
    ' Get AutoCAD
    Dim acadApp As Object
    Set acadApp = GetObject(, "AutoCAD.Application")
    Dim acadDoc As Object
    Set acadDoc = acadApp.ActiveDocument
    
    ' Create frame object
    Dim frame As New clsDTSFrame
    frame.StartPoint.Init 0, 0, 0
    frame.EndPoint.Init 5000, 0, 0
    frame.sectionName = "W12X26"
    frame.Base.layer = "DTS_FRAMES"
    frame.Base.color = 1
    
    ' Add custom properties
    frame.Base.Properties.Add "material", "A992"
    frame.Base.Properties.Add "fireRating", "2HR"
    
    ' Draw frame (includes XData save)
    Dim lineObj As Object
    Set lineObj = LibDTS_DriverCAD.DrawFrame(frame, acadDoc)
    
    Debug.Print "Frame drawn with GUID: " & frame.Base.guid
End Sub
```

### Read Frame from XData

```vba
Public Sub Example_ReadFrameFromXData()
    ' Select entity in AutoCAD
    Dim acadDoc As Object
    Set acadDoc = GetObject(, "AutoCAD.Application").ActiveDocument
    
    ' Get selection
    Dim selSet As Object
    Set selSet = acadDoc.SelectionSets.Add("TEMP_SEL")
    selSet.SelectOnScreen
    
    If selSet.Count > 0 Then
        Dim ent As Object
        Set ent = selSet.Item(0)
        
        ' Read XData
        Dim xdata As Object
        Set xdata = LibDTS_DriverCAD.ReadXData(ent)
        
        Debug.Print "Schema Version: " & xdata("SCHEMA_VER")
        Debug.Print "GUID: " & xdata("DTS_GUID")
        Debug.Print "Section: " & xdata("SECTION")
        
        ' If JSON properties exist
        If xdata.Exists("PROPS_JSON") Then
            Dim props As Object
            Set props = LibDTS_Base.ParseJson(xdata("PROPS_JSON"))
            
            Debug.Print "Material: " & props("material")
            Debug.Print "Fire Rating: " & props("fireRating")
        End If
        
        ' Check if legacy
        If xdata.Exists("LEGACY_FLAG") Then
            Debug.Print "WARNING: This is legacy XData. Consider upgrading."
        End If
    End If
    
    selSet.Delete
End Sub
```

---

**Document Version:** 1.0  
**Last Updated:** 2025-11-23  
**Schema Version:** 2.0
