# DTS_VBA API Consolidation - Inventory Report

## Executive Summary

This document provides a comprehensive inventory of all public APIs in the DTS_VBA repository, grouped by behavior and domain. The purpose is to identify duplicate implementations and consolidate them into canonical driver modules.

**Repository Statistics:**
- Total VBA Files: 56 (bas/cls)
- Total Public APIs: 200+
- Driver Modules: 3 (LibDTS_DriverSAP, LibDTS_DriverCAD, LibDTS_DriverDB)
- Core Classes: 10 (clsDTS*)
- Legacy Modules: 20+ (m*.bas, n*.bas)

## Duplicate API Groups

### 1. SAP2000 Connection Management (4 implementations)

**Behavior:** Connect to SAP2000 instance, manage connection lifecycle

**Implementations:**
1. `LibDTS_DriverSAP.bas` - `Public Function Connect() As Boolean` (Line 14)
   - Modern driver approach, late-binding
   - Initializes new model with units
   
2. `m01_SAP2000_Connection.bas` - `Public Function ConnectSAP2000() As Boolean` (Line 8)
   - Legacy approach with helper object version detection
   - More robust fallback chain through multiple SAP versions
   - Sets present units to kN_mm_C
   
3. `m01_SAP2000_Connection.bas` - `Public Sub DisconnectSAP2000()` (Line 69)
   - Legacy disconnect with optional message
   
4. `Core_Utils.bas` - `Public Function IsSAPConnected(SapModel As Object) As Boolean` (Line 1635)
   - Validation helper, checks if connection exists

**Recommendation:** Consolidate into `LibDTS_DriverSAP` with enhanced version detection and error handling

---

### 2. CAD Entity Reading (7 implementations)

**Behavior:** Read entities from AutoCAD drawing and extract data

**Implementations:**
1. `LibDTS_DriverCAD.bas` - `Public Function ReadFrame(ent As Object) As clsDTSFrame` (Line 95)
   - Reads single Frame entity with XData
   - Returns typed clsDTSFrame object
   
2. `Core_XData_Reader.bas` - `Public Function ReadPointsFromCAD()` (Line 49)
   - Batch read all Points from drawing
   - Returns array of CADPoint UDTs
   
3. `Core_XData_Reader.bas` - `Public Function ReadFramesFromCAD()` (Line 90)
   - Batch read all Frames from drawing
   - Returns array of CADFrame UDTs
   
4. `Core_XData_Reader.bas` - `Public Function ReadAreasFromCAD()` (Line 131)
   - Batch read all Areas from drawing
   - Returns array of CADArea UDTs
   
5. `m04_SAP2000_Joints_Frames.bas` - `Public Sub ExtractPoints()` (Line 8)
   - Legacy SAP data extraction to global arrays
   
6. `m04_SAP2000_Joints_Frames.bas` - `Public Sub ExtractFrames()` (Line 33)
   - Legacy SAP frame extraction to global arrays
   
7. `m05_SAP2000_Areas.bas` - `Public Sub ExtractAreas()` (Line 8)
   - Legacy SAP area extraction to global arrays

**Recommendation:** Consolidate into `LibDTS_DriverCAD` with unified ReadEntity pattern
- Add ReadPoint, ReadFrame, ReadArea (single entity)
- Add ReadAllPoints, ReadAllFrames, ReadAllAreas (batch operations)
- Return Core objects (clsDTS*) instead of UDTs where possible

---

### 3. CAD Entity Drawing (4 implementations)

**Behavior:** Create/draw entities in AutoCAD with metadata

**Implementations:**
1. `LibDTS_DriverCAD.bas` - `Public Function DrawFrame()` (Line 11)
   - Draws single Frame from clsDTSFrame object
   - Saves XData automatically
   
2. `LibDTS_DriverCAD.bas` - `Public Function DrawTag()` (Line 40)
   - Draws text annotation
   - Links to host via XData
   
3. `Core_CAD_Plotter.bas` - `Public Function PlotPoints()` (Line 49)
   - Batch plot points with labels
   - Lower-level implementation
   
4. `n00_ACAD_Main_Integration.bas` - `Public Sub PlotModelToNewDrawing()` (Line 14)
   - High-level orchestrator
   - Combines connection, sync, and plotting

**Recommendation:** Keep high-level orchestrators, consolidate low-level drawing into `LibDTS_DriverCAD`
- Add DrawPoint, DrawFrame, DrawArea, DrawTag
- All should accept Core objects and options (dryRun, layer, color)

---

### 4. XData Operations (6 implementations)

**Behavior:** Read/Write extended data to CAD entities

**Implementations:**
1. `LibDTS_DriverCAD.bas` - `Public Sub SaveXData(dtsObj, ent)` (Line 64)
   - Generic save for any DTS object
   - Registers app, serializes, writes XData
   
2. `LibDTS_DriverCAD.bas` - Private `GetRawXData()` (Line 125)
   - Helper to extract raw XData string
   
3. `clsDTSElement.cls` - `Public Function SerializeBase()` (Line ?)
   - Serializes base properties to JSON
   
4. `clsDTSElement.cls` - `Public Sub DeserializeBase()` (Line ?)
   - Deserializes base properties from JSON
   
5. `clsDTSFrame.cls` - `Public Sub FromXData()` (Line ?)
   - Specialized Frame deserialization
   
6. `clsDTSRebar.cls` - `Public Sub FromXData()` (Line ?)
   - Specialized Rebar deserialization

**Recommendation:** Keep current structure with enhancements
- Ensure consistent app name registration
- Add validation and versioning to XData format
- Implement XData migration for schema changes

---

### 5. SAP2000 Push Operations (3+ implementations)

**Behavior:** Create entities in SAP2000 model

**Implementations:**
1. `LibDTS_DriverSAP.bas` - `Public Function PushFrame()` (Line 37)
   - Creates Frame in SAP from clsDTSFrame
   - Handles point creation automatically
   
2. Legacy modules in `m*.bas` files create points/frames directly via SAP API
   - Scattered throughout m04, m05, etc.
   
**Recommendation:** Consolidate into `LibDTS_DriverSAP`
- Add PushPoint, PushFrame, PushArea, PushRebar
- Add batch operations (PushFrames, PushAreas)
- Implement dryRun mode for validation

---

### 6. Sync Operations (2 implementations)

**Behavior:** Bi-directional sync between CAD and SAP

**Implementations:**
1. `Core_Sync_Manager.bas` - `Public Sub SyncSAPToCADWithFilters()` (Line 235)
   - SAP → CAD with filtering rules
   - Complex implementation with callbacks
   
2. `Core_Sync_Manager.bas` - `Public Sub SyncCADToSAP()` (Line 1993)
   - CAD → SAP with tolerance and scaling
   - Handles point merging and mapping

**Recommendation:** Keep in Core_Sync_Manager but refactor to use driver APIs
- Delegate connection to drivers
- Delegate entity creation to drivers
- Focus on orchestration and conflict resolution

---

### 7. Database Operations (3 implementations)

**Behavior:** Load/Save settings and data to persistent storage

**Implementations:**
1. `LibDTS_DriverDB.bas` - `Public Function LoadSettings()` (Line 7)
   - Loads JSON settings from AppData
   - Returns Dictionary object
   
2. `LibDTS_DriverDB.bas` - `Public Sub SaveSettings()` (Line 29)
   - Saves Dictionary to JSON file
   
3. `m11_SAP2000_DataBase.bas` - `Public Sub ExportTableToActiveSheet()` (Line 294)
   - Exports SAP table to Excel
   - Different purpose but overlapping storage concerns

**Recommendation:** Extend `LibDTS_DriverDB` with more operations
- Add LoadMapping, SaveMapping (GUID ↔ external ID)
- Add LoadConfig, SaveConfig
- Add database versioning and migration support

---

## Mapping Priority Matrix

| Behavior Group | Implementations | Consolidation Target | Priority | Risk |
|----------------|-----------------|---------------------|----------|------|
| SAP Connection | 4 | LibDTS_DriverSAP | HIGH | LOW |
| CAD Read | 7 | LibDTS_DriverCAD | HIGH | MEDIUM |
| CAD Draw | 4 | LibDTS_DriverCAD | HIGH | LOW |
| SAP Push | 3+ | LibDTS_DriverSAP | HIGH | MEDIUM |
| XData Ops | 6 | LibDTS_DriverCAD | MEDIUM | LOW |
| Sync | 2 | Keep Core_Sync_Manager | LOW | HIGH |
| Database | 3 | LibDTS_DriverDB | MEDIUM | LOW |

**Risk Assessment:**
- LOW: Simple delegation, no side effects
- MEDIUM: Behavioral differences, needs careful testing
- HIGH: Complex orchestration, many dependencies

---

## Legacy Module Migration Path

### Phase 1: Critical Path (Week 1)
- m01_SAP2000_Connection → LibDTS_DriverSAP
- m04_SAP2000_Joints_Frames → LibDTS_DriverSAP (read operations)
- Core_XData_Reader → LibDTS_DriverCAD (consolidate read methods)

### Phase 2: CAD Operations (Week 2)
- Core_CAD_Plotter → LibDTS_DriverCAD (drawing methods)
- n00_ACAD_Main_Integration → Update to use drivers

### Phase 3: Data Operations (Week 3)
- m11_SAP2000_DataBase → Extend LibDTS_DriverDB
- Implement GUID mapping persistence

### Phase 4: Testing & Validation (Week 4)
- Smoke tests for each driver
- Backward compatibility validation
- Performance benchmarking

---

## API Surface Design

### LibDTS_DriverSAP (Enhanced)

```vba
' CONNECTION
Public Function Connect(Optional version As String = "auto") As Boolean
Public Function Disconnect() As Boolean
Public Function IsConnected() As Boolean

' MODELING (with dryRun support)
Public Function PushPoint(pt As clsDTSPoint, Optional dryRun As Boolean = False) As String
Public Function PushFrame(frame As clsDTSFrame, Optional dryRun As Boolean = False) As String
Public Function PushArea(area As clsDTSArea, Optional dryRun As Boolean = False) As String
Public Function PushRebar(rebar As clsDTSRebar, Optional dryRun As Boolean = False) As String

' READING
Public Function ReadPoint(pointName As String) As clsDTSPoint
Public Function ReadFrame(frameName As String) As clsDTSFrame
Public Function ReadArea(areaName As String) As clsDTSArea

' GUID MAPPING
Public Function MapGUIDToElement(guid As String, sapName As String, elementType As String) As Boolean
Public Function FindElementByGUID(guid As String) As String
Public Function RemoveGUIDMapping(guid As String) As Boolean

' RESULTS QUERY
Public Function GetFrameForces(frameName As String, loadCase As String) As Dictionary
Public Function GetAreaStresses(areaName As String, loadCase As String) As Dictionary
```

### LibDTS_DriverCAD (Enhanced)

```vba
' DRAWING (with dryRun support)
Public Function DrawPoint(pt As clsDTSPoint, acadDoc As Object, Optional dryRun As Boolean = False) As Object
Public Function DrawFrame(frame As clsDTSFrame, acadDoc As Object, Optional dryRun As Boolean = False) As Object
Public Function DrawArea(area As clsDTSArea, acadDoc As Object, Optional dryRun As Boolean = False) As Object
Public Function DrawTag(tag As clsDTSTag, acadDoc As Object, Optional dryRun As Boolean = False) As Object

' READING
Public Function ReadPoint(ent As Object) As clsDTSPoint
Public Function ReadFrame(ent As Object) As clsDTSFrame
Public Function ReadArea(ent As Object) As clsDTSArea

' BATCH OPERATIONS
Public Function ReadAllPoints(acadDoc As Object) As Collection
Public Function ReadAllFrames(acadDoc As Object) As Collection
Public Function ReadAllAreas(acadDoc As Object) As Collection

' XDATA OPERATIONS
Public Sub SaveXData(dtsObj As Object, ent As Object)
Public Function ReadXData(ent As Object, appName As String) As String
Public Function HasXData(ent As Object, appName As String) As Boolean

' GUID MAPPING
Public Function FindEntityByGUID(acadDoc As Object, guid As String) As Object
Public Function MapGUIDToHandle(guid As String, handle As String) As Boolean
```

### LibDTS_DriverDB (Enhanced)

```vba
' SETTINGS
Public Function LoadSettings() As Object
Public Sub SaveSettings(settingsDict As Object)

' MAPPING PERSISTENCE
Public Function LoadGUIDMapping() As Dictionary
Public Sub SaveGUIDMapping(mappingDict As Dictionary)
Public Function GetMappedElement(guid As String) As Variant
Public Sub SetMappedElement(guid As String, elementInfo As Variant)

' CONFIGURATION
Public Function LoadConfig(Optional configName As String = "default") As clsDTSConfig
Public Sub SaveConfig(config As clsDTSConfig)

' VALIDATION
Public Function ValidateMappingIntegrity() As Collection ' Returns list of broken mappings
Public Sub RepairMapping(guidList As Variant)
```

---

## Missing Dependencies Status

### CRITICAL: JsonConverter.bas
**Status:** ❌ NOT FOUND
**Impact:** HIGH - Referenced by clsDTSFrame, clsDTSRebar, LibDTS_DriverDB
**Action:** Created LibDTS_Base.bas with simple JSON parser as interim solution
**Recommendation:** Add proper JsonConverter library or use native approach

### CRITICAL: LibDTS_Base.bas
**Status:** ✅ CREATED
**Contains:** GUID generation, JSON utilities, validation helpers
**Dependencies:** None (standalone)

### CRITICAL: LibDTS_Security.bas
**Status:** ✅ CREATED
**Contains:** XOR encryption, Base64 encoding for XData protection
**Dependencies:** MSXML2.DOMDocument (optional)

---

## Rollback Procedures

### If Driver Migration Fails:
1. Revert to legacy modules (keep originals with `.legacy` suffix)
2. Restore global variable declarations
3. Re-enable direct API calls

### If XData Format Changes Break:
1. Implement XData version detection
2. Add migration layer for old format
3. Support dual read (new/old format)

### If GUID Mapping Corrupts:
1. Export current mapping to CSV
2. Rebuild from XData scan
3. Manual reconciliation via Excel

---

## Next Steps

1. ✅ Create missing utility modules (LibDTS_Base, LibDTS_Security)
2. 📝 Document existing driver APIs
3. 🔧 Enhance drivers with defensive programming
4. 🧪 Create smoke test suite
5. 🔄 Begin Phase 1 migration (connection APIs)

---

## Appendix: Complete API List

See `/tmp/inventory_report.md` for machine-generated complete API inventory.

---

**Document Version:** 1.0  
**Date:** 2025-11-23  
**Author:** DTS Consolidation Agent  
