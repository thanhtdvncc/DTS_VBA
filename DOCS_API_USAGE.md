# DTS Core System - API Documentation

## Table of Contents
1. [System Overview](#system-overview)
2. [Architecture](#architecture)
3. [Getting Started](#getting-started)
4. [Core Classes](#core-classes)
5. [Driver Modules](#driver-modules)
6. [Repository Pattern](#repository-pattern)
7. [Usage Examples](#usage-examples)
8. [Best Practices](#best-practices)

---

## System Overview

The DTS Core System is a VBA-based framework for managing structural engineering elements in AutoCAD with SAP2000 integration. It implements Clean Architecture principles with:

- **Pure Domain Models**: No external dependencies
- **Self-Healing Identity**: Automatic GUID regeneration on copy detection
- **Encrypted Persistence**: JSON data stored in XData with encryption
- **2-Way Sync**: Bidirectional synchronization with SAP2000

### Key Features

✅ Hexagonal Architecture (Ports & Adapters)
✅ Object-Oriented Design with Composition
✅ Strong typing with Enums
✅ Comprehensive error handling
✅ Memory leak prevention
✅ JSON serialization with encryption

---

## Architecture

### Layer Structure

```
┌─────────────────────────────────────┐
│     User Interface Layer (Forms)     │
└──────────────┬──────────────────────┘
               │
┌──────────────▼──────────────────────┐
│    Repository Layer (Orchestration)  │
│       clsDTSRepository               │
└──────────────┬──────────────────────┘
               │
    ┌──────────┼──────────┐
    │          │           │
┌───▼──┐  ┌───▼──┐  ┌────▼────┐
│ CAD  │  │ SAP  │  │ DB      │  Driver Layer
│Driver│  │Driver│  │ Driver  │  (External APIs)
└───┬──┘  └───┬──┘  └────┬────┘
    │          │           │
┌───▼──────────▼───────────▼──┐
│    Core Domain Layer         │
│  (Element, Frame, Area, etc) │
└──────────────┬───────────────┘
               │
┌──────────────▼───────────────┐
│    Logic Layer (Pure Math)   │
│  (Geometry, RebarAlgo, Base) │
└──────────────────────────────┘
```

### Module Categories

#### Infrastructure
- `LibDTS_Global` - Constants, enums, singletons
- `LibDTS_Base` - GUID generation, JSON utilities
- `LibDTS_Security` - Encryption/decryption
- `LibDTS_Logger` - File logging
- `LibDTS_Bootstrap` - System initialization
- `clsDTSConfig` - Settings management

#### Core Domain (Class Modules)
- `clsDTSElement` - Base class with identity management
- `clsDTSPoint` - 3D geometry primitive
- `clsDTSSection` - Section properties
- `clsDTSFrame` - Beam/column elements
- `clsDTSArea` - Slab/wall elements
- `clsDTSTag` - Annotations
- `clsDTSRebar` - Reinforcement

#### Logic (Pure Algorithms)
- `LibDTS_RebarAlgo` - Rebar calculations
- `LibDTS_Geometry` - Geometric operations

#### Drivers (External APIs)
- `LibDTS_DriverCAD` - AutoCAD integration
- `LibDTS_DriverSAP` - SAP2000 integration
- `LibDTS_DriverDB` - Database operations

#### Repository
- `clsDTSRepository` - CRUD orchestrator

---

## Getting Started

### 1. Initialize the System

```vba
Sub InitializeDTS()
    ' Initialize system (creates folders, loads config)
    LibDTS_Bootstrap.InitializeSystem
    
    ' Get configuration
    Dim cfg As clsDTSConfig
    Set cfg = LibDTS_Global.Config
    
    ' Verify initialization
    Debug.Print "DTS System Initialized - Version: " & LibDTS_Global.DTS_VERSION
End Sub
```

### 2. Create a Repository Instance

```vba
Sub SetupRepository()
    ' Get AutoCAD document
    Dim acadDoc As Object
    Set acadDoc = ThisDrawing ' Or GetObject(, "AutoCAD.Application").ActiveDocument
    
    ' Create repository
    Dim repo As New clsDTSRepository
    repo.Initialize acadDoc
    
    Debug.Print "Repository ready"
End Sub
```

---

## Core Classes

### clsDTSElement (Base Class)

All structural elements inherit from this base class via composition.

**Properties:**
- `GUID` (Read-only) - Unique identifier
- `OwnerHandle` - CAD entity handle
- `ElementType` - Type enum (FRAME, AREA, etc.)
- `Layer` - AutoCAD layer name
- `Properties` - Dictionary for custom attributes
- `IsDirty` - Tracks if save is needed

**Methods:**
- `ValidateIdentity(handle)` - Self-healing check
- `SerializeBase()` - Returns dictionary
- `DeserializeBase(dict)` - Loads from dictionary
- `MarkAsClean()` - Clear dirty flag
- `MarkAsDirty()` - Set dirty flag

### clsDTSPoint (Geometry Primitive)

Lightweight 3D point for geometric calculations. **Does not inherit from Element.**

**Properties:**
- `X`, `Y`, `Z` - Coordinates

**Methods:**
```vba
pt.Init(x, y, z)              ' Initialize
dist = pt.DistanceTo(otherPt) ' Calculate distance
newPt = pt.Offset(dx, dy, dz) ' Create offset point
isEqual = pt.IsEqual(otherPt) ' Check equality
clonePt = pt.Clone()          ' Deep copy
arr = pt.ToArray()            ' For serialization
```

### clsDTSFrame (Structural Element)

Represents beams and columns.

**Properties:**
- `Base` - clsDTSElement
- `StartPoint` - clsDTSPoint
- `EndPoint` - clsDTSPoint
- `Section` - clsDTSSection
- `FrameType` - Enum (BEAM, COLUMN, etc.)
- `Length` (Read-only) - Calculated

**Methods:**
```vba
frame.IsValid()         ' Check validity
midPt = frame.GetMidPoint()  ' Get center point
json = frame.ToJson()        ' Serialize (encrypted)
frame.FromXData(json)        ' Deserialize
```

### clsDTSArea (Slab/Wall Element)

Represents 2D area elements.

**Properties:**
- `Base` - clsDTSElement
- `BoundaryPoints` - Collection of clsDTSPoint
- `Thickness` - Element thickness (mm)
- `Material` - Material name
- `AreaType` - "SLAB", "WALL", "FOUNDATION"
- `LoadBearing` - Boolean

**Methods:**
```vba
area.AddBoundaryPoint(pt)     ' Add point to boundary
area.ClearBoundary()          ' Remove all points
area.IsValid()                ' Check validity (>= 3 points)
area.IsClosed()               ' Check if closed
calcArea = area.CalculateArea()      ' Calculate area (m²)
perim = area.CalculatePerimeter()    ' Calculate perimeter
centroid = area.GetCentroid()        ' Get center point
clone = area.Clone()                 ' Deep copy
json = area.ToJson()                 ' Serialize
```

### clsDTSTag (Annotation)

Text annotations linked to host elements.

**Properties:**
- `Base` - clsDTSElement
- `HostGUID` - Parent element GUID
- `TextContent` - Annotation text
- `Position` - clsDTSPoint
- `Rotation` - Radians
- `Height` - Text height
- `Style` - Text style name

**Methods:**
```vba
tag.IsValid()           ' Check if text exists
tag.HasHost()           ' Check if linked
json = tag.ToJson()     ' Serialize
clone = tag.Clone()     ' Deep copy
```

### clsDTSRebar (Reinforcement)

Represents steel reinforcement.

**Properties:**
- `Base` - clsDTSElement
- `Diameter` - Bar diameter (mm)
- `Quantity` - Number of bars
- `Spacing` - Center-to-center spacing (mm)
- `Mark` - Bar mark/number
- `RebarClass` - Steel grade (e.g., "CB400V")
- `HostGUID` - Parent element GUID
- `Shape` - clsDTSRebarShape

**Methods:**
```vba
weight = rebar.GetTotalWeight()  ' Calculate weight (kg)
json = rebar.ToJson()            ' Serialize
```

---

## Driver Modules

### LibDTS_DriverCAD

AutoCAD integration functions.

**Drawing Functions:**
```vba
' Draw a frame (returns Line entity)
Set lineObj = LibDTS_DriverCAD.DrawFrame(frame, acadDoc, dryRun:=False)

' Draw an area (returns Polyline entity)
Set polyObj = LibDTS_DriverCAD.DrawArea(area, acadDoc, dryRun:=False)

' Draw a tag (returns Text entity)
Set textObj = LibDTS_DriverCAD.DrawTag(tag, acadDoc, dryRun:=False)
```

**Reading Functions:**
```vba
' Read frame from Line entity
Set frame = LibDTS_DriverCAD.ReadFrame(lineEntity)

' Read area from Polyline entity
Set area = LibDTS_DriverCAD.ReadArea(polyEntity)

' Check if entity has XData
hasData = LibDTS_DriverCAD.HasXData(entity)
```

**XData Operations:**
```vba
' Save XData to entity (automatic serialization & encryption)
LibDTS_DriverCAD.SaveXData frame, lineEntity

' Read raw XData
jsonStr = LibDTS_DriverCAD.ReadXData(entity)
```

### LibDTS_DriverSAP

SAP2000 integration functions.

**Connection Management:**
```vba
' Connect to SAP2000
success = LibDTS_DriverSAP.Connect(version:="auto", startNew:=False)

' Disconnect
LibDTS_DriverSAP.Disconnect saveModel:=True
```

**Data Transfer:**
```vba
' Push frame to SAP2000
success = LibDTS_DriverSAP.PushFrame(frame)

' Sync GUID mapping
LibDTS_DriverSAP.SyncGUID sapFrameName, dtsGUID
```

### LibDTS_DriverDB

Database and settings persistence.

**Settings Management:**
```vba
' Load settings
Set settings = LibDTS_DriverDB.LoadSettings()

' Save settings
LibDTS_DriverDB.SaveSettings settingsDict
```

---

## Repository Pattern

The `clsDTSRepository` class orchestrates all CRUD operations.

### Initialization

```vba
Dim repo As New clsDTSRepository
repo.Initialize acadDoc
```

### Saving Elements

```vba
' Save a frame (automatically draws if new, updates if exists)
success = repo.Save(frame)

' Save with explicit entity
success = repo.Save(frame, existingLineEntity)

' Save an area
success = repo.Save(area)

' Save a tag
success = repo.Save(tag)
```

### Loading Elements

```vba
' Load by CAD handle
Set element = repo.LoadByHandle("ABC123")

' Load by GUID
Set element = repo.LoadByGUID("{guid-string}")

' Load from entity
Set element = repo.LoadFromEntity(lineEntity)
```

### SAP2000 Synchronization

```vba
' Sync single element to SAP2000
success = repo.SyncToSAP(frame)

' Check connection status
If repo.SAPConnected Then
    Debug.Print "SAP2000 connected"
End If
```

---

## Usage Examples

### Example 1: Create and Save a Frame

```vba
Sub CreateBeam()
    ' Initialize repository
    Dim repo As New clsDTSRepository
    repo.Initialize ThisDrawing
    
    ' Create frame
    Dim beam As New clsDTSFrame
    beam.Base.Layer = "DTS_BEAM"
    
    ' Set geometry
    Dim pt1 As New clsDTSPoint
    Dim pt2 As New clsDTSPoint
    pt1.Init 0, 0, 0
    pt2.Init 5000, 0, 0
    Set beam.StartPoint = pt1
    Set beam.EndPoint = pt2
    
    ' Set section
    beam.Section.Name = "B300x600"
    beam.Section.Width = 300
    beam.Section.Depth = 600
    beam.Section.Material = "C30"
    
    ' Save to CAD
    If repo.Save(beam) Then
        Debug.Print "Beam created: " & beam.Base.GUID
    End If
End Sub
```

### Example 2: Create a Slab Area

```vba
Sub CreateSlab()
    Dim repo As New clsDTSRepository
    repo.Initialize ThisDrawing
    
    ' Create area
    Dim slab As New clsDTSArea
    slab.Base.Layer = "DTS_SLAB"
    slab.AreaType = "SLAB"
    slab.Thickness = 150
    slab.Material = "C30"
    
    ' Define boundary (rectangle)
    Dim pt1 As New clsDTSPoint: pt1.Init 0, 0, 0
    Dim pt2 As New clsDTSPoint: pt2.Init 6000, 0, 0
    Dim pt3 As New clsDTSPoint: pt3.Init 6000, 4000, 0
    Dim pt4 As New clsDTSPoint: pt4.Init 0, 4000, 0
    
    slab.AddBoundaryPoint pt1
    slab.AddBoundaryPoint pt2
    slab.AddBoundaryPoint pt3
    slab.AddBoundaryPoint pt4
    slab.AddBoundaryPoint pt1 ' Close polygon
    
    ' Calculate properties
    Debug.Print "Area: " & slab.CalculateArea() & " mm²"
    Debug.Print "Perimeter: " & slab.CalculatePerimeter() & " mm"
    
    ' Save
    If repo.Save(slab) Then
        Debug.Print "Slab created: " & slab.Base.GUID
    End If
End Sub
```

### Example 3: Add Rebar with Calculations

```vba
Sub AddRebarToBeam()
    ' Create rebar
    Dim rebar As New clsDTSRebar
    rebar.Base.Layer = "DTS_REBAR_MAIN"
    rebar.Diameter = 20
    rebar.Quantity = 4
    rebar.Spacing = 200
    rebar.Mark = "T1"
    rebar.RebarClass = "CB400V"
    
    ' Set shape (stirrup)
    Dim dims As Object
    Set dims = CreateObject("Scripting.Dictionary")
    dims.Add "Width", 250
    dims.Add "Height", 550
    
    ' Calculate cut length
    Dim cutLength As Double
    cutLength = LibDTS_RebarAlgo.CalculateCutLength(51, dims, rebar.Diameter)
    Debug.Print "Cut length: " & cutLength & " mm"
    
    ' Calculate weight
    Dim weight As Double
    weight = LibDTS_RebarAlgo.CalculateWeight(rebar.Diameter, cutLength, rebar.Quantity)
    Debug.Print "Total weight: " & weight & " kg"
End Sub
```

### Example 4: Sync to SAP2000

```vba
Sub SyncFrameToSAP()
    Dim repo As New clsDTSRepository
    repo.Initialize ThisDrawing
    
    ' Load existing frame by handle
    Dim handle As String
    handle = "ABC123" ' Get from selection
    
    Set frame = repo.LoadByHandle(handle)
    
    If Not frame Is Nothing Then
        ' Sync to SAP2000
        If repo.SyncToSAP(frame) Then
            Debug.Print "Frame synced to SAP2000"
        Else
            Debug.Print "Sync failed: " & repo.LastError
        End If
    End If
End Sub
```

### Example 5: Batch Processing

```vba
Sub ProcessAllFrames()
    Dim repo As New clsDTSRepository
    repo.Initialize ThisDrawing
    
    ' Get all line entities
    Dim entity As Object
    For Each entity In ThisDrawing.ModelSpace
        If TypeName(entity) = "AcDbLine" Then
            ' Try to load as frame
            Dim frame As clsDTSFrame
            Set frame = repo.LoadFromEntity(entity)
            
            If Not frame Is Nothing Then
                ' Process frame
                Debug.Print "Frame: " & frame.Base.GUID
                Debug.Print "  Length: " & frame.Length & " mm"
                Debug.Print "  Layer: " & frame.Base.Layer
                
                ' Example: Update section
                frame.Section.Material = "C35"
                frame.Base.MarkAsDirty
                repo.Save frame
            End If
        End If
    Next entity
End Sub
```

---

## Best Practices

### 1. Always Initialize the System

```vba
' At application startup
LibDTS_Bootstrap.InitializeSystem
```

### 2. Use Repository for All CRUD Operations

❌ Don't bypass the repository:
```vba
' BAD - Direct driver access loses benefits
Set line = LibDTS_DriverCAD.DrawFrame(frame, doc)
```

✅ Use repository:
```vba
' GOOD - Repository manages everything
success = repo.Save(frame)
```

### 3. Check for Errors

```vba
If Not repo.Save(frame) Then
    Debug.Print "Error: " & repo.LastError
    ' Handle error
End If
```

### 4. Handle Self-Healing

The system automatically handles copied elements:

```vba
' When loading, identity is validated automatically
Set frame = repo.LoadByHandle(handle)
' If handle mismatch detected, new GUID generated automatically
```

### 5. Use Strong Typing

```vba
' Good - Strong typing
Dim frame As clsDTSFrame
Set frame = New clsDTSFrame

' Avoid - Late binding
Dim frame As Object
Set frame = New clsDTSFrame
```

### 6. Clean Up Resources

```vba
Sub CleanupExample()
    Dim frame As clsDTSFrame
    Set frame = New clsDTSFrame
    
    ' Do work...
    
    ' Clean up (automatic, but explicit is good practice)
    Set frame = Nothing
End Sub
```

### 7. Validate Before Saving

```vba
If frame.IsValid() Then
    repo.Save frame
Else
    Debug.Print "Frame is invalid"
End If
```

### 8. Use Configuration

```vba
' Access global config
Dim cfg As clsDTSConfig
Set cfg = LibDTS_Global.Config

' Get layer names from config
Dim beamLayer As String
beamLayer = cfg.GetVal("Layer_Beam", "DTS_BEAM")

' Update config
cfg.SetVal "Custom_Setting", "Value"
cfg.Save ' Persist to disk
```

### 9. Leverage Pure Algorithms

```vba
' Rebar calculations are pure functions
Dim spacing As Double
spacing = LibDTS_RebarAlgo.CalculateActualSpacing(5000, 25, 50, 50)

Dim positions As Collection
Set positions = LibDTS_RebarAlgo.GenerateBarPositions(5000, 200, 100)
```

### 10. Log Important Operations

```vba
' Logging happens automatically in drivers
' Manual logging when needed
LibDTS_Logger.Log "Custom operation completed", DTS_INFO
LibDTS_Logger.Log "Warning message", DTS_WARNING
LibDTS_Logger.Log "Error details", DTS_ERROR
```

---

## Performance Tips

1. **Batch Operations**: Group multiple saves together
2. **Dry Run Mode**: Test operations without committing
3. **Lazy Loading**: Only load data when needed
4. **Cache Results**: Store frequently accessed data
5. **Minimize CAD Redraws**: Use transactions when available

---

## Troubleshooting

### Issue: "Type Mismatch" Error

**Cause**: Missing class module or incorrect type
**Solution**: Ensure all class modules are imported

### Issue: XData Not Persisting

**Cause**: Application not registered
**Solution**: Check XDATA_APPNAME constant, ensure RegApp exists

### Issue: Self-Healing Not Working

**Cause**: Handle not passed to ValidateIdentity
**Solution**: Repository handles this automatically - use repo.Load methods

### Issue: SAP2000 Connection Failed

**Cause**: SAP not running or version mismatch
**Solution**: 
```vba
' Check version
success = LibDTS_DriverSAP.Connect("auto", startNew:=True)
```

---

## Version History

**2.0.0** (Current)
- Complete Big Bang rewrite
- Hexagonal architecture
- Self-healing identity
- Encrypted persistence
- Full SAP2000 sync
- Comprehensive rebar algorithms

---

## Support & Contact

For issues and questions:
- Review log files in %TEMP%\DTS_System.log
- Check configuration in %APPDATA%\DTS_Core\settings.json
- Consult the README files in the repository

---

*DTS Core System - Professional Structural Engineering Automation*
