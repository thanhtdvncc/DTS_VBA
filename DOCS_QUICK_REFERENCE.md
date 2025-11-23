# DTS Core System - Quick Reference Card

## System Initialization

```vba
' 1. Initialize system
LibDTS_Bootstrap.InitializeSystem

' 2. Create repository
Dim repo As New clsDTSRepository
repo.Initialize ThisDrawing
```

## Common Operations

### Create Frame (Beam/Column)

```vba
Dim frame As New clsDTSFrame
frame.Base.Layer = "DTS_BEAM"

Dim pt1 As New clsDTSPoint
Dim pt2 As New clsDTSPoint
pt1.Init 0, 0, 0
pt2.Init 5000, 0, 0
Set frame.StartPoint = pt1
Set frame.EndPoint = pt2

frame.Section.Width = 300
frame.Section.Depth = 600

repo.Save frame
```

### Create Area (Slab/Wall)

```vba
Dim area As New clsDTSArea
area.AreaType = "SLAB"
area.Thickness = 150

Dim pt1 As New clsDTSPoint: pt1.Init 0, 0, 0
Dim pt2 As New clsDTSPoint: pt2.Init 6000, 0, 0
Dim pt3 As New clsDTSPoint: pt3.Init 6000, 4000, 0
Dim pt4 As New clsDTSPoint: pt4.Init 0, 4000, 0

area.AddBoundaryPoint pt1
area.AddBoundaryPoint pt2
area.AddBoundaryPoint pt3
area.AddBoundaryPoint pt4

repo.Save area
```

### Create Tag (Annotation)

```vba
Dim tag As New clsDTSTag
tag.TextContent = "B1"
tag.Height = 2.5
tag.Position.Init 1000, 500, 0
tag.HostGUID = frame.Base.GUID

repo.Save tag
```

### Create Rebar

```vba
Dim rebar As New clsDTSRebar
rebar.Diameter = 20
rebar.Quantity = 4
rebar.Spacing = 200
rebar.Mark = "T1"

repo.Save rebar
```

## Load Operations

```vba
' Load by handle
Set element = repo.LoadByHandle("ABC123")

' Load by GUID
Set element = repo.LoadByGUID("{guid}")

' Load from entity
Set element = repo.LoadFromEntity(lineEntity)
```

## Rebar Calculations

```vba
' Calculate cut length for stirrup
Dim dims As Object
Set dims = CreateObject("Scripting.Dictionary")
dims.Add "Width", 250
dims.Add "Height", 550

cutLen = LibDTS_RebarAlgo.CalculateCutLength(51, dims, 12)

' Calculate weight
weight = LibDTS_RebarAlgo.CalculateWeight(diameter, length, quantity)

' Calculate bar count from spacing
count = LibDTS_RebarAlgo.CalculateBarCount(totalLength, spacing)

' Generate bar positions
Set positions = LibDTS_RebarAlgo.GenerateBarPositions(5000, 200, 50)
```

## SAP2000 Sync

```vba
' Connect to SAP2000
LibDTS_DriverSAP.Connect

' Sync frame
repo.SyncToSAP frame

' Disconnect
LibDTS_DriverSAP.Disconnect saveModel:=True
```

## Geometry Operations

```vba
' Point operations
dist = pt1.DistanceTo(pt2)
newPt = pt1.Offset(100, 200, 0)
isEqual = pt1.IsEqual(pt2)
clonePt = pt1.Clone()

' Frame operations
length = frame.Length
midPt = frame.GetMidPoint()
isValid = frame.IsValid()

' Area operations
area = area.CalculateArea()
perimeter = area.CalculatePerimeter()
centroid = area.GetCentroid()
isClosed = area.IsClosed()
```

## Configuration

```vba
' Get config
Dim cfg As clsDTSConfig
Set cfg = LibDTS_Global.Config

' Read setting
value = cfg.GetVal("Layer_Beam", "DTS_BEAM")

' Write setting
cfg.SetVal "Custom_Key", "Custom_Value"
cfg.Save
```

## Constants & Enums

### Element Types
```vba
DTS_ELEM_UNKNOWN = 0
DTS_ELEM_FRAME = 1
DTS_ELEM_AREA = 2
DTS_ELEM_NODE = 3
DTS_ELEM_ANNOTATION = 4
DTS_ELEM_REBAR = 5
```

### Frame Types
```vba
DTS_FRM_BEAM = 1
DTS_FRM_COLUMN = 2
DTS_FRM_BRACE = 3
DTS_FRM_PILE = 4
```

### Shape Types
```vba
DTS_SHP_RECTANGLE = 1
DTS_SHP_CIRCLE = 2
DTS_SHP_I_SECTION = 3
DTS_SHP_T_SECTION = 4
DTS_SHP_L_SECTION = 5
```

### Rebar Shape Codes
```vba
DTS_RBR_00 = 0   ' Unknown
DTS_RBR_01 = 1   ' Straight
DTS_RBR_02 = 2   ' Straight with hooks
DTS_RBR_18 = 18  ' U-Shape
DTS_RBR_51 = 51  ' Stirrup
```

### Log Types
```vba
DTS_INFO = 0
DTS_WARNING = 1
DTS_ERROR = 2
```

## Error Handling

```vba
On Error GoTo ErrHandler

' Your code here

Exit Sub

ErrHandler:
    LibDTS_Logger.Log "Error: " & Err.Description, DTS_ERROR
    MsgBox "Operation failed: " & Err.Description
```

## Validation Checks

```vba
' Check if element is valid
If frame.IsValid() Then
    ' Process
End If

' Check if area is closed
If area.IsClosed() Then
    ' Calculate properties
End If

' Check if tag has host
If tag.HasHost() Then
    ' Link to parent
End If
```

## Serialization

```vba
' Serialize to JSON (encrypted)
jsonStr = frame.ToJson()

' Deserialize from JSON
frame.FromXData jsonStr

' Clone object (deep copy)
Set newFrame = frame.Clone()
```

## CAD Operations

```vba
' Draw elements
Set line = LibDTS_DriverCAD.DrawFrame(frame, acadDoc)
Set poly = LibDTS_DriverCAD.DrawArea(area, acadDoc)
Set text = LibDTS_DriverCAD.DrawTag(tag, acadDoc)

' Read elements
Set frame = LibDTS_DriverCAD.ReadFrame(lineEntity)
Set area = LibDTS_DriverCAD.ReadArea(polyEntity)

' Check XData
hasData = LibDTS_DriverCAD.HasXData(entity)
```

## Standard Rebar Sizes (mm)

```
6, 8, 10, 12, 16, 20, 25, 32, 40
```

## Standard Concrete Cover (mm)

```vba
' Beams: 25mm
' Columns: 40mm  
' Slabs: 20mm
```

## Common Formulas

### Rebar Weight
```
Weight (kg) = (D² / 162) × L(m) × Quantity
where D = diameter (mm), L = length (meters)
```

### Development Length (Simplified)
```
Ld = 1.25 × (fy / √fc') × diameter
where fy = steel grade, fc' = concrete grade
```

## File Locations

```
Logs:     %TEMP%\DTS_System.log
Settings: %APPDATA%\DTS_Core\settings.json
```

## System Constants

```vba
DTS_APP_NAME = "DTS_CORE_DATA"
DTS_VERSION = "2.0.0"
DTS_PRECISION = 0.0001
DTS_XDATA_APPNAME = "DTS_CORE"
```

## Tips & Tricks

1. **Always use repository** - Don't bypass repo.Save()
2. **Check IsValid()** before saving
3. **Handle errors** with On Error GoTo
4. **Clean up objects** with Set obj = Nothing
5. **Use enums** instead of magic numbers
6. **Validate spacing** with LibDTS_RebarAlgo.ValidateSpacing()
7. **Round spacing** with LibDTS_RebarAlgo.RoundSpacing()
8. **Log important operations** with LibDTS_Logger.Log()

## Performance Notes

- Points are lightweight (no Base element)
- Frames, Areas, Tags, Rebar have full persistence
- Use dry-run mode for validation: `DrawFrame(frame, doc, dryRun:=True)`
- Batch saves when possible
- Clear collections when done

## Common Pitfalls to Avoid

❌ `frame.StartPoint.Init(x, y, z)` - Won't work, StartPoint is read-only
✅ Create point first: `Set frame.StartPoint = pt`

❌ `pt.Base.GUID` - Points don't have Base
✅ Points are pure geometry primitives

❌ Direct XData manipulation
✅ Use Repository: `repo.Save(element)`

❌ Hardcoded values
✅ Use config: `cfg.GetVal("Layer_Beam")`

---

*Quick Reference v2.0 - For detailed documentation see DOCS_API_USAGE.md*
