Attribute VB_Name = "LibDTS_RebarAlgo"
' Module: LibDTS_RebarAlgo
' Purpose: Pure algorithms for rebar calculations (cut lengths, spacing, geometry)
' This module contains NO external API calls - only mathematics
' Dependencies: None (pure logic)
Option Explicit

' ==========================================
' CONSTANTS
' ==========================================

' Standard bend radii (multiples of diameter)
Public Const REBAR_BEND_RADIUS_90 As Double = 3#     ' 90-degree bend = 3*diameter
Public Const REBAR_BEND_RADIUS_135 As Double = 3#    ' 135-degree bend = 3*diameter
Public Const REBAR_BEND_RADIUS_180 As Double = 5#    ' 180-degree bend (hook) = 5*diameter

' Standard hook extensions
Public Const REBAR_HOOK_EXTENSION As Double = 12#    ' Hook extension = 12*diameter

' Concrete cover constants (mm)
Public Const DEFAULT_COVER_BEAM As Double = 25#
Public Const DEFAULT_COVER_COLUMN As Double = 40#
Public Const DEFAULT_COVER_SLAB As Double = 20#

' ==========================================
' CUT LENGTH CALCULATIONS
' ==========================================

' Calculate cut length for straight bar
' Parameters:
'   length: Required length in mm
'   diameter: Bar diameter in mm
'   hasHooks: True if hooks at both ends
' Returns: Total cut length including hooks
Public Function CalculateStraightBarLength(length As Double, _
                                          diameter As Double, _
                                          Optional hasHooks As Boolean = False) As Double
    On Error GoTo ErrHandler
    
    Dim cutLength As Double
    cutLength = length
    
    ' Add hook extensions if needed
    If hasHooks Then
        Dim hookLength As Double
        hookLength = CalculateHookLength(diameter, 180#)
        cutLength = cutLength + (2 * hookLength)
    End If
    
    CalculateStraightBarLength = cutLength
    Exit Function
    
ErrHandler:
    ' Note: Cannot log from pure logic layer (violates Clean Architecture)
    ' Caller should handle and log errors
    CalculateStraightBarLength = length
End Function

' Calculate hook length
' Parameters:
'   diameter: Bar diameter in mm
'   angle: Bend angle in degrees (90, 135, 180)
' Returns: Hook length in mm
Public Function CalculateHookLength(diameter As Double, angle As Double) As Double
    On Error GoTo ErrHandler
    
    Dim bendRadius As Double
    Dim extensionLength As Double
    Dim arcLength As Double
    
    ' Determine bend radius based on angle
    Select Case angle
        Case 90
            bendRadius = REBAR_BEND_RADIUS_90 * diameter
            extensionLength = REBAR_HOOK_EXTENSION * diameter
        Case 135
            bendRadius = REBAR_BEND_RADIUS_135 * diameter
            extensionLength = REBAR_HOOK_EXTENSION * diameter
        Case 180
            bendRadius = REBAR_BEND_RADIUS_180 * diameter
            extensionLength = REBAR_HOOK_EXTENSION * diameter
        Case Else
            bendRadius = REBAR_BEND_RADIUS_90 * diameter
            extensionLength = REBAR_HOOK_EXTENSION * diameter
    End Select
    
    ' Calculate arc length: L = R * θ (θ in radians)
    Dim angleRad As Double
    angleRad = angle * 3.14159265358979 / 180#
    arcLength = bendRadius * angleRad
    
    ' Total hook length = arc length + extension
    CalculateHookLength = arcLength + extensionLength
    Exit Function
    
ErrHandler:
    
    CalculateHookLength = 0#
End Function

' Calculate U-shape (shape code 18) cut length
' Parameters:
'   width: Internal width in mm
'   height: Leg height in mm
'   diameter: Bar diameter in mm
' Returns: Total cut length in mm
Public Function CalculateUShapeLength(width As Double, _
                                     height As Double, _
                                     diameter As Double) As Double
    On Error GoTo ErrHandler
    
    Dim bendRadius As Double
    Dim straightLength As Double
    Dim bendAllowance As Double
    
    bendRadius = REBAR_BEND_RADIUS_90 * diameter
    
    ' Calculate bend allowance for 90-degree bends
    ' Bend allowance = π * R * (angle/180) - (R * tan(angle/2))
    bendAllowance = 3.14159265358979 * bendRadius * 0.5 - bendRadius
    
    ' Total length = 2 * height + width + 2 * bend allowances
    straightLength = (2 * height) + width + (2 * bendAllowance)
    
    CalculateUShapeLength = straightLength
    Exit Function
    
ErrHandler:
    
    CalculateUShapeLength = 0#
End Function

' Calculate stirrup (shape code 51) cut length
' Parameters:
'   width: Internal width in mm
'   height: Internal height in mm
'   diameter: Bar diameter in mm
'   hookLength: Extension length for hooks (default = 75mm)
' Returns: Total cut length in mm
Public Function CalculateStirrupLength(width As Double, _
                                      height As Double, _
                                      diameter As Double, _
                                      Optional hookLength As Double = 75#) As Double
    On Error GoTo ErrHandler
    
    Dim bendRadius As Double
    Dim perimeter As Double
    Dim bendAllowance As Double
    
    bendRadius = REBAR_BEND_RADIUS_90 * diameter
    
    ' Calculate perimeter (inside dimensions)
    perimeter = 2 * (width + height)
    
    ' Bend allowance for 4 corners (90-degree bends)
    bendAllowance = 4 * (3.14159265358979 * bendRadius * 0.5 - bendRadius)
    
    ' Add hook extensions (2 hooks)
    Dim totalLength As Double
    totalLength = perimeter + bendAllowance + (2 * hookLength)
    
    CalculateStirrupLength = totalLength
    Exit Function
    
ErrHandler:
    
    CalculateStirrupLength = 0#
End Function

' General shape cut length calculator (dispatches to specific functions)
' Parameters:
'   shapeCode: Standard shape code (1=straight, 18=U, 51=stirrup, etc.)
'   dimensions: Dictionary with required dimensions
'   diameter: Bar diameter in mm
' Returns: Total cut length in mm
Public Function CalculateCutLength(shapeCode As Long, _
                                  dimensions As Object, _
                                  diameter As Double) As Double
    On Error GoTo ErrHandler
    
    Dim cutLength As Double
    
    Select Case shapeCode
        Case 1 ' Straight bar
            If dimensions.exists("Length") Then
                cutLength = CalculateStraightBarLength(CDbl(dimensions("Length")), diameter, False)
            End If
            
        Case 2 ' Straight with hooks
            If dimensions.exists("Length") Then
                cutLength = CalculateStraightBarLength(CDbl(dimensions("Length")), diameter, True)
            End If
            
        Case 18 ' U-shape
            If dimensions.exists("Width") And dimensions.exists("Height") Then
                cutLength = CalculateUShapeLength(CDbl(dimensions("Width")), _
                                                 CDbl(dimensions("Height")), _
                                                 diameter)
            End If
            
        Case 51 ' Stirrup
            If dimensions.exists("Width") And dimensions.exists("Height") Then
                Dim hookLen As Double
                hookLen = 75# ' Default
                If dimensions.exists("HookLength") Then
                    hookLen = CDbl(dimensions("HookLength"))
                End If
                cutLength = CalculateStirrupLength(CDbl(dimensions("Width")), _
                                                  CDbl(dimensions("Height")), _
                                                  diameter, _
                                                  hookLen)
            End If
            
        Case Else
            ' Unknown shape - return 0
            cutLength = 0#
    End Select
    
    CalculateCutLength = cutLength
    Exit Function
    
ErrHandler:
    
    CalculateCutLength = 0#
End Function

' ==========================================
' SPACING & DISTRIBUTION CALCULATIONS
' ==========================================

' Calculate number of bars based on spacing
' Parameters:
'   totalLength: Total length to distribute bars (mm)
'   spacing: Center-to-center spacing (mm)
'   startOffset: Offset from start (mm)
'   endOffset: Offset from end (mm)
' Returns: Number of bars
Public Function CalculateBarCount(totalLength As Double, _
                                 spacing As Double, _
                                 Optional startOffset As Double = 0#, _
                                 Optional endOffset As Double = 0#) As Long
    On Error GoTo ErrHandler
    
    If spacing <= 0 Or totalLength <= 0 Then
        CalculateBarCount = 0
        Exit Function
    End If
    
    Dim effectiveLength As Double
    effectiveLength = totalLength - startOffset - endOffset
    
    If effectiveLength <= 0 Then
        CalculateBarCount = 0
        Exit Function
    End If
    
    ' Number of bars = (effective length / spacing) + 1
    Dim count As Long
    count = Int(effectiveLength / spacing) + 1
    
    CalculateBarCount = count
    Exit Function
    
ErrHandler:
    
    CalculateBarCount = 0
End Function

' Calculate actual spacing based on number of bars
' Parameters:
'   totalLength: Total length to distribute bars (mm)
'   barCount: Number of bars
'   startOffset: Offset from start (mm)
'   endOffset: Offset from end (mm)
' Returns: Actual center-to-center spacing (mm)
Public Function CalculateActualSpacing(totalLength As Double, _
                                      barCount As Long, _
                                      Optional startOffset As Double = 0#, _
                                      Optional endOffset As Double = 0#) As Double
    On Error GoTo ErrHandler
    
    If barCount <= 1 Or totalLength <= 0 Then
        CalculateActualSpacing = 0#
        Exit Function
    End If
    
    Dim effectiveLength As Double
    effectiveLength = totalLength - startOffset - endOffset
    
    If effectiveLength <= 0 Then
        CalculateActualSpacing = 0#
        Exit Function
    End If
    
    ' Actual spacing = effective length / (count - 1)
    Dim actualSpacing As Double
    actualSpacing = effectiveLength / (barCount - 1)
    
    CalculateActualSpacing = actualSpacing
    Exit Function
    
ErrHandler:
    
    CalculateActualSpacing = 0#
End Function

' Generate bar positions along a line
' Parameters:
'   totalLength: Total length (mm)
'   spacing: Center-to-center spacing (mm)
'   startOffset: Offset from start (mm)
' Returns: Collection of positions (Double values)
Public Function GenerateBarPositions(totalLength As Double, _
                                    spacing As Double, _
                                    Optional startOffset As Double = 0#) As Collection
    On Error GoTo ErrHandler
    
    Dim positions As New Collection
    
    If spacing <= 0 Or totalLength <= 0 Then
        Set GenerateBarPositions = positions
        Exit Function
    End If
    
    Dim currentPos As Double
    currentPos = startOffset
    
    Do While currentPos <= totalLength
        positions.Add currentPos
        currentPos = currentPos + spacing
    Loop
    
    Set GenerateBarPositions = positions
    Exit Function
    
ErrHandler:
    
    Set GenerateBarPositions = New Collection
End Function

' ==========================================
' WEIGHT CALCULATIONS
' ==========================================

' Calculate weight of rebar
' Parameters:
'   diameter: Bar diameter in mm
'   length: Total length in mm
'   quantity: Number of bars
' Returns: Total weight in kg
Public Function CalculateWeight(diameter As Double, _
                               length As Double, _
                               Optional quantity As Long = 1) As Double
    On Error GoTo ErrHandler
    
    ' Weight formula: W = (D^2 / 162) * L(meters) * Quantity
    ' Where: D = diameter in mm, L = length in meters
    
    If diameter <= 0 Or length <= 0 Then
        CalculateWeight = 0#
        Exit Function
    End If
    
    Dim lengthMeters As Double
    lengthMeters = length / 1000#
    
    Dim weight As Double
    weight = (diameter * diameter / 162#) * lengthMeters * quantity
    
    CalculateWeight = weight
    Exit Function
    
ErrHandler:
    
    CalculateWeight = 0#
End Function

' ==========================================
' ANCHORAGE & DEVELOPMENT LENGTH
' ==========================================

' Calculate development length (simplified ACI/Eurocode approximation)
' Parameters:
'   diameter: Bar diameter in mm
'   concreteGrade: Concrete strength (e.g., 30 for C30)
'   steelGrade: Steel yield strength (e.g., 400 for Grade 400)
' Returns: Development length in mm
Public Function CalculateDevelopmentLength(diameter As Double, _
                                          concreteGrade As Double, _
                                          steelGrade As Double) As Double
    On Error GoTo ErrHandler
    
    If diameter <= 0 Or concreteGrade <= 0 Or steelGrade <= 0 Then
        CalculateDevelopmentLength = 0#
        Exit Function
    End If
    
    ' Simplified formula: Ld = k * (fy/sqrt(fc')) * diameter
    ' Where k ≈ 1.25 for tension, fy = steel grade, fc' = concrete grade
    
    Dim k As Double
    k = 1.25
    
    Dim fcSqrt As Double
    fcSqrt = Sqr(concreteGrade)
    
    Dim developmentLength As Double
    developmentLength = k * (steelGrade / fcSqrt) * diameter
    
    CalculateDevelopmentLength = developmentLength
    Exit Function
    
ErrHandler:
    
    CalculateDevelopmentLength = 0#
End Function

' ==========================================
' COVER CALCULATIONS
' ==========================================

' Calculate effective rebar position considering cover
' Parameters:
'   elementWidth: Element dimension (mm)
'   cover: Concrete cover (mm)
'   diameter: Bar diameter (mm)
' Returns: Distance from face to bar center (mm)
Public Function CalculateEffectiveCover(elementWidth As Double, _
                                       cover As Double, _
                                       diameter As Double) As Double
    On Error GoTo ErrHandler
    
    ' Effective cover = cover + diameter/2
    CalculateEffectiveCover = cover + (diameter / 2#)
    Exit Function
    
ErrHandler:
    
    CalculateEffectiveCover = cover
End Function

' Calculate available space for rebar distribution
' Parameters:
'   elementWidth: Total width (mm)
'   coverStart: Cover at start side (mm)
'   coverEnd: Cover at end side (mm)
' Returns: Available space for rebar distribution (mm)
Public Function CalculateAvailableSpace(elementWidth As Double, _
                                       coverStart As Double, _
                                       coverEnd As Double) As Double
    On Error GoTo ErrHandler
    
    Dim availableSpace As Double
    availableSpace = elementWidth - coverStart - coverEnd
    
    If availableSpace < 0 Then availableSpace = 0#
    
    CalculateAvailableSpace = availableSpace
    Exit Function
    
ErrHandler:
    
    CalculateAvailableSpace = 0#
End Function

' ==========================================
' VALIDATION
' ==========================================

' Check if spacing meets minimum requirements
' Parameters:
'   spacing: Proposed spacing (mm)
'   diameter: Bar diameter (mm)
'   aggregateSize: Maximum aggregate size (mm)
' Returns: True if spacing is valid
Public Function ValidateSpacing(spacing As Double, _
                               diameter As Double, _
                               Optional aggregateSize As Double = 20#) As Boolean
    On Error GoTo ErrHandler
    
    ' Minimum spacing rules:
    ' 1. Greater than bar diameter
    ' 2. Greater than aggregate size + 5mm
    ' 3. Greater than 25mm (typical minimum)
    
    Dim minSpacing As Double
    minSpacing = diameter
    
    If aggregateSize + 5# > minSpacing Then
        minSpacing = aggregateSize + 5#
    End If
    
    If 25# > minSpacing Then
        minSpacing = 25#
    End If
    
    ValidateSpacing = (spacing >= minSpacing)
    Exit Function
    
ErrHandler:
    
    ValidateSpacing = False
End Function

' ==========================================
' UTILITY FUNCTIONS
' ==========================================

' Round spacing to standard increment
' Parameters:
'   spacing: Calculated spacing (mm)
'   increment: Rounding increment (default = 10mm)
' Returns: Rounded spacing
Public Function RoundSpacing(spacing As Double, _
                            Optional increment As Double = 10#) As Double
    On Error GoTo ErrHandler
    
    If increment <= 0 Then increment = 10#
    
    RoundSpacing = Round(spacing / increment) * increment
    Exit Function
    
ErrHandler:
    
    RoundSpacing = spacing
End Function

' Get standard bar diameter from nominal size
' Returns closest standard diameter
Public Function GetStandardDiameter(nominalSize As Double) As Double
    ' Standard bar sizes (mm): 6, 8, 10, 12, 16, 20, 25, 32, 40
    Dim standardSizes As Variant
    standardSizes = Array(6, 8, 10, 12, 16, 20, 25, 32, 40)
    
    Dim i As Long
    Dim closestSize As Double
    Dim minDiff As Double
    
    closestSize = 12# ' Default
    minDiff = 999999#
    
    For i = LBound(standardSizes) To UBound(standardSizes)
        Dim diff As Double
        diff = Abs(nominalSize - standardSizes(i))
        If diff < minDiff Then
            minDiff = diff
            closestSize = standardSizes(i)
        End If
    Next i
    
    GetStandardDiameter = closestSize
End Function
