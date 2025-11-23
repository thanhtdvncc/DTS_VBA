Attribute VB_Name = "m00_SAP2000_State"
Option Explicit
'===============================================================
' Module: modSAP2000_State
' Purpose: Global state & feature flags
'===============================================================

Public SapApp As Object             ' SAP2000 main application object
Public SapModel As cSapModel             ' SAP2000 SapModel object

' Joint cache
Public gPointNames() As String
Public gPointCount As Long
Public gDictX As Object
Public gDictY As Object
Public gDictZ As Object

' Frame cache
Public gFrameNames() As String
Public gFrameCount As Long
Public gFrameP1() As String
Public gFrameP2() As String
Public gFrameProp() As String
Public gFrameAngle() As Double

' Area cache
Public gAreaNames() As String
Public gAreaCount As Long
Public gAreaProp() As String
Public gAreaNumPts() As Long
Public gAreaPointStr() As String
Public gAreaCentX() As Double
Public gAreaCentY() As Double
Public gAreaCentZ() As Double
Public gAreaGeomArea() As Double
Public gAreaNx() As Double
Public gAreaNy() As Double
Public gAreaNz() As Double

' Feature flags
Public Const ENABLE_AREAS As Boolean = True
Public Const ENABLE_FRAME_SECTIONS As Boolean = True
Public Const ENABLE_AREA_SECTIONS As Boolean = True
Public Const ENABLE_GROUPS As Boolean = True
Public Const ENABLE_LOADCASES As Boolean = True

' Units (kN - mm - C) adjust if needed
Public Const UNIT_KN_MM_C As Long = 5

Public Enum eFrameDictIndex
    eFrameStart = 1
    eFrameEnd = 2
    eFrameProp = 3
    eFrameAngle = 4
End Enum
