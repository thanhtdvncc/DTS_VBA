Attribute VB_Name = "Core_XData_Reader"

Option Explicit
'===============================================================
' Core Module: Core_XData_Reader
' Purpose: Read entity data from AutoCAD
' Dependencies: None (pure AutoCAD reading)
' Note: Avoid storing UDT in Variant/Collection to prevent VBA coercion issues.
'===============================================================

Private Const APP_NAME As String = "DTS_SAP2000"
Private Const XD_APPNAME As Integer = 1001
Private Const XD_STRING As Integer = 1000
Private Const XD_REAL As Integer = 1040
Private Const XD_LONG As Integer = 1071

' ----------------------------
' Data Structures
' ----------------------------
Public Type CADPoint
    entityHandle As String
    layerName As String
    nodeName As String
    X As Double
    Y As Double
    Z As Double
    springData As String
End Type

Public Type CADFrame
    entityHandle As String
    layerName As String
    frameName As String
    Point1Name As String
    Point2Name As String
    sectionName As String
End Type

Public Type CADArea
    entityHandle As String
    layerName As String
    areaName As String
    sectionName As String
    pointList As String ' comma-separated
End Type

' Read all Points from drawing
' Fill pointsArr() with CADPoint UDTs and return count
Public Function ReadPointsFromCAD(acadDoc As Object, ByRef pointsArr() As CADPoint) As Long
    On Error GoTo ErrHandler

    Dim ms As Object
    Set ms = acadDoc.ModelSpace

    Debug.Print "ReadPointsFromCAD: ModelSpace.Count ="; ms.count

    Dim ent As Object
    Dim count As Long
    count = 0
    Erase pointsArr ' ensure empty

    For Each ent In ms
        Dim entType As String
        entType = LCase(TypeName(ent))
        
        ' Flexible match for circle-like types
        If InStr(entType, "circle") > 0 Then
            Dim pt As CADPoint
            If ExtractPointData(ent, pt) Then
                If count = 0 Then
                    ReDim pointsArr(0 To 0)
                Else
                    ReDim Preserve pointsArr(0 To count)
                End If
                pointsArr(count) = pt
                count = count + 1
            End If
        End If
    Next ent

    ReadPointsFromCAD = count
    Exit Function

ErrHandler:
    Debug.Print "Error in ReadPointsFromCAD: " & err.number & " - " & err.description
    ReadPointsFromCAD = 0
End Function

' Read all Frames from drawing
Public Function ReadFramesFromCAD(acadDoc As Object, ByRef framesArr() As CADFrame) As Long
    On Error GoTo ErrHandler

    Dim ms As Object
    Set ms = acadDoc.ModelSpace

    Debug.Print "ReadFramesFromCAD: ModelSpace.Count ="; ms.count

    Dim ent As Object
    Dim count As Long
    count = 0
    Erase framesArr

    For Each ent In ms
        Dim entType As String
        entType = LCase(TypeName(ent))
        
        ' Flexible match for line-like types (frame may be line or polyline depending on app)
        If InStr(entType, "line") > 0 Or InStr(entType, "polyline") > 0 Then
            Dim fr As CADFrame
            If ExtractFrameData(ent, fr) Then
                If count = 0 Then
                    ReDim framesArr(0 To 0)
                Else
                    ReDim Preserve framesArr(0 To count)
                End If
                framesArr(count) = fr
                count = count + 1
            End If
        End If
    Next ent

    ReadFramesFromCAD = count
    Exit Function

ErrHandler:
    Debug.Print "Error in ReadFramesFromCAD: " & err.number & " - " & err.description
    ReadFramesFromCAD = 0
End Function

' Read all Areas from drawing
Public Function ReadAreasFromCAD(acadDoc As Object, ByRef areasArr() As CADArea) As Long
    On Error GoTo ErrHandler

    Dim ms As Object
    Set ms = acadDoc.ModelSpace

    Debug.Print "ReadAreasFromCAD: ModelSpace.Count ="; ms.count

    Dim ent As Object
    Dim count As Long
    count = 0
    Erase areasArr

    For Each ent In ms
        Dim entType As String
        entType = LCase(TypeName(ent))
        
        ' Match polyline/hatch/region types more flexibly
        If InStr(entType, "lwpolyline") > 0 Or InStr(entType, "polyline") > 0 Or InStr(entType, "3dpolyline") > 0 Or InStr(entType, "hatch") > 0 Then
            Dim ar As CADArea
            If ExtractAreaData(ent, ar) Then
                If count = 0 Then
                    ReDim areasArr(0 To 0)
                Else
                    ReDim Preserve areasArr(0 To count)
                End If
                areasArr(count) = ar
                count = count + 1
            End If
        End If
    Next ent

    ReadAreasFromCAD = count
    Exit Function

ErrHandler:
    Debug.Print "Error in ReadAreasFromCAD: " & err.number & " - " & err.description
    ReadAreasFromCAD = 0
End Function

Private Function ExtractPointData(ent As Object, ByRef pt As CADPoint) As Boolean
    On Error GoTo ErrHandler
    ExtractPointData = False
    
    Dim xdType As Variant
    Dim xdVal As Variant
    
    ent.GetXData "DTS_SAP2000", xdType, xdVal
    
    If IsEmpty(xdVal) Or Not IsArray(xdVal) Then
        Debug.Print "ExtractPointData FAIL: No XData for handle " & ent.Handle
        Exit Function
    End If
    
    If UBound(xdVal) < 4 Then
        Debug.Print "ExtractPointData FAIL: Insufficient XData elements (" & UBound(xdVal) & ") for handle " & ent.Handle
        Exit Function
    End If
    
    pt.entityHandle = ent.Handle
    pt.layerName = ent.layer
    pt.nodeName = Trim$(CStr(xdVal(1)))
    
    If Len(pt.nodeName) = 0 Then
        Debug.Print "ExtractPointData FAIL: Empty nodeName for handle " & ent.Handle
        Exit Function
    End If
    
    pt.X = CDbl(xdVal(2))
    pt.Y = CDbl(xdVal(3))
    pt.Z = CDbl(xdVal(4))
    
    If UBound(xdVal) >= 5 Then
        pt.springData = CStr(xdVal(5))
    Else
        pt.springData = ""
    End If
    
    ExtractPointData = True
    Exit Function

ErrHandler:
    Debug.Print "ExtractPointData ERROR: " & err.number & " - " & err.description & " (Handle: " & ent.Handle & ")"
    ExtractPointData = False
End Function

Private Function ExtractFrameData(ent As Object, ByRef fr As CADFrame) As Boolean
    On Error GoTo ErrHandler
    ExtractFrameData = False
    
    Dim xdType As Variant
    Dim xdVal As Variant
    
    ent.GetXData "DTS_SAP2000", xdType, xdVal
    
    If IsEmpty(xdVal) Or Not IsArray(xdVal) Then
        Debug.Print "ExtractFrameData FAIL: No XData for handle " & ent.Handle
        Exit Function
    End If
    
    If UBound(xdVal) < 4 Then
        Debug.Print "ExtractFrameData FAIL: Insufficient XData for handle " & ent.Handle
        Exit Function
    End If
    
    fr.entityHandle = ent.Handle
    fr.layerName = ent.layer
    fr.frameName = CStr(xdVal(1))
    fr.Point1Name = CStr(xdVal(2))
    fr.Point2Name = CStr(xdVal(3))
    fr.sectionName = CStr(xdVal(4))
    
    ExtractFrameData = True
    Exit Function

ErrHandler:
    Debug.Print "ExtractFrameData ERROR: " & err.number & " - " & err.description & " (Handle: " & ent.Handle & ")"
    ExtractFrameData = False
End Function

Private Function ExtractAreaData(ent As Object, ByRef ar As CADArea) As Boolean
    On Error GoTo ErrHandler
    ExtractAreaData = False
    
    Dim xdType As Variant
    Dim xdVal As Variant
    
    ent.GetXData "DTS_SAP2000", xdType, xdVal
    
    If IsEmpty(xdVal) Or Not IsArray(xdVal) Then
        Debug.Print "ExtractAreaData FAIL: No XData for handle " & ent.Handle
        Exit Function
    End If
    
    If UBound(xdVal) < 3 Then
        Debug.Print "ExtractAreaData FAIL: Insufficient XData for handle " & ent.Handle
        Exit Function
    End If
    
    ar.entityHandle = ent.Handle
    ar.layerName = ent.layer
    ar.areaName = CStr(xdVal(1))
    ar.sectionName = CStr(xdVal(2))
    ar.pointList = CStr(xdVal(3))
    
    ExtractAreaData = True
    Exit Function

ErrHandler:
    Debug.Print "ExtractAreaData ERROR: " & err.number & " - " & err.description & " (Handle: " & ent.Handle & ")"
    ExtractAreaData = False
End Function

Public Sub ExportDTSMetadataToExcel(acadDoc As Object)
    On Error GoTo ErrHandler

    Const SHEET_NAME As String = "DTS_Metadata"
    Dim ws As Worksheet
    Dim wb As Workbook
    Set wb = ThisWorkbook

    ' prepare worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(SHEET_NAME)
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count))
        ws.Name = SHEET_NAME
    Else
        ws.Cells.Clear
    End If
    On Error GoTo ErrHandler

    ' Headers
    ws.Range("A1").Value = "CAD_Handle"
    ws.Range("B1").Value = "CAD_Type"
    ws.Range("C1").Value = "Layer"
    ws.Range("D1").Value = "Entity_Type" ' e.g., Point/Frame/Area/Other
    ws.Range("E1").Value = "XData_App"
    ws.Range("F1").Value = "XData_Summary" ' serialized Type:Value pairs
    ws.Range("G1").Value = "X"
    ws.Range("H1").Value = "Y"
    ws.Range("I1").Value = "Z"
    ws.Range("J1").Value = "Notes"

    Dim row As Long
    row = 2

    Dim ms As Object
    Set ms = acadDoc.ModelSpace

    Dim ent As Object
    For Each ent In ms
        Dim xdType As Variant
        Dim xdVal As Variant
        Dim hasXData As Boolean
        hasXData = False

        ' Try get XData for DTS application
        On Error Resume Next
        ent.GetXData "DTS_SAP2000", xdType, xdVal
        If err.number = 0 Then
            If Not IsEmpty(xdVal) And IsArray(xdVal) Then
                Dim k As Long
                For k = LBound(xdVal) To UBound(xdVal)
                    If CStr(xdVal(k)) = APP_NAME Then
                        hasXData = True
                        Exit For
                    End If
                Next k
                If Not hasXData Then
                    If IsArray(xdType) Then
                        If UBound(xdType) >= 0 Then
                            If xdType(0) = XD_APPNAME Then
                                hasXData = True
                            End If
                        End If
                    End If
                End If
            End If
        End If
        err.Clear
        On Error GoTo ErrHandler

        If Not hasXData Then
            ' skip entities that are not from the DTS app
            GoTo NextEntity
        End If

        ' Determine friendly CAD entity type
        Dim entTypeName As String
        entTypeName = TypeName(ent)

        Dim friendlyType As String
        Select Case entTypeName
            Case "AcadCircle"
                friendlyType = "Point"
            Case "AcadLine"
                friendlyType = "Frame"
            Case "AcadLWPolyline", "AcadPolyline", "Acad3DPolyline"
                friendlyType = "Area"
            Case Else
                friendlyType = entTypeName
        End Select

        ' Basic fields
        Dim handleVal As String
        On Error Resume Next
        handleVal = CStr(ent.Handle)
        If err.number <> 0 Then handleVal = "<unknown>"
        err.Clear
        On Error GoTo ErrHandler

        ws.Cells(row, 1).Value = handleVal
        ws.Cells(row, 2).Value = entTypeName
        ws.Cells(row, 3).Value = ent.layer
        ws.Cells(row, 4).Value = friendlyType

        ' XData app name column
        Dim appNameRead As String
        appNameRead = ""
        If IsArray(xdVal) Then
            On Error Resume Next
            If UBound(xdVal) >= 0 Then appNameRead = CStr(xdVal(0))
            err.Clear
            On Error GoTo ErrHandler
        End If
        ws.Cells(row, 5).Value = appNameRead

        ' Serialize XData as TypeCode:Value | ...
        Dim summary As String
        summary = ""
        If IsArray(xdType) And IsArray(xdVal) Then
            Dim ti As Long
            Dim maxIdx As Long
            maxIdx = Application.WorksheetFunction.Min(UBound(xdVal), UBound(xdType))
            For ti = LBound(xdVal) To UBound(xdVal)
                Dim tcode As String
                Dim tVal As String
                On Error Resume Next
                If IsArray(xdType) Then
                    If ti <= UBound(xdType) Then
                        tcode = CStr(xdType(ti))
                    Else
                        tcode = ""
                    End If
                Else
                    tcode = ""
                End If
                tVal = CStr(xdVal(ti))
                On Error GoTo ErrHandler

                If summary <> "" Then summary = summary & " | "
                summary = summary & tcode & ":" & tVal
            Next ti
        Else
            summary = "<unreadable XData>"
        End If
        ws.Cells(row, 6).Value = summary

        ' Try fill X,Y,Z if present
        On Error Resume Next
        Dim maybeX As Double, maybeY As Double, maybeZ As Double
        maybeX = 0#: maybeY = 0#: maybeZ = 0#
        If IsArray(xdVal) Then
            If UBound(xdVal) >= 4 Then
                maybeX = CDbl(xdVal(2))
                maybeY = CDbl(xdVal(3))
                maybeZ = CDbl(xdVal(4))
                ws.Cells(row, 7).Value = maybeX
                ws.Cells(row, 8).Value = maybeY
                ws.Cells(row, 9).Value = maybeZ
            Else
                Select Case entTypeName
                    Case "AcadCircle", "AcadPoint"
                        Dim center As Variant
                        center = ent.center
                        ws.Cells(row, 7).Value = center(0)
                        ws.Cells(row, 8).Value = center(1)
                        ws.Cells(row, 9).Value = center(2)
                    Case "AcadLine"
                        Dim sp As Variant, ep As Variant
                        sp = ent.StartPoint
                        ep = ent.EndPoint
                        ws.Cells(row, 7).Value = sp(0)
                        ws.Cells(row, 8).Value = sp(1)
                        ws.Cells(row, 9).Value = sp(2)
                    Case Else
                        ' leave blank
                End Select
            End If
        End If
        err.Clear
        On Error GoTo ErrHandler

        ws.Cells(row, 10).Value = "" ' placeholder for notes
        row = row + 1

NextEntity:
    Next ent

    ' Autofit columns for readability
    ws.Columns("A:J").AutoFit

    Exit Sub

ErrHandler:
    MsgBox "Error in ExportDTSMetadataToExcel: " & err.number & " - " & err.description, vbCritical, "Export Error"
End Sub

Public Sub ExportDTSMetadataToExcel_Run()
    On Error GoTo ErrHandler

    Dim acadApp As Object
    Dim acadDoc As Object

    ' Try get or create AutoCAD
    Set acadApp = Core_Utils.GetOrCreateAutoCAD()
    If acadApp Is Nothing Then
        MsgBox "AutoCAD not available.", vbExclamation, "Export DTS Metadata"
        Exit Sub
    End If

    ' Use active document
    On Error Resume Next
    Set acadDoc = acadApp.ActiveDocument
    On Error GoTo ErrHandler

    If acadDoc Is Nothing Then
        MsgBox "No active drawing in AutoCAD.", vbExclamation, "Export DTS Metadata"
        Exit Sub
    End If

    ' Call the real export routine
    ExportDTSMetadataToExcel acadDoc

    Exit Sub

ErrHandler:
    MsgBox "Error in ExportDTSMetadataToExcel_Run: " & err.number & " - " & err.description, vbCritical, "Export Error"
End Sub
