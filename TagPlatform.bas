Attribute VB_Name = "TagPlatform"
Type Point
    X As Double
    Y As Double
    Z As Double
End Type

Sub Main()
    '''
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    '''
    Dim ListPipe As Collection
    Set ListPipe = New Collection
    Call SAP2000_Connectv16
    Call FilterPipeData(ListPipe)
    MsgBox ("")
    Call SAP2000_Disconnect
    '''
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub


Sub ImportCSV()
    Set ws = Sheets("PipeData")
    ws.Cells.Clear
    strFile = Application.GetOpenFilename("Text Files (*.csv),*.csv", , "Please select text file...")
    With ws.QueryTables.Add(Connection:= _
        "TEXT;" & strFile, Destination:= _
        Range("$A$1"))
        .TextFileParseType = xlDelimited
        .TextFileCommaDelimiter = True
        .Refresh BackgroundQuery:=False
    End With
    ws.QueryTables(1).SaveData = False
    ws.QueryTables.item(1).Delete
    Columns("G:L").Replace What:="mm", Replacement:="", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
End Sub


Function GetStrVolume() As Collection
    Set temp = New Collection
    lastRow = Sheets("GeometryData").Cells.Find("*", SearchOrder:=xlByRows, searchdirection:=xlPrevious).row
    'FramesData = Application.Transpose(Application.Transpose(Sheets("GeometryData").Range("G3:L" & LastRow).Value2))
    maxX = Application.Max(Sheets("GeometryData").Range("G3:G" & lastRow).Value2, Sheets("GeometryData").Range("J3:J" & lastRow).Value2) + 1000
    minX = Application.Min(Sheets("GeometryData").Range("G3:G" & lastRow).Value2, Sheets("GeometryData").Range("J3:J" & lastRow).Value2) - 1000
    maxY = Application.Max(Sheets("GeometryData").Range("H3:H" & lastRow).Value2, Sheets("GeometryData").Range("K3:K" & lastRow).Value2) + 1000
    minY = Application.Min(Sheets("GeometryData").Range("H3:H" & lastRow).Value2, Sheets("GeometryData").Range("K3:K" & lastRow).Value2) - 1000
    maxZ = Application.Max(Sheets("GeometryData").Range("I3:I" & lastRow).Value2, Sheets("GeometryData").Range("L3:L" & lastRow).Value2) + 500
    minZ = Application.Min(Sheets("GeometryData").Range("I3:I" & lastRow).Value2, Sheets("GeometryData").Range("L3:L" & lastRow).Value2) - 500
    temp.Add maxX: temp.Add minX: temp.Add maxY: temp.Add minY: temp.Add maxZ: temp.Add minZ
    Set GetStrVolume = temp
End Function

Function GetPipesPoints() As Object
    lastRow = Sheets("PipeData").Cells.Find("*", SearchOrder:=xlByRows, searchdirection:=xlPrevious).row
    data = Application.Transpose(Application.Transpose(Sheets("PipeData").Range("A4:Q" & lastRow).Value2))
    '''
    Set PointsName = CreateObject("Scripting.Dictionary")
    Set PointsX = CreateObject("Scripting.Dictionary")
    Set PointsY = CreateObject("Scripting.Dictionary")
    Set pointsZ = CreateObject("Scripting.Dictionary")
    j = 100
    For i = LBound(data) To UBound(data) 'i = 7 To 7
        Dim p1 As Point, p2 As Point
        p1.X = data(i, 7): p1.Y = data(i, 8): p1.Z = data(i, 9)
        p2.X = data(i, 10): p2.Y = data(i, 11): p2.Z = data(i, 12)
        If Not PointsName.exists(p1.X & "-" & p1.Y & "-" & p1.Z) Then
            PointsName.Add p1.X & "-" & p1.Y & "-" & p1.Z, j
            PointsX.Add j, p1.X
            PointsY.Add j, p1.Y
            pointsZ.Add j, p1.Z
            j = j + 1
        End If
        
        If Not PointsName.exists(p2.X & "-" & p2.Y & "-" & p2.Z) Then
            PointsName.Add p2.X & "-" & p2.Y & "-" & p2.Z, j
            PointsX.Add j, p2.X
            PointsY.Add j, p2.Y
            pointsZ.Add j, p2.Z
            j = j + 1
        End If
        Sheets("PipeData").Range("R" & i + 3).Value2 = PointsName(p1.X & "-" & p1.Y & "-" & p1.Z)
        Sheets("PipeData").Range("S" & i + 3).Value2 = PointsName(p2.X & "-" & p2.Y & "-" & p2.Z)
    Next
    Set temp = New Collection
    temp.Add PointsName
    temp.Add PointsX
    temp.Add PointsY
    temp.Add pointsZ
    Set GetPipesPoints = temp
    Set temp = Nothing
End Function

Sub FilterPipeData(ListPipe As Collection)
    lastRow = Sheets("PipeData").Cells.Find("*", SearchOrder:=xlByRows, searchdirection:=xlPrevious).row
    data = Application.Transpose(Application.Transpose(Sheets("PipeData").Range("A4:Q" & lastRow).Value2))
    '''
    'Set Pipelines = ArrayToDictionary_ValueBased(RemoveDuplicatesInArray(Application.Transpose(Sheets("PipeData").Range("B4:B" & LastRow).Value2)))
    '''
    Set vol = GetStrVolume
    maxX = vol(1): minX = vol(2): maxY = vol(3): minY = vol(4): maxZ = vol(5): minZ = vol(6)
    maxX = 9999999: minX = -9999999: maxY = 9999999: minY = -9999999: maxZ = 9999999: minZ = -9999999:
    '''
    'Set PipesPoints = CreateObject("Scripting.Dictionary")
    'Set PipesDia = CreateObject("Scripting.Dictionary")
    'Set PipesEmpty = CreateObject("Scripting.Dictionary")
    'Set PipesOperating = CreateObject("Scripting.Dictionary")
    'Set PipesTest = CreateObject("Scripting.Dictionary")
    'For Each item In Pipelines
        'Set temp = CreateObject("Scripting.Dictionary")
        'PipesPoints.Add Pipelines(item), temp
        'Set temp = Nothing
        'Index = Application.Match(item, Sheets("PipeData").Range("B:B"), 0)
        'PipesDia.Add Pipelines(item), Data(Index, 3)
        'PipesEmpty.Add Pipelines(item), Data(Index, 14)
        'PipesOperating.Add Pipelines(item), Data(Index, 15)
        'PipesTest.Add Pipelines(item), Data(Index, 16)
    'Next
    '''
    Dim p1 As Point, p2 As Point
    'ret = SapModel.SetPresentUnits(8) 'kN_mm_C
    For i = LBound(data) To UBound(data) 'i = 7 To 7
        p1.X = data(i, 7): p1.Y = data(i, 8): p1.Z = data(i, 9)
        p2.X = data(i, 10): p2.Y = data(i, 11): p2.Z = data(i, 12)
        x1 = Application.Max(p1.X, p2.X)
        x2 = Application.Min(p1.X, p2.X)
        y1 = Application.Max(p1.Y, p2.Y)
        y2 = Application.Min(p1.Y, p2.Y)
        z1 = Application.Max(p1.Z, p2.Z)
        z2 = Application.Min(p1.Z, p2.Z)
        '''
        
        If (((p1.X > minX And p1.X < maxX) And (p2.X > minX And p2.X < maxX)) _
            Or ((p1.Y > minY And p1.Y < maxY) And (p2.Y > minY And p2.Y < maxY))) _
                And ((p1.Z > minZ And p1.Z < maxZ) And (p2.Z > minZ And p2.Z < maxZ)) Then
                If Not (x1 < minX Or x2 > maxX Or y1 < minY Or y2 > maxY Or z1 < minZ Or z2 > maxZ) Then
                    Name = DrawPipe_SAP2000(p1, p2, data(i, 1) & "-" & data(i, 2) & "-" & i, data(i, 3), data(i, 14) * 0.00001, data(i, 15) * 0.00001, data(i, 16) * 0.00001)
                    Sheets("PipeData").Range("R" & i + 3).Value2 = Name
                    ListPipe.Add Name
                End If
        End If
        
    Next
    'ret = SapModel.SetPresentUnits(5) 'kN_mm_C
End Sub

Function DrawPipe_SAP2000(StartPoint As Point, EndPoint As Point, UserName As String, Diameter, EmptyLoadVal, OperatingLoadVal, TestLoadVal) As String
    Dim Name As String
    '''
    ret = SapModel.PropMaterial.SetMaterial("Pipe", MATERIAL_STEEL)
    '''
    ret = SapModel.PropFrame.SetPipe("PIPE" & Diameter, "Pipe", Diameter, 0.01)
    '''
    ret = SapModel.frameObj.AddByCoord(StartPoint.X, StartPoint.Y, StartPoint.Z, EndPoint.X, EndPoint.Y, EndPoint.Z, Name, "PIPE" & Diameter, UserName)
    '''
    ret = SapModel.loadPatterns.Add("PE", 2)
    ret = SapModel.frameObj.SetLoadDistributed(Name, "PE", 1, 10, 0, 1, EmptyLoadVal, EmptyLoadVal, "Global", True, True, 0)
    '''
    ret = SapModel.loadPatterns.Add("PO", 2)
    ret = SapModel.frameObj.SetLoadDistributed(Name, "PO", 1, 10, 0, 1, OperatingLoadVal, OperatingLoadVal, "Global", True, True, 0)
    '''
    ret = SapModel.loadPatterns.Add("PT", 2)
    ret = SapModel.frameObj.SetLoadDistributed(Name, "PT", 1, 10, 0, 1, TestLoadVal, TestLoadVal, "Global", True, True, 0)
    '''
    DrawPipe_SAP2000 = Name
End Function

Function GetShortestDistanceBetween2Lines(p1 As Point, p2 As Point, p3 As Point, p4 As Point)


End Function
