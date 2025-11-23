Attribute VB_Name = "ApplyLoad_Auto"
Sub ApplyStressLoad()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    Set wsgeo = Sheets("GeometryData")
    Set wsaf = Sheets("StressInput")
    Set wsoutput = Sheets("StressOutput")
    Dim i As Integer, l As Integer, m As Integer, n As Integer
    '''
    LastColumn = wsaf.Cells(3, wsaf.Columns.count).End(xlToLeft).Column
    wsaf.Range(Number2Letter(LastColumn + 1) & "4:" & Number2Letter(LastColumn + 2) & "1048576").clearContents
    wsoutput.Range("A4:K1048576").clearContents
    l = 4
    n = 4
    Do Until wsaf.Range("A" & n).Value2 = ""
        m = 3
        x0 = wsaf.Range("F" & n).Value2
        y0 = wsaf.Range("G" & n).Value2
        z0 = wsaf.Range("H" & n).Value2
        Do Until wsgeo.Range("A" & m).Value2 = ""
            x1 = wsgeo.Range("G" & m).Value2
            y1 = wsgeo.Range("H" & m).Value2
            z1 = wsgeo.Range("I" & m).Value2
            x2 = wsgeo.Range("J" & m).Value2
            y2 = wsgeo.Range("K" & m).Value2
            z2 = wsgeo.Range("L" & m).Value2
            If (Application.Min(z1, z2) <= z0 And z0 <= Application.Max(z1, z2)) And ((Application.Min(x1, x2) <= x0 And x0 <= Application.Max(x1, x2)) _
                                  Or (Application.Min(y1, y2) <= y0 And y0 <= Application.Max(y1, y2))) Then
                Length = Math.Round(((x1 - x2) ^ 2 + (y1 - y2) ^ 2 + (z1 - z2) ^ 2) ^ 0.5, 2)
                Length1 = Math.Round(((x1 - x0) ^ 2 + (y1 - y0) ^ 2 + (z1 - z0) ^ 2) ^ 0.5, 2)
                Length2 = Math.Round(((x2 - x0) ^ 2 + (y2 - y0) ^ 2 + (z2 - z0) ^ 2) ^ 0.5, 2)
                dist = Abs((x2 - x1) * (y1 - y0) - (x1 - x0) * (y2 - y1)) / Length
                If dist < 1000 And Length1 < Length And Length2 < Length Then
                    Length3 = Math.Round((Length1 ^ 2 - dist ^ 2) ^ 0.5, 0)
                    For i = 9 To LastColumn
                        If wsaf.Range(Number2Letter(i) & n).Value2 <> 0 And wsaf.Range(Number2Letter(i) & n).Value2 <> "" Then
                            wsoutput.Range("A" & l).Value2 = wsgeo.Range("A" & m).Value2
                            wsoutput.Range("B" & l).Value2 = wsaf.Range(Number2Letter(i) & 2).Value2
                            wsoutput.Range("C" & l).Value2 = "GLOBAL"
                            wsoutput.Range("D" & l).Value2 = "Force"
                            wsoutput.Range("E" & l).Value2 = UCase(Right(wsaf.Range(Number2Letter(i) & 3).Value2, 1))
                            wsoutput.Range("F" & l).Value2 = "RelDist"
                            wsoutput.Range("G" & l).Value2 = CStr(Math.Round(Length3 / Length, 3))
                            wsoutput.Range("H" & l).Value2 = CStr(Length3)
                            wsoutput.Range("I" & l).Value2 = wsaf.Range(Number2Letter(i) & n).Value2
                            wsoutput.Range("K" & l).Value2 = "Stress" & "-" & wsaf.Range(Number2Letter(i) & 2).Value2 & "-" & wsaf.Range(Number2Letter(i) & 3).Value2 & "-" & wsaf.Range("B" & n).Value2 & "-" & wsaf.Range("C" & n).Value2
                            l = l + 1
                        End If
                    Next
                    wsaf.Range(Number2Letter(LastColumn + 1) & n).Value2 = "Done"
                    Exit Do
                End If
            End If
            m = m + 1
        Loop
        If wsaf.Range(Number2Letter(LastColumn + 1) & n).Value2 = "" Then: wsaf.Range(Number2Letter(LastColumn + 1) & n).Value2 = "Not Found"
        n = n + 1
    Loop
    MsgBox ("DONE")
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub

Sub ApplyWindLoad()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    Set wsgeo = Sheets("GeometryData")
    Set wsWind = Sheets("WindIntensity")
    Set wsoutput = Sheets("WindLoad")
    Set wscal = Sheets("WindCal")
    Set wssection = Sheets("Section")
    Set wsgroup = Sheets("Groups")
    wsoutput.Range("A4:O1048576").clearContents
    wscal.Range("A3:BM1048576").clearContents
    '''
    groupindex = Application.match(wsWind.Range("K4").Value2, wsgroup.Range("1:1").Value2, 0)
    Set FramesInGroup = CreateObject("Scripting.Dictionary")
    Set TypeList = CreateObject("Scripting.Dictionary")
    TypeList.Add 1, "Point"
    TypeList.Add 2, "Frame"
    TypeList.Add 3, "Cable"
    TypeList.Add 4, "Tendon"
    TypeList.Add 5, "Area"
    TypeList.Add 6, "Solid"
    TypeList.Add 7, "Link"
    i = 2
    Do Until wsgroup.Range(Number2Letter(groupindex) & i).Value2 = ""
        If wsgroup.Range(Number2Letter(groupindex + 1) & i).Value2 = "2" Then
            FramesInGroup.Add wsgroup.Range(Number2Letter(groupindex) & i).Value2, TypeList(wsgroup.Range(Number2Letter(groupindex + 1) & i).Value2)
        End If
        i = i + 1
    Loop
    '''
    If wsWind.Range("I16").Value2 = 0 Then
        elechange = wsWind.Range("L28").Value2 * 1000
    ElseIf wsWind.Range("I16").Value2 = 1 Then
        elechange = wsWind.Range("H28").Value2 * 1000
    End If
    '''
    g = wsWind.Range("J13").Value2
    Cf1 = wsWind.Range("G13").Value2
    Cf2 = wsWind.Range("G14").Value2
    NameLoadCaseX = wsWind.Range("G6").Value2
    NameLoadCaseY = wsWind.Range("G7").Value2
    '''
    GroundEle = wsWind.Range("H9").Value2 * 1000
    '''
    n = 3
    Set T3_Section = CreateObject("Scripting.Dictionary")
    Set T2_Section = CreateObject("Scripting.Dictionary")
    Set FP_Section = CreateObject("Scripting.Dictionary")
    Set Name_Section = CreateObject("Scripting.Dictionary")
    '''
    Do Until wssection.Range("A" & n).Value2 = ""
        T3_Section.Add wssection.Range("A" & n).Value2, wssection.Range("C" & n).Value2
        T2_Section.Add wssection.Range("A" & n).Value2, wssection.Range("D" & n).Value2
        FP_Section.Add wssection.Range("A" & n).Value2, wssection.Range("L" & n).Value2
        str_src = wssection.Range("A" & n).Value2
        temp_str = ""
        For Index = 1 To Len(str_src)
            If Not IsNumeric(mid(str_src, Index, 1)) Then
                temp_str = temp_str & mid(str_src, Index, 1)
            Else
                Exit For
            End If
        Next
        Name_Section.Add wssection.Range("A" & n).Value2, temp_str
        n = n + 1
    Loop
    '''
    Set ListX = CreateObject("Scripting.Dictionary")
    Set ListY = CreateObject("Scripting.Dictionary")
    n = 3
    m = 4
    Do Until wsgeo.Range("A" & n).Value2 = ""
        If FramesInGroup.exists(wsgeo.Range("A" & n).Value2) Then
            x1 = wsgeo.Range("G" & n).Value2
            y1 = wsgeo.Range("H" & n).Value2
            z1 = wsgeo.Range("I" & n).Value2
            x2 = wsgeo.Range("J" & n).Value2
            y2 = wsgeo.Range("K" & n).Value2
            z2 = wsgeo.Range("L" & n).Value2
            angle = wsgeo.Range("E" & n).Value2
            Name = wsgeo.Range("D" & n).Value2
            Length = wsgeo.Range("F" & n).Value2
            tx = Abs(T3_Section(Name) * Cos(angle * 0.0174532925) + T2_Section(Name) * Sin(angle * 0.0174532925)) + FP_Section(Name)
            ty = Abs(T3_Section(Name) * Sin(angle * 0.0174532925) + T2_Section(Name) * Cos(angle * 0.0174532925)) + FP_Section(Name)
            dx = Abs(x1 - x2)
            dy = Abs(y2 - y1)
            '''
            If Abs(z1 - GroundEle) < 1000 Then
                Cf = Cf1
            Else
                Cf = Cf2
            End If
            '''
            Reserve = False
            If z2 < z1 Then
                ele1 = z2
                ele2 = z1
                Reserve = True
            Else
                ele1 = z1
                ele2 = z2
            End If
            '''
            elestart = ele1
            If ele1 <= elechange Then
                q1 = Round(VInterpolate(ele1, wsWind.Range("A4:D5"), 2) * g * Cf, 3)
                If ele1 <= GroundEle Then
                    q1 = Round(VInterpolate(GroundEle, wsWind.Range("A4:D5"), 2) * g * Cf, 3)
                    elestart = GroundEle
                End If
            Else
                q1 = Round(VInterpolate(ele1, wsWind.Range("A6:D100"), 2) * g * Cf, 3)
            End If
            '''
            If ele2 <= elechange Then
                q2 = Round(VInterpolate(ele2, wsWind.Range("A4:D5"), 2) * g * Cf, 3)
            Else
                q2 = Round(VInterpolate(ele2, wsWind.Range("A6:D100"), 2) * g * Cf, 3)
            End If
            '''
            
            CountX = ""
            CountY = ""
    
            If ele2 > GroundEle Then
                q = Round((q1 + q2) / 2, 3)
                If ele1 < ele2 Then
                    If Abs(x1 - x2) < 0.1 And Abs(y1 - y2) < 0.1 Then
                        Mark_Section = "Col"
                        If (ele1 < elechange And ele2 > elechange) Then
                            q = Round((q1 * (elechange - elestart) + q * (ele2 - elechange)) / (ele2 - elestart), 3)
                        End If
                        CountX = elestart & "|" & Mark_Section & "|" & Name & "|" & ele2 - elestart & "|" & ty & "|" & q
                        CountY = elestart & "|" & Mark_Section & "|" & Name & "|" & ele2 - elestart & "|" & tx & "|" & q
                    Else
                        Mark_Section = "VBrace"
                        CountX = ele1 & "|" & Mark_Section & "|" & Name & "|" & Length & "|" & ty & "|" & q
                        CountY = ele1 & "|" & Mark_Section & "|" & Name & "|" & Length & "|" & tx & "|" & q
                    End If
                Else
                    Mark_Section = "Beam"
                    If dy <> 0 Then: CountX = ele1 & "|" & Mark_Section & "|" & Name & "|" & Length & "|" & tx & "|" & q
                    If dx <> 0 Then: CountY = ele1 & "|" & Mark_Section & "|" & Name & "|" & Length & "|" & tx & "|" & q
                End If
            End If
            '''
            If CountX <> "" Then
                If Not ListX.exists(CountX) Then
                    ListX.Add CountX, 1
                Else
                    ListX(CountX) = ListX(CountX) + 1
                End If
                '''
            End If
            If CountY <> "" Then
                If Not ListY.exists(CountY) Then
                    ListY.Add CountY, 1
                Else
                    ListY(CountY) = ListY(CountY) + 1
                End If
            End If
            '''
            If CountX <> "" Then
                wsoutput.Range("A" & m).Value2 = wsgeo.Range("A" & n).Value2
                wsoutput.Range("B" & m).Value2 = NameLoadCaseX
                wsoutput.Range("C" & m).Value2 = "GLOBAL"
                wsoutput.Range("D" & m).Value2 = "FORCE"
                wsoutput.Range("E" & m).Value2 = "X"
                wsoutput.Range("F" & m).Value2 = "RelDist"
            If (ele1 < elechange And ele2 > elechange) And Mark_Section = "Col" Then
                    If Reserve = False Then
                        wsoutput.Range("G" & m).Value2 = Abs(elestart - z1) / Length
                        wsoutput.Range("H" & m).Value2 = Abs(elechange - z1) / Length
                        wsoutput.Range("I" & m).Value2 = Abs(elestart - z1)
                        wsoutput.Range("J" & m).Value2 = Abs(elechange - z1)
                        wsoutput.Range("K" & m).Value2 = q1 * ty * 0.000001
                        wsoutput.Range("L" & m).Value2 = q1 * ty * 0.000001
                        wsoutput.Range("A" & m & ":F" & m + 1).FillDown
                        m = m + 1
                        wsoutput.Range("G" & m).Value2 = Abs(elechange - z1) / Length
                        wsoutput.Range("H" & m).Value2 = Abs(ele2 - z1) / Length
                        wsoutput.Range("I" & m).Value2 = Abs(elechange - z1)
                        wsoutput.Range("J" & m).Value2 = Abs(ele2 - z1)
                        wsoutput.Range("K" & m).Value2 = q1 * ty * 0.000001
                        wsoutput.Range("L" & m).Value2 = q2 * ty * 0.000001
                    Else
                        wsoutput.Range("G" & m).Value2 = Abs(elechange - z1) / Length
                        wsoutput.Range("H" & m).Value2 = Abs(elestart - z1) / Length
                        wsoutput.Range("I" & m).Value2 = Abs(elechange - z1)
                        wsoutput.Range("J" & m).Value2 = Abs(elestart - z1)
                        wsoutput.Range("K" & m).Value2 = q1 * ty * 0.000001
                        wsoutput.Range("L" & m).Value2 = q1 * ty * 0.000001
                        wsoutput.Range("A" & m & ":F" & m + 1).FillDown
                        m = m + 1
                        wsoutput.Range("G" & m).Value2 = Abs(ele2 - z1) / Length
                        wsoutput.Range("H" & m).Value2 = Abs(elechange - z1) / Length
                        wsoutput.Range("I" & m).Value2 = Abs(ele2 - z1)
                        wsoutput.Range("J" & m).Value2 = Abs(elechange - z1)
                        wsoutput.Range("K" & m).Value2 = q2 * ty * 0.000001
                        wsoutput.Range("L" & m).Value2 = q1 * ty * 0.000001
                    End If
                ElseIf Mark_Section = "VBrace" Then
                    wsoutput.Range("G" & m).Value2 = 0
                    wsoutput.Range("H" & m).Value2 = 1
                    wsoutput.Range("I" & m).Value2 = 0
                    wsoutput.Range("J" & m).Value2 = Length
                    If Reserve = False Then
                        wsoutput.Range("K" & m).Value2 = q1 * ty * 0.000001
                        wsoutput.Range("L" & m).Value2 = q2 * ty * 0.000001
                    ElseIf Reserve = True Then
                        wsoutput.Range("K" & m).Value2 = q2 * ty * 0.000001
                        wsoutput.Range("L" & m).Value2 = q1 * ty * 0.000001
                    End If
                ElseIf Mark_Section = "Col" Then
                    If Reserve = False Then
                        wsoutput.Range("G" & m).Value2 = Abs(elestart - z1) / Length
                        wsoutput.Range("H" & m).Value2 = Abs(ele2 - z1) / Length
                        wsoutput.Range("I" & m).Value2 = Abs(elestart - z1)
                        wsoutput.Range("J" & m).Value2 = Abs(ele2 - z1)
                        wsoutput.Range("K" & m).Value2 = q1 * ty * 0.000001
                        wsoutput.Range("L" & m).Value2 = q2 * ty * 0.000001
                    Else
                        wsoutput.Range("G" & m).Value2 = Abs(ele2 - z1) / Length
                        wsoutput.Range("H" & m).Value2 = Abs(elestart - z1) / Length
                        wsoutput.Range("I" & m).Value2 = Abs(ele2 - z1)
                        wsoutput.Range("J" & m).Value2 = Abs(elestart - z1)
                        wsoutput.Range("K" & m).Value2 = q2 * ty * 0.000001
                        wsoutput.Range("L" & m).Value2 = q1 * ty * 0.000001
                    End If
                ElseIf Mark_Section = "Beam" Then
                    wsoutput.Range("G" & m).Value2 = 0
                    wsoutput.Range("H" & m).Value2 = 1
                    wsoutput.Range("I" & m).Value2 = 0
                    wsoutput.Range("J" & m).Value2 = Length
                    wsoutput.Range("K" & m).Value2 = q1 * tx * 0.000001
                    wsoutput.Range("L" & m).Value2 = q2 * tx * 0.000001
                End If
                'wsoutput.Range("M" & m).Value2 = "WindStr" & "-" & wsgeo.Range("A" & n).Value2 & "-" & NameLoadCaseX & "-" & Round((q1 + q2) / 2, 2) & "-" & Round(ele1, 2)
                m = m + 1
            End If
            If CountY <> "" Then
                wsoutput.Range("A" & m).Value2 = wsgeo.Range("A" & n).Value2
                wsoutput.Range("B" & m).Value2 = NameLoadCaseY
                wsoutput.Range("C" & m).Value2 = "GLOBAL"
                wsoutput.Range("D" & m).Value2 = "FORCE"
                wsoutput.Range("E" & m).Value2 = "Y"
                wsoutput.Range("F" & m).Value2 = "RelDist"
                If (ele1 < elechange And ele2 > elechange) And Mark_Section = "Col" Then
                    If Reserve = False Then
                        wsoutput.Range("G" & m).Value2 = Abs(elestart - z1) / Length
                        wsoutput.Range("H" & m).Value2 = Abs(elechange - z1) / Length
                        wsoutput.Range("I" & m).Value2 = Abs(elestart - z1)
                        wsoutput.Range("J" & m).Value2 = Abs(elechange - z1)
                        wsoutput.Range("K" & m).Value2 = q1 * tx * 0.000001
                        wsoutput.Range("L" & m).Value2 = q1 * tx * 0.000001
                        wsoutput.Range("A" & m & ":F" & m + 1).FillDown
                        m = m + 1
                        wsoutput.Range("G" & m).Value2 = Abs(elechange - z1) / Length
                        wsoutput.Range("H" & m).Value2 = Abs(ele2 - z1) / Length
                        wsoutput.Range("I" & m).Value2 = Abs(elechange - z1)
                        wsoutput.Range("J" & m).Value2 = Abs(ele2 - z1)
                        wsoutput.Range("K" & m).Value2 = q1 * tx * 0.000001
                        wsoutput.Range("L" & m).Value2 = q2 * tx * 0.000001
                    Else
                        wsoutput.Range("G" & m).Value2 = Abs(elechange - z1) / Length
                        wsoutput.Range("H" & m).Value2 = Abs(elestart - z1) / Length
                        wsoutput.Range("I" & m).Value2 = Abs(elechange - z1)
                        wsoutput.Range("J" & m).Value2 = Abs(elestart - z1)
                        wsoutput.Range("K" & m).Value2 = q1 * tx * 0.000001
                        wsoutput.Range("L" & m).Value2 = q1 * tx * 0.000001
                        wsoutput.Range("A" & m & ":F" & m + 1).FillDown
                        m = m + 1
                        wsoutput.Range("G" & m).Value2 = Abs(ele2 - z1) / Length
                        wsoutput.Range("H" & m).Value2 = Abs(elechange - z1) / Length
                        wsoutput.Range("I" & m).Value2 = Abs(ele2 - z1)
                        wsoutput.Range("J" & m).Value2 = Abs(elechange - z1)
                        wsoutput.Range("K" & m).Value2 = q2 * tx * 0.000001
                        wsoutput.Range("L" & m).Value2 = q1 * tx * 0.000001
                    End If
                ElseIf Mark_Section = "VBrace" Then
                    If Reserve = False Then
                        wsoutput.Range("K" & m).Value2 = q1 * tx * 0.000001
                        wsoutput.Range("L" & m).Value2 = q2 * tx * 0.000001
                    ElseIf Reserve = True Then
                        wsoutput.Range("K" & m).Value2 = q2 * tx * 0.000001
                        wsoutput.Range("L" & m).Value2 = q1 * tx * 0.000001
                    End If
                ElseIf Mark_Section = "Col" Then
                    If Reserve = False Then
                        wsoutput.Range("G" & m).Value2 = Abs(elestart - z1) / Length
                        wsoutput.Range("H" & m).Value2 = Abs(ele2 - z1) / Length
                        wsoutput.Range("I" & m).Value2 = Abs(elestart - z1)
                        wsoutput.Range("J" & m).Value2 = Abs(ele2 - z1)
                        wsoutput.Range("K" & m).Value2 = q1 * tx * 0.000001
                        wsoutput.Range("L" & m).Value2 = q2 * tx * 0.000001
                    Else
                        wsoutput.Range("G" & m).Value2 = Abs(ele2 - z1) / Length
                        wsoutput.Range("H" & m).Value2 = Abs(elestart - z1) / Length
                        wsoutput.Range("I" & m).Value2 = Abs(ele2 - z1)
                        wsoutput.Range("J" & m).Value2 = Abs(elestart - z1)
                        wsoutput.Range("K" & m).Value2 = q2 * tx * 0.000001
                        wsoutput.Range("L" & m).Value2 = q1 * tx * 0.000001
                    End If
                ElseIf Mark_Section = "Beam" Then
                    wsoutput.Range("G" & m).Value2 = 0
                    wsoutput.Range("H" & m).Value2 = 1
                    wsoutput.Range("I" & m).Value2 = 0
                    wsoutput.Range("J" & m).Value2 = Length
                    wsoutput.Range("K" & m).Value2 = q1 * tx * 0.000001
                    wsoutput.Range("L" & m).Value2 = q2 * tx * 0.000001
                End If
                'wsoutput.Range("M" & m).Value2 = "WindStr" & "-" & wsgeo.Range("A" & n).Value2 & "-" & NameLoadCaseX & "-" & Round((q1 + q2) / 2, 2) & "-" & Round(ele1, 2)
                m = m + 1
            End If
        End If
        n = n + 1
    Loop
    
    '''
    j = 3
    For Each item In ListX.keys
        temp = Split(item, "|")
        wscal.Range("A" & j).Value2 = temp(0)
        wscal.Range("E" & j).Value2 = temp(1)
        '''
        wscal.Range("K" & j).Value2 = T3_Section(temp(2))
        wscal.Range("N" & j).Value2 = T2_Section(temp(2))
        wscal.Range("M" & j).Value2 = "x"
        wscal.Range("I" & j).Value2 = Name_Section(temp(2))
        '''
        wscal.Range("R" & j).Value2 = temp(3)
        wscal.Range("U" & j).Value2 = temp(5)
        wscal.Range("Y" & j).Value2 = CDbl(temp(3)) * CDbl(temp(4)) * 0.000001
        wscal.Range("AB" & j).Value2 = ListX(item)
        wscal.Range("AD" & j).FormulaR1C1 = "=RC[-9]*RC[-5]*RC[-2]"
        j = j + 1
    Next
    j = 3
    For Each item In ListY.keys
        temp = Split(item, "|")
        wscal.Range("AH" & j).Value2 = temp(0)
        wscal.Range("AL" & j).Value2 = temp(1)
        '''
        wscal.Range("AR" & j).Value2 = T3_Section(temp(2))
        wscal.Range("AU" & j).Value2 = T2_Section(temp(2))
        wscal.Range("AT" & j).Value2 = "x"
        wscal.Range("AP" & j).Value2 = Name_Section(temp(2))
        '''
        wscal.Range("AY" & j).Value2 = temp(3)
        wscal.Range("BB" & j).Value2 = temp(5)
        wscal.Range("BF" & j).Value2 = CDbl(temp(3)) * CDbl(temp(4)) * 0.000001
        wscal.Range("BI" & j).Value2 = ListY(item)
        wscal.Range("BK" & j).FormulaR1C1 = "=RC[-9]*RC[-5]*RC[-2]"
        j = j + 1
    Next
    wscal.rows("3:3").Copy
    wscal.rows("4:" & Application.Max(ListY.count, ListX.count, 4) + 2).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    MsgBox ("DONE")
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub


