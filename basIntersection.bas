Attribute VB_Name = "basIntersection"
'
' Algebra taken from various sources on the WWW
'
Option Explicit
Public Function IntersectComplex(x1 As Double, y1 As Double, x2 As Double, y2 As Double, LineCoordinates As Range, axis As Boolean) As Variant
Attribute IntersectComplex.VB_Description = "Calculates whether a line segments intersects a line.\nAxis=True returns X coord Axis=False returns Y coord"
Attribute IntersectComplex.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Complex Intersect.
' Because the line segments are not uniformly spaced the (xy,y1)(x2,y2) could cross
' at any point along the other line
'
' Return
' If intersection
'    requested coordinate
' else
'    nothing
' endif
' Axis=True returns X value
' Axis=False returns Y value
'
    Dim dblCrossX As Double
    Dim dblCrossY As Double
    Dim dblTestx1 As Double
    Dim dblTesty1 As Double
    Dim dblTestx2 As Double
    Dim dblTesty2 As Double
    Dim intSegment As Integer
    
    With LineCoordinates
        For intSegment = 1 To .rows.count - 1
            dblTestx1 = .Cells(intSegment, 1)
            dblTesty1 = .Cells(intSegment, 2)
            dblTestx2 = .Cells(intSegment + 1, 1)
            dblTesty2 = .Cells(intSegment + 1, 2)
            If m_CalculateIntersection(x1, y1, x2, y2, dblTestx1, dblTesty1, dblTestx2, dblTesty2, dblCrossX, dblCrossY) Then
                If axis Then
                    IntersectComplex = dblCrossX
                Else
                    IntersectComplex = dblCrossY
                End If
                Exit Function
            End If
        Next
    
        ' Special check for last pairing
        intSegment = .rows.count
        dblTestx1 = .Cells(intSegment, 1)
        dblTesty1 = .Cells(intSegment, 2)
        dblTestx2 = .Cells(intSegment, 1)
        dblTesty2 = .Cells(intSegment, 2)
        If m_CalculateIntersection(x1, y1, x2, y2, dblTestx1, dblTesty1, dblTestx2, dblTesty2, dblCrossX, dblCrossY) Then
            If axis Then
                IntersectComplex = dblCrossX
            Else
                IntersectComplex = dblCrossY
            End If
            Exit Function
        End If
        
    End With
    IntersectComplex = CVErr(xlErrNA)    ' Null
    
End Function
Private Function m_CalculateIntersection(x1 As Double, y1 As Double, x2 As Double, y2 As Double, _
    x3 As Double, y3 As Double, x4 As Double, y4 As Double, _
    ByRef CrossX As Double, ByRef CrossY As Double) As Variant

'Call with x1,y1,x2,y2,x3,y3,x4,y4 and returns intersect,x,y
'
'Where:
' x1,y1,x2,y2,x3,y3,x4,y4 are the end points of two line segments
'Returns:
' intersect is true/false, and x,y is the interecting point if intersect is true
'
'Description:
'
'Equations for the lines are:
' Pa = P1 + Ua(P2 - P1)
' Pb = P3 + Ub(P4 - P3)
'
'Solving for the point where Pa = Pb gives the following equations for ua and ub
'
' Ua = ((x4 - x3) * (y1 - y3) - (y4 - y3 ) * (x1 - x3)) / ((y4 - y3) * (x2 - x1)
'     - (x4 - x3) * (y2 - y1))
' Ub = ((x2 - x1) * (y1 - y3) - (y2 - y1 ) * (x1 - x3)) / ((y4 - y3) * (x2 - x1)
'     - (x4 - x3) * (y2 - y1))
'
'Substituting either of these into the corresponding equation for the line gives
'     the intersection point.
'For example the intersection point (x,y) is
' x = x1 + Ua(x2 - x1)
' y = y1 + Ua(y2 - y1)
'
'Notes:
' - The denominators are the same.
'
' - If the denominator above is 0 then the two lines are parallel.
'
' - If the denominator and numerator are 0 then the two lines are coincident.
'
' - The equations above apply to lines,
'     if the intersection of line segments is
'     required then it is only necessary to test if ua and ub lie between 0 and 1.
'     Whichever one lies within that range then the corresponding line segment
'     contains the intersection point. If both lie within the range of 0 to 1 then
'     the intersection point is within both line segments.
'
    Dim dblDenominator As Double
    Dim dblUa As Double
    Dim dblUb As Double
    'Pre calc the denominator, if zero then
    '     both lines are parallel and there is no
    '     intersection
    dblDenominator = ((y4 - y3) * (x2 - x1) - (x4 - x3) * (y2 - y1))

    If dblDenominator <> 0 Then
        'Solve for the simultaneous equations
        dblUa = ((x4 - x3) * (y1 - y3) - (y4 - y3) * (x1 - x3)) / dblDenominator
        dblUb = ((x2 - x1) * (y1 - y3) - (y2 - y1) * (x1 - x3)) / dblDenominator
    Else
    
        If (x1 = x3) And (y1 = y3) Then
            CrossX = x1
            CrossY = y1
            m_CalculateIntersection = True
        Else
            m_CalculateIntersection = False
        End If
        Exit Function
    End If
    
    'Could the lines intersect?
    If dblUa >= 0 And dblUa <= 1 And dblUb >= 0 And dblUb <= 1 Then
        'Calculate the intersection point
        CrossX = x1 + dblUa * (x2 - x1)
        CrossY = y1 + dblUa * (y2 - y1)
        'Yes, they do
        m_CalculateIntersection = True
    Else
        'No, they do not
        m_CalculateIntersection = False
    End If
    
End Function
