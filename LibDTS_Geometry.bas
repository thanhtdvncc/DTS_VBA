Attribute VB_Name = "LibDTS_Geometry"
' Module: LibDTS_Geometry
Option Explicit

Private Const PI As Double = 3.14159265358979

' Calculate distance between two 3D coordinates
' Tinh khoang cach giua hai toa do 3D
Public Function Distance(x1 As Double, y1 As Double, z1 As Double, _
                         x2 As Double, y2 As Double, z2 As Double) As Double
    Distance = Sqr((x2 - x1) ^ 2 + (y2 - y1) ^ 2 + (z2 - z1) ^ 2)
End Function

' Calculate Midpoint
' Tinh trung diem
Public Function GetMidPoint(p1 As clsDTSPoint, p2 As clsDTSPoint) As clsDTSPoint
    Dim midPt As New clsDTSPoint
    midPt.Init (p1.X + p2.X) / 2, (p1.Y + p2.Y) / 2, (p1.Z + p2.Z) / 2
    Set GetMidPoint = midPt
End Function

' Calculate Angle in Radians (XY Plane)
' Tinh goc theo Radian (Mat phang XY)
Public Function GetAngleXY(p1 As clsDTSPoint, p2 As clsDTSPoint) As Double
    Dim dx As Double, dy As Double
    dx = p2.X - p1.X
    dy = p2.Y - p1.Y
    
    If Abs(dx) < LibDTS_Global.DTS_PRECISION Then
        GetAngleXY = PI / 2 ' 90 degrees
        If dy < 0 Then GetAngleXY = 3 * PI / 2 ' 270 degrees
    Else
        GetAngleXY = Atn(dy / dx)
    End If
End Function

' Check if lines intersect (2D)
' Kiem tra hai duong thang co cat nhau khong (2D)
Public Function GetIntersection2D(x1 As Double, y1 As Double, x2 As Double, y2 As Double, _
                                  x3 As Double, y3 As Double, x4 As Double, y4 As Double) As Variant
    Dim d As Double
    d = (x1 - x2) * (y3 - y4) - (y1 - y2) * (x3 - x4)
    
    If Abs(d) < LibDTS_Global.DTS_PRECISION Then
        GetIntersection2D = Empty ' Parallel
        Exit Function
    End If
    
    Dim px As Double, py As Double
    px = ((x1 * y2 - y1 * x2) * (x3 - x4) - (x1 - x2) * (x3 * y4 - y3 * x4)) / d
    py = ((x1 * y2 - y1 * x2) * (y3 - y4) - (y1 - y2) * (x3 * y4 - y3 * x4)) / d
    
    GetIntersection2D = Array(px, py)
End Function

' Offset a point by angle and distance
' Dich chuyen diem theo goc va khoang cach
Public Function PolarPoint(origin As clsDTSPoint, angle As Double, dist As Double) As clsDTSPoint
    Dim p As New clsDTSPoint
    p.Init origin.X + Cos(angle) * dist, origin.Y + Sin(angle) * dist, origin.Z
    Set PolarPoint = p
End Function

