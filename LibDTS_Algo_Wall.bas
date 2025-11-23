Attribute VB_Name = "LibDTS_Algo_Wall"
' ==============================================================================
' Module: LibDTS_Algo_Wall
' Purpose: Pure algorithmic logic for wall-to-frame synchronization
' Architecture: Logic Layer (MVVM)
' Dependencies: Uses Core Classes (clsDTSFrame, clsDTSElement) only
' ==============================================================================
Option Explicit

' --- CONSTANTS ---
Private Const COORD_TOLERANCE As Double = 250    ' mm - proximity detection
Private Const ANGLE_TOLERANCE As Double = 0.174533  ' 10 degrees in radians
Private Const MIN_OVERLAP_RATIO As Double = 0.15
Private Const PI As Double = 3.14159265358979

' --- MODULE VARIABLES ---
Private m_LastError As String

' ==============================================================================
' PUBLIC FUNCTION: GetSelectionAsFrames
' Purpose: Convert selected AutoCAD entities to Collection of clsDTSFrame
' Input: acadDoc - AutoCAD Document object
' Output: Collection of clsDTSFrame objects
' ==============================================================================
Public Function GetSelectionAsFrames(acadDoc As Object) As Collection
    On Error GoTo ErrHandler
    
    Set GetSelectionAsFrames = New Collection
    
    ' Validate input
    If acadDoc Is Nothing Then
        m_LastError = "AutoCAD Document is Nothing"
        LibDTS_Logger.Log "LibDTS_Algo_Wall.GetSelectionAsFrames: " & m_LastError, DTS_ERROR
        Exit Function
    End If
    
    ' Prompt user to select entities
    Dim ssName As String
    ssName = "WALL_SEL_" & Format(Now, "hhmmss")
    
    ' Clean up any existing selection set
    On Error Resume Next
    acadDoc.SelectionSets.item(ssName).Delete
    On Error GoTo ErrHandler
    
    Dim ss As Object
    Set ss = acadDoc.SelectionSets.Add(ssName)
    
    ' Filter for lines on DTS_WALL_DIAGRAM layer
    Dim gpCode(0 To 1) As Integer
    Dim dataVal(0 To 1) As Variant
    gpCode(0) = 0: dataVal(0) = "LINE"
    gpCode(1) = 8: dataVal(1) = "DTS_WALL_DIAGRAM"
    
    On Error Resume Next
    acadDoc.Utility.prompt vbCrLf & "Select wall lines to process: "
    ss.SelectOnScreen gpCode, dataVal
    On Error GoTo ErrHandler
    
    If ss.count = 0 Then
        ss.Delete
        Exit Function
    End If
    
    ' Convert each entity to clsDTSFrame
    Dim i As Long
    For i = 0 To ss.count - 1
        Dim ent As Object
        Set ent = ss.item(i)
        
        ' Parse entity using LibDTS_DriverCAD
        Dim frame As clsDTSFrame
        Set frame = LibDTS_DriverCAD.ParseFrame(ent)
        
        If Not frame Is Nothing Then
            ' Validate frame
            If frame.IsValid Then
                GetSelectionAsFrames.Add frame
            End If
        End If
    Next i
    
    ss.Delete
    
    LibDTS_Logger.Log "LibDTS_Algo_Wall.GetSelectionAsFrames: Processed " & GetSelectionAsFrames.count & " frames", DTS_INFO
    Exit Function
    
ErrHandler:
    m_LastError = "GetSelectionAsFrames error: " & Err.Description
    LibDTS_Logger.Log "LibDTS_Algo_Wall.GetSelectionAsFrames: " & m_LastError, DTS_ERROR
    On Error Resume Next
    If Not ss Is Nothing Then ss.Delete
End Function

' ==============================================================================
' PUBLIC FUNCTION: AnalyzeWallConnections
' Purpose: Analyze spatial relationships between wall frames
' Algorithm: Uses LibDTS_Geometry for intersection/overlap detection
' Input: frames - Collection of clsDTSFrame objects
' Output: Updates clsDTSFrame.Properties with connection info
' ==============================================================================
Public Function AnalyzeWallConnections(frames As Collection) As Boolean
    On Error GoTo ErrHandler
    
    AnalyzeWallConnections = False
    
    ' Validate input
    If frames Is Nothing Then
        m_LastError = "Frames collection is Nothing"
        LibDTS_Logger.Log "LibDTS_Algo_Wall.AnalyzeWallConnections: " & m_LastError, DTS_ERROR
        Exit Function
    End If
    
    If frames.count = 0 Then
        AnalyzeWallConnections = True
        Exit Function
    End If
    
    ' Iterate through all frame pairs
    Dim i As Long, j As Long
    For i = 1 To frames.count - 1
        Dim frame1 As clsDTSFrame
        Set frame1 = frames(i)
        
        For j = i + 1 To frames.count
            Dim frame2 As clsDTSFrame
            Set frame2 = frames(j)
            
            ' Check if frames are parallel
            If AreFramesParallel(frame1, frame2) Then
                ' Check for overlap
                If CheckFrameOverlap(frame1, frame2) Then
                    ' Mark both frames as overlapping
                    frame1.Base.Properties("HasOverlap") = True
                    frame1.Base.Properties("OverlapWith") = frame2.Base.guid
                    
                    frame2.Base.Properties("HasOverlap") = True
                    frame2.Base.Properties("OverlapWith") = frame1.Base.guid
                    
                    LibDTS_Logger.Log "LibDTS_Algo_Wall: Overlap detected between " & frame1.Base.guid & " and " & frame2.Base.guid, DTS_WARNING
                End If
            End If
        Next j
    Next i
    
    AnalyzeWallConnections = True
    Exit Function
    
ErrHandler:
    m_LastError = "AnalyzeWallConnections error: " & Err.Description
    LibDTS_Logger.Log "LibDTS_Algo_Wall.AnalyzeWallConnections: " & m_LastError, DTS_ERROR
    AnalyzeWallConnections = False
End Function

' ==============================================================================
' PUBLIC SUB: SyncFramesToSAP
' Purpose: Synchronize frame collection to SAP2000
' Uses: clsDTSRepository.SyncToSAP for each frame
' Input: frames - Collection of clsDTSFrame objects
' ==============================================================================
Public Sub SyncFramesToSAP(frames As Collection)
    On Error GoTo ErrHandler
    
    ' Validate input
    If frames Is Nothing Then
        m_LastError = "Frames collection is Nothing"
        LibDTS_Logger.Log "LibDTS_Algo_Wall.SyncFramesToSAP: " & m_LastError, DTS_ERROR
        Exit Sub
    End If
    
    If frames.count = 0 Then
        LibDTS_Logger.Log "LibDTS_Algo_Wall.SyncFramesToSAP: No frames to sync", DTS_INFO
        Exit Sub
    End If
    
    ' Get repository instance
    Dim repo As clsDTSRepository
    Set repo = New clsDTSRepository
    
    ' Initialize repository (assumes global CAD doc is available)
    Dim acadDoc As Object
    On Error Resume Next
    Set acadDoc = GetObject(, "AutoCAD.Application").ActiveDocument
    On Error GoTo ErrHandler
    
    If acadDoc Is Nothing Then
        m_LastError = "Cannot connect to AutoCAD"
        LibDTS_Logger.Log "LibDTS_Algo_Wall.SyncFramesToSAP: " & m_LastError, DTS_ERROR
        Exit Sub
    End If
    
    repo.Initialize acadDoc
    
    ' Sync each frame
    Dim successCount As Long
    Dim failCount As Long
    successCount = 0
    failCount = 0
    
    Dim frame As Variant
    For Each frame In frames
        ' Call Repository to sync to SAP
        If repo.SyncToSAP(frame) Then
            successCount = successCount + 1
            
            ' Update frame properties
            frame.Base.Properties("SyncedToSAP") = True
            frame.Base.Properties("SyncTime") = Now
        Else
            failCount = failCount + 1
            frame.Base.Properties("SyncedToSAP") = False
        End If
    Next frame
    
    LibDTS_Logger.Log "LibDTS_Algo_Wall.SyncFramesToSAP: Synced " & successCount & " frames, Failed " & failCount, DTS_INFO
    Exit Sub
    
ErrHandler:
    m_LastError = "SyncFramesToSAP error: " & Err.Description
    LibDTS_Logger.Log "LibDTS_Algo_Wall.SyncFramesToSAP: " & m_LastError, DTS_ERROR
End Sub

' ==============================================================================
' PRIVATE HELPER: AreFramesParallel
' Purpose: Check if two frames are parallel (within angle tolerance)
' ==============================================================================
Private Function AreFramesParallel(frame1 As clsDTSFrame, frame2 As clsDTSFrame) As Boolean
    On Error Resume Next
    
    AreFramesParallel = False
    
    ' Calculate direction vectors
    Dim dx1 As Double, dy1 As Double
    dx1 = frame1.EndPoint.X - frame1.StartPoint.X
    dy1 = frame1.EndPoint.Y - frame1.StartPoint.Y
    
    Dim dx2 As Double, dy2 As Double
    dx2 = frame2.EndPoint.X - frame2.StartPoint.X
    dy2 = frame2.EndPoint.Y - frame2.StartPoint.Y
    
    ' Calculate angles
    Dim angle1 As Double, angle2 As Double
    angle1 = Atan2(dy1, dx1)
    angle2 = Atan2(dy2, dx2)
    
    ' Calculate angle difference
    Dim angleDiff As Double
    angleDiff = Abs(angle1 - angle2)
    
    ' Normalize to [0, PI]
    If angleDiff > PI Then angleDiff = 2 * PI - angleDiff
    
    ' Check if parallel (or opposite direction)
    If angleDiff <= ANGLE_TOLERANCE Or Abs(angleDiff - PI) <= ANGLE_TOLERANCE Then
        AreFramesParallel = True
    End If
End Function

' ==============================================================================
' PRIVATE HELPER: CheckFrameOverlap
' Purpose: Check if two parallel frames overlap in space
' ==============================================================================
Private Function CheckFrameOverlap(frame1 As clsDTSFrame, frame2 As clsDTSFrame) As Boolean
    On Error Resume Next
    
    CheckFrameOverlap = False
    
    ' Calculate perpendicular distance between frames
    Dim dist As Double
    dist = PointToLineDistance( _
        frame2.StartPoint.X, frame2.StartPoint.Y, _
        frame1.StartPoint.X, frame1.StartPoint.Y, _
        frame1.EndPoint.X, frame1.EndPoint.Y)
    
    ' If frames are too far apart, no overlap
    If dist > COORD_TOLERANCE Then Exit Function
    
    ' Project frame2 endpoints onto frame1 axis
    Dim t1 As Double, t2 As Double
    Dim dx As Double, dy As Double, len As Double
    
    dx = frame1.EndPoint.X - frame1.StartPoint.X
    dy = frame1.EndPoint.Y - frame1.StartPoint.Y
    len = Sqr(dx * dx + dy * dy)
    
    If len < LibDTS_Global.DTS_PRECISION Then Exit Function
    
    ' Unit vector
    dx = dx / len
    dy = dy / len
    
    ' Project frame2.StartPoint
    t1 = (frame2.StartPoint.X - frame1.StartPoint.X) * dx + _
         (frame2.StartPoint.Y - frame1.StartPoint.Y) * dy
    
    ' Project frame2.EndPoint
    t2 = (frame2.EndPoint.X - frame1.StartPoint.X) * dx + _
         (frame2.EndPoint.Y - frame1.StartPoint.Y) * dy
    
    ' Ensure t1 < t2
    If t1 > t2 Then
        Dim temp As Double
        temp = t1: t1 = t2: t2 = temp
    End If
    
    ' Calculate overlap
    Dim overlapStart As Double, overlapEnd As Double
    If t1 > 0 Then
        overlapStart = t1
    Else
        overlapStart = 0
    End If
    
    If t2 < len Then
        overlapEnd = t2
    Else
        overlapEnd = len
    End If
    
    Dim overlapLen As Double
    overlapLen = overlapEnd - overlapStart
    
    ' Check if overlap is significant
    If overlapLen > len * MIN_OVERLAP_RATIO Then
        CheckFrameOverlap = True
    End If
End Function

' ==============================================================================
' PRIVATE HELPER: PointToLineDistance
' Purpose: Calculate perpendicular distance from point to line
' ==============================================================================
Private Function PointToLineDistance(px As Double, py As Double, _
                                     x1 As Double, y1 As Double, _
                                     x2 As Double, y2 As Double) As Double
    Dim dx As Double, dy As Double, len2 As Double
    dx = x2 - x1
    dy = y2 - y1
    len2 = dx * dx + dy * dy
    
    If len2 < LibDTS_Global.DTS_PRECISION Then
        PointToLineDistance = Sqr((px - x1) ^ 2 + (py - y1) ^ 2)
    Else
        PointToLineDistance = Abs(dy * (px - x1) - dx * (py - y1)) / Sqr(len2)
    End If
End Function

' ==============================================================================
' PRIVATE HELPER: Atan2
' Purpose: Calculate arctangent of y/x
' ==============================================================================
Private Function Atan2(Y As Double, X As Double) As Double
    If X > 0 Then
        Atan2 = Atn(Y / X)
    ElseIf X < 0 And Y >= 0 Then
        Atan2 = Atn(Y / X) + PI
    ElseIf X < 0 And Y < 0 Then
        Atan2 = Atn(Y / X) - PI
    ElseIf X = 0 And Y > 0 Then
        Atan2 = PI / 2
    ElseIf X = 0 And Y < 0 Then
        Atan2 = -PI / 2
    Else
        Atan2 = 0
    End If
End Function

' ==============================================================================
' PUBLIC PROPERTY: GetLastError
' Purpose: Retrieve last error message for debugging
' ==============================================================================
Public Property Get LastError() As String
    LastError = m_LastError
End Property
