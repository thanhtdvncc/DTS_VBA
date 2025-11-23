Attribute VB_Name = "Database"

'Interpolates a lookup table vertically
Public Function VInterpolate(LookupValue As Double, LookupTable As Range, Index As Integer) As Double
    
    Dim x1 As Double, x2 As Double
    Dim y1 As Double, y2 As Double
    Dim i As Integer
    
    For i = 1 To LookupTable.rows.count - 1
        
        If LookupValue >= LookupTable.Cells(i, 1) And LookupValue <= LookupTable.Cells(i + 1, 1) Then
            
            x1 = LookupTable.Cells(i, 1)
            x2 = LookupTable.Cells(i + 1, 1)
            y1 = LookupTable.Cells(i, Index)
            y2 = LookupTable.Cells(i + 1, Index)
            
            Exit For
        
        End If
        
    Next i
    
    VInterpolate = (y2 - y1) / (x2 - x1) * (LookupValue - x1) + y1
    
End Function

'Interpolates a lookup table horizontally
Public Function HInterpolate(LookupValue As Double, LookupTable As Range, Index As Integer) As Double
    
    Dim x1 As Double, x2 As Double
    Dim y1 As Double, y2 As Double
    Dim i As Integer
    
    For i = 1 To LookupTable.Columns.count - 1
        
        If LookupValue >= LookupTable.Cells(1, i) And LookupValue <= LookupTable.Cells(1, i + 1) Then
            
            x1 = LookupTable.Cells(1, i)
            x2 = LookupTable.Cells(1, i + 1)
            y1 = LookupTable.Cells(Index, i)
            y2 = LookupTable.Cells(Index, i + 1)
            
            Exit For
        
        End If
        
    Next i
    
    HInterpolate = (y2 - y1) / (x2 - x1) * (LookupValue - x1) + y1
    
End Function

'Interpolates a lookup table in two directions (vertically and then horizontally)
Public Function DualInterpolate(VLookupValue As Double, VLookupTable As Range, HLookupValue As Double, HLookupTable As Range) As Double
    
    Dim x1 As Double, x2 As Double
    Dim y1A As Double, y1B As Double
    Dim y2A As Double, y2B As Double
    Dim y1 As Double, y2 As Double
    Dim i As Integer, j As Integer
    
    For i = 1 To VLookupTable.rows.count - 1
        
        If VLookupValue >= VLookupTable.Cells(i, 1) And VLookupValue <= VLookupTable.Cells(i + 1, 1) Then
            
            For j = 1 To HLookupTable.Columns.count - 1
            
                If HLookupValue >= HLookupTable.Cells(1, j) And HLookupValue <= HLookupTable.Cells(1, j + 1) Then
                    
                    x1 = VLookupTable.Cells(i, 1)
                    x2 = VLookupTable.Cells(i + 1, 1)
                    
                    y1A = HLookupTable.Cells(1 + i, j)
                    y1B = HLookupTable.Cells(1 + i + 1, j)
                    
                    y2A = HLookupTable.Cells(1 + i, j + 1)
                    y2B = HLookupTable.Cells(1 + i + 1, j + 1)
                    
                    y1 = (y1B - y1A) / (x2 - x1) * (VLookupValue - x1) + y1A
                    y2 = (y2B - y2A) / (x2 - x1) * (VLookupValue - x1) + y2A
                    
                    x1 = HLookupTable.Cells(1, j)
                    x2 = HLookupTable.Cells(1, j + 1)
                    
                    Exit For
                    
                End If
                
            Next j
            
            Exit For
        
        End If
        
    Next i
    
    DualInterpolate = (y2 - y1) / (x2 - x1) * (HLookupValue - x1) + y1
    
End Function


