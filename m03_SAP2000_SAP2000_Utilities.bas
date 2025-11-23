Attribute VB_Name = "m03_SAP2000_SAP2000_Utilities"
Option Explicit
'===============================================================
' Module: modSAP2000_Utilities
' Purpose: Internal helper utilities not already in Helpers
'===============================================================

Public Sub CheckRet(ByVal ret As Long, ByVal ctx As String)
    If ret <> 0 Then LogMsg "Non-zero ret=" & ret & " context=" & ctx
End Sub

Public Function SheetOrCreate(ByVal sheetName As String, Optional allowCreate As Boolean = True) As Worksheet
    On Error Resume Next
    Set SheetOrCreate = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If SheetOrCreate Is Nothing And allowCreate Then
        On Error Resume Next
        Set SheetOrCreate = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        If Not SheetOrCreate Is Nothing Then SheetOrCreate.Name = sheetName
        On Error GoTo 0
    End If
End Function

Public Function ToVerticalVariant(ByVal arr As Variant) As Variant
    If Not IsArray(arr) Then
        Dim singleVal(1 To 1, 1 To 1)
        singleVal(1, 1) = arr
        ToVerticalVariant = singleVal
        Exit Function
    End If
    If IsArrayEmpty(arr) Then
        Dim emptyVal(1 To 1, 1 To 1)
        emptyVal(1, 1) = ""
        ToVerticalVariant = emptyVal
        Exit Function
    End If
    Dim lb As Long, ub As Long, i As Long
    lb = LBound(arr): ub = UBound(arr)
    Dim outArr() As Variant
    ReDim outArr(1 To ub - lb + 1, 1 To 1)
    For i = lb To ub
        outArr(i - lb + 1, 1) = arr(i)
    Next
    ToVerticalVariant = outArr
End Function

Public Function SafeTranspose(ByVal arr As Variant) As Variant
    If Not IsArrayEmpty(arr) Then
        On Error Resume Next
        SafeTranspose = Application.WorksheetFunction.Transpose(arr)
        If err.number <> 0 Then
            err.Clear
            SafeTranspose = ToVerticalVariant(arr)
        End If
        On Error GoTo 0
    Else
        Dim tmp(1 To 1, 1 To 1)
        tmp(1, 1) = ""
        SafeTranspose = tmp
    End If
End Function



