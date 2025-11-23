Attribute VB_Name = "m03_SAP2000_SAP2000_Helper"
'Option Compare Text
Function ArrayToDictionary_KeysBased(ByVal arr_Value As Variant) As Object

    ' Check that array is One Dimensional
    On Error Resume Next
    Dim ret As Long
    ret = -1
    ret = UBound(arr_Value, 2)
    On Error GoTo 0
    If ret <> -1 Then
        err.Raise vbObjectError + 513, "Combine2ArrayToDictionary" _
                                      , "The array can only have one 1 dimension"
    End If

    ' Create the Dictionary
    Dim dic As Object
    Set dic = CreateObject("Scripting.Dictionary")
    'dic.CompareMode = vbTextCompare
    On Error Resume Next
    ' Add items to the Dictionary
    For Each item In arr_Value
        dic.Add item, item
    Next
    
    On Error GoTo 0
    ' Return the new ArrayList
    Set ArrayToDictionary_KeysBased = dic
    
End Function

Function ArrayToDictionary_ValueBased(ByVal arr_Value As Variant) As Object

    ' Check that array is One Dimensional
    On Error Resume Next
    Dim ret As Long
    ret = -1
    ret = UBound(arr_Value, 2)
    On Error GoTo 0
    If ret <> -1 Then
        err.Raise vbObjectError + 513, "Combine2ArrayToDictionary" _
                                      , "The array can only have one 1 dimension"
    End If

    ' Create the Dictionary
    Dim dic As Object
    Set dic = CreateObject("Scripting.Dictionary")
    'dic.CompareMode = vbTextCompare
    On Error Resume Next
    ' Add items to the Dictionary
    Dim i As Long
    i = 0
    For Each item In arr_Value
        dic.Add item, i
        i = i + 1
    Next
    
    On Error GoTo 0
    ' Return the new ArrayList
    Set ArrayToDictionary_ValueBased = dic
    
End Function

Public Function CollectiontoArray(ByVal coll As Collection)
    Dim temp() As Variant
    ReDim temp(coll.count - 1)
    i = 0
    For Each item In coll
        temp(i) = item
        i = i + 1
    Next
    CollectiontoArray = temp
End Function


Function Combine2ArrayToDictionary(ByVal arr_Key As Variant, ByVal arr_Value As Variant) As Object

    ' Check that array is One Dimensional
    On Error Resume Next
    Dim ret As Long
    ret = -1
    ret = Application.Max(UBound(arr_Key, 2), UBound(arr_Value, 2))
    On Error GoTo 0
    If ret <> -1 Then
        err.Raise vbObjectError + 513, "Combine2ArrayToDictionary" _
                                      , "The array can only have one 1 dimension"
    End If

    ' Create the Dictionary
    Dim dic As Object
    Set dic = CreateObject("Scripting.Dictionary")
    'dic.CompareMode = vbTextCompare
    On Error Resume Next
    ' Add items to the Dictionary
    Dim i As Long
    For i = LBound(arr_Key, 1) To UBound(arr_Key, 1)
        dic.Add arr_Key(i), arr_Value(i)
    Next i
    
    On Error GoTo 0
    ' Return the new ArrayList
    Set Combine2ArrayToDictionary = dic
    
End Function
Public Function IsArrayEmpty(arr As Variant) As Boolean
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' IsArrayEmpty
    ' This function tests whether the array is empty (unallocated). Returns TRUE or FALSE.
    '
    ' The VBA IsArray function indicates whether a variable is an array, but it does not
    ' distinguish between allocated and unallocated arrays. It will return TRUE for both
    ' allocated and unallocated arrays. This function tests whether the array has actually
    ' been allocated.
    '
    ' This function is really the reverse of IsArrayAllocated.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim lb As Long
    Dim ub As Long

    err.Clear
    On Error Resume Next
    If IsArray(arr) = False Then
        ' we weren't passed an array, return True
        IsArrayEmpty = True
    End If

    ' Attempt to get the UBound of the array. If the array is
    ' unallocated, an error will occur.
    ub = UBound(arr, 1)
    If (err.number <> 0) Then
        IsArrayEmpty = True
    Else
        ''''''''''''''''''''''''''''''''''''''''''
        ' On rare occassion, under circumstances I
        ' cannot reliably replictate, Err.Number
        ' will be 0 for an unallocated, empty array.
        ' On these occassions, LBound is 0 and
        ' UBound is -1.
        ' To accomodate the weird behavior, test to
        ' see if LB > UB. If so, the array is not
        ' allocated.
        ''''''''''''''''''''''''''''''''''''''''''
        err.Clear
        lb = LBound(arr)
        If lb > ub Then
            IsArrayEmpty = True
        Else
            IsArrayEmpty = False
        End If
    End If
    On Error GoTo 0
End Function

Function BlankRemover(ArrayToCondense As Variant) As Variant()

    Dim ArrayWithoutBlanks() As Variant
    Dim CellsInArray As Long
    Dim ArrayWithoutBlanksIndex As Long

    ArrayWithoutBlanksIndex = 0

    For CellsInArray = LBound(ArrayToCondense) To UBound(ArrayToCondense)

        If ArrayToCondense(CellsInArray) <> "" Then

            ReDim Preserve ArrayWithoutBlanks(ArrayWithoutBlanksIndex)

            ArrayWithoutBlanks(ArrayWithoutBlanksIndex) = ArrayToCondense(CellsInArray)

            ArrayWithoutBlanksIndex = ArrayWithoutBlanksIndex + 1

        End If

    Next CellsInArray

    'ArrayWithoutBlanks = Application.Transpose(ArrayWithoutBlanks)
    BlankRemover = ArrayWithoutBlanks

End Function                                     'BlankRemover

Public Sub QuickSortArray(ByRef SortArray As Variant, Optional lngMin As Long = -1, Optional lngMax As Long = -1, Optional lngColumn As Long = 0)
    On Error Resume Next

    'Sort a 2-Dimensional array

    ' SampleUsage: sort arrData by the contents of column 3
    '
    '   QuickSortArray arrData, , , 3

    '
    'Posted by Jim Rech 10/20/98 Excel.Programming

    'Modifications, Nigel Heffernan:

    '       ' Escape failed comparison with empty variant
    '       ' Defensive coding: check inputs

    Dim i As Long
    Dim j As Long
    Dim varMid As Variant
    Dim arrRowTemp As Variant
    Dim lngColTemp As Long

    If IsEmpty(SortArray) Then
        Exit Sub
    End If
    If InStr(TypeName(SortArray), "()") < 1 Then 'IsArray() is somewhat broken: Look for brackets in the type name
        Exit Sub
    End If
    If lngMin = -1 Then
        lngMin = LBound(SortArray, 1)
    End If
    If lngMax = -1 Then
        lngMax = UBound(SortArray, 1)
    End If
    If lngMin >= lngMax Then                     ' no sorting required
        Exit Sub
    End If

    i = lngMin
    j = lngMax

    varMid = Empty
    varMid = SortArray((lngMin + lngMax) \ 2, lngColumn)

    ' We  send 'Empty' and invalid data items to the end of the list:
    If IsObject(varMid) Then                     ' note that we don't check isObject(SortArray(n)) - varMid *might* pick up a valid default member or property
        i = lngMax
        j = lngMin
    ElseIf IsEmpty(varMid) Then
        i = lngMax
        j = lngMin
    ElseIf IsNull(varMid) Then
        i = lngMax
        j = lngMin
    ElseIf varMid = "" Then
        i = lngMax
        j = lngMin
    ElseIf VarType(varMid) = vbError Then
        i = lngMax
        j = lngMin
    ElseIf VarType(varMid) > 17 Then
        i = lngMax
        j = lngMin
    End If

    While i <= j
        While SortArray(i, lngColumn) < varMid And i < lngMax
            i = i + 1
        Wend
        While varMid < SortArray(j, lngColumn) And j > lngMin
            j = j - 1
        Wend

        If i <= j Then
            ' Swap the rows
            ReDim arrRowTemp(LBound(SortArray, 2) To UBound(SortArray, 2))
            For lngColTemp = LBound(SortArray, 2) To UBound(SortArray, 2)
                arrRowTemp(lngColTemp) = SortArray(i, lngColTemp)
                SortArray(i, lngColTemp) = SortArray(j, lngColTemp)
                SortArray(j, lngColTemp) = arrRowTemp(lngColTemp)
            Next lngColTemp
            Erase arrRowTemp

            i = i + 1
            j = j - 1
        End If
    Wend

    If (lngMin < j) Then Call QuickSortArray(SortArray, lngMin, j, lngColumn)
    If (i < lngMax) Then Call QuickSortArray(SortArray, i, lngMax, lngColumn)
    On Error GoTo 0
End Sub

'redim preserve both dimensions for a multidimension array *ONLY
Public Function ReDimPreserve(aArrayToPreserve, nNewFirstUBound, nNewLastUBound)
    ReDimPreserve = False
    'check if its in array first
    If IsArray(aArrayToPreserve) Then
        'create new array
        ReDim aPreservedArray(nNewFirstUBound, nNewLastUBound)
        'get old lBound/uBound
        nOldFirstUBound = UBound(aArrayToPreserve, 1)
        nOldLastUBound = UBound(aArrayToPreserve, 2)
        'loop through first
        For nFirst = LBound(aArrayToPreserve, 1) To nNewFirstUBound
            For nLast = LBound(aArrayToPreserve, 2) To nNewLastUBound
                'if its in range, then append to new array the same way
                If nOldFirstUBound >= nFirst And nOldLastUBound >= nLast Then
                    aPreservedArray(nFirst, nLast) = aArrayToPreserve(nFirst, nLast)
                End If
            Next
        Next
        'return the array redimmed
        If IsArray(aPreservedArray) Then ReDimPreserve = aPreservedArray
    End If
End Function

Public Function IsInArray(valToBeFound As Variant, arr As Variant) As Boolean
    'DEVELOPER: Ryan Wells (wellsr.com)
    'DESCRIPTION: Function to check if a value is in an array of values
    'INPUT: Pass the function a value to search for and an array of values of any data type.
    'OUTPUT: True if is in array, false otherwise
    Dim element As Variant
    On Error GoTo IsInArrayError:                'array is empty
    For Each element In arr
        If CStr(element) = CStr(valToBeFound) Then
            IsInArray = True
            Exit Function
        End If
    Next element
    Exit Function
IsInArrayError:
    On Error GoTo 0
    IsInArray = False
End Function

Public Function RemoveDuplicatesInArray(ByVal SrcArray As Variant) As Variant
    Set dict = CreateObject("Scripting.Dictionary")
    For Each item In SrcArray
        If Not dict.exists(item) Then: dict.Add item, 0
    Next
    RemoveDuplicatesInArray = dict.keys
    Set dict = Nothing
End Function

Function ArrayToArrayList(arr As Variant) As Object

    ' Check that array is One Dimensional
    On Error Resume Next
    Dim ret As Long
    ret = -1
    ret = UBound(arr, 2)
    On Error GoTo 0
    If ret <> -1 Then
        err.Raise vbObjectError + 513, "ArrayToArrayList" _
                                      , "The array can only have one 1 dimension"
    End If

    ' Create the ArrayList
    Dim coll As Object
    Set coll = CreateObject("System.Collections.ArrayList")
    
    ' Add items to the ArrayList
    Dim i As Long
    'For i = LBound(arr, 1) To UBound(arr, 1)
        'coll.Add arr(i)
    'Next i
    On Error Resume Next
    For Each item In arr
        coll.Add item
    Next
    On Error GoTo 0
    ' Return the new ArrayList
    Set ArrayToArrayList = coll
    
End Function

' returns index of item if found, returns 0 if not found
Public Function IndexOf(ByVal coll As Collection, ByVal item As Variant) As Long
    Dim i As Long
    For i = 1 To coll.count
        If coll(i) = item Then
            IndexOf = i
            Exit Function
        End If
    Next
End Function

Public Function ExistsInCollection(col As Collection, key As Variant) As Boolean
    On Error GoTo err
    ExistsInCollection = True
    IsObject (col.item(key))
    Exit Function
err:
    ExistsInCollection = False
End Function

Function IsInCollection(oCollection As Collection, sItem As Variant) As Boolean
    Dim vItem As Variant
    For Each vItem In oCollection
        If vItem = sItem Then
            IsInCollection = True
            Exit Function
        End If
    Next vItem
    IsInCollection = False
End Function

Function Number2Letter(ColumnNumber As Variant) As String
    'PURPOSE: Convert a given number into it's corresponding Letter Reference
    'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault
    'Convert To Column Letter
    Number2Letter = Split(Cells(1, ColumnNumber).Address, "$")(1)
End Function


