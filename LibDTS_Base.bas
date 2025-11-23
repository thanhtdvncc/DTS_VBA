Attribute VB_Name = "LibDTS_Base"
' Module: LibDTS_Base
' Purpose: Base utilities for GUID generation and JSON handling
' Dependencies: Microsoft Scripting Runtime (for Dictionary)
Option Explicit

' ==========================================
' GUID GENERATION
' ==========================================

' Generate a new GUID using Scriptlet.TypeLib
Public Function GenerateGUID() As String
    On Error GoTo Fallback
    
    Dim typeLib As Object
    Set typeLib = CreateObject("Scriptlet.TypeLib")
    GenerateGUID = Left$(typeLib.guid, 38) ' Remove trailing braces
    Exit Function
    
Fallback:
    ' Fallback: Generate pseudo-GUID using timestamp + random
    Dim timestamp As String
    Dim random1 As String
    Dim random2 As String
    
    timestamp = Format$(Now, "yyyymmddhhnnss")
    random1 = Format$(Int(Rnd * 10000), "0000")
    random2 = Format$(Int(Rnd * 10000), "0000")
    
    GenerateGUID = "{" & timestamp & "-" & random1 & "-" & random2 & "}"
End Function

' ==========================================
' JSON UTILITIES
' ==========================================

' Parse JSON string to Dictionary/Collection
' Simple JSON parser without external dependencies
Public Function ParseJson(jsonStr As String) As Object
    On Error GoTo ErrHandler
    
    ' Remove whitespace
    jsonStr = Trim$(jsonStr)
    
    ' Check if it's an object or array
    If Left$(jsonStr, 1) = "{" Then
        Set ParseJson = ParseJsonObject(jsonStr)
    ElseIf Left$(jsonStr, 1) = "[" Then
        Set ParseJson = ParseJsonArray(jsonStr)
    Else
        ' Return empty dictionary for invalid JSON
        Set ParseJson = CreateObject("Scripting.Dictionary")
    End If
    Exit Function
    
ErrHandler:
    LibDTS_Logger.Log "Error parsing JSON: " & err.description, DTS_ERROR
    Set ParseJson = CreateObject("Scripting.Dictionary")
End Function

' Convert Dictionary/Collection to JSON string
Public Function ToJson(obj As Object) As String
    On Error GoTo ErrHandler
    
    If TypeName(obj) = "Dictionary" Then
        ToJson = DictToJson(obj)
    ElseIf TypeName(obj) = "Collection" Then
        ToJson = CollectionToJson(obj)
    Else
        ToJson = "{}"
    End If
    Exit Function
    
ErrHandler:
    LibDTS_Logger.Log "Error converting to JSON: " & err.description, DTS_ERROR
    ToJson = "{}"
End Function

' ==========================================
' PRIVATE HELPERS
' ==========================================

Private Function ParseJsonObject(jsonStr As String) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Remove outer braces
    jsonStr = mid$(jsonStr, 2, Len(jsonStr) - 2)
    
    ' Simple key-value parsing
    Dim pairs() As String
    pairs = Split(jsonStr, ",")
    
    Dim i As Long
    For i = LBound(pairs) To UBound(pairs)
        Dim pair As String
        pair = Trim$(pairs(i))
        
        Dim colonPos As Long
        colonPos = InStr(pair, ":")
        
        If colonPos > 0 Then
            Dim key As String
            Dim value As String
            
            key = Trim$(mid$(pair, 1, colonPos - 1))
            value = Trim$(mid$(pair, colonPos + 1))
            
            ' Remove quotes from key and value
            key = Replace(key, """", "")
            value = Replace(value, """", "")
            
            dict.Add key, value
        End If
    Next i
    
    Set ParseJsonObject = dict
End Function

Private Function ParseJsonArray(jsonStr As String) As Object
    Dim coll As Object
    Set coll = New Collection
    
    ' Remove outer brackets
    jsonStr = mid$(jsonStr, 2, Len(jsonStr) - 2)
    
    ' Split by comma (simple approach)
    Dim items() As String
    items = Split(jsonStr, ",")
    
    Dim i As Long
    For i = LBound(items) To UBound(items)
        Dim item As String
        item = Trim$(items(i))
        item = Replace(item, """", "")
        coll.Add item
    Next i
    
    Set ParseJsonArray = coll
End Function

Private Function DictToJson(dict As Object) As String
    Dim result As String
    result = "{"
    
    Dim key As Variant
    Dim firstItem As Boolean
    firstItem = True
    
    For Each key In dict.keys
        If Not firstItem Then result = result & ","
        firstItem = False
        
        result = result & """" & CStr(key) & """:"
        result = result & ValueToJson(dict(key))
    Next key
    
    result = result & "}"
    DictToJson = result
End Function

Private Function CollectionToJson(coll As Object) As String
    Dim result As String
    result = "["
    
    Dim item As Variant
    Dim firstItem As Boolean
    firstItem = True
    
    For Each item In coll
        If Not firstItem Then result = result & ","
        firstItem = False
        
        result = result & ValueToJson(item)
    Next item
    
    result = result & "]"
    CollectionToJson = result
End Function

Private Function ValueToJson(value As Variant) As String
    If IsObject(value) Then
        If TypeName(value) = "Dictionary" Then
            ValueToJson = DictToJson(value)
        ElseIf TypeName(value) = "Collection" Then
            ValueToJson = CollectionToJson(value)
        Else
            ValueToJson = "null"
        End If
    ElseIf IsNumeric(value) Then
        ValueToJson = CStr(value)
    ElseIf VarType(value) = vbBoolean Then
        ValueToJson = IIf(value, "true", "false")
    ElseIf IsNull(value) Or IsEmpty(value) Then
        ValueToJson = "null"
    Else
        ' String value - escape quotes
        Dim strValue As String
        strValue = CStr(value)
        strValue = Replace(strValue, """", "\""")
        strValue = Replace(strValue, vbCr, "")
        strValue = Replace(strValue, vbLf, "\n")
        ValueToJson = """" & strValue & """"
    End If
End Function

' ==========================================
' VALIDATION HELPERS
' ==========================================

Public Function IsValidGUID(guid As String) As Boolean
    ' Check if GUID format is valid (basic check)
    IsValidGUID = (Len(guid) >= 32 And Len(guid) <= 40)
End Function

Public Function IsEmptyGUID(guid As String) As Boolean
    IsEmptyGUID = (Len(Trim$(guid)) = 0 Or guid = "{00000000-0000-0000-0000-000000000000}")
End Function
