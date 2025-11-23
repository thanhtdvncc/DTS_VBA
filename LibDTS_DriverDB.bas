Attribute VB_Name = "LibDTS_DriverDB"
' Module: LibDTS_DriverDB
Option Explicit

' Load Settings from JSON file
' Tai cai dat tu file JSON
Public Function LoadSettings() As Object
    Dim fso As Object, ts As Object
    Dim jsonStr As String
    Dim path As String
    
    path = Environ("APPDATA") & "\DTS_Core\settings.json"
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FileExists(path) Then
        Set ts = fso.OpenTextFile(path, 1)
        jsonStr = ts.ReadAll
        ts.Close
        Set LoadSettings = LibDTS_Base.ParseJson(jsonStr)
    Else
        ' Return empty dictionary if no file
        ' Tra ve dictionary rong neu khong co file
        Set LoadSettings = CreateObject("Scripting.Dictionary")
    End If
End Function

' Save Settings to JSON file
' Luu cai dat vao file JSON
Public Sub SaveSettings(settingsDict As Object)
    Dim fso As Object, ts As Object
    Dim jsonStr As String
    Dim path As String
    
    path = Environ("APPDATA") & "\DTS_Core\settings.json"
    jsonStr = LibDTS_Base.ToJson(settingsDict)
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Create folder if not exists
    ' Tao thu muc neu chua co
    If Not fso.FolderExists(Environ("APPDATA") & "\DTS_Core") Then
        fso.CreateFolder Environ("APPDATA") & "\DTS_Core"
    End If
    
    Set ts = fso.CreateTextFile(path, True)
    ts.Write jsonStr
    ts.Close
End Sub
