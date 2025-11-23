Attribute VB_Name = "LibDTS_Bootstrap"
' Module: LibDTS_Bootstrap
Option Explicit

' Main Entry Point to Initialize System
' Diem vao chinh de khoi tao he thong
Public Sub InitializeSystem()
    On Error GoTo ErrHandler
    
    ' 1. Check System Folders
    ' 1. Kiem tra thu muc he thong
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim appPath As String
    appPath = Environ("APPDATA") & "\DTS_Core"
    
    If Not fso.FolderExists(appPath) Then
        fso.CreateFolder appPath
    End If
    
    ' 2. Initialize Config (Force Load)
    ' 2. Khoi tao cau hinh (Bat buoc tai)
    Dim cfg As clsDTSConfig
    Set cfg = LibDTS_Global.Config ' This triggers auto-load
    
    ' 3. Check License (Optional stub)
    ' 3. Kiem tra ban quyen (Ham cho)
    ' If Not CheckLicense() Then Err.Raise 999, "DTS", "License Invalid"
    
    Debug.Print "[DTS] System Initialized Successfully."
    Exit Sub
    
ErrHandler:
    MsgBox "Critical Error Initializing DTS System: " & err.description, vbCritical
End Sub

' Called when AutoCAD/Excel closes
' Goi khi AutoCAD/Excel dong
Public Sub TerminateSystem()
    ' Save config if dirty
    ' Luu cau hinh neu co thay doi
    If Not LibDTS_Global.Config Is Nothing Then
        LibDTS_Global.Config.Save
    End If
    
    LibDTS_Global.ResetGlobals
End Sub
