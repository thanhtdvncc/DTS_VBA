Attribute VB_Name = "LibDTS_Logger"
' Module: LibDTS_Logger
Option Explicit

Private Const LOG_FILE_NAME As String = "DTS_System.log"

' Log type enum
Public Enum DTSLogType
    DTS_INFO = 0
    DTS_WARNING = 1
    DTS_ERROR = 2
End Enum

' Write log to text file in Temp folder
' Ghi log vao file text trong thu muc Temp
Public Sub Log(msg As String, Optional lType As DTSLogType = DTS_INFO)
    Dim fso As Object, ts As Object
    Dim path As String
    Dim typeStr As String
    
    On Error Resume Next
    
    Select Case lType
        Case DTS_INFO: typeStr = "[INFO]"
        Case DTS_WARNING: typeStr = "[WARN]"
        Case DTS_ERROR: typeStr = "[ERROR]"
    End Select
    
    path = Environ("TEMP") & "\" & LOG_FILE_NAME
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Open for appending (8)
    Set ts = fso.OpenTextFile(path, 8, True)
    ts.WriteLine Format(Now, "yyyy-mm-dd HH:nn:ss") & " " & typeStr & " " & msg
    ts.Close
    
    ' If Error, also print to Immediate window
    If lType = DTS_ERROR Then Debug.Print "DTS ERROR: " & msg
End Sub
