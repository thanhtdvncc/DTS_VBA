VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDTSEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class Module: clsDTSEvents
Option Explicit

' Define core events
' Dinh nghia cac su kien cot loi
Public Event ElementCreated(elem As clsDTSElement)
Public Event ElementModified(elem As clsDTSElement)
Public Event ElementDeleted(guid As String)
Public Event SelectionChanged(count As Long)

' Method to raise creation event
' Phuong thuc kich hoat su kien tao moi
Public Sub NotifyElementCreated(elem As clsDTSElement)
    RaiseEvent ElementCreated(elem)
End Sub

' Method to raise modification event
' Phuong thuc kich hoat su kien chinh sua
Public Sub NotifyElementModified(elem As clsDTSElement)
    RaiseEvent ElementModified(elem)
End Sub
