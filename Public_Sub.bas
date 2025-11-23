Attribute VB_Name = "Public_Sub"

Public Sub Remove_Project_Reference(RefName As String)
    Dim ref As Variant
    With ThisWorkbook.VBProject
        For Each ref In .References
            If ref.Name = RefName Then .References.Remove ref
        Next
    End With
End Sub

Public Sub AddReferenceSAP2000()                 'Khong can thiet phai add reference voi late binding
    On Error Resume Next
    With ThisWorkbook.VBProject
        .References.AddFromGuid "{0002E157-0000-0000-C000-000000000046}", 2, 0
        .References.AddFromFile "D:\Program Files\Computers and Structures\SAP2000 19\SAP2000v19.tlb"
        .References.AddFromFile "C:\Program Files\Computers and Structures\SAP2000 19\SAP2000v19.tlb"
        .References.AddFromFile "D:\Program Files\Computers and Structures\SAP2000 22\SAP2000v1.tlb"
        .References.AddFromFile "D:\Program Files\Computers and Structures\SAP2000 22\SAP2000.tlb"
        .References.AddFromFile "C:\Program Files\Computers and Structures\SAP2000 22\SAP2000v1.tlb"
        .References.AddFromFile "C:\Program Files\Computers and Structures\SAP2000 22\SAP2000.tlb"
        .References.AddFromFile "D:\Program Files (x86)\Computers and Structures\SAP2000 16\SAP2000.exe"
        .References.AddFromFile "C:\Program Files (x86)\Computers and Structures\SAP2000 16\SAP2000.exe"
    End With
    err.Clear
    On Error GoTo 0
End Sub

Public Sub RemoveReferenceSAP2000()
    Remove_Project_Reference "SAP2000v16"
    Remove_Project_Reference "SAP2000v19"
    Remove_Project_Reference "SAP2000"
    Remove_Project_Reference "SAP2000v1"
End Sub

Public Sub RemoveMissingReference()
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Dim oRefS As Object, oRef As Object
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

Set oRefS = ActiveWorkbook.VBProject.References
For Each oRef In oRefS
    If oRef.IsBroken Then
        Call oRefS.Remove(oRef)
    End If
Next

End Sub

