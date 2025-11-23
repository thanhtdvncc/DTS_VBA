Attribute VB_Name = "SelectLoadCase"
Sub Select_LC3xxx()
    Call SAP2000_Connectv16
    '''
    Dim NumberNames As Long
    Dim MyName() As String
    '''
    ret = SapModel.RespCombo.GetNameList(NumberNames, MyName)
    ret = SapModel.results.Setup.DeselectAllCasesAndCombosForOutput
    ret = SapModel.SetPresentUnits(12) 'Ton_m_C
    For i = 0 To NumberNames - 1
        If MyName(i) Like "LC3*" And Not MyName(i) Like "LC3*A" Then
            ret = SapModel.results.Setup.SetComboSelectedForOutput(MyName(i))
        End If
    Next i
    Call SAP2000_Disconnect
End Sub

Sub CreateNewSession()

Dim xlApp As Excel.Application
Set xlApp = New Excel.Application

xlApp.Workbooks.Add
xlApp.Visible = True

Set xlApp = Nothing

End Sub
