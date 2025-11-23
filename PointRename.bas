Attribute VB_Name = "PointRename"
Sub PointRename()
    SwitchOff True
    Call SAP2000_Connectv16
    
    Set ws = Sheets("RenameNodes")
    i = 2
    ret = SapModel.SetPresentUnits(12) 'Ton_m_C
    Do Until ws.Range("A" & i).Value2 = ""
        ret = SapModel.pointObj.ChangeName(ws.Range("A" & i).Value2, ws.Range("B" & i).Value2)
        i = i + 1
    Loop

    Call SAP2000_Disconnect
    SwitchOff False

End Sub

