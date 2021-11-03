Set oWMP = CreateObject("WMPlayer.OCX.7" )
Set colCDROMs = oWMP.cdromCollection

if colCDROMs.Count >= 1 then
        For i = 0 to colCDROMs.Count - 1
                colCDROMs.Item(i).Eject
        Next ' cdrom
End If

 Function sayHello()
   msgbox("Here is your cupholder buddy!")
 End Function

 Call sayHello()