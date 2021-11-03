Dim LocNum: LocNum = Wscript.Arguments.Named("prnr")

Set AFCECPrnt = CreateObject("WScript.Network") 

If LocNum = "8" Then
'Location = CEOM
AFCECPrnt.AddWindowsPrinterConnection "\\tyncesaapspd02\COO Xerox 5330" 
AFCECPrnt.AddWindowsPrinterConnection "\\tyncesaapspd02\CO HP Color 5525"
AFCECPrnt.SetDefaultPrinter "\\tyncesaapspd02\COO Xerox 5330"
End If

wscript.quit 