Dim LocNum: LocNum = Wscript.Arguments.Named("prnr")

Set AFCECPrnt = CreateObject("WScript.Network") 

If LocNum = "1" Then
'Location = CEX
AFCECPrnt.AddWindowsPrinterConnection "\\tyncesaapspd02\CX HP Color 5525"
AFCECPrnt.AddWindowsPrinterConnection "\\tyncesaapspd02\CX Xerox 5330"
AFCECPrnt.SetDefaultPrinter "\\tyncesaapspd02\CX Xerox 5330"
End If

If LocNum = "2" Then
'Location = VAULT
AFCECPrnt.AddWindowsPrinterConnection "\\tyncesaapspd02\CX Vault Bizhub 363" 
AFCECPrnt.SetDefaultPrinter "\\tyncesaapspd02\CX Vault Bizhub 363"
End If

If LocNum = "3" Then
'Location = CC
AFCECPrnt.AddWindowsPrinterConnection "\\tyncesaapspd02\CC HP B&W 4250"
AFCECPrnt.AddWindowsPrinterConnection "\\tyncesaapspd02\CC HP Color 5525"
AFCECPrnt.AddWindowsPrinterConnection "\\tyncesaapspd02\CC Bizhub 363" 
AFCECPrnt.SetDefaultPrinter "\\tyncesaapspd02\CC Bizhub 363"
End If

If LocNum = "4" Then
'Location = CEXR
AFCECPrnt.AddWindowsPrinterConnection "\\tyncesaapspd02\CXR HP Color 5525"
AFCECPrnt.AddWindowsPrinterConnection "\\tyncesaapspd02\CXR Bizhub 363" 
AFCECPrnt.SetDefaultPrinter "\\tyncesaapspd02\CXR Bizhub 363"
End If

If LocNum = "5" Then
'Location = CEB w/o 7500
AFCECPrnt.AddWindowsPrinterConnection "\\tyncesaapspd02\DS Xerox 5740"
AFCECPrnt.AddWindowsPrinterConnection "\\tyncesaapspd02\CXR HP COlor 5525" 
AFCECPrnt.SetDefaultPrinter "\\tyncesaapspd02\DS Xerox 5740"
End If

If LocNum = "6" Then
'Location = CEB w/ 7500
AFCECPrnt.AddWindowsPrinterConnection "\\tyncesaapspd02\DS Xerox 5740"
AFCECPrnt.AddWindowsPrinterConnection "\\tyncesaapspd02\DS Xerox Phaser 7500" 
AFCECPrnt.SetDefaultPrinter "\\tyncesaapspd02\DS Xerox 5740"
End If

If LocNum = "7" Then
'Location = CEO
AFCECPrnt.AddWindowsPrinterConnection "\\tyncesaapspd02\CO Bizhub 363" 
AFCECPrnt.AddWindowsPrinterConnection "\\tyncesaapspd02\CO HP Color 5525" 
AFCECPrnt.SetDefaultPrinter "\\tyncesaapspd02\CO Bizhub 363"
End If

If LocNum = "8" Then
'Location = CEOM
AFCECPrnt.AddWindowsPrinterConnection "\\tyncesaapspd02\COO Xerox 5330" 
AFCECPrnt.AddWindowsPrinterConnection "\\tyncesaapspd02\CO HP Color 5525"
AFCECPrnt.SetDefaultPrinter "\\tyncesaapspd02\COO Xerox 5330"
End If

If LocNum = "9" Then
'Location = Trailer B Front
AFCECPrnt.AddWindowsPrinterConnection "\\tyncesaapspd02\Trailer B HP COlor 5525" 
AFCECPrnt.AddWindowsPrinterConnection "\\tyncesaapspd02\Trailer B Xerox 5740 FRONT"
AFCECPrnt.AddWindowsPrinterConnection "\\tyncesaapspd02\Trailer B Xerox 5740 BACK"
AFCECPrnt.SetDefaultPrinter "\\tyncesaapspd02\Trailer B Xerox 5740 FRONT"
End If

If LocNum = "10" Then
'Location = Trailer B Back
AFCECPrnt.AddWindowsPrinterConnection "\\tyncesaapspd02\Trailer B HP COlor 5525"
AFCECPrnt.AddWindowsPrinterConnection "\\tyncesaapspd02\Trailer B Xerox 5740 FRONT"
AFCECPrnt.AddWindowsPrinterConnection "\\tyncesaapspd02\Trailer B Xerox 5740 BACK"
AFCECPrnt.SetDefaultPrinter "\\tyncesaapspd02\Trailer B Xerox 5740 BACK"
End If

If LocNum = "11" Then
'Location = Trailer A
AFCECPrnt.AddWindowsPrinterConnection "\\tyncesaapspd02\Trailer A Bizhub 423" 
AFCECPrnt.SetDefaultPrinter "\\tyncesaapspd02\Trailer A Bizhub 423"
End If

wscript.quit 