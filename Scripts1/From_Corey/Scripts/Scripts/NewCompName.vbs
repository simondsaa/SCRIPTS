 Dim Wmi :Set Wmi = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2") 
 Dim Arg, Col, Obj 
 Dim LetterType
 
  For Each Col In Wmi.ExecQuery("Select * from Win32_SystemEnclosure") 
   For Each Obj In Col.ChassisTypes 
     Select Case Obj 
      Case 1  :Arg = "Other" 
      Case 2  :Arg = "Unknown" 
      Case 3  :Arg = "Desktop" 
      Case 4  :Arg = "Low Profile Desktop" 
      Case 5  :Arg = "Pizza Box" 
      Case 6  :Arg = "Mini Tower" 
      Case 7  :Arg = "Tower" 
      Case 8  :Arg = "Portable" 
      Case 9  :Arg = "Laptop" 
      Case 10 :Arg = "Notebook" 
      Case 11 :Arg = "Handheld" 
      Case 12 :Arg = "Docking Station" 
      Case 13 :Arg = "All-in-One" 
      Case 14 :Arg = "Sub-Notebook" 
      Case 15 :Arg = "Space Saving" 
      Case 16 :Arg = "Lunch Box" 
      Case 17 :Arg = "Main System Chassis" 
      Case 18 :Arg = "Expansion Chassis" 
      Case 19 :Arg = "Sub-Chassis" 
      Case 20 :Arg = "Bus Expansion Chassis" 
      Case 21 :Arg = "Peripheral Chassis" 
      Case 22 :Arg = "Storage Chassis" 
      Case 23 :Arg = "Rack Mount Chassis" 
      Case 24 :Arg = "Sealed-Case PC" 
      Case Else  
       Arg = "Unknown" 
    End Select 
   Next 
  Next 
   If Arg = "Notebook" Then
LetterType = "L"
 ElseIf Arg = "Laptop" Then
LetterType = "L"
 ElseIf Arg = "Portable" Then
LetterType = "L"
 ElseIf Arg = "Desktop" Then
LetterType = "W"
 ElseIf Arg = "Mini Tower" Then
LetterType = "W"
 ElseIf Arg = "Low Profile Desktop" Then
LetterType = "W"
 Else
LetterType = "_"
End If

Set fso = CreateObject("Scripting.FileSystemObject")
'OpenTextFile Parameters:
'-Filename
'-The 2 is for writing... 1 is reading and 8 is appending
'-The "True" is to create if not already there.
Set wshNetwork = WScript.CreateObject( "WScript.Network" )
strComputerName = wshNetwork.ComputerName

Set fl = fso.OpenTextFile("\\tyncesaapspd02\afcesashared$\admin\net_admin\DumpLogs\Computers\ByComputername\" & strComputerName & "\tempCompName.txt", 2, True) 
fl.Write(LetterType)
fl.Close : Set fl = Nothing
Set fso = Nothing