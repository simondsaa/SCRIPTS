On Error Resume Next

Dim strTarget, strPingResults

'Set Constants
Const ForReading = 1

'Dim Objects
Dim objFSO

'Dim Strings
Dim i
Dim strComputer

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set InputFile = objFSO.OpenTextFile("MachineList.Txt")

intcol = 0
introw = 1

'Create Excel Document
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
objExcel.Workbooks.Add

objExcel.Cells(IntRow, IntCol+1).Value = "Machine Name"
objExcel.Cells(IntRow, IntCol+2).Value = "Machine Type"
objExcel.Cells(IntRow, IntCol+3).Value = "Serial Number"
objExcel.Cells(IntRow, IntCol+4).Value = "Logged On"
objExcel.Cells(introw, intcol+5).Value = "Monitor 1 Brand"
objExcel.Cells(introw, intcol+6).Value = "Monitor 1 Model"
objExcel.Cells(introw, intcol+7).Value = "Monitor 1 Serial Number"
objExcel.Cells(introw, intcol+8).Value = "Monitor 2 Brand"
objExcel.Cells(introw, intcol+9).Value = "Monitor 2 Model"
objExcel.Cells(introw, intcol+10).Value = "Monitor 2 Serial Number"
objExcel.Cells(introw, intcol+11).Value = "Monitor 3 Brand"
objExcel.Cells(introw, intcol+12).Value = "Monitor 3 Model"
objExcel.Cells(introw, intcol+13).Value = "Monitor 3 Serial Number"

Do While Not (InputFile.atEndOfStream)

   strComputer = InputFile.ReadLine

   intCol = 1
   intRow = Introw + 1
   objExcel.Cells(intRow, IntCol).Value = strComputer

' --------- Check If PC is on -------------
Set WshShell = WScript.CreateObject("WScript.Shell")
Set WshExec = WshShell.Exec("ping -n 1 -w 1000 " & strComputer) 'send 3 echo requests, waiting 2secs each
strPingResults = LCase(WshExec.StdOut.ReadAll)
If InStr(strPingResults, "reply from") Then
         'Successful ping


'Set objWMIService = GetObject("winmgmts:" & strComputer & "\root\cimv2")
Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

	Call GetSerialNumber() 		'Get Serial Number
	Call GetLoggedonUser()		'Get Logged On User
	Call GetMachineModel()		'Get Machine Model
	Call GetMonitorInfo() 		'Get Monitor Information

      Else
         'Unsuccessful ping
   objExcel.Cells(intRow, 2).Value ="Offline"
strNewContents = strNewContents & strComputer & vbCrLf

      End If

loop


'Call TidyExcelDocument() 'Tidy Excel Document


'Function TidyExcelDocument()

	' --------- Tidy Excel Document -------------
	objExcel.Range("A1:P1").Select
	objExcel.Selection.Interior.ColorIndex = 19
	objExcel.Selection.Font.ColorIndex = 11
	objExcel.Selection.Font.Bold = True
	objExcel.Cells.EntireColumn.AutoFit

'end function
Set objFile = objFSO.OpenTextFile("MachineList.txt", 2)
objFile.Write strNewContents

Function GetSerialNumber()

	' --------- Get Serial Number -------------
	Set colSMBIOS = objWMIService.ExecQuery _
        	    ("Select * from Win32_BIOS")
	   For Each objSMBIOS in colSMBIOS     

       objExcel.Cells(intRow, 3).Value = objSMBIOS.SerialNumber
Next
end function

Function GetLoggedonUser()

' --------- Get Loggedon User -------------

Set colComputer = objWMIService.ExecQuery("Select * from Win32_ComputerSystem") 

For Each objComputer in colComputer
  
     objExcel.Cells(intRow, 4).Value = objComputer.UserName 
Next 


End Function

Function GetMachineModel()

Set objWMIService = GetObject("winmgmts:" _
& "{impersonationLevel=impersonate}!\\" _ 
& strComputer & "\root\cimv2") 
Set colComputer = objWMIService.ExecQuery _
("Select * from Win32_ComputerSystem") 

For Each objComputer in colComputer

     objExcel.Cells(intRow, 2).Value =  objComputer.Model     

Next 


End Function










Const DISPLAY_REGKEY="HKLM\SYSTEM\CurrentControlSet\Enum\DISPLAY\"
'sets the debug outfile (use format like c:\debug.txt)
Const DEBUGFILE="NUL"
'if set to 1 then output debug info to DEBUGFILE (also writes debug to screen if running under cscript.exe)
Const DEBUGMODE=0
'The ForceCscript subroutine forces execution under CSCRIPT.EXE/Prevents execution
'under WSCRIPT.EXE -- useful when debugging
'ForceCscript

DebugOut "Execution Started " & cstr(now)

DebugOut "Execution Completed " & cstr(now)

'This is the main function. It calls everything else
'in the correct order.
Function GetMonitorInfo()
debugout "Getting all display devices"
arrAllDisplays=GetAllDisplayDevicesInReg()
debugout "Filtering display devices to monitors"
arrAllMonitors=GetAllMonitorsFromAllDisplays(arrAllDisplays)
debugout "Filtering monitors to active monitors"
arrActiveMonitors=GetActiveMonitorsFromAllMonitors(arrAllMonitors)
if ubound(arrActiveMonitors)=0 and arrActiveMonitors(0)="{ERROR}" then
debugout "No active monitors found"
strFormattedMonitorInfo="[Monitor_1]" & vbcrlf & "Monitor=Not Found" & vbcrlf & vbcrlf
else
debugout "Found active monitors"
debugout "Retrieving EDID for all active monitors"
arrActiveEDID=GetEDIDFromActiveMonitors(arrActiveMonitors)
debugout "Parsing EDID/Windows data"
arrParsedMonitorInfo=GetParsedMonitorInfo(arrActiveEDID,arrActiveMonitors)
debugout "Formatting parsed data"
strFormattedMonitorInfo=GetFormattedMonitorInfo(arrParsedMonitorInfo)
end if
debugout "Data retrieval completed"
GetMonitorInfo=strFormattedMonitorInfo
end function







'this function formats the parsed array for display
'this is where the final output is generated
'it is the one you will most likely want to
'customize to suit your needs

Function GetFormattedMonitorInfo(arrParsedMonitorInfo)




for tmpctr=0 to ubound(arrParsedMonitorInfo)
tmpResult=split(arrParsedMonitorInfo(tmpctr),"|||")
'tmpOutput=tmpOutput & "[Monitor_" & cstr(tmpctr+1) & "]" & vbcrlf
'tmpOutput=tmpOutput & "EDID_VESAManufacturerID=" & tmpResult(1) & vbcrlf
'tmpOutput=tmpOutput & "EDID_DeviceID=" & tmpResult(3) & vbcrlf
'tmpOutput=tmpOutput & "EDID_ManufactureDate=" & tmpResult(2) & vbcrlf
'tmpOutput=tmpOutput & "EDID_SerialNumber=" & tmpResult(0) & vbcrlf
'tmpOutput=tmpOutput & "EDID_ModelName=" & tmpResult(4) & vbcrlf
'tmpOutput=tmpOutput & "EDID_Version=" & tmpResult(5) & vbcrlf
'tmpOutput=tmpOutput & "Windows_VESAID=" & tmpResult(6) & vbcrlf
'tmpOutput=tmpOutput & "Windows_PNPID=" & tmpResult(7) & vbcrlf & vbcrlf










objExcel.Cells(introw, intcol+4).Value =  tmpResult(1)


objExcel.Cells(introw, intcol+5).Value =  tmpResult(4)


objExcel.Cells(introw, intcol+6).Value =  tmpResult(0)


intcol = intcol + 3



next
GetFormattedMonitorInfo=tmpOutput

'intCol = intCol +1












End Function

'This function returns an array of all subkeys of the 
'regkey defined by DISPLAY_REGKEY
'(typically this should be "HKLM\SYSTEM\CurrentControlSet\Enum\DISPLAY")
Function GetAllDisplayDevicesInReg()
dim arrResult()
redim arrResult(0)
intArrResultIndex=-1
arrtmpkeys=RegEnumKeys(DISPLAY_REGKEY)
if vartype(arrtmpkeys)<>8204 then
arrResult(0)="{ERROR}"
GetAllDisplayDevicesInReg=false
debugout "Display=Can't enum subkeys of display regkey"
else
for tmpctr=0 to ubound(arrtmpkeys)
arrtmpkeys2=RegEnumKeys(DISPLAY_REGKEY & arrtmpkeys(tmpctr))
for tmpctr2 = 0 to ubound(arrtmpkeys2)
intArrResultIndex=intArrResultIndex+1
redim preserve arrResult(intArrResultIndex)
arrResult(intArrResultIndex)=DISPLAY_REGKEY & arrtmpkeys(tmpctr) & "\" & arrtmpkeys2(tmpctr2)
debugout "Display=" & arrResult(intArrResultIndex)
next 
next
end if
GetAllDisplayDevicesInReg=arrResult


























End Function





'This function is passed an array of regkeys as strings
'and returns an array containing only those that have a
'hardware id value appropriate to a monitor.
Function GetAllMonitorsFromAllDisplays(arrRegKeys)
dim arrResult()
redim arrResult(0)
intArrResultIndex=-1
for tmpctr=0 to ubound(arrRegKeys)
if IsDisplayDeviceAMonitor(arrRegKeys(tmpctr)) then
intArrResultIndex=intArrResultIndex+1
redim preserve arrResult(intArrResultIndex)
arrResult(intArrResultIndex)=arrRegKeys(tmpctr)
debugout "Monitor=" & arrResult(intArrResultIndex)
end if
next
if intArrResultIndex=-1 then
arrResult(0)="{ERROR}"
debugout "Monitor=Unable to locate any monitors"
end if
GetAllMonitorsFromAllDisplays=arrResult
End Function

'this function is passed a regsubkey as a string
'and determines if it is a monitor
'returns boolean
Function IsDisplayDeviceAMonitor(strDisplayRegKey)
arrtmpResult=RegGetMultiStringValue(strDisplayRegKey,"HardwareID")
strtmpResult="|||" & join(arrtmpResult,"|||") & "|||"
if instr(lcase(strtmpResult),"|||monitor\")=0 then
debugout "MonitorCheck='" & strDisplayRegKey & "'|||is not a monitor"
IsDisplayDeviceAMonitor=false
else
debugout "MonitorCheck='" & strDisplayRegKey & "'|||is a monitor"
IsDisplayDeviceAMonitor=true
end if
End Function

'This function is passed an array of regkeys as strings
'and returns an array containing only those that have a
'subkey named "Control"...establishing that they are current.
Function GetActiveMonitorsFromAllMonitors(arrRegKeys)
dim arrResult()
redim arrResult(0)
intArrResultIndex=-1
for tmpctr=0 to ubound(arrRegKeys)
if IsMonitorActive(arrRegKeys(tmpctr)) then
intArrResultIndex=intArrResultIndex+1
redim preserve arrResult(intArrResultIndex)
arrResult(intArrResultIndex)=arrRegKeys(tmpctr)
debugout "ActiveMonitor=" & arrResult(intArrResultIndex)
end if
next

if intArrResultIndex=-1 then
arrResult(0)="{ERROR}"
debugout "ActiveMonitor=Unable to locate any active monitors"
end if
GetActiveMonitorsFromAllMonitors=arrResult
End Function

'this function is passed a regsubkey as a string
'and determines if it is an active monitor
'returns boolean
Function IsMonitorActive(strMonitorRegKey)
arrtmpResult=RegEnumKeys(strMonitorRegKey)
strtmpResult="|||" & join(arrtmpResult,"|||") & "|||"
if instr(lcase(strtmpResult),"|||control|||")=0 then
debugout "ActiveMonitorCheck='" & strMonitorRegKey & "'|||is not active"
IsMonitorActive=false
else
debugout "ActiveMonitorCheck='" & strMonitorRegKey & "'|||is active"
IsMonitorActive=true
end if
End Function

'This function is passed an array of regkeys as strings
'and returns an array containing the corresponding contents
'of the EDID value (in string format) for the "Device Parameters" 
'subkey of the specified key
Function GetEDIDFromActiveMonitors(arrRegKeys)
dim arrResult()
redim arrResult(0)
intArrResultIndex=-1
for tmpctr=0 to ubound(arrRegKeys)
strtmpResult=GetEDIDForMonitor(arrRegKeys(tmpctr))
intArrResultIndex=intArrResultIndex+1
redim preserve arrResult(intArrResultIndex)
arrResult(intArrResultIndex)=strtmpResult
debugout "GETEDID=" & arrRegKeys(tmpctr) & "|||EDID,Yes"
next

if intArrResultIndex=-1 then
arrResult(0)="{ERROR}"
debugout "EDID=Unable to retrieve any edid"
end if
GetEDIDFromActiveMonitors=arrResult
End Function

'given the regkey of a specific monitor
'this function returns the EDID info
'in string format
Function GetEDIDForMonitor(strMonitorRegKey)
arrtmpResult=RegGetBinaryValue(strMonitorRegKey & "\Device Parameters","EDID")
if vartype(arrtmpResult) <> 8204 then
debugout "GetEDID=No EDID Found|||" & strMonitorRegKey
GetEDIDForMonitor="{ERROR}"
else
for each bytevalue in arrtmpResult
strtmpResult=strtmpResult & chr(bytevalue)
next
debugout "GetEDID=EDID Found|||" & strMonitorRegKey
debugout "GetEDID_Result=" & GetHexFromString(strtmpResult)
GetEDIDForMonitor=strtmpResult
end if
End Function

'passed a given string this function 
'returns comma seperated hex values 
'for each byte
Function GetHexFromString(strText)
for tmpctr=1 to len(strText)
tmpresult=tmpresult & right( "0" & hex(asc(mid(strText,tmpctr,1))),2) & ","
next
GetHexFromString=left(tmpresult,len(tmpresult)-1)
End Function

'this function should be passed two arrays with the same
'number of elements. array 1 should contain the
'edid information that corresponds to the active monitor
'regkey found in the same element of array 2
'Why not use a 2D array or a dictionary object?.
'I guess I'm just lazy
Function GetParsedMonitorInfo(arrActiveEDID,arrActiveMonitors)
dim arrResult()
for tmpctr=0 to ubound(arrActiveEDID)
strSerial=GetSerialFromEDID(arrActiveEDID(tmpctr))
strMfg=GetMfgFromEDID(arrActiveEDID(tmpctr))
strMfgDate=GetMfgDateFromEDID(arrActiveEDID(tmpctr))
strDev=GetDevFromEDID(arrActiveEDID(tmpctr))
strModel=GetModelFromEDID(arrActiveEDID(tmpctr))
strEDIDVer=GetEDIDVerFromEDID(arrActiveEDID(tmpctr))
strWinVesaID=GetWinVESAIDFromRegKey(arrActiveMonitors(tmpctr))
strWinPNPID=GetWinPNPFromRegKey(arrActiveMonitors(tmpctr))
redim preserve arrResult(tmpctr)
arrResult(tmpctr)=arrResult(tmpctr) & strSerial & "|||"
arrResult(tmpctr)=arrResult(tmpctr) & strMfg & "|||"
arrResult(tmpctr)=arrResult(tmpctr) & strMfgDate & "|||"
arrResult(tmpctr)=arrResult(tmpctr) & strDev & "|||"
arrResult(tmpctr)=arrResult(tmpctr) & strModel & "|||"
arrResult(tmpctr)=arrResult(tmpctr) & strEDIDVer & "|||"
arrResult(tmpctr)=arrResult(tmpctr) & strWinVesaID & "|||"
arrResult(tmpctr)=arrResult(tmpctr) & strWinPNPID
debugout arrResult(tmpctr)
next
GetParsedMonitorInfo=arrResult
End Function

'this is a simple string function to break the VESA monitor ID
'from the registry key
Function GetWinVESAIDFromRegKey(strRegKey)
if strRegKey="{ERROR}" then
GetWinVESAIDFromRegKey="Bad Registry Info"
exit function
end if
strtmpResult=right(strRegKey,len(strRegkey)-len(DISPLAY_REGKEY))
strtmpResult=left(strtmpResult,instr(strtmpResult,"\")-1) 
GetWinVESAIDFromRegKey=strtmpResult
End Function

'this is a simple string function to break windows PNP device id
'from the registry key
Function GetWinPNPFromRegKey(strRegKey)
if strRegKey="{ERROR}" then
GetWinPNPFromRegKey="Bad Registry Info"
exit function
end if 
strtmpResult=right(strRegKey,len(strRegkey)-len(DISPLAY_REGKEY))
strtmpResult=right(strtmpResult,len(strtmpResult)-instr(strtmpResult,"\"))
GetWinPNPFromRegKey=strtmpResult
End Function

'utilizes the GetDescriptorBlockFromEDID function
'to retrieve the serial number block
'from the EDID data
Function GetSerialFromEDID(strEDID)
'a serial number descriptor will start with &H00 00 00 ff
strTag=chr(&H00) & chr(&H00) & chr(&H00) & chr(&Hff)
GetSerialFromEDID=GetDescriptorBlockFromEDID(strEDID,strTag)
End Function

'utilizes the GetDescriptorBlockFromEDID function
'to retrieve the model description block
'from the EDID data
Function GetModelFromEDID(strEDID)
'a model number descriptor will start with &H00 00 00 fc
strTag=chr(&H00) & chr(&H00) & chr(&H00) & chr(&Hfc)
GetModelFromEDID=GetDescriptorBlockFromEDID(strEDID,strTag)
End Function

'This function parses a string containing EDID data
'and returns the information contained in one of the
'4 custom "descriptor blocks" providing the data in the
'block is tagged wit a certain prefix
'if no descriptor is tagged with the specified prefix then
'function returns "Not Present in EDID"
'otherwise it returns the data found in the descriptor
'trimmed of its prefix tag and also trimmed of
'leading NULLs (chr(0)) and trailing linefeeds (chr(10))
Function GetDescriptorBlockFromEDID(strEDID,strTag)
if strEDID="{ERROR}" then
GetDescriptorBlockFromEDID="Bad EDID"
exit function
end if

'*********************************************************************
'There are 4 descriptor blocks in edid at offset locations
'&H36 &H48 &H5a and &H6c each block is 18 bytes long
'the model and serial numbers are stored in the vesa descriptor
'blocks in the edid.
'*********************************************************************
dim arrDescriptorBlock(3)
arrDescriptorBlock(0)=mid(strEDID,&H36+1,18)
arrDescriptorBlock(1)=mid(strEDID,&H48+1,18)
arrDescriptorBlock(2)=mid(strEDID,&H5a+1,18)
arrDescriptorBlock(3)=mid(strEDID,&H6c+1,18)

if instr(arrDescriptorBlock(0),strTag)>0 then
strFoundBlock=arrDescriptorBlock(0)
elseif instr(arrDescriptorBlock(1),strTag)>0 then
strFoundBlock=arrDescriptorBlock(1)
elseif instr(arrDescriptorBlock(2),strTag)>0 then
strFoundBlock=arrDescriptorBlock(2)
elseif instr(arrDescriptorBlock(3),strTag)>0 then
strFoundBlock=arrDescriptorBlock(3)
else
GetDescriptorBlockFromEDID="Not Present in EDID"
exit function
end if

strResult=right(strFoundBlock,14)
'the data in the descriptor block will either fill the 
'block completely or be terminated with a linefeed (&h0a)
if instr(strResult,chr(&H0a))>0 then
strResult=trim(left(strResult,instr(strResult,chr(&H0a))-1))
else
strResult=trim(strResult)
end if

'although it is not part of the edid spec (as far as i can tell) it seems as though the
'information in the descriptor will frequently be preceeded by &H00, this
'compensates for that
if left(strResult,1)=chr(0) then strResult=right(strResult,len(strResult)-1)

GetDescriptorBlockFromEDID=strResult
End Function

'This function parses a string containing EDID data
'and returns the VESA manufacturer ID as a string
'the manufacturer ID is a 3 character identifier
'assigned to device manufacturers by VESA
'I guess that means you're not allowed to make an EDID
'compliant monitor unless you belong to VESA.
Function GetMfgFromEDID(strEDID)
if strEDID="{ERROR}" then
GetMfgFromEDID="Bad EDID"
exit function
end if

'the mfg id is 2 bytes starting at EDID offset &H08
'the id is three characters long. using 5 bits to represent
'each character. the bits are used so that 1=A 2=B etc..
'
'get the data
tmpEDIDMfg=mid(strEDID,&H08+1,2) 
Char1=0 : Char2=0 : Char3=0 
Byte1=asc(left(tmpEDIDMfg,1)) 'get the first half of the string 
Byte2=asc(right(tmpEDIDMfg,1)) 'get the first half of the string
'now shift the bits
'shift the 64 bit to the 16 bit
if (Byte1 and 64) > 0 then Char1=Char1+16 
'shift the 32 bit to the 8 bit
if (Byte1 and 32) > 0 then Char1=Char1+8 
'etc....
if (Byte1 and 16) > 0 then Char1=Char1+4 
if (Byte1 and 8) > 0 then Char1=Char1+2 
if (Byte1 and 4) > 0 then Char1=Char1+1 

'the 2nd character uses the 2 bit and the 1 bit of the 1st byte
if (Byte1 and 2) > 0 then Char2=Char2+16 
if (Byte1 and 1) > 0 then Char2=Char2+8 
'and the 128,64 and 32 bits of the 2nd byte
if (Byte2 and 128) > 0 then Char2=Char2+4 
if (Byte2 and 64) > 0 then Char2=Char2+2 
if (Byte2 and 32) > 0 then Char2=Char2+1 

'the bits for the 3rd character don't need shifting
'we can use them as they are
Char3=Char3+(Byte2 and 16) 
Char3=Char3+(Byte2 and 8) 
Char3=Char3+(Byte2 and 4) 
Char3=Char3+(Byte2 and 2) 
Char3=Char3+(Byte2 and 1) 
tmpmfg=chr(Char1+64) & chr(Char2+64) & chr(Char3+64)
GetMfgFromEDID=tmpmfg
End Function

'This function parses a string containing EDID data
'and returns the manufacture date in mm/yyyy format
Function GetMfgDateFromEDID(strEDID)
if strEDID="{ERROR}" then
GetMfgDateFromEDID="Bad EDID"
exit function
end if

'the week of manufacture is stored at EDID offset &H10
tmpmfgweek=asc(mid(strEDID,&H10+1,1))

'the year of manufacture is stored at EDID offset &H11
'and is the current year -1990
tmpmfgyear=(asc(mid(strEDID,&H11+1,1)))+1990

'store it in month/year format 
tmpmdt=month(dateadd("ww",tmpmfgweek,datevalue("1/1/" & tmpmfgyear))) & "/" & tmpmfgyear
GetMfgDateFromEDID=tmpmdt
End Function

'This function parses a string containing EDID data
'and returns the device ID as a string
Function GetDevFromEDID(strEDID)
if strEDID="{ERROR}" then
GetDevFromEDID="Bad EDID"
exit function
end if
'the device id is 2bytes starting at EDID offset &H0a
'the bytes are in reverse order.
'this code is not text. it is just a 2 byte code assigned
'by the manufacturer. they should be unique to a model
tmpEDIDDev1=hex(asc(mid(strEDID,&H0a+1,1)))
tmpEDIDDev2=hex(asc(mid(strEDID,&H0b+1,1)))
if len(tmpEDIDDev1)=1 then tmpEDIDDev1="0" & tmpEDIDDev1
if len(tmpEDIDDev2)=1 then tmpEDIDDev2="0" & tmpEDIDDev2
tmpdev=tmpEDIDDev2 & tmpEDIDDev1
GetDevFromEDID=tmpdev
End Function

'This function parses a string containing EDID data
'and returns the EDID version number as a string
'I should probably do this first and then not return any other data
'if the edid version exceeds 1.3 since most if this code probably
'won't work right if they change the spec drastically enough (which they probably
'won't do for backward compatability reasons thus negating my need to check and
'making this comment somewhat redundant)
Function GetEDIDVerFromEDID(strEDID)
if strEDID="{ERROR}" then
GetEDIDVerFromEDID="Bad EDID"
exit function
end if

'the version is at EDID offset &H12
tmpEDIDMajorVer=asc(mid(strEDID,&H12+1,1))

'the revision level is at EDID offset &H13
tmpEDIDRev=asc(mid(strEDID,&H13+1,1))

tmpver=chr(48+tmpEDIDMajorVer) & "." & chr(48+tmpEDIDRev)
GetEDIDVerFromEDID=tmpver
End Function

'simple function to provide an
'easier interface to the wmi registry functions
Function RegEnumKeys(RegKey)
hive=SetHive(RegKey)
set objReg=GetWMIRegProvider()
strKeyPath = right(RegKey,len(RegKey)-instr(RegKey,"\"))
objReg.EnumKey Hive, strKeyPath, arrSubKeys
RegEnumKeys=arrSubKeys
End Function

'simple function to provide an
'easier interface to the wmi registry functions
Function RegGetStringValue(RegKey,RegValueName)
hive=SetHive(RegKey)
set objReg=GetWMIRegProvider()
strKeyPath = right(RegKey,len(RegKey)-instr(RegKey,"\"))
tmpreturn=objReg.GetStringValue(Hive, strKeyPath, RegValueName, RegValue)
if tmpreturn=0 then
RegGetStringValue=RegValue
else
RegGetStringValue="~{{<ERROR>}}~"
end if
End Function

'simple function to provide an
'easier interface to the wmi registry functions
Function RegGetMultiStringValue(RegKey,RegValueName)
hive=SetHive(RegKey)
set objReg=GetWMIRegProvider()
strKeyPath = right(RegKey,len(RegKey)-instr(RegKey,"\"))
tmpreturn=objReg.GetMultiStringValue(Hive, strKeyPath, RegValueName, RegValue)
if tmpreturn=0 then
RegGetMultiStringValue=RegValue
else
RegGetMultiStringValue="~{{<ERROR>}}~"
end if
End Function

'simple function to provide an
'easier interface to the wmi registry functions
Function RegGetBinaryValue(RegKey,RegValueName)
hive=SetHive(RegKey)
set objReg=GetWMIRegProvider()
strKeyPath = right(RegKey,len(RegKey)-instr(RegKey,"\"))
tmpreturn=objReg.GetBinaryValue(Hive, strKeyPath, RegValueName, RegValue)
if tmpreturn=0 then
RegGetBinaryValue=RegValue
else
RegGetBinaryValue="~{{<ERROR>}}~"
end if
End Function

'simple function to provide a wmi registry provider
'to all the other registry functions (regenumkeys, reggetstringvalue, etc...)
Function GetWMIRegProvider()
'strComputer = "."

Set GetWMIRegProvider=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
End Function

'function to parse the specified hive
'from the registry functions above
'to all the other registry functions (regenumkeys, reggetstringvalue, etc...)
Function SetHive(RegKey)
HKEY_CLASSES_ROOT=&H80000000
HKEY_CURRENT_USER=&H80000001
HKEY_CURRENT_CONFIG=&H80000005
HKEY_LOCAL_MACHINE=&H80000002
HKEY_USERS=&H80000003
strHive=left(RegKey,instr(RegKey,"\"))
if strHive="HKCR\" or strHive="HKR\" then SetHive=HKEY_CLASSES_ROOT
if strHive="HKCU\" then SetHive=HKEY_CURRENT_USER
if strHive="HKCC\" then SetHive=HKEY_CURRENT_CONFIG
if strHive="HKLM\" then SetHive=HKEY_LOCAL_MACHINE
if strHive="HKU\" then SetHive=HKEY_USERS
End Function

'this sub forces execution under cscript
'it can be useful for debugging if your machine's
'default script engine is set to wscript
Sub ForceCScript
strCurrScriptHost=lcase(right(wscript.fullname,len(wscript.fullname)-len(wscript.path)-1))
if strCurrScriptHost<>"cscript.exe" then
set objFSO=CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")
Set objArgs = WScript.Arguments
strExecCmdLine=wscript.path & "\cscript.exe //nologo " & objfso.getfile(wscript.scriptfullname).shortpath
For argctr = 0 to objArgs.Count - 1
strExecArg=objArgs(argctr)
if instr(strExecArg," ")>0 then strExecArg=chr(34) & strExecArg & chr(34)
strExecAllArgs=strExecAllArgs & " " & strExecArg
Next
objShell.run strExecCmdLine & strExecAllArgs,1,false
set objFSO = nothing
Set objShell = nothing
Set objArgs = nothing
wscript.quit
end if
End Sub

'allows for a pause at the end of execution
'currently used only for debugging
Sub Pause
set objStdin=wscript.stdin
set objStdout=wscript.stdout
objStdout.write "Press ENTER to continue..."
strtmp=objStdin.readline
end Sub

'if debugmode=1 the writes dubug info to the specified
'file and if running under cscript also writes it to screen.
Sub DebugOut(strDebugInfo)
if DEBUGMODE=0 then exit sub
strCurrScriptHost=lcase(right(wscript.fullname,len(wscript.fullname)-len(wscript.path)-1))
if strCurrScriptHost="cscript.exe" then wscript.echo "Debug: " & strDebugInfo
AppendFileMode=8
set objDebugFSO=CreateObject("Scripting.FileSystemObject")
set objDebugStream=objDebugFSO.OpenTextFile(DEBUGFILE,AppendFileMode,True,False)
objDebugStream.writeline strDebugInfo
objDebugStream.Close
set objDebugStream=Nothing
set objDebugFSO=Nothing
End Sub 

