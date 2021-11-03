CONST HKEY_LOCAL_MACHINE = &H80000002
CONST msiInstallStateAbsent=2
Dim strValue, strValue1, strValue3, strValue4, oReg, WshShell, strLogFile, strResults
Dim msi, re, strList

On Error Resume Next
Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\root\default:StdRegProv")
Set WshShell = CreateObject("Wscript.Shell")
Set fso = CreateObject("scripting.filesystemobject")
strLogFile = WshShell.ExpandEnvironmentStrings("%windir%") & "\SMSWORK\AppUninstall(Flash).log"
Set LogFile = fso.OpenTextFile(strLogFile, 2, True)

log "                 Beginning AppUninstall....                   "

strKeyPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
oReg.EnumKey HKEY_LOCAL_MACHINE, strKeyPath, arrSubKeys

log "Enumerating all subkeys under " & strKeyPath
For Each subkey In arrSubKeys
    oReg.GetStringValue HKEY_LOCAL_MACHINE, strKeyPath & "\" & subkey, "DisplayName", strValue

	If InStr(strValue, "Flash") > 0 or InStr(strValue, "Macromedia Flash") > 0 then
		log "Found a probably instance of a Flash install."
		log "Display name = " & strValue
		'WScript.Echo strValue
		oReg.GetStringValue HKEY_LOCAL_MACHINE, strKeyPath & "\" & subkey, "UninstallString", strValue1
		log "UninstallString value = " & strValue1
		
		If strValue1 <> vbNullString Then
		    strValue1 = Replace(strValue1, " /I", " /x")
		    strValue1 = Replace(strValue1, " /qn", " /x")
		    strValue = Replace(strValue, " ", "")
		    strValue = Replace(strValue, "(TM)", "")
		    strValue1 = Replace(strValue1, "xec.exe ", "xec.exe /qn ") &  " /L* " & Chr(34) & "C:\Windows\SMSWORK\" & strValue & "Uninstall.log" & Chr(34)
		    log "Modifying uninstall string with appropriate syntax."
			log "New uninstall string = " & strValue1
			' WScript.Echo strValue1
			log "Launching command " & Chr(34) & strValue1 & Chr(34)
		    WshShell.Run strValue1, 0, 1
			If err <> 0 then
				log "Error returned when running previous command."
				log "Error description = " & err.description
				err.clear
			else
				log "Command completed successfully."
			End if
		Else
			log "UninstallString value was Null. Proceeding to the next subkey."
		End if
	End if
	err.clear
Next

Set subkey = Nothing
strKeyPath1 = "SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
oReg.EnumKey HKEY_LOCAL_MACHINE, strKeyPath1, arrSubKeys1

log vbCrLf
log "-----------------------------------------------------------------"
log vbCrLf
log "Enumerating all subkeys under " & strKeyPath1

For Each subkey In arrSubKeys1
    oReg.GetStringValue HKEY_LOCAL_MACHINE, strKeyPath1 & "\" & subkey, "DisplayName", strValue3

	If InStr(strValue3, "Flash") > 0 or InStr(strValue3, "Macromedia Flash") > 0 then
		log "Found a probably instance of a Flash install."
		log "Display name = " & strValue3
		'WScript.Echo strValue
		oReg.GetStringValue HKEY_LOCAL_MACHINE, strKeyPath1 & "\" & subkey, "UninstallString", strValue4
		log "UninstallString value = " & strValue4
		
		If strValue4 <> vbNullString Then
		    strValue4 = Replace(strValue4, " /I", " /x")
		    strValue4 = Replace(strValue4, " /qn", " /x")
		    strValue3 = Replace(strValue3, " ", "")
		    strValue3 = Replace(strValue3, "(TM)", "")
		    strValue4 = Replace(strValue4, "xec.exe ", "xec.exe /qn ") &  " /L* " & Chr(34) & "C:\Windows\SMSWORK\" & strValue3 & "Uninstall.log" & Chr(34)
		    log "Modifying uninstall string with appropriate syntax."
			log "New uninstall string = " & strValue4
			' WScript.Echo strValue1
			log "Launching command " & Chr(34) & strValue4 & Chr(34)
		    WshShell.Run strValue4, 0, 1
			If err <> 0 then
				log "Error returned when running previous command."
				log "Error description = " & err.description
				err.clear
			else
				log "Command completed successfully."
			End if
		Else
			log "UninstallString value was Null. Proceeding to the next subkey."
		End if
	End if
	err.clear
Next

log vbCrLf
log "-----------------------------------------------------------------"
log vbCrLf

log "Enumerating MSI installs..."

Set msi = CreateObject("WindowsInstaller.Installer")
Set re = New RegExp
msi.UILevel = 2 
 
For Each msipackage In msi.Products
	info = msipackage & " = " & msi.ProductInfo(msipackage, "ProductName")
	If InStr(Info, "Flash") > 0 AND InStr(Info, "Macromedia Flash") = 0 AND InStr(Info, "Development Kit") = 0 then
	    strList = msipackage  & "|" & strList
	    log "Found a probably instance of a Flash install."
		log Info
	End If
Next

strList = Replace (strList, "{", "")
strList = Replace (strList, "}", "")
If Right (strList, 1) = "|" then
    strList = Mid(strList, 1, (Len(strList) - 1))
End if

re.pattern="{(" & strList & ")}"

For Each msipackage In msi.Products
	if re.test(msipackage) Then
		log "Uninstalling " & msi.ProductInfo(msipackage, "ProductName")
		msi.ConfigureProduct msipackage, 0, msiInstallStateAbsent
		if err <> 0 then
			log "Error uninstalling " & msi.ProductInfo(msipackage, "ProductName") & "."
			log err.description
			err.clear
		else
			log msi.ProductInfo(msipackage, "ProductName") & " uninstalled successfully."
		End if
	End If
	err.clear
Next


log "Script completed successfully."

Set oReg = Nothing
strKeyPath = vbNullString
strValueName = vbNullString
strValue = vbNullString

'WScript.Echo "Done"
WScript.Quit

Sub log(strResults)
	LogFile.WriteLine(Now & " >>>> " & strResults) 
End Sub