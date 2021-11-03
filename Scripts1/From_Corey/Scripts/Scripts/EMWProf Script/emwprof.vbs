On Error Resume Next

'=======================================================================================
'  Initailize variables and constants
'=======================================================================================


Dim strUserName, objADSI, objUser, strTargetAddress, objReg, strProfilePath, wshOKOnly
Dim strDefaultProfile, strMailServer, wshExclamation, strWarning, objWSH, intReturn, strComplete
Dim strCacheMode, strCacheModeFormat, strSwitched
Const HKEY_CURRENT_USER = &H80000001

'=======================================================================================
' Logging has been added for Troubleshooting
'=======================================================================================

Set fso = CreateObject("Scripting.FileSystemObject")
Set oShell = CreateObject("Wscript.Shell") 
strUserProfile = oShell.ExpandEnvironmentStrings("%USERPROFILE%") 
Set fl1 = fso.createTextFile(strUserProfile & "\EMWProf_VBS_Time.log")

fl1.Writeline Time & " VBS Step1: EMWProf.vbs is starting"

'=======================================================================================
' Query targetAddress attribute to see if Mailbox has been Switched
'=======================================================================================

Set objADSI    		=	CreateObject("ADSystemInfo")
strUserName    		=	objADSI.username
strUserName    		=	replace(strUserName, "/", "\/")
Set objUser    		=	GetObject("LDAP://" & strUserName)

If len(objUser.targetAddress) > 0 Then
	strTargetAddress = objUser.targetAddress
	'Do Nothing: Mailbox has not been switched!
	strSwitched = "Not Switched"
	fl1.Writeline Time & " VBS Step2a: User has a TargetAddress of " & strTargetAddress
Else

fl1.Writeline Time & " VBS Step2b: Users TargetAddress is null"
strSwitched = "Switched"

'=======================================================================================
' Mailbox has been switched: Find Default Outlook profile for current user
'=======================================================================================

	Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
	strProfilePath = "Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles\"
	objReg.GetStringValue HKEY_CURRENT_USER, strProfilePath, "DefaultProfile", strDefaultProfile
	objReg.EnumKey HKEY_CURRENT_USER, (strProfilePath & strDefaultProfile & "\"),arrSubKeys
	For Each subkey IN arrSUbKeys
		objReg.Enumvalues HKEY_CURRENT_USER, (strProfilePath & strDefaultProfile & "\" & subkey & "\"),arrSubValues
		on error resume next
		For EACh valueKey in arrSubValues
			if vAlueKey = "001e6602" then
				objREg.GetStringvalue HKEY_CURRENT_USER, (strProfilePath & strDefaultProfile & "\" & subkey), "001e6602", strMailServer
			end if
		Next
		On Error goto 0
	Next
	
fl1.Writeline Time & " VBS Step3: Checked Outlook Profile: Users mail server = " & strMailServer

'=======================================================================================
' See if Default Outlook profile has been updated, if not run EMWProf
'=======================================================================================

	If Instr(lcase(strMailServer),"us.af.mil") Then
		'Do Nothing: Profile has already been updated!
		fl1.Writeline Time & " VBS Step4a: Mail server name contains us.af.mil"

	Else

		ProcessesToKill = Array("communicator.exe","desktopmgr.exe","outlook.exe")
		fl1.Writeline Time & ": [START] Terminating MAPI processes"
		For each process in ProcessesToKill
			Set objWMIService = GetObject("winmgmts:" _
				& "{impersonationLevel=impersonate}!\\.\root\cimv2")
			Set colProcess = objWMIService.ExecQuery _
				("Select * from Win32_Process Where Name = '" & process & "'")
			If colProcess.Count = 0 Then
				' process was not running, log it and move on
				fl1.Writeline Time & ": MAPI processes '" & process & "' not running."
			Else
				For Each objProcess in colProcess
					fl1.Writeline Time & ": Terminating MAPI process '" & process & "'"
					objProcess.Terminate()
					fl1.Writeline Time & ": Terminated"
				Next
			End If
		Next

		fl1.Writeline Time & ": [STOP] Terminating MAPI processes"		

		wshOKOnly = 0
		wshExclamation = 48
		strWarning = "Your email account has been migrated, and EMW Profile Update Tool is now running. " & vbcrlf _
			& "Please DO NOT open Outlook or Blackberry software until EMW Client Updater is finished running."
		Set objWSH = CreateObject("WScript.Shell")
        	intReturn = objWSH.Popup(strWarning, 12, "E-Mail Migration In Progress...", wshOKOnly + wshExclamation)
		
		fl1.Writeline Time & " VBS Step4b: Mail server name does not contain us.af.mil.  Time to Start Emwprof.bat"
		
		objWSH.Run "Emwprof.bat", 7, True

		fl1.Writeline Time & " VBS Step4c: Emwprof.bat has completed"
		
		
		strComplete = "Outlook Profile Updated Successfully!" & vbcrlf & "It is now safe to open Outlook"
		intReturn = objWSH.Popup(strComplete, 12, "E-Mail Migration Complete ...", wshOKOnly + wshExclamation)
		'objWSH.Run chr(34) & "C:\Program Files\Microsoft Office Communicator\communicator.exe" & chr(34)

'=======================================================================================
' Enable cache mode in Outlook after emwprof runs
'=======================================================================================

		Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
		strProfilePath = "Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles\"
		objReg.GetStringValue HKEY_CURRENT_USER, strProfilePath, "DefaultProfile", strDefaultProfile
		objReg.EnumKey HKEY_CURRENT_USER, (strProfilePath & strDefaultProfile & "\"),arrSubKeys
		For Each subkey IN arrSUbKeys
			objReg.Enumvalues HKEY_CURRENT_USER, (strProfilePath & strDefaultProfile & "\" & subkey & "\"),arrSubValues
			on error resume next
			For Each valueKey in arrSubValues
				if valueKey = "00036601" then
					objReg.GetBinaryValue HKEY_CURRENT_USER, (strProfilePath & strDefaultProfile & "\" & subkey), "00036601", arrCacheMode

					if arrCacheMode(0) = &H04 Then
						arrCacheMode(0) = arrCacheMode(0) Or &H80
						arrCacheMode(1) = arrCacheMode(1) Or &H09
						objReg.SetBinaryValue HKEY_CURRENT_USER, (strProfilePath & strDefaultProfile & "\" & subkey), "00036601", arrCacheMode
						fl1.Writeline Time & " CACHE MODE WAS DISABLED, SUCCESSFULLY ENABLED"
					else
						fl1.Writeline Time & " CACHE MODE ALREADY ENABLED"
					end if
	
				end if
			Next
			On Error goto 0
		Next
	End If
End If

fl1.Writeline Time & " VBS Step5: Emwprof.vbs has completed"

'=======================================================================================
'  Clean up
'=======================================================================================

Set objADSI    	= Nothing
Set objUser    	= Nothing
Set objReg 	= Nothing
Set objWSH	= Nothing
