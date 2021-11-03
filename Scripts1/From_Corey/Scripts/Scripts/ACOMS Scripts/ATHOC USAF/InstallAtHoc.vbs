'****************************************************************************************
'*
'*   InstallAtHoc.vbs
'*   Version 3.0 08/12/2013
'*   Created by Derek Morehead
'*   Modified by SSgt Colantuono
'*   AtHoc v. 6.2.27.268
'*
'*   Purpose: Enumerates Domain name from WMI and installs intended version of AtHoc
'*	      client using BaseURL and PID associated with client domain/base.
'*
'*
'*
'****************************************************************************************


Const ForReading = 1
Const ForWriting = 2

On Error Resume Next

Set WShShell = CreateObject("WScript.Shell")
Set objWMI = GetObject("winmgmts:\root\cimv2")
Set colItems = objWMI.ExecQuery("Select * from Win32_NTDomain", , 48)

For Each objItem in colItems
    strDomain = objItem.ClientSiteName
If strDomain <> "" Then Exit For
	Next

Set objWMI = Nothing
Set colItems = Nothing
Set objItem = Nothing

strWinDir = WshShell.ExpandEnvironmentStrings("%windir%")

Set FSO = CreateObject("Scripting.FileSystemObject")
strFolder = FSO.GetFile(WScript.ScriptFullname).ParentFolder

If NOT fso.FolderExists(strWinDir & "\SMSWork") then
	fso.CreateFolder(strWinDir & "\SMSWork")
End if

strLogFile = strWinDir & "\SMSWork\AtHocUSAFDSW6_2_27_268_VB.log"

Set LogFile = fso.CreateTextFile(strLogFile, 2, True)

Log "Beginning AtHocInstall.vbs."
Log "Looking for the client site name of the system."
Log "Site name found is " & strDomain & "."

Select Case LCase(strDomain)
'*******************   ACC SITES   *****************
	Case "muhj-langley-esul","muhj-langley","langley-va","langley-rhs","nosc","nosccoop"
		strURL = "BASEURL=https://alertswest.acc.af.mil/config/baseurl.asp"
		strPID = "PID=50295402"
	Case "shaw-sc","vlsb-shaw"
		strURL = "BASEURL=https://alertswest.acc.af.mil/config/baseurl.asp"
		strPID = "PID=50258778"
	Case "ghzo-fewarren"
		strURL = "BASEURL=https://alertswest.acc.af.mil/config/baseurl.asp" 
		strPID = "PID=50318565"
	Case "nzas-malmstrom"
		strURL = "BASEURL=https://alertswest.acc.af.mil/config/baseurl.asp" 
		strPID = "PID=50290792"
	Case "sj-nc","vkag-seymourjohnson"
		strURL = "BASEURL=https://alertswest.acc.af.mil/config/baseurl.asp"
		strPID = "PID=50264356"
	Case "fnwz-dyess","dyess-tx"
		strURL = "BASEURL=https://alertswest.acc.af.mil/config/baseurl.asp"
		strPID = "PID=50323782"
	Case "moody-ga","qseu-moody"
		strURL = "BASEURL=https://alertswest.acc.af.mil/config/baseurl.asp"
		strPID = "PID=50276250"
	Case "baey-beale","beale-ca"
		strURL = "BASEURL=https://alertswest.acc.af.mil/config/baseurl.asp"
		strPID = "PID=50010104"
	Case "gyzh-mountainhome","mh-id"
		strURL = "BASEURL=https://alertswest.acc.af.mil/config/baseurl.asp"
		strPID = "PID=50010107"
	Case "nellis-nv","rkmf-nellis"
		strURL = "BASEURL=https://alertswest.acc.af.mil/config/baseurl.asp"
		strPID = "PID=50010119"
	Case "dm-az","fbnv-davismonthan"
		strURL = "BASEURL=https://alertswest.acc.af.mil/config/baseurl.asp"
		strPID = "PID=50010113"
	Case "ellsworth-sd","fxbm-ellsworth"
		strURL = "BASEURL=https://alertswest.acc.af.mil/config/baseurl.asp"
		strPID = "PID=50010116"
	Case "holloman-nm","kwrd-holloman"
		strURL = "BASEURL=https://alertswest.acc.af.mil/config/baseurl.asp"
		strPID = "PID=50010125"
	Case "offutt-ne","sgbp-offutt"
		strURL = "BASEURL=https://alertswest.acc.af.mil/config/baseurl.asp"
		strPID = "PID=50010100"
'*******************  USAFE SITES  *******************
	Case "ashe-aviano"
		strURL =  "BASEURL=https://alerts.aviano.af.mil/config/baseurl.asp"
		strPID = "PID=2010773"
	Case "ljyc-incirlik"
		strURL =  "BASEURL=https://alerts.incirlik.af.mil/config/baseurl.asp"
		strPID = "PID=2010782"
	Case "mqna-lajes"
		strURL =  "BASEURL=https://alerts.lajes.af.mil/config/baseurl.asp"
		strPID = "PID=2050329"
	Case "exss-croughton","gkux-fairford","gfpy-mildenhall","gfpy-mildenhall-gsu-pzkz-menwithhill"
		strURL =  "BASEURL=https://alerts.lakenheath.af.mil/config/baseurl.asp"
		strPID = "PID=2010776"
	Case "msek-lakenheath"
		strURL =  "BASEURL=https://alerts.lakenheath.af.mil/config/baseurl.asp"
		strPID = "PID=2010779"
	Case "tyfq-ramstein","tyfr-ramstein"
		strURL =  "BASEURL=https://alerts.ramstein.af.mil/config/baseurl.asp"
		strPID = "PID=2010764"
	Case "sdhm-spangdahlem"
		strURL =  "BASEURL=https://alerts.spangdahlem.af.mil/config/baseurl.asp"
		strPID = "PID=2010110"
'****************	AFMC SITES  *********************
	Case "anzw-arnold"
		strURL = "BASEURL=https://alertseast.afmc.af.mil/config/baseurl.asp"
		strPID = "PID=2029223"
	Case "ftfa-eglin"
		strURL = "BASEURL=https://alertseast.afmc.af.mil/config/baseurl.asp"
		strPID = "PID=2010110"
	Case "jumj-gunter"
		strURL = "BASEURL=https://alertseast.afmc.af.mil/config/baseurl.asp"
		strPID = "PID=2112396"
	Case "catd-hanscom"
		strURL = "BASEURL=https://alertseast.afmc.af.mil/config/baseurl.asp"
		strPID = "PID=2029226"
	Case "uhhy-robins"
		strURL = "BASEURL=https://alertseast.afmc.af.mil/config/baseurl.asp"
		strPID = "PID=2018077"
	Case "zhtv-wrightpatterson","zhtx-wrightpatterson-apc"
		strURL = "BASEURL=https://alertseast.afmc.af.mil/config/baseurl.asp"
		strPID = "PID=2029220"
	Case "fspm-edwards"
		strURL = "BASEURL=https://alertswest.afmc.af.mil/config/baseurl.asp"
		strPID = "PID=2010110"
	Case "krsm-hill"
		strURL = "BASEURL=https://alertswest.afmc.af.mil/config/baseurl.asp"
		strPID = "PID=2023293"
	Case "mhmv-kirtland"
		strURL = "BASEURL=https://alertswest.afmc.af.mil/config/baseurl.asp"
		strPID = "PID=2023288"
	Case "wwyk-tinker"
		strURL = "BASEURL=https://alertswest.afmc.af.mil/config/baseurl.asp"
		strPID = "PID=2023296"
	Case "uldf-rome"
		strURL = "BASEURL=https://alertseast.afmc.af.mil/config/baseurl.asp"
		strPID = "PID=2112399"
	Case "xlwu-tyndall"
		strURL = "BASEURL=https://alertswest.acc.af.mil/config/baseurl.asp"
		strPID = "PID=50250670"
'****************	AFRC SITES  *********************
	Case "uhhz-hq-afrc"
		strURL = "Baseurl=https://alerts.afrc.af.mil/csi/"
		strPID = "PID=2010576"
End Select

strCmd = "msiexec.exe /qn /i " & Chr(34) & strFolder & "\AtHocUSAFDSW6.2.27.268.msi" & Chr(34) & " /l*vx " & strWinDir & "\SMSWork\AtHocUSAFDSW6.2.27.268.log " & strUrl & " " & strPID & " RUNAFTERINSTALL=N DESKBAR=N TOOLBAR=N SILENT=Y VALIDATECERT=N MANDATESSL=N UNINSTALLOPTION=N"

Log "Using the following string for installation at base: " & strDomain & "."
Log "------------------------------------------------------------------------------------"
Log strCmd
Log "------------------------------------------------------------------------------------"

err.clear
If fso.FileExists(strFolder & "\AtHocUSAFDSW6.2.27.268.msi") then
	If strURL <> vbNullString and strPID <> vbNullString then
		iResult = WshShell.Run(strCmd, 0, True)

		If iResult <> 0 then
			Log "Installation failed with error: " & err.description
			Log "Return Code: " & iResult
		Else
			Log "Installation completed successfully."
			Log "Return Code: " & iResult
		End if
	Else
		Log "URL or PID was null, which means the domain for this system wasn't found."
	End if
Else
	Log "The installation file, " & strFolder & "\AtHocUSAFDSW6.2.27.268.msi could not be found."
	Log "Exiting installation process."
End if

Log "Work finished. Ending AtHocInstall.vbs."
f.close
Set FSO = Nothing
Set oRootDSE = Nothing
Set oMyDomain = Nothing
Set WShShell = Nothing
Set LogFile = Nothing

WScript.Quit(iReturn)


Sub log(strResults)
	LogFile.WriteLine(Now & " >>>> " & strResults) 
End Sub