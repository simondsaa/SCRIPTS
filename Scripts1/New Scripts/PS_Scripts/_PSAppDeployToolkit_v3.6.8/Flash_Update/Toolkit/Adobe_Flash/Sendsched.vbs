'********************************************************************
'
'   SendSched.vbs
'
'   Syntax:
'       cscript.exe SendSched.vbs ScheduleID [Machine]
'       wscript.exe SendSched.vbs ScheduleID [Machine]
'
'       ScheduleID is the ID of the schedule
'       Machine is an optional parameter specifying the name of the machine on which to run
'********************************************************************

dim sMachine
dim sScheduleID
dim oCCMNamespace
dim oInstance
dim bCscript
dim oParams

'********************************************************************
'
'   OUT function
'
'********************************************************************

sub OUT(sText)
    if (bCscript = true) then
        WScript.Echo(sText)
    end if
end sub

'********************************************************************
'
'   OutputUsage
'
'********************************************************************

sub OutputUsage
    WScript.Echo("Usage: cscript SendSched.vbs ScheduleID [machine name]")
end sub

'********************************************************************
'
'   Main
'
'********************************************************************

'
'   Determine script host
'
if (InStr(lcase(WScript.FullName), "cscript") > 0) then
    bCScript = true
else
    bCScript = false
end if

'
'   Get target machine name
'
if (WScript.Arguments.Count = 0) then
    OutputUsage
    WScript.Quit -1
elseif (WScript.Arguments.Count = 1) then
    if ((StrComp(WScript.Arguments(0), lcase("/?"), vbTextCompare)=0) or _
        (StrComp(WScript.Arguments(0), lcase("-?"), vbTextCompare)=0)) then
        OutputUsage
        WScript.Quit -1
    else
        sScheduleID = WScript.Arguments(0)
        sMachine = "."
        OUT "Connecting to local machine"
    end if
else
    if ((StrComp(WScript.Arguments(0), lcase("/?"), vbTextCompare)=0) or _
        (StrComp(WScript.Arguments(0), lcase("-?"), vbTextCompare)=0)) then
        OutputUsage
        WScript.Quit -1
    else
        sScheduleID = WScript.Arguments(0)
        sMachine = WScript.Arguments(1)
        OUT "Connecting to machine " & sMachine
    end if
end if


'
'   Connect to the machine's CCM namespace
'
on error resume next
set oCCMNamespace = GetObject("winmgmts://" & sMachine & "/root/ccm")
if (Err.number <> 0) then
    OUT "Failed to connect to WMI on " & sMachine & ": " & Err.Description & " (" & Err.number & ")"
    WScript.Quit -1
end if
OUT "Successfully connected to WMI"



'
'   Invoke SMS_Client.TriggerSchedule
'
OUT "Triggering Schedule"
on error resume next
Err.Clear
set oInstance = oCCMNamespace.Get("SMS_Client")
set oParams = oInstance.Methods_("TriggerSchedule").inParameters.SpawnInstance_()
oParams.sScheduleID = sScheduleID
oCCMNamespace.ExecMethod "SMS_Client", "TriggerSchedule", oParams
if (Err.number <> 0) then
    OUT "Trigger Schedule Method failed: " & Err.Description & " (" & Err.number & ")"
    WScript.Quit -1
else
    OUT "Schedule " & sScheduleID & " was successfully triggered"
end if


