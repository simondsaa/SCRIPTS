Param($computer)

##Create outfile using today's date
$date = get-date -format ddMMMyyyy
$outfile = "$date-SDC-Servicing-outputfile.txt"

##Perform ping on target to verify that it is online or offline 
$Ping = new-object system.net.networkinformation.ping
$reply = $ping.send($Computer)
if ($reply.status -eq "Success"){

"$computer is online"

##The XML that is from the exported task, inputted here as a herestring
$task = [xml]@'
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.2" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Date>2017-07-21T13:13:40.211176</Date>
    <Author>AREA52\1034133855.adm</Author>
    <Description>This will execute the Upgrade-Staging.ps1 from the SCCM management point on a system from a SYSTEM task run as SYSTEM at a random time between 1801 and 2159.</Description>
    <URI>\5.2 SDC Servicing</URI>
  </RegistrationInfo>
  <Triggers>
    <TimeTrigger>
      <StartBoundary>2017-07-21T19:00:00</StartBoundary>
      <Enabled>true</Enabled>
    </TimeTrigger>
  </Triggers>
  <Principals>
    <Principal id="Author">
      <UserId>S-1-5-18</UserId>
      <RunLevel>HighestAvailable</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>true</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>true</StopIfGoingOnBatteries>
    <AllowHardTerminate>true</AllowHardTerminate>
    <StartWhenAvailable>false</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>
    <IdleSettings>
      <StopOnIdleEnd>true</StopOnIdleEnd>
      <RestartOnIdle>false</RestartOnIdle>
    </IdleSettings>
    <AllowStartOnDemand>true</AllowStartOnDemand>
    <Enabled>true</Enabled>
    <Hidden>false</Hidden>
    <RunOnlyIfIdle>false</RunOnlyIfIdle>
    <WakeToRun>false</WakeToRun>
    <ExecutionTimeLimit>PT72H</ExecutionTimeLimit>
    <Priority>7</Priority>
  </Settings>
  <Actions Context="Author">
    <Exec>
      <Command>Powershell.exe</Command>
      <Arguments>-noninteractive -noprofile -executionpolicy bypass "&amp;{\\52tdka-cm-003v.area52.afnoapps.usaf.mil\smspkgd$\TDK00057\Upgrade-Staging.ps1}"</Arguments>
    </Exec>
  </Actions>
</Task>
'@

##get random hour and minute 
$hour = get-random (18..21)
$minute = get-random (01..59)
$starttime = (Get-Date -format "yyyy-MM-ddT") + $hour + ":" + $minute + ":00"

"Scheduling Upgrade for $computer at $starttime"

$task.task.triggers.timetrigger.startboundary = $starttime
$task.save("\\$computer\c$\windows\temp\SDC-Servicing.xml")
$createtask = Schtasks.exe /S $computer /Create /TN "SDC-Servicing-5.2-to-5.3.1" /XML "\\$computer\c$\windows\temp\SDC-Servicing.xml"
$computer >> $outfile
$createtask >> $outfile
$createtask
""
}
Else{"$computer offline" >> $outfile}