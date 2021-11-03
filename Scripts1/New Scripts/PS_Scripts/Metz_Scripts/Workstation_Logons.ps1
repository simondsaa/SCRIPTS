$Query = "select * FROM __InstanceCreationEvent WHERE TargetInstance ISA 'Win32_NTLogEvent' AND TargetInstance.LogFile = 'Security' AND (TargetInstance.EventIdentifier = '4624' OR TargetInstance.EventIdentifier = '4642' OR TargetInstance.EventIdentifier = '4648')"
$computername = "tdkaw-wcsse1"

$ManagementScope = New-Object management.ManagementScope("\\$computername\root\cimv2")
$Eventwatcher = New-Object management.managementEventWatcher($ManagementScope, $Query)

do {
$Event = $Eventwatcher.waitForNextEvent()
$event.TargetInstance.message
Send-MailMessage -Smtpserver "131.15.70.12" -From "MyWorkstation@us.af.mil" -To "andrew.metzger.4@us.af.mil" -Subject "Logon to Workstation" -Body $event.TargetInstance.message
$Eventwatcher.start()
}
while ($true)
