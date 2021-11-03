Param($computer)
$Ping = new-object system.net.networkinformation.ping
$reply = $ping.send($Computer)
if ($reply.status -eq "Success"){
	[xml]$task = gc "C:\SDCServicing\5.2-SDC-Servicing-scanonly.xml"
	$task.task.triggers.timetrigger.startboundary = [string]((Get-Date).Addminutes(1) | Get-Date -format "yyyy-MM-ddTHH:mm:00")
	if(test-path "\\$computer\c$\windows\temp\SDC-Servicing-Scanonly.xml"){ri "\\$computer\c$\windows\temp\SDC-Servicing-Scanonly.xml"}
	$task.save("\\$computer\c$\windows\temp\SDC-Servicing-Scanonly.xml")
	Schtasks.exe /S $computer /Create /TN "SDC-Servicing-5.2-to-5.3.1-scanonly" /XML "\\$computer\c$\windows\temp\SDC-Servicing-Scanonly.xml"
}
"Sleeping 2 minutes"
Sleep 120
if(test-path "\\$computer\c$\upgrade_os_logs\*_Upgrade_OS.log"){
	if(!(test-path C:\SDCServicing\PreFlight-SUCCESS\)){md C:\SDCServicing\PreFlight-SUCCESS\}
	if(!(test-path C:\SDCServicing\PreFlight-FAIL\)){md C:\SDCServicing\PreFlight-FAIL\}
	$content = gc "\\$computer\c$\upgrade_os_logs\*_Upgrade_OS.log"
	If($content -like "*Preflight checks have been successfully Validated.*"){
	cp "\\$computer\c$\upgrade_os_logs\*_Upgrade_OS.log" C:\SDCServicing\PreFlight-SUCCESS\
	$computer >> C:\SDCServicing\PreFlight-SUCCESS\ServicingReady.txt}
	If($content -like "*Preflight checks have failed*"){
	cp "\\$computer\c$\upgrade_os_logs\*_Upgrade_OS.log" C:\SDCServicing\PreFlight-FAIL\
	$computer >> C:\SDCServicing\PreFlight-SUCCESS\ServicingNOTReady.txt}
}