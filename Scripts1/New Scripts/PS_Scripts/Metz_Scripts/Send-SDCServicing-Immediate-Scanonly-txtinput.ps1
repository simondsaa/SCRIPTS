Param($file)
$computers = gc $file
$reached = @()
foreach($computer in $computers){
$Ping = new-object system.net.networkinformation.ping
$reply = $ping.send($Computer)
if ($reply.status -eq "Success"){
	$reached += $computer
	$computer
	[xml]$task = gc "C:\OperationBOHIC\Pre-Flight-All\5.2-SDC-Servicing-scanonly.xml"
	$task.task.triggers.timetrigger.startboundary = [string]((Get-Date).Addminutes(1) | Get-Date -format "yyyy-MM-ddTHH:mm:00")
	if(test-path "\\$computer\c$\windows\temp\SDC-Servicing-Scanonly.xml"){ri "\\$computer\c$\windows\temp\SDC-Servicing-Scanonly.xml"}
	$task.save("\\$computer\c$\windows\temp\SDC-Servicing-Scanonly.xml")
	Schtasks.exe /S $computer /Create /TN "SDC-Servicing-5.2-to-5.3.1-17Aug2017-scanonly" /XML "\\$computer\c$\windows\temp\SDC-Servicing-Scanonly.xml"
}
else{$computer >> c:\operationBOHIC\pre-flight-all\pre-flight-fail\offline.txt
"$computer Offline"} 
}
"Sleeping 2 minutes"
Sleep 120
foreach($computer in $reached){
if(test-path "\\$computer\c$\upgrade_os_logs\*_Upgrade_OS.log"){
	"removing task on $computer and copying files"
	schtasks.exe /delete /tn "SDC-Servicing-5.2-to-5.3.1-31July-scanonly" /S $computer /F
	$content = gc "\\$computer\c$\upgrade_os_logs\*_Upgrade_OS.log"
	If($content -like "*Preflight checks have been successfully Validated.*"){
	$computer >> C:\OperationBOHIC\Pre-Flight-All\Pre-Flight-SUCCESS\ServicingReady.txt
	cp "\\$computer\c$\upgrade_os_logs\*_Upgrade_OS.log" C:\OperationBOHIC\Pre-Flight-All\Pre-Flight-SUCCESS\}
	If($content -like "*Preflight checks have failed*"){
	$computer >> C:\OperationBOHIC\Pre-Flight-All\Pre-Flight-FAIL\ServicingFAILReady.txt
	cp "\\$computer\c$\upgrade_os_logs\*_Upgrade_OS.log" C:\OperationBOHIC\Pre-Flight-All\Pre-Flight-FAIL\}
}}
"Script Complete"
	
	
	

