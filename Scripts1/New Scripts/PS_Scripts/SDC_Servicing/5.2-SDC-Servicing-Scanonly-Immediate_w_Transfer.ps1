Param($computer)
$Ping = new-object system.net.networkinformation.ping
$reply = $ping.send($Computer)
$reached = @()
if ($reply.status -eq "Success"){

    $reached += $computer
	$computer
	
    $sourcefiles1 = "\\xlwu-fs-dfs1v\Tyndall\SDC_531_Upgrade\Upgrade-Staging.ps1"
    $sourcefiles2 = "\\xlwu-fs-dfs1v\Tyndall\SDC_531_Upgrade\Write-PSLogs\Write-PSLogs.psm1"
    $destination = "\\$computer\c$\windows\temp\"


    If (!(Test-Path -path $destination))
            {                        
                New-Item $destination -Type Directory -Force
            }            
                Copy-Item -Path $sourcefiles1 -Destination $destination
                Copy-Item -Path $sourcefiles2 -Destination $destination


    [xml]$task = gc "C:\Operation_UPGRADE\5.2-SDC-Servicing-scanonly.xml"
    $task.task.triggers.timetrigger.startboundary = [string]((Get-Date).Addseconds(10) | Get-Date -format "yyyy-MM-ddTHH:mm:00")
    $task.save("\\$computer\c$\windows\temp\5.2-SDC-Servicing-scanonly.xml")
    Schtasks.exe /S "$computer" /Create /TN "SDC-Servicing to 5.3.1 Scan Only"  /XML "\\$computer\c$\windows\temp\5.2-SDC-Servicing-scanonly.xml"
}
"Sleeping 2 minutes"
Sleep 60
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