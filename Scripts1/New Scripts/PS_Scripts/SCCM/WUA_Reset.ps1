param($computer)

$Ping = new-object system.net.networkinformation.ping
$reply = $ping.send($computer)
if ($reply.status -eq "Success")
{

write-host "Attempting to reset Windows Update Agent on $computer" -foregroundcolor cyan

$servicelist = "wuauserv","appidsvc","cryptsvc","bits"

$services = get-wmiobject win32_service -computername $computer

Function Stop-wuservices($service)
{
$targetservice = $services | ?{$_.name -eq "$service"}
Write-host "Stopping the $service service"
$stoppingservice = $targetservice.stopservice()
if($stoppingservice.returnvalue -ne 0){
	write-host "$service service was not stopped" -foregroundcolor yellow
	$state = $targetservice.state
	write-host "$service state is $state"}
sleep 10
if((get-service $service -computername $computer).status -eq "Stopped"){Write-host "$service service is Stopped." -foregroundcolor green}
}

foreach($entry in $servicelist){stop-wuservices $entry}

remove-item "\\$computer\c$\programdata\Application Data\Microsoft\Network\Downloader\qmgr*.dat"
If((Test-path "\\$computer\c$\programdata\Application Data\Microsoft\Network\Downloader\qmgr*") -ne "True"){Write-host "Sucessfully removed the QMGR files" -foregroundcolor green}


write-host "Renaming SoftwareDistribution and Catroot2 folders"
if((Test-path \\$computer\c$\windows\SoftwareDistribution) -eq "True"){rename-item \\$computer\c$\windows\SoftwareDistribution \\$computer\c$\windows\SoftwareDistribution.bak}
if((Test-path \\$computer\c$\windows\SoftwareDistribution) -ne "True"){Write-host "SoftwareDistribution folder sucessfully renamed" -foregroundcolor green}
if((Test-path \\$computer\c$\windows\system32\catroot2) -eq "True"){rename-item \\$computer\c$\windows\system32\catroot2 \\$computer\c$\windows\system32\catroot2.bak}
if((Test-path \\$computer\c$\windows\system32\catroot2) -ne "True"){Write-host "Catroot2 folder sucessfully renamed" -foregroundcolor green}

Write-host "Sending remote SC.exe commands to the target system"
$RemoteProcess=([wmiclass]"\\$computer\root\cimv2:win32_process")
$Service1 = $RemoteProcess.create("sc.exe sdset bits D:(A;;CCLCSWRPWPDTLOCRRC;;;SY)(A;;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;BA)(A;;CCLCSWLOCRRC;;;AU)(A;;CCLCSWRPWPDTLOCRRC;;;PU)")
$Service2 = $RemoteProcess.create("sc.exe sdset wuauserv D:(A;;CCLCSWRPWPDTLOCRRC;;;SY)(A;;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;BA)(A;;CCLCSWLOCRRC;;;AU)(A;;CCLCSWRPWPDTLOCRRC;;;PU)")


$dlls = "atl.dll","urlmon.dll","mshtml.dll","shdocvw.dll","browseui.dll","jscript.dll","vbscript.dll","scrrun.dll","msxml.dll","msxml3.dll","msxml6.dll","actxprxy.dll","softpub.dll","wintrust.dll","dssenh.dll","rsaenh.dll","gpkcsp.dll","sccbase.dll","slbcsp.dll","cryptdlg.dll","oleaut32.dll","ole32.dll","shell32.dll","initpki.dll","wuapi.dll","wuaueng.dll","wuaueng1.dll","wucltui.dll","wups.dll","wups2.dll","wuweb.dll","qmgr.dll","qmgrprxy.dll","wucltux.dll","muweb.dll","wuwebv.dll"
write-host "Registering DLLs in the C:\WIndows\System32 directory"
Foreach($entry in $dlls){
$process = $remoteprocess.create("regsvr32.exe /s c:\windows\system32\$entry")}

Write-host "Sending remote NETSH commands"
$netsh = $remoteprocess.create("netsh winsock reset")
$netsh2 = $remoteprocess.create("Netsh winhttp reset proxy")

"Actions complete"
}
else{"$computer is offline"}














