Param($filename)

$jobdata = gc "C:\Users\1274873341C\Desktop\Desktop\PS_Scripts\Remediation_Tracking\No_SCCM_Client.txt"

$joblimit = 30


foreach($computer in $jobdata)
{
	$entry = ""
	$computer
	$IP = ""	
	$IP = [System.Net.Dns]::gethostaddresses("$computer")
	$entry = $IP.IPAddresstostring
	if($entry -eq ""){"DNS has no entry for $computer"}
	else{
	$entry
	
	$runningjobs = (get-job | where{$_.state -eq "Running"}).count
	"Jobs: " + $runningjobs
	do
	{
	"Sleeping 1 seconds"
	""
	sleep 1
	}
	while ((get-job | where{$_.state -eq "Running"}).count -gt $joblimit)
	
	"**Start Job**"
	
	If(test-Connection -ComputerName $entry -Count 3 -Quiet ){
	
	start-job -args $entry -name $entry -scriptblock{Param($entry)

$sysarch = (gwmi win32_operatingsystem -computername $entry).OSArchitecture
If($sysarch -eq "32-bit")
{$ccmsetuppath = "C:\windows\system32\ccmsetup\ccmsetup.exe"}
else
{$ccmsetuppath = "c:\windows\ccmsetup\ccmsetup.exe"}


$WMISCCMCache = get-wmiobject -class cacheinfoex -namespace ROOT\ccm\Softmgmtagent -computername $entry
foreach($cache in $WMISCCMCache)
{
	If((test-path ("\\$entry\c$" + $CACHE.location.substring(2))) -ne "true")
	{
		"Creating cache folder " + ("\\$entry\c$" + $CACHE.location.substring(2))
		md ("\\$entry\c$" + $CACHE.location.substring(2))
	}
	else
	{
		test-path ("\\$entry\c$" + $CACHE.location.substring(2))
	}
}
$startprocess = ([wmiClass]"\\$entry\ROOT\CIMV2:win32_process")
$startprocess.create.Invoke("$ccmsetuppath /uninstall")
sleep 2
Do
{
	If((get-wmiobject win32_process -computername $entry | where {$_.name -eq 'ccmsetup.exe'}) -ne $null)
	{
		"CCMSETUP.EXE still running, waiting 30 Seconds"
		Sleep 30	
	}
}
Until ((get-wmiobject win32_process -computername $entry | where {$_.name -eq 'ccmsetup.exe'}) -eq $null)
write-host "CCMSETUP.EXE has completed on " $computername -foregroundcolor green

Sleep 2
""
"Installing The Client"
""
Robocopy \\xlwu-fs-05pv\Tyndall_PUBLIC\Applications\SCCM_CB\ccmsetup.exe \\$entry\c$\windows\temp\SMSClient /e /R:5 /W:5

$startprocess = ([wmiClass]"\\$entry\ROOT\CIMV2:win32_process")
$startprocess.create.Invoke("c:\windows\temp\smsclient\ccmsetup.exe SMSSITECODE=AUTO")
	sleep 2
	Do
	{
		If((get-wmiobject win32_process -computername $entry | where {$_.name -eq 'ccmsetup.exe'}) -ne $null)
		{
			"CCMSETUP.EXE still running, waiting 30 Seconds"
			Sleep 30	
		}
	}
	Until ((get-wmiobject win32_process -computername $entry | where {$_.name -eq 'ccmsetup.exe'}) -eq $null)
	write-host "CCMSETUP.EXE has completed on " $entry -foregroundcolor green

}
}
}
}