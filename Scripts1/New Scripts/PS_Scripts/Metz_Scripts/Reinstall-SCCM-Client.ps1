Param($computername = (read-host "Enter System name: "))

#Ping the target system to verify it is online

$ping = new-object system.net.networkinformation.ping
$reply = $ping.send($computername)


If ($reply.status -eq "Success")
{
""
"$computername is online"

#get system architecture to verify where

$sysarch = (gwmi win32_operatingsystem -computername $computername).OSArchitecture

"System Architecture: $sysarch"

If($sysarch -eq "32-bit")
{$ccmsetuppath = "C:\windows\system32\ccmsetup\ccmsetup.exe"}

Else
{$ccmsetuppath = "c:\windows\ccmsetup\ccmsetup.exe"}

#Get the WMI class to repopulate the sccm cache folders pror to client uninstall

$WMISCCMCache = get-wmiobject -class cacheinfoex -namespace ROOT\ccm\Softmgmtagent -computername $computername

#Recreate the missing cache folders so the uninstall doesn't hang up

ForEach($cache in $WMISCCMCache)
{
	If((test-path ("\\$computername\c$" + $CACHE.location.substring(2))) -ne "true")
	{
		"Creating cache folder " + ("\\$computername\c$" + $CACHE.location.substring(2))
		md ("\\$computername\c$" + $CACHE.location.substring(2))
	}
	
    ElseIf((test-path ("\\$computername\c$" + $CACHE.location.substring(2))) -eq "true")
	{
		$cache.Location + " exits"
	}
}

#Start the ccmsetup uninstall

"Starting the SCCM Client uninstall"

$startprocess = ([wmiClass]"\\$computername\ROOT\CIMV2:win32_process")
$remoteprocess = $startprocess.create.Invoke("$ccmsetuppath /uninstall")

If ($remoteprocess.returnvalue -eq 0) {     
        Write-Host "Successfully launched SCCM Uninstall on $computername" -ForegroundColor GREEN
    Write-Host "Process ID: " $remoteprocess.processid -ForegroundColor GREEN
    } 

Else {     
        Write-Host "Failed to launch SCCM Uninstall on $computername. ReturnValue is" $remoteprocess.ReturnValue -ForegroundColor RED 
	break
    } 

sleep 2

Do
{
	If((get-wmiobject win32_process -computername $computername | where {$_.name -eq 'ccmsetup.exe'}) -ne $null)
	{
		write-host "CCMSETUP.EXE still running, waiting 30 Seconds"
		Sleep 30	
	}
}
Until ((get-wmiobject win32_process -computername $computername | where {$_.name -eq 'ccmsetup.exe'}) -eq $null)
write-host "CCMSETUP.EXE has completed on " $computername -foregroundcolor green

Sleep 2
""
"Copying SCCM Client to $computername"
""
#Copy the client from the SCCM Management point

Robocopy \\52xlwu-cm-001v\SMSClient \\$computername\c$\windows\temp\SMSClient /e /R:5 /W:5

#Starting the installation process

"Starting the install with the new client"

$startprocess = ([wmiClass]"\\$computername\ROOT\CIMV2:win32_process")
$remoteprocess = $startprocess.create.Invoke("c:\windows\temp\smsclient\ccmsetup.exe SMSSITECODE=AUTO")

If ($remoteprocess.returnvalue -eq 0) {     
        Write-Host "Successfully launched SCCM install on $computername" -ForegroundColor GREEN
    Write-Host "Process ID: " $remoteprocess.processid -ForegroundColor GREEN
    } 
    Else {     
        Write-Host "Failed to launch SCCM install on $computername. ReturnValue is" $remoteprocess.ReturnValue -ForegroundColor RED 
	Break
    } 
	sleep 2
	
    Do
	{
		If((get-wmiobject win32_process -computername $computername | where {$_.name -eq 'ccmsetup.exe'}) -ne $null)
		{
			"CCMSETUP.EXE still running, waiting 30 Seconds"
			Sleep 30	
		}
	}
	Until ((get-wmiobject win32_process -computername $computername | where {$_.name -eq 'ccmsetup.exe'}) -eq $null)

	write-host "CCMSETUP.EXE has completed on $computername" -foregroundcolor green


}
