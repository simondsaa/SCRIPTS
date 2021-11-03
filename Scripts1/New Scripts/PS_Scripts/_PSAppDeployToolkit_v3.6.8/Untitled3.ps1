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
{
$activexsetuppath = "C:\TEMP\Flash_x64"
$pluginsetuppath = "C:\TEMP\Flash_x64"
}

Else
{$ccmsetuppath = "C:\TEMP\Flash_x64"}

#Start the ccmsetup uninstall

"Starting the Adobe Flash update"


Sleep 2
""
"Copying installation files to '\\$computername\$activexsetuppath'"
""
#Copy the flash files from Tyndall_PUBLIC\Applications directory

Robocopy \\xlwu-fs-05pv\Tyndall_PUBLIC\Applications\Adobe Flash\Adobe Flash 21.0.0.242\* \\$computername\c$\TEMP\Flash_x64 /e /R:5 /W:5

#Starting the installation process

$startprocess = ([wmiClass]"\\$computername\ROOT\CIMV2:win32_process")
$remoteprocess = $startprocess.create.Invoke("c:\windows\temp\smsclient\ccmsetup.exe SMSSITECODE=AUTO")

If ($remoteprocess.returnvalue -eq 0) {     
        Write-Host "Successfully launched Adobe Flash Update on $computername" -ForegroundColor GREEN
    Write-Host "Process ID: " $remoteprocess.processid -ForegroundColor GREEN
    } 
    Else {     
        Write-Host "Failed to launch Adobe Flash Update on $computername. ReturnValue is" $remoteprocess.ReturnValue -ForegroundColor RED 
	Break
    } 
	sleep 2
	
    Do
	{
		If((get-wmiobject win32_process -computername $computername | where {$_.name -eq '*Flash*'}) -ne $null)
		{
			"Adobe Flash Update still running, waiting 30 Seconds"
			Sleep 30	
		}
	}
	Until ((get-wmiobject win32_process -computername $computername | where {$_.name -eq '*Flash*'}) -eq $null)

	write-host "Adobe Flash update has completed on $computername" -foregroundcolor green


}
