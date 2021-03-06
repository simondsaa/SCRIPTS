# ------------------------------------------------------------------------
# NAME: Install-McAfeeFramePkg.ps1
# AUTHOR: Andrew Metzger, 21 CS
# DATE:20 July 2015
#
#
# COMMENTS: This script will copy the McAfee Framepackage you
# specify in the executable variables to a remote system 
# and install the framepacakge. 
#
# ------------------------------------------------------------------------
#PARAM([string]$computername=$(Read-Host -prompt "ComputerName?"))
Param($filename) 
#$Computers = gc $filename
$Computers = gc "C:\Users\1274873341C\Desktop\Desktop\PS_Scripts\HBSS\No_Agent_targets.txt"
#$computers = "xlwuw-121x7z"

foreach($computername in $computers)
{
   
#Set executable variables
$path = "\\xlwu-fs-05pv\Tyndall_PUBLIC\Applications\McAfee_Install\ePO9FramePkg\5.4FramePkg1"
$executable = "FramePkg.exe"
$switches = "/INSTALL=AGENT /S /FORCEINSTALL"

$cmd = "\\$computername\c$\windows\temp\$executable $switches"
$fullpath = $path + "\" + $executable
$ErrorActionPreference="SilentlyContinue"

$Ping = new-object system.net.networkinformation.ping
$reply = $ping.send($computername)
if ($reply.status -eq "Success")
{
    Write-Host "$computername - System online" -ForegroundColor GREEN
    copy-item $fullpath -destination \\$computername\c$\windows\temp
    Write-Host "Source $path\$executable"
    Write-Host "Destination \\$computername\c$\windows\temp"
    Trap {write-Warning "There was an error connecting to the remote computer or creating the process"; continue}      
    Write-Host "Connecting to $computername"
    Write-Host "Process to create is $cmd"
    $wmi=([wmiclass]"\\$computername\root\cimv2:win32_process")  
    #bail out if the object didn't get created 
    if (!$wmi) {return}  
    $remote=$wmi.Create($cmd)  
    if ($remote.returnvalue -eq 0) {     
        Write-Host "Successfully launched $cmd on $computername" -ForegroundColor GREEN
    Write-Host "Process ID: " $remote.processid -ForegroundColor GREEN
    } 
    else {     
        Write-Host "Failed to launch $cmd on $computername. ReturnValue is" $remote.ReturnValue -ForegroundColor RED 
    } 
}
else {
    Write-Host "$computername - System offline" -ForegroundColor RED
}
}
