

$Computers = gc "C:\Users\1274873341C\Desktop\Desktop\PS_Scripts\HBSS\No_Agent_targets.txt"
#$computers = "xlwuw-121x7z"

foreach($computername in $computers)
{
   
#Set executable variables
$path = "\\xlwu-fs-05pv\Tyndall_PUBLIC\Applications\McAfee_Install\ePO9FramePkg"
$executable = "FramePkg.exe"
$switches = "/INSTALL=AGENT /S /FORCEINSTALL"

$cmd = "\\$computername\c$\Program Files (x86)\McAfee\Common Framework\$executable $switches"
$fullpath = $path + "\" + $executable
$ErrorActionPreference="SilentlyContinue"

$Ping = new-object system.net.networkinformation.ping
$reply = $ping.send($computername)
if ($reply.status -eq "Success")
{
    Write-Host "$computername - System online" -ForegroundColor GREEN
    copy-item $fullpath -destination \\$computername\c$\windows\temp
    
    Trap {write-Warning "There was an error connecting to the remote computer or creating the process"; continue}      
    
    Write-Host "Connecting to $computername"
    
    $wmi=([wmiclass]"\\$computername\root\cimv2:win32_process")  
    #bail out if the object didn't get created 
    
    if (!$wmi) {return}  
    
    $remote=$wmi.Create($cmd)  
    
    if ($remote.returnvalue -eq 0) {     
        Write-Host "Successfully launched on $computername" -ForegroundColor GREEN
        } 
    else {     
        Write-Host "Failed to launch $cmd on $computername. ReturnValue is" $remote.ReturnValue -ForegroundColor RED 
    } 
}
else {
    Write-Host "$computername - System offline" -ForegroundColor RED
}
}
