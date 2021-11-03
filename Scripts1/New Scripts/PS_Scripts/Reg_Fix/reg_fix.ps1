$computers = gc "C:\Users\1274873341C\Desktop\Desktop\PS_Scripts\Reg_Fix\targets.txt"

foreach($computer in $computers) {

$svc = Get-WmiObject -Class Win32_Service -Computer $computer -Filter "Name='RemoteRegistry'" -ErrorAction SilentlyContinue 

    "Connecting to $computer..."
    
    if ($svc){
    "Connected @ $svc"
    }

    if (-not $svc) {
    "Cannot connect to $computer."
    }

if ($svc.State -eq 'Stopped') { Get-WmiObject -Class Win32_Service -Filter "Name='RemoteRegistry'" | Start-Service }
if ($svc.State -eq 'Running') { "Remote Registry is running" }
}