$computer = "xlwuw-421nkx"
$svc = Get-WmiObject -Class Win32_Service -Computer $computer -Filter "Name='RemoteRegistry'" -ErrorAction SilentlyContinue 

    if ($svc){
    "Connected @ $svc"
    exit 1
    }

    if (-not $svc) {
    "Cannot connect to $computer."
    exit 1
    }

if ($svc.State -eq 'Stopped') { $svc.StartService() }