$computers = "52xlwul3-410k93"

#gc "C:\Users\1274873341C\Desktop\Desktop\PS_Scripts\Reg_Fix\targets.txt"

foreach($computer in $computers) {

$svc = Get-Service -Name "RemoteRegistry" -ComputerName "$computer" 
$svc_status = Get-Service -ComputerName "$computer" -DisplayName "Remote Registry" | select Status
$svc_start = Get-Service -ComputerName "$computer" -Name RemoteRegistry | Start-Service
 


    "Connecting to $computer..."
    
    if ($svc_status = "Running"){
    "Connected @ $computer"
    }

    if ($svc_status = "Stopped") {
    "Remote Registry not running on $computer, attempting to start"
    $svc_start
    }

}
