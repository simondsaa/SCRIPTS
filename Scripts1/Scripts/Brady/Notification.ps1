$LogPath = "\\xlwu-fs-05pv\Tyndall_PUBLIC\Stats\Notifications\ACOMS\" + $env:COMPUTERNAME + ".txt"

$Date = Get-Date -Format "dd-MMM-yy hh:mm"

$EDI = (Get-WmiObject Win32_ComputerSystem).UserName.TrimStart("AREA52\")
$User = (Get-ADUser "$EDI" -Properties DisplayName).DisplayName

$c = New-Object -Comobject wscript.shell
$b = $c.popup("All personnel are requested to log into all PC's, in all positions, for all networks between the hours of 0800-1600 today",0,"Cyber Tuesday Reminder",0)
If ($b -eq 1)
{
    $Log = "ACKNOWLEDGED $Date 

Username: $User
Computer: $env:COMPUTERNAME"
    Out-File -FilePath $LogPath -Force -InputObject $Log
}