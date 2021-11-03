$c.Cells.Item(1,1) = "Machine Name"
$c.Cells.Item(1,2) = "Manufacturer"
$c.Cells.Item(1,3) = "System Model"
$c.Cells.Item(1,4) = "Operating System"
$c.Cells.Item(1,5) = "SDC Version"
$c.Cells.Item(1,6) = "System Arch"
$c.Cells.Item(1,7) = "IP Address"
$c.Cells.Item(1,8) = "MAC Address"
$c.Cells.Item(1,9) = "RAM"


Write-Output "Computer Name      : $env:COMPUTERNAME" | Out-File "C:\Temp\System Info.txt" -Force
Write-Output "System Model       : $Man $Model"  | Out-File "C:\Temp\System Info.txt" -Append -Force
Write-Output "Operating System   : $Caption SP $SP"  | Out-File "C:\Temp\System Info.txt" -Append -Force
Write-Output "Installed On       : $Installed"  | Out-File "C:\Temp\System Info.txt" -Append -Force
Write-Output "Added to Domain    : $AD"  | Out-File "C:\Temp\System Info.txt" -Append -Force
Write-Output "SDC Version        : $SDC" | Out-File "C:\Temp\System Info.txt" -Append -Force
Write-Output "System Bit         : $Bit"  | Out-File "C:\Temp\System Info.txt" -Append -Force
Write-Output "IP Address         : $IP"  | Out-File "C:\Temp\System Info.txt" -Append -Force
Write-Output "MAC Address        : $MAC"  | Out-File "C:\Temp\System Info.txt" -Append -Force


$CompMod = Get-WmiObject Win32_ComputerSystem -ErrorAction SilentlyContinue
    $Man = $Comp.Manufacturer
    $Model = $Comp.Model
    $Bit = $Comp.SystemType