$RegOpen = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$env:COMPUTERNAME)
$RegKey = $RegOpen.OpenSubKey('SOFTWARE\Microsoft\Windows\CurrentVersion\OEMInformation')
$SDC = $RegKey.GetValue('Model')
$Comp = Get-WmiObject Win32_ComputerSystem -ErrorAction SilentlyContinue
    $Man = $Comp.Manufacturer
    $Model = $Comp.Model
    $Bit = $Comp.SystemType
$OS = Get-WmiObject Win32_OperatingSystem -ErrorAction SilentlyContinue
    $Caption = $OS.Caption
    $SP = $OS.ServicePackMajorVersion
    $Installed = $OS.ConvertToDateTime($OS.InstallDate)

$AD = (Get-ADComputer -LDAPFilter "(name=$env:COMPUTERNAME)" -Properties whenCreated).whenCreated 


$NIC = Get-WmiObject Win32_NetworkAdapterConfiguration -filter "IPEnabled='True'" -ErrorAction SilentlyContinue |
Where-Object {$_.IPAddress -like "$IPRange"}
    $IP = $NIC.IPAddress
    $MAC = $NIC.MACAddress

Write-Output "Computer Name      : $env:COMPUTERNAME" | Out-File "C:\Temp\System Info.txt" -Force
Write-Output "System Model       : $Man $Model"  | Out-File "C:\Temp\System Info.txt" -Append -Force
Write-Output "Operating System   : $Caption SP $SP"  | Out-File "C:\Temp\System Info.txt" -Append -Force
Write-Output "Installed On       : $Installed"  | Out-File "C:\Temp\System Info.txt" -Append -Force
Write-Output "Added to Domain    : $AD"  | Out-File "C:\Temp\System Info.txt" -Append -Force
Write-Output "SDC Version        : $SDC" | Out-File "C:\Temp\System Info.txt" -Append -Force
Write-Output "System Bit         : $Bit"  | Out-File "C:\Temp\System Info.txt" -Append -Force
Write-Output "IP Address         : $IP"  | Out-File "C:\Temp\System Info.txt" -Append -Force
Write-Output "MAC Address        : $MAC"  | Out-File "C:\Temp\System Info.txt" -Append -Force