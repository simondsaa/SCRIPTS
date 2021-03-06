#-----------------------------------------------------------------------------------------#
#                                  Written by SrA Timothy Brady                           #
#                                  Tyndall AFB, Panama City, FL                           #
#                                   Created December 18, 2013                             #
#-----------------------------------------------------------------------------------------#
$strComputer = Read-Host "Computer Name"
Write-Host
Write-Host -ForegroundColor Yellow "BEGIN REPORT--------------------------------------------------------"
$Comp = Get-WmiObject Win32_ComputerSystem -comp $strComputer -ErrorAction SilentlyContinue
Write-Host "Computer Name    :" $Comp.Name
Write-Host "User Logged On   :" $Comp.UserName   
    $OS = Get-Wmiobject Win32_OperatingSystem -comp $strComputer -ErrorAction SilentlyContinue
    Write-Host "Operating System :" $OS.Caption"SP"$OS.ServicePackMajorVersion
    Write-Host "Installed On     :" $OS.ConvertToDateTime($OS.InstallDate)  
        $NIC = Get-WmiObject -Class Win32_NetworkAdapterConfiguration -filter "IPEnabled='True'" -cn $strComputer -ErrorAction SilentlyContinue |
        Where-Object {$_.IPAddress -like "131.55*"}
        Write-Host "IP Address       :" $NIC.IPAddress
        Write-Host "MAC Address      :" $NIC.MACAddress
Write-Host -ForegroundColor Yellow "END OF REPORT-------------------------------------------------------"
Write-Host