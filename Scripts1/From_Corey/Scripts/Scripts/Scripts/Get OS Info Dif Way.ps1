#-----------------------------------------------------------------------------------------#
#                                  Written by SrA Timothy Brady                           #
#                                  Tyndall AFB, Panama City, FL                           #
#-----------------------------------------------------------------------------------------#
Write-Host
Write-Host "BEGIN REPORT--------------------------------------------------------"
$strComputer = "52xlwuw3-410yzp"
$Comp = Get-WmiObject Win32_ComputerSystem -comp $strComputer -ErrorAction SilentlyContinue
Write-Host -ForeGroundColor Black "Computer Name:    " $Comp.Name
Write-Host -ForeGroundColor Black "User Logged On:   " $Comp.UserName   
    $OS = Get-Wmiobject Win32_OperatingSystem -comp $strComputer -ErrorAction SilentlyContinue
    Write-Host -ForeGroundColor Black "Operating System: " $OS.Caption"SP" $OS.ServicePackMajorVersion
    Write-Host -ForeGroundColor Black "Installed On:     " $OS.ConvertToDateTime($OS.InstallDate)  
        $NIC = Get-WmiObject -Class Win32_NetworkAdapterConfiguration -filter "IPEnabled='True'" -cn $strComputer -ErrorAction SilentlyContinue |
        Where-Object {$_.IPAddress -like "131.55*"}
        Write-Host -ForeGroundColor Black "IP Address:       " $NIC.IPAddress
        Write-Host -ForeGroundColor Black "MAC Address:      " $NIC.MACAddress
Write-Host "END OF REPORT-------------------------------------------------------"
Write-Host