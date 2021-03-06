#-----------------------------------------------------------------------------------------#
#                                  Written by SrA Timothy Brady                           #
#                                  Tyndall AFB, Panama City, FL                           #
#-----------------------------------------------------------------------------------------#
# Specify the Computers you want to query:
$Computers = Get-Content "C:\Users\timothy.brady\Music\Smith.txt"
ForEach ($Computer in $Computers)
{
    If (Test-Connection $Computer -quiet -BufferSize 16 -Ea 0 -count 1)
    {
        Write-Host
        #Write-Host -ForeGroundColor Yellow "BEGIN REPORT-------------------------------------------------------"
        #Write-Host -ForeGroundColor Green "Status           : Available"
        $Comp = Get-WmiObject Win32_ComputerSystem -comp $Computer -ErrorAction SilentlyContinue
        Write-Host -ForeGroundColor White "Computer Name    :" $Comp.Name
        #Write-Host -ForeGroundColor White "User Logged On   :" $Comp.UserName   
        $OS = Get-Wmiobject Win32_OperatingSystem -comp $Computer -ErrorAction SilentlyContinue
        Write-Host -ForeGroundColor White "Operating System :" $OS.Caption"SP" $OS.ServicePackMajorVersion
        #Write-Host -ForeGroundColor White "Installed On     :" $OS.ConvertToDateTime($OS.InstallDate)  
        #$NIC = Get-WmiObject -Class Win32_NetworkAdapterConfiguration -filter "IPEnabled='True'" -cn $Computer -ErrorAction SilentlyContinue |
        #Where-Object {$_.IPAddress -like "131.55*"}
        #Write-Host -ForeGroundColor White "IP Address       :" $NIC.IPAddress
        #Write-Host -ForegroundColor White "MAC Address      :" $NIC.MACAddress
        #Write-Host -ForeGroundColor Yellow "END OF REPORT------------------------------------------------------"
    }
    Else 
    {
        Write-host
        #Write-Host -ForeGroundColor Yellow "BEGIN REPORT-------------------------------------------------------"
        #Write-Host -ForeGroundColor Red "Status           : Unavailable"
        Write-Host "Computer Name    : $Computer"
        #Write-Host "User Logged On   : N/A"
        #Write-Host "Operating System : N/A"
        #Write-Host "Installed On     : N/A"
        Write-Host "IP Address       : N/A"
        #Write-Host "MAC Address      : N/A"
        #Write-Host -ForeGroundColor Yellow "END OF REPORT------------------------------------------------------"
    }
}