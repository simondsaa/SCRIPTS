#-----------------------------------------------------------------------------------------#
#                                  Written by SrA Timothy Brady                           #
#                                  Tyndall AFB, Panama City, FL                           #
#                                    Created January 28, 2014                             #
#-----------------------------------------------------------------------------------------#
$Computer = "52XLWUW3-DKPVV1"
$User = Get-WmiObject Win32_ComputerSystem -cn $Computer
$Comp = Get-WmiObject Win32_ComputerSystem -cn $Computer -ErrorAction SilentlyContinue
$OS = Get-Wmiobject Win32_OperatingSystem -cn $Computer -ErrorAction SilentlyContinue
$NIC = Get-WmiObject Win32_NetworkAdapterConfiguration -filter "IPEnabled='True'" -cn $Computer -ErrorAction SilentlyContinue | Where-Object {$_.IPAddress -like "131.55*"}
$EDI = $User.UserName
If ($User.UserName -ne $Null)
    {
    $UserInfo = Get-ADUser -Filter {name -like $EDI} -Properties DisplayName, City, gigID, EmailAddress, LockedOut, Enabled, OfficePhone -ErrorAction SilentlyContinue
    }
Write-Host
Write-Host -ForegroundColor Yellow "BEGIN REPORT-------------------------------------------------------------"
Write-Host -ForegroundColor Cyan "-Computer:"
Write-Host "Computer Name    :" $Comp.Name
Write-Host "Operating System :" $OS.Caption"SP" $OS.ServicePackMajorVersion
Write-Host "Installed On     :" $OS.ConvertToDateTime($OS.InstallDate)
Write-Host "IP Address       :" $NIC.IPAddress
Write-Host "MAC Address      :" $NIC.MACAddress
Write-Host -ForegroundColor Cyan "-User:"
If ($User -ne $Null)
    {
    Write-Host "User Logged On   :" $UserInfo.DisplayName
    Write-Host "Base             :" $UserInfo.City
    Write-Host "EDI Number       :" $UserInfo.gigID
    Write-Host "Email Address    :" $UserInfo.EmailAddress
    Write-Host "Telephone number :" $UserInfo.OfficePhone
    Write-Host "Locked Out       :" $UserInfo.LockedOut
    Write-Host "Enabled          :" $UserInfo.Enabled
    }
Else
    {
    Write-Host "No user currently logged on."
    }
Write-Host -ForegroundColor Yellow "END OF REPORT------------------------------------------------------------"
Write-Host