#----------------------------------------------------------------------------------
#                           Written by SrA Timothy Brady
#                           Tyndall AFB, Panama City, FL
#                             Created March 5, 2014
# 
# Edited 9-Jul-14 to add more variables for easier sharing
#----------------------------------------------------------------------------------

#Edit this to your user profile
$UserName = "timothy.brady"

#Edit this to your IP range for your base
$IPRange = "131.55*"

Import-Module C:\Users\$UserName\Documents\WindowsPowerShell\Modules\PSTerminalServices\PSTerminalServices.psm1
Import-Module C:\Users\$UserName\Documents\WindowsPowerShell\Modules\PSTerminalServices\PSTerminalServices.psd1
[System.Reflection.Assembly]::LoadFrom("C:\Users\$UserName\Documents\WindowsPowerShell\Modules\PSTerminalServices\Bin\Cassia.dll")

$Message1 = "Unable to pull user information due to invalid Registry Setting."
$Message2 = "Please follow the path below and change 'AllowRemoteRPC' to a '1'"

Write-Host
$strComputer = Read-Host "Computer Name"

If (Test-Connection $strComputer -Quiet -BufferSize 16 -Ea 0 -Count 1)
{
    $Comp = Get-WmiObject Win32_ComputerSystem -cn $strComputer -ErrorAction SilentlyContinue
    $OS = Get-Wmiobject Win32_OperatingSystem -cn $strComputer -ErrorAction SilentlyContinue
    $AD = Get-ADComputer -LDAPFilter "(name=$strComputer)" -Properties whenCreated
    $sysuptime = (Get-Date) – [System.Management.ManagementDateTimeconverter]::ToDateTime($OS.LastBootUpTime)
    $NIC = Get-WmiObject Win32_NetworkAdapterConfiguration -filter "IPEnabled='True'" -cn $strComputer -ErrorAction SilentlyContinue |
    Where-Object {$_.IPAddress -like "$IPRange"}
    Write-Host
    Write-Host -ForegroundColor Yellow "BEGIN REPORT--------------------------------------------------------"
    Write-Host -ForegroundColor Green "Computer -           Online" 
    Write-Host "Computer Name      :" $Comp.Name
    Write-Host "Operating System   :" $OS.Caption"SP"$OS.ServicePackMajorVersion
    Write-Host "Installed On       :" $OS.ConvertToDateTime($OS.InstallDate)
    Write-Host "Added to Domain    :" $AD.whenCreated
    Write-Host "Systmem Bit        :" $Comp.SystemType  
    Write-Host "IP Address         :" $NIC.IPAddress
    Write-Host "MAC Address        :" $NIC.MACAddress
    Write-Host "System Uptime      :" $sysuptime.days"Days"$sysuptime.hours"Hours"$sysuptime.minutes"Min"
    Write-Host
    Try 
    {
        $User = Get-TSSession -ComputerName $strComputer 
        If ($User.UserName -ne "$null")
        {
            Write-Output $User.UserName | Out-File C:\Users\$UserName\Desktop\EDI.txt
            $EDIS = Get-Content C:\Users\$UserName\Desktop\EDI.txt
            ForEach ($EDI in $EDIS)
            {
                If ($EDI -ne "$null")
                {
                    $State = Get-TSSession -ComputerName $strComputer | Where {$_.UserName -Like "$EDI"}
                    $UserInfo = Get-ADUser "$EDI" -Properties DisplayName, City, gigID, EmailAddress, extensionAttribute5, mDBOverHardQuotaLimit, LockedOut, Enabled, OfficePhone, MemberOf -ErrorAction SilentlyContinue
                    $MailSize = ($UserInfo.mDBOverHardQuotaLimit/1024)
                    Write-Host -ForegroundColor Green "User -              " $State.State"- ID"$State.SessionId
                    Write-Host "Display Name       :" $UserInfo.DisplayName
                    Write-Host "Pre-Windows 2000   :" $UserInfo.SamAccountName
                    Write-Host "Base Name          :" $UserInfo.City
                    Write-Host "Email Address      :" $UserInfo.EmailAddress
                    Write-Host "Mail Category      :" $UserInfo.extensionAttribute5
                    Write-Host "Box Size Limit     : $MailSize MB"
                    Write-Host "Account Locked Out :" $UserInfo.LockedOut
                    Write-Host "Account Enabled    :" $UserInfo.Enabled
                    Write-Host "Office Phone       :" $UserInfo.OfficePhone
                    Try 
                    {
                        $Groups = Get-ADPrincipalGroupMembership "$EDI" -ErrorAction SilentlyContinue
                        $GroupName = ""
                        ForEach ($Group in $Groups)
                        {
                            $GroupName += $Group.Name + "
                     "
                        }
                    Write-Host -ForegroundColor Green "Groups -             Loaded"
                    Write-Host "Group Membership   :" $GroupName
                    }
                    Catch
                    {
                        Write-Host -ForegroundColor DarkGreen "Groups -             Not Loaded"
                        Write-Host
                    }
                }
            }
        Write-Host -ForegroundColor Yellow "END OF REPORT-------------------------------------------------------"
        }
        Else 
        {
            Write-Host -ForegroundColor DarkGreen "User -               Null"
            Write-Host -ForegroundColor Yellow "END OF REPORT-------------------------------------------------------"
            Write-Host
        }    
    }
    Catch
    {
        Write-Host -ForegroundColor DarkGreen "User -               Error"
        Write-Host -ForegroundColor Red "Error              : $Message1"
        Write-Host -ForegroundColor Red "                     $Message2"
        Write-Host "Path               : HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server"
        Write-Host "Key                : AllowRemoteRPC     REG_DWORD     0x00000000"
        Write-Host -ForegroundColor Yellow "END OF REPORT-------------------------------------------------------"
    }
}
Else
{
Write-Host
Write-Host -ForegroundColor Yellow "BEGIN REPORT--------------------------------------------------------"
Write-Host -ForegroundColor DarkGreen "Computer -           Offline"
Write-Host "Computer Name      :" $strComputer
Write-Host -ForegroundColor Yellow "END OF REPORT-------------------------------------------------------"
Write-Host
}