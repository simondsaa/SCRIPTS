# Written by SrA Timothy Brady
# Tyndall AFB, Panama City, FL
# Created July 10, 2014

# Ref all Mods outlined within script
# Mod 1: 14-Jul-14 changed menu for better functionality and added option 4
# Mod 2: 29-Jul-14 added SDC to system information for CSTs *doesn't always work*
# Mod 3: 15-Oct-14 removed Modules and editted user gathering portion. Changed the logging location to ADM profile to limit modification requirements
# Mod 4: 16-Dec-14 added a call feature, automatically calls user using Avaya *has specific requirements*

# Edit this to your IP range for your base
$IPRange = "131.55*"

$LogPath = "$env:USERPROFILE\Documents\Logs"

# ================================================================================================================
Function ComputerInfo
{
    If (Test-Connection $Computer -Quiet -BufferSize 16 -Ea 0 -Count 1)
    {
        $Comp = Get-WmiObject Win32_ComputerSystem -cn $Computer -ErrorAction SilentlyContinue
        $OS = Get-WmiObject Win32_OperatingSystem -cn $Computer -ErrorAction SilentlyContinue
        # Mod 2:
        # ------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        $SDC = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$Computer).OpenSubKey('SOFTWARE\Microsoft\Windows\CurrentVersion\OEMInformation').GetValue('Model')
        # ------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        $AD = Get-ADComputer -LDAPFilter "(name=$Computer)" -Properties whenCreated
        $sysuptime = (Get-Date) – [System.Management.ManagementDateTimeconverter]::ToDateTime($OS.LastBootUpTime)
        $NIC = Get-WmiObject Win32_NetworkAdapterConfiguration -filter "IPEnabled='True'" -cn $Computer -ErrorAction SilentlyContinue |
        Where-Object {$_.IPAddress -like "$IPRange"}
        Write-Host "SYSTEM INFO" -ForegroundColor Yellow
        Write-Host "Computer Name      :" $Comp.Name
        Write-Host "System Model       :" $Comp.Manufacturer$Comp.Model
        Write-Host "Operating System   :" $OS.Caption"SP"$OS.ServicePackMajorVersion
        Write-Host "Installed On       :" $OS.ConvertToDateTime($OS.InstallDate)
        Write-Host "Added to Domain    :" $AD.whenCreated
        Write-Host "SDC Version        :" $SDC
        Write-Host "System Bit         :" $Comp.SystemType  
        Write-Host "IP Address         :" $NIC.IPAddress
        Write-Host "MAC Address        :" $NIC.MACAddress
        Write-Host "System Uptime      :" $sysuptime.days"Days"$sysuptime.hours"Hours"$sysuptime.minutes"Min"
    }
    Else
    {
        Write-Host -ForegroundColor Red "$Computer is offline and has been logged in $LogPath\Offline_Systems.txt"
        Out-File "$LogPath\Offline_Systems.txt" -Force -InputObject $Computer -Append
    }
}

# ================================================================================================================
Function ComputerUser
{
    If (Test-Connection $Computer -Quiet -BufferSize 16 -Ea 0 -Count 1)
    {
        # Mod 3:
        # ----------------------------------------------------------------
        $User = Get-WmiObject Win32_ComputerSystem -ComputerName $Computer
        If ($User.UserName -ne $null)
        {
            $DomainEDI = $User.UserName
            $EDI = $DomainEDI.split("\") -replace ".*AREA52"
        # ----------------------------------------------------------------        
            $UserInfo = Get-ADUser "$EDI" -Properties DisplayName, City, gigID, EmailAddress, extensionAttribute5, mDBOverHardQuotaLimit, LockedOut, Enabled, OfficePhone, MemberOf -ErrorAction SilentlyContinue
            $MailSize = ($UserInfo.mDBOverHardQuotaLimit/1024)
            Write-Host "USER INFO           " -ForegroundColor Yellow 
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
                Write-Host "Group Membership   :" $GroupName
            }
            Catch
            {
                $GroupName = "Unavailable"
                Write-Host "Group Membership   :" $GroupName
            }
                    
                    
        }
        Else
        {
            Write-Host "No user logged on." -ForegroundColor Yellow
        }
    }
    Else
    {
        Write-Host -ForegroundColor Red "$Computer is offline and has been logged in $LogPath\Offline_Systems.txt"
        Out-File "$LogPath\Offline_Systems.txt" -Force -InputObject $Computer -Append
    }
    # Mod 4:
    # ---------------------------------------------------------
    $Call = Read-Host "Call User? Y/N"
    If ($Call -eq "Y")
    {
        $number = $UserInfo.OfficePhone.Split("-") -replace "-"
        $a = Get-Process | Where-Object {$_.Name -eq "SMC"}
        $wshell = New-Object -ComObject wscript.shell
        $wshell.AppActivate($a.Id)
        $wshell.sendKeys("{TAB}$number{ENTER}")
    }
    # ---------------------------------------------------------
}

# ================================================================================================================
Function UserInfo
{
    $UserInfo = Get-ADUser "$EDI" -Properties DisplayName, City, gigID, EmailAddress, extensionAttribute5, mDBOverHardQuotaLimit, LockedOut, Enabled, OfficePhone, MemberOf -ErrorAction SilentlyContinue
    $MailSize = ($UserInfo.mDBOverHardQuotaLimit/1024)
    Write-Host "USER INFO" -ForegroundColor Yellow
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
        Write-Host "Group Membership   :" $GroupName
    }
    Catch
    { 
        $GroupName = "Unavailable"
        Write-Host "Group Membership   :" $GroupName
    }
    # Mod 4:
    # ---------------------------------------------------------
    $Call = Read-Host "Call User? Y/N"
    If ($Call -eq "Y")
    {
        $number = $UserInfo.OfficePhone.Split("-") -replace "-"
        $a = Get-Process | Where-Object {$_.Name -eq "SMC"}
        $wshell = New-Object -ComObject wscript.shell
        $wshell.AppActivate($a.Id)
        $wshell.sendKeys("{TAB}$number{ENTER}")
    }
    # ---------------------------------------------------------
}

# ================================================================================================================
# Mod 1:
# ---------------------------------------------------
Do
{
    Cls
    Write-Host
    Write-Host "Log Path: $LogPath"
    Write-Host " "
    Write-Host "1 - Computer Information"
    Write-Host "2 - Logged on User Information"
    Write-Host "3 - User Information"
    Write-Host "4 - Computer & Logged on User Info"
    Write-Host "5 - Exit"
    Write-Host " "

    $Ans = Read-Host "Make Selection"
    
    If ($Ans -eq 1)
    {
        Write-Host
        $Computer = Read-Host "Computer"
        Cls  
        ComputerInfo
        Write-Host
        Pause
    }
    If ($Ans -eq 2)
    {
        Write-Host
        $Computer = Read-Host "Computer"
        Cls
        ComputerUser
        Write-Host
        Pause
    }
    If ($Ans -eq 3)
    {
        Write-Host
        $EDI = Read-Host "EDI Number (eg 1234567891A)"
        Cls
        UserInfo
        Write-Host
        Pause
    }
    If ($Ans -eq 4)
    {
        Write-Host
        $Computer = Read-Host "Computer"
        Cls
        ComputerInfo
        ComputerUser
        Write-Host
        Pause
    }
}
Until ($Ans -eq 5)
# ---------------------------------------------------