# Written by SrA Timothy Brady
# Tyndall AFB, Panama City, FL
# Created July 10, 2014

# MODIFICATIONS
# -------------
# REF: All Mods are outlined within script (#---)
# Mod 1: 14 Jul 14 changed menu for better functionality and added option 4
# Mod 2: 29 Jul 14 added SDC to system information for CSTs *doesn't always work*
# Mod 3: 15 Oct 14 removed Modules and changed the logging location to ADM profile to limit changes needed
# Mod 4: 16 Dec 14 added a call feature, automatically calls user using Avaya *has specific requirements*
# Mod 5: 19 Mar 15 made Mod 4 a Function to be omitted when requirements aren't met
# Mod 6: 22 Sep 15 added option 4 to open the Log folder to view the logs

# NOTES
# -----
# Mod 2 - This will not work with WinXP and some Server versions, I haven't been able to test them all.
# Mod 4 - Specific Requirements are below:
#   1. Must have Avaya UC Client Release 8.1 Version 8.1.5126
#   2. The "Theme" must be set to "standard"
#      a. Select Tools > Preferences > User Interface > Select Theme = standard > Apply

# CHANGES
# -------

# If Mod 4 Requiremetns have been met, change "False" to "True"
$IncludeCallUser = "True"

# Edit this to your IP range for your network
$IPRange = "131.55*"

# SCRIPT BEGINS
# -------------

# Creates Log Path
If (!(Test-Path "$env:USERPROFILE\Documents\Logs"))
{
    New-Item -Path $env:USERPROFILE\Documents\Logs -Type Directory -Force
}

# Mod 3:
# ------------------------------------------
$LogPath = "$env:USERPROFILE\Documents\Logs"
# ------------------------------------------

# ================================================================================================================
Function ComputerInfo
{
    If (Test-Connection $Computer -Quiet -BufferSize 16 -Ea 0 -Count 1)
    {
        
        # Mod 2:
        # ------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	    Try
	    {        
	        $RegOpen = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$Computer)
	        $RegKey = $RegOpen.OpenSubKey('SOFTWARE\Microsoft\Windows\CurrentVersion\OEMInformation')
	        $SDC = $RegKey.GetValue('Model')
	    }
	    Catch
	    {
	        $SDC = "N/A"
	    }
        # ------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	    $Comp = Get-WmiObject Win32_ComputerSystem -cn $Computer -ErrorAction SilentlyContinue
        $OS = Get-WmiObject Win32_OperatingSystem -cn $Computer -ErrorAction SilentlyContinue
        $Disk = Get-WmiObject Win32_LogicalDisk -cn $Computer -Filter "DriveType=3"
        $RAM = [Math]::Round($Comp.TotalPhysicalMemory/1048576, 0)
        $Size = [Math]::Round($Disk.Size/1073741824, 0)
        $FreeSpace = [Math]::Round($Disk.FreeSpace/1073741824, 0)       
	    $AD = Get-ADComputer -LDAPFilter "(name=$Computer)" -Properties whenCreated
        $sysuptime = (Get-Date) – [System.Management.ManagementDateTimeconverter]::ToDateTime($OS.LastBootUpTime)
        $NIC = Get-WmiObject Win32_NetworkAdapterConfiguration -filter "IPEnabled='True'" -cn $Computer -ErrorAction SilentlyContinue |
        Where-Object {$_.IPAddress -like "$IPRange"}
        $Profiles = Get-ChildItem \\$Computer\C$\Users
        $AdminProf = 0
        ForEach ($Profile in $Profiles)
        {
            If ($Profile -like "*.adm")
            {
                $AdminProf += 1
            }
        }
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
        Write-Host "Disk Space         : $Size GB total | $FreeSpace GB free"
        Write-Host "RAM                : $RAM MB"
        Write-Host "Number of Profiles :" $Profiles.Count"total | $AdminProf admin profile(s)"
    }
    Else
    {
        Write-Host -ForegroundColor Red "$Computer is offline and has been logged in $LogPath\Offline_Systems.txt"
        $Space = " "
	    Out-File "$LogPath\Offline_Systems.txt" -Force -InputObject $Date$Space$Computer -Append
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
            $EDI = $User.UserName.TrimStart("AREA52\")
        # ----------------------------------------------------------------        
            $UserInfo = Get-ADUser "$EDI" -Properties DisplayName, City, gigID, EmailAddress, extensionAttribute5, mDBOverHardQuotaLimit, LockedOut, Enabled, OfficePhone, MemberOf, distinguishedName -ErrorAction SilentlyContinue
            $MailSize = ($UserInfo.mDBOverHardQuotaLimit/1024)
            $OU = ($UserInfo.distinguishedName -split ",OU=")[1]
            Write-Host "USER INFO           " -ForegroundColor Yellow 
            Write-Host "Display Name       :" $UserInfo.DisplayName
            Write-Host "Pre-Windows 2000   :" $UserInfo.SamAccountName
            Write-Host "Base Name          :" $UserInfo.City
            Write-Host "Email Address      :" $UserInfo.EmailAddress
            Write-Host "Mail Category      :" $UserInfo.extensionAttribute5
            Write-Host "Box Size Limit     : $MailSize MB"
            Write-Host "Account Locked Out :" $UserInfo.LockedOut
            Write-Host "Account Enabled    :" $UserInfo.Enabled
            Write-Host "Organizational Unit:" $OU
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
	        If ($IncludeCallUser -eq "True")
	        {
		        CallUser
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
        $Space = " "
	    Out-File "$LogPath\Offline_Systems.txt" -Force -InputObject $Date$Space$Computer -Append
    }
}

# ================================================================================================================
Function UserInfo
{
    $UserInfo = Get-ADUser "$EDI" -Properties DisplayName, City, EmailAddress, extensionAttribute5, mDBOverHardQuotaLimit, LockedOut, Enabled, OfficePhone -ErrorAction SilentlyContinue
    $MailSize = ($UserInfo.mDBOverHardQuotaLimit/1024)
    $OU = ($UserInfo.distinguishedName -split ",OU=")[1]
    Write-Host "USER INFO" -ForegroundColor Yellow
    Write-Host "Display Name       :" $UserInfo.DisplayName
    Write-Host "Pre-Windows 2000   :" $UserInfo.SamAccountName
    Write-Host "Base Name          :" $UserInfo.City
    Write-Host "Email Address      :" $UserInfo.EmailAddress
    Write-Host "Mail Category      :" $UserInfo.extensionAttribute5
    Write-Host "Box Size Limit     : $MailSize MB"
    Write-Host "Account Locked Out :" $UserInfo.LockedOut
    Write-Host "Account Enabled    :" $UserInfo.Enabled
    Write-Host "Organizational Unit:" $OU
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
    If ($IncludeCallUser -eq "True")
    {
	    CallUser
    }
}

# ================================================================================================================
# Mod 5:
# -------------------------------------------------------------
Function CallUser
{
    # Mod 4:
    # ---------------------------------------------------------
    $Call = Read-Host "Call User? Y/N"
    If ($Call -eq "Y")
    {
        $Number = $UserInfo.OfficePhone.Split("-") -replace "-"
        $a = Get-Process | Where-Object {$_.Name -eq "SMC"}
        $wshell = New-Object -ComObject wscript.shell
        $wshell.AppActivate($a.Id)
        $wshell.sendKeys("{TAB}$Number{ENTER}")
    }
    # ---------------------------------------------------------
}
# -------------------------------------------------------------

# ================================================================================================================
# Mod 1:
# ---------------------------------------------------
Do
{
    $Date = Get-Date -Format "MMMMM dd, yyyy HH:mm"
    Cls
    Write-Host
    Write-Host " $Date"
    Write-Host
    Write-Host " Log Path: $LogPath"
    Write-Host
    Write-Host " 1 - Computer Information"
    Write-Host " 2 - Computer & Logged on User Info"
    Write-Host " 3 - User Information"
    Write-Host " 4 - Open Log Folder"
    Write-Host " 5 - Exit"
    Write-Host

    $Ans = Read-Host " Make Selection"
    
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
        ComputerInfo
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
    # Mod 6:
    # -------------------
    If ($Ans -eq 4)
    {
        Explorer $LogPath
    }
    # -------------------
}
Until ($Ans -eq 5)
# ---------------------------------------------------