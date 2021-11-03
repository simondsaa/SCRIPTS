# Written by SSgt Timothy Brady
# Tyndall AFB, Panama City, FL
# Created May 3, 2016

# MODIFICATIONS
# -------------

# NOTES
# -----

# This script has to be run with an Admin account that has Admin rights to the remote systems.
# There will be a lot of Red text at the end of the script due to failed access to systems hard drive when pulling profiles, it can be ignored.

# CHANGES
# -------

# Text file with the list of computers to run against
$Computers = Get-Content "\\XLWUW3-DKPVV1\C$\Users\1392134782A\Desktop\BaseComputers.txt"

# Save path for the results
$Date = Get-Date -Format "dd MMM yy"
$Path = "\\XLWUW3-DKPVV1\C$\Users\1392134782A\Documents\System Stats $Date.csv"

# Max number of Jobs to run at one time (60 was found to be the sweet spot)
$MaxThreads = 60

# Number of seconds before the script cancels any hung up jobs
$TimeOut = 150

# SCRIPT BEGINS
# -------------

$Start = Get-Date

$i = $null
$TotalJobs = $Computers.Count
$Counter = $null

$ScriptBlock = {
    If (Test-Connection $args[0] -Quiet -Count 1 -BufferSize 16 -Ea 0)
    {
        $Ping = "Online"
        Try
        {
            $Disk = Get-WmiObject Win32_LogicalDisk -cn $args[0] -ErrorAction SilentlyContinue | Where {$_.DeviceID -like "C:"}
            $FreeSpace = [Math]::Round($Disk.FreeSpace/1073741824, 0)
        }
        Catch
        {
            $FreeSpace = "Failed"
        }
        Try
        {
            $CS = Get-WmiObject Win32_ComputerSystem -cn $args[0] -ErrorAction SilentlyContinue
            $Comp = $CS.Name
            $RAM = [Math]::Round($CS.TotalPhysicalMemory/1048576, 0)
            $User = $CS.UserName.TrimStart("AREA52\")
            If ($User -eq $null)
            {
                $User = "No user logged on"
            }
        }
        Catch
        {
            $Comp = $args[0]
            $RAM = "Failed"
            $User = "Failed"
        }
        Try
        {
            $OS = Get-WmiObject Win32_OperatingSystem -cn $args[0] -ErrorAction SilentlyContinue
            $Bit = $OS.OSArchitecture
            $Uptime = (Get-Date) – [System.Management.ManagementDateTimeconverter]::ToDateTime($OS.LastBootUpTime)
            
            $Days = $Uptime.Days
            $Hours = $Uptime.Hours
            $Minutes = $Uptime.Minutes
            $UptimeF = "$Days" +  ":" + $Hours + ":" + $Minutes

            If ((Get-Service -cn $args[0] -Name CcmExec).Status -eq "Running")
            {
                $SCCM = "Running"
            }
            Else
            {
                $SCCM = "Not running"
            }
        }
        Catch
        {
            $Bit = "Failed"
            $UptimeF = "Failed"
        }
        
        $Profiles = (Get-ChildItem "\\$($args[0])\C$\Users").Count
    }
    Else
    {
        $Comp = $args[0]
        $Ping = "Offline"
    }

    $Domain = "OU=Tyndall AFB Computers,OU=Tyndall AFB,OU=AFCONUSEAST,OU=Bases,DC=AREA52,DC=AFNOAPPS,DC=USAF,DC=MIL"
    $objDomain = [adsi]("LDAP://" + $domain)
    $Search = New-Object System.DirectoryServices.DirectorySearcher
    $Search.SearchRoot = $objDomain
    $Search.Filter = "(&(objectClass=computer)(samAccountName=*$($args[0])*))"
    $Search.SearchScope = "Subtree"
    $Results = $Search.FindAll()
    ForEach($Item in $Results)
    {
        $objComputer = $Item.GetDirectoryEntry()
        $Org = (($objComputer.o) | Out-String).Trim()
        $Bldg = ($objComputer.location).Split(";")[0]
        $Room = ($objComputer.location).Split(";")[1].TrimStart(" ")
    }

    $Results = [PSCustomObject]@{
        System = $Comp
        Ping = $Ping
        SystemBit = $Bit
        RAM_MB = $RAM
        FreeDiskSpace_GB = $FreeSpace
        Uptime = $UptimeF
        SCCM_Service = $SCCM
        LoggedOnUser = $User
        Organization = $Org
        Building = $Bldg
        Room = $Room
        Profiles = $Profiles
        }
    
    $Results
}

ForEach ($Computer in $Computers)
{
    Write-Host "Starting Job on: $Computer" -ForegroundColor Cyan
    $i++
    Write-Host "Status: $i / $TotalJobs" -ForegroundColor Yellow

    Start-Job -Name $Computer -ScriptBlock $ScriptBlock -ArgumentList $Computer | Out-Null

    While ($(Get-Job -State Running).Count -ge $MaxThreads)
    {
        Get-Job | Wait-Job -Any | Out-Null
    }
}

While ($(Get-Job -State Running).Count -ne 0)
{
    $JobCount = (Get-Job -State Running).Count
    Start-Sleep -Seconds 1
    $Counter++
    Write-Host "Waiting for $JobCount Jobs to complete: $Counter" -ForegroundColor DarkYellow

    If ($Counter -gt $TimeOut)
    {
        Write-Host "Exiting loop; $JobCount Jobs did not complete"
        Get-Job -State Running | Select Name
        Break
    }
}

$Outcome = Get-Job | Receive-Job
$Outcome | Select System, Ping, SystemBit, RAM_MB, FreeDiskSpace_GB, Uptime, SCCM_Service, LoggedOnUser, Organization, Building, Room, Profiles -ExcludeProperty RunspaceId | Export-Csv $Path -Force
Import-Csv $Path | OGV

$Stop = Get-Date
$TimeS = ($Stop - $Start).Seconds
$TimeM = [Math]::Round(($Stop - $Start).TotalMinutes, 0)
Write-Host
Write-Host "Elapsed Time: $TimeM min $TimeS sec" -ForegroundColor Cyan

Get-Process | Where {$_.Name -like "powershell"} | Stop-Process -Force