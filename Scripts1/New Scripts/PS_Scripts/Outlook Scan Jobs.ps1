# Written by SSgt Timothy Brady
# Tyndall AFB, Panama City, FL
# Created February 16, 2016

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
$Program = Read-Host "Program"
$Date = Get-Date -UFormat "%d-%b-%g %H%M"
$Path = "\\XLWUW3-DKPVV1\C$\Users\1392134782A\Documents\$Program Scan Jobs $Date.csv"

# Max number of Jobs to run at one time (60 was found to be the sweet spot)
$MaxThreads = 200

# Number of seconds before the script cancles and hung up jobs (for a small number of computers it can be reduced to ~30)
$TimeOut = 150

# SCRIPT BEGINS
# -------------

$Start = Get-Date

$i = $null
$TotalJobs = $Computers.Count
$Counter = $null

$ScriptBlock = {
    If (Test-Connection $args[0] -Quiet -Count 2 -BufferSize 16 -Ea 0)
    {
        $Ping = "Online"
        
        Try
        {
            $OSInfo = Get-Wmiobject Win32_OperatingSystem -ComputerName $args[0] -ErrorAction SilentlyContinue
            
            If ($OSInfo.OSArchitecture -eq "64-bit"){$RegPath = "Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"}
            ElseIf ($OSInfo.OSArchitecture -eq "32-bit"){$RegPath = "Software\Microsoft\Windows\CurrentVersion\Uninstall"}        
            
            $Reg = [microsoft.win32.registrykey]::OpenRemoteBaseKey('LocalMachine',$args[0])
            $RegKey = $Reg.OpenSubKey($RegPath)
            $SubKeys = $RegKey.GetSubKeyNames()
            
            ForEach($Key in $SubKeys)
            {
                $ThisKey = $RegPath+"\"+$Key 
                $ThisSubKey = $Reg.OpenSubKey($ThisKey)
                $DisplayNames = $ThisSubKey.GetValue("DisplayName")
                ForEach ($Displayname in $DisplayNames)
                {
                    If ($Displayname -like "$($args[1])*")
                    {
                        $Prog = $Displayname
                        $InstallDate = $ThisSubKey.GetValue("InstallDate")
                    }
                }
            }
        }
        
        Catch { $Ping = "No Access" }
    }
    
    Else
    {
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
        System = $args[0]
        Ping = $Ping
        Outlook = $Prog
        Install_Date = $InstallDate
        Organization = $Org
        Building = $Bldg
        Room = $Room
        }

    $Results
}

ForEach ($Computer in $Computers)
{
    Write-Host "Starting Job on: $Computer" -ForegroundColor Cyan
    $i++
    Write-Host "Status: $i / $TotalJobs" -ForegroundColor Yellow

    Start-Job -Name $Computer -ScriptBlock $ScriptBlock -ArgumentList $Computer, $Program | Out-Null

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
        Write-Host "Exiting loop $JobCount Jobs did not complete"
        Get-Job -State Running | Select Name
        Break
    }
}

$Outcome = Get-Job | Receive-Job
$Outcome | Select System, Ping, Outlook, Install_Date, Organization, Building, Room -ExcludeProperty RunspaceId | Export-Csv $Path -Force
Import-Csv $Path | OGV

$Stop = Get-Date
$TimeS = ($Stop - $Start).Seconds
$TimeM = [Math]::Round(($Stop - $Start).TotalMinutes, 0)
Write-Host
Write-Host "Elapsed Time: $TimeM min $TimeS sec" -ForegroundColor Cyan

Get-Process | Where {$_.Name -like "powershell"} | Stop-Process -Force