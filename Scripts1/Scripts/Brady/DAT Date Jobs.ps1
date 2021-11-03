# Written by SSgt Timothy Brady
# Tyndall AFB, Panama City, FL
# Created September 25, 2015

# MODIFICATIONS
# -------------

# NOTES
# -----

# This script has to be run with an Admin account that has Admin rights to the remote systems.
# There will be a lot of Red text at the end of the script due to failed access to systems hard drive when pulling profiles, it can be ignored.

# CHANGES
# -------

# Text file with the list of computers to run against
$Computers = Get-Content "C:\Users\1180219788A\Desktop\dat.txt"

# Save path for the results
$Date = Get-Date -UFormat "%d-%b-%g %H%M"
$Path = "C:\Users\1180219788A\Desktop\DAT Date Scan Jobs $Date.csv"

# Max number of Jobs to run at one time (60 was found to be the sweet spot)
$MaxThreads = 200

# Number of seconds before the script cancles and hung up jobs (for a small number of computers it can be reduced to ~30)
$TimeOut = 240

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
            
            If ($OSInfo.OSArchitecture -eq "64-bit"){$key = "Software\Wow6432Node\McAfee\AVEngine"}
            ElseIf ($OSInfo.OSArchitecture -eq "32-bit"){$key = "Software\McAfee\AVEngine"}        
            
            $regkey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $args[0])
            $regKey = $regKey.OpenSubKey($key)
            $DATDate = $regKey.GetValue("AVDatDate")
            $DATVer = $regKey.GetValue("AVDatVersion")
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
        DAT_Version = $DATVer
        DAT_Date = $DATDate
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
    Write-Host "________________Status: $i / $TotalJobs" -ForegroundColor Yellow

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
$Outcome | Select System, Ping, DAT_Version, DAT_Date, Organization, Building, Room -ExcludeProperty RunspaceId | Export-Csv $Path -Force
Import-Csv $Path | OGV

$Stop = Get-Date
$TimeS = ($Stop - $Start).Seconds
$TimeM = [Math]::Round(($Stop - $Start).TotalMinutes, 0)
Write-Host
Write-Host "Elapsed Time: $TimeM min $TimeS sec" -ForegroundColor Cyan

Get-Job | Remove-Job -Force