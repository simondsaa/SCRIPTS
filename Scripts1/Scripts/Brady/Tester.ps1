# Written by SrA Timothy Brady
# Tyndall AFB, Panama City, FL
# Created September 25, 2015

# MODIFICATIONS
# -------------

# NOTES
# -----

# CHANGES
# -------

$Computers = Get-Content "C:\Users\1392134782A\Desktop\Comps.txt"

$Path = "C:\Users\1392134782A\Documents\ADInfo.csv"

$MaxThreads = 70

# SCRIPT BEGINS
# -------------

$Start = Get-Date

$i = $null
$TotalJobs = $Computers.Count
$Counter = $null

$ScriptBlock = {
    $AD = Get-ADComputer -Identity $args[0] -Properties o
    $Org = $AD.o
    
    $Results = [PSCustomObject]@{
        System = $args[0]
        Organization = $Org
        }
    
    $Results
}

ForEach ($Computer in $Computers)
{
    Write-Host "Starting Job on: $Computer" -ForegroundColor Cyan
    $i++
    Write-Host "________________Status: $i / $TotalJobs" -ForegroundColor Yellow

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

    If ($Counter -gt 100)
    {
        Write-Host "Exiting loop $JobCount Jobs did not complete"
        Get-Job -State Running | Select Name
        Break
    }
}

$Outcome = Get-Job | Receive-Job
$Outcome | Select System, Organization -ExcludeProperty RunspaceId | Export-Csv $Path -Force
Import-Csv $Path | OGV -Title "AD Shit"

$Stop = Get-Date
$TimeS = ($Stop - $Start).Seconds
$TimeM = [Math]::Round(($Stop - $Start).TotalMinutes, 0)
Write-Host
Write-Host "Elapsed Time: $TimeM min $TimeS sec" -ForegroundColor Cyan

Get-Job | Remove-Job -Force