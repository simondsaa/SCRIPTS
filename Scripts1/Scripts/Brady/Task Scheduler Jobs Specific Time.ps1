# Written by SSgt Timothy Brady
# Tyndall AFB, Panama City, FL
# Created March 14, 2016

# MODIFICATIONS
# -------------

# NOTES
# -----

# This script has to be run with an Admin account that has Admin rights to the remote systems.
# There will be a lot of Red text at the end of the script due to failed access to systems hard drive when pulling profiles, it can be ignored.

# CHANGES
# -------

# Text file with the list of computers to run against
$Computers = Get-Content "C:\Users\1392134782A\Desktop\Comps.txt"

# Save path for the results
$Date = Get-Date -UFormat "%d-%b-%g %H%M"
$Path = "C:\Users\1392134782A\Documents\Task Schedule $Date.csv"

# Max number of Jobs to run at one time (60 was found to be the sweet spot)
$MaxThreads = 60

# Number of seconds before the script cancles and hung up jobs (for a small number of computers it can be reduced to ~30)
$TimeOut = 150

# SCRIPT BEGINS
# -------------

$Start = Get-Date

$i = $null
$TotalJobs = $Computers.Count
$Counter = $null

$ScriptBlock = {
    $ErrorActionPreference = 'Stop'
    
    Try
    {    
        If (Test-Connection $args[0] -Quiet -Count 1 -BufferSize 16 -Ea 0)
        {
            $Ping = "Online"
        
            $Task = schtasks.exe /CREATE /TN "Reboot" /S $args[0] /SC ONCE /ST 18:00 /RU "INTERACTIVE" /TR "powershell.exe -ExecutionPolicy Unrestricted -WindowStyle Hidden -noprofile -File '\\xlwu-fs-05pv\Tyndall_PUBLIC\NCC Admin\Progess Bar Reboot.ps1'" /F
            Sleep -s 1
            #$Run = schtasks.exe /RUN /TN "HBSS Install" /S $args[0] 
            #Sleep -s 1
            #$Delete = schtasks.exe /DELETE /TN "HBSS Install" /s  $args[0] /F
            $Success = "True"
        }
        Else
        {
            $Ping = "Offline"
        }
    }
    Catch
    {
        $Stop = $Error.exception.message
        $Status = "Failed"
    }

    $RemoteObj = [PSCustomObject]@{
            System = $args[0]
            Ping = $ping
            Task_Status = $run
            Error = $stop
            Success = $success
            }

    $RemoteObj
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

    If ($Counter -gt $TimeOut)
    {
        Write-Host "Exiting loop $JobCount Jobs did not complete"
        Get-Job -State Running | Select Name
        Break
    }
}

$Outcome = Get-Job | Receive-Job
$Outcome | Select System, Ping, Task_Status, Error, Status -ExcludeProperty RunspaceId | Export-Csv $Path -Force
Import-Csv $Path | OGV

$Stop = Get-Date
$TimeS = ($Stop - $Start).Seconds
$TimeM = [Math]::Round(($Stop - $Start).TotalMinutes, 0)
Write-Host
Write-Host "Elapsed Time: $TimeM min $TimeS sec" -ForegroundColor Cyan

Get-Process | Where {$_.Name -like "powershell"} | Stop-Process