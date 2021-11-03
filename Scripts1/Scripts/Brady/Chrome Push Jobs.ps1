# Written by SrA Timothy Brady
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
$Computers = Get-Content "\\XLWUW3-DKPVV1\C$\Users\1392134782A\Desktop\Comps.txt"

# Save path for the results
$Path = "\\XLWUW3-DKPVV1\C$\Users\1392134782A\Documents\Chrome_Push.csv"

# Install package
$Installer = "\\xlwu-fs-05pv\Tyndall_PUBLIC\NCC_Admin\Chrome_43.0.2357.130.msi"

# Max number of Jobs to run at one time (60 was found to be the sweet spot)
$MaxThreads = 20

# Number of seconds before the script cancles and hung up jobs (for a small number of computers it can be reduced to ~30)
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
            $STask = schtasks.exe /CREATE /TN "Chrome" /S $args[0] /SC WEEKLY /D SAT /ST 23:59 /RL HIGHEST /RU SYSTEM /TR "powershell.exe -ExecutionPolicy Unrestricted -WindowStyle Hidden -noprofile -command &{Start-Process Msiexec.exe -Argumentlist /i ,'$($args[1])', /quiet}" /F
            $Run = schtasks.exe /RUN /TN "Chrome" /S $args[0] 
            $Delete = schtasks.exe /DELETE /TN "Chrome" /s  $args[0] /F

            $Task = "Scheduled"
        }

        Catch
        {
            $Ping = "No Access"
            $Task = "Failed"
        }
    }
    
    Else
    {
        $Ping = "Offline"
    }

    $AD = Get-ADComputer -Identity $args[0] -Properties location, o
    $Org = $AD.o
    $Bldg = ($AD.location).Split(";")[0]
    $Room = ($AD.location).Split(";")[1].TrimStart(" ")

    $Results = [PSCustomObject]@{
        System = $args[0]
        Ping = $Ping
        Task = $Task
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

    Start-Job -Name $Computer -ScriptBlock $ScriptBlock -ArgumentList $Computer, $Installer | Out-Null

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
$Outcome | Select System, Ping, Task, Organization, Building, Room -ExcludeProperty RunspaceId | Export-Csv $Path -Force
Import-Csv $Path | OGV

$Stop = Get-Date
$TimeS = ($Stop - $Start).Seconds
$TimeM = [Math]::Round(($Stop - $Start).TotalMinutes, 0)
Write-Host
Write-Host "Elapsed Time: $TimeM min $TimeS sec" -ForegroundColor Cyan

Get-Job | Remove-Job -Force