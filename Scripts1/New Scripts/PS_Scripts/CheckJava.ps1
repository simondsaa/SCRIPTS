# Written by SrA Timothy Brady
# Tyndall AFB, Panama City, FL
# Created September 25, 2015

# MODIFICATIONS
# -------------

# NOTES
# -----

# CHANGES
# -------

$Computers = Get-Content "C:\Users\1392134782A\Desktop\Java.txt"

$Path = "C:\Users\1392134782A\Documents\Java.csv"

$MaxThreads = 60

# SCRIPT BEGINS
# -------------

$Start = Get-Date

$i = $null
$TotalJobs = $Computers.Count
$Counter = $null

$ScriptBlock = {
    If (Test-Connection $args[0] -Quiet -BufferSize 16 -Ea 0 -Count 1)
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
                If ($Key -like "{26A24AE4-039D-4CA4-87B4*")
                {
                    $ThisKey = $RegPath+"\"+$Key 
                    $ThisSubKey = $Reg.OpenSubKey($ThisKey)
                    $Java = $thisSubKey.GetValue("DisplayName")
                    If ($Java -like "Java 8 Update 51*")
                    {
                        $Status = "Good"
                    }
                    Else
                    {
                        $Status = "Bad"
                    }
                }
            }
        }
        Catch { }
    }
    Else
    {
        $Ping = "Offline"
    }

    $Results = [PSCustomObject]@{
        System = $args[0]
        Ping = $Ping
        Java = $Java
        Status = $Status
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

    If ($Counter -gt 120)
    {
        Write-Host "Exiting loop $JobCount Jobs did not complete"
        Get-Job -State Running | Select Name
        Break
    }
}

$Outcome = Get-Job | Receive-Job
$Outcome | Select System, Ping, Java, Status -ExcludeProperty RunspaceId | Export-Csv $Path -Force
Import-Csv $Path | OGV -Title "Java Check"

$Stop = Get-Date
$TimeS = ($Stop - $Start).Seconds
$TimeM = [Math]::Round(($Stop - $Start).TotalMinutes, 0)
Write-Host
Write-Host "Elapsed Time: $TimeM min $TimeS sec" -ForegroundColor Cyan

Get-Job | Remove-Job -Force