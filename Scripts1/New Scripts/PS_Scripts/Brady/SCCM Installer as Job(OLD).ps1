# Written by SSgt Timothy Brady
# Tyndall AFB, Panama City, FL
# Created July 14, 2016

# CHANGES
# -------

# Max number of Jobs to run at one time (60 was found to be the sweet spot on a Quad Core @ 3.3 GHz)
$MaxThreads = 60

# Number of seconds before the script cancles any hung up jobs (for a small number of computers, <500, it can be reduced to ~30)
$TimeOut = 150

# SCRIPT BEGINS
# -------------

# Formats the date for file naming
$Date = Get-Date -UFormat "%d-%b-%g %H%M"

# Get your logged on info to query text files and save results
$LocalUser = Get-WmiObject Win32_ComputerSystem
$LocalEDI = $LocalUser.UserName.Split("\")[1]
$LocalName = (Get-ADUser "$LocalEDI" -Properties DisplayName).DisplayName

# Save path for the results
$Path = "C:\Users\$LocalEDI\Documents\Task Schedule $Date.csv"

# List text files to chose a list of computers from (deaults to your desktop)
[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
$dialog = New-Object System.Windows.Forms.OpenFileDialog
$dialog.Filter = 'Text Files|*.txt|All Files|*.*'
$dialog.FilterIndex = 0
$dialog.InitialDirectory = "C:\Users\$LocalEDI\Desktop"
$dialog.Multiselect = $false
$dialog.RestoreDirectory = $true
$dialog.Title = "Select File with Computer Names"
$dialog.ValidateNames = $true
$dialog.ShowDialog()

Try 
{
    $Computers = Get-Content $dialog.FileName
}

Catch
{
    Exit
}

$Start = Get-Date

$i = $null
$i2 = $null
$i3 = $null
$TotalJobs = $Computers.Count
$Counter = $null

cls

# Gathers available packages currently in RAP
$SCCM = New-Object -ComObject UIResource.UIResourceMgr
$Apps = $SCCM.GetAvailableApplications() | Select ID, PackageID, PackageName

# Builds an array to save all the packages in
$Array = @()
ForEach ($App in $Apps)
{
    $i3 ++
    $Array += "$i3 " + ":" + " " + $App.PackageName
}

# Displays all the packages and prompts for which one you want to run
Write-Host
Write-Host "Available Packages in RAP:"
Write-Host
$Array
Write-Host
$Selection = (Read-Host -Prompt "Select the number to run the program") - 1

$Pick = $Array.Get("$Selection")
$Package = $Pick.Split(":")[1].TrimStart(" ")

Write-Host
Write-Host "`"$Package`" will be ran..."

# Sets the installer to the package you selected
$Installer = $Apps | Where {$_.PackageName -like "$Package"}

$ID = $Installer.Id
If ($ID -eq "`*")
{
    $ID = '`' + $Installer.Id
}

$PackageID = $Installer.PackageId

# Builds the script for the SCCM install
$Script = "`$SCCM = New-Object -ComObject UIResource.UIResourceMgr
" + "`$SCCM.ExecuteProgram(`"$ID`",`"$PackageID`",`$true)" | Out-File -FilePath "\\xlwu-fs-05pv\Tyndall_PUBLIC\NCC Admin\SCCM_Installer.ps1" -Force

$Script = "SCCM_Installer.ps1"

# Prompts for a reboot if needed/wanted
Write-Host
$Reboot = Read-Host -Prompt "Does this require a reboot? Yes/No"

# Script block for the remote system
$ScriptBlock = {
    $ErrorActionPreference = 'Stop'
    
    Try
    {    
        If (Test-Connection $args[0] -Quiet -Count 4 -BufferSize 16 -Ea 0)
        {
            $Ping = "Online"
            
            # Creates a "C:\temp" folder if it doesn't already exist on teh system
            If (!(Test-Path "\\$($args[0])\c$\temp"))
            {
                New-Item -Path "\\$($args[0])\c$\temp" -Type Directory -Force
            }
            # Copies the script to the local system to prevent any security warning for the remote user
            Copy-Item -Path "\\xlwu-fs-05pv\Tyndall_PUBLIC\NCC Admin\$($args[1])" -Destination "\\$($args[0])\c$\temp" -Force
        
            # Schedules the task to run the script, runs the task, then deletes the task
            $Task = schtasks.exe /CREATE /TN "SCCM" /S $args[0] /SC ONLOGON /RU "INTERACTIVE" /TR "powershell.exe -WindowStyle Hidden -noprofile -File 'C:\temp\$($args[1])'" /F
            Sleep -s 1
            $Run = schtasks.exe /RUN /TN "SCCM" /S $args[0] 
            Sleep -s 1
            $Delete = schtasks.exe /DELETE /TN "SCCM" /S $args[0] /F
            $Status = "Successful"

            # Deletes the script from the "C:\temp" folder
            Remove-Item -Path "\\$($args[0])\c$\temp\$($args[1])" -Force

            # If you answered "Yes" to the reboot, this is where it's done
            If ($args[2] -eq "Yes")
            {
                # Copies the reboot script to the local system to prevent any security warning for the remote user
                Copy-Item -Path "\\xlwu-fs-05pv\Tyndall_PUBLIC\NCC Admin\Progress Bar Reboot.ps1" -Destination "\\$($args[0])\c$\temp" -Force
        
                # Schedules the task to run the script, runs the task, then deletes the task
                $Task = schtasks.exe /CREATE /TN "Reboot" /S $args[0] /SC ONLOGON /RU "INTERACTIVE" /TR "powershell.exe -WindowStyle Hidden -noprofile -File 'C:\temp\Progress Bar Reboot.ps1'" /F
                Sleep -s 1
                $Run = schtasks.exe /RUN /TN "Reboot" /S $args[0] 
                Sleep -s 1
                $Delete = schtasks.exe /DELETE /TN "Reboot" /S $args[0] /F
                $Status = "Successful"

                # Deletes the reboot script from the "C:\temp" folder
                Remove-Item -Path "\\$($args[0])\c$\temp\Progress Bar Reboot.ps1" -Force
            }  
        }
        Else
        {
            $Ping = "Offline"
        }
    }
    Catch
    {
        # Error catching for end report
        $Stop = $Error.exception.message
        $Status = "Failed"
    }

    # Result output for report
    $Results = [PSCustomObject]@{
            System = $args[0]
            Ping = $Ping
            Task_Status = $Run
            Error = $Stop
            Status = $Status
            }

    $Results
}

ForEach ($Computer in $Computers)
{
    $i++

    # Progress bar to show installer status
    [int]$pct = ($i/$TotalJobs)*100
    Write-Progress -Activity "Running on System - $Computer" -Status "Starting $i / $TotalJobs" -PercentComplete $pct
    
    Start-Job -Name $Computer -ScriptBlock $ScriptBlock -ArgumentList $Computer, $Script, $Reboot | Out-Null

    # Waiting point for jobs once they reach the max threads
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

    # Progress bar countdown for jobs still running
    [Int]$pct2 = ($Counter/$TimeOut)*100
    ForEach ($Sec in $TimeOut)
    {
        $i2++
        $SecLeft = $TimeOut - $i2
        $Min = [Int](([String]($SecLeft/60)).split('.')[0])
    }
            
    Write-Progress -Activity "Waiting for $JobCount jobs to complete..." -Status "$Min minute(s) $($SecLeft % 60) seconds left" -PercentComplete $pct2

    # Timeout for jobs that hang up and do not complete before the specified timeout
    If ($Counter -gt $TimeOut)
    {
        Write-Host "Exiting loop $JobCount Jobs did not complete"
        Get-Job -State Running | Select Name
        Break
    }
}

# Output for the end report
$Outcome = Get-Job | Receive-Job
$Outcome | Select System, Ping, Task_Status, Error, Status -ExcludeProperty RunspaceId | Export-Csv $Path -Force
Import-Csv $Path | OGV

# Timelaps for script
$Stop = Get-Date
$TimeS = ($Stop - $Start).Seconds
$TimeM = [Math]::Round(($Stop - $Start).TotalMinutes, 0)
Write-Host
Write-Host "Elapsed Time: $TimeM min $TimeS sec" -ForegroundColor Cyan

# Closing all the jobs
Get-Job | Remove-Job -Force