<# Creates and executes a scheduled task.
# This Scheduled task will contain a script that will contain tasks, such as install/uninstall/etc.
#
# Object details for the final variable of $outcome
#
# Target_ID - Variable input from list
# Ping - if ONLINE or OFFLINE
# Task_Status - Did the task schedule
# Error - Recorded Error Message
# Success - True or False value to track and record numbers from push
#>

cls
$nl = [Environment]::newline

# Timer
$sw = new-object system.diagnostics.stopwatch
$sw.reset()
$sw.Start()

# Target List
$Comps = Get-Content "C:\Users\1252862141.adm\Desktop\Scripts1\Pop.txt"

# Script to be pushed to machines
# Setup in a cookie cutter to call an external script that does the job. 

$scriptblock = {
 IF (Test-Connection $args -Quiet -Count 1 -buffersize 16) {
 $ping = "Online"
    $ErrorActionPreference = 'Stop'
    Try {
    $create1 = schtasks.exe /CREATE /SC ONLOGON /TN "Comply2Connect" /S "$Computer" /TR "cscript.exe '\\xlwu-fs-05pv\Tyndall_PUBLIC\NCC_Admin\Tyndall AFB C2C\c2c.vbs'" /F
    $run1 = schtasks.exe /RUN /TN "Comply2Connect" /S "$args"
    Sleep -Milliseconds 10
    $delete1  = schtasks.exe /DELETE /TN "Comply2Connect" /s  "$args" /F
    $success = "True"

    }
    Catch {
                $stop = $error.exception.message
                $success = "False"
                
    } # END CATCH

} 
ELSE {$ping = "OFFLINE"
}
    # Information to be passed to the console and collected
    $RemoteObj = [PSCustomObject]@{
                     Target_ID = $args
                     Ping = $ping
                     Task_Status = $run
                     Error = $stop
                     Success = $success
                     } 
    # Print to console
    $RemoteObj
}

###########################
# JOB CREATION AND CONFIG #
###########################
$i = 0 # Counter
$totalJobs = $comps.Count # Used for calculating counter
$MaxThreads = 60 # Max number of concurrent threads
$Timeout = 60 # Set value if timeout is desired, add on Wait-Job area as -Timeout $timeout

 Foreach ($comp in $comps) {
        Write-Host "Starting Job on: $comp" -ForegroundColor Cyan -BackgroundColor DarkGray
        $i++
        Write-Host "________________Job Number: $i / $totalJobs" -ForegroundColor Yellow -BackgroundColor DarkGray

        Start-Job -name $comp -ScriptBlock $scriptblock -argumentlist $comp |Out-Null
        
        While($(Get-Job -State Running).Count -ge $MaxThreads) {Get-Job | Wait-Job -Any |Out-Null}
} # End ForEach

# $jobs | Add-Member -MemberType NoteProperty -Name Started -Value $null

While ($(Get-Job -state running).count -ne 0){Get-Job -State Running| Wait-Job -timeout 5|Stop-Job
    Write-Host "Remaining Systems : $((Get-job -State Running).count)" -ForegroundColor DarkYellow
}


$outcome = Get-job | Receive-Job # Pull data into $outcome                      
           Get-Job | Remove-Job -force # Delete all jobs
$sw.stop()

#$outcome | Select-Object -Property * -ExcludeProperty RunspaceID |ogv #Print Job 

$objSuccess = ($outcome.Success |where {$_ -eq "True"}).count
$objTotalAttempts = ($outcome.ping |where {$_ -eq "ONLINE"}).count
$num = $objSuccess / $objTotalAttempts
$percent = "{0:N2}%" -f ($num * 100)

#cls
$Date = Get-Date -UFormat "%d-%b-%g %H%M" 

$outcome | Export-Csv C:\Users\1180219788A\Desktop\Job_Results_$date.csv

#Write-Warning "Data may be viewed and manipulated with object named: OUTCOME"
$nl
Write-Host "Statistics" -ForegroundColor Cyan
Write-Host "Total Systems: $(($comps).count)" -ForegroundColor Gray
Write-Host "Systems Online: $(($outcome.ping |where {$_ -eq "ONLINE"}).count)" -foregroundcolor Green
Write-Host "Systems Offline: $(($outcome.ping |where {$_ -eq "OFFLINE"}).count)" -foregroundcolor DarkYellow
Write-Host "Fix Attempted on: $(($outcome.ping |where {$_ -eq "ONLINE"}).count)" -foregroundcolor Gray
Write-Host "Fix CLEARED on: $(($outcome.Success |where {$_ -eq "True"}).count)" -foregroundcolor Green
Write-host "Fix ERRORED on: $(($outcome.Success |where {$_ -eq "False"}).count)" -foregroundcolor Red
Write-host "Success Rate: $percent" -ForegroundColor Yellow -BackgroundColor DarkGray
$nl
Write-host "Elapsed Time: $($sw.Elapsed.Minutes) Minutes" -ForegroundColor Cyan -BackgroundColor DarkGray