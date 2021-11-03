cls

# Jobs - Scheduled Task push - Template
# SSgt Crill, Christian 325 CS/SCOO
# 8/20/2015
#
#
# Object details for the final variable of $outcome
#
# Target_ID - Variable input from list
# Ping - if ONLINE or OFFLINE
# Task_Status - Did the task schedule
# Error - Recorded Error Message
# Success - True or False value to track and record numbers from push
#
# 

$nl = [Environment]::newline

#Timer
$sw = new-object system.diagnostics.stopwatch
$sw.reset()
$sw.Start()

#Target List
$Comps = Get-Content "c:\comps.txt"

#Script to be pushed to machines
#Setup in a cookie cutter to call an external script that does the job. 


$scriptblock = {
    $ErrorActionPreference = 'Stop'
    #TRY to accomplish tasks
    TRY {
            #Check if machine ONLINE if OFFLINE do nothing    
            IF (Test-Connection $args -Quiet -Count 1 -buffersize 16) {
                #ONLINE
                $ping = "Online"
                $create = schtasks.exe /CREATE /TN "Fix Vuln" /S $args /SC ONLOGON /RL HIGHEST /RU SYSTEM /TR "powershell.exe -noprofile -File '\\xlwu-fs-05pv\Tyndall_PUBLIC\Update Image Path.ps1'" /F  
                Sleep -s 1
                $run = schtasks.exe /RUN /TN "Fix Vuln" /S $args 
                Sleep -s 1
                $delete = schtasks.exe /DELETE /TN "Fix Vuln" /s  $args /F
                $success = "True"
            } # END IF
            ELSE {
                $ping = "Offline" 
            } # END ELSE
    } #END Try
    Catch {
                $stop = $error.exception.message
                $success = "False"
                
    } # END CATCH

    #Information to be passed to the console and collected
    $RemoteObj = [PSCustomObject]@{
                     Target_ID = $args
                     Ping = $ping
                     Task_Status = $run
                     Error = $stop
                     Success = $success
                     } 
    #Print to console
    $RemoteObj

    }


###########################
# JOB CREATION AND CONFIG #
###########################
$i = 0 #Counter
$totalJobs = $comps.Count #Used for calculating counter
$MaxThreads = 20 #Max amount of threads
$Timeout = 60 #Set value if timeout is desired, add on Wait-Job area as -Timeout $timeout

#Job setup
 Foreach ($comp in $comps) {
 #Begining
        Write-Host "Starting Job on: $comp" -ForegroundColor Cyan -BackgroundColor DarkGray
        $i++
        Write-Host "________________Status :$i / $totalJobs" -ForegroundColor Yellow -BackgroundColor DarkGray

        Start-Job -name $comp -ScriptBlock $scriptblock -argumentlist $comp |Out-Null
#During
# If jobs that are still running are greater than or equal to the max threads, it will wait, else it will make more jobs
        While($(Get-Job -State Running).Count -ge $MaxThreads) {Get-Job | Wait-Job -Any -timeout $Timeout |Out-Null}
} #End ForEach

#ENDING
# Wait until no more jobs are listed as Running
# IF they never end add a timeout
DO {
    write-host "Waiting on: $((Get-Job -State Running).count)" 
    Sleep -Milliseconds 500 
    }

Until ($(Get-Job -State Running).count -eq 0) #Once there are no more running jobs, pull all information

$outcome = Get-job | Receive-Job #Pull data into $outcome                      
get-Job | Remove-Job -force #Delete all jobs
$sw.stop()

$outcome | Select-Object -Property * -ExcludeProperty RunspaceID |ogv #Print Job 

$objSuccess = ($outcome.Success |where {$_ -eq "True"}).count
$objTotalAttempts = ($outcome.ping |where {$_ -eq "ONLINE"}).count
$num = $objSuccess / $objTotalAttempts
$percent = "{0:N2}%" -f ($num * 100)

cls



Write-Warning "Data can be viewed and manipulated with object named: OUTCOME"
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
$nl
Write-host "Elapsed Time: $($sw.Elapsed.Minutes)" -ForegroundColor Cyan -BackgroundColor DarkGray                        
