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
$DirTOP = Get-Childitem "C:\Users\1252862141.adm\Desktop\Scripts1\Pop.txt" -Directory |Sort-object

$comps = $dirtop
#Script to be pushed to machines
#Setup in a cookie cutter to call an external script that does the job. 


$scriptblock = {
    #TRY to accomplish tasks
    TRY {
             $success = "True"
             $task = $args | Get-Childitem -directory -recurse  -ea SilentlyContinue | Get-ACL -ea SilentlyContinue | Where {$_.Access} | Select @{Name="Path";Expression={$_.PSPath.Substring($_.PSPath.IndexOf(“:”)+2) }},Owner,AccesstoString
    } #END Try
    Catch {
                $stop = $error.exception.message
                $success = "False"
                
    } # END CATCH
    $remoteobj = @()
    $remoteobj += $stop
    $remoteobj += $success 
                
                    
$remoteobj += $task
$remoteobj
   } #End Scriptblock


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

#$cleanup = $outcome | select Path,Owner,AccessToString 
#$outcome = $null

cls

$nl
Write-host "Elapsed Time: $($sw.Elapsed.Minutes)" -ForegroundColor Cyan -BackgroundColor DarkGray                        

#$cleanup | ogv