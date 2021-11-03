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

Write-host "Testing PS Version" -ForegroundColor DarkYellow
Sleep -s 1
$nl
Write-host "Script works best with Powershell Version 3 or greater" -ForegroundColor DarkYellow
Sleep -s 1
$nl 
IF ($PSVersionTable.PSVersion.Major -ge "3") {
    Write-Host "Powershell version is $($PSVersionTable.PSVersion.Major)" -ForegroundColor Green -BackgroundColor DarkGray
    $Version = "True"
}
Else {
    Write-Host "Powershell version is $($PSVersionTable.PSVersion.Major)" -ForegroundColor Red -BackgroundColor Yellow
    $version = "False"
    }





Write-Host "Starting Timer" -ForegroundColor DarkYellow
#Timer
$sw = new-object system.diagnostics.stopwatch
$sw.reset()
$sw.Start()

$nl
Write-Host "Gathering Initial Directories" -ForegroundColor DarkYellow

#Target List
IF ($version -eq "True") { #PSVersion 3 or greater scan
$DirTOP = Get-Childitem "C:\work\bowling.txt" -Directory |Sort-object | Select Fullname
} 
ELSE { # Universal PSVersion scan
$DirTOP = Get-Childitem "C:\work\bowling.txt" | where {$_.PsIsContainer} | Sort-object | Select-Object FullName
}


#Convert Dir list to universal variable
$comps = $dirtop.fullname
#Script to be pushed to machines
#Setup in a cookie cutter to call an external script that does the job. 


$scriptblock = {
$items = (New-Object -com "WMPlayer.OCX.7").cdromcollection.item(0)            
$items.eject()  

   } #End Scriptblock


###########################
# JOB CREATION AND CONFIG #
###########################
$i = 0 #Counter
$totalJobs = $comps.Count #Used for calculating counter
$MaxThreads = 20 #Max amount of threads
$Timeout = $null #Set value if timeout is desired, add on Wait-Job area as -Timeout $timeout


Write-Host "Starting Jobs" -ForegroundColor DarkYellow
Write-Host "Cancel job now if you dislike the Thread count" -ForegroundColor DarkYellow
$nl;$nl
Write-Host "Thread Count is: $MaxThreads" -ForegroundColor cyan
Write-Host "3 Seconds to Cancel job before threads are created" -ForegroundColor DarkYellow
Sleep -s 3







#Job setup
 Foreach ($comp in $comps) {
 #Begining
        Write-Host "Starting Job on: $comp" -ForegroundColor Cyan -BackgroundColor DarkGray
        $i++
        Write-Host "________________Status :$i / $totalJobs" -ForegroundColor Yellow -BackgroundColor DarkGray

        Start-Job -name $comp -ScriptBlock $scriptblock -argumentlist $comp |Out-Null
#During
# If jobs that are still running are greater than or equal to the max threads, it will wait, else it will make more jobs
        While($(Get-Job -State Running).Count -ge $MaxThreads) {Get-Job | Wait-Job -Any |Out-Null}
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

$nl
Write-host "Elapsed Time: $($sw.Elapsed.TotalMinutes)" -ForegroundColor Cyan -BackgroundColor DarkGray                        

