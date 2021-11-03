﻿cls

# Jobs - Is Program Installed - Template
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
$Comps = "localhost"
#Program key
$program = "{26A24AE4-039D-4CA4-87B4*"

#Script to be pushed to machines
#Setup in a cookie cutter to call an external script that does the job. 


$scriptblock = {
    $ErrorActionPreference = 'Stop'
    #TRY to accomplish tasks
    TRY {
            #Check if machine ONLINE if OFFLINE do nothing    
            IF (Test-Connection $args[0] -Quiet -Count 1 -buffersize 16) {
                #ONLINE
                $ping = "Online"
                $OSInfo = Get-WmiObject Win32_OperatingSystem #-ComputerName $args[0]
                $success = "True"
                    #Check if Program is Installed 
                    If ($OSInfo.OSArchitecture -eq "64-bit"){
                            $RegPath = "Software\\Wow6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall"
                             }
                    ElseIf ($OSInfo.OSArchitecture -eq "32-bit"){
                            $RegPath = "Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall"}        
                            $Reg = [microsoft.win32.registrykey]::OpenRemoteBaseKey('LocalMachine',$args[0])
                            $RegKey = $Reg.OpenSubKey($RegPath)
                            $SubKeys = $RegKey.GetSubKeyNames()
                            $Array = @()
                                    ForEach($Key in $SubKeys){
                                        If ($Key -like "*") {
                                            $ThisKey = $RegPath+"\"+$Key 
                                            $ThisSubKey = $Reg.OpenSubKey($ThisKey)
                                            $ImagePath = $thisSubKey.GetValue("ImagePath",$null,'DoNotExpandEnvironmentNames')
                                            }
                                        If ($ImagePath -like "*exe*") {
                                             Write-Host "$Computer - $ImagePath"
                                             #"$ThisKey ; $ImagePath" | Out-File \\XLWU-FS-05pv\Tyndall_PUBLIC\ImagePath.txt -Append -Force
                                             }
                                        }

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
                     Target_ID = $args[0]
                     Ping = $ping
                     Architecture = $OSInfo.OSArchitecture
                     Error = $stop
                     Success = $success
                     Imagepath = $ImagePath
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
$Timeout = "2" #Set value if timeout is desired, add on Wait-Job area as -Timeout $timeout

#Job setup
 Foreach ($comp in $comps) {
 #Begining
        Write-Host "Starting Job on: $comp" -ForegroundColor Cyan -BackgroundColor DarkGray
        $i++
        Write-Host "________________Status :$i / $totalJobs" -ForegroundColor Yellow -BackgroundColor DarkGray

        Start-Job -name $comp -ScriptBlock $scriptblock -argumentlist $comp,$program |Out-Null
#During
# If jobs that are still running are greater than or equal to the max threads, it will wait, else it will make more jobs
        While($(Get-Job -State Running).Count -ge $MaxThreads) {Get-Job | Wait-Job -Any |Out-Null}
} #End ForEach

#ENDING
# Wait until no more jobs are listed as Running
#Timer
Write-host "Setting Timers on remaining jobs" -foregroundcolor Darkyellow
$jobs = Foreach ($job in (Get-Job -state Running)){
            $job_timer = new-object system.diagnostics.stopwatch
            $job_timer.Reset()
            $job_timer.Start()
            }

DO {
    write-host "Waiting on: $((Get-Job -State Running).count)"
        Foreach ($job in $job) {
            If ($job_timer.elapsed.seconds -ge "20") { 
            $job | Stop-Job |Out-Null
            $Job_timer.stop()} 
            }
    Sleep -s 1
    }

Until ($(Get-Job -State Running).count -eq 0) #Once there are no more running jobs, pull all information

$outcome = Get-job | Receive-Job #Pull data into $outcome                      
get-Job | Remove-Job -force #Delete all jobs
$sw.stop()

#$outcome | Select-Object -Property * -ExcludeProperty RunspaceID |ogv #Print Job 

$objSuccess = ($outcome.Success |where {$_ -eq "True"}).count
$objTotalAttempts = ($outcome.ping |where {$_ -eq "ONLINE"}).count
$num = $objSuccess / $objTotalAttempts
$percent = "{0:N2}%" -f ($num * 100)

#cls



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
