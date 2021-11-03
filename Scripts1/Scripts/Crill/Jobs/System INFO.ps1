cls
$nl = [Environment]::newline

#Timer
$sw = new-object system.diagnostics.stopwatch
$sw.reset()
$sw.Start()

#Target List
$Comps =  get-content "C:\users\1394844760A\desktop\offline.txt"



$scriptblock = {
    sleep -Seconds 5
    $obj = Get-ADComputer -Identity $args[0] -Properties CN, LastLogonDate,IPv4Address
    #TRY to accomplish tasks
    TRY {
            #Check if machine ONLINE if OFFLINE do nothing    
            IF (Test-Connection $args[0] -Quiet -Count 1 -buffersize 16) {
                #ONLINE
                $ping = "Online"
                $success = "True"
            ######### True Code block ##########
            $Bit = (Get-WmiObject Win32_OperatingSystem -cn $args[0] -ErrorAction SilentlyContinue).OSArchitecture
            $RAM = [Math]::Round((Get-WmiObject Win32_ComputerSystem -cn $args[0] -ErrorAction SilentlyContinue).TotalPhysicalMemory/1048576, 0)
            $Disk = Get-WmiObject Win32_LogicalDisk -cn $args[0] -Filter "DriveType=3" -ErrorAction SilentlyContinue
            $FreeSpace = [Math]::Round($Disk.FreeSpace/1073741824, 0)
            $Profiles = (Get-ChildItem "\\$($args[0])\C$\Users").Count
            $AD = Get-ADComputer -Identity $args[0] -Properties location, o
            $Org = $AD.o
            $Bldg = ($AD.location).Split(";")[0]
            $Room = ($AD.location).Split(";")[1]
            }
 
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
                     AD_Checkin = $obj.LastLogonDate
                     IP = $obj.IPv4Address
                     SystemBit = $Bit
                     RAM_MB = $RAM
                     FreeDiskSpace_GB = $FreeSpace
                     Organization = $Org
                     Building = $Bldg
                     Room = $Room
                     Profiles = $Profiles
                     Error = $stop
                     Success = $success
                     } 
    #Print to console
    $RemoteObj

    }

###########################
# JOB CREATION AND CONFIG #
###########################
$i = $null #Counter
$totalJobs = $comps.Count #Used for calculating counter
$MaxThreads = 60 #Max amount of threads
$counter = $null


 Foreach ($comp in $comps) {
        Write-Host "Starting Job on: $comp" -ForegroundColor Cyan -BackgroundColor DarkGray
        $i++
        Write-Host "________________Status :$i / $totalJobs" -ForegroundColor Yellow -BackgroundColor DarkGray

        Start-Job -name $comp -ScriptBlock $scriptblock -argumentlist $comp,$program |Out-Null
        
        While($(Get-Job -State Running).Count -ge $MaxThreads) {Get-Job | Wait-Job -Any |Out-Null}
} #End ForEach


While ($(Get-Job -state running).count -ne 0){
$jobcount = (Get-Job -state running).count
Write-Host "Waiting for $jobcount Jobs to Complete: $counter" -foregroundcolor DarkYellow
Start-Sleep -seconds 5
$Counter++

    if ($Counter -gt 40) {
                Write-Host "Exiting loop $jobCount Jobs did not complete"
                get-job  -state Running | select Name
                break
            }

}

$outcome = Get-job | Receive-Job #Pull data into $outcome                      

$sw.stop()


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

$outcome | Export-Csv C:\Offlines.csv
"Results Exported"
Get-Job | Remove-Job -force #Delete all jobs