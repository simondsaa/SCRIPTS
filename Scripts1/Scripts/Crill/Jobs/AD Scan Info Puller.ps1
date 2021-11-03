cls
$nl = [Environment]::newline

#Timer
$sw = new-object system.diagnostics.stopwatch
$sw.reset()
$sw.Start()

#Target List
"Pulling Computers from ADUC"
$AD = Get-ADComputer -Filter *  -Searchbase "OU=TYNDALL AFB, OU=AFCONUSEAST, OU=BASES, DC=AREA52, DC=AFNOAPPS, DC=USAF,DC=MIL"
$Comps = $AD.Name





$scriptblock = {
    #TRY to accomplish tasks
    Start-Sleep -s 10
    TRY {
            #Check if machine ONLINE if OFFLINE do nothing    
            IF (Test-Connection $args[0] -Quiet -Count 1 -buffersize 16) {
                #ONLINE
                $ping = "Online"
                $OSInfo = Get-WmiObject Win32_OperatingSystem -ComputerName $args[0]
                $success = "True"
                    #Check if Program is Installed 
                          $last = (Get-ChildItem "\\$($args[0])\c$\users\*\ntuser.dat" -Force | select @{e={(Split-path $_.Directory -Leaf)}},last* | sort lastwritetime -Descending)[0]
                          $Program_Name = $last.'(Split-path $_.Directory -Leaf)'
                          $ram = (get-wmiobject Win32_PhysicalMemory -ComputerName $args[0]| Measure-Object -Property Capacity -Sum).Sum/1gb
                          $FreeC = (get-wmiobject win32_volume -ComputerName $args[0]| Measure-Object -Property FreeSpace -Sum).Sum/1gb
                          $serial = get-wmiobject Win32_BIOS -computerName $args[0] -Property SerialNumber
                          $user = $Program_Name
                          $trimUser = $user
                          $trimmed = $trimuser.TrimStart("AREA52\")
                          $obj = Get-ADUser -Identity $trimmed -Properties *
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
                     Architecture = $OSInfo.OSArchitecture
                     LastLoggedOnUser = $obj.cn
                     Organization = $obj.o
                     Serial = $serial.SerialNumber
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

$outcome | Export-Csv C:\LastLoggedonUserV2.csv
"Results Exported"
Get-Job | Remove-Job -force #Delete all jobs