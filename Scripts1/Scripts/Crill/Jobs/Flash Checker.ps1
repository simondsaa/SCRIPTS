##########
# Jobs to determine success #
##########

Write-host "Pulling Information on success or fail"



#Target List
$Comps = get-content "c:\users\1394844760A\desktop\office.txt"

#Program key

$program0 = "{A4488E5C-1022-432A-8066*"  #Adobe Flash Player 18 ActiveX
$program1 = "{A580818A-6519-4120-AB1C*"  #Adobe Flash Player 18 NPAPI (Plugin)

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
                $OSInfo = Get-WmiObject Win32_OperatingSystem -ComputerName $args[0]
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
                                    If ($Key -like "$($args[1])"){ #Program0
                                        $ThisKey = $RegPath+"\\"+$Key 
                                        $ThisSubKey = $Reg.OpenSubKey($ThisKey)
                                        $Program_Name0 = $thisSubKey.GetValue("DisplayName")
                                        $Version0 = $thisSubKey.GetValue("DisplayVersion")
                                    }
                                    If ($Key -like "$($args[2])"){ #Program1
                                        $ThisKey = $RegPath+"\\"+$Key 
                                        $ThisSubKey = $Reg.OpenSubKey($ThisKey)
                                        $Program_Name1 = $thisSubKey.GetValue("DisplayName")
                                        $Version1 = $thisSubKey.GetValue("DisplayVersion")
                                    }                                 }
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
                     Program0 = $Program_Name0
                     Version0 = $Version0
                     Program1 = $Program_Name1
                     Version1 = $Version1
                                          } 
    #Print to console
    $RemoteObj

    }


###########################
# JOB CREATION AND CONFIG #
###########################
$i = 0 #Counter
$totalJobs = $comps.Count #Used for calculating counter
$MaxThreads = 40 #Max amount of threads
$counter = 0

 Foreach ($comp in $comps) {
        Write-Host "Starting Job on: $comp" -ForegroundColor Cyan -BackgroundColor DarkGray
        $i++
        Write-Host "________________Status :$i / $totalJobs" -ForegroundColor Yellow -BackgroundColor DarkGray

        Start-Job -name $comp -ScriptBlock $scriptblock -argumentlist $comp, $program0,$program1 |Out-Null
        
        While($(Get-Job -State Running).Count -ge $MaxThreads) {Get-Job | Wait-Job -Any |Out-Null}
} #End ForEach
$ErrorActionPreference = 'SilentlyContinue'

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
           Get-Job | Remove-Job -force #Delete all jobs
$sw.stop()

$num = ($outcome | where{$_.success -eq "True"} |where{$_.Version0 -eq "18.0.0.232" -and $_.Version1 -eq "18.0.0.232"}).count / ($comps).count
$percent = "{0:N2}%" -f ($num * 100)


Write-Warning "Data can be viewed and manipulated with object named: OUTCOME"
$nl
Write-Host "Statistics" -ForegroundColor Cyan
Write-Host "Total Systems: $(($comps).count)" -ForegroundColor Gray
Write-Host "Systems Online: $(($outcome.ping |where {$_ -eq "ONLINE"}).count)" -foregroundcolor Green
Write-Host "Systems Offline: $(($outcome.ping |where {$_ -eq "OFFLINE"}).count)" -foregroundcolor DarkYellow
Write-Host "Fix Attempted on: $(($outcome.ping |where {$_ -eq "ONLINE"}).count)" -foregroundcolor Gray
Write-Host "Fix CLEARED on:$(($outcome | where{$_.success -eq "True"} |where{$_.Version0 -eq "18.0.0.232" -and $_.Version1 -eq "18.0.0.232"}).count)" -foregroundcolor Green
Write-host "Fix ERRORED on: $(($outcome | where{$_.success -eq "True"} |where{$_.Version0 -ne "18.0.0.232" -or $_.Version1 -ne "18.0.0.232"}).count)" -foregroundcolor Red
Write-host "Success Rate: $percent" -ForegroundColor Yellow -BackgroundColor DarkGray
$nl
$nl
Write-host "Elapsed Time: $($sw.Elapsed.Minutes)" -ForegroundColor Cyan -BackgroundColor DarkGray

$outcome | Export-Csv -append "\\xlwu-fs-05pv\Tyndall_PUBLIC\ncc admin\flash\Flash_Install.csv"
$outcome |ogv                        

