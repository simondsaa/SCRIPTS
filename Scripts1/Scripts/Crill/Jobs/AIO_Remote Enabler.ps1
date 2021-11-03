#####################################################################################################
#                                                                                                   #
#                                                                                                   #
#                                                                                                   #
#                                                                                                   #
#                                                                                                   #
#                                                                                                   #
#                                                                                                   #
#####################################################################################################
#####################################################################################################
# Script: Workflow Remoting.ps1
# Purpose: Gathers information for Stat folder .CSVs, attempts to repair the PSRemoting an
#          WMI remoting that is inactive on most Tyndall Machines. Utilizing Workflows to speed
#          Up the job.
#          Planning to adopt PSJobs or Runspaces in the future.
# Creator: SSgt Crill, Christian 325 CS/SCOO
# Date: 8/12/2015
#                                                                                                   #
#                                                                                                   #
#                                                                                                   #
#                                                                                                   #
#                                                                                                   #
#                                                                                                   #
#                                                                                                   #
#                                                                                                   #
cls
#                                                                                                   #
#                                                                                                   #
#                                                                                                   #
#                                                                                                   #
#                                                                                                   #
#####################################################################################################
#                                   WORKFLOW AND ENVIRONMENT VARIABLES                              #
#                                                                                                   #
#####################################################################################################
#Use as Remote-Fix TARGET

#Variables
$nl = [Environment]::newline
$DOY = (Get-date).DayOfYear
$continue = $false


#                                                                                                   
#                                                                                                   
#                                                                                                   
#                                                                                                   
##########################################################################################################
#                                           BODY                                                    
# Ping list of Inactive machines
# Filter so the fix is only applied to online machines that have inactive remoting
#
#
#
###########################################################################################################
###########################################################################################################
$nl;$nl;$nl
Do
{
    Cls
    $nl
    $nl
    $nl
    $nl
    Write-Host " "
    Write-Host "Please select your target."
    Write-Host "1 - Single Machine"
    Write-Host "2 - All Inactive machines in Stats folder"
    Write-Host "3 - List of machines"
    Write-Host "4 - Exit"

    Write-Host " "

    $Ans = Read-Host "Make Selection"
    
    If ($Ans -eq 1) #Single
    {
            Write-Host "Enter Computer Name Below" -ForegroundColor Cyan -BackgroundColor DarkGray
            $computer = Read-Host "Target Computername or IP"
            $scriptBlock = {
                    $comp = $args[0]
                    #Reset Variables
                    $WinRMTest = $null
                    $WinRM = $null
                    $method = $null
                    #Test if online
                         IF (Test-Connection -computername $comp -Quiet -ErrorAction SilentlyContinue -BufferSize 16 -Count 1) {
                                Try {
                                $ping = "Online"
                                $comp = $args[0]
                                #const
                                $quote= [char]34
                                #Remote Commands to execute to enable our flavor of remoting.  To be executed via WMI.
                                $command_1 = "schtasks.exe /CREATE /TN 'Minion-Enable-WSRemoting1' /SC ONCE /ST 17:00 /RL HIGHEST /RU SYSTEM /TR $quote powershell.exe -noprofile -command Enable-PSRemoting -Force$quote /F"
                                $command_2 = "schtasks.exe /CREATE /TN 'Minion-Enable-WSRemoting2' /SC ONCE /ST 17:00 /RL HIGHEST /RU SYSTEM /TR $quote powershell.exe -noprofile -command Set-WSManQuickConfig -Force$quote /F"
                                $command_3 = "schtasks.exe /CREATE /TN 'Minion-Enable-WSRemoting3' /SC ONCE /ST 17:00 /RL HIGHEST /RU SYSTEM /TR $quote powershell.exe -noprofile -command Enable-WSManCredSSP -Role Server -Force$quote /F"
                                $command_run1 = "schtasks.exe /RUN /TN 'Minion-Enable-WSRemoting1'"
                                $command_run2 = "schtasks.exe /RUN /TN 'Minion-Enable-WSRemoting2'"
                                $command_run3 = "schtasks.exe /RUN /TN 'Minion-Enable-WSRemoting3'"
                                $command_delete1 = "schtasks.exe /DELETE /TN 'Minion-Enable-WSRemoting1' /F"
                                $command_delete2 = "schtasks.exe /DELETE /TN 'Minion-Enable-WSRemoting2' /F"
                                $command_delete3 = "schtasks.exe /DELETE /TN 'Minion-Enable-WSRemoting3' /F"
                                $process = [WMICLASS]"\\$comp\ROOT\CIMV2:Win32_Process"
                                $result1 = $process.Create($command_1)
                                $result2 = $process.Create($command_run1)
                                $result3 = $process.Create($command_delete1)
                                $result4 = $process.Create($command_2)
                                $result5 = $process.Create($command_run2)
                                $result6 = $process.Create($command_delete2)
                                $result7 = $process.Create($command_3)
                                $result8 = $process.Create($command_run3)
                                $result9 = $process.Create($command_delete3)
                                Write-Host "Completed PSRemote Fix via SCHTASK: $($args[0])" -ForegroundColor Green -BackgroundColor darkgray
                                $method = "SCHTASK"
                                } 
                                CATCH {
                                        Try { 
                                        Write-Host "Running PSEXEC Fix: $($args[0])" -ForegroundColor Cyan -BackgroundColor DarkGray
                                        psexec \\$comp /accepteula -d c:\windows\system32\winrm.cmd quickconfig -quiet
                                        $method = "PSEXEC"}
                                        Catch {
                                        Write-Host "Full Failure" -ForegroundColor Cyan -BackgroundColor darkRed
                                        $method = "FULL FAILURE"}
                                        }
                                FINALLY {
                                        Try {
                                        #Stat Update -TYNDALL ONLY
                                        $Create = schtasks.exe /CREATE /TN "Update Stat" /S "$comp" /SC ONCE /ST 17:00 /RL HIGHEST /RU SYSTEM /TR "powershell.exe -noprofile -File \\area52.afnoapps.usaf.mil\Tyndall_AFB\Logon_Scripts\On_Demand_Stats.ps1" /F
                                        $run = schtasks.exe /RUN /TN "Update Stat" /S "$comp"
                                        Sleep -Milliseconds 10
                                        $delete  = schtasks.exe /DELETE /TN "Update Stat" /s  "$comp" /F
                                        }
                                        Catch {
                                        $method = "Stat Update Fail"
                                         }
         #WinRM test and object output
         $WinRMTest =  Get-Service winrm
              IF ($WinRMTest.Status -eq "Running") {
                    $WinRM = "Active"
             } ELSE {
                    $WINRM = "Inactive"
                    }
                }
        } ELSE {
        $ping = "OFFLINE"
        $method = "OFFLINE"
        Write-Host "System OFFLINE: $comp" -ForegroundColor Red -BackgroundColor Yellow 
         }


        $RemoteObj = [PSCustomObject]@{
                     Computername = $comp
                     Method = $Method
                     Remoting = $WINRM
                     Ping = $ping
                     }  
        $RemoteObj
        $RemoteObj = $null  
}
#
# JOB CONFIG INFO
#
$MaxThreads = 1
$timeout = 60
$Job_Result = $null
$Job_Result = @()
$totalJobs = $computer.Count
            foreach ($comp in $computer) {
                        Start-Job -ScriptBlock $scriptblock -argumentlist $comp |Out-Null
                        While($(Get-Job -State Running).Count -ge $MaxThreads) {
                        Get-Job | Wait-Job -Any -timeout $timeout |Out-Null 
                        }
                        Get-Job -State Completed | % {
                        $outcome= Receive-Job $_ -AutoRemoveJob -Wait
                        }
                        $i++
                        Write-Host "________________Status :$i / $totalJobs" -ForegroundColor Yellow -BackgroundColor DarkGray
                        $Job_Result += $outcome
                        }
        
 Get-job | Remove-Job -force
$i = $null
$continue = $true
    }
        
    
    If ($Ans -eq 2) #All Inactive in Stats
    {
            $directory = "\\xlwu-fs-05pv\Tyndall_PUBLIC\Stats\Current\Computer_stats\Windows\*.*"
            $csvFiles = get-childitem $directory -filter *.csv

            $results = @();


            Foreach ($csv in $csvFiles) {
                $results += import-csv $csv
                $i++
                Write-Host "." -ForegroundColor Cyan
                Write-Progress -activity “Combining Information” -status “Status: $i/$($csvfiles.count)” -PercentComplete (($i / $csvFiles.count)*100)
                 }

            $i = $null



            #Establish collections to be used
            $Active_List = @()
            $Active_List +=  $Results | where {$_.Remoting -eq "Active"} | select ComputerName, Remoting, DayOfYear,IPAddress
            $Inactive_List = @()
            $Inactive_List +=  $Results | where {$_.Remoting -eq "InActive"} | select ComputerName, Remoting, DayOfYear,IPAddress

            #Repair

            #Initial Stats

            $nl;$nl;$nl;$nl;$nl;$nl;$nl;$nl;$nl;$nl
            Write-Host "Initial Statistics"  -ForegroundColor Gray
            Write-Host "Active: $(($Active_List).count)"  -ForegroundColor Green
            Write-Host "Inactive: $(($Inactive_List).count)"  -ForegroundColor Yellow
            Write-Host "Total Machines: $(($results).count)"  -ForegroundColor Cyan
            Sleep -s 1
            #Fix remoting for all logged machines
            Write-Host "Attempting to fix: $(($Inactive_List).count)" -Foregroundcolor Gray
            Sleep -s 1

            $scriptBlock = {
                    $comp = $args[0]
                    #Reset Variables
                    $WinRMTest = $null
                    $WinRM = $null
                    $method = $null
                    #Test if online
                         IF (Test-Connection -computername $comp -Quiet -ErrorAction SilentlyContinue -BufferSize 16 -Count 1) {
                                Try {
                                $ping = "Online"
                                $comp = $args[0]
                                #const
                                $quote= [char]34
                                #Remote Commands to execute to enable our flavor of remoting.  To be executed via WMI.
                                $command_1 = "schtasks.exe /CREATE /TN 'Minion-Enable-WSRemoting1' /SC ONCE /ST 17:00 /RL HIGHEST /RU SYSTEM /TR $quote powershell.exe -noprofile -command Enable-PSRemoting -Force$quote /F"
                                $command_2 = "schtasks.exe /CREATE /TN 'Minion-Enable-WSRemoting2' /SC ONCE /ST 17:00 /RL HIGHEST /RU SYSTEM /TR $quote powershell.exe -noprofile -command Set-WSManQuickConfig -Force$quote /F"
                                $command_3 = "schtasks.exe /CREATE /TN 'Minion-Enable-WSRemoting3' /SC ONCE /ST 17:00 /RL HIGHEST /RU SYSTEM /TR $quote powershell.exe -noprofile -command Enable-WSManCredSSP -Role Server -Force$quote /F"
                                $command_run1 = "schtasks.exe /RUN /TN 'Minion-Enable-WSRemoting1'"
                                $command_run2 = "schtasks.exe /RUN /TN 'Minion-Enable-WSRemoting2'"
                                $command_run3 = "schtasks.exe /RUN /TN 'Minion-Enable-WSRemoting3'"
                                $command_delete1 = "schtasks.exe /DELETE /TN 'Minion-Enable-WSRemoting1' /F"
                                $command_delete2 = "schtasks.exe /DELETE /TN 'Minion-Enable-WSRemoting2' /F"
                                $command_delete3 = "schtasks.exe /DELETE /TN 'Minion-Enable-WSRemoting3' /F"
                                $process = [WMICLASS]"\\$comp\ROOT\CIMV2:Win32_Process"
                                $result1 = $process.Create($command_1)
                                $result2 = $process.Create($command_run1)
                                $result3 = $process.Create($command_delete1)
                                $result4 = $process.Create($command_2)
                                $result5 = $process.Create($command_run2)
                                $result6 = $process.Create($command_delete2)
                                $result7 = $process.Create($command_3)
                                $result8 = $process.Create($command_run3)
                                $result9 = $process.Create($command_delete3)
                                Write-Host "Completed PSRemote Fix via SCHTASK: $($args[0])" -ForegroundColor Green -BackgroundColor darkgray
                                $method = "SCHTASK"
                                } 
                                CATCH {
                                        Try { 
                                        Write-Host "Running PSEXEC Fix: $($args[0])" -ForegroundColor Cyan -BackgroundColor DarkGray
                                        psexec \\$comp /accepteula -d c:\windows\system32\winrm.cmd quickconfig -quiet
                                        $method = "PSEXEC"}
                                        Catch {
                                        Write-Host "Full Failure" -ForegroundColor Cyan -BackgroundColor darkRed
                                        $method = "FULL FAILURE"}
                                        }
                                FINALLY {
                                        Try {
                                        #Stat Update -TYNDALL ONLY
                                        $Create = schtasks.exe /CREATE /TN "Update Stat" /S "$comp" /SC ONCE /ST 17:00 /RL HIGHEST /RU SYSTEM /TR "powershell.exe -noprofile -File \\area52.afnoapps.usaf.mil\Tyndall_AFB\Logon_Scripts\On_Demand_Stats.ps1" /F
                                        $run = schtasks.exe /RUN /TN "Update Stat" /S "$comp"
                                        Sleep -Milliseconds 10
                                        $delete  = schtasks.exe /DELETE /TN "Update Stat" /s  "$comp" /F
                                        }
                                        Catch {
                                        $method = "Stat Update Fail"
                                         }
         #WinRM test and object output
         $WinRMTest =  Get-Service winrm
              IF ($WinRMTest.Status -eq "Running") {
                    $WinRM = "Active"
             } ELSE {
                    $WINRM = "Inactive"
                    }
                }
        } ELSE {
        $ping = "OFFLINE"
        $method = "OFFLINE"
        Write-Host "System OFFLINE: $comp" -ForegroundColor Red -BackgroundColor Yellow 
         }


        $RemoteObj = [PSCustomObject]@{
                     Computername = $comp
                     Method = $Method
                     Remoting = $WINRM
                     Ping = $ping
                     }  
        $RemoteObj
        $RemoteObj = $null  
}
#
# JOB CONFIG INFO
#
$MaxThreads = 10
$timeout = 60
$Job_Result = $null
$Job_Result = @()
$totalJobs = $Inactive_List.Count
            foreach ($comp in $Inactive_list.IpAddress) {
                        Start-Job -ScriptBlock $scriptblock -argumentlist $comp |Out-Null
                        While($(Get-Job -State Running).Count -ge $MaxThreads) {
                        Get-Job | Wait-Job -Any -timeout $timeout |Out-Null 
                        }
                        Get-Job -State Completed | % {
                        $outcome= Receive-Job $_ -AutoRemoveJob -Wait
                        }
                        $i++
                        Write-Host "________________Status :$i / $totalJobs" -ForegroundColor Yellow -BackgroundColor DarkGray
                        $Job_Result += $outcome
                        }
        
 Get-job | Remove-Job -force
$i = $null
$continue = $true

}
    If ($Ans -eq 3) #List of Machines
    {
            Write-Host "Enter absolute path to your .TXT of computernames" -ForegroundColor Cyan -BackgroundColor darkgray
            $nl
            $path = Read-Host "Path:"
            $computers = Get-Content "$path"
            #Initial Stats
            $nl;$nl;$nl;$nl;$nl;$nl;$nl;$nl;$nl;$nl
            Write-Host "Initial Statistics"  -ForegroundColor Gray
            Write-Host "Inactive: $(($computers).count)"  -ForegroundColor Yellow
            Write-Host "Total Machines: $(($computers).count)"  -ForegroundColor Cyan
            Sleep -s 1
            #Fix remoting for all logged machines
            Write-Host "Attempting to fix: $(($computers).count)" -Foregroundcolor Gray
            Sleep -s 1

            $scriptBlock = {
                    $comp = $args[0]
                    #Reset Variables
                    $WinRMTest = $null
                    $WinRM = $null
                    $method = $null
                    #Test if online
                         IF (Test-Connection -computername $comp -Quiet -ErrorAction SilentlyContinue -BufferSize 16 -Count 1) {
                                Try {
                                $ping = "Online"
                                $comp = $args[0]
                                #const
                                $quote= [char]34
                                #Remote Commands to execute to enable our flavor of remoting.  To be executed via WMI.
                                $command_1 = "schtasks.exe /CREATE /TN 'Minion-Enable-WSRemoting1' /SC ONCE /ST 17:00 /RL HIGHEST /RU SYSTEM /TR $quote powershell.exe -noprofile -command Enable-PSRemoting -Force$quote /F"
                                $command_2 = "schtasks.exe /CREATE /TN 'Minion-Enable-WSRemoting2' /SC ONCE /ST 17:00 /RL HIGHEST /RU SYSTEM /TR $quote powershell.exe -noprofile -command Set-WSManQuickConfig -Force$quote /F"
                                $command_3 = "schtasks.exe /CREATE /TN 'Minion-Enable-WSRemoting3' /SC ONCE /ST 17:00 /RL HIGHEST /RU SYSTEM /TR $quote powershell.exe -noprofile -command Enable-WSManCredSSP -Role Server -Force$quote /F"
                                $command_run1 = "schtasks.exe /RUN /TN 'Minion-Enable-WSRemoting1'"
                                $command_run2 = "schtasks.exe /RUN /TN 'Minion-Enable-WSRemoting2'"
                                $command_run3 = "schtasks.exe /RUN /TN 'Minion-Enable-WSRemoting3'"
                                $command_delete1 = "schtasks.exe /DELETE /TN 'Minion-Enable-WSRemoting1' /F"
                                $command_delete2 = "schtasks.exe /DELETE /TN 'Minion-Enable-WSRemoting2' /F"
                                $command_delete3 = "schtasks.exe /DELETE /TN 'Minion-Enable-WSRemoting3' /F"
                                $process = [WMICLASS]"\\$comp\ROOT\CIMV2:Win32_Process"
                                $result1 = $process.Create($command_1)
                                $result2 = $process.Create($command_run1)
                                $result3 = $process.Create($command_delete1)
                                $result4 = $process.Create($command_2)
                                $result5 = $process.Create($command_run2)
                                $result6 = $process.Create($command_delete2)
                                $result7 = $process.Create($command_3)
                                $result8 = $process.Create($command_run3)
                                $result9 = $process.Create($command_delete3)
                                Write-Host "Completed PSRemote Fix via SCHTASK: $($args[0])" -ForegroundColor Green -BackgroundColor darkgray
                                $method = "SCHTASK"
                                } 
                                CATCH {
                                        Try { 
                                        Write-Host "Running PSEXEC Fix: $($args[0])" -ForegroundColor Cyan -BackgroundColor DarkGray
                                        psexec \\$comp /accepteula -d c:\windows\system32\winrm.cmd quickconfig -quiet
                                        $method = "PSEXEC"}
                                        Catch {
                                        Write-Host "Full Failure" -ForegroundColor Cyan -BackgroundColor darkRed
                                        $method = "FULL FAILURE"}
                                        }
                                FINALLY {
                                        Try {
                                        #Stat Update -TYNDALL ONLY
                                        $Create = schtasks.exe /CREATE /TN "Update Stat" /S "$comp" /SC ONCE /ST 17:00 /RL HIGHEST /RU SYSTEM /TR "powershell.exe -noprofile -File \\area52.afnoapps.usaf.mil\Tyndall_AFB\Logon_Scripts\On_Demand_Stats.ps1" /F
                                        $run = schtasks.exe /RUN /TN "Update Stat" /S "$comp"
                                        Sleep -Milliseconds 10
                                        $delete  = schtasks.exe /DELETE /TN "Update Stat" /s  "$comp" /F
                                        }
                                        Catch {
                                        $method = "Stat Update Fail"
                                         }
         #WinRM test and object output
         $WinRMTest =  Get-Service winrm
              IF ($WinRMTest.Status -eq "Running") {
                    $WinRM = "Active"
             } ELSE {
                    $WINRM = "Inactive"
                    }
                }
        } ELSE {
        $ping = "OFFLINE"
        $method = "OFFLINE"
        Write-Host "System OFFLINE: $comp" -ForegroundColor Red -BackgroundColor Yellow 
         }


        $RemoteObj = [PSCustomObject]@{
                     Computername = $comp
                     Method = $Method
                     Remoting = $WINRM
                     Ping = $ping
                     }  
        $RemoteObj
        $RemoteObj = $null  
}
#
# JOB CONFIG INFO
#
$MaxThreads = 10
$timeout = 60
$Job_Result = $null
$Job_Result = @()
$totalJobs = $computers.Count
            foreach ($comp in $computers) {
                        Start-Job -ScriptBlock $scriptblock -argumentlist $comp |Out-Null
                        While($(Get-Job -State Running).Count -ge $MaxThreads) {
                        Get-Job | Wait-Job -Any -timeout $timeout |Out-Null 
                        }
                        Get-Job -State Completed | % {
                        $outcome= Receive-Job $_ -AutoRemoveJob -Wait
                        }
                        $i++
                        Write-Host "________________Status :$i / $totalJobs" -ForegroundColor Yellow -BackgroundColor DarkGray
                        $Job_Result += $outcome
                        }
        
 Get-job | Remove-Job -force
$i = $null
$continue = $true

}
}
Until ($continue -eq $true)
$Job_Result | ogv