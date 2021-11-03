$directory = "\\xlwu-fs-05pv\Tyndall_PUBLIC\Stats\Current\UL_Clients\*.*"
$csvFiles = get-childitem $directory -filter *.csv

$resultsCSV = @();

Foreach ($csv in $csvFiles) {
    $resultsCSV += Import-Csv $csv
    $i++
    Write-Host "." -ForegroundColor Cyan
    Write-Progress -activity “Combining Information” -status “Status: $i/$($csvfiles.count)” -PercentComplete (($i / $csvFiles.count)*100)
        }

#List of Computers
$ComputerList = $Results | where {$_.Remoting -eq "InActive"} | select ComputerName


[int]$Throttle = 100
[int]$SleepTimer = 200
[int]$runspaceTimeout = 15
[int]$maxQueue = $(
                if($runspaceTimeout -ne 0){$Throttle}
                else{$throttle * 3}
                )

#Commands to be ran      
$ScriptBlock = {
    Param ($Computer)
psexec \\$computer /accepteula -s -n 10 c:\windows\system32\winrm.cmd quickconfig -quiet
Enter-pssession $computer
$RemoteObj = New-Object -TypeName PSobject

    $WinRMTest =  Get-Service winrm
    If ($WinRMTest.Status -eq "Running") {
    $WinRM = "Active"
    }
    ELSE {
    $WINRM = "Inactive"
         }

          [PSCustomObject]@{
                Computername = $Computer
                Remoting = $WinRM
                                    }  
 }         

#Clean up function        
Function Get-RunspaceData {
                param( [switch]$Wait,$Results )

                #loop through runspaces
                #if $wait is specified, keep looping until all complete
                Do {

                    #set more to false for tracking completion
                    $more = $false

                    #Progress bar if we have inputobject count (bound parameter)
                    Write-Progress  -Activity "Running Query"`
                        -Status "Starting threads"`
                        -CurrentOperation "$startedCount threads defined - $totalCount input objects - $script:completedCount input objects processed"`
                        -PercentComplete ($script:completedCount / $totalCount * 100)

                    #run through each runspace.           
 Foreach($runspace in $runspaces) {
                    
                        #get the duration - inaccurate
                        $currentdate = get-date
                        $runtime = $currentdate - $runspace.startTime
                        $runMin = [math]::round( $runtime.totalminutes ,2 )

                        #If runspace completed, end invoke, dispose, recycle, counter++
                        If ($runspace.Runspace.isCompleted) {
                          
                            $script:completedCount++
                            #everything is logged, clean up the runspace
                            $runspace.powershell.EndInvoke($runspace.Runspace) | Export-CSV "C:\Windows\Temp\RunspaceDump.CSV" -Delimiter ";" -Append
                            $runspace.powershell.dispose()
                            $runspace.Runspace = $null
                            $runspace.powershell = $null

                        }

                        #If runtime exceeds max, dispose the runspace
                        ElseIf ( $runspaceTimeout -ne 0 -and $runtime.totalseconds -gt $runspaceTimeout) {
                            
                            $script:completedCount++
                            
                            #Depending on how it hangs, we could still get stuck here as dispose calls a synchronous method on the powershell instance
                            $runspace.powershell.dispose()
                            $runspace.Runspace = $null
                            $runspace.powershell = $null
                            $completedCount++

                        }
                   
                        #If runspace isn't null set more to true  
                        ElseIf ($runspace.Runspace -ne $null ) {
                            $more = $true
                        }
                    }

                    #Clean out unused runspace jobs
                    $temphash = $runspaces.clone()
                    $temphash | Where { $_.runspace -eq $Null } | ForEach {
                        $Runspaces.remove($_)
                    }

                    #sleep for a bit if we will loop again
                    if($PSBoundParameters['Wait']){ start-sleep -milliseconds $SleepTimer }

                #Loop again only if -wait parameter and there are more runspaces to process
                } while ($more -and $PSBoundParameters['Wait'])
         }
            


#Create runspace pool with specified throttle
Write-Verbose "Creating runspace pool and session states"
$sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
$runspacepool = [runspacefactory]::CreateRunspacePool(1, 100, $sessionstate, $Host)
$runspacepool.Open() 

Write-Verbose "Creating empty collection to hold runspace jobs"
$Script:runspaces = New-Object System.Collections.ArrayList        
        
#If inputObject is bound get a total count and set bound to true
$global:__bound = $false
$allObjects = @()
if( $PSBoundParameters.ContainsKey("inputObject") ){
                $global:__bound = $true
}
#add piped objects to all objects or set all objects to bound input object parameter
if( -not $global:__bound ){
            $allObjects += $ComputerList
}
else{
            $allObjects = $ComputerList
}
      


        
#counts for progress
$totalCount = $allObjects.count
$script:completedCount = 0
$startedCount = 0

foreach($object in $allObjects){
        
#region add scripts to runspace pool
                
#Create the powershell instance and supply the scriptblock with the other parameters
$powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($object)
    
#Add the runspace into the powershell instance
$powershell.RunspacePool = $runspacepool
    
#Create a temporary collection for each runspace
$temp = "" | Select-Object PowerShell, StartTime, object, Runspace
$temp.PowerShell = $powershell
$temp.StartTime = get-date
$temp.object = $object
    
#Save the handle output when calling BeginInvoke() that will be used later to end the runspace
$temp.Runspace = $powershell.BeginInvoke()
$startedCount++

#Add the temp tracking info to $runspaces collection
Write-Verbose ( "Adding {0} to collection at {1}" -f $temp.object, $temp.starttime.tostring() )
$runspaces.Add($temp) | Out-Null
            
#loop through existing runspaces one time
Get-RunspaceData

#If we have more running than max queue (used to control timeout accuracy)
$firstRun = $true
while ($runspaces.count -ge $maxQueue) {

#give verbose output
if($firstRun){
Write-Verbose "$($runspaces.count) items running - exceeded $maxQueue limit."
}
$firstRun = $false
                    
#run get-runspace data and sleep for a short while
Get-RunspaceData
Start-Sleep -milliseconds $sleepTimer
}
#endregion add scripts to runspace pool
 }
                     
Write-Verbose ( "Finish processing the remaining runspace jobs: {0}" -f (@(($runspaces | Where {$_.Runspace -ne $Null}).Count)) )
Get-RunspaceData -wait

Write-Verbose "Closing the runspace pool"
$runspacepool.close()    
$Results = Import-csv "C:\Windows\Temp\RunspaceDump.CSV" -Delimiter ";"
Remove-Item "C:\Windows\Temp\RunspaceDump.CSV" -Force

#Results are in this variable
$Results |OGV