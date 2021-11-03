#require -version 2.0       
# create a pool of 3 runspaces    
$pool = [runspacefactory]::CreateRunspacePool(1, 3)    
$pool.Open()    
   
write-host "Available Runspaces: $($pool.GetAvailableRunspaces())"   
  
$jobs = @()    
$ps = @()    
.$wait = @()    
   
# run 6 background pipelines    
for ($i = 0; $i -lt 6; $i++) {    
       
  # create a "powershell pipeline runner"    
 $ps += [powershell]::create()    
      
 # assign our pool of 3 runspaces to use    
  $ps[$i].runspacepool = $pool   
       
  $freq = 840 + ($i * 10)    
  $sleep = (1 * ($i + 1))    
       
   # test command: beep and wait a certain time    
  [void]$ps[$i].AddScript(    
       "[console]::Beep($freq, 30); sleep -seconds $sleep")    
       
   # start job    
   write-host "Job $i will run for $sleep second(s)"   
  $jobs += $ps[$i].BeginInvoke();    
      
   write-host "Available runspaces: $($pool.GetAvailableRunspaces())"   
       
   # store wait handles for WaitForAll call    
  $wait += $jobs[$i].AsyncWaitHandle    
}    
   
# wait 20 seconds for all jobs to complete, else abort    
$success = [System.Threading.WaitHandle]::WaitAll($wait, 20000)    
   
write-host "All completed? $success"   
   
# end async call    
for ($i = 0; $i -lt 6; $i++) {    
 
   write-host "Completing async pipeline job $i"   
   
    try {    
   
       # complete async job    
        $ps[$i].EndInvoke($jobs[$i])    
   
   } catch {    
        
        # oops-ee!    
        write-warning "error: $_"   
    }    
   
    # dump info about completed pipelines    
    $info = $ps[$i].InvocationStateInfo    
   
    write-host "State: $($info.state) ; Reason: $($info.reason)"   
}    
   
# should show 3 again.    
write-host "Available runspaces: $($pool.GetAvailableRunspaces())"  
