#Sort all systems array ascending
#$Computername = $AllComputerNames | Sort-Object
$Computername = Get-Content "C:\Users\1394844760A\Desktop\Scripting Test Bed\Remediation.txt"
Measure-Command {

$scriptblock =
{
  Param ($computer)
  $GatherIAVM = Get-WmiObject -Namespace "Root\ccm\softwareupdates\updatesstore" -Class CCM_UpdateStatus -ComputerName $Computer
          [PSCustomObject]@{
            Server = $GatherIAVM.__SERVER
            Status = $GatherIAVM.Status
            Bulletin = $GatherIAVM.Bulletin
            Article = $GatherIAVM.Article
            Title = $GatherIAVM.Title
            UniqueID = $GatherIAVM.UniqueID
            ScanTme = $GatherIAVM.ScanTime
                                                }
}

$RunspacePool = [RunspaceFactory]::CreateRunspacePool(1,100)
$RunspacePool.Open()
$Jobs =
   foreach ( $computer in $Computername)
    {
     $Job = [powershell]::Create().AddScript($ScriptBlock).AddArgument($computer)
     $Job.RunspacePool = $RunspacePool

     [PSCustomObject]@{
      Pipe = $Job
      Result = $Job.BeginInvoke()
     }
}

Write-Host 'Working..'

Do {
   Write-host "Still Working"
   Start-Sleep -Seconds 1
} While ( $Jobs.Result.IsCompleted -contains $false)

Write-Host ' Done! Writing output file.'
Write-host "C:\Windows\Temp\PingResults.csv"
$RunspacePool.Close()

$IAVMResults =ForEach ($Job in $Jobs) { $Job.Pipe.EndInvoke($Job.Result) }
$RunspacePool.Close()
$RunspacePool.Dispose()
}






         $IAVMResults= $IAVMResults | Select __SERVER, Status, Bulletin, Article, Title, UniqueID, ScanTime | Sort-Object
         $IAVMInstalled= $IAVMResults | Select __SERVER, Status, Bulletin, Article, Title, UniqueID, ScanTime | Where {$_.Status -eq "Installed"}
         $IAVMMissing= $IAVMResults | Select __SERVER, Status, Bulletin, Article, Title, UniqueID, ScanTime | Where {$_.Status -eq "Missing"}
         $IAVM_MSBulletin = $IAVMResults | Select __SERVER, Status, Bulletin, Article, Title, UniqueID, ScanTime | Where {$_.Bulletin -like "MS*"}
         $IAVMMissing_MS = $IAVMResults | Select __SERVER, Status, Bulletin, Article, Title, UniqueID, ScanTime | Where {$_.Status -eq "Missing"} | Where {$_.Bulletin -like "MS*"}
         $IAVMInstalled_MS = $IAVMResults | Select __SERVER, Status, Bulletin, Article, Title, UniqueID, ScanTime | Where {$_.Status -eq "Installed"} | Where {$_.Bulletin -like "MS*"}

             $stoptimer = Get-Date
         Write-Host
         "Total Time for IAVM Enumeration: {0} Minutes" -f [math]::round(($stoptimer - $starttimer).TotalMinutes , 2)
         Write-Host
         "Total Systems: {0} " -f ($computername | Measure-Object).count
         Write-Host
         "Total Patches Enumerated  : {0} " -f ($IAVMResults | Measure-Object).count
         "Total Patches Installed   : {0} " -f ($IAVMInstalled | Measure-Object).count
         "Total Patches Missing     : {0} " -f ($IAVMMissing |Measure-Object).count
         Write-Host
         "Total Patched   : {0}%  " -f [math]::Round(($IAVMInstalled.count / $IAVMResults.count) * 100,2)
         "Total Unpatched : {0}%  " -f [math]::Round(($IAVMMissing.count / $IAVMResults.count) * 100,2)
         Write-Host
         "Total MS Bulletins Installed   : {0} " -f ($IAVMInstalled_MS | Measure-Object).count
         "Total MS Bulletins Missing     : {0} " -f ($IAVMMissing_MS |Measure-Object).count
         "MS Bulletins Patch Compliance  : {0}%  " -f [math]::Round(($IAVMInstalled_MS.count / $IAVM_MSBulletin.count) * 100,2)
          Return $IAVMResults | Out-GridView
