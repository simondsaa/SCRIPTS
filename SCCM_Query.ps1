#Server list 
$Servers = "xlwuw-759072"
 
#Define array
$Updates = @()
  
#Checking updates in Software Center
Try{
    $Updates = Invoke-Command -cn $Servers {
        $Application =  Get-WmiObject -Namespace "root\ccm\softwareupdates" -Class CCM_RemoteControlManager

        If(!$Application){
                $Object = New-Object PSObject -Property ([ordered]@{      
                        ArticleId         = " - "
                        Publisher         = " - "
                        Software          = " - "
                        Description       = " - "
                        State             = " - "
                        StartTime         = " - "
                        DeadLine          = " - "
                })
  
                $Object
        }
        Else{
            Foreach ($App in $Application){
  
                $EvState = Switch ( $App.EvaluationState  ) {
                        '0'  { "None" } 
                        '1'  { "Available" } 
                        '2'  { "Submitted" } 
                        '3'  { "Detecting" } 
                        '4'  { "PreDownload" } 
                        '5'  { "Downloading" } 
                        '6'  { "WaitInstall" } 
                        '7'  { "Installing" } 
                        '8'  { "PendingSoftReboot" } 
                        '9'  { "PendingHardReboot" } 
                        '10' { "WaitReboot" } 
                        '11' { "Verifying" } 
                        '12' { "InstallComplete" } 
                        '13' { "Error" } 
                        '14' { "WaitServiceWindow" } 
                        '15' { "WaitUserLogon" } 
                        '16' { "WaitUserLogoff" } 
                        '17' { "WaitJobUserLogon" } 
                        '18' { "WaitUserReconnect" } 
                        '19' { "PendingUserLogoff" } 
                        '20' { "PendingUpdate" } 
                        '21' { "WaitingRetry" } 
                        '22' { "WaitPresModeOff" } 
                        '23' { "WaitForOrchestration" } 
  
  
                        DEFAULT { "Unknown" }
                }
  
                $Object = New-Object PSObject -Property ([ordered]@{      
                        ArticleId         = $App.ArticleID
                        Publisher         = $App.Publisher
                        Software          = $App.Name
                        Description       = $App.Description
                        State             = $EvState
                         
                })
  
                $Object
            }
        }
  
    } -ErrorAction Stop | select @{n='ServerName';e={$_.pscomputername}},ArticleID,Publisher,Software,Description,State,StartTime,DeadLine
}
Catch [System.Exception]{
    Write-Host "Error" -BackgroundColor Red -ForegroundColor Yellow
    $_.Exception.Message
}
  
#Display results
$Updates | Out-GridView -Title "Updates"
 
#Export results to CSV
$Updates | Export-Csv $env:USERPROFILE\desktop\updates.csv -Force -NoTypeInformation