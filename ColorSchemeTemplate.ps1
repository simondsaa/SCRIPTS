    {
     Write-Host " "
     $PW = Write-Host "Press" -NoNewline 
           Write-Host " 1" -ForegroundColor Green -NoNewline
           Write-Host " to" -NoNewline
           Write-Host " CREATE A TASK" -ForegroundColor Green -NoNewline
           Write-Host " on" -NoNewline
           Write-Host " one" -ForegroundColor Green -NoNewline
           Write-Host " PC." -NoNewline
           Write-Host " Press" -NoNewline 
           Write-Host " 2" -ForegroundColor Cyan -NoNewline
           Write-Host " to" -NoNewline
           Write-Host " CREATE A TASK" -ForegroundColor Cyan -NoNewline
           Write-Host " on" -NoNewline
           Write-Host " multiple" -ForegroundColor Cyan -NoNewline
           Write-Host " PCs:  " -NoNewline
           $Ans = Read-Host
If ($Ans -eq 1){
     $CN = Write-Host "What is the" -NoNewline
           Write-Host " COMPUTER NAME" -ForegroundColor Green -NoNewline
           Write-Host " that needs a task started?:  " -NoNewline
           $Comp = Read-Host
     $TN = Write-Host "What is the" -NoNewline
           Write-Host " TASK NAME" -ForegroundColor Green -NoNewline
           Write-Host " ? (This must be" -NoNewline
           Write-Host " exact" -ForegroundColor Green -NoNewline
           Write-Host ")" -NoNewline
           $TaskName = Read-Host
           StartTask
}
If ($Ans -eq 2){
     $CN = Write-Host "What is the" -NoNewline
           Write-Host " DIRECTORY" -ForegroundColor Cyan -NoNewline
           Write-Host " to the" -NoNewline
           Write-Host " LIST OF COMPUTERS" -ForegroundColor Cyan -NoNewline
           Write-Host " needing a task started?:  " -NoNewline
           $Path = Read-Host
           $Computers = Get-Content $Path
     $TN = Write-Host "What is the" -NoNewline
           Write-Host " TASK NAME" -ForegroundColor Cyan -NoNewline
           Write-Host " ? (This must be" -NoNewline
           Write-Host " exact" -ForegroundColor Cyan -NoNewline
           Write-Host ")" -NoNewline
           $TaskName = Read-Host
foreach ($comp in $Computers){
Start-ScheduledTask -CimSession $Comp -TaskName "$TaskName"
}
}
}