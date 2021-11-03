$comp = Read-Host "PC w/ profile" 
$source = Read-Host "EDI + designator (N, A, E, C, V)"   
$pc = Read-Host "PC needing profile (will be in Temp folder)"
$computerName = $env:computerName 
$date = get-date -format M.d.yyyy 
$logDir = "C:Temp\Logs\ProfileMoves"
$logName = $comp+"_"+$source+"_"+$date
 
New-Item -Type directory -Path $logDir -Force  
 
ROBOCOPY \\$comp\C$\Users\$source \\$pc\C$\Temp\$source\ /E /B /MT:8 /NP /R:3 /W:1 /log:"C:\Temp\Logs\ProfileMoves\$logname.txt"
