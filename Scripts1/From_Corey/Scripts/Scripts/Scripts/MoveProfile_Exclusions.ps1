Write-Host -ForegroundColor White -BackgroundColor Red "NOTE:  Run 'WinRm qc' and restart before running this on a newly PXE'd PC"
$comp = Read-Host "PC w/ profile" 
$source = Read-Host "EDI + designator (N, A, E, C, V). Ex:  1325127891N"   
$pc = Read-Host "PC needing profile (will be in Temp folder)" 
$date = get-date -format M.d.yyyy 
$logDir = "C:Temp\Logs\ProfileMoves"
$logName = $comp+"_"+$source+"_"+$date

New-Item -Type directory -Path $logDir -Force  

ROBOCOPY \\$comp\C$\Users\$source \\$pc\C$\Temp\$source\ /E /Mir /B /R:3 /W:1 /XD "\\$comp\C$\Users\$source\Local Settings" /XD "\\$comp\C$\Users\$source\Contacts" /XD "\\$comp\C$\Users\$source\Links" /XD "\\$comp\C$\Users\$source\Application Data" /XD "\\$comp\C$\Users\$source\mcafee dlp quarantined files" /XD "\\$comp\C$\Users\$source\saved games" /XD "\\$comp\C$\Users\$source\searches" /XD "\\$comp\C$\Users\$source\AppData" /XD "\\$comp\C$\Users\$source\NetHood" /XD "\\$comp\C$\Users\$source\PrintHood" /XD "\\$comp\C$\Users\$source\SendTo" /XD "\\$comp\C$\Users\$source\Start Menu" /XD "\\$comp\C$\Users\$source\Templates" /XD "\\$comp\C$\Users\$source\Recent Items" /XF "\\$comp\C$\Users\$Source\NTUSER.dat" /log:"C:\Temp\Logs\ProfileMoves\$logname.txt"
