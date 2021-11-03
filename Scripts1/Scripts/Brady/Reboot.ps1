#$Computers = Get-Content C:\work\141_Computers.txt
#Set-Alias .\psshutdown.exe C:\Users\1180219788A\Documents\WindowsPowerShell\pstools\psshutdown.exe
$Computers = Read-Host "Computer Name"

$Message = "Restartin' yur 'pewter, kid."  

Please click OK to acknowledge this message."

ForEach ($Computer in $Computers)

    { Shutdown /m \\$Computer /r /f /t 10 /c "$Message" }
