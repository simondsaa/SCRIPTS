$Computers = Get-Content "C:\work\bowling.txt"
$Message = Read-Host "Message"
ForEach($Computer in $Computers){
    Msg Console /Server:$Computer $Message}$Message =
"Would you like to go bowling? Click yes or no."
If ($b -eq 6){Write-Host "You clicked Yes"}
ElseIf ($b -eq 7){Write-Host "You clicked No"}