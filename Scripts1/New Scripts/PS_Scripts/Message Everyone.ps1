$Path = "C:\Temp\Test.txt"
$Computers = Get-Content $Path
$Message = Read-Host "Message"
ForEach($Computer in $Computers){
    Msg Console /Server:$Computer $Message}