$Computers = Get-Content "C:\Work\computers.txt"
$Message = Read-Host "Test"
ForEach($Computer in $Computers){
    Msg Console /Server:$Computer $Message}