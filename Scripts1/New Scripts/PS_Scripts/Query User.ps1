$Comp = Read-Host "Name"
$User = Get-WmiObject Win32_ComputerSystem -comp $Comp -ErrorAction SilentlyContinue
Write-Host
Write-Host "User:    " $User.UserName