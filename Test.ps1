#============
Function MAC
{
$Path = Read-Host "PC List"
$Computers = Get-Content $Path
foreach ($comp in $Computers){
getmac /S $Comp}
}
#============
Function PingSweep
{
$Path = Read-Host "PC"
$Computername = Get-Content $Path
ping $Computername
}
#============
Do
{
Write-Host ""
Write-Host "1 - Mac"
Write-Host "2 - Ping"
Write-Host "3 - Quit"
$Ans = Read-Host "Pick one"
If ($Ans -eq 1)
{
  MAC
}
If ($Ans -eq 2)
{
  PingSweep
}
}
Until ($Ans -eq 3)
Cls
