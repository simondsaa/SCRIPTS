cls
$computers = gc "C:\Users\1274873341C\Desktop\Desktop\PS_Scripts\Reg_Fix\targets.txt"
foreach ($computer in $computers)
{
if (Test-Connection -count 1 -computer $computer -quiet){
Write-Host "Updating system" $computer "....." -ForegroundColor Green
Set-Service –Name RemoteRegistry –Computer $computer -StartupType Automatic
Get-Service -Name RemoteRegistry -Computer $computer | start-service
}
else
{
Write-Host "System Offline " $computer "....." -ForegroundColor Red
echo $computer >> "C:\Users\1274873341C\Desktop\Desktop\PS_Scripts\Reg_Fix\offlineRemoteRegStartup.txt"}
}