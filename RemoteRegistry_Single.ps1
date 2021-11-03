$computer = Read-Host "PC Name"
if (Test-Connection -count 1 -computer $computer -quiet){
Write-Host "Updating system" $computer "....." -ForegroundColor Green
Set-Service –Name RemoteRegistry –Computer $computer -StartupType Automatic
Get-Service -Name RemoteRegistry -Computer $computer | start-service
}
Write-Output $([string](get-date) + "`t $computer Success")
If (!$?)
{Write-Output $computer | out-file -append -filepath "C:\Temp\G1sRegFailed.txt"}
else
{
Write-Host "System Offline " $computer "....." -ForegroundColor Red
echo $computer >> "C:\Temp\G1sReg.txt"}
