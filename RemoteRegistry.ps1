$computers = get-content "C:\Temp\g1s.txt"
foreach ($computer in $computers)
{
if (Test-Connection -count 1 -computer $computer -quiet){
Write-Host "Updating system" $computer "....." -ForegroundColor Green
Set-Service –Name RemoteRegistry –Computer $computer -StartupType Automatic
Get-Service -Name RemoteRegistry -Computer $computer | start-service
}
If (!$?)
{Write-Output $computer | out-file -append -filepath "C:\Temp\G1sRegFailed.txt"}
else
{
Write-Host "System Offline " $computer "....." -ForegroundColor Red
echo $computer >> "C:\Temp\G1sReg.txt"}
}