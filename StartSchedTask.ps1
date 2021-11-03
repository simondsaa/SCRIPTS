$Path = Read-Host "Path to PCs"
$servers = get-content $Path
foreach ($server in $servers)
{
  Start-ScheduledTask -CimSession $server -TaskName "GoogleEarthPro_Install"
}