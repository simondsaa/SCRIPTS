$Path = Read-Host "Path to PCs"
$servers = get-content $Path
foreach ($server in $servers)
{
  Unregister-ScheduledTask -CimSession $server -TaskName "MSOffice16_Install" -Confirm:$false
}