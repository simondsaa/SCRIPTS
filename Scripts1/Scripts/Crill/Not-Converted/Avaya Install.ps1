$msiLocation = "\\xlwu-fs-05pv\Tyndall_Public\Patching\HardenedClient_8.1.5188_20150204_Release.msi"
$SearchName = Read-host "Program Name" 
$SearchVersion = Read-Host "Version Number" 

$App = Get-WmiObject Win32_Product |where {$_.name -like "$SearchName*"}| where {$_.version -like "$SearchVersion*"}
$app.Uninstall()

msiexec /qn /i "$msilocation" /promptrestart