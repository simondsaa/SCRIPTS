$App = Get-WmiObject Win32_Product | Where {$_.Name -like "*Java*"}
$App.Uninstall()