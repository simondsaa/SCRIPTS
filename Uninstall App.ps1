$App = Get-WmiObject Win32_Product | Where {$_.Name -like "*Adobe AIR*"}
$App.Uninstall()