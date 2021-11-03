$App = Get-WmiObject Win32_Product | Where {$_.Name -like "*ActiveClient*"}
$App.Uninstall()