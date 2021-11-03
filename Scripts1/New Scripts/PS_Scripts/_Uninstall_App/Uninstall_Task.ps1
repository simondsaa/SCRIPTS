
#BE SURE AND MODIFY THE APP NAME BEFORE EXECUTING



$App = Get-WmiObject Win32_Product | Where {$_.Name -like "*WinZIP*"}
$App.Uninstall()