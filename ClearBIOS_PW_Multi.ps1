$Path = "C:\Temp\Ballentine.txt"
$Computers = gc $Path
ForEach ($Computer in $Computers)
{
(Get-WmiObject -computername $Computer -Namespace root/hp/instrumentedBIOS -Class HP_BIOSSettingInterface).SetBIOSSetting('Setup Password','<utf-16/>','<utf-16/>password')
}