$computers=Get-Content -Path C:\Temp\g2.txt
foreach ($computer in $computers) {
(Get-WmiObject -computername $Computer -Namespace root/hp/instrumentedBIOS -Class HP_BIOSSettingInterface).SetBIOSSetting('Setup Password','<utf-16/>','<utf-16/>WEAp0ns1')
}