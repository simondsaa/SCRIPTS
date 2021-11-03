$computers=Get-Content -Path C:\Users\1252862141.adm\Desktop\Scripts\BIOS_Enable_Audio.txt
foreach ($computer in $computers) {
$passChange=Get-WmiObject -computername $computer -Namespace root/hp/instrumentedBIOS -Class HP_BIOSSettingInterface
$passChange.SetBIOSSetting('Setup Password','<utf-16/>','<utf-16/>WEAp0ns1')
}