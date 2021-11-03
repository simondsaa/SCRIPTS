$computers=Get-Content -Path C:\Users\1252862141.adm\Desktop\Scripts1\BIOS_Enable_Audio.txt
foreach ($computer in $computers) {
$AudioDevice=Get-WmiObject -computername $computer -Namespace root/hp/instrumentedBIOS -Class HP_BIOSSettingInterface
$AudioDevice.SetBIOSSetting("Audio Device","Disable")
}