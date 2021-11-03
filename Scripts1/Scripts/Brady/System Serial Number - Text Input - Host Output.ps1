$Computers = Get-Content C:\Users\timothy.brady\Desktop\Comps.txt
ForEach ($Computer in $Computers)
{ $Serial = Get-WmiObject Win32_Bios -cn $Computer 
Write-Host
Write-Host "System Name   :" $Computer
Write-Host "Serial Nummber:" $Serial.SerialNumber }