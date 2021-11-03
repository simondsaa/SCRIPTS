$Comp = "Tynmoswk10zw903"
$Task = schtasks.exe /CREATE /TN "Enable youTube in Chrome" /S $Comp /SC ONLOGON /RL HIGHEST /RU SYSTEM /TR "powershell.exe -ExecutionPolicy Unrestricted -WindowStyle Hidden -noprofile -File 'C:\Users\1393356126A\Documents\Enable YouTube in Chrome.ps1'" /F
$Run = schtasks.exe /RUN /TN "Enable youTube in Chrome" /S $Comp