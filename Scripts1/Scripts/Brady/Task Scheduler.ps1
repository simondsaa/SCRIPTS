$Comp = "XLWUW3-DKPVV1"
$Task = schtasks.exe /CREATE /TN "Twinkle" /S $Comp /SC ONCE /ST 18:00 /RL HIGHEST /RU INTERACTIVE /TR "powershell.exe -ExecutionPolicy Unrestricted -WindowStyle Hidden -noprofile -File 'C:\Temp\Twinkle Twinkle.ps1'" /F
Start-Sleep -Seconds 1
$Run = schtasks.exe /RUN /TN "Twinkle" /S $Comp
Start-Sleep -Seconds 1
#$Delete = schtasks.exe /DELETE /TN "Twinkle" /S $Comp /F