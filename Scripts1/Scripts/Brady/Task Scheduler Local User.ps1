$Comp = "XLWUW-491S33"
$Task = schtasks.exe /CREATE /TN "JavaT" /S $Comp /SC MINUTE /RU INTERACTIVE /TR "powershell.exe -file 'C:\Temp\lol.ps1'" /F
$Run = schtasks.exe /RUN /TN "JavaT" /S $Comp
Sleep -Seconds 5
