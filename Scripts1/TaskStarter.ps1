$Comp = "XLWUW-491S33"
$Task = schtasks.exe /CREATE /TN "JavaT" /S $Comp /SC MINUTE /mo 30 /RU INTERACTIVE /TR "powershell.exe -file 'C:\Temp\TaskStarter.ps1'" /F
$Run = schtasks.exe /RUN /TN "JavaT" /S $Comp
Sleep -Seconds 5