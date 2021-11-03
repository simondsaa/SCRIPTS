$Comp = "XLWUL-42093D"
$Task = schtasks.exe /CREATE /TN "Minimize" /S $Comp /SC ONLOGON /RL HIGHEST /RU SYSTEM /TR "powershell.exe -noprofile -File '\\xlwu-fs-01pv\Tyndall_ANG\Shared\Minimize.ps1'" /F  
$Run = schtasks.exe /RUN /TN "Minimize" /S $Comp 
Sleep -Seconds 5 
#$Delete = schtasks.exe /DELETE /TN "Fix" /s  $Comp /F