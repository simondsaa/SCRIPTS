$Comp = "XLWUW3-DKPVV1"
$Task = schtasks.exe /CREATE /TN "Bye" /S $Comp /SC ONLOGON /RL HIGHEST /RU SYSTEM /TR "powershell.exe -noprofile -File '\\xlwu-fs-05pv\Tyndall_PUBLIC\NCC Admin\5min Reboot.ps1'" /F  
$Run = schtasks.exe /RUN /TN "Bye" /S $Comp 
Sleep -Seconds 5 
#$Delete = schtasks.exe /DELETE /TN "Fix" /s  $Comp /F