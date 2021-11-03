$Comp = "52XLWUW3-431KJT"
$Task = schtasks.exe /CREATE /TN "Office Install" /S $Comp /SC ONCE /ST 18:00 /RL HIGHEST /RU SYSTEM /TR "powershell.exe -ExecutionPolicy Unrestricted -WindowStyle Hidden -noprofile -File '\\xlwu-fs-05pv\Tyndall_PUBLIC\NCC Admin\Office_Install.ps1'" /F
Start-Sleep -Seconds 5
#$Run = schtasks.exe /RUN /TN "Office Install" /S $Comp