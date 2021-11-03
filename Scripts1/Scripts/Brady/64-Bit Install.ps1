$Comp = "52XLWUW3-3818Y8"
$Task = schtasks.exe /CREATE /TN "64Bit Test" /S $Comp /SC ONLOGON /RL HIGHEST /RU SYSTEM /TR "powershell.exe -noprofile -File '\\xlwu-fs-05pv\Tyndall_PUBLIC\NCC Admin\64Bit_Install.ps1'" /F
$Run = schtasks.exe /RUN /TN "64Bit Test" /S $Comp