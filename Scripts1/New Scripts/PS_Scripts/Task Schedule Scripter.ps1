﻿$Comp = "XLWUW3-DKPVV1"
$Task = schtasks.exe /CREATE /TN "Scripter" /S $Comp /SC ONLOGON /RL HIGHEST /RU INTERACTIVE /TR "PowerShell.exe -ExecutionPolicy Unrestricted -WindowStyle Hidden -noprofile 'Start-Process powershell.exe -WindowStyle Hidden -ArgumentList ' -file C:\Users\1392134782A\Documents\Scripter.ps1' -verb RunAs'" /F