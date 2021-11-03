$Comp = "52XLWUW3-372CKY"
$Task = schtasks.exe /CREATE /TN "Deleted Old Profiles" /S $Comp /SC WEEKLY /D SAT /ST 23:59 /RU "SYSTEM" /TR "powershell.exe -ExecutionPolicy Unrestricted -WindowStyle Hidden -noprofile -File '\\xlwu-fs-05pv\Tyndall_PUBLIC\NCC Admin\Delete Old Profiles.ps1'" /F
#$Run = schtasks.exe /RUN /TN "Notification" /S $Comp
#Sleep -Seconds 5
#$Delete = schtasks.exe /DELETE /TN "Twinkle" /S  $Comp /F