$Path = "C:\Temp\Reboot.txt"
$Computers = Get-Content $Path
$Message = "Reboot will commence in 15 minutes. Please save all data! This will fix the current, base-wide logon issues. -101ACOMS"
ForEach ($Computer in $Computers)
    { Shutdown /m \\$Computer /r /f /t 900 /c "$Message" }