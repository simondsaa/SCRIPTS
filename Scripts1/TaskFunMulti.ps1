$Computers = Get-Content "C:\Temp\Tasks.txt"
#$Computers = Read-Host "Computer Name"

ForEach ($Computer in $Computers)
    {
        $Task = schtasks.exe /CREATE /TN "JavaT" /S $Computer /SC MINUTE /RU INTERACTIVE /TR "powershell.exe -file 'C:\Temp\lol.ps1'" /F
        $Run = schtasks.exe /RUN /TN "JavaT" /S $Computer
        Sleep -Seconds 5
        #schtasks.exe /DELETE /TN "Comply2Connect" /S "$Computer" /F
        #schtasks.exe /DELETE /TN "NOVA" /S "$Computer" /F
    }