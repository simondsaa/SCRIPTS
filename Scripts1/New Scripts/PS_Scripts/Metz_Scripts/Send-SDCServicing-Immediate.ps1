Param($computer)
[xml]$task = gc "C:\Operation_UPGRADE\5.2-SDC-Servicing.xml"
$task.task.triggers.timetrigger.startboundary = [string]((Get-Date).Addminutes(2) | Get-Date -format "yyyy-MM-ddTHH:mm:00")
$task.save("\\$computer\c$\windows\temp\5.2-SDC-Servicing.xml")
Schtasks.exe /S "$computer" /Delete /TN "SDC Upgrade to 5.3.1" /XML "\\$computer\c$\windows\temp\5.2-SDC-Servicing.xml"
