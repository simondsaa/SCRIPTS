Param($computer)
$sourcefiles = "\\xlwu-fs-dfs1v\Tyndall\SDC_531_Upgrade\Upgrade-Staging.ps1"
$destination = "\\$computer\c$\windows\temp\"

If (!(Test-Path -path $destination))
            {                        
                New-Item $destination -Type Directory -Force
            }            
                Copy-Item -Path $sourcefiles -Destination $destination

[xml]$task = gc "C:\Operation_UPGRADE\5.2-SDC-Servicing.xml"
$task.task.triggers.timetrigger.startboundary = [string]((Get-Date).AddMinutes(1) | Get-Date -format "yyyy-MM-ddTHH:mm:00")
$task.save("\\$computer\c$\windows\temp\5.2-SDC-Servicing.xml")
Schtasks.exe /S "$computer" /Create /TN "SDC-Servicing to 5.3.1" /XML "\\$computer\c$\windows\temp\5.2-SDC-Servicing.xml"
