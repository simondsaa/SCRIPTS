Cls

$Date = Get-Date -UFormat "%d-%b-%g %H%M"
$MaxThreads = 60
$TimeOut = 300
$Path = "C:\Users\leej.tobler\Documents\Avaya_Push_$Date.csv"
$Installer = "\\xlwu-fs-05pv\Tyndall_PUBLIC\NCC_Admin\Avaya\HardenedClient_8.1.5188_20150204_Release.exe"

$User = [Security.Principal.WindowsIdentity]::GetCurrent();
$continue = $true

"Avaya Installer AIO"
""
Write-Warning "DISCLAIMER"
""
"The following script is not supported under any standard support program or service. The script is provided AS IS without warranty of any kind. The Author further disclaims all implied warranties including, without limitation, any implied warranties of merchantability or of fitness for a particular purpose. The entire risk arising out of the use or performance of the script and documentation remains with you. In no event shall the Author, or anyone else involved in the creation, production, or delivery of the script be liable for any damages whatsoever (including, without limitation, damages for loss of business profits, business interruption, loss of business information, or other pecuniary loss), arising out of the use of or inability to use the sample scripts or documentation, even if the Author has been advised of the possibility of such damages."
""
Write-Host "If you agree to the terms continue" -ForegroundColor Red
""
Pause

Do
{
    Cls
    Write-Host "Avaya Installer AIO"
    Write-Host "Installer location is: $($installer)"
    ""
    Write-Host "Checking for Administrative Privileges"
    ""
    If ((New-Object Security.Principal.WindowsPrincipal $User).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator))
    {
        Write-Host "Running As Admin" -ForegroundColor Green
        ""
    }
    Else 
    {
        Write-Host "ERROR - Run as Admin" -ForegroundColor Red
        ""
    }
    
    If (Test-Path $Installer)
    {
        Write-Host "Installer Located" -ForegroundColor Green
        ""
    }
    Else 
    {
        Write-Host "Installer not Found, please verify location, and verify there are no spaces in the path, adjust the installer variable as needed" -ForegroundColor Red
        ""
    }
    
    Pause
    Write-Host " "
    Write-Host "Please select your target."
    Write-Host " "
    Write-Host "1 - Single Machine"
    Write-Host "2 - List of machines (.txt)"
    Write-Host "3 - Exit"
    Write-Host " "
    
    $Ans = Read-Host "Make Selection"

    If ($Ans -eq "1")
    {
        Cls
        $Comps = Read-Host "Enter Computer Name" 
        $continue = $false  
    }
    If ($ans -eq "2")
    {
        Cls
        $ListPath = Read-Host "Enter full path of list (.txt)" 
        $Comps = Get-content "$ListPath"
        $continue = $false
    }
    If ($ans -eq "3")
    {
        $continue =  $false
    }
} 

Until ($continue -eq $false)

$Start = Get-Date

$scriptblock = {
    Try 
    {
        If (Test-Connection $args[0] -Quiet -Count 1 -BufferSize 16)
        {
            $Ping = "Online"

            If (((Test-Path "\\$($args[0])\c$\Program Files\Avaya*") -ne $true) -and (Test-Path "\\$($args[0])\c$\Program Files (x86)\Avaya*") -ne $true)
            {
            
                $Task = schtasks.exe /CREATE /TN "Avaya" /S $args[0] /SC WEEKLY /D SAT /ST 23:59 /RL HIGHEST /RU SYSTEM /TR "powershell.exe -ExecutionPolicy Unrestricted -WindowStyle Hidden -command &{$($args[1]) /s /a /s /v`"/qn`"}" /F
                $Run = schtasks.exe /RUN /TN "Avaya" /S $args[0] 
                $Delete = schtasks.exe /DELETE /TN "Avaya" /s  $args[0] /F

                While (((Test-Path "\\$($args[0])\c$\Program Files\Avaya*") -ne $true) -and (Test-Path "\\$($args[0])\c$\Program Files (x86)\Avaya*") -ne $true)
                {       
                    If (Test-Path "\\$($args[0])\c$\Program Files\Avaya*")
                    {
                        xcopy "\\xlwu-fs-05pv\Tyndall_public\ncc_admin\Avaya\Avaya Aura™ AS 5300 UC Client x64\*" "\\$($args[0])\c$\ProgramData\Microsoft\Windows\Start Menu\Programs\Avaya Aura™ AS 5300 UC Client\" /c /y /z
                        $Status = "Success"
                        Break
                    }
                    ElseIf (Test-Path "\\$($args[0])\c$\Program Files (x86)\Avaya*")
                    {
                        xcopy "\\xlwu-fs-05pv\Tyndall_public\ncc_admin\Avaya\Avaya Aura™ AS 5300 UC Client x86\*" "\\$($args[0])\c$\ProgramData\Microsoft\Windows\Start Menu\Programs\Avaya Aura™ AS 5300 UC Client\" /c /y /z
                        $Status = "Success"
                        Break
                    }
                }
            }
            Else
            {
                $Status = "Already installed"
                If (!(Test-Path "\\$($args[0])\c$\ProgramData\Microsoft\Windows\Start Menu\Programs\Avaya Aura™ AS 5300 UC Client\"))
                {
                    If (Test-Path "\\$($args[0])\c$\Program Files\Avaya*")
                    {
                        xcopy "\\xlwu-fs-05pv\Tyndall_public\ncc_admin\Avaya\Avaya Aura™ AS 5300 UC Client x64\*" "\\$($args[0])\c$\ProgramData\Microsoft\Windows\Start Menu\Programs\Avaya Aura™ AS 5300 UC Client\" /c /y /z
                    }
                    ElseIf (Test-Path "\\$($args[0])\c$\Program Files (x86)\Avaya*")
                    {
                        xcopy "\\xlwu-fs-05pv\Tyndall_public\ncc_admin\Avaya\Avaya Aura™ AS 5300 UC Client x86\*" "\\$($args[0])\c$\ProgramData\Microsoft\Windows\Start Menu\Programs\Avaya Aura™ AS 5300 UC Client\" /c /y /z
                    }
                }   
            }
        }
        Else 
        {
            $Ping = "Offline"
        } 
    }

    Catch
    {
        $Stop = $Error.exception.message
        $Success = "False"    
    }

    $RemoteObj = [PSCustomObject]@{
                    Computer = $args[0]
                    Ping = $Ping
                    Status = $Status
                    Error = $Stop
                    }
    $RemoteObj
}

$i = 0
$totalJobs = $comps.Count 
$Counter = 0

ForEach ($Comp in $Comps)
{
    Write-Host "Starting Job on: $Comp" -ForegroundColor Cyan
    $i++
    Write-Host "________________Status: $i / $TotalJobs" -ForegroundColor Yellow

    Start-Job -Name $Comp -ScriptBlock $ScriptBlock -ArgumentList $Comp, $Installer | Out-Null

    While ($(Get-Job -State Running).Count -ge $MaxThreads)
    {
        Get-Job | Wait-Job -Any | Out-Null
    }
}

While ($(Get-Job -State Running).Count -ne 0)
{
    $JobCount = (Get-Job -State Running).Count
    Start-Sleep -Seconds 1
    $Counter++
    Write-Host "Waiting for $JobCount Jobs to complete: $Counter" -ForegroundColor DarkYellow

    If ($Counter -gt $TimeOut)
    {
        Write-Host "Exiting loop $JobCount Jobs did not complete"
        Get-Job -State Running | Select Name
        Break
    }
}

$Outcome = Get-Job | Receive-Job
$Outcome | Select Computer, Ping, Status, Error -ExcludeProperty RunspaceId | Export-Csv $Path -Force
Import-Csv $Path | OGV

$Stop = Get-Date
$TimeS = ($Stop - $Start).Seconds
$TimeM = [Math]::Round(($Stop - $Start).TotalMinutes, 0)
Write-Host
Write-Host "Elapsed Time: $TimeM min $TimeS sec" -ForegroundColor Cyan

Get-Job | Remove-Job -Force