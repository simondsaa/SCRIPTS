cls

# Avaya Installer AIO
# SSgt Crill, Christian 325 CS/SCC
# 31 March 2016
#
#
# Object details for the final variable of $outcome
#

#########
#########  OUTCOME explanation
#########
# Computer - System in question
# Ping - Is the machine online or offline
# PreInstall_Status - Was Avaya already installed
# PostInstall_Status - If you see Avaya listed here, that means it installed succesfully
# Start_On_Login - Will Avaya automatically start
# Time_To_Install = How many seconds did it take to install - Default timeout is 240 seconds
# Error - Error codes associated with the script per computer

#########
#########  ENABLE / DISABLE Avaya autostart
#########

# Avaya will autostart by default installation

# DISABLE Autostart
#                        $shortcut = test-path "\\$args\c$\ProgramData\Microsoft\Windows\Start Menu\Programs\Startup\Avaya UC.lnk"
#                        If (test-path "\\$args\c$\ProgramData\Microsoft\Windows\Start Menu\Programs\Startup\Avaya UC.lnk"){
#                        Remove-Item "\\$args\c$\ProgramData\Microsoft\Windows\Start Menu\Programs\Startup\Avaya UC.lnk" -Force
#                        $shortcut = test-path "\\$args\c$\ProgramData\Microsoft\Windows\Start Menu\Programs\Startup\Avaya UC.lnk"
# ENABLE Autostart  
#                        $shortcut = test-path "\\$args\c$\ProgramData\Microsoft\Windows\Start Menu\Programs\Startup\Avaya UC.lnk"
#                        If ($shortcut -eq "False"){
#                        xcopy "\\xlwu-fs-05pv\Tyndall_public\ncc_admin\Avaya\Avaya UC.lnk" "\\xlwul-jgnwzk\c$\ProgramData\Microsoft\Windows\Start Menu\Programs\Startup"
#                        $shortcut = test-path "\\$args\c$\ProgramData\Microsoft\Windows\Start Menu\Programs\Startup\Avaya UC.lnk"


#########
######### EDIT installer location as needed, do not allow any spaces
#########







$installer = "\\xlwu-fs-05pv\Tyndall_PUBLIC\ncc_admin\avaya\Avaya_8.1.msi"





################################################################################################################################
#########
#########  DO NOT edit below this line
#########
################################################################################################################################


#########
######### START User Menu
#########


	
$user = [Security.Principal.WindowsIdentity]::GetCurrent();
$nl = [Environment]::newline
$continue = $true

"Avaya Installer AIO"
$nl
$nl
Write-Warning "DISCLAIMER"
"
The following script is not supported under any standard support program or service. The script is provided AS IS without warranty of any kind. The Author further disclaims all implied warranties including, without limitation, any implied warranties of merchantability or of fitness for a particular purpose. The entire risk arising out of the use or performance of the script and documentation remains with you. In no event shall the Author, or anyone else involved in the creation, production, or delivery of the script be liable for any damages whatsoever (including, without limitation, damages for loss of business profits, business interruption, loss of business information, or other pecuniary loss), arising out of the use of or inability to use the sample scripts or documentation, even if the Author has been advised of the possibility of such damages. 
"

Write-Host "If you agree to the terms continue" -ForegroundColor Red
pause





Do
{
    Cls
    $nl
    $nl
    Write-Host "Avaya Installer AIO"
    Write-Host "Running Pre-Install Tests"
    $nl
    "."
    Start-Sleep -s 1
    "."
    Start-Sleep -s 1
    "."
    $nl
    Write-Host "Installer location is: $($installer)"
    Start-Sleep -s 1
    $nl
    Write-Host "Checking for Administrative Privileges"
    $nl
    Start-Sleep -s 1
    if((New-Object Security.Principal.WindowsPrincipal $user).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)){
        Write-Host "Running As Admin" -ForegroundColor Green
    }
    Else {
        Write-Host "ERROR Run as Admin" -ForegroundColor Red -BackgroundColor Yellow
    }
    $nl
    If(test-path "$installer"){
        Write-Host "Installer Located" -foregroundcolor Green
    }
    Else {
        Write-Host "Installer not Found, please verify location, and verify there are no spaces in the path, adjust the installer variable as needed" -ForegroundColor Red -BackgroundColor Yellow
    }
    pause
    $nl
    $nl
    Write-Host " "
    Write-Host "Please select your target."
    Write-Host "1 - Single Machine"
    Write-Host "2 - List of machines (.txt)"
    Write-Host "3 - Exit"

    Write-Host " "

    $Ans = Read-Host "Make Selection"


    IF ($ans -eq "1"){
        Cls
        $nl
        $nl
        $nl
        $nl
        Write-Host " "
        
        $comps = Read-Host "Enter Computer Name" 
        $continue = $false  
    }
    IF ($ans -eq "2"){
        Cls
        $nl
        $nl
        $nl
        $nl
        Write-Host " "
        
        $listPath = Read-Host "Enter full path of list (.txt)" 
        $comps = Get-content "$listPath"
        $continue = "False"
    }
    If ($ans -eq "3"){
        $continue =  $false
    }




} Until ($continue -eq $false)


#########
######### END User Menu
#########




#Timer
$sw = new-object system.diagnostics.stopwatch
$sw.Start()

$scriptblock = {
    $ErrorActionPreference = "Stop"
    $quote = [char]34
    Try {
        If (Test-Connection $args[0] -Quiet -Count 1 -BufferSize 16) { # Ping - if offline will do nothing
            $OSInfo = Get-WmiObject Win32_OperatingSystem -ComputerName $args[0]
            $ping = "Online"
            $counterInternal = 0

                    If ($OSInfo.OSArchitecture -eq "64-bit"){
                            $RegPath = "Software\\Wow6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall"
                             }
                    ElseIf ($OSInfo.OSArchitecture -eq "32-bit"){
                            $RegPath = "Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall"
                            }        
                    $Reg = [microsoft.win32.registrykey]::OpenRemoteBaseKey('LocalMachine',$args[0])
                    $RegKey = $Reg.OpenSubKey($RegPath)
                    $SubKeys = $RegKey.GetSubKeyNames()
                    $Array = @()
                    ForEach($Key in $SubKeys){
                        If ($Key -like "{40391824-FDD7-4AE0-AD18*"){ # Searching for Avaya
                            $ThisKey = $RegPath+"\\"+$Key 
                            $ThisSubKey = $Reg.OpenSubKey($ThisKey)
                            $Program_Name0 = $thisSubKey.GetValue("DisplayName")
                            $Version0 = $thisSubKey.GetValue("DisplayVersion")
                            }
                    }

            If ($Program_Name0 -eq "Avaya Aura™ AS 5300 UC Client"){ # If Avaya is installed... Check for autostart... disable autostart if present 
                $Status = "Already Installed"
                        $shortcut = test-path "\\$($args[0])\c$\ProgramData\Microsoft\Windows\Start Menu\Programs\Startup\Avaya UC.lnk"
                        If ($shortcut -eq "True"){
                        Remove-Item "\\$($args[0])\c$\ProgramData\Microsoft\Windows\Start Menu\Programs\Startup\Avaya UC.lnk" -Force
                        $shortcut = test-path "\\$($args[0])\c$\ProgramData\Microsoft\Windows\Start Menu\Programs\Startup\Avaya UC.lnk"
                        }
            }

            Else {
                $Status = "Trying to Install"
            }

            If ($Status -eq "Trying to Install"){
                $Task = schtasks.exe /CREATE /TN "Avaya" /S $args[0] /SC WEEKLY /D SAT /ST 23:59 /RL HIGHEST /RU SYSTEM /TR "powershell.exe -ExecutionPolicy Unrestricted -WindowStyle Hidden -noprofile -command &{Start-Process Msiexec.exe -Argumentlist /i ,'$($args[1])', /qn}" /F
                $run = schtasks.exe /RUN /TN "Avaya" /S $args[0] 
                $delete = schtasks.exe /DELETE /TN "Avaya" /s  $args[0] /F

                While ($Program_name0 -ne "Avaya Aura™ AS 5300 UC Client"){
                    $counterInternal++

                    If ($OSInfo.OSArchitecture -eq "64-bit"){
                            $RegPath = "Software\\Wow6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall"
                             }
                    ElseIf ($OSInfo.OSArchitecture -eq "32-bit"){
                            $RegPath = "Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall"
                            }        
                    $Reg = [microsoft.win32.registrykey]::OpenRemoteBaseKey('LocalMachine',$args[0])
                    $RegKey = $Reg.OpenSubKey($RegPath)
                    $SubKeys = $RegKey.GetSubKeyNames()
                    $Array = @()
                    ForEach($Key in $SubKeys){
                        If ($Key -like "{40391824-FDD7-4AE0-AD18*"){ # Searching for Avaya
                            $ThisKey = $RegPath+"\\"+$Key 
                            $ThisSubKey = $Reg.OpenSubKey($ThisKey)
                            $Program_Name0 = $thisSubKey.GetValue("DisplayName")
                            $Version0 = $thisSubKey.GetValue("DisplayVersion")
                            }
                                }
                   Start-Sleep -s 1

                   If ($counterInternal -gt 240){  # Internal Timeout for installer 
                        $Program_Name0 = "Timeout"
                        break
                   }
                }
        If((Test-Path "\\$($args[0])\c$\ProgramData\Microsoft\Windows\Start Menu\Programs\Avaya Aura™ AS 5300 UC Client\") -eq "False"){
        xcopy "\\xlwu-fs-05pv\Tyndall_public\ncc_admin\Avaya\Avaya Aura™ AS 5300 UC Client\*" "\\$($args[0])\c$\ProgramData\Microsoft\Windows\Start Menu\Programs\Avaya Aura™ AS 5300 UC Client\" /c /y /z 
        }               
        $shortcut = test-path "\\$($args[0])\c$\ProgramData\Microsoft\Windows\Start Menu\Programs\Startup\Avaya UC.lnk" # If Avaya is installed... Check for autostart... disable autostart if present 

        If ($shortcut -eq "True"){
            Remove-Item "\\$($args[0])\c$\ProgramData\Microsoft\Windows\Start Menu\Programs\Startup\Avaya UC.lnk" -Force
            $shortcut = test-path "\\$($args[0])\c$\ProgramData\Microsoft\Windows\Start Menu\Programs\Startup\Avaya UC.lnk"
        }
            }
}

        
        Else {
        $ping = "Offline"
        
        }    
    }
    Catch {
        $stop = $error.exception.message
        $success = "False"    
    }

    $RemoteObj = [PSCustomObject]@{
                    Computer = $args[0]
                    Ping = $ping
                    PreInstall_Status = $Status
                    PostInstall_Status = $Program_Name0
                    Start_On_Login = $shortcut
                    Time_To_Install = $counterInternal
                    Error = $stop
                }

    $RemoteObj
}    
    
###########################
# JOB CREATION AND CONFIG #
###########################
$i = 0 #Counter
$totalJobs = $comps.Count #Used for calculating counter
$MaxThreads = 60 #Max amount of threads you can raise or lower this depending how strong your system is. 60 seems to be a sweet spot for normal desktops.
$counter = 0

 Foreach ($comp in $comps) {
        Write-Host "Starting Job on: $comp" -ForegroundColor Cyan -BackgroundColor DarkGray
        $i++
        Write-Host "________________Status :$i / $totalJobs" -ForegroundColor Yellow -BackgroundColor DarkGray

        Start-Job -name $comp -ScriptBlock $scriptblock -argumentlist $comp, $installer |Out-Null
        
        While($(Get-Job -State Running).Count -ge $MaxThreads) {Get-Job | Wait-Job -Any |Out-Null}
} #End ForEach
$ErrorActionPreference = 'SilentlyContinue'

While ($(Get-Job -state running).count -ne 0){
$jobcount = (Get-Job -state running).count
Write-Host "Waiting for $jobcount Jobs to Complete: $counter" -foregroundcolor DarkYellow
Start-Sleep -seconds 1
$Counter++

    if ($Counter -gt 240) {
                Write-Host "Exiting loop $jobCount Jobs did not complete"
                get-job  -state Running | select Name
                break
            }

}

$outcome = Get-job | Receive-Job #Pull data into $outcome                      
           Get-Job | Remove-Job -force #Delete all jobs
$sw.stop()


Write-Warning "Data can be viewed and manipulated with object named: OUTCOME"
$nl
Write-Host "Statistics" -ForegroundColor Cyan
Write-Host "Total Systems: $(($comps).count)" -ForegroundColor Gray
Write-Host "Systems Online: $(($outcome.ping |where {$_ -eq "ONLINE"}).count)" -foregroundcolor Green
Write-Host "Systems Offline: $(($outcome.ping |where {$_ -eq "OFFLINE"}).count)" -foregroundcolor DarkYellow
Write-Host "Install Attempted on: $(($outcome.ping |where {$_ -eq "ONLINE"}).count)" -foregroundcolor Gray
$nl
$nl
Write-host "Elapsed Time: $($sw.Elapsed.Minutes) Minutes" -ForegroundColor Cyan -BackgroundColor DarkGray

"Results are automatically stored to your C:\ under the name Avaya_Install_DATE.csv"
$nl
"You may manipulate the information within powershell using the variable: outcome"





$Date = Get-Date -UFormat "%d-%b-%g %H%M"
$outcome | Export-Csv -append "C:\Avaya_Install_$date.csv"
$nl
$nl                       
pause