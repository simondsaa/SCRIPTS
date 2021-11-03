$ErrorActionPreference = 'SilentlyContinue'
$ProgramCheck = $null
$task = "Pending Install" 
$OSInfo = Get-WmiObject Win32_OperatingSystem 

IF($OSInfo.OSArchitecture -eq "64-bit")
{
    $InstallCheck = Test-Path "C:\Program Files (x86)\Avaya Aura AS 5300 UC Client"
        IF ($InstallCheck) {
            $Task = "Already Installed"
            }
    }

IF($OSInfo.OSArchitecture -eq "32-bit")
{
    $InstallCheck = Test-Path "C:\Program Files\Avaya Aura AS 5300 UC Client"
        IF ($InstallCheck) {
            $Task = "Already Installed"
            }
    }



IF($InstallCheck -eq $false){
    $Task = "Installing Program"
    Start-Process "\\xlwu-fs-05pv\Tyndall_PUBLIC\ncc admin\Avaya\Avaya_8.1.msi" /qn -wait
    Copy-item "\\xlwu-fs-04pv\Tyndall_325_msg\325 CS\SCO\SCOO\Scripts\Crill\SCHDTASK targets\Avaya UC.lnk" "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Startup\" -Force
    Copy-item "\\xlwu-fs-04pv\Tyndall_325_msg\325 CS\SCO\SCOO\Scripts\Crill\SCHDTASK targets\Avaya UC.lnk" "C:\Users\Public\Desktop" -Force

}


IF ($OSInfo.OSArchitecture -eq "64-bit")
    {IF (Test-Path "C:\Program Files (x86)\Avaya Aura AS 5300 UC Client")
        {$ProgramCheck = "SUCCESS"}
        ELSE { $ProgramCheck = "FAILED"}
        
    }
IF ($OSInfo.OSArchitecture -eq "32-bit")
    {IF (Test-Path "C:\Program Files\Avaya Aura AS 5300 UC Client")
        {$ProgramCheck = "SUCCESS"}
        ELSE { $ProgramCheck = "FAILED"}
        
    }

"$env:ComputerName : $Task : $($OSInfo.OSArchitecture) : $ProgramCheck : $((get-date).DateTime)" | out-file -append "\\xlwu-fs-05pv\Tyndall_PUBLIC\ncc admin\Avaya\Avaya_Tests.txt"
