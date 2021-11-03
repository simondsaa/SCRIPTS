#* FileName: Intelligent_Patcher.ps1
#*=============================================================================
#* Script Name: InteliPatch
#* Created: 09/11/2014
#* Author: SSgt Crill, Christian
#* Company: USAF
#* Email: christian.crill@us.af.mil
#* Web: 
#* Reqrmnts:
#* Keywords:
#*=============================================================================
#  WORKFLOW
# Ping LOGGED
# Manage Remote machine LOGGED
# If failed repair machine LOGGED
# Manage repaired Remote machine LOGGED
# 
#      NEED:  Patching script addition / combine logs into one CSV
#*=============================================================================
#*=============================================================================
#*                            REVISION HISTORY
#*=============================================================================
#* Date: 10/16/2014
#* Time: 12:00
#* Streamlined process, Added Minion-Enable-PSRemoting, Added Patch program
#* Planning to add Runspace Pool
#*
#*=============================================================================
#*=============================================================================
#*                           FUNCTION / VARIABLE LISTING
#*=============================================================================

#Function: $comps
#Names of remote machines
$comps = Get-Content "C:\Users\1394844760A\Desktop\Remediation\Remediation.txt" | Sort-Object

#Function: $updatedir
#Specify the location of the *.msu files
$updatedir = "\\xlwu-fs-002\tyndall$\Applications\MSBatch\Crill Patch"

#Function: $Domain
#Fully Qualified Domain Name (FQDN)
$Domain = ".area52.afnoapps.usaf.mil"


#=============================================================================
#                                  MENU
#=============================================================================
    CLS
    Write-Host "InteliPatch, Created by SSgt Crill, Christian 325 CS/SCOO" -ForegroundColor Yellow
    Write-Host "Program will Ping machines, enable PSRemoting if needed, then push patches to the machines" -For Green
    Write-Host
    Write-Host "SCRIPT REQUIREMENTS will be in red" -BackgroundColor Yellow -Foregroundcolor Red
    Write-Host "Run as Administrator" -Foregroundcolor Red 
    Write-Host "PSExec on local machine" -ForegroundColor Red

    Write-host "Please Verify your target list" -BackgroundColor Yellow -Foregroundcolor Red
    $comps
    pause

#=============================================================================
#                               SCRIPT BODY
#=============================================================================


$Results = @()

ForEach ($comp in $comps)
{

$Ping = Test-Connection $comp -Quiet -Buffersize 16 -Ea 0 -Count 1
Write-Host "Pinging $($comp)" -Foregroundcolor darkyellow
if ($Ping) {
    $WinMRM = psexec \\$comp -s -n 10 c:\windows\system32\winrm.cmd quickconfig -quiet
    $Outcome = New-Object PSObject -Property @{
    Computer = $comp
    Result   = $Ping
    Enable = $WinRM
     }
    }
$Results += $Outcome
}





Write-Host "Total systems Online: $(($Results.Result -eq $True).count)" -ForegroundColor Green
Write-Host "Total systems Offline: $(($Results.Result -eq $False).count)" -ForegroundColor Red

Start-Sleep -milliseconds 100
$Results | ogv



Start-Sleep -Milliseconds 100
$Results | ogv 
$Repaired = ($Results | where {$_.Result -eq $True}).computer


ForEach ($S_comp in $Repaired){
robocopy /E /Z /R:0 "C:\Users\1394844760A\Desktop\Remediation\Patches" "\\$S_comp\c$\Patching"
Invoke-Command -ComputerName ($S_comp) -Scriptblock {
$files = Get-ChildItem "C:\Patching" -Recurse
    foreach ($file in $files)
        {
                Write-Host "Installing update $file ..."
                $fullname = $file.fullname
                # Specify the command line parameters for wusa.exe
                $parameters = $fullname + " /quiet /norestart"
                # Start wusa.exe and pass in the parameters
                $install = [System.Diagnostics.Process]::Start( "wusa",$parameters )
                $install.WaitForExit()
                Write-Host -ForegroundColor Green "Finished installing $file"
    }
Test-Path "C:\Patching"
Remove-Item -Recurse -Force "C:\Patching"
Test-Path "C:\Patching"
    }
  }
#=============================================================================
# END OF SCRIPT: [InteliPatch]
#=============================================================================

