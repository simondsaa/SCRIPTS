################################################################################
# Java Repair for ULUC2 Clients
# Requires PSEXEC 
# Targets single machine / or a listing of machines
# Sets a SYSTEM task at startup to change JAVA_HOME variable every reboot.
#
# SSgt Crill, Christian 325 CS/SCOO
# 22 July 2015
################################################################################
#=============================================================================
#                                  MENU
#=============================================================================
    CLS
    Write-Host "Java Repair, Created by SSgt Crill, Christian 325 CS/SCOO" -ForegroundColor Yellow
    Write-Host
    Write-Host "SCRIPT REQUIREMENTS will be in red" -BackgroundColor Yellow -Foregroundcolor Red
    Write-Host "Run as Administrator" -Foregroundcolor Red 
    Write-Host "PSExec on local machine" -ForegroundColor Red

    Write-host "Please Verify your target list" -BackgroundColor Yellow -Foregroundcolor Red
    pause

################################################################################
# Network share for below batch jobs must have Domain Computers with Read + Execute as NTFS Permissions

Do
{
    Cls
    Write-Host " "
    Write-Host "1 - Single Machine"
    Write-Host "2 - List of Machines"
    Write-Host "3 - Exit"
    Write-Host " "

    $Ans = Read-Host "Make Selection"
    
    If ($Ans -eq 1)
    {
        Write-Host
        $Computername = Read-Host "Target Computer" 
        psexec.exe \\$computername -c -f "\\xlwu-fs-05pv\Tyndall_PUBLIC\Logons\ULUC2\main.bat"
        psexec.exe \\$computername -c -f "\\xlwu-fs-05pv\Tyndall_PUBLIC\Logons\ULUC2\ULUC2_BootStrap.bat"
        Write-Host "$computername Patched" -ForegroundColor Green
        pause
    }
    If ($Ans -eq 2)
    {
        Write-Host
        $List_Path = Read-Host "FULL path to computer list" 
        $ComputerList = get-content "$list_path"
        ForEach ($comp in $computerlist) {
        psexec.exe \\$computername -c -f "\\xlwu-fs-05pv\Tyndall_PUBLIC\Logons\ULUC2\main.bat"
        psexec.exe \\$comp -c -f "\\xlwu-fs-05pv\Tyndall_PUBLIC\Logons\ULUC2\ULUC2_BootStrap.bat"
        Write-Host "$comp Patched" -ForegroundColor Green
        }
        pause
}
}
Until ($Ans -eq 3)
pause
