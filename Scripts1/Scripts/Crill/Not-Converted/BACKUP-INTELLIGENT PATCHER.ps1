#* FileName: Intelligent_Patcher.ps1
#*=============================================================================
#* Script Name: Intelligent Patcher
#* Created: 09/04/2014
#* Author: SSgt Crill, Christian
#* Company: USAF
#* Email: christian.crill@us.af.mil
#* Web: 
#* Reqrmnts:
#* Keywords:
#*=============================================================================
#  WORKFLOW
# Ping if alive do below LOGGED
# Manage Remote machine LOGGED
# If failed repair machine LOGGED
# Manage repaired Remote machine LOGGED
# 
#      NEED:  Patching script addition 
#*=============================================================================
#*=============================================================================
#* REVISION HISTORY
#*=============================================================================
#* Date: [DATE_MDY]
#* Time: [TIME]
#* Issue:
#* Solution:
#*
#*=============================================================================
#*=============================================================================
#* FUNCTION LISTINGS
#*=============================================================================


#Function: $comps
#Names of remote machines6
$comps = Get-Content "C:\Users\1394844760A\Desktop\Scripting Test Bed\names.txt"

#Function: $updatedir
#Specify the location of the *.msu files
$updatedir = "\\xlwu-fs-002\tyndall$\Applications\MSBatch\Sept"

#Function: $LogPath
#Specify the location of the Log Files
$LogPath = "C:\Users\1394844760A\Desktop\Scripting Test Bed\Logs"

#Function: $Domain
#Fully Qualified Domain Name (FQDN)
$Domain = ".area52.afnoapps.usaf.mil"

#Function: cPing
# Ping
Function cPing 
{
  Test-Connection $comp -Quiet -Buffersize 16 -Ea 0 -Count 1
}
# =============================================================================
# Function: cArc_Patch
# Created: [07/17/2014]
# Author: SrA Roberson, David
# Arguments:
# =============================================================================
# Purpose: Check if computer is x86 or x64 in order to push the correct patches.
# eventually work this into the invoke-command
# =============================================================================

Function cArc_Patch 
{
    #Get the Operating System architecture
    $OS = (Get-WmiObject Win32_OperatingSystem).OSArchitecture

    $files = Get-ChildItem $updatedir -exclude *.bat,*.exe -Recurse

    If ($OS -eq "64-bit")
    {
        ForEach ($file in $files)
        {
            If ($file.Name -like "*x64*")
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
        }
    }
    ElseIf ($OS -eq "32-bit")
    {
        ForEach ($file in $files)
        {
            If ($file.Name -like "*x86*")
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
        }
    }
    Else
    {
        Write-Host -ForegroundColor Yellow "OS Architecture UNKNOWN, quitting..."
        Write-Host -ForegroundColor Red "Press any key to Exit..."
        $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyUp") > $null
        Exit
    }
}

#=============================================================================
# SCRIPT BODY
#=============================================================================
#Call the “Pinging” function / Patch if online / Log if not online.

ForEach ($comp in $comps)
{
    If (cPing)
    {
        Try
        {
             Invoke-Command -ErrorAction Stop -Computer ($comp) -Scriptblock {hostname}
             Out-File "$LogPath\Success.txt" -Append -InputObject "$comp, REMOTE" -Force
        }
        Catch
        {
             Write-Host "Running PSRemoting Fix"
             psexec \\$comp -s -h -d powershell Enable-PSRemoting -Force
             Out-File "$LogPath\Attempted_Repair.txt" -Append -InputObject "$comp" -Force
             $R_Comps = Get-Content "$LogPath\Attempted_Repair.txt" 
        }
        }
    Else
    {
        Out-File "$LogPath\No_Ping.txt" -Append -InputObject "$comp, OFFLINE" -Force
    }
}

ForEach ($R_comp in $R_comps)
{
    Try
    {
    Invoke-Command -ErrorAction Stop -Computer ($R_comp) -Scriptblock {hostname}
    Out-File "$LogPath\Repaired Success.txt" -Append -InputObject "$R_comp,REMOTE Repaired" -Force
    }
    Catch
    {
    Out-File "$LogPath\Repaired Fail.txt" -Append -InputObject "$R_comp,REMOTE Failed" -Force
    }
}
#=============================================================================
# END OF SCRIPT: [Intelligent Patcher]
#=============================================================================