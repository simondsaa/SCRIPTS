# Written by SSgt Timothy Brady
# Tyndall AFB, Panama City, FL
# Created January 28, 2016

# MODIFICATIONS
# -------------
# 01 Feb 16 - Added logging of System Name and which profiles are being removed
# 02 Feb 16 - Added the date the last time the profile was used to the log
# 10 Feb 16 - Changed the log paths to work properly with old versions of PowerShell
# 11 Feb 16 - Moved the "No profiles removed" out of the ForEach loop to stop repetitiveness in the logs

# NOTES
# -----
# Deletes user profiles that are older than the $Days variable. 
# The date is based off of the "Last Use Date" of the profile itself, not the last modify date of the user folder.

# CHANGES
# -------

$Days = 365

$LogPath = "\\xlwu-fs-05pv\Tyndall_PUBLIC\Stats\Profile Cleanup\" + $env:COMPUTERNAME + ".txt"
$OldLogPath = "\\xlwu-fs-05pv\Tyndall_PUBLIC\Stats\Profile Cleanup\Old\" + $env:COMPUTERNAME + ".txt"

# SCRIPT BEGINS
# -------------

If ((Test-Path -Path $LogPath) -eq $true)
{
    Move-Item -Path $LogPath -Destination $OldLogPath -Force
}

$Date = (Get-Date).AddDays(-$Days)

$Profiles = Get-WmiObject Win32_UserProfile | Select *

ForEach ($Profile in $Profiles)
{
    If (($Profile.LocalPath -like "C:\Users*") -and ($Profile.LocalPath -notlike "*Admin*"))
    {
        $ProfDate = [System.Management.ManagementDateTimeconverter]::ToDateTime($Profile.LastUseTime)
    
        If ($ProfDate -lt $Date)
        {
            $ProfilePath = $Profile.LocalPath
            
            $Today = Get-Date
	        Out-File -FilePath $LogPath -Force -InputObject "$ProfilePath - last used $ProfDate - removed $Today" -Append
            Get-WmiObject Win32_UserProfile | Where {$_.LocalPath -eq $Profile.LocalPath} | ForEach {$_.Delete()}
        }
        Else
        {
            $Output = $true
        }
    }
}

If ($Output -eq $true)
{
    $Today = Get-Date
    Out-File -FilePath $LogPath -Force -InputObject "No profiles removed - $Today" -Append
}