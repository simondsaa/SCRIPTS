# Written by SrA Timothy Brady
# Tyndall AFB, Panama City, FL
# Created August 11, 2015

# NOTES
# -----
# RISK: Password will be typed in plain text.
# You need to modify the path below to the text file with a list of computer names

$Path = "C:\Computers.txt"

# SCRIPT BEGINS
# -------------
$Date = Get-Date -Format "MMMMM dd, yyyy HH:mm"
$Space = " "

# Gets your EDI and creates a log path for Offline Systems
$User = Get-WmiObject Win32_ComputerSystem
If ($User.UserName -ne $null)
{
    $EDI = $User.UserName.TrimStart("AREA52\")
}
If (!(Test-Path "C:\Users\$EDI\Documents\Logs"))
{
    New-Item -Path C:\Users\$EDI\Documents\Logs -Type Directory -Force
}
$LogPath = "C:\Users\$EDI\Documents\Logs"

# Promtps for the password you want to set
$Password = Read-Host -Prompt "Enter New Password"
$Computers = Get-Content -Path $Path
ForEach ($Computer in $Computers)
{
    If (Test-Connection $Computer -Quiet -BufferSize 16 -Ea 0 -Count 1)
    {
         $Admin = [adsi]"WinNT://$Computer/USAF_Admin,user"
         $Admin.SetPassword($Password)
         $Admin.SetInfo()
    }
    Else
    {
        Out-File "$LogPath\Offline_Systems.txt" -Force -InputObject $Date$Space$Computer -Append
    }
}