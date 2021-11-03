##############################################################################
#
# Reboot Script - User Policy
#
# Authors: 561 NOS
#
# Created:  November 4th, 2013
#
# Revision: 1.3
#
##############################################################################

#You will need to comment the exit out in order for the script to work.(WARNING: Only do this once you have everything 
#configured for your base). You will put a # before exit.
#Exit



###Script Admin defined variables###

#The exemption group to place users or machines in so they won't be affected by the script.  They may be the exact same group
$Computerexemptiongroup = "GLS_schriever_reboot_exempt"
$Userexemptiongroup = "GLS_schriever_reboot_exempt"

#This variable is used to define how you want time measured. It needs to be one of the following variables:
#Days, Hours, Minutes
$Timemeasure = "days"

#The amount of time a system needs to be up before it is rebooted. (This variable will be in Days, Hours, or Minutes depending upon
# what you define in the $Timemeasure variable above).
$Timelength = 7

####################################



#Gather local information, and determine what users are actively logged into the machine.
$users = Get-WmiObject Win32_Process -filter "Name = 'explorer.exe'" | foreach -process {$_.GetOwner().User}

#Exit script if any users are logged on
If($users -ne $Null)
    {
    Write-host "Test"
    }

#Determine system uptime in days
$time=[DateTime]::Now - [Management.ManagementDateTimeConverter]::ToDateTime((Get-WmiObject Win32_OperatingSystem).LastBootUpTime)
$uptime = $time.$Timemeasure

#Determine if system uptime exceeded time limit 
If($uptime -le $Timelength)
    {
    exit
    }
Else
    {
    Invoke-Expression "C:\Windows\System32\shutdown.exe -f -t 300 -r"
    }