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
$Timemeasure = "Days"

#The amount of time a system needs to be up before it is rebooted. (This variable will be in Days, Hours, or Minutes depending upon
# what you define in the $Timemeasure variable above).
$Timelength = 7

#The amount of time for the users to wait before their system is rebooted. (Note: This is in minutes)
$countdowntime = 60

#This is the Air Force installation in which this script will be ran at.
$AFB = "Peterson Air Force Base"

#The amount of time to give the users, if they need additional time. (Note: This is in minutes.)
$addtionaltime = 60

####################################



#Gather local information, and determine what users are actively logged into the machine.
$Computergroups = ([adsisearcher]"(&(objectCategory=computer)(cn=$env:COMPUTERNAME))").FindOne().Properties.memberof -replace '^CN=([^,]+).+$','$1'
$users = Get-WmiObject Win32_Process -filter "Name = 'explorer.exe'" | foreach -process {$_.GetOwner().User}
$Computergroups
#Determine if there is an exempted account currently logged on to the system.  Exit if true.
If($Computergroups -contains $Computerexemptiongroup)
    {
    exit
    }
Else
    {
    Foreach($user in $users)
        {
        $user
        $usergroups = ([adsisearcher]"(&(objectCategory=user)(objectClass=user)(SamAccountName=$user))").FindOne().Properties.memberof -replace '^CN=([^,]+).+$','$1'
        $usergroups
        If($usergroups -contains $Userexemptiongroup)
            {
            exit
            }
        $usergroups = $null
        }
    }

#Determine system uptime speceific by the varible of "$timemeasure" (I.E. Days, Hours, Minutes)
$time=[DateTime]::Now - [Management.ManagementDateTimeConverter]::ToDateTime((Get-WmiObject Win32_OperatingSystem).LastBootUpTime)
$uptime = $time.$Timemeasure

#Determine if system uptime exceeded time limit 
If($uptime -ge $Timelength)
    {
    exit
    }

#Run shutdown function
[int]$timeleft = $countdowntime
[System.Reflection.Assembly]::LoadWithPartialName("System.Diagnostics")
$countdowntimer = New-Object system.diagnostics.stopwatch 
While($timeleft -gt 0)
    {
    $countdowntimer.start()
    $a = new-object -comobject wscript.shell
    $intAnswer = $a.popup("Your system exceeded the uptime requirments of $AFB.  The machine will reboot in $timeleft minutes.  If you require more time, please click Yes for an additional $addtionaltime minutes.  If you wish to reboot immediately, click No.  If you wish to close the window and allow the system to reboot on it's own in the specified time period, click Cancel",60,"System Restart",3 + 4096)
    If($intAnswer -eq 6)
        {
        $countdowntimer.stop()
        Start-Sleep -Seconds ($addtionaltime * 60)
        $countdowntimer.reset()
        While($timeleft -gt 0)
            {
            $countdowntimer.start()
            $a = new-object -comobject wscript.shell
            $intAnswer = $a.popup("Your system exceeded the uptime requirments of $AFB.  The machine will reboot in $timeleft minutes.  You are not authorized addtional time.  Click OK to permanently close this box.",60,"System Restart",0 + 4096)
            If ($intAnswer -eq 1)
                {
                Invoke-Expression "C:\Windows\System32\shutdown.exe -f -t ($timeleft * 60) -r"
                Exit
                }
            $countdowntimer.reset()
            $timeleft--
            }
        }
    ElseIf($intAnswer -eq 7)
        {
        Invoke-Expression "C:\Windows\System32\shutdown.exe -f -r"
        Exit
        }
    ElseIf($intAnswer -eq 2)
        {
        Invoke-Expression "C:\Windows\System32\shutdown.exe -f -t ($timeleft * 60) -r"
        Exit
        }
  $countdowntimer.reset()
  $timeleft--
    }
Invoke-Expression "C:\Windows\System32\shutdown.exe -f -r"