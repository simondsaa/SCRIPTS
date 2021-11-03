####################################################################################
#  Scriptblock
####################################################################################

$ScriptBlock = {

param([String]$strComputer, [string]$targetfolder, [array]$FileList, [string]$scriptpath )



####################################################################################
#  Start doing stuff!
####################################################################################

$reply = Get-WmiObject -Class Win32_PingStatus -Filter "Address='$strComputer'"

    if ($reply.statuscode -eq 0)
    {
    $pingIPaddress = $reply.IPV4Address.ipaddresstostring
    $PingResponse = "Success"
                    
    $sysArch = (Get-WmiObject -Class Win32_OperatingSystem -ComputerName $strComputer -ea 0).OSArchitecture
        if ($sysArch -eq "64-bit")
        {
        $sysArch = "x64"
        }
        if ($sysarch -eq "32-bit")
        {
        $sysArch = "x86"
        }
   
#######################################################################################
#  Disable the on-access file scanner
#  
#######################################################################################
        
       if ($sysArch -eq "x86")
       {   
       $Onaccess = (([WMICLASS]"\\$strComputer\ROOT\CIMV2:win32_process").Create('"c:\Program Files\McAfee\VirusScan Enterprise\shstat.exe" -disable'))
       $Onaccess = (([WMICLASS]"\\$strComputer\ROOT\CIMV2:win32_process").Create('"c:\Program Files\McAfee\VirusScan Enterprise\shstat.exe" -disable'))
       $Onaccess = (([WMICLASS]"\\$strComputer\ROOT\CIMV2:win32_process").Create('"c:\Program Files\McAfee\VirusScan Enterprise\shstat.exe" -disable'))
       $Onaccess = (([WMICLASS]"\\$strComputer\ROOT\CIMV2:win32_process").Create('"c:\Program Files\McAfee\VirusScan Enterprise\shstat.exe" -disable'))
       }
       else
       {
       $Onaccess = (([WMICLASS]"\\$strComputer\ROOT\CIMV2:win32_process").Create('"c:\Program Files (x86)\McAfee\VirusScan Enterprise\shstat.exe" -disable'))
       $Onaccess = (([WMICLASS]"\\$strComputer\ROOT\CIMV2:win32_process").Create('"c:\Program Files (x86)\McAfee\VirusScan Enterprise\shstat.exe" -disable'))
       $Onaccess = (([WMICLASS]"\\$strComputer\ROOT\CIMV2:win32_process").Create('"c:\Program Files (x86)\McAfee\VirusScan Enterprise\shstat.exe" -disable'))
       $Onaccess = (([WMICLASS]"\\$strComputer\ROOT\CIMV2:win32_process").Create('"c:\Program Files (x86)\McAfee\VirusScan Enterprise\shstat.exe" -disable'))
       }
   
                       
#######################################################################################
#  Loop through install files and run the correct install function
#  
#######################################################################################
        Write-Host "Creating Cyber Protection Folder on Target System." -Foregroundcolor Green
        New-Item -Path \\$strComputer\c$\CyberProtection -ItemType directory -force
        
        Foreach ($objFile in $FileList)
            {
            $filename = $objFile.name
            $filepath = $objFile.fullname
            Copy-Item $filepath \\$strComputer\c$\CyberProtection\$filename -force
            }
            
        Copy-Item $scriptpath\installer.cmd \\$strComputer\c$\installer.cmd -force

        $install = (([WMICLASS]"\\$strComputer\ROOT\CIMV2:win32_process").Create('c:\installer.cmd')).processid

        $wait1 = get-process -cn $strComputer -pid $install -ea SilentlyContinue
                        
          while ($wait1 -ne $null)
              {
               start-sleep -seconds 5
               write-host "waiting for Install to complete"
               $wait1 = get-process -cn $strComputer -pid $install -ea SilentlyContinue
              }
        
        start-sleep -seconds 10
        
        write-host "Beginning File Deletion, Executing Patch Deleter" -Foregroundcolor Green
        
        copy-item $scriptpath\patchdeleter.bat \\$strComputer\c$\patchdeleter.bat -force

        $install2 = (([WMICLASS]"\\$strComputer\ROOT\CIMV2:win32_process").Create('cmd /c start c:\patchdeleter.bat')).processid

        $deletewait = get-process -cn $strComputer -pid $install2 -ea SilentlyContinue

            while ($deletewait -ne $null)
                {
                 start-sleep -seconds 5
                 write-host "Waiting for Patch Deleter to finish clean up."
                 $deletewait = get-process -cn $strComputer -pid $install2 -ea SilentlyContinue
                }

        
        #remove delete file
        remove-item \\$strComputer\c$\patchdeleter.bat
    

#######################################################################################
#  Re-enable Onaccess scanning
#
#  
#######################################################################################     
     
       if ($sysArch -eq "x86")
       {   
       $Onaccess = (([WMICLASS]"\\$strComputer\ROOT\CIMV2:win32_process").Create('"c:\Program Files\McAfee\VirusScan Enterprise\shstat.exe" -enable'))
       }
       else
       {
       $Onaccess = (([WMICLASS]"\\$strComputer\ROOT\CIMV2:win32_process").Create('"c:\Program Files (x86)\McAfee\VirusScan Enterprise\shstat.exe" -enable'))
       }

     
     }
    
    Else  #means we had a bad ping
    {
    #Report the reason for the bad ping based on WMI ping error code
    
        if ($reply.statuscode -eq $null)   #null status code means no DNS entry was found.  Delete that shit from ADUC (unless it's the three-star's laptop...)                 
        {                   
        $PingResponse = "No DNS Entry"
        $pingIPaddress = $null
        }
        #Report the reason for the bad ping based on WMI ping error code    
        If ($reply.statuscode -eq "11010")
        {
        #Ping timeouts still return the IP address from DNS        
        $pingIPaddress = $reply.IPV4Address.ipaddresstostring

        $PingResponse = "Request Timed Out"
        }
        #Report the reason for the bad ping based on WMI ping error code    
        If ($reply.statuscode -eq "11003")
        {
        $pingIPaddress = $reply.IPV4Address.ipaddresstostring
                        
        $PingResponse = "Unable to reach host"
        
        }

     }                  
 

####################################################################################
#  End of scriptblock
#  
#   
#
####################################################################################

            $ResultProps = @{
                ComputerName =  $strComputer
                PingResponse =  $PingResponse
                ErrorCode = $ReturnCode

            }
          
          $return = New-Object PSObject -Property $ResultProps

return $return

}


####################################################################################
#
#  Pull Target Folder and Target List
#   
#  Init variables
#
####################################################################################

Clear
write-host "Power Patcher Revision 5 - by HA/FOSTER/HICKORY/Parson" -ForegroundColor Green
$instructions = @"

--Power Patcher Instructions--

Provide a target list with all hostnames you wish to patch.

Enter the full path to the directory containing your patches.

This script supports dynamic pushing for the following common programs:
-All standard Microsoft patches
-Adobe Flash
-.NET Framework
-Silverlight

This script does not support the pushing of Java.
DO NOT ATTEMPT TO PUSH JAVA WITH THIS SCRIPT!


To exit this script, press CTRL+SCROLL LOCK




-Hickory


"@

$instructions

Pause
$targetlistpath = read-host "Enter location of target list.  Example:  C:\Scripts\Osan\1TargetSystems.txt"
$targetList = get-content $targetlistpath

$targetFolder = "C:\Users\1180219788A\Downloads\17-10"

$BreakCounter = 0
$results = @()
$MaxConcurrentJobs = 19
$counter = 1
$starttimer = Get-Date
$local = hostname
$sysArch = (Get-WmiObject -Class Win32_OperatingSystem -ComputerName $local -ea 0).OSArchitecture
if ($sysArch -eq "64-bit")
    {
    $sysArch = "x64"
    }
if ($sysarch -eq "32-bit")
    {
    $sysArch = "x86"
    }
if ($sysArch -eq "x86")
    {   
    $Onaccess = (([WMICLASS]"\\$local\ROOT\CIMV2:win32_process").Create('"c:\Program Files\McAfee\VirusScan Enterprise\shstat.exe" -disable'))
    $Onaccess = (([WMICLASS]"\\$local\ROOT\CIMV2:win32_process").Create('"c:\Program Files\McAfee\VirusScan Enterprise\shstat.exe" -disable'))
    $Onaccess = (([WMICLASS]"\\$local\ROOT\CIMV2:win32_process").Create('"c:\Program Files\McAfee\VirusScan Enterprise\shstat.exe" -disable'))
    $Onaccess = (([WMICLASS]"\\$local\ROOT\CIMV2:win32_process").Create('"c:\Program Files\McAfee\VirusScan Enterprise\shstat.exe" -disable'))
    }
else
    {
    $Onaccess = (([WMICLASS]"\\$local\ROOT\CIMV2:win32_process").Create('"c:\Program Files (x86)\McAfee\VirusScan Enterprise\shstat.exe" -disable'))
    $Onaccess = (([WMICLASS]"\\$local\ROOT\CIMV2:win32_process").Create('"c:\Program Files (x86)\McAfee\VirusScan Enterprise\shstat.exe" -disable'))
    $Onaccess = (([WMICLASS]"\\$local\ROOT\CIMV2:win32_process").Create('"c:\Program Files (x86)\McAfee\VirusScan Enterprise\shstat.exe" -disable'))
    $Onaccess = (([WMICLASS]"\\$local\ROOT\CIMV2:win32_process").Create('"c:\Program Files (x86)\McAfee\VirusScan Enterprise\shstat.exe" -disable'))
    }

######################################################################################
#
# Expandomatic - H&R BLOCK
#
######################################################################################

$checkmsu = Get-Childitem -Name $targetfolder\*.msu
If ($checkmsu -ne $null)
    {
    Write-Host "Expanding .MSU files. Please wait..." -Foregroundcolor Green
    
    New-Item -Path "c:\CABFILES" -ItemType directory -force | Out-Null
    $cabdir = "c:\CABFILES"
    
    Foreach ($file in $checkmsu)
        {
        Expand -F:* $targetfolder\$file $cabdir | Out-Null
        Write-Host "Files expanding, please wait..."
        }
    
    Remove-Item $cabdir\wsusscan.cab
    Remove-Item $targetfolder\*.msu
    
    Copy-Item $cabdir/*.cab -destination $targetfolder
    Remove-Item $cabdir -Force -Recurse
    
    $cabcount = (Get-Childitem $targetfolder\*.cab).Count
    Write-Host "Expansion complete. $cabcount files expanded." -Foregroundcolor Green
    }

$FileList = Get-ChildItem -path $targetFolder -recurse

if ($sysArch -eq "x86")
{   
$Onaccess = (([WMICLASS]"\\$local\ROOT\CIMV2:win32_process").Create('"c:\Program Files\McAfee\VirusScan Enterprise\shstat.exe" -enable'))
}
else
{
$Onaccess = (([WMICLASS]"\\$local\ROOT\CIMV2:win32_process").Create('"c:\Program Files (x86)\McAfee\VirusScan Enterprise\shstat.exe" -enable'))
}

###############################################################################################
<#

 This creates in the install bat file.  If you have multiple install files,
 put them all in here.  

 The install strings will be explicit and based on the remote machine's C Drive  
 Ensure that all File names are set appropriately.

 
 ###############################################################################################
 Ensure that all DISM, MSIEXEC and EXE install lines are executed with the appropriate
 install variables to ensure patch execution and application to vulnerable client.
 
 Run the patch via command line with a "/?" to check the install switches.
 
 EXE files with an embedded MSI will need a seperate install line.

 Ensure you test the variables for the file on a client that can be physically controlled by you
 to verify that the switches function as they should.

#>
#################################################################################################

Write-Host "Performing Patch Installation." -ForegroundColor Black

$scriptpath = (get-location).path
$cablist = Get-Childitem $targetfolder\*.cab
$exelist = Get-Childitem $targetfolder\*.exe -Exclude "*jre*", "*flash*", "*NDP20*", "*NDP40*", "*NDP45*", "*Silverlight*", "*air*"
$msilist = Get-Childitem $targetfolder\*.msi
$msplist = Get-Childitem $targetfolder\*.msp
$flashlist = Get-Childitem $targetfolder\*flash*
$airlist = Get-Childitem $targetfolder\*air*
$ndplist = Get-Childitem $targetfolder\*NDP40*
$silverlist = Get-ChildItem $targetfolder\*Silverlight*
$javalist = Get-Childitem $targetfolder\*jre*

clear-content $scriptpath\installer.cmd -ea SilentlyContinue

If ($ndplist -ne $null)
    {
        Foreach ($file in $ndplist)
            {
            $ndp = $file.name
            $ndpline = "C:\CyberProtection\$ndp /quiet /norestart"
            Add-Content $scriptpath\installer.cmd -Value $ndpline
            }
    }

If ($cablist -ne $null)
    {
        Foreach ($file in $cablist)
            {
            $cab = $file.name
            $dismline = "DISM.exe /Online /Add-Package /PackagePath:C:\CyberProtection\$cab /quiet /norestart"
            add-content $scriptpath\installer.cmd -value $dismline  
            }
    }
    
If ($exelist -ne $null)
    {
        Foreach ($file in $exelist)
            {
             $exe = $file.name
             $exeline = "C:\CyberProtection\$exe /quiet /norestart"
             add-content $scriptpath\installer.cmd -value $exeline
            }
     }

If ($silverlist -ne $null)
    {
        Foreach ($file in $silverlist)
            {
            $silver = $file.name
            $silverline = "C:\Cyberprotection\$silver /q"
            Add-Content $scriptpath\installer.cmd -Value $silverline
            }
    }

If ($flashlist -ne $null)
    {
        Foreach ($file in $flashlist)
            {
            $flash = $file.name
            $flashline = "C:\CyberProtection\$flash -install"
            Add-Content $scriptpath\installer.cmd -value $flashline
            }
    }
    
If ($airlist -ne $null)
    {
        Foreach ($file in $airlist)
            {
            $air = $file.name
            $airline = "C:\CyberProtection\$air -uninstall"
            Add-Content $scriptpath\installer.cmd -value $airline
            }
    }
         
If ($msilist -ne $null)
    {
        Foreach ($file in $msilist)
            {
            $msi = $file.name
            $msiline = "msiexec /install C:\CyberProtection\$msi /quiet /norestart"
            add-content $scriptpath\installer.cmd -value $msiline
            }
    }
            
If ($msplist -ne $null)
    {
        Foreach ($file in $msplist)
            {
            $msp = $file.name
            $mspline = "msiexec /update C:\CyberProtection\$msp /quiet /norestart"
            add-content $scriptpath\installer.cmd -value $mspline
            }
    }

$wipeshockwave = @"

wmic product where "name like '%%shockwave%%'" call uninstall 

MsiExec.exe /X{612C34C7-5E90-47D8-9B5C-0F717DD82726} /passive
MsiExec.exe /X{BCFB58FF-181E-472F-A9DB-827B75C1EDF7} /passive
MsiExec.exe /X{3B834B54-EC4B-48E2-BFC6-03FF5DA06F62} /passive

"C:\windows\system32\Adobe\Shockwave 11\uninstaller.exe"
"C:\windows\syswow64\Adobe\Shockwave 11\uninstaller.exe"


reg query hklm\software\classes\installer\products /f "shockwave" /s | find "HKEY_LOCAL_MACHINE" > delshock.txt
for /f "tokens=* delims= " %%a in (delshock.txt) do reg delete %%a /f
del delshock.txt

reg delete "HKEY_LOCAL_MACHINE\SOFTWARE\Adobe\Shockwave 11" /f
reg delete "HKEY_LOCAL_MACHINE\SOFTWARE\Adobe\Shockwave 12" /f
reg delete "HKEY_LOCAL_MACHINE\SOFTWARE\Macromedia\Shockwave 10" /f

"@

$exitline = "exit"
Add-Content $scriptpath\installer.cmd -Value $wipeshockwave
add-content $scriptpath\installer.cmd -value $exitline 


####################################################################################
#
#  This is the Patch Deleter bat, This file back 
#  
#
####################################################################################

$cmdTT = @"


cd c:\
rmdir C:\CyberProtection /S /Q
del installer.cmd /Q
exit
"@


clear-content $scriptpath\patchdeleter.bat -ea SilentlyContinue

add-content $scriptpath\patchdeleter.bat -value $cmdTT

####################################################################################
#
#  Multi thread start
#   
#  
####################################################################################

    foreach ($machineName in $targetList)
  {
    Write-host "$counter `t $machineName Starting Job..."
    $counter ++

    start-job -name ("InstallJob-" + $machineName) -scriptblock $scriptblock -ArgumentList $machineName, $targetFolder, $FileList, $scriptpath | out-null

 
        while (((get-job | where-object { $_.Name -like "InstallJob*" -and $_.State -eq "Running" }) | measure).Count -gt $MaxConcurrentJobs)
	    {
        "{0} Concurrent jobs running, sleeping 5 seconds" -f $MaxConcurrentJobs
	    Start-Sleep -seconds 5
	    }


    get-job | where { $_.Name -like "InstallJob*" -and $_.state -eq "Completed" } | % { $results += Receive-Job $_ ; Remove-Job $_ }
    
    }


	while (((get-job | where-object { $_.Name -like "InstallJob*" -and $_.state -eq "Running" }) | measure).count -gt 0)
	{

    get-job | where { $_.Name -like "InstallJob*" -and $_.state -eq "Completed" } | % { $results += Receive-Job $_ ; Remove-Job $_ }
        
		$jobcount = ((get-job | where-object { $_.Name -like "InstallJob*" -and $_.state -eq "Running" }) | measure).count
		Write-Host "Waiting for $jobcount Jobs to Complete sleeping 5 seconds" 
		Start-Sleep -seconds 5
        
            $BreakCounter++
        if ($BreakCounter -gt 100) {
            Write-Host "Exiting loop $jobCount Jobs did not complete"
            get-job | where-object { $_.Name -like "*InstallJob*" -and $_.state -eq "Running" } | select Name
            break
            }    
    }


get-job | where { $_.Name -like "*InstallJob*" -and $_.state -eq "Completed" } | % { $results += Receive-Job $_ ; Remove-Job $_ }




####################################################################################
#
#  END OF MULTITHREAD
#   
#  
####################################################################################


####################################################################################
#
#  Send results to grid view
#   
#  Print total runtime  
#
####################################################################################

$results | select ComputerName, PingResponse | Sort-Object PingResponse | out-gridview

                    
#Pulls end time and prints total time for actions
$stoptimer = Get-Date
"Total time for JOBs: {0} Minutes" -f [math]::round(($stoptimer - $starttimer).TotalMinutes , 2)