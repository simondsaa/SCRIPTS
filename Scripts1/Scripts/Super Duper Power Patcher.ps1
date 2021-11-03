#######################################################################################
#  Scriptblock
#######################################################################################

$ScriptBlock = {

param([String]$strComputer, [string]$targetfolder, [array]$FileList, [string]$scriptpath )

$reply = Get-WmiObject -Class Win32_PingStatus -Filter "Address='$strComputer'"

    If ($reply.statuscode -eq 0)
        {
        $pingIPaddress = $reply.IPV4Address.ipaddresstostring
        $PingResponse = "Success"
                    
        $sysArch = (Get-WmiObject -Class Win32_OperatingSystem -ComputerName $strComputer -ea 0).OSArchitecture
   
#######################################################################################
#  Loop through install files and run the correct install function  
#######################################################################################

    New-Item -Path \\$strComputer\c$\Temp\CyberProtection -ItemType Directory -Force

        foreach ($objFile in $FileList)
            {
            $filename = $objFile.name
            $filepath = $objFile.fullname
            Copy-Item $filepath \\$strComputer\c$\Temp\CyberProtection\$filename -Force
            }
            
        Copy-Item $scriptpath\installer.cmd \\$strComputer\c$\Temp\installer.cmd -Force

        $install = (([WMICLASS]"\\$strComputer\ROOT\CIMV2:win32_process").Create('c:\Temp\installer.cmd')).processid

        $wait1 = Get-Process -Cn $strComputer -PID $install -ea SilentlyContinue
                        
          while ($wait1 -ne $null)
              {
               Start-Sleep -seconds 5
               Write-Host "waiting for Install to complete operation."
               $wait1 = Get-Process -Cn $strComputer -PID $install -ea SilentlyContinue
              }
        
        Start-Sleep -seconds 10
        
        Write-Host "Beginning File Deletion." -ForegroundColor Green
        
        Copy-Item $scriptpath\patchdeleter.bat \\$strComputer\c$\Temp\patchdeleter.bat -Force

        $install2 = (([WMICLASS]"\\$strComputer\ROOT\CIMV2:win32_process").Create('cmd /c start c:\Temp\patchdeleter.bat')).processid

        $deletewait = Get-Process -Cn $strComputer -PID $install2 -ea SilentlyContinue

            while ($deletewait -ne $null)
                {
                 Start-Sleep -seconds 5
                 Write-Host "Waiting for Patch Deleter to complete operation."
                 $deletewait = Get-Process -Cn $strComputer -PID $install2 -ea SilentlyContinue
                }                           
     }
    
    Else  # means we had a bad ping
        {
        if ($reply.statuscode -eq $null)   # null status code means no DNS entry was found.  Delete that shit from ADUC (unless it's the three-star's laptop...)                 
            {                   
            $PingResponse = "No DNS Entry"
            $pingIPaddress = $null
            }
        # Report the reason for the bad ping based on WMI ping error code    
        If ($reply.statuscode -eq "11010")
            {
            # Ping timeouts still return the IP address from DNS        
            $pingIPaddress = $reply.IPV4Address.ipaddresstostring
            $PingResponse = "Request Timed Out"
            }
        # Report the reason for the bad ping based on WMI ping error code    
        If ($reply.statuscode -eq "11003")
            {
            $pingIPaddress = $reply.IPV4Address.ipaddresstostring
            $PingResponse = "Unable to reach host"
            }
        }                  
 
#######################################################################################
#  End of scriptblock
#######################################################################################

            $ResultProps = @{
                ComputerName =  $strComputer
                PingResponse =  $PingResponse
                ErrorCode = $ReturnCode
                }
          
          $return = New-Object PSObject -Property $ResultProps

return $return
}

#######################################################################################
#  Pull Target Folder and Target List
#  Init variables
#######################################################################################

clear
Write-Host "SUPER DUPER POWER PATCHER v1.0" -ForegroundColor Yellow -BackgroundColor Black
Write-Host "Created By Robie" -ForegroundColor Yellow -BackgroundColor Black

$targetlistpath = Read-Host "Enter path of target list."
$targetList = Get-Content $targetlistpath

$targetFolder = "\\xlwu-fs-04pv\Tyndall_325_MSG\325 CS\SCO\SCOO\PowerShell"

$BreakCounter = 0
$results = @()
$MaxConcurrentJobs = 19
$counter = 1
$starttimer = Get-Date
$local = hostname

#######################################################################################
# Expandomatic - H&R BLOCK
#######################################################################################

$CheckMSU = Get-ChildItem -Name $targetfolder\PowerShellv3_Install.msu

If ($CheckMSU -ne $null)
    {
    Write-Host "Expanding .MSU files. Please Wait..." -ForegroundColor Green
    
    New-Item -Path "C:\Temp\CABFILES" -ItemType directory -Force | Out-Null
    $cabdir = "c:\Temp\CABFILES"
    
    Foreach ($file in $CheckMSU)
        {
        Expand -F:* $targetfolder\$file $cabdir | Out-Null
        }
    
    Remove-Item $cabdir\wsusscan.cab
    Remove-Item $targetfolder\*.msu
    
    Copy-Item $cabdir/*.cab -Destination $targetfolder
    Remove-Item $cabdir -Force -Recurse
    
    $cabcount = (Get-ChildItem $targetfolder\*.cab).Count
    Write-Host "Expansion complete. $cabcount files expanded."
    }

$FileList = Get-ChildItem -Path $targetFolder -Recurse

#######################################################################################
<#

 This creates in the install bat file.  If you have multiple install files,
 put them all in here.  

 The install strings will be explicit and based on the remote machine's C Drive  
 Ensure that all File names are set appropriately.

 
#######################################################################################
 Ensure that all DISM, MSIEXEC and EXE install lines are executed with the appropriate
 install variables to ensure patch execution and application to vulnerable client.
 
 Run the patch via command line with a "/?" to check the install switches.
 
 EXE files with an embedded MSI will need a seperate install line.

 Ensure you test the variables for the file on a client that can be physically controlled by you
 to verify that the switches function as they should.

#>
#######################################################################################

Write-Host "Performing Patch Installation..." -ForegroundColor Green

$scriptpath = (Get-Location).path
$cablist = Get-ChildItem $targetfolder\*.cab
$exelist = Get-ChildItem $targetfolder\*.exe -Exclude "*jre*", "*flash*", "*NDP20*", "*NDP40*", "*NDP45*", "*Silverlight*", "*air*"
$msilist = Get-ChildItem $targetfolder\*.msi
$msplist = Get-ChildItem $targetfolder\*.msp
$flashlist = Get-ChildItem $targetfolder\*flash*
$airlist = Get-ChildItem $targetfolder\*air*
$ndplist = Get-ChildItem $targetfolder\*NDP40*
$silverlist = Get-ChildItem $targetfolder\*Silverlight*
$javalist = Get-ChildItem $targetfolder\*jre*

Clear-Content $scriptpath\installer.cmd -ea SilentlyContinue

If ($ndplist -ne $null)
    {
        Foreach ($file in $ndplist)
            {
            $ndp = $file.name
            $ndpline = "C:\Temp\CyberProtection\$ndp /quiet /norestart"
            Add-Content $scriptpath\installer.cmd -Value $ndpline
            }
    }

If ($cablist -ne $null)
    {
        Foreach ($file in $cablist)
            {
            $cab = $file.name
            $dismline = "DISM.exe /Online /Add-Package /PackagePath:C:\Temp\CyberProtection\$cab /quiet /norestart"
            Add-Content $scriptpath\installer.cmd -Value $dismline  
            }
    }
    
If ($exelist -ne $null)
    {
        Foreach ($file in $exelist)
            {
             $exe = $file.name
             $exeline = "C:\Temp\CyberProtection\$exe /quiet /norestart"
             Add-Content $scriptpath\installer.cmd -Value $exeline
            }
     }

If ($silverlist -ne $null)
    {
        Foreach ($file in $silverlist)
            {
            $silver = $file.name
            $silverline = "C:\Temp\Cyberprotection\$silver /q"
            Add-Content $scriptpath\installer.cmd -Value $silverline
            }
    }

If ($flashlist -ne $null)
    {
        Foreach ($file in $flashlist)
            {
            $flash = $file.name
            $flashline = "C:\Temp\CyberProtection\$flash -install"
            Add-Content $scriptpath\installer.cmd -Value $flashline
            }
    }
    
If ($airlist -ne $null)
    {
        Foreach ($file in $airlist)
            {
            $air = $file.name
            $airline = "C:\Temp\CyberProtection\$air -uninstall"
            Add-Content $scriptpath\installer.cmd -Value $airline
            }
    }
         
If ($msilist -ne $null)
    {
        Foreach ($file in $msilist)
            {
            $msi = $file.name
            $msiline = "msiexec /install C:\Temp\CyberProtection\$msi /quiet /norestart"
            Add-Content $scriptpath\installer.cmd -Value $msiline
            }
    }
            
If ($msplist -ne $null)
    {
        Foreach ($file in $msplist)
            {
            $msp = $file.name
            $mspline = "msiexec /update C:\Temp\CyberProtection\$msp /quiet /norestart"
            Add-Content $scriptpath\installer.cmd -Value $mspline
            }
    }

$exitline = "exit"
Add-Content $scriptpath\installer.cmd -Value $exitline 

#######################################################################################
#  This is the Patch Deleter bat, This file back 
#######################################################################################

$cmdTT = @"


cd c:\
rmdir C:\Temp\CyberProtection /S /Q
del C:\Temp\installer.cmd /Q
exit
"@

Clear-Content $scriptpath\patchdeleter.bat -ea SilentlyContinue

Add-Content $scriptpath\patchdeleter.bat -Value $cmdTT

#######################################################################################
#  Multi thread start
#######################################################################################

    Foreach ($machineName in $targetList)
  {
    Write-Host "$counter `t $machineName Starting Job..."
    $counter ++

    Start-Job -Name ("InstallJob-" + $machineName) -ScriptBlock $scriptblock -ArgumentList $machineName, $targetFolder, $FileList, $scriptpath | Out-Null

 
        While (((Get-Job | Where-Object { $_.Name -like "InstallJob*" -and $_.State -eq "Running" }) | measure).Count -gt $MaxConcurrentJobs)
	    {
        "{0} Concurrent jobs running" -f $MaxConcurrentJobs
	    Start-Sleep -Seconds 5
	    }


    Get-Job | Where { $_.Name -like "InstallJob*" -and $_.state -eq "Completed" } | % { $results += Receive-Job $_ ; Remove-Job $_ }
    }

	While (((Get-Job | Where-Object { $_.Name -like "InstallJob*" -and $_.state -eq "Running" }) | measure).count -gt 0)
	{

    Get-Job | Where { $_.Name -like "InstallJob*" -and $_.state -eq "Completed" } | % { $results += Receive-Job $_ ; Remove-Job $_ }
        
		$jobcount = ((Get-Job | Where-Object { $_.Name -like "InstallJob*" -and $_.state -eq "Running" }) | measure).count
		Write-Host "Waiting for $jobcount Job(s) to complete." 
		Start-Sleep -Seconds 5
        
            $BreakCounter++
        If ($BreakCounter -gt 100) {
            Write-Host "Exiting loop $jobCount Job(s) did not complete."
            Get-Job | Where-Object { $_.Name -like "*InstallJob*" -and $_.state -eq "Running" } | Select Name
            Break
            }    
    }

Get-Job | Where { $_.Name -like "*InstallJob*" -and $_.state -eq "Completed" } | % { $results += Receive-Job $_ ; Remove-Job $_ }

#######################################################################################
#  END OF MULTITHREAD
#######################################################################################

#######################################################################################
#  Send results to grid view
#  Print total runtime  
#######################################################################################

$results | Select ComputerName, PingResponse | Sort-Object PingResponse | Out-GridView
                    
# Pulls end time and prints total time for actions
$stoptimer = Get-Date
"Total time for JOBs: {0} Minutes" -f [math]::round(($stoptimer - $starttimer).TotalMinutes , 2)