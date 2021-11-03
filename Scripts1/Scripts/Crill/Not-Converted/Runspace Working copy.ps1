cls
$ADsPath = "LDAP://OU=Tyndall AFB Computers,OU=Tyndall AFB,OU=AFCONUSEAST,OU=Bases,DC=AREA52,DC=AFNOAPPS,DC=USAF,DC=MIL"
$strFilter = "(&(objectCategory=Computer))"
$showError=0
#This is total number of jobs that will run at one time.
$MaxConcurrentJobs = 100

$quote= [char]34
#$joe_command = "cscript.exe $quote\\ang.ds.af.mil\sysvol\ang.ds.af.mil\Policies\{****GUID****}\User\Scripts\Logon\angtcno.vbs$quote"
#$joe_command = "cscript.exe $quote\\110fw-fs-01\ULI\Login_Scripts\angtcno_110awmod.vbs$quote"
$reboot_command = "shutdown -r -t 600"
#$90MeterInstall_command = "cscript.exe $quote\\110fw-fs-01\uli\Login_Scripts\90MeterInstallation.vbs$quote"

$SCCM_ManagementPoint = "52XLWU-CM-004v.area52.afnoapps.usaf.mil"
$CCM_Uninstall_command_x86 = "C:\Windows\System32\ccmsetup\ccmsetup.exe /uninstall"
$CCM_Install_command_x86 = "C:\Windows\System32\ccmsetup\ccmsetup.exe /mp:"+$SCCM_ManagementPoint+" smssitecode=XLW"
$CCM_Uninstall_command_x64 = "C:\Windows\ccmsetup\ccmsetup.exe /uninstall"
$CCM_Install_command_x64 = "C:\Windows\ccmsetup\ccmsetup.exe /mp:"+$SCCM_ManagementPoint+" smssitecode=XLW"

[String[]]$AllComputerNames=@()
[String[]]$Awake=@()
[String[]]$Asleep=@()
[String[]]$Excluded = @()
[String[]]$Target=@()

$ADObj = New-Object PSObject;
$ADInfo = @();
$PSRemotingStatuses=@{}
$CombinedADandPingResults=@()
$Global:ADPE411=@()
$Global:IAVMResults=@()
$Global:IAVMInstalled=@()
$Global:IAVMMissing=@()

$SCCMDayHashTable=@{0="None";1="Sunday";2="Monday";3="Tuesday";4="Wednesday";5="Thurdsay";6="Friday";7="Saturday";8="Daily";}
$SCCMHourHashTable=@{0="0000";1="0100";2="0200";3="0300";4="0400";5="0500";6="0600";7="0700";8="0800";9="0900";10="1000";11="1100";12="1200";13="1300";14="1400";15="1500";16="1600";17="1700";18="1800";19="1900";20="2000";21="2100";22="2200";23="2300";}
$SCCMProgressHashTable=@{0="Update_Progress_None";1="Update_Progress_Optional_Install";2="Update_Progress_Mandatory_Install"}
#Options Flags for SCCM Update Install.  These are bit-wise flags.
$SCCMInstallOptions=0x0001 -BOR 0x0002 -BOR 0x0008 -BOR 0x0010 -BOR 0x0020

#Variables for setting the ClientUI "Install Required Updates on a Schedule" value.  Hash tables used to convert to COM Object usable values.
$SCCMDesiredInstallDay="Daily"
$SCCMDesiredInstallHour="1700"

$DNSDomain = $ENV:USERDNSDOMAIN
$DelegateDomain = "*."+$DNSDomain
$objDomain = New-Object System.DirectoryServices.DirectoryEntry
$objOU = New-Object System.DirectoryServices.DirectoryEntry($ADsPath)

$objSearcher = New-Object System.DirectoryServices.DirectorySearcher
$objSearcher.SearchRoot = $objOU
$objSearcher.PageSize = 1000
$objSearcher.Filter = $strFilter
$objSearcher.SearchScope = "Subtree"

$colProplist = "name","description","lastLogon","lastLogonTimestamp","logonCount","pwdLastSet","whenCreated"
foreach ($i in $colPropList){$objSearcher.PropertiesToLoad.Add($i)}

$CollSystems = $objSearcher.FindAll()
Write-Host
"Approx {0} Systems for Availability check" -f ($CollSystems | Measure-Object).count    
Write-Host
    ForEach($objResult in $CollSystems) 
    {
        $objItem = $objResult.Properties; $AllComputerNames +=$objItem.name;$ADObj = New-Object PSObject;$ADObj | 
        Add-Member NoteProperty Name $objItem.name -Force; $ADObj | 
        Add-Member NoteProperty Description $objItem.description -Force; $ADObj | 
        Add-Member NoteProperty LogonCount $objItem.logoncount -Force; $mydatetime = $objItem.lastlogon; $ADObj | 
        Add-Member NoteProperty LastLogon ([datetime]::FromFileTime($($mydatetime))) -Force; $mydatetime = $objItem.lastlogontimestamp; $ADObj | 
        Add-Member NoteProperty LastLogonTimeStamp ([datetime]::FromFileTime($($mydatetime))) -Force; $mydatetime = $objItem.pwdlastset; $ADObj | 
        Add-Member NoteProperty PasswordLastSet ([datetime]::FromFileTime($($mydatetime))) -Force;$mydatetime = $objItem.whencreated; $ADObj | 
        Add-Member NoteProperty Created $mydatetime -Force;$ADInfo += $ADObj
    }

#Sort all systems array ascending
$Computername = $AllComputerNames | Sort-Object
Measure-Command {
#$Computername = Get-Content "C:\Users\1394844760A\Desktop\Scripting Test Bed\names.txt"
$scriptblock = 
{
  Param ($computer)
  if (Test-Connection $computer -Quiet -Count 1 -EA 0 ){
        $GatherWMI = Get-WmiObject -Query "Select * from Win32_PingStatus Where timeout=3000 and Address='$Computer'"
      }
          [PSCustomObject]@{
                Computer_Name = $computer
                IP_Address = $GatherWMI.IPv4Address
                                    }
}

$RunspacePool = [RunspaceFactory]::CreateRunspacePool(100,100)
$RunspacePool.Open()
$Jobs = 
   foreach ( $computer in $Computername)
    {
     $Job = [powershell]::Create().AddScript($ScriptBlock).AddArgument($computer)
     $Job.RunspacePool = $RunspacePool

     [PSCustomObject]@{
      Pipe = $Job
      Result = $Job.BeginInvoke()
     }
}

Write-Host 'Working..' 

Do {
   Write-host "Still Working" 
   Start-Sleep -Seconds 1
} While ( $Jobs.Result.IsCompleted -contains $false)

Write-Host ' Done! Writing output file.'
Write-host "C:\Windows\Temp\PingResults.csv"
$RunspacePool.Close()

$(ForEach ($Job in $Jobs)
{ $Job.Pipe.EndInvoke($Job.Result) }) |Export-CSV "C:\Windows\Temp\PingResults.csv"
$RunspacePool.Close()
$RunspacePool.Dispose()
}

$PingResults = Import-Csv "C:\Windows\Temp\PingResults.csv"

$PingOffline = $PingResults | where{-Not $_.IP_Address}
$PingOnline = $PingResults | where-object{$_ -ne $null}

    Write-Host
	Write-Host "Availability check complete!"
    "Total Systems : {0}" -f ($Computername | Measure-Object).count 
    Write-Host
    "Total Systems Offline: {0}" -f (($PingResults | where{-Not $_.IP_Address} | measure-object).count)
    "Total Systems Online : {0}" -f (($PingResults | where-object{$_ -ne $null} | measure-object).count)
    $PingResults | ogv

   
    
    $Awake += $PingOnline
   # $Target = $Awake | Where {$Excluded -notcontains $_}
    

    
function Minion-Enable-PSRemoting-Client
    {
        param(
            [parameter(Mandatory = $false)]
            $computername
        )
        
        Enable-PSRemoting -Force
        #Enable-WSManCredSSP -Role Client -DelegateComputer $DelegateDomain -Force
        Return
    }
    
    function Minion-Enable-PSRemoting-Server
    {
        param(
            [parameter(Mandatory = $true)]
            $computername
        )
        #const
        $quote= [char]34
        #Remote Commands to execute to enable our flavor of remoting.  To be executed via WMI.
        $command_1 = "schtasks.exe /CREATE /TN 'Minion-Enable-WSRemoting1' /SC ONCE /ST 17:00 /RL HIGHEST /RU SYSTEM /TR $quote powershell.exe -noprofile -command Enable-PSRemoting -Force$quote /F"
        $command_2 = "schtasks.exe /CREATE /TN 'Minion-Enable-WSRemoting2' /SC ONCE /ST 17:00 /RL HIGHEST /RU SYSTEM /TR $quote powershell.exe -noprofile -command Set-WSManQuickConfig -Force$quote /F"
        $command_3 = "schtasks.exe /CREATE /TN 'Minion-Enable-WSRemoting3' /SC ONCE /ST 17:00 /RL HIGHEST /RU SYSTEM /TR $quote powershell.exe -noprofile -command Enable-WSManCredSSP -Role Server -Force$quote /F"
        $command_run1 = "schtasks.exe /RUN /TN 'Minion-Enable-WSRemoting1'"
        $command_run2 = "schtasks.exe /RUN /TN 'Minion-Enable-WSRemoting2'"
        $command_run3 = "schtasks.exe /RUN /TN 'Minion-Enable-WSRemoting3'"
        $command_delete1 = "schtasks.exe /DELETE /TN 'Minion-Enable-WSRemoting1' /F"
        $command_delete2 = "schtasks.exe /DELETE /TN 'Minion-Enable-WSRemoting2' /F"
        $command_delete3 = "schtasks.exe /DELETE /TN 'Minion-Enable-WSRemoting3' /F"
                
        $process = [WMICLASS]"\\$computername\ROOT\CIMV2:Win32_Process"
        $result1 = $process.Create($command_1)
        $result2 = $process.Create($command_run1)
        Write-Host "Enabling PSRemoting on:        " $computername
        Start-Sleep -s 5
        $result3 = $process.Create($command_delete1)
        Start-Sleep -s 1
        $result4 = $process.Create($command_2)
        $result5 = $process.Create($command_run2)
        Write-Host "Configuring WSMan on:          " $computername
        Start-Sleep -s 5
        $result6 = $process.Create($command_delete2)
        Start-Sleep -s 1
        $result7 = $process.Create($command_3)
        $result8 = $process.Create($command_run3)
        Write-Host "Configuring CredSSP Server on: " $computername
        Start-Sleep -s 5
        $result9 = $process.Create($command_delete3)
        Return 
    }
    
    function Minion-Disable-PSRemoting-Server
    {
        param(
            [parameter(Mandatory = $true)]
            $computername
        )
        #const
        $quote= [char]34
        #Remote Commands to execute to disable our flavor of remoting.  To be executed via WMI.
        $command_5 = "schtasks.exe /CREATE /TN 'Minion-Disable-WSRemoting5' /SC ONCE /ST 17:00 /RL HIGHEST /RU SYSTEM /TR $quote powershell.exe -noprofile -command Disable-PSRemoting -Force$quote /F"
        $command_4 = "schtasks.exe /CREATE /TN 'Minion-Disable-WSRemoting4' /SC ONCE /ST 17:00 /RL HIGHEST /RU SYSTEM /TR $quote sc config WinRM start= disabled $quote /F"
        $command_3 = "schtasks.exe /CREATE /TN 'Minion-Disable-WSRemoting3' /SC ONCE /ST 17:00 /RL HIGHEST /RU SYSTEM /TR $quote net stop \$quote"+"Windows Remote Management (WS-Management)\$quote $quote /F"
        $command_2 = "schtasks.exe /CREATE /TN 'Minion-Disable-WSRemoting2' /SC ONCE /ST 17:00 /RL HIGHEST /RU SYSTEM /TR $quote powershell.exe -noprofile -command Remove-WSManInstance winrm/config/listener -selectorset @{Address=\\\$quote*\\\$quote;Transport=\\\"+$quote+"http\\\$quote} $quote /F"
        $command_1 = "schtasks.exe /CREATE /TN 'Minion-Disable-WSRemoting1' /SC ONCE /ST 17:00 /RL HIGHEST /RU SYSTEM /TR $quote powershell.exe -noprofile -command Disable-WSManCredSSP -Role Server -Force$quote /F"
        $command_run1 = "schtasks.exe /RUN /TN 'Minion-Disable-WSRemoting1'"
        $command_run2 = "schtasks.exe /RUN /TN 'Minion-Disable-WSRemoting2'"
        $command_run3 = "schtasks.exe /RUN /TN 'Minion-Disable-WSRemoting3'"
        $command_run4 = "schtasks.exe /RUN /TN 'Minion-Disable-WSRemoting4'"
        $command_run5 = "schtasks.exe /RUN /TN 'Minion-Disable-WSRemoting5'"
        $command_delete1 = "schtasks.exe /DELETE /TN 'Minion-Disable-WSRemoting1' /F"
        $command_delete2 = "schtasks.exe /DELETE /TN 'Minion-Disable-WSRemoting2' /F"
        $command_delete3 = "schtasks.exe /DELETE /TN 'Minion-Disable-WSRemoting3' /F"
        $command_delete4 = "schtasks.exe /DELETE /TN 'Minion-Disable-WSRemoting4' /F"
        $command_delete5 = "schtasks.exe /DELETE /TN 'Minion-Disable-WSRemoting5' /F"
                
        $process = [WMICLASS]"\\$computername\ROOT\CIMV2:Win32_Process"
        $result1 = $process.Create($command_1)
        $result2 = $process.Create($command_run1)
        Write-Host "Disabling WSManCredSSP on:  " $computername
        Start-Sleep -s 5
        $result3 = $process.Create($command_delete1)
        Start-Sleep -s 1
        $result4 = $process.Create($command_2)
        $result5 = $process.Create($command_run2)
        Write-Host "Removing Listener on:       " $computername
        Start-Sleep -s 5
        $result6 = $process.Create($command_delete2)
        Start-Sleep -s 1
        $result7 = $process.Create($command_3)
        $result8 = $process.Create($command_run3)
        Write-Host "Stopping WSMan Service on:  " $computername
        Start-Sleep -s 5
        $result9 = $process.Create($command_delete3)
        Start-Sleep -s 1
        $result10 = $process.Create($command_4)
        $result11 = $process.Create($command_run4)
        Write-Host "Disabling WSMan Service on: " $computername
        Start-Sleep -s 5
        $result12 = $process.Create($command_delete4)
        Start-Sleep -s 1
        $result13 = $process.Create($command_5)
        $result14 = $process.Create($command_run5)
        Write-Host "Disabling PSRemoting on:    " $computername
        Start-Sleep -s 2
        $result15 = $process.Create($command_delete5)
        Return 
    }
    
    function Minion-Enable-PSRemoting-Server-MultiThreaded
    {
        param(
            [parameter(Mandatory = $true)]
            $paraArray
        )
        $starttimer = Get-Date
        $paraArray = $paraArray | Sort-Object
        
    #Initialization Script for MultiThreaded Start-Job -Begin
        $iniScript = { function Minion-Enable-PSRemoting-Server
    {
        param(
            [parameter(Mandatory = $true)]
            $computername
        )
        #const
        $quote= [char]34
        #Remote Commands to execute to enable our flavor of remoting.  To be executed via WMI.
        $command_1 = "schtasks.exe /CREATE /TN 'Minion-Enable-WSRemoting1' /SC ONCE /ST 17:00 /RL HIGHEST /RU SYSTEM /TR $quote powershell.exe -noprofile -command Enable-PSRemoting -Force$quote /F"
        $command_2 = "schtasks.exe /CREATE /TN 'Minion-Enable-WSRemoting2' /SC ONCE /ST 17:00 /RL HIGHEST /RU SYSTEM /TR $quote powershell.exe -noprofile -command Set-WSManQuickConfig -Force$quote /F"
        $command_3 = "schtasks.exe /CREATE /TN 'Minion-Enable-WSRemoting3' /SC ONCE /ST 17:00 /RL HIGHEST /RU SYSTEM /TR $quote powershell.exe -noprofile -command Enable-WSManCredSSP -Role Server -Force$quote /F"
        $command_run1 = "schtasks.exe /RUN /TN 'Minion-Enable-WSRemoting1'"
        $command_run2 = "schtasks.exe /RUN /TN 'Minion-Enable-WSRemoting2'"
        $command_run3 = "schtasks.exe /RUN /TN 'Minion-Enable-WSRemoting3'"
        $command_delete1 = "schtasks.exe /DELETE /TN 'Minion-Enable-WSRemoting1' /F"
        $command_delete2 = "schtasks.exe /DELETE /TN 'Minion-Enable-WSRemoting2' /F"
        $command_delete3 = "schtasks.exe /DELETE /TN 'Minion-Enable-WSRemoting3' /F"
                
        $process = [WMICLASS]"\\$computername\ROOT\CIMV2:Win32_Process"
        $result1 = $process.Create($command_1)
        $result2 = $process.Create($command_run1)
        Write-Host "Enabling PSRemoting on:        " $computername
        Start-Sleep -s 5
        $result3 = $process.Create($command_delete1)
        Start-Sleep -s 1
        $result4 = $process.Create($command_2)
        $result5 = $process.Create($command_run2)
        Write-Host "Configuring WSMan on:          " $computername
        Start-Sleep -s 5
        $result6 = $process.Create($command_delete2)
        Start-Sleep -s 1
        $result7 = $process.Create($command_3)
        $result8 = $process.Create($command_run3)
        Write-Host "Configuring CredSSP Server on: " $computername
        Start-Sleep -s 5
        $result9 = $process.Create($command_delete3)
        Return $P
    }
    }
    function Minion-Get-ADPE
    {
        param(
            [parameter(Mandatory = $true)]
            $paraArray
        )
        $starttimer = Get-Date
        $paraArray = $paraArray | Sort-Object
        Write-Host ""        
        "Enumerating ADPE information for {0} Systems." -f ($paraArray | Measure-Object).count
        Write-Host ""
        $paraArray | foreach {       
        if ($_.length -gt 0)
		    {
			     "Enumerating ADPE information for {0} " -f $_
                 start-job -scriptblock {$wmi = Get-WmiObject Win32_OperatingSystem -comp $args[0] | Select CSName, Version, CSDVersion, Caption, Description, OSArchitecture; $obj=New-Object PSObject; $obj | Add-Member NoteProperty CSName ($wmi.CSName); $obj | Add-Member NoteProperty Version ($wmi.Version); $obj | Add-Member NoteProperty SPVersion ($wmi.CSDVersion); $obj | Add-Member NoteProperty Caption ($wmi.Caption); $obj | Add-Member NoteProperty Description ($wmi.Description); $obj | Add-Member NoteProperty OSArchitecture ($wmi.OSArchitecture); $wmi = Get-WmiObject Win32_BIOS -comp $args[0] | Select SerialNumber; $obj | Add-Member NoteProperty BIOSSerial ($wmi.SerialNumber); $wmi = Get-WmiObject Win32_NetworkAdapterConfiguration -comp $args[0]; $IPAddress = $wmi | Where {$_.IPAddress} | Select -Expand IPAddress; $DefaultIPGateway = $wmi | Where {$_.DefaultIPGateway} | Select -Expand DefaultIPGateway; $SubnetMask = $wmi | Where {$_.IPSubnet} | Select -Expand IPSubnet; $Description = $wmi | Where {$_.IPAddress} | Select -Expand Description; $obj | Add-Member NoteProperty IPAddress ($IPAddress); $obj | Add-Member NoteProperty SubnetMask ($SubnetMask); $obj | Add-Member NoteProperty DefaultIPGateway ($DefaultIPGateway); $obj | Add-Member NoteProperty NICDesc ($Description);$wmi = Get-WmiObject AF_ImageRevision -comp $args[0] | Select ImageRevision; $obj | Add-Member NoteProperty AF_Revision ($wmi.ImageRevision); $wmi = Get-WmiObject AF_Revision_Detail -comp $args[0] | Select CurrentBuild; $obj | Add-Member NoteProperty CurrentBuild ($wmi.CurrentBuild);$wmi = Get-WmiObject Win32_ComputerSystem -comp $args[0] | Select Manufacturer,Model,UserName; $obj | Add-Member NoteProperty Manufacturer ($wmi.Manufacturer); $obj | Add-Member NoteProperty Model ($wmi.Model); $obj | Add-Member NoteProperty CurrentUser ($wmi.UserName); $obj;} -name("ADPE-" + $_) -argumentlist $_ | Out-Null
            }    		
		    while (((get-job | where-object { $_.Name -like "ADPE-*" -and $_.State -eq "Running" }) | measure).Count -gt $MaxConcurrentJobs)
		    {
                "{0} Concurrent jobs running, sleeping 5 seconds" -f $MaxConcurrentJobs
			    Start-Sleep -seconds 5
		    }
	    }
        while (((get-job | where-object { $_.Name -like "ADPE-*" -and $_.state -eq "Running" }) | measure).count -gt 0)
	    {
		  $jobcount = ((get-job | where-object { $_.Name -like "ADPE-*" -and $_.state -eq "Running" }) | measure).count
		  Write-Host "Waiting for $jobcount Jobs to Complete" 
		  Start-Sleep -seconds 5
          $Counter++
            if ($Counter -gt 40) {
                Write-Host "Exiting loop $jobCount Jobs did not complete"
                get-job | where-object { $_.Name -like "ADPE-*" -and $_.state -eq "Running" } | select Name
                break
            }
	     }
         $Global:ADPEResults = @()
         get-job | where { $_.Name -like "ADPE-*" -and $_.state -eq "Completed" } | % { $Global:ADPEResults += Receive-Job $_ ; Remove-Job $_ }
	     $stoptimer = Get-Date
         $Global:ADPE411 = $Global:ADPEResults | Select CSName,BIOSSerial,CurrentUser,Manufacturer,Model,Version,SPVersion,Caption,AF_Revision,CurrentBuild,Description,OSArchitecture,IPAddress,SubnetMask,DefaultIPGateway,NICDesc | Where {$_.CSName -ne $Null} | Sort-Object CSName
         $Global:ADPE34MigrationEligible  = $ADPE411 | Where-Object {$_.CurrentBuild -like '3.*' -and $_.CurrentBuild -notlike '*3.4*'}
                 
         "Total Time for ADPE Enumeration: {0} Minutes" -f [math]::round(($stoptimer - $starttimer).TotalMinutes , 2)
         Write-Host
         "Total Systems: {0} " -f ($Awake | Measure-Object).count
         Write-Host
         "Total Systems 3.4 Migration Eligible  : {0} " -f ($ADPE34MigrationEligible | Measure-Object).count
         "Total Systems 3.4 : {0}" -f ($ADPE411 | Where-Object {$_.CurrentBuild -like '*3.4*'}).count
         Write-Host
         "Total Systems !Enumerated : {0} " -f [math]::Abs(($Awake | Measure-Object).count - ($ADPE411 | Measure-Object).count)
         "Total Systems Enumerated  : {0} " -f ($ADPE411 | Measure-Object).count
         
                  
         Return $ADPE411 | Out-GridView
    }
    
function Minion-Get-Progs
    {
        param(
            [parameter(Mandatory = $true)]
            $paraArray
        )
        #if ($cred -eq $null) {Write-Host "No Credentials supplied, requesting presently..." ; $Global:cred = Minion-Get-Cred -Domain $DNSDomain} else {Write-Host "Credential check pass...proceeding"}
        $starttimer = Get-Date
        $paraArray = $paraArray | Sort-Object
        Write-Host ""        
        "Enumerating Installed Applications for {0} Systems." -f ($paraArray | Measure-Object).count
        Write-Host ""
        $paraArray | foreach {       
        if ($_.length -gt 0)
		    {
			     "Enumerating Installed Applications for {0} " -f $_
                 start-job -scriptblock {$session = New-PSSession -cn ($args[0] +"."+ $args[1]); Invoke-Command -Session $session -ScriptBlock {$OSArch = (Get-WmiObject Win32_OperatingSystem).OSArchitecture; $OSArch_WoW6432 = 'WoW6462';$ComputerName = gc env:computername;$APPS_Native = gci "hklm:\software\microsoft\windows\currentversion\uninstall" | foreach { gp $_.PSPath } | select @{name='KeyName';expression={$_.PSChildName}},@{name='Architecture';expression={$OSArch}},@{name='SystemName';expression={$ComputerName}},DisplayName,DisplayVersion,InstallDate,ModifyPath,Publisher,UninstallString,Language;If($OSArch -eq '64-bit'){$APPS_WoW6432 = gci "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\" | foreach { gp $_.PSPath } | select @{name='KeyName';expression={$_.PSChildName}},@{name='Architecture';expression={$OSArch_WoW6432}},@{name='SystemName';expression={$ComputerName}},DisplayName,DisplayVersion,InstallDate,ModifyPath,Publisher,UninstallString,Language};$APPS = $APPS_Native+$APPS_WoW6432 | Sort-Object DisplayName;$APPS}} -name("Progs-" + $_) -argumentlist $_ , $DNSDomain | Out-Null
            }    		
		    while (((get-job | where-object { $_.Name -like "Progs-*" -and $_.State -eq "Running" }) | measure).Count -gt $MaxConcurrentJobs)
		    {
                "{0} Concurrent jobs running, sleeping 5 seconds" -f $MaxConcurrentJobs
			    Start-Sleep -seconds 5
		    }
	    }
        while (((get-job | where-object { $_.Name -like "Progs-*" -and $_.state -eq "Running" }) | measure).count -gt 0)
	    {
		  $jobcount = ((get-job | where-object { $_.Name -like "Progs-*" -and $_.state -eq "Running" }) | measure).count
		  Write-Host "Waiting for $jobcount Jobs to Complete" 
		  Start-Sleep -seconds 5
          $Counter++
            if ($Counter -gt 40) {
                Write-Host "Exiting loop $jobCount Jobs did not complete"
                get-job | where-object { $_.Name -like "Progs-*" -and $_.state -eq "Running" } | select Name
                break
            }
	     }
         $Global:ProgsResults = @()
         get-job | where { $_.Name -like "Progs-*" -and $_.state -eq "Completed" } | % { $Global:ProgsResults += Receive-Job $_ ; Remove-Job $_ }
	     $stoptimer = Get-Date
         #Commented out - $Global:ADPE411 = $Global:ADPEResults | Select CSName,BIOSSerial,CurrentUser,Manufacturer,Model,Version,SPVersion,Caption,AF_Revision,Description,OSArchitecture,IPAddress,SubnetMask,DefaultIPGateway,NICDesc | Where {$_.CSName -ne $Null} | Sort-Object CSName
         
         "Total Time for Installed Application Enumeration: {0} Minutes" -f [math]::round(($stoptimer - $starttimer).TotalMinutes , 2)
         Write-Host
         "Total Systems: {0} " -f ($Awake | Measure-Object).count
         Write-Host
         "Total Systems !Enumerated : {0} " -f [math]::Abs(($Awake | Measure-Object).count - ($ProgsResults | Measure-Object).count)
         "Total Systems Enumerated  : {0} " -f ($ProgsResults | Measure-Object).count
         
                  
         Return $ProgsResults | Out-GridView
    }
    
    function Minion-Get-QFE
    {
        param(
            [parameter(Mandatory = $true)]
            $paraArray
        )
        #if ($cred -eq $null) {Write-Host "No Credentials supplied, requesting presently..." ; $Global:cred = Minion-Get-Cred -Domain $DNSDomain} else {Write-Host "Credential check pass...proceeding"}
        $starttimer = Get-Date
        $paraArray = $paraArray | Sort-Object
        Write-Host ""        
        "Enumerating QuickFixEngineering (Security Updates, HotFixes, Updates) for {0} Systems." -f ($paraArray | Measure-Object).count
        Write-Host ""
        $paraArray | foreach {       
        if ($_.length -gt 0)
		    {
			     "Enumerating HotFixes for {0} " -f $_
                 start-job -scriptblock {Get-Hotfix -ComputerName ($args[0] +"."+ $args[1])} -name("HotFixQFE-" + $_) -argumentlist $_ , $DNSDomain | Out-Null
            }    		
		    while (((get-job | where-object { $_.Name -like "HotFixQFE-*" -and $_.State -eq "Running" }) | measure).Count -gt $MaxConcurrentJobs)
		    {
                "{0} Concurrent jobs running, sleeping 5 seconds" -f $MaxConcurrentJobs
			    Start-Sleep -seconds 5
		    }
	    }
        while (((get-job | where-object { $_.Name -like "HotFixQFE-" -and $_.state -eq "Running" }) | measure).count -gt 0)
	    {
		  $jobcount = ((get-job | where-object { $_.Name -like "HotFixQFE-*" -and $_.state -eq "Running" }) | measure).count
		  Write-Host "Waiting for $jobcount Jobs to Complete" 
		  Start-Sleep -seconds 5
          $Counter++
            if ($Counter -gt 40) {
                Write-Host "Exiting loop $jobCount Jobs did not complete"
                get-job | where-object { $_.Name -like "HotFixQFE-*" -and $_.state -eq "Running" } | select Name
                break
            }
	     }
         $Global:QFEResults = @()
         get-job | where { $_.Name -like "HotFixQFE-*" -and $_.state -eq "Completed" } | % { $Global:QFEResults += Receive-Job $_ ; Remove-Job $_ }
	     $stoptimer = Get-Date
                 
         "Total Time for Installed Application Enumeration: {0} Minutes" -f [math]::round(($stoptimer - $starttimer).TotalMinutes , 2)
         Write-Host
         "Total Systems: {0} " -f ($Awake | Measure-Object).count
         Write-Host
         "Total Systems !Enumerated : {0} " -f [math]::Abs(($Awake | Measure-Object).count - ($QFEResults|Select -Unique CSName|Measure-Object).count)
         "Total Systems Enumerated  : {0} " -f ($QFEResults|Select -Unique CSName|Measure-Object).count
         
                  
         Return $QFEResults | Out-GridView
    }
    
    
    function Minion-Get-IAVM
    {
        param(
            [parameter(Mandatory = $true)]
            $paraArray
        )
        #if ($cred -eq $null) {Write-Host "No Credentials supplied, requesting presently..." ; $Global:cred = Minion-Get-Cred -Domain $DNSDomain} else {Write-Host "Credential check pass...proceeding"}
        $starttimer = Get-Date
        $paraArray = $paraArray | Sort-Object
        Write-Host ""        
        "Enumerating IAVM information for {0} Systems." -f ($paraArray | Measure-Object).count
        Write-Host ""
        $paraArray | foreach {       
        if ($_.length -gt 0)
		    {
			     "Enumerating IAVM information for {0} " -f $_
                 start-job -scriptblock {$session = New-PSSession -cn ($args[0] +"."+ $args[1]); Invoke-Command -Session $session -ScriptBlock {Get-WmiObject -Namespace "Root\ccm\softwareupdates\updatesstore" -Class CCM_UpdateStatus | Select __SERVER, Status, Bulletin, Article, Title, UniqueID, ScanTime}} -name("IAVM-" + $_) -argumentlist $_ , $DNSDomain | Out-Null
            }    		
		    while (((get-job | where-object { $_.Name -like "IAVM-*" -and $_.State -eq "Running" }) | measure).Count -gt $MaxConcurrentJobs)
		    {
                "{0} Concurrent jobs running, sleeping 5 seconds" -f $MaxConcurrentJobs
			    Start-Sleep -seconds 5
		    }
	    }
        while (((get-job | where-object { $_.Name -like "IAVM-*" -and $_.state -eq "Running" }) | measure).count -gt 0)
	    {
		  $jobcount = ((get-job | where-object { $_.Name -like "IAVM-*" -and $_.state -eq "Running" }) | measure).count
		  Write-Host "Waiting for $jobcount Jobs to Complete" 
		  Start-Sleep -seconds 5
          $Counter++
            if ($Counter -gt 40) {
                Write-Host "Exiting loop $jobCount Jobs did not complete"
                get-job | where-object { $_.Name -like "IAVM-*" -and $_.state -eq "Running" } | select Name
                break
            }
	     }
         
         $Global:IAVMResults=@()
         $Global:IAVMInstalled=@()
         $Global:IAVMMissing=@()
         get-job | where { $_.Name -like "IAVM-*" -and $_.state -eq "Completed" } | % { $Global:IAVMResults += Receive-Job $_ ; Remove-Job $_ }
         $Global:IAVMResults= $IAVMResults | Select __SERVER, Status, Bulletin, Article, Title, UniqueID, ScanTime | Sort-Object
         $Global:IAVMInstalled= $IAVMResults | Select __SERVER, Status, Bulletin, Article, Title, UniqueID, ScanTime | Where {$_.Status -eq "Installed"}
         $Global:IAVMMissing= $IAVMResults | Select __SERVER, Status, Bulletin, Article, Title, UniqueID, ScanTime | Where {$_.Status -eq "Missing"}
         $Global:IAVM_MSBulletin = $IAVMResults | Select __SERVER, Status, Bulletin, Article, Title, UniqueID, ScanTime | Where {$_.Bulletin -like "MS*"}
         $Global:IAVMMissing_MS = $IAVMResults | Select __SERVER, Status, Bulletin, Article, Title, UniqueID, ScanTime | Where {$_.Status -eq "Missing"} | Where {$_.Bulletin -like "MS*"}
         $Global:IAVMInstalled_MS = $IAVMResults | Select __SERVER, Status, Bulletin, Article, Title, UniqueID, ScanTime | Where {$_.Status -eq "Installed"} | Where {$_.Bulletin -like "MS*"}
         
	     $stoptimer = Get-Date
         #$Global:ADPE411 = $Global:ADPEResults | Select CSName,BIOSSerial,CurrentUser,Manufacturer,Model,Version,SPVersion,Caption,AF_Revision,Description,OSArchitecture,IPAddress,DefaultIPGateway,NICDesc | Where {$_.CSName -ne $Null} | Sort-Object CSName
         Write-Host
         "Total Time for IAVM Enumeration: {0} Minutes" -f [math]::round(($stoptimer - $starttimer).TotalMinutes , 2)
         Write-Host
         "Total Systems: {0} " -f ($paraArray | Measure-Object).count
         Write-Host
         "Total Patches Enumerated  : {0} " -f ($IAVMResults | Measure-Object).count
         "Total Patches Installed   : {0} " -f ($IAVMInstalled | Measure-Object).count
         "Total Patches Missing     : {0} " -f ($IAVMMissing |Measure-Object).count
         Write-Host
         "Total Patched   : {0}%  " -f [math]::Round(($IAVMInstalled.count / $IAVMResults.count) * 100,2)
         "Total Unpatched : {0}%  " -f [math]::Round(($IAVMMissing.count / $IAVMResults.count) * 100,2)
         Write-Host
         "Total MS Bulletins Installed   : {0} " -f ($IAVMInstalled_MS | Measure-Object).count
         "Total MS Bulletins Missing     : {0} " -f ($IAVMMissing_MS |Measure-Object).count
         "MS Bulletins Patch Compliance  : {0}%  " -f [math]::Round(($IAVMInstalled_MS.count / $IAVM_MSBulletin.count) * 100,2)
          Return $IAVMResults | Out-GridView
    }
    
    function Minion-Get-SCCMStatus
    {
        param(
            [parameter(Mandatory = $true)]
            $paraArray
        )
        #if ($cred -eq $null) {Write-Host "No Credentials supplied, requesting presently..." ; $Global:cred = Minion-Get-Cred -Domain $DNSDomain} else {Write-Host "Credential check pass...proceeding"}
        $starttimer = Get-Date
        $paraArray = $paraArray | Sort-Object
        Write-Host ""        
        "Enumerating SCCM Status for {0} Systems." -f ($paraArray | Measure-Object).count
        Write-Host ""
        $paraArray | foreach {       
        if ($_.length -gt 0)
		    {
			     "Enumerating SCCM Status for {0} " -f $_
                 #Perform WMI query for OS Architecture prior to PSSession instantiation to ensure 32bit COM can be loaded in 32 bit WoW Process via PS configuration
                 start-job -scriptblock {$wmi = Get-WmiObject Win32_OperatingSystem -comp $args[0] | Select OSArchitecture; if ($wmi.OSArchitecture -eq '64-bit') {$session = New-PSSession -cn ($args[0] +"."+ $args[1]) -ConfigurationName Microsoft.PowerShell32}; if ($wmi.OSArchitecture -eq '32-bit') {$session = New-PSSession -cn ($args[0] +"."+ $args[1])}; Invoke-Command -Session $session -ScriptBlock {$SCCMUpdate = New-Object -ComObject UDA.CCMUpdatesDeployment;$SCCMClientUI = New-Object -ComObject 'CPAPPLET.CPAppletMgr'; $SCCMClientProps = $SCCMClientUI.GetClientProperties();$SCCMClientVersion = ($SCCMClientProps | Where-Object { $_.Name -eq 'ClientVersion' }).value; $SCCMClientCurrentMP = ($SCCMClientProps | Where-Object { $_.Name -eq 'CurrentManagementPoint' }).value;$SCCMClientCurrentUser = ($SCCMClientProps | Where-Object { $_.Name -eq 'UserName' }).value;$SCCMClientADSite = ($SCCMClientProps | Where-Object { $_.Name -eq 'ADSiteName' }).value;$hostname=hostname;$SCCMDayHashTable = $args[0]; $SCCMHourHashTable = $args[1]; $SCCMProgressHashTable = $args[2]; [ref]$Progress = $NULL;[ref]$SCCMDay = $NULL; [ref]$SCCMHour = $NULL; $updates=$SCCMUpdate.EnumerateUpdates(2,1,$Progress);$UpdateCount = $updates.getcount(); $SCCMUpdate.GetUserDefinedSchedule($SCCMDay,$SCCMHour);$ProgressHRF = $SCCMProgressHashTable.Get_Item($Progress.value); $SCCMDayHRF = $SCCMDayHashTable.Get_Item($SCCMDay.value); $SCCMHourHRF = $SCCMHourHashTable.Get_Item($SCCMHour.Value); $SCCMScanTime = (Get-WmiObject -Namespace "Root\ccm\softwareupdates\updatesstore" -Class CCM_UpdateStatus | Select ScanTime | Sort-Object -Unique).ScanTime;if ($args[3] -eq '64-bit') {$SCUPReboot = Get-ChildItem 'HKLM:\SOFTWARE\WoW6432Node\Microsoft\SMS\Mobile Client\Updates Management\Handler\UpdatesRebootStatus'}; if ($args[3] -eq '32-bit') {$SCUPReboot = Get-ChildItem 'HKLM:\SOFTWARE\Microsoft\SMS\Mobile Client\Updates Management\Handler\UpdatesRebootStatus'};$RebootStatus=$NULL;if($SCUPReboot -ne $NULL){$RebootStatus="Pending";$RebootPendingIDs=@(); $SCUPReboot|foreach $_.Name { $RebootPendingIDs += ($_.Name).Replace("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\SMS\Mobile Client\Updates Management\Handler\UpdatesRebootStatus\","")+","}} ; $obj=New-Object PSObject; $obj | Add-Member NoteProperty CSName ($hostname);$obj | Add-Member NoteProperty CurrentInstallProgress ($ProgressHRF);$obj | Add-Member NoteProperty UpdateCount ($UpdateCount);$obj | Add-Member NoteProperty SCCMInstallDaySchedule ($SCCMDayHRF);$obj | Add-Member NoteProperty SCCMClientVersion ($SCCMClientVersion);$obj | Add-Member NoteProperty SCCMCurrentMP ($SCCMClientCurrentMP);$obj | Add-Member NoteProperty SCCMCurrentUser ($SCCMClientCurrentUser);$obj | Add-Member NoteProperty SCCMADSite ($SCCMClientADSite);$obj | Add-Member NoteProperty SCCMInstallHourSchedule ($SCCMHourHRF);$obj | Add-Member NoteProperty SCCMScanTime ($SCCMScanTime);$obj | Add-Member NoteProperty RebootStatus ($RebootStatus);$obj | Add-Member NoteProperty PendingRebootIDs ($RebootPendingIDs); $obj;} -Args $args[3], $args[4], $args[5], $wmi.OSArchitecture} -name("SCCMStatus-" + $_) -argumentlist $_ , $DNSDomain, $cred, $SCCMDayHashTable, $SCCMHourHashTable, $SCCMProgressHashTable | Out-Null
            }    		
		    while (((get-job | where-object { $_.Name -like "SCCMStatus-*" -and $_.State -eq "Running" }) | measure).Count -gt $MaxConcurrentJobs)
		    {
                "{0} Concurrent jobs running, sleeping 5 seconds" -f $MaxConcurrentJobs
			    Start-Sleep -seconds 5
		    }
	    }
        while (((get-job | where-object { $_.Name -like "SCCMStatus-*" -and $_.state -eq "Running" }) | measure).count -gt 0)
	    {
		  $jobcount = ((get-job | where-object { $_.Name -like "SCCMStatus-*" -and $_.state -eq "Running" }) | measure).count
		  Write-Host "Waiting for $jobcount Jobs to Complete" 
		  Start-Sleep -seconds 5
          $Counter++
            if ($Counter -gt 40) {
                Write-Host "Exiting loop $jobCount Jobs did not complete"
                get-job | where-object { $_.Name -like "SCCMStatus-*" -and $_.state -eq "Running" } | select Name
                break
            }
	     }
         
         $Global:SCCMResults=@()
         get-job | where { $_.Name -like "SCCMStatus-*" -and $_.state -eq "Completed" } | % { $Global:SCCMResults += Receive-Job $_ ; Remove-Job $_ }
         $Global:SCCMResults = $SCCMResults | Select CSName, SCCMCurrentUser, CurrentInstallProgress, UpdateCount, SCCMInstallDaySchedule, SCCMInstallHourSchedule, SCCMClientVersion, SCCMCurrentMP, SCCMADSite, SCCMScanTime, RebootStatus, PendingRebootIDs
	     $stoptimer = Get-Date
         Write-Host
         "Total Time for SCCM Status Enumeration: {0} Minutes" -f [math]::round(($stoptimer - $starttimer).TotalMinutes , 2)
         Write-Host
         "Total Systems: {0} " -f ($paraArray | Measure-Object).count
          Return $SCCMResults | Out-GridView
    }
    
    function Minion-Invoke-SCCM
    {
        param(
            [parameter(Mandatory = $true)]
            $paraArray
        )
        Write-Host ""
        #if ($cred -eq $null) {Write-Host "No Credentials supplied, requesting presently..." ; $Global:cred = Minion-Get-Cred -Domain $DNSDomain} else {Write-Host "Credential check pass...proceeding"}
        Write-Host ""
        ##Menu##
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")

        $x=$NULL

        $objForm = New-Object System.Windows.Forms.Form
        $objForm.Text = "Select a SCCM function to Invoke"
        $objForm.Size = New-Object System.Drawing.Size(400,500)
        $objForm.StartPosition = "CenterScreen"

        $objForm.KeyPreview = $True

        $objForm.Add_KeyDown({if ($_.KeyCode -eq "Enter")
            {$x=$objListBox.SelectedItem;$objForm.Close()}})
        $objForm.Add_KeyDown({if ($_.KeyCode -eq "Escape")
            {$objForm.Close()}})
    
        $OKButton = New-Object System.Windows.Forms.Button
        $OKButton.Location = New-Object System.Drawing.Size(75,435)
        $OKButton.Size = New-Object System.Drawing.Size(75,23)
        $OKButton.Text = "OK"
        $OKButton.Add_Click({$x=$objListBox.SelectedItem;$objForm.Close()})
        $objForm.Controls.Add($OKButton)

        $CancelButton = New-Object System.Windows.Forms.Button
        $CancelButton.Location = New-Object System.Drawing.Size(150,435)
        $CancelButton.Size = New-Object System.Drawing.Size(75,23)
        $CancelButton.Text = "Cancel"
        $CancelButton.Add_Click({$objForm.Close()})
        $objForm.Controls.Add($CancelButton)

        $objLabel = New-Object System.Windows.Forms.Label
        $objLabel.Location = New-Object System.Drawing.Size(10,20)
        $objLabel.Size = New-Object System.Drawing.Size(280,20)
        $objLabel.Text = "Please select a Function:"
        $objForm.Controls.Add($objLabel)
 
        $objListBox = New-Object System.Windows.Forms.ListBox
        $objListBox.Location = New-Object System.Drawing.Size(10,40)
        $objListBox.Size = New-Object System.Drawing.Size(350,20)
        $objListBox.Height = 400

        [void] $objListBox.Items.Add("Set User Defined Schedule")
        [void] $objListBox.Items.Add("Install Pending Updates")
        [void] $objListBox.Items.Add("Initate 3.4 Migration")
        [void] $objListBox.Items.Add("Hardware Inventory Collection Task")
        [void] $objListBox.Items.Add("Software Inventory Collection Task")
        [void] $objListBox.Items.Add("Heartbeat Discovery Cycle")
        [void] $objListBox.Items.Add("Software Inventory File Collection Task")
        [void] $objListBox.Items.Add("Machine Policy Assignments Request")
        [void] $objListBox.Items.Add("Evaluate Machine Policy Assignments")
        [void] $objListBox.Items.Add("Refresh Default MP Task")
        [void] $objListBox.Items.Add("Refresh Location Services Task")
        [void] $objListBox.Items.Add("Location Services Cleanup Task")
        [void] $objListBox.Items.Add("Software Metering Report Cycle")
        [void] $objListBox.Items.Add("Source Update Manage Update Cycle")
        [void] $objListBox.Items.Add("Policy Agent Cleanup Cycle")
        [void] $objListBox.Items.Add("Validate Machine Policy Assignments")
        [void] $objListBox.Items.Add("Certificate Maintenance Cycle")
        [void] $objListBox.Items.Add("Peer Distribution Point Status Task")
        [void] $objListBox.Items.Add("Peer Distribution Point Provisioning Status Task")
        [void] $objListBox.Items.Add("Compliance Interval Enforcement")
        [void] $objListBox.Items.Add("Software Updates Deployment Agent Assignment Evaluation Cycle")
        [void] $objListBox.Items.Add("Send Unsent State Messages")
        [void] $objListBox.Items.Add("State Message Manager Task")
        [void] $objListBox.Items.Add("Force Software Update Scan")
        [void] $objListBox.Items.Add("Software Update Store")
        [void] $objListBox.Items.Add("AMT Provision Cycle")        

        $objForm.Controls.Add($objListBox)

        $objForm.TopMost = $True

        $objForm.Add_Shown({$objForm.Activate()})
        [void] $objForm.ShowDialog()

        $x
        ##End of Menu##
        
        If ($x -eq $NULL) {Write-Host "No Function Selected...Exiting"}
        
        If ($x -ne $NULL) {             
        $starttimer = Get-Date
        $paraArray = $paraArray | Sort-Object
        Write-Host ""        
        "Invoking SCCM Function '$x' for {0} Systems." -f ($paraArray | Measure-Object).count
        Write-Host ""
        
        #CCM ScheduleIDs for Trigger Method from SCCM 2007 SDK (TriggerSchedule Method in Class SMS_Client)
        #Query to enum: Get-WmiObject CCM_Scheduler_ScheduledMessage -Namespace root\ccm\policy\machine\actualconfig | Select-Object ScheduledMessageID, TargetEndPoint | Where-Object {$_.TargetEndPoint -ne "direct:execmgr"}
        
        $WMITrigger = $NULL
        If ($x -eq "Hardware Inventory Collection Task"){$WMITrigger = "{00000000-0000-0000-0000-000000000001}"}
        If ($x -eq "Software Inventory Collection Task"){$WMITrigger = "{00000000-0000-0000-0000-000000000002}"}
        If ($x -eq "Heartbeat Discovery Cycle"){$WMITrigger = "{00000000-0000-0000-0000-000000000003}"}
        If ($x -eq "Software Inventory File Collection Task"){$WMITrigger = "{00000000-0000-0000-0000-000000000010}"}
        If ($x -eq "Machine Policy Assignments Request"){$WMITrigger = "{00000000-0000-0000-0000-000000000021}"}
        If ($x -eq "Evaluate Machine Policy Assignments"){$WMITrigger = "{00000000-0000-0000-0000-000000000022}"}
        If ($x -eq "Refresh Default MP Task"){$WMITrigger = "{00000000-0000-0000-0000-000000000023}"}
        If ($x -eq "Refresh Location Services Task"){$WMITrigger = "{00000000-0000-0000-0000-000000000024}"}
        If ($x -eq "Location Services Cleanup Task"){$WMITrigger = "{00000000-0000-0000-0000-000000000025}"}
        If ($x -eq "Software Metering Report Cycle"){$WMITrigger = "{00000000-0000-0000-0000-000000000031}"}
        If ($x -eq "Source Update Manage Update Cycle"){$WMITrigger = "{00000000-0000-0000-0000-000000000032}"}
        If ($x -eq "Policy Agent Cleanup Cycle"){$WMITrigger = "{00000000-0000-0000-0000-000000000040}"}
        If ($x -eq "Validate Machine Policy Assignments"){$WMITrigger = "{00000000-0000-0000-0000-000000000042}"}
        If ($x -eq "Certificate Maintenance Cycle"){$WMITrigger = "{00000000-0000-0000-0000-000000000051}"}
        If ($x -eq "Peer Distribution Point Status Task"){$WMITrigger = "{00000000-0000-0000-0000-000000000061}"}
        If ($x -eq "Peer Distribution Point Provisioning Status Task"){$WMITrigger = "{00000000-0000-0000-0000-000000000062}"}
        If ($x -eq "Compliance Interval Enforcement"){$WMITrigger = "{00000000-0000-0000-0000-000000000071}"} #**does not work**
        If ($x -eq "Software Updates Deployment Agent Assignment Evaluation Cycle"){$WMITrigger = "{00000000-0000-0000-0000-000000000108}"}
        If ($x -eq "Send Unsent State Messages"){$WMITrigger = "{00000000-0000-0000-0000-000000000111}"}
        If ($x -eq "State Message Manager Task"){$WMITrigger = "{00000000-0000-0000-0000-000000000112}"}
        If ($x -eq "Force Software Update Scan"){$WMITrigger = "{00000000-0000-0000-0000-000000000113}"}
        If ($x -eq "Software Update Store"){$WMITrigger = "{00000000-0000-0000-0000-000000000114}"}
        If ($x -eq "AMT Provision Cycle"){$WMITrigger = "{00000000-0000-0000-0000-000000000120}"}
                
        $paraArray | foreach {       
        if ($_.length -gt 0)
		    {			     
                 If ($x -eq "Set User Defined Schedule"){
                 "Invoking SCCM Function: '$x' on {0}.  Configuring for: $SCCMDesiredInstallDay @ $SCCMDesiredInstallHour" -f $_
                 start-job -scriptblock {$wmi = Get-WmiObject Win32_OperatingSystem -comp $args[0] | Select OSArchitecture; if ($wmi.OSArchitecture -eq '64-bit') {$session = New-PSSession -cn ($args[0] +"."+ $args[1])  -ConfigurationName Microsoft.PowerShell32}; if ($wmi.OSArchitecture -eq '32-bit') {$session = New-PSSession -cn ($args[0] +"."+ $args[1]) }; Invoke-Command -Session $session -ScriptBlock {$SCCMUpdate = New-Object -ComObject 'UDA.CCMUpdatesDeployment'; $SCCMDayHashTable=$args[0];$SCCMHourHashTable=$args[1];$SCCMDesiredInstallDay=$args[2];$SCCMDesiredInstallHour=$args[3]; $SCCMUpdate.SetUserDefinedSchedule(($SCCMDayHashTable.GetEnumerator() | ?{$_.Value -eq $SCCMDesiredInstallDay}).name, ($SCCMHourHashTable.GetEnumerator() | ?{$_.Value -eq $SCCMDesiredInstallHour}).name)} -Args $args[3], $args[4], $args[5], $args[6]} -name("SCCMInvoke-" + $_) -argumentlist $_ , $DNSDomain, $cred, $SCCMDayHashTable, $SCCMHourHashTable, $SCCMDesiredInstallDay, $SCCMDesiredInstallHour | Out-Null
                 }
                 If ($x -eq "Install Pending Updates"){
                 "Invoking SCCM Function: '$x' on {0} " -f $_
                 start-job -scriptblock {$wmi = Get-WmiObject Win32_OperatingSystem -comp $args[0] | Select OSArchitecture; if ($wmi.OSArchitecture -eq '64-bit') {$session = New-PSSession -cn ($args[0] +"."+ $args[1])  -ConfigurationName Microsoft.PowerShell32}; if ($wmi.OSArchitecture -eq '32-bit') {$session = New-PSSession -cn ($args[0] +"."+ $args[1]) }; Invoke-Command -Session $session -ScriptBlock {$SCCMUpdate = New-Object -ComObject 'UDA.CCMUpdatesDeployment';$hostname=hostname;[ref]$Progress=$Null;$updates = $SCCMUpdate.EnumerateUpdates(2,1,$Progress); if($Progress.Value -eq 0) {$UpdateCount=$updates.GetCount(); if($UpdateCount -ne 0) {[string[]]$UpdateIDs=For($i=0;$i -lt $UpdateCount;$i++){$updates.GetUpdate($i).GetID()};$SCCMUpdate.InstallUpdates($UpdateIDs,0,$args[0])};Write-Host;Write-Host "$hostname : Installing the following Updates ($UpdateCount in Total): $UpdateIDs"} if($Progress.Value -ne 0){Write-Host;Write-Host "$hostname : Already Currently Installing (recommend check for Hang)"} } -Args $args[3] } -name("SCCMInvoke-" + $_) -argumentlist $_ , $DNSDomain, $cred, $SCCMInstallOptions | Out-Null
                 }
                 If ($x -eq "Initate 3.4 Migration"){
                 "Invoking SCCM Function: '$x' on {0}. " -f $_
                 start-job -scriptblock {$wmi = Get-WmiObject Win32_OperatingSystem -comp $args[0] | Select OSArchitecture; if ($wmi.OSArchitecture -eq '64-bit') {$session = New-PSSession -cn ($args[0] +"."+ $args[1])  -ConfigurationName Microsoft.PowerShell32}; if ($wmi.OSArchitecture -eq '32-bit') {$session = New-PSSession -cn ($args[0] +"."+ $args[1]) }; Invoke-Command -Session $session -ScriptBlock {$AdvID = 'ANG202C0';$strQuery = "Select * From CCM_Scheduler_ScheduledMessage Where ScheduledMessageID like '" + $AdvID + "%'";$objSMSchID = Get-WmiObject -Query $strQuery -Namespace root\ccm\policy\machine\actualconfig;foreach($instance in $objSMSchID){$strScheduleID=$instance.ScheduledMessageID};$strQuery = "Select * From CCM_SoftwareDistribution Where ADV_AdvertisementID = '" + $AdvID + "'";Get-WmiObject -Query $strQuery -Namespace root\ccm\policy\machine\actualconfig | ForEach-Object {$_.ADV_MandatoryAssignments='TRUE';$_.ADV_RepeatRunBehavior='RerunAlways';$_.Put()};$WMIPath='\\.\root\ccm:SMS_Client';$SMSwmi=[wmiclass]$WMIPath;$SMSwmi.TriggerSchedule($strScheduleID);Write-Host $strScheduleID initated on $args[0];} -Args $args[0]} -name("SCCMInvoke-" + $_) -argumentlist $_ , $DNSDomain, $cred | Out-Null
                 }
                 If ($WMITrigger -ne $NULL) {
                 Write-Host "Invoking SCCM Function: " $x - $WMITrigger "on" $_ 
                 Start-Job -ScriptBlock {$session = New-PSSession -cn ($args[0] +"."+ $args[1]) ; Invoke-Command -Session $session -ScriptBlock {([wmiclass]'root\ccm:SMS_Client').TriggerSchedule($args)} -Args $args[3] } -name("SCCMInvoke-" + $_) -argumentlist $_ , $DNSDomain, $cred, $WMITrigger | Out-Null
                 }
            }    		
		    while (((get-job | where-object { $_.Name -like "SCCMInvoke-*" -and $_.State -eq "Running" }) | measure).Count -gt $MaxConcurrentJobs)
		    {
                "{0} Concurrent jobs running, sleeping 5 seconds" -f $MaxConcurrentJobs
			    Start-Sleep -seconds 5
		    }
	    }
        while (((get-job | where-object { $_.Name -like "SCCMInvoke-*" -and $_.state -eq "Running" }) | measure).count -gt 0)
	    {
		  $jobcount = ((get-job | where-object { $_.Name -like "SCCMInvoke-*" -and $_.state -eq "Running" }) | measure).count
		  Write-Host "Waiting for $jobcount Jobs to Complete" 
		  Start-Sleep -seconds 5
          $Counter++
            if ($Counter -gt 40) {
                Write-Host "Exiting loop $jobCount Jobs did not complete"
                get-job | where-object { $_.Name -like "SCCMInvoke-*" -and $_.state -eq "Running" } | select Name
                break
            }
	     }
         
         $Global:SCCMResults=@()
         get-job | where { $_.Name -like "SCCMInvoke-*" -and $_.state -eq "Completed" } | % { $Global:SCCMResults += Receive-Job $_ ; Remove-Job $_ }
         $stoptimer = Get-Date
         Write-Host
         "Total Time for SCCM Invoke: {0} Minutes" -f [math]::round(($stoptimer - $starttimer).TotalMinutes , 2)
         Write-Host
         "Total Systems: {0} " -f ($paraArray | Measure-Object).count
          Return #$SCCMResults | Out-GridView
          }
    }    
        
        #invoke Update Deployment via PS Sessions
        #$session = New-PSSession -cn 110AOG0AISCD202 -Cred $cred; Invoke-Command -Session $session -ScriptBlock {([wmiclass]'root\ccm:SMS_Client').TriggerSchedule('{00000000-0000-0000-0000-000000000108}')}
        
    function Minion-Invoke-CMD-MultiThreaded
    {
        param(
            [parameter(Mandatory = $true)]
            $paraArray
        )
        Write-Host ""
        #if ($cred -eq $null) {Write-Host "No Credentials supplied, requesting presently..." ; $Global:cred = Minion-Get-Cred -Domain $DNSDomain} else {Write-Host "Credential check pass...proceeding"}
        Write-Host ""
        ##Menu##
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")

        $x=$NULL

        $objForm = New-Object System.Windows.Forms.Form
        $objForm.Text = "Select a SCCM function to Invoke"
        $objForm.Size = New-Object System.Drawing.Size(400,500)
        $objForm.StartPosition = "CenterScreen"

        $objForm.KeyPreview = $True

        $objForm.Add_KeyDown({if ($_.KeyCode -eq "Enter")
            {$x=$objListBox.SelectedItem;$objForm.Close()}})
        $objForm.Add_KeyDown({if ($_.KeyCode -eq "Escape")
            {$objForm.Close()}})
    
        $OKButton = New-Object System.Windows.Forms.Button
        $OKButton.Location = New-Object System.Drawing.Size(75,435)
        $OKButton.Size = New-Object System.Drawing.Size(75,23)
        $OKButton.Text = "OK"
        $OKButton.Add_Click({$x=$objListBox.SelectedItem;$objForm.Close()})
        $objForm.Controls.Add($OKButton)

        $CancelButton = New-Object System.Windows.Forms.Button
        $CancelButton.Location = New-Object System.Drawing.Size(150,435)
        $CancelButton.Size = New-Object System.Drawing.Size(75,23)
        $CancelButton.Text = "Cancel"
        $CancelButton.Add_Click({$objForm.Close()})
        $objForm.Controls.Add($CancelButton)

        $objLabel = New-Object System.Windows.Forms.Label
        $objLabel.Location = New-Object System.Drawing.Size(10,20)
        $objLabel.Size = New-Object System.Drawing.Size(280,20)
        $objLabel.Text = "Please select a Function:"
        $objForm.Controls.Add($objLabel)
 
        $objListBox = New-Object System.Windows.Forms.ListBox
        $objListBox.Location = New-Object System.Drawing.Size(10,40)
        $objListBox.Size = New-Object System.Drawing.Size(350,20)
        $objListBox.Height = 400

        [void] $objListBox.Items.Add("Joe Vande's Script")
        [void] $objListBox.Items.Add("Reboot - 10 minute countdown")
        [void] $objListBox.Items.Add("SMS Client: Uninstall")
        [void] $objListBox.Items.Add("SMS Client: Install")
        #[void] $objListBox.Items.Add("Initate 3.4 Migration")
        [void] $objListBox.Items.Add("Software Inventory File Collection Task")
        [void] $objListBox.Items.Add("Machine Policy Assignments Request")
        [void] $objListBox.Items.Add("Evaluate Machine Policy Assignments")
        [void] $objListBox.Items.Add("Refresh Default MP Task")
        [void] $objListBox.Items.Add("Refresh Location Services Task")
        [void] $objListBox.Items.Add("Location Services Cleanup Task")
        [void] $objListBox.Items.Add("Software Metering Report Cycle")
        [void] $objListBox.Items.Add("Source Update Manage Update Cycle")
        [void] $objListBox.Items.Add("Policy Agent Cleanup Cycle")
        [void] $objListBox.Items.Add("Validate Machine Policy Assignments")
        [void] $objListBox.Items.Add("Certificate Maintenance Cycle")
        [void] $objListBox.Items.Add("Peer Distribution Point Status Task")
        [void] $objListBox.Items.Add("Peer Distribution Point Provisioning Status Task")
        [void] $objListBox.Items.Add("Compliance Interval Enforcement")
        [void] $objListBox.Items.Add("Software Updates Deployment Agent Assignment Evaluation Cycle")
        [void] $objListBox.Items.Add("Send Unsent State Messages")
        [void] $objListBox.Items.Add("State Message Manager Task")
        [void] $objListBox.Items.Add("Force Software Update Scan")
        [void] $objListBox.Items.Add("Software Update Store")
        [void] $objListBox.Items.Add("AMT Provision Cycle")
        


        $objForm.Controls.Add($objListBox)

        $objForm.TopMost = $True

        $objForm.Add_Shown({$objForm.Activate()})
        [void] $objForm.ShowDialog()

        $x
        ##End of Menu##
        
        If ($x -eq $NULL) {Write-Host "No Function Selected...Exiting"}
        
        If ($x -ne $NULL) {     
        
        $starttimer = Get-Date
        $paraArray = $paraArray | Sort-Object
        
        $threadedCMD = $NULL
        If ($x -eq "Joe Vande's Script"){$threadedCMD = $joe_command}
        If ($x -eq "Reboot - 10 minute countdown"){$threadedCMD = $reboot_command}
        If ($x -eq "SMS Client: Uninstall"){$threadedCMD = "CCM_Client_Uninstall"}
        If ($x -eq "SMS Client: Install"){$threadedCMD = "CCM_Client_Install"}
        
        #$threadedCMD = $90MeterInstall_command
        #$threadedCMD=$CCM_Uninstall_command
        #$threadedCMD=$CCM_Install_command
        
        Write-Host ""        
        "Running Command: '" + $threadedCMD + "' on {0} Systems." -f ($paraArray | Measure-Object).count
        Write-Host ""
        "Utilizing Credentials: {0}" -f $cred.username
        Write-Host ""
        $paraArray | foreach {       
        if ($_.length -gt 0)
		    {
			     
                 "Invoking command on {0}" -f $_
                 start-job -scriptblock {$session = New-PSSession -cn ($args[0] +"."+ $args[1]); $args[3]| Invoke-Command -Session $session -ScriptBlock { cmd }}  -name("PSInvokeCMD-" + $_) -ArgumentList $_ , $DNSDomain, $cred, $threadedCMD | Out-Null
                 
            }    		
		    while (((get-job | where-object { $_.Name -like "PSInvokeCMD*" -and $_.State -eq "Running" }) | measure).Count -gt $MaxConcurrentJobs)
		    {
                "{0} Concurrent jobs running, sleeping 5 seconds" -f $MaxConcurrentJobs
			    Start-Sleep -seconds 5
		    }
	    }
        while (((get-job | where-object { $_.Name -like "PSInvokeCMD*" -and $_.state -eq "Running" }) | measure).count -gt 0)
	    {
		  $jobcount = ((get-job | where-object { $_.Name -like "PSInvokeCMD*" -and $_.state -eq "Running" }) | measure).count
		  Write-Host "Waiting for $jobcount Jobs to Complete"
		  Start-Sleep -seconds 5
          $Counter++
            if ($Counter -gt 40) {
                Write-Host "Exiting loop $jobCount Jobs did not complete"
                get-job | where-object { $_.Name -like "PSInvokeCMD*" -and $_.state -eq "Running" } | select Name
                break
            }
	     }
         $PSInvokeCMDResults = @()
	
	     #Import all job state into $PingResults Array
	     get-job | where { $_.Name -like "PSInvokeCMD*" -and $_.state -eq "Completed" } | % { $PSInvokeCMDResults += Receive-Job $_ ; Remove-Job $_ }
	     $stoptimer = Get-Date
         Write-Host
         "Total Time for Execution: {0} Minutes" -f [math]::round(($stoptimer - $starttimer).TotalMinutes , 2)
         Write-Host
         "Command: '" + $threadedCMD + " ' on {0} Systems." -f ($paraArray | Measure-Object).count
         Return
         }
    }
