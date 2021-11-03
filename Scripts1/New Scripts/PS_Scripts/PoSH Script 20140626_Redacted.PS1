$ADsPath = "LDAP://OU=Tyndall AFB Computers,OU=Tyndall AFB,OU=AFCONUSEAST,OU=Bases,DC=AREA52,DC=AFNOAPPS,DC=USAF,DC=MIL"
$strFilter = "(&(objectCategory=Computer))"
$showError=0
$starttimer = Get-Date

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
    $AllComputerNames = $AllComputerNames | Sort-Object
    
    #now ping each system - multi-threaded, max of $MaxConcurrentJobs at a time.
    $starttimer = Get-Date
    $SB={
        Param (
            [string]$Computer
        )
        $Pinger = Get-WmiObject -Query "Select * from Win32_PingStatus Where timeout=3000 and Address='$Computer'"
        $RunResult = Invoke-Expression $Pinger
        Return $RunResult
    }
        
    ForEach ($Computer in $AllComputerNames)
    {
    	If ($Computer.length -gt 0)
		{
			"Checking Availability of {0}" -f $Computer
            $Pinger2 = Get-WmiObject -Query "Select * from Win32_PingStatus Where timeout=3000 and Address='$args'"
            Start-Job -ScriptBlock {Invoke-Expression $Pinger2} -Name("Ping-" + $_) -argumentlist $_ | Out-Null
        }    		
		While (((Get-Job | Where-Object { $_.Name -like "Ping*" -and $_.State -eq "Running" }) | measure).Count -gt $MaxConcurrentJobs)
		{
            "{0} Concurrent jobs running, sleeping 5 seconds" -f $MaxConcurrentJobs
			Start-Sleep -seconds 5
		}
	}
    
    #wait for all jobs that start with 'Ping*' to complete...
	while (((get-job | where-object { $_.Name -like "Ping*" -and $_.state -eq "Running" }) | measure).count -gt 0)
	{
		$jobcount = ((get-job | where-object { $_.Name -like "Ping*" -and $_.state -eq "Running" }) | measure).count
		Write-Host "Waiting for $jobcount Jobs to Complete" 
		Start-Sleep -seconds 5
        $Counter++
        if ($Counter -gt 40) 
        {
            Write-Host "Exiting loop $jobCount Jobs did not complete"
            get-job | where-object {$_.Name -like "Ping*" -and $_.state -eq "Running"} | select Name
            break
        }
	}
	
	$PingResults = @()
	
	#Import all job state into $PingResults Array
	Get-Job | Where {$_.Name -like "Ping" -and $_.state -eq "Completed"} | % {$PingResults += Receive-Job $_ ; Remove-Job $_}
	
	$stoptimer = Get-Date
    
    #Display info, and display in GridView
    Write-Host
	Write-Host "Availability check complete!"
    Write-Host
    "Execution Time: {0} Minutes" -f [math]::round(($stoptimer - $starttimer).TotalMinutes , 2)
    "Total Systems : {0}" -f ($CollSystems | Measure-Object).count 
    Write-Host
    $DidnotResolveIP = if((($PingResults | where-object {$_.statusCode -eq $Null}) | measure-object ).Count -eq $null) {0} else {(($PingResults | where-object {$_.statusCode -eq $Null}) | measure-object ).Count}
    Write-Host $DidnotResolveIP "Systems Offline, and did not properly resolve an IP."
    $DidnotResolveMAC = if((($PingResults | where-object {$_.statusCode -eq 11003}) | measure-object ).Count -eq $null) {0} else {(($PingResults | where-object {$_.statusCode -eq 11003}) | measure-object ).Count}
    Write-Host $DidnotResolveMAC "Systems Offline that resolved an IP, but did not successfully ARP a MAC."    
    $notonline = if((($PingResults | where-object {$_.statusCode -eq 11010}) | measure-object ).Count -eq $null) {0} else {(($PingResults | where-object {$_.statusCode -eq 11010}) | measure-object ).Count}
    Write-Host $notonline "Systems Offline, resolved an IP and successfully ARP'd a MAC."
    Write-Host
    "Total Systems Offline: {0}" -f [math]::Round($DidnotResolveIP + $DidnotResolveMAC + $notonline , 1)
    "Total Systems Online : {0}" -f (($PingResults | where-object {$_.statusCode -eq 0}) | measure-object ).Count
    "{0}% of All Active Directory Systems @ '$ADsPath' are currently available." -f ([math]::Round((($PingResults | where-object {$_.statusCode -eq 0}) | measure-object ).Count/ $PingResults.count , 2)*100)
    $CombinedADandPingResults = $PingResults | foreach { $PingResult=$_;$match = $ADInfo | Where {$_.Name -eq $PingResult.Address} | select -unique;if($match -ne $NULL){""|select @{n="Name";e={$PingResult.Address}}, @{n="IPAddress";e={$PingResult.IPV4Address}}, @{n="StatusCode";e={$PingResult.StatusCode}}, @{n="Location";e={$match.description}}, @{n="LastLogon";e={$match.lastlogon}},@{n="LastLogonTimeStamp";e={$match.LastLogonTimeStamp}}, @{n="LogonCount";e={$match.logoncount}},@{n="PasswordChanged";e={$match.pwdLastSet}}, @{n="Created";e={$match.Created}}}}
    $CombinedADandPingResults | Sort-Object StatusCode | Out-GridView
    $PingResults | select Address, IPV4Address, StatusCode  | Sort-Object StatusCode | foreach {$_} 
    
    $Awake += ($PingResults | where-object {$_.statusCode -eq 0} | Select -ExpandProperty Address)
    $Target = $Awake | Where {$Excluded -notcontains $_}
    
    #$cred = Minion-Get-Cred -Domain $DNSDomain
    #Minion-Enable-PSRemoting-Server -computername $computername
    #$session = New-PSSession -cn ($computername +"."+ $DNSDomain) -Cred $cred
    #$session = New-PSSession -cn ($computername +"."+ $DNSDomain) -Cred $cred -Authentication CredSSP
    #Invoke-Command -Session $session -ScriptBlock { Test-Path \\110fw-fs-01\ULI }
    #$joe_command | Invoke-Command -Session $session -ScriptBlock { cmd }
    #Minion-Disable-PSRemoting-Server -computername $computername
    #$ADPE411 = $ADPEResults | Select CSName,BuildNumber,SPVersion,BIOSSerial | Where {$_.CSName -ne $Null}
    #$Year = Get-Date -UFormat "%Y - "; $JulianDate = (Get-Date).dayofyear; $TCNODay = $Year + $JulianDate; $TCNODay; Get-Date -UFormat "Zulu offset: %Z "
    #Get-Job | Remove-Job -Force
    
    function Minion-Get-Cred
    {
        param(
            [parameter(Mandatory = $true)]
            $Domain
        )

        $Global:cred = Get-Credential -Credential ($Domain + "\")
        Return $cred
    }
    
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
    #Initialization Script for MultiThreaded Start-Job -End
        
        "Enabling PSRemoting on {0} Systems." -f ($paraArray | Measure-Object).count
        $paraArray | foreach {       
        if ($_.length -gt 0)
		    {
			     "Enabling PS Remoting on {0}" -f $_
                 start-job -scriptblock {Minion-Enable-PSRemoting-Server -computername $args}  -InitializationScript $iniScript -name("PSRemotingEnable-" + $_) -ArgumentList $_ | Out-Null
                 $PSRemotingStatuses[$_]="**On**"
            }    		
		    while (((get-job | where-object { $_.Name -like "PSRemotingEnable*" -and $_.State -eq "Running" }) | measure).Count -gt $MaxConcurrentJobs)
		    {
                "{0} Concurrent jobs running, sleeping 5 seconds" -f $MaxConcurrentJobs
			    Start-Sleep -seconds 5
		    }
	    }
        while (((get-job | where-object { $_.Name -like "PSRemotingEnable*" -and $_.state -eq "Running" }) | measure).count -gt 0)
	    {
		  $jobcount = ((get-job | where-object { $_.Name -like "PSRemotingEnable*" -and $_.state -eq "Running" }) | measure).count
		  Write-Host "Waiting for $jobcount Jobs to Complete" 
		  Start-Sleep -seconds 5
          $Counter++
            if ($Counter -gt 40) {
                Write-Host "Exiting loop $jobCount Jobs did not complete"
                get-job | where-object { $_.Name -like "PSRemotingEnable*" -and $_.state -eq "Running" } | select Name
                break
            }
	     }
         #$PSRemotingEnableResults = @()
	
	     #Import all job state into $PingResults Array
	     get-job | where { $_.Name -like "PSRemotingEnable*" -and $_.state -eq "Completed" } | % { $PSRemotingEnableResults += Receive-Job $_ ; Remove-Job $_ }
	     $stoptimer = Get-Date

         Return $PSRemotingStatuses.GetEnumerator() | Sort-Object Name
    }
    
    function Minion-Disable-PSRemoting-Server-MultiThreaded
    {
        param(
            [parameter(Mandatory = $true)]
            $paraArray
        )
        $starttimer = Get-Date
        $paraArray = $paraArray | Sort-Object
        
    #Initialization Script for MultiThreaded Start-Job -Begin
        $iniScript = { function Minion-Disable-PSRemoting-Server
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
    }
    #Initialization Script for MultiThreaded Start-Job -End
        
        "Disabling PSRemoting on {0} Systems." -f ($paraArray | Measure-Object).count
        $paraArray | foreach {       
        if ($_.length -gt 0)
		    {
			     "Disabling PS Remoting on {0}" -f $_
                 start-job -scriptblock {Minion-Disable-PSRemoting-Server -computername $args}  -InitializationScript $iniScript -name("PSRemotingDisable-" + $_) -ArgumentList $_ | Out-Null
                 $PSRemotingStatuses[$_]="**Off**"
            }    		
		    while (((get-job | where-object { $_.Name -like "PSRemotingDisable*" -and $_.State -eq "Running" }) | measure).Count -gt $MaxConcurrentJobs)
		    {
                "{0} Concurrent jobs running, sleeping 5 seconds" -f $MaxConcurrentJobs
			    Start-Sleep -seconds 5
		    }
	    }
        while (((get-job | where-object { $_.Name -like "PSRemotingDisable*" -and $_.state -eq "Running" }) | measure).count -gt 0)
	    {
		  $jobcount = ((get-job | where-object { $_.Name -like "PSRemotingDisable*" -and $_.state -eq "Running" }) | measure).count
		  Write-Host "Waiting for $jobcount Jobs to Complete" 
		  Start-Sleep -seconds 5
          $Counter++
            if ($Counter -gt 40) {
                Write-Host "Exiting loop $jobCount Jobs did not complete"
                get-job | where-object { $_.Name -like "PSRemotingDisable*" -and $_.state -eq "Running" } | select Name
                break
            }
	     }
         $PSRemotingDisableResults = @()
	
	     #Import all job state into $PingResults Array
	     get-job | where { $_.Name -like "PSRemotingDisable*" -and $_.state -eq "Completed" } | % { $PSRemotingDisableResults += Receive-Job $_ ; Remove-Job $_ }
	     $stoptimer = Get-Date

         Return
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
#Write-Host 'Easy Button';Minion-Enable-PSRemoting-Server-MultiThreaded -paraArray $Awake;Minion-Get-ADPE -paraArray $Awake;Minion-Get-SCCMStatus -paraArray $Awake;Minion-Get-IAVM -paraArray $Awake;Minion-Get-Progs -paraArray $Awake;$Java=$ProgsResults|Where-Object{$_.DisplayName -like '*Java 7*' -or $_.DisplayName -like '*Java(TM)*'}; $AdobeAcrobat=$ProgsResults|Where-Object{$_.DisplayName -like '*Acrobat*' -or $_.Displayname -like '*Adobe Reader*'} ; $Chrome=$ProgsResults|Where-Object{$_.DisplayName -like '*Chrome*'} ; $FireFox=$ProgsResults|Where-Object{$_.DisplayName -like '*FireFox*'}; $ThunderBird=$ProgsResults|Where-Object{$_.DIsplayName -like '*Thunderbird*'}; $Flash=$ProgsResults|Where-Object{$_.DisplayName -like '*Adobe Flash*'}; $Shockwave=$ProgsResults|Where-Object{$_.DisplayName -like '*Adobe Shockwave*'}; $IBMForms=$ProgsResults|Where-Object{$_.DisplayName -like '*IBM Forms*'}; $Java|Out-GridView;$AdobeAcrobat|Out-GridView;$Chrome|Out-GridView;$FireFox|Out-GridView;$ThunderBird|Out-GridView;$Flash|Out-GridView;$Shockwave|Out-GridView;$IBMForms|Out-GridView;Write-Host '...Well that was easy'
# SIG # Begin signature block
# MIIacQYJKoZIhvcNAQcCoIIaYjCCGl4CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUW3nbejCJ5YG+4HNJpuKScMUH
# rvmgghXrMIIDcDCCAligAwIBAgIBBTANBgkqhkiG9w0BAQUFADBbMQswCQYDVQQG
# EwJVUzEYMBYGA1UEChMPVS5TLiBHb3Zlcm5tZW50MQwwCgYDVQQLEwNEb0QxDDAK
# BgNVBAsTA1BLSTEWMBQGA1UEAxMNRG9EIFJvb3QgQ0EgMjAeFw0wNDEyMTMxNTAw
# MTBaFw0yOTEyMDUxNTAwMTBaMFsxCzAJBgNVBAYTAlVTMRgwFgYDVQQKEw9VLlMu
# IEdvdmVybm1lbnQxDDAKBgNVBAsTA0RvRDEMMAoGA1UECxMDUEtJMRYwFAYDVQQD
# Ew1Eb0QgUm9vdCBDQSAyMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEA
# wCzB9o07rP8/PNZxvrh0IgfscEEV/KtA4weqwcPYn/7aTDq/P8jYKHtLNgHArEUl
# w9IOCo+FGGQQPRoTcCpvjtfcjZOzQQ84Ic2tq8I9KgXTVxE3Dc2MUfmT48xGSSGO
# FLTNyxQ+OM1yMe6rEvJl6jQuVl3/7mN1y226kTT8nvP0LRy+UMRC31mI/2qz+qhs
# PctWcXEFlrufgOWARVlnQbDrw61gpIB1BhecDvRD4JkOG/t/9bPMsoGCsf0ywbi+
# QaRktWA6WlEwjM7eQSwZR1xJEGS5dKmHQa99brrBuKG/ZTE6BGf5tbuOkooAY7ix
# 5ow4X4P/UNU7ol1rshDMYwIDAQABoz8wPTAdBgNVHQ4EFgQUSXS7DF66ev4CVO97
# oMaVxgmAcJYwCwYDVR0PBAQDAgGGMA8GA1UdEwEB/wQFMAMBAf8wDQYJKoZIhvcN
# AQEFBQADggEBAJiRjT+JyLv1wGlzKTs1rLqzCHY9cAmS6YREIQF9FHYb7lFsHY0V
# Ny17MWn0mkS4r0bMNPojywMnGdKDIXUr5+AbmSbchECV6KjSzPZYXGbvP0qXEIId
# ugqi3VsGK52nZE7rLgE1pLQ/E61V5NVzqGmbEfGY8jEeb0DU+HifjpGgb3AEkGaq
# BivO4XqStX3h4NGW56E6LcyxnR8FRO2HmdNNGnA5wQQM5X7Z8a/XIA7xInolpHOZ
# zD+kByeWqKKV7YK5FtOeC4fCwfKI9WLfaN/HvGlR7bFc3FRUKQ8JOZqsA8HbDE2u
# bwp6Fknxv5HSOJTT9pUst2zJQraNypCNhdkwggPuMIIDV6ADAgECAhB+k+v7fMZO
# WepLmnfUBvw7MA0GCSqGSIb3DQEBBQUAMIGLMQswCQYDVQQGEwJaQTEVMBMGA1UE
# CBMMV2VzdGVybiBDYXBlMRQwEgYDVQQHEwtEdXJiYW52aWxsZTEPMA0GA1UEChMG
# VGhhd3RlMR0wGwYDVQQLExRUaGF3dGUgQ2VydGlmaWNhdGlvbjEfMB0GA1UEAxMW
# VGhhd3RlIFRpbWVzdGFtcGluZyBDQTAeFw0xMjEyMjEwMDAwMDBaFw0yMDEyMzAy
# MzU5NTlaMF4xCzAJBgNVBAYTAlVTMR0wGwYDVQQKExRTeW1hbnRlYyBDb3Jwb3Jh
# dGlvbjEwMC4GA1UEAxMnU3ltYW50ZWMgVGltZSBTdGFtcGluZyBTZXJ2aWNlcyBD
# QSAtIEcyMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAsayzSVRLlxwS
# CtgleZEiVypv3LgmxENza8K/LlBa+xTCdo5DASVDtKHiRfTot3vDdMwi17SUAAL3
# Te2/tLdEJGvNX0U70UTOQxJzF4KLabQry5kerHIbJk1xH7Ex3ftRYQJTpqr1SSwF
# eEWlL4nO55nn/oziVz89xpLcSvh7M+R5CvvwdYhBnP/FA1GZqtdsn5Nph2Upg4XC
# YBTEyMk7FNrAgfAfDXTekiKryvf7dHwn5vdKG3+nw54trorqpuaqJxZ9YfeYcRG8
# 4lChS+Vd+uUOpyyfqmUg09iW6Mh8pU5IRP8Z4kQHkgvXaISAXWp4ZEXNYEZ+VMET
# fMV58cnBcQIDAQABo4H6MIH3MB0GA1UdDgQWBBRfmvVuXMzMdJrU3X3vP9vsTIAu
# 3TAyBggrBgEFBQcBAQQmMCQwIgYIKwYBBQUHMAGGFmh0dHA6Ly9vY3NwLnRoYXd0
# ZS5jb20wEgYDVR0TAQH/BAgwBgEB/wIBADA/BgNVHR8EODA2MDSgMqAwhi5odHRw
# Oi8vY3JsLnRoYXd0ZS5jb20vVGhhd3RlVGltZXN0YW1waW5nQ0EuY3JsMBMGA1Ud
# JQQMMAoGCCsGAQUFBwMIMA4GA1UdDwEB/wQEAwIBBjAoBgNVHREEITAfpB0wGzEZ
# MBcGA1UEAxMQVGltZVN0YW1wLTIwNDgtMTANBgkqhkiG9w0BAQUFAAOBgQADCZuP
# ee9/WTCq72i1+uMJHbtPggZdN1+mUp8WjeockglEbvVt61h8MOj5aY0jcwsSb0ep
# rjkR+Cqxm7Aaw47rWZYArc4MTbLQMaYIXCp6/OJ6HVdMqGUY6XlAYiWWbsfHN2qD
# IQiOQerd2Vc/HXdJhyoWBl6mOGoiEqNRGYN+tjCCBIowggNyoAMCAQICAyThXjAN
# BgkqhkiG9w0BAQUFADBXMQswCQYDVQQGEwJVUzEYMBYGA1UEChMPVS5TLiBHb3Zl
# cm5tZW50MQwwCgYDVQQLEwNEb0QxDDAKBgNVBAsTA1BLSTESMBAGA1UEAxMJRE9E
# IENBLTMwMB4XDTEyMTEyMTAwMDAwMFoXDTE1MTEyMDIzNTk1OVowgYExCzAJBgNV
# BAYTAlVTMRgwFgYDVQQKEw9VLlMuIEdvdmVybm1lbnQxDDAKBgNVBAsTA0RvRDEM
# MAoGA1UECxMDUEtJMQ0wCwYDVQQLEwRVU0FGMS0wKwYDVQQDEyRQQVVMU09OLktF
# VklOLkNIUklTVE9QSEVSLjEyOTAzMjcyMDkwggEiMA0GCSqGSIb3DQEBAQUAA4IB
# DwAwggEKAoIBAQDDQyktZrxfWiu+yHbFeLApkeDWuIO0JGBTWynEeMpoHc7SumE8
# 4B0R3rp03N/U0SghKpe7QPIrULyGnMmkHvFnRr0sLfU5408LTuLZcsQbQpvWqwNO
# do7Qcj2KF5c9XIWgS5FowC0oBfYwzj7eRD06nrBIxiyjcDDxGA+htMA2zDMTGtNP
# x8NRY6ct76hhk9KiaKJxbR2GpjpD0oXqn7uiDP+ZYD0mqLFp4AnsgNwt2krq7v56
# 0Kwxs2IwaYL6pCRCE+Fnniwp/mheJL9O7j4XOWxt6ahBNUlTozQAx2HXrYVsIX2Y
# dwIVH8KIWDcRUcWgT7qBRmRh2KQQV8j3CSc7AgMBAAGjggEyMIIBLjAfBgNVHSME
# GDAWgBQITtWkPCoEm5MbtwQIjnS5BnwNozA1BgNVHR8ELjAsMCqgKKAmhiRodHRw
# Oi8vY3JsLmRpc2EubWlsL2NybC9ET0RDQV8zMC5jcmwwDgYDVR0PAQH/BAQDAgbA
# MCMGA1UdIAQcMBowCwYJYIZIAWUCAQsJMAsGCWCGSAFlAgELEzAdBgNVHQ4EFgQU
# CMN/CfXor0kvPE6gRQ2aBrgDOygwYwYIKwYBBQUHAQEEVzBVMDEGCCsGAQUFBzAC
# hiVodHRwOi8vY3JsLmRpc2EubWlsL3NpZ24vRE9EQ0FfMzAuY2VyMCAGCCsGAQUF
# BzABhhRodHRwOi8vb2NzcC5kaXNhLm1pbDAbBgNVHQkEFDASMBAGCCsGAQUFBwkE
# MQQTAlVTMA0GCSqGSIb3DQEBBQUAA4IBAQBddHGEDnkZnP4oGNwP99KTYhAEDTmJ
# iZPa872iK5WlqJFX4iSP0t7eop3XhfEMGz/BoKIsoK2L8T+FAsFc1WA27CT4JLNR
# Dk8aZnMQg5Zq/YsQlNe20EMrx4YSdP/rnduKmmeA3APnyoGPT+Ecy4RN90JnEmNZ
# qntIja4ukY0c9PlwKmrKydHHqnu1/J7pvidOuQ36M/wnmeKZpFNORcxxe/sJpjzl
# HrHNSLxIdUVXDpCcH2vmAaCXpxiv2wG9KnzomCgQEBYVOU+i3YeXI2bpPJCt5eiT
# O8z+oYtMABMxYfDtWaUcuECUQw3HZ3xAz4jOXDc3TXl97zR+JVqTenk8MIIEozCC
# A4ugAwIBAgIQDs/0OMj+vzVuBNhqmBsaUDANBgkqhkiG9w0BAQUFADBeMQswCQYD
# VQQGEwJVUzEdMBsGA1UEChMUU3ltYW50ZWMgQ29ycG9yYXRpb24xMDAuBgNVBAMT
# J1N5bWFudGVjIFRpbWUgU3RhbXBpbmcgU2VydmljZXMgQ0EgLSBHMjAeFw0xMjEw
# MTgwMDAwMDBaFw0yMDEyMjkyMzU5NTlaMGIxCzAJBgNVBAYTAlVTMR0wGwYDVQQK
# ExRTeW1hbnRlYyBDb3Jwb3JhdGlvbjE0MDIGA1UEAxMrU3ltYW50ZWMgVGltZSBT
# dGFtcGluZyBTZXJ2aWNlcyBTaWduZXIgLSBHNDCCASIwDQYJKoZIhvcNAQEBBQAD
# ggEPADCCAQoCggEBAKJjCzlEuLsjp0RJuw7/ofBhClOTsJjbrSwPSsVu/4Y8U1UP
# Fc4EPyv9qZaW2b5heQtbyUyGduXgQ0sile7CK0PBn9hotI5AT+6FOLkRxSPyZFjw
# FTJvTlehroikAtcqHs1L4d1j1ReJMluwXplaqJ0oUA4X7pbbYTtFUR3PElYLkkf8
# q672Zj1HrHBy55LnX80QucSDZJQZvSWA4ejSIqXQugJ6oXeTW2XD7hd0vEGGKtwI
# TIySjJEtnndEH2jWqHR32w5bMotWizO92WPISZ06xcXqMwvS8aMb9Iu+2bNXizve
# BKd6IrIkri7HcMW+ToMmCPsLvalPmQjhEChyqs0CAwEAAaOCAVcwggFTMAwGA1Ud
# EwEB/wQCMAAwFgYDVR0lAQH/BAwwCgYIKwYBBQUHAwgwDgYDVR0PAQH/BAQDAgeA
# MHMGCCsGAQUFBwEBBGcwZTAqBggrBgEFBQcwAYYeaHR0cDovL3RzLW9jc3Aud3Mu
# c3ltYW50ZWMuY29tMDcGCCsGAQUFBzAChitodHRwOi8vdHMtYWlhLndzLnN5bWFu
# dGVjLmNvbS90c3MtY2EtZzIuY2VyMDwGA1UdHwQ1MDMwMaAvoC2GK2h0dHA6Ly90
# cy1jcmwud3Muc3ltYW50ZWMuY29tL3Rzcy1jYS1nMi5jcmwwKAYDVR0RBCEwH6Qd
# MBsxGTAXBgNVBAMTEFRpbWVTdGFtcC0yMDQ4LTIwHQYDVR0OBBYEFEbGaaMOShQe
# 1UzaUmMXP142vA3mMB8GA1UdIwQYMBaAFF+a9W5czMx0mtTdfe8/2+xMgC7dMA0G
# CSqGSIb3DQEBBQUAA4IBAQB4O7SRKgBM8I9iMDd4o4QnB28Yst4l3KDUlAOqhk4l
# n5pAAxzdzuN5yyFoBtq2MrRtv/QsJmMz5ElkbQ3mw2cO9wWkNWx8iRbG6bLfsund
# IMZxD82VdNy2XN69Nx9DeOZ4tc0oBCCjqvFLxIgpkQ6A0RH83Vx2bk9eDkVGQW4N
# sOo4mrE62glxEPwcebSAe6xp9P2ctgwWK/F/Wwk9m1viFsoTgW0ALjgNqCmPLOGy
# 9FqpAa8VnCwvSRvbIrvD/niUUcOGsYKIXfA9tFGheTMrLnu53CAJE3Hrahlbz+il
# MFcsiUk/uc9/yb8+ImhjU5q9aXSsxR08f5Lgw7wc2AR1MIIFTDCCBDSgAwIBAgIC
# AbUwDQYJKoZIhvcNAQEFBQAwWzELMAkGA1UEBhMCVVMxGDAWBgNVBAoTD1UuUy4g
# R292ZXJubWVudDEMMAoGA1UECxMDRG9EMQwwCgYDVQQLEwNQS0kxFjAUBgNVBAMT
# DURvRCBSb290IENBIDIwHhcNMTEwOTA4MTU1OTI0WhcNMTcwOTA4MTU1OTI0WjBX
# MQswCQYDVQQGEwJVUzEYMBYGA1UEChMPVS5TLiBHb3Zlcm5tZW50MQwwCgYDVQQL
# EwNEb0QxDDAKBgNVBAsTA1BLSTESMBAGA1UEAxMJRE9EIENBLTMwMIIBIjANBgkq
# hkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAzTyTZyYPHuaBXu6fzvlQHt1iohWEJeV3
# VsJTx6DgyUJTKuZOZ1I+cF3GaLgVZcjddtCy1ZrJqizjxAkiBPd9iaSI2cKD7Fl7
# SRDvmo3Ihvlz3fIOYHqc2Y9Pd4N4DEtMLd7tn7GvHEMyrLDQODpUniYPFEuNwW71
# JpUkN4ft7eDD1e/A8A119W+avv1kPCoirzgSK3MtDQl+Eer8azJzTVzEWRfaxFmB
# BgS2CwLQZ7OWnHkTQxUkXsSV/VDRXgieH7ShlpI5K2isvYw+hokuPrbrReC8HJsr
# C3jvbfEaYN3mR/h19PLKRKj7gFngUWOFC7b7Fizj8/9v92q+m801gQIDAQABo4IC
# HDCCAhgwDgYDVR0PAQH/BAQDAgGGMB8GA1UdIwQYMBaAFEl0uwxeunr+AlTve6DG
# lcYJgHCWMB0GA1UdDgQWBBQITtWkPCoEm5MbtwQIjnS5BnwNozASBgNVHRMBAf8E
# CDAGAQH/AgEAMAwGA1UdJAQFMAOAAQAwZgYDVR0gBF8wXTALBglghkgBZQIBCwUw
# CwYJYIZIAWUCAQsJMAsGCWCGSAFlAgELETALBglghkgBZQIBCxIwCwYJYIZIAWUC
# AQsTMAwGCmCGSAFlAwIBAxowDAYKYIZIAWUDAgEDGzA3BgNVHR8EMDAuMCygKqAo
# hiZodHRwOi8vY3JsLmRpc2EubWlsL2NybC9ET0RST09UQ0EyLmNybDCCAQEGCCsG
# AQUFBwEBBIH0MIHxMDoGCCsGAQUFBzAChi5odHRwOi8vY3JsLmRpc2EubWlsL2lz
# c3VlZHRvL0RPRFJPT1RDQTJfSVQucDdjMCAGCCsGAQUFBzABhhRodHRwOi8vb2Nz
# cC5kaXNhLm1pbDCBkAYIKwYBBQUHMAKGgYNsZGFwOi8vY3JsLmdkcy5kaXNhLm1p
# bC9jbiUzZERvRCUyMFJvb3QlMjBDQSUyMDIlMmNvdSUzZFBLSSUyY291JTNkRG9E
# JTJjbyUzZFUuUy4lMjBHb3Zlcm5tZW50JTJjYyUzZFVTP2Nyb3NzQ2VydGlmaWNh
# dGVQYWlyO2JpbmFyeTANBgkqhkiG9w0BAQUFAAOCAQEAmffXrLElu68fAzW/Vnv1
# oWCm2pTuj93MMtE1DZ/lXqZ0En8BKlozIcDXBsq/3RtmVE8CVfym32gX0r/0XWuO
# +chz21tUOt294WnZ+pHbKloPx46INQgjq2Rn298fa/yOX3Kfl4GHgeWlIX3YT/4x
# m6F5pCZUQfBFkK9fQsEelof5z8ekGkRTkRE00IBktNkT1i0OiMepsSAkVnwH+8R7
# 9PmcerUORLcyVzpNg5HEdRiUls9f9m82K65zGfjg/GnOhn//QiE++TjDXnqZKN6Y
# LLCciBCyNB6qCArLTgHFZOtNpafzCD0LenU6lkr3/c8cr3JMcULZ/iO5WrStVwX9
# JDGCA/AwggPsAgEBMF4wVzELMAkGA1UEBhMCVVMxGDAWBgNVBAoTD1UuUy4gR292
# ZXJubWVudDEMMAoGA1UECxMDRG9EMQwwCgYDVQQLEwNQS0kxEjAQBgNVBAMTCURP
# RCBDQS0zMAIDJOFeMAkGBSsOAwIaBQCgWjAYBgorBgEEAYI3AgEMMQowCKACgACh
# AoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMCMGCSqGSIb3DQEJBDEWBBQp
# cu0f8CDB9TI47JKNAIbZxiKm3TANBgkqhkiG9w0BAQEFAASCAQAJnm3hnZeeIj/7
# vBItIgOfAXmoxrnweEygK1fLbNIHXblWN1C7ZsFYk2E+up5QLI+8Wm48T4BNqKRk
# EukCJyn5ZrOOkaOtlZvS+rhPEG3iunahrCbPwP8GlFORjhgrAxtRBlKY8E9otPnk
# jQOFnjX2ycJu28TFY5xHZJM54Uo/twxwG7lmShYunc24PyKer36cBCJUl2vcay64
# BC8asOqNYGLA1sKlurxyD8B4AJ3qZSu2ZRoaqeJ9OG27ecdRxn9K5Kg82AcbYwyo
# dbrFqGcd756VB/ecQ9khLlJYsMtAHsRTQUzrF0P9OujokgddsEfBNWchj2H/Mx8a
# t26OEN/GoYICCzCCAgcGCSqGSIb3DQEJBjGCAfgwggH0AgEBMHIwXjELMAkGA1UE
# BhMCVVMxHTAbBgNVBAoTFFN5bWFudGVjIENvcnBvcmF0aW9uMTAwLgYDVQQDEydT
# eW1hbnRlYyBUaW1lIFN0YW1waW5nIFNlcnZpY2VzIENBIC0gRzICEA7P9DjI/r81
# bgTYapgbGlAwCQYFKw4DAhoFAKBdMBgGCSqGSIb3DQEJAzELBgkqhkiG9w0BBwEw
# HAYJKoZIhvcNAQkFMQ8XDTE0MDYyNzAzMDk1MlowIwYJKoZIhvcNAQkEMRYEFOL2
# Czn9nFhCyrBZtqon9M2MQOafMA0GCSqGSIb3DQEBAQUABIIBACmXCa/q+Yu9TQd0
# 6Zxx4kWKTTJ4crpQqPTT452vMnLJrJfwRmnKMk62jjFPF6IUVjZ5E4WccCEw0Mks
# JIhr5Ai5AddExS+MqhD7HKpyTKcJWTtnJ8D6a9yCt0UMLbivQsqf9j/ixw5Hly6B
# r0WyT7A8eu+k2I8kGxJiB1Z+lRLUHl6eqEoa3fAT67TsuMWGCr6byvWzksDHimPm
# cyNv3hIc4eJBApoota4xMMpggroTRRuGKjZ3RR/KmsOCRWNRjkV3BHCn0K5ytk8t
# INyBpjBsI5ATGB7GPCUO5CkbWLCGIL7WTNRGdBg6g+UfIEnhMhQXhU5/BZ5KYy0M
# GX0Ji18=
# SIG # End signature block
