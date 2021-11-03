#############################################################################################

#Patch Script Folder - Only change this line for different bases/new folder locations	
$PatchScriptFolder = "\\XLWU-FS-004\325 MSG\325 CS\SCO\SCOO\Super Patch"
#Patch Script Logs Folder - Where the patch logs, setup files, and script running logs are located.
$PatchScriptLogs = "$PatchScriptFolder\Patch Logs\"
#Fully Qualified Domain - Your base's Fully Qualified Domain (Ex. Area52.afnoapps.usaf.mil)
$Domain = ".area52.afnoapps.usaf.mil"

#############################################################################################

#No Error Displays
$ErrorActionPreference = "SilentlyContinue"

#Functions
function Use-RunAs 
{    
    # Check if script is running as Adminstrator and if not use RunAs 
    # Use Check Switch to check if admin 
     
    param([Switch]$Check) 
     
    $IsAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()` 
        ).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator") 
         
    if ($Check) { return $IsAdmin }     
 
    if ($MyInvocation.ScriptName -ne "") 
    {  
        if (-not $IsAdmin)  
        {  
            try 
            {  
                $arg = "-file `"$($MyInvocation.ScriptName)`"" 
                Start-Process "$psHome\powershell.exe" -Verb Runas -ArgumentList $arg -ErrorAction 'stop'  
            } 
            catch 
            { 
                Write-Warning "Error - Failed to restart script with runas"  
                break               
            } 
            exit # Quit this session of powershell 
        }  
    }  
    else  
    {  
        Write-Warning "Error - Script must be saved as a .ps1 file first"  
        break  
    }  
} 

Function Duplicate([ref]$CurrentDup)
    {
		#Check for duplicates and if already patched
		$DC = ($i -eq $Computername[$DupInd])
		If ($DC -eq $true)           
			{$CurrentDup.Value = "1"}         
		Else            
			{$CurrentDup.Value = "0"}
    }

Function Patch([ref]$CurrentPatch)
    {
		#Check if already patched		
        $TP = test-path "$FolderPath\Good\$PatchLogFile"
        If ($TP -eq $true) 
		    {$CurrentPatch.value = "1"}    
        Else               
            {$CurrentPatch.value = "0"}

    }

Function ArrayInit([ref]$Array)
    {
        for ($n = 0; $n -lt $LastIP; $n++)
        {$Array.value += ,""}
    }

#Check if ran as admin
    Use-RunAs

#Popup Variable
	$a = new-object -comobject wscript.shell

#Date and Time
	$dt = get-date -format MMM-dd-yyyy_HH_mm

#Copy PSexec files to local computer
Write-Output "Checking if local computer has Psexec."
$PSexecCheck = test-path "c:\Windows\System32\psexec.exe"

If ($PSexecCheck -eq $false) 
    {
    Write-Output "Psexec file is copying. This may take up to 15 seconds."
    Copy-Item "$PatchScriptLogs\Setup_Files\psexec.exe" "C:\Windows\System32\" 
    Write-Output "PSexec file copied.`n"
    }
Else
    {Write-Output "Psexec file already exists."`n}

#Check for Script Running Log File/If Log exists
    $LogCheck = Test-Path "$PatchScriptLogs\Script_Running\*.csv"
        
    If ($LogCheck -eq $true)
        {
        $RunLogCheck = Get-ChildItem -Path ("$PatchScriptLogs\Script_Running") -filter *.csv -name
        $ScriptRunner = ($RunLogCheck.ToString().Replace(".csv","")) 
         $intAnswer = $a.popup("The MST is being ran currently on $ScriptRunner. If you know it is not being run by that system, delete the run log. Do you want to delete the Script Running Log?", 
         0,"Delete Files",4)
        If ($intAnswer -eq 6)
			{
			Remove-Item "$PatchScriptLogs\Script_Running\*" -recurse
			$a.popup("The Script Running Log has been deleted, the script will continue.", 
			0,"Delete Files",0) | Out-Null
			}          
        Else
			{
			$a.popup("Script is currently running, please try again later.", 
			0,"Delete Files",0) | Out-Null
			exit
			}
        }

#User Popups Yes/No
$intAnswer = $a.popup("Do you want to display user popups on remote workstations?", `
0,"Popups",260)
If ($intAnswer -eq 6)
    {
    $UserInput = "Yes"
    } 
else
    {
    $UserInput = "No"
    }

If ($userinput -eq "Yes")
    {
    #Kicker Popup
        $object = New-Object -comObject Shell.Application  
        $kicker = $object.BrowseForFolder(0, 'Select a Kicker Folder.', 0, $PatchScriptFolder) 
        If ($kicker -ne $null) 
            {
		    $KickerPath = ($kicker.self.Path)
            }
        Else
            {
		    $a.popup("You did not select a kicker folder.  The script will now end.", 0, "Error", 0) | Out-Null
            Remove-Item "$PatchScriptLogs\Script_Running\$env:COMPUTERNAME by $env:username.csv" | Out-Null
            exit    
            }
     }

#Delayed Patching Yes/No
$intAnswer = $a.popup("Do you want to set a time delay for the patch?", `
0,"Time Delay",260)
If ($intAnswer -eq 6)
    {
    $DelayedPatch = "Yes"
    } 
else 
    {
    $DelayedPatch = "No"
    }       

#Input Popup
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 

$objForm = New-Object System.Windows.Forms.Form 
$objForm.Text = "Enter Log Name"
$objForm.Size = New-Object System.Drawing.Size(300,150) 
$objForm.StartPosition = "CenterScreen"

$objForm.KeyPreview = $True
$objForm.Add_KeyDown({if ($_.KeyCode -eq "Enter") 
    {$objForm.Close()}})
$objForm.Add_KeyDown({if ($_.KeyCode -eq "Escape") 
    {$objForm.Close()}})

$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Size(50,75)
$OKButton.Size = New-Object System.Drawing.Size(75,23)
$OKButton.Text = "OK"
$OKButton.Add_Click({$objForm.Close()})
$objForm.Controls.Add($OKButton)

$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Size(150,75)
$CancelButton.Size = New-Object System.Drawing.Size(75,23)
$CancelButton.Text = "Cancel"
$CancelButton.Add_Click({$objForm.Close()})
$objForm.Controls.Add($CancelButton)

$objLabel = New-Object System.Windows.Forms.Label
$objLabel.Location = New-Object System.Drawing.Size(10,20) 
$objLabel.Size = New-Object System.Drawing.Size(280,20) 
$objLabel.Text = "Please enter the information in the space below:"
$objForm.Controls.Add($objLabel) 

$objTextBox = New-Object System.Windows.Forms.TextBox 
$objTextBox.Location = New-Object System.Drawing.Size(10,40) 
$objTextBox.Size = New-Object System.Drawing.Size(260,20) 
$objForm.Controls.Add($objTextBox) 

$objForm.Topmost = $True

$objForm.Add_Shown({$objForm.Activate();$objTextBox.focus()})
[void] $objForm.ShowDialog()

#Grab InputBox String
$Input=$objTextBox.Text

#Patch Folder Popup
    $object = New-Object -comObject Shell.Application  
    $folder = $object.BrowseForFolder(0, 'Select a folder', 0, $PatchScriptFolder) 
    If ($folder -ne $null) 
        {
		$FolderPath = ($folder.self.Path)
        $FolderName = ($folder.self.Name)
        }
    Else
        {
		$a.popup("You did not select a folder.  The script will now end.", 0, "Error", 0) | Out-Null
        Remove-Item "$PatchScriptLogs\Script_Running\$env:COMPUTERNAME by $env:username.csv" | Out-Null
        exit    
        }

#Computer List File Popup
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $PatchScriptFolder
    $OpenFileDialog.filter = "Text Files (*.txt)| *.txt"
    $FileCheck = $OpenFileDialog.ShowDialog()
    $ComputerList = $OpenFileDialog.filename
    If ($FileCheck -eq "Cancel")
		{
        $a.popup("You did not select a file.  The script will now end.", 0, "Error", 0) | Out-Null
        Remove-Item "$PatchScriptLogs\Script_Running\$env:COMPUTERNAME by $env:username.csv" | Out-Null
        exit
        }

#Time Delay
If ($DelayedPatch -eq "Yes")
    {
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 

    $objForm = New-Object System.Windows.Forms.Form 
    $objForm.Text = "Select a Time"
    $objForm.Size = New-Object System.Drawing.Size(300,300) 
    $objForm.StartPosition = "CenterScreen"

    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = New-Object System.Drawing.Size(100,220)
    $OKButton.Size = New-Object System.Drawing.Size(75,23)
    $OKButton.Text = "Select"
    $OKButton.Add_Click({$objForm.Close()})
    $objForm.Controls.Add($OKButton)

    $objLabel = New-Object System.Windows.Forms.Label
    $objLabel.Location = New-Object System.Drawing.Size(10,20) 
    $objLabel.Size = New-Object System.Drawing.Size(280,20) 
    $objLabel.Text = "Please select a military time:"
    $objForm.Controls.Add($objLabel) 

    $objListBoxHour = New-Object System.Windows.Forms.ListBox 
    $objListBoxHour.Location = New-Object System.Drawing.Size(10,40) 
    $objListBoxHour.Size = New-Object System.Drawing.Size(130,20) 
    $objListBoxHour.Height = 160

    $objListBoxMin = New-Object System.Windows.Forms.ListBox 
    $objListBoxMin.Location = New-Object System.Drawing.Size(141,40) 
    $objListBoxMin.Size = New-Object System.Drawing.Size(130,20) 
    $objListBoxMin.Height = 160

    [void] $objListBoxHour.Items.Add("0")
    [void] $objListBoxHour.Items.Add("1")
    [void] $objListBoxHour.Items.Add("2")
    [void] $objListBoxHour.Items.Add("3")
    [void] $objListBoxHour.Items.Add("4")
    [void] $objListBoxHour.Items.Add("5")
    [void] $objListBoxHour.Items.Add("6")
    [void] $objListBoxHour.Items.Add("7")
    [void] $objListBoxHour.Items.Add("8")
    [void] $objListBoxHour.Items.Add("9")
    [void] $objListBoxHour.Items.Add("10")
    [void] $objListBoxHour.Items.Add("11")
    [void] $objListBoxHour.Items.Add("12")
    [void] $objListBoxHour.Items.Add("13")
    [void] $objListBoxHour.Items.Add("14")
    [void] $objListBoxHour.Items.Add("15")
    [void] $objListBoxHour.Items.Add("16")
    [void] $objListBoxHour.Items.Add("17")
    [void] $objListBoxHour.Items.Add("18")
    [void] $objListBoxHour.Items.Add("19")
    [void] $objListBoxHour.Items.Add("20")
    [void] $objListBoxHour.Items.Add("21")
    [void] $objListBoxHour.Items.Add("22")
    [void] $objListBoxHour.Items.Add("23")

    [void] $objListBoxMin.Items.Add("00")
    [void] $objListBoxMin.Items.Add("05")
    [void] $objListBoxMin.Items.Add("10")
    [void] $objListBoxMin.Items.Add("15")
    [void] $objListBoxMin.Items.Add("20")
    [void] $objListBoxMin.Items.Add("25")
    [void] $objListBoxMin.Items.Add("30")
    [void] $objListBoxMin.Items.Add("35")
    [void] $objListBoxMin.Items.Add("40")
    [void] $objListBoxMin.Items.Add("45")
    [void] $objListBoxMin.Items.Add("50")
    [void] $objListBoxMin.Items.Add("55")


    $objForm.Controls.Add($objListBoxHour) 
    $objForm.Controls.Add($objListBoxMin)

    $objForm.Topmost = $True

    $objForm.Add_Shown({$objForm.Activate()})
    [void] $objForm.ShowDialog()

    # Date and Time for Delay
    $DelayedTime = $null
    $dtdelay = get-date -format HH:mm
    $dtdelaysplit = $dtdelay.Split(":")
    [int]$dtHour = [int]$dtdelaysplit[0]
    [int]$dtMin = [int]$dtdelaysplit[1]

    [int]$H = $objListBoxHour.SelectedItem
    [int]$M = $objListBoxMin.SelectedItem
    
    $H = ($H * 60)
    $MinTotalSel = ($H + $M)
    $dtHour = ($dtHour * 60)
    $dtTotal = ($dtHour + $dtMin)
           
    $DelayedTime = (($MinTotalSel - $dtTotal) * 60)
           
    }

Start-Sleep -s $DelayedTime

#Create Script Running Log
    New-Item -ItemType file -Path "$PatchScriptLogs\Script_Running\$env:COMPUTERNAME by $env:username.csv" | Out-Null

#timer start
$starttimer = Get-Date

#Patch .cmd variables
    If ($userinput -eq "yes")
    {$PatchCmd = Get-ChildItem -Path $KickerPath -filter *.cmd -name}
    Else
    {$PatchCmd = Get-ChildItem -Path $FolderPath -filter *.cmd -name}  

    $BatchCmdPath = "C$\installfiles\"
	$CMDPath = "$BatchCmdPath\$PatchCmd"
    $PSExecCmdPath = "C:\installfiles\"
	
#Creates MST Log
    $MSTLOG = "$PatchScriptLogs\$Input-$FolderName-MST-$dt.csv"
    New-Item -ItemType file -Path $MSTLOG | Out-Null
    Add-Content -Value "Computer Name,DNS IP Resolution,Online,DNS Reverse Lookup,DNS Compare,MAC Address,Date of Last Login,Last User to Login,Patch Status" -Path $MSTLOG

#Put Computer Names from .txt into Array and Get total count of systems
    $ComputerArray = get-content $OpenFileDialog.filename
    $LastIP = ($ComputerArray.getupperbound(0) + 1)
    if ($LastIP -eq $null)
        {
        if ($ComputerArray -ne $null)
            {
            $LastIP = 1
            }
        else
            {
            $LastIP = 0
            }
        }

#Reset Arrays
$DupArray = @()
$PatchArray = @()
$Mac = @()
$DisplayName = @()
$Date = @()
$IP = @()
$IPLookup = @()
$ComputerName = @()
ArrayInit([ref]$IPLookup)
ArrayInit([ref]$DupArray)
ArrayInit([ref]$PatchArray)
ArrayInit([ref]$Mac)
ArrayInit([ref]$IP)
ArrayInit([ref]$ComputerName)
ArrayInit([ref]$DisplayName)
ArrayInit([ref]$Date)
#Reset variables
    $DupInd=$WOLTotal=$AlreadyPatched=$Counter=$Dup=$Patch=$MacDisplay = 0

#Start WOL
  
    #Run each item in array
    Write-Output "Checking List for current IP's"

        Foreach ($i in $ComputerArray)
        {
        #Check if computer list contains fully qualified domain name
        If ($i.contains("$domain"))
        {
        $isplit = $i.split(".")
        $i = $isplit[0]
        }
        #Check if computer list contains IPs already 
        If ($i.contains("132.10."))
            {
			#Resolve DNSName if IP
            $ns = $null
            $ns = [system.net.dns]::Gethostbyaddress("$i").hostname.Split(".")
            $i = $ns[0]
            }
            $Computername[$counter] = $i
            $IPLookup[$counter] = [System.Net.Dns]::GetHostentryAsync("$i")

          $counter += 1
          If ($Counter -eq $LastIP)
            {
                Start-Sleep -s 3
                Write-Output "Saving IPs!"
                For($Count = 0;$Count -lt $LastIP; $Count++)
                {
                $IP[$count] = $IPLookup[$count].Result.AddressList.IPAddressToString
                $countdisplay += 1
                }
            }
          }
          $counter = 0
          Write-Output "Sending WOL Packets!"
          Foreach ($i in $ComputerName)
            {
        $PatchLogFile = "$i.csv"
        $DupInd += 1
        Duplicate([ref]$Dup)
        Patch([ref]$Patch)
        $DupArray[$Counter] = $Dup
        $PatchArray[$Counter] = $Patch                            

                #Read Audit logs
			    $AuditData = "\\XLWU-FS-004\325 MSG\325 CS\SCO\SCOO\Super Patch\Patch_Logs\Audit_Logs\$i.csv"
			    $AD = test-path $AuditData
				If ($AD -eq $true)
					{
                    $Audit = Import-Csv $AuditData
					$Mac[$Counter] = ($Audit[-1]."MAC Address")

					$DisplayName[$Counter] = ($Audit[-1]."Display Name")
                    $Date[$Counter] = $Audit[-1]."Date"
					$TrueIP = ($IP[$Counter] -Match "132.10.")
					$AM = ($Mac[$Counter] -ne $null)
                    $MacDisplay = $Mac[$Counter]
                    $IPDisplay = $IP[$Counter]  
					If ($TrueIP -eq $true) 
						{
						If ($AM -eq $true)
							{
                            If ($PatchArray[$Counter] -eq "0")
                                {
                                If ($DupArray[$Counter] -eq "0")                    
                                    {				   
				                    #Run mc-wol.exe
				                    $WolCmd = "\\XLWU-FS-004\325 MSG\325 CS\SCO\SCOO\Super Patch\Software\mc-wol.exe"
				                    Invoke-Expression "$WolCmd $MACDisplay $IPDisplay" | Out-Null
				                    Write-Output "WOL sent to $i MAC - $MacDisplay IP - $IPDisplay"
				                    $WOLTotal = ($WOLTotal + 1)
                                    }
                                Else {Write-Output "$i is a duplicate, skipping system"}  
			                    }				
		                        Else			
		                            {
                                    Write-output "$i was Already Patched, no WOL sent"
                                    $AlreadyPatched += 1
                                    }
							}                         
						}
                    Else {Write-Output "MAC or IP was unavailable for $i"}
                    }
                    $Counter += 1                              
        }	    

    #Display total WOL sent
        Write-Output "Total WOLs Sent = $WOLTotal/$LastIP Already Patched = $AlreadyPatched/$LastIP`n" #unfinished variables
       
		If ($Input.contains("test"))
			{
			Write-Output "Waiting 5 seconds for testing`n"
			start-sleep -s 5
			}
		Else
			{
			Write-Output "Waiting 2 minutes for systems to start`n"
			start-sleep -s 120
			}

	#Start Ping
		$DupInd=$SysDone=$SysPatched=$Counter = 0

    #Run each item in array
        Foreach ($i in $ComputerName)
			{
            #Check if computer list contains IPs already 
            If ($i.contains("131.55."))
                {
			    #Resolve DNSName if IP
                $ns = $null
                $ns = [system.net.dns]::Gethostbyaddress("$i").hostname.Split(".")
                $i = $ns[0]
                }
            $PatchLogFile = "$i.csv"
			$DupInd += 1
            $SysDone += 1 
            $MacDisplay = $Mac[$Counter] 
            $DateDisplay = $Date[$Counter]
            $DisplayNameDisplay = $DisplayName[$Counter]
            
            $TestConnection = test-connection -ComputerName $i -count 1 -Quiet -timetolive 3
            If ($TestConnection -eq $true)
            {$Online = "Online"}
            Else
                {$Online = "Offline"}

            If ($PatchArray[$Counter] -eq "0")
                {If ($Online -eq "Online")
                    {$APDisplay = "Patched"}
                Else {$APDisplay = "Offline - Not Patched"}
                }
            Else
                {$APDisplay = "Already Patched"}

            If ($DupArray[$Counter] -eq "1")
                {Write-Output "$i was a duplicate computer, skipping system"}
            ElseIf ($PatchArray[$Counter] -eq "1")
                {Write-Output "$i was already patched, skipping system"}
            ElseIf ($Online -eq "Online") 
                {Write-Output "$i is currently patching"}
            Else 
                {Write-Output "$i was offline, could not patch"}
           
            # DNS Reverse Lookup/Compare
            $DNSRL = $IP[$Counter]
            $ns = ""
            $ns = [system.net.dns]::Gethostbyaddress("$DNSRL").hostname.Split(".")
            $ReverseLookup = $ns[0]

            If ($ReverseLookup -eq $i)
                {$DnsCompare = "Correct"}
            Else {$DnsCompare = "Incorrect"}
            
        If ($PatchArray[$Counter] -eq "0")
            {
            If ($DupArray[$Counter] -eq "0")
                {
                    If ($Online -eq "Online")
                        {                           
                        #Copy and run patch .cmd
                        If ($userinput -eq "yes")
                        {$PatcherPatchDir = "$KickerPath\$PatchCmd"}
                        Else
                        {$PatcherPatchDir = "$FolderPath\$PatchCmd"}

                        $PatcherFolderDir = "\\$i\$BatchCmdPath$FolderName\"
                        $PatcherPsExec = "\\$i $PSExecCmdPath$FolderName\$PatchCmd"
	                    Write-Output "Copying Files to $i\$BatchCmdPath$FolderName\"
                        
                        #Runspace
                        If ($userinput -eq "yes")
                        {$code = "New-Item -ItemType directory -Path " + (Get-Variable -ValueOnly PatcherFolderDir) + "`n" + "Copy-Item " + (Get-Variable -ValueOnly PatcherPatchDir) + " " + (Get-Variable -ValueOnly PatcherFolderDir) + "`n" + "invoke-expression " + '"' +  "psexec.exe -s -d -h " + (Get-Variable -ValueOnly PatcherPsExec) + " " + $i + " " + $FolderName + '"' + " 2>&1"}
                        Else
                        {$code = "New-Item -ItemType directory -Path " + (Get-Variable -ValueOnly PatcherFolderDir) + "`n" + "Copy-Item " + (Get-Variable -ValueOnly PatcherPatchDir) + " " + (Get-Variable -ValueOnly PatcherFolderDir) + "`n" + "invoke-expression " + '"' +  "psexec.exe -s -d -h " + (Get-Variable -ValueOnly PatcherPsExec) + '"' + " 2>&1"}
                        
                        $timedout = $null
                        $timer = $null
                        $newPowerShell = [PowerShell]::Create().AddScript($code)
                        $runspace = $newPowerShell.BeginInvoke() 
                        While (-Not $runspace.IsCompleted) {Start-Sleep -s 1; $timer += 1; $runspacecomplete = ($timer -eq 4); If ($runspacecomplete -eq $true) {Write-output "Access denied to C$, Did not patch.";$APDisplay = "Access denied to C$";$timedout = $true; break}}
                        $newPowershell.BeginStop | Out-Null
                        $newPowerShell.Dispose($runspace)
                        If ($timedout -eq $null)
                        {$SysPatched += 1}
                        }
                }
            }
            
            If ($i.Contains("132.10."))
                        {#Check if $i is IP already
	                    Add-Content -Value "$i,$i,$Online,$ReverseLookup,$DnsCompare,$MacDisplay,$DateDisplay,$DisplayNameDisplay,$APDisplay" -Path $MSTLOG
	                    }                                   
                    Else                
	                    {#Resolve DNS to IP
                        $DNStoIP = ""
	                    $DNStoIP = [System.Net.Dns]::GetHostAddresses("$i").IPAddressToString
	                    Add-Content -Value "$i,$DnstoIP,$Online,$ReverseLookup,$DnsCompare,$MacDisplay,$DateDisplay,$DisplayNameDisplay,$APDisplay" -Path $MSTLOG
                        }

            Write-Output "Sys Done = $SysDone/$LastIP, Sys Patched = $SysPatched/$LastIP, Sys Already Patched = $AlreadyPatched/$LastIP`n"         
            $Counter += 1
            
            }
#timer end
$stoptimer = Get-Date
"Total time: {0} Minutes" -f [math]::round(($stoptimer - $starttimer).TotalMinutes , 2)
 #Delete Script Running Log       
    Remove-Item "$PatchScriptLogs\Script_Running\$env:COMPUTERNAME by $env:username.csv" | Out-Null
    $a.popup("Script has been completed.", 
	0,"Complete!",0) | Out-Null 
pause 