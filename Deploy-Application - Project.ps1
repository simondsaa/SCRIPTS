<#
.SYNOPSIS
	This script performs the installation or uninstallation of an application(s).
.DESCRIPTION
	The script is provided as a template to perform an install or uninstall of an application(s).
	The script either performs an "Install" deployment type or an "Uninstall" deployment type.
	The install deployment type is broken down into 3 main sections/phases: Pre-Install, Install, and Post-Install.
	The script dot-sources the AppDeployToolkitMain.ps1 script which contains the logic and functions required to install or uninstall an application.
.PARAMETER DeploymentType
	The type of deployment to perform. Default is: Install.
.PARAMETER DeployMode
	Specifies whether the installation should be run in Interactive, Silent, or NonInteractive mode. Default is: Interactive. Options: Interactive = Shows dialogs, Silent = No dialogs, NonInteractive = Very silent, i.e. no blocking apps. NonInteractive mode is automatically set if it is detected that the process is not user interactive.
.PARAMETER AllowRebootPassThru
	Allows the 3010 return code (requires restart) to be passed back to the parent process (e.g. SCCM) if detected from an installation. If 3010 is passed back to SCCM, a reboot prompt will be triggered.
.PARAMETER TerminalServerMode
	Changes to "user install mode" and back to "user execute mode" for installing/uninstalling applications for Remote Destkop Session Hosts/Citrix servers.
.PARAMETER DisableLogging
	Disables logging to file for the script. Default is: $false.
.PARAMETER Repair
	Repairs an existing installation. Default is: $false.
.EXAMPLE
	Deploy-Application.ps1
.EXAMPLE
	Deploy-Application.ps1 -DeployMode 'Silent'
.EXAMPLE
	Deploy-Application.ps1 -AllowRebootPassThru -AllowDefer
.EXAMPLE
	Deploy-Application.ps1 -DeploymentType Uninstall
.NOTES
	Toolkit Exit Code Ranges:
	60000 - 68999: Reserved for built-in exit codes in Deploy-Application.ps1, Deploy-Application.exe, and AppDeployToolkitMain.ps1
	69000 - 69999: Recommended for user customized exit codes in Deploy-Application.ps1
	70000 - 79999: Recommended for user customized exit codes in AppDeployToolkitExtensions.ps1
.LINK
	http://psappdeploytoolkit.codeplex.com
#>
[CmdletBinding()]
Param (
	[Parameter(Mandatory=$false)]
	[ValidateSet('Install','Uninstall')]
	[string]$DeploymentType = 'Install',
	[Parameter(Mandatory=$false)]
	[ValidateSet('Interactive','Silent','NonInteractive')]
	[string]$DeployMode = 'Interactive',
	[Parameter(Mandatory=$false)]
	[switch]$AllowRebootPassThru = $false,
	[Parameter(Mandatory=$false)]
	[switch]$TerminalServerMode = $false,
	[Parameter(Mandatory=$false)]
	[switch]$Repair = $false,
	[Parameter(Mandatory=$false)]
	[switch]$DisableLogging = $false,
	[switch]$addComponentsOnly = $false, # Specify whether running in Component Only Mode
	[switch]$addInfoPath = $false, # Add InfoPath to the install
	[switch]$addOneNote = $false, # Add OneNote to the install
	[switch]$addOutlook = $false, # Add Outlook to the install
	[switch]$addPublisher = $false, # Add Publisher to the install
	[switch]$addSharepointWorkspace = $false, # Add Sharepoint Workspace to the install
	[switch]$applyPolicy = $false #If specified, apply STIG policy and exit the script
)

Try {
	## Set the script execution policy for this process
	Try { Set-ExecutionPolicy -ExecutionPolicy 'ByPass' -Scope 'Process' -Force -ErrorAction 'Stop' } Catch {}
	
	##*===============================================
	##* VARIABLE DECLARATION
	##*===============================================
	## Variables: Application
	[string]$appVendor = 'Nil - '
	[string]$appName = 'MS Project 16'
	[string]$appVersion = ''
	[string]$appArch = 'x86'
	[string]$appLang = 'EN'
	[string]$appRevision = '01'
	[string]$appScriptVersion = ''
	[string]$appScriptDate = '11/24/2015'
	[string]$appScriptAuthor = 'AFECMO Enterprise Products Team'
	##*===============================================
	
				##* AFECMO Enterprise Products Team Region Additions
	
#region Script information variables
# Grab the script path and name for logging purposes>
$ScriptDir = Split-Path $MyInvocation.MyCommand.Path
$ScriptName = $MyInvocation.MyCommand.Name
# Grab locations of Powershell
$ps64 = "$env:windir\sysnative\WindowsPowerShell\v1.0\powershell.exe"
$ps32 = "$env:windir\syswow64\WindowsPowerShell\v1.0\powershell.exe"
#endregion

#region Get-ScriptDirectory
function Get-ScriptDirectory
{
$Invocation = (Get-Variable MyInvocation -Scope 1).Value
Split-Path $Invocation.MyCommand.Path
}

#endregion

#region Environment 

##Grab the environment and shell variables
$OSArch = (gwmi Win32_OperatingSystem).OSArchitecture
$ShellArch = [intptr]::Size                 # returns 4 for 32-bit and 8 for 64-bit
$Windows = [environment]::ExpandEnvironmentVariables("%windir%")
$WindowsTemp = (Join-Path $Windows -ChildPath \temp)
$CTemp = [environment]::ExpandEnvironmentVariables("%SYSTEMDRIVE%\temp\")
$Sys32 = [environment]::ExpandEnvironmentVariables("%WINDIR%\System32\")

#endregion

#region Logging
<#
    Generate Logname based on application name
#>
$logName = $appVendor + $appName + $appVersion + ".log"
#endregion
	
				##* END AFECMO Enterprise Products Team Region Additions
	
	##* Do not modify section below
	#region DoNotModify
	
	## Variables: Exit Code
	[int32]$mainExitCode = 0
	
	## Variables: Script
	[string]$deployAppScriptFriendlyName = 'Deploy Application'
	[version]$deployAppScriptVersion = [version]'3.6.1'
	[string]$deployAppScriptDate = '03/20/2015'
	[hashtable]$deployAppScriptParameters = $psBoundParameters
	
	## Variables: Environment
	[string]$scriptDirectory = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
	
	## Dot source the required App Deploy Toolkit Functions
	Try {
		[string]$moduleAppDeployToolkitMain = "$scriptDirectory\AppDeployToolkit\AppDeployToolkitMain.ps1"
		If (-not (Test-Path -Path $moduleAppDeployToolkitMain -PathType Leaf)) { Throw "Module does not exist at the specified location [$moduleAppDeployToolkitMain]." }
		If ($DisableLogging) { . $moduleAppDeployToolkitMain -DisableLogging } Else { . $moduleAppDeployToolkitMain }
	}
	Catch {
		[int32]$mainExitCode = 69001
		Write-Error -Message "Module [$moduleAppDeployToolkitMain] failed to load: `n$($_.Exception.Message)`n `n$($_.InvocationInfo.PositionMessage)" -ErrorAction 'Continue'
		Exit $mainExitCode
	}
	
	#endregion
	##* Do not modify section above
	##*===============================================
	##* END VARIABLE DECLARATION
	##*===============================================
	
	#  Set the initial Office folder
	[string] $dirOffice = Join-Path -Path "$envProgramFilesX86" -ChildPath "Microsoft Office"
    [string] $dirLync = Join-Path -Path "$envProgramFilesX86" -ChildPath "Microsoft Lync"
	[string] $dirCommunicator = Join-Path -Path "$envProgramFilesX86" -ChildPath "Microsoft Office Communicator"
	
	If ($deploymentType -ine 'Uninstall') {
		##*===============================================
		##* PRE-INSTALLATION
		##*===============================================
		[string]$installPhase = 'Pre-Installation'
		
		#Versioning information
		$PkgVersion = "180315"
		Write-Log -Message "Beginning installation of SDC NIPR and SIPR - Microsoft Office 2016 package version $PkgVersion"
		
		#If the applyPolicy switch is specified, apply the STIG pol files and exit the script
		if ($applyPolicy) {
			Write-Log -Message "Apply Policy switch specified. Applying STIG policy files and exiting."
			#Remove deprecated policy settings if present
			Write-Log -Message "Removing old policy settings if present."
			$polCleanScript = Get-ChildItem -Path "$dirSupportFiles\Cleanup-GPOConfig" | Where-Object {$_.Name -like "*.ps1"} | Select-Object -ExpandProperty FullName
			$polCleanFile = Get-ChildItem -Path "$dirSupportFiles\Cleanup-GPOConfig" | Where-Object {$_.Name -like "*.txt"} | Select-Object -ExpandProperty FullName
			Execute-Process -Path "$envSystem32Directory\WindowsPowerShell\v1.0\powershell.exe" -Parameters "-File `"$polCleanScript`" -SourceFile `"$polCleanFile`""
			#Apply GPOs
			$PolDef = [environment]::ExpandEnvironmentVariables('%systemroot%\PolicyDefinitions')
			$PolDefLang = [environment]::ExpandEnvironmentVariables('%systemroot%\PolicyDefinitions\en-us')
			Write-Log -Message "Copying ADMX and ADML files to Group Policy definitions."
			Write-Log -Message "Policy definitions folder is $PolDef."
			Write-Log -Message "Language folder is $PolDefLang."
			xcopy $dirSupportFiles\ADMX $PolDef /i /y /c
			xcopy $dirSupportFiles\ADML $PolDefLang /i /y /c
			$machpols = Get-ChildItem $dirSupportfiles | Where {$_.name -like "*mach.pol"} | Select -Expand Name
				If ($machpols){
				Write-Log -Message "Creating Machine Group Policy settings."
				foreach ($machpol in $machpols){
				Execute-Process -Path "$dirSupportFiles\ImportRegPol.exe" -Parameters "-m $machpol"
					}
				}
			$userpols = Get-ChildItem $dirSupportfiles | Where {$_.name -like "*user.pol"} | Select -Expand Name
				If ($userpols){
				Write-Log -Message "Creating User Group Policy settings."
				foreach ($userpol in $userpols){
				Execute-Process -Path "$dirSupportFiles\ImportRegPol.exe" -Parameters "-u $userpol"
					}
				}
			#Versioning information
			Write-Log -Message "Writing versioning information to the registry"
			Set-RegistryKey -Key "HKEY_LOCAL_MACHINE\SOFTWARE\USAF\SDC\Applications" -Name "SdcNiprAndSiprMicrosoftOffice2016Version" -Value $PkgVersion -Type DWORD
			Write-Log -Message "Writing versioning information to WMI"
			$mofFile = Get-ChildItem -Path "$dirSupportFiles" | Where-Object { $_.Name -like "*.mof" } | Select-Object -ExpandProperty FullName
			if ($envOSArchitecture -like "*64*") {
				Execute-Process -Path "$envWinDir\SysWOW64\wbem\mofcomp.exe" -Parameters "`"$mofFile`""
			} else {
				Execute-Process -Path "$envSystem32Directory\wbem\mofcomp.exe" -Parameters "`"$mofFile`""
			}
			Exit-Script -exitcode 69013
		}
		
		#Perform Repair if selected
		If ($Repair){
			Show-InstallationProgress -StatusMessage 'Repairing Office Professional. This may take some time. Please wait...'
			Write-Log -Message "Repairing Office 2016."
			Execute-Process -path "$dirfiles\Setup.exe" -Parameters "/repair Proplus /config `"$dirFiles\ProPlus.WW\SilentRepairConfig.xml`"" -WindowStyle Hidden
			Exit-Script -ExitCode $mainExitCode
		}
		
		
		## Check whether running in Add Components Only mode
		If ($addComponentsOnly) {
			#  Verify that components were specified on the command-line
			If ((-not $addInfoPath) -and (-not $addSharepointWorkspace) -and (-not $addOneNote) -and (-not $addOutlook) -and (-not $addPublisher)) {
				Show-InstallationPrompt -Message 'No addon components were specified' -ButtonRightText 'OK' -Icon 'Error'
				Write-Log -Message "No addon components were specified. Cancelling installation."
				Exit-Script -ExitCode 69007
			}
			
			#  Verify that Office 2016 is already installed
			$officeVersion = Get-ItemProperty -Path 'HKLM:SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{90160000-0011-0000-0000-0000000FF1CE}' -ErrorAction 'SilentlyContinue' | Select-Object -ExpandProperty DisplayName
			
			#  If not found, display an error and exit
			If (-not $officeVersion) {
				Show-InstallationPrompt -Message 'Unable to add the requested components as Office 2016 is not currently installed.' -ButtonRightText 'OK' -Icon 'Error'
				Write-Log -Message "Unable to add the requested components as Office 2016 is not installed."
			}
		}
		
		## Show Welcome Message, close Microsoft Office applications if required, allow up to 3 deferrals, and verify there is enough disk space to complete the install
		Show-InstallationWelcome -CloseApps 'DATABASECOMPARE,MSACCESS,EXCEL,INFOPATH,SETLANG,MSOUC,ONENOTE,OUTLOOK,POWERPNT,MSPUB,WINWORD,ONENOTEM,SPREADSHEETCOMPARE,COMMUNICATOR' -AllowDefer -DeferTimes 3 -CheckDiskSpace
		
		#Region Prerequisite check for Windows 7 and 8.x systems
		$W7x86Patch = Get-ChildItem -Path "$dirSupportFiles\W7x86Patch" | Where-Object { $_.Name -like "*.msu*" } | Select-Object -ExpandProperty FullName
		$W7x64Patch = Get-ChildItem -Path "$dirSupportFiles\W7x64Patch" | Where-Object { $_.Name -like "*.msu*" } | Select-Object -ExpandProperty FullName
		$W8x64Patch = Get-ChildItem -Path "$dirSupportFiles\W8x64Patch" | Where-Object { $_.Name -like "*.msu*" } | Select-Object -ExpandProperty FullName
		$KBNum = "KB2999226"
		$WinVer = "$envOSVersionMajor" + "." + "$envOSVersionMinor"
		
		#Check if target OS is Windows 8.x. If Win 8.x and KB2999226 is not installed, install it.
		If ($WinVer -eq "6.3") {
			$HotfixInstalled = Get-HotFix -Id $KBNum
			If ($HotfixInstalled -eq $null) {
				Write-Log -Message "KB2999226 is not installed. Installing."
				Execute-Process -Path "$envSystem32Directory\wusa.exe" -Parameters "`"$W8x64Patch`" /quiet /norestart"
			}
		}
		#Check if target OS is Windows 7. If so, check if 64- or 32-bit. For each, check if KB2999226 is installed. If not, install it.
		If ($WinVer -eq 6.1) {
			$HotfixInstalled = Get-HotFix -Id $KBNum
			If ($envOSArchitecture -like "*64*") {
				If ($HotfixInstalled -eq $null) {
					Write-Log -Message "KB2999226 is not installed. Installing."
					Execute-Process -Path "$envSystem32Directory\wusa.exe" -Parameters "`"$W7x64Patch`" /quiet /norestart"
				}
			} else {
				If ($HotfixInstalled -eq $null) {
					Write-Log -Message "KB2999226 is not installed. Installing."
					Execute-Process -Path "$envSystem32Directory\wusa.exe" -Parameters "`"$W7x86Patch`" /quiet /norestart"
				}
			}
		}
		#endregion
		
		## Display Pre-Install cleanup status
			Show-InstallationProgress -StatusMessage 'Performing Pre-Install cleanup. This may take some time. Please wait...'
		
			# Remove any previous version of Office (if required)
			[string[]]$officeExecutables = 'excel.exe', 'groove.exe', 'infopath.exe', 'onenote.exe', 'outlook.exe', 'mspub.exe', 'powerpnt.exe', 'winword.exe', 'winproj.exe', 'visio.exe' , 'setlang.exe' , 'msouc.exe' , 'onenotem.exe'  
			$RegUninstallPaths = @( 
	    	'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall', 
	    	'HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall') 
	 		$UninstallSearchFilter = { ($_.GetValue('DisplayName') -like '*Microsoft Lync 2010*') } 
		
			ForEach ($officeExecutable in $officeExecutables) {
				If (Test-Path -Path (Join-Path -Path $dirOffice -ChildPath "Office12\$officeExecutable") -PathType Leaf) {
					Write-Log -Message 'Microsoft Office 2007 was detected. Will be uninstalled.' -Source $deployAppScriptFriendlyName
					Execute-Process -Path 'cscript.exe' -Parameters "`"$dirSupportFiles\OffScrub07.vbs`" ProPlus /S /Q /NoCancel" -WindowStyle Hidden -IgnoreExitCodes '1,2,3'
					Write-Log -Message "Removing old policy settings if present."
					$polCleanScript = Get-ChildItem -Path "$dirSupportFiles\Cleanup-GPOConfig" | Where-Object {$_.Name -like "*.ps1"} | Select-Object -ExpandProperty FullName
					$polCleanFile = Get-ChildItem -Path "$dirSupportFiles\Cleanup-GPOConfig" | Where-Object {$_.Name -like "*.txt"} | Select-Object -ExpandProperty FullName
					Execute-Process -Path "$envSystem32Directory\WindowsPowerShell\v1.0\powershell.exe" -Parameters "-File `"$polCleanScript`" -SourceFile `"$polCleanFile`""
					Break
				}
			}
			ForEach ($officeExecutable in $officeExecutables) {
				If (Test-Path -Path (Join-Path -Path $dirOffice -ChildPath "Office14\$officeExecutable") -PathType Leaf) {
					Write-Log -Message 'Microsoft Office 2010 was detected. Will be uninstalled.' -Source $deployAppScriptFriendlyName
					Execute-Process -Path "cscript.exe" -Parameters "`"$dirSupportFiles\OffScrub10.vbs`" ProPlus /S /Q /NoCancel" -WindowStyle Hidden -IgnoreExitCodes '1,2,3'
					Write-Log -Message "Removing old policy settings if present."
					$polCleanScript = Get-ChildItem -Path "$dirSupportFiles\Cleanup-GPOConfig" | Where-Object {$_.Name -like "*.ps1"} | Select-Object -ExpandProperty FullName
					$polCleanFile = Get-ChildItem -Path "$dirSupportFiles\Cleanup-GPOConfig" | Where-Object {$_.Name -like "*.txt"} | Select-Object -ExpandProperty FullName
					Execute-Process -Path "$envSystem32Directory\WindowsPowerShell\v1.0\powershell.exe" -Parameters "-File `"$polCleanScript`" -SourceFile `"$polCleanFile`""
					Break
				}
			}
			If (Test-Path -Path $dirLync) {
                Write-Log -Message "Microsoft Lync 2010 was detected and will be uninstalled." -Source $deployAppScriptFriendlyName
				foreach ($Path in $RegUninstallPaths) { 
		    		if (Test-Path $Path) { 
		       			 Get-ChildItem $Path | Where $UninstallSearchFilter |  
		        Foreach { Start-Process "$env:systemdrive\Windows\System32\msiexec.exe" "/x $($_.PSChildName) /qn /norestart" -Wait} 
		    		} 
				}
			}
			If (Test-Path -Path $dirCommunicator) {
                Write-Log -Message "Microsoft Communicator 2007 was detected and will be uninstalled." -Source $deployAppScriptFriendlyName
				$UninstallSearchFilter = { ($_.GetValue('DisplayName') -like '*Communicator 2007*') } 
				foreach ($Path in $RegUninstallPaths) { 
		    		if (Test-Path $Path) { 
		       			 Get-ChildItem $Path | Where $UninstallSearchFilter |  
		        Foreach { Start-Process "$env:systemdrive\Windows\System32\msiexec.exe" "/x $($_.PSChildName) /qn /norestart" -Wait} 
		    		} 
				}
			}
		
			ForEach ($officeExecutable in $officeExecutables) {
				If (Test-Path -Path (Join-Path -Path $dirOffice -ChildPath "Office15\$officeExecutable") -PathType Leaf) {
					Write-Log -Message 'Microsoft Office 2013 was detected. Will be uninstalled.' -Source $deployAppScriptFriendlyName
					Execute-Process -Path "cscript.exe" -Parameters "`"$dirSupportFiles\OffScrub13.vbs`" ProPlus /S /Q /NoCancel" -WindowStyle Hidden -IgnoreExitCodes '1,2,3'
					Write-Log -Message "Removing old policy settings if present."
					$polCleanScript = Get-ChildItem -Path "$dirSupportFiles\Cleanup-GPOConfig" | Where-Object {$_.Name -like "*.ps1"} | Select-Object -ExpandProperty FullName
					$polCleanFile = Get-ChildItem -Path "$dirSupportFiles\Cleanup-GPOConfig" | Where-Object {$_.Name -like "*.txt"} | Select-Object -ExpandProperty FullName
					Execute-Process -Path "$envSystem32Directory\WindowsPowerShell\v1.0\powershell.exe" -Parameters "-File `"$polCleanScript`" -SourceFile `"$polCleanFile`""
					Break
				}
			}
			ForEach ($officeExecutable in $officeExecutables) {
				If (Test-Path -Path (Join-Path -Path $dirOffice -ChildPath "Office16\$officeExecutable") -PathType Leaf) {
					Write-Log -Message 'Microsoft Office 2016 was detected and will be uninstalled.' -Source $deployAppScriptFriendlyName
					Execute-Process -Path "cscript.exe" -Parameters "`"$dirSupportFiles\OffScrub16.vbs`" ProPlus /S /Q /NoCancel" -WindowStyle Hidden -IgnoreExitCodes '1,2,3'
					Write-Log -Message "Removing old policy settings if present."
					$polCleanScript = Get-ChildItem -Path "$dirSupportFiles\Cleanup-GPOConfig" | Where-Object {$_.Name -like "*.ps1"} | Select-Object -ExpandProperty FullName
					$polCleanFile = Get-ChildItem -Path "$dirSupportFiles\Cleanup-GPOConfig" | Where-Object {$_.Name -like "*.txt"} | Select-Object -ExpandProperty FullName
					Execute-Process -Path "$envSystem32Directory\WindowsPowerShell\v1.0\powershell.exe" -Parameters "-File `"$polCleanScript`" -SourceFile `"$polCleanFile`""
					Break
				}
			}
		
		#Import Certificates
		$RootCerts = Get-ChildItem "$dirSupportFiles\Certificates\Root" | Where-Object {$_.Name -like "*.cer"} | Select-Object -ExpandProperty FullName
        $IntermediateCerts = Get-ChildItem "$dirSupportFiles\Certificates\Intermediate" | Where-Object {$_.Name -like "*.cer"} | Select-Object -ExpandProperty FullName
		$trustedPubCerts = Get-ChildItem "$dirSupportFiles\Certificates\TrustedPub" | Where-Object {$_.Name -like "*.cer"} | Select-Object -ExpandProperty FullName
		Write-Log -Message "Root certs are $RootCerts"
		Write-Log -Message "Intermediate certs are $IntermediateCerts"
		Write-Log -Message "Trusted Publisher certs are $trustedPubCerts"
		
        ForEach ($RootCert in $RootCerts) {
			$rootCertPrint = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
			$rootCertPrint.Import($RootCert)
			$rootMatch = Get-ChildItem -Path "cert:\LocalMachine\Root" | Where-Object { $_.Thumbprint -eq $rootCertPrint.Thumbprint }
			if ($rootMatch) {
				Write-Log -Message "$RootCert is already in the store. Skipping."
			} else {
	            Try {
	                certutil.exe -addstore -f 'root' `"$RootCert`" 
	                Write-Log -Message "Successfully imported `"$RootCert`""
	            } catch {
	                Write-Log -Message "Failed to import `"$RootCert`""
	            }					
			}
        }

        ForEach ($IntermediateCert in $IntermediateCerts) {
			$intCertPrint = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
			$intCertPrint.Import($IntermediateCert)
			$intMatch = Get-ChildItem -Path "cert:\LocalMachine\CA" | Where-Object { $_.Thumbprint -eq $intCertPrint.Thumbprint }
			if ($intMatch) {
				Write-Log -Message "$IntermediateCert is already in the store. Skipping."
			} else {
	            Try {
	                certutil.exe -addstore -f 'ca' `"$IntermediateCert`" 
	                Write-Log -Message "Successfully imported `"$IntermediateCert`""
	            } catch {
	                Write-Log -Message "Failed to import `"$IntermediateCert`""
	            }					
			}
        }

        ForEach ($trustedPubCert in $trustedPubCerts) {
			$tpCertPrint = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
			$tpCertPrint.Import($trustedPubCert)
			$tpMatch = Get-ChildItem -Path "cert:\LocalMachine\TrustedPublisher" | Where-Object { $_.Thumbprint -eq $tpCertPrint.Thumbprint }
			if ($tpMatch) {
				Write-Log -Message "$trustedPubCert is already in the store. Skipping."
			} else {
	            Try {
	                certutil.exe -addstore -f 'TrustedPublisher' `"$trustedPubCert`" 
	                Write-Log -Message "Successfully imported `"$trustedPubCert`""
	            } catch {
	                Write-Log -Message "Failed to import `"$trustedPubCert`""
	            }					
			}
        }			
		
		
		
		##*===============================================
		##* INSTALLATION
		##*===============================================
		[string]$installPhase = 'Installation'
		
		
		## Check whether running in Add Components Only mode
		If (-not $addComponentsOnly) {
	  		Show-InstallationProgress -StatusMessage 'Installing Office Professional. This may take some time. Please wait...'
			Write-Log -Message "Beginning Office 2016 installation."
			Execute-Process -Path "$dirFiles\Setup.exe" -Parameters "/adminfile `"$dirFiles\Config\Office2016ProPlus.MSP`" /config `"$dirFiles\ProPlus.WW\Config.xml`"" -WindowStyle Hidden -IgnoreExitCodes '3010'
		}
		
		
		##*===============================================
		##* POST-INSTALLATION
		##*===============================================
		[string]$installPhase = 'Post-Installation'
		
		# Enterprise Product Team additions
		# Copy Admin Templates and apply policy

		$PolDef = [environment]::ExpandEnvironmentVariables('%systemroot%\PolicyDefinitions')
		$PolDefLang = [environment]::ExpandEnvironmentVariables('%systemroot%\PolicyDefinitions\en-us')
		$vsupdates = Get-ChildItem $dirSupportfiles\VSUpdate\ | where {$_.name -like "*.exe"} | Select-Object -ExpandProperty Name 
		If ($vsupdates) {
			foreach ($vsupdate in $vsupdates) {
		Write-Log -Message "Installing Visual Studio Tools update." 
		Write-Log -Message "Update installer is $vsupdate."
		Execute-Process -Path "$dirSupportFiles\VSUpdate\$vsupdate" -Parameters "/q" -WindowStyle Hidden
		}
			}
		Write-Log -Message "Beginning Visual Studio Tools 2010 patch."
		Write-Log -Message "Copying ADMX and ADML files to Group Policy definitions."
		Write-Log -Message "Policy definitions folder is $PolDef."
		Write-Log -Message "Language folder is $PolDefLang."
		xcopy $dirSupportFiles\ADMX $PolDef /i /y /c
		xcopy $dirSupportFiles\ADML $PolDefLang /i /y /c
		$machpols = Get-ChildItem $dirSupportfiles | Where {$_.name -like "*mach.pol"} | Select -Expand Name
			If ($machpols){
			Write-Log -Message "Creating Machine Group Policy settings."
			foreach ($machpol in $machpols){
			Execute-Process -Path "$dirSupportFiles\ImportRegPol.exe" -Parameters "-m $machpol"
				}
			}
		$userpols = Get-ChildItem $dirSupportfiles | Where {$_.name -like "*user.pol"} | Select -Expand Name
			If ($userpols){
			Write-Log -Message "Creating User Group Policy settings."
			foreach ($userpol in $userpols){
			Execute-Process -Path "$dirSupportFiles\ImportRegPol.exe" -Parameters "-u $userpol"
				}
			}
	
		#Run token activation 
	
		$activationFile = Get-ChildItem $dirSupportFiles | Where-Object {$_.name -like "*.xrm-ms"} 
		Write-Log -Message "Token Activation file is $activationFile"
		Write-Log -Message "Running token activation."
		If((Get-WmiObject win32_operatingsystem).OSArchitecture -like "64-bit") {
			Execute-Process -Path "$envSystem32Directory\cscript.exe" -Parameters "`"$envProgramFilesX86\Microsoft Office\Office16\ospp.vbs`" /inslic:`"$dirSupportFiles\$activationFile`""
		} else {
			Execute-Process -Path "$envSystem32Directory\cscript.exe" -Parameters "`"$envProgramFiles\Microsoft Office\Office16\ospp.vbs`" /inslic:`"$dirSupportFiles\$activationFile`""
		}
	
		# Activate Office components (if running as a user)
		If (-not $osdMode) {
			If (Test-Path -Path (Join-Path -Path $dirOffice -ChildPath 'Office15\OSPP.VBS') -PathType Leaf) {
				Write-Log -Message "Activating Office 2016."
				Show-InstallationProgress -StatusMessage 'Activating Microsoft Office components. This may take some time. Please wait...'
				Execute-Process -Path 'cscript.exe' -Parameters "`"$dirOffice\Office15\OSPP.VBS`" /ACT" -WindowStyle Hidden
			}
		}
		
		# Prompt for a restart (if running as a user, not installing components and not running on a server)
		If ((-not $addComponentsOnly) -and ($deployMode -eq 'Interactive') -and (-not $IsServerOS)) {
			Show-InstallationRestartPrompt
		}
	
	#Versioning information
	Write-Log -Message "Writing versioning information to the registry"
	Set-RegistryKey -Key "HKEY_LOCAL_MACHINE\SOFTWARE\USAF\SDC\Applications" -Name "SdcNiprAndSiprMicrosoftOffice2016Version" -Value $PkgVersion -Type DWORD
	Write-Log -Message "Writing versioning information to WMI"
	$mofFile = Get-ChildItem -Path "$dirSupportFiles" | Where-Object { $_.Name -like "*.mof" } | Select-Object -ExpandProperty FullName
	if ($envOSArchitecture -like "*64*") {
		Execute-Process -Path "$envWinDir\SysWOW64\wbem\mofcomp.exe" -Parameters "`"$mofFile`""
	} else {
		Execute-Process -Path "$envSystem32Directory\wbem\mofcomp.exe" -Parameters "`"$mofFile`""
	}
	
	} ElseIf ($deploymentType -ieq 'Uninstall') {
		##*===============================================
		##* PRE-UNINSTALLATION
		##*===============================================
		[string]$installPhase = 'Pre-Uninstallation'
		
		## Show Welcome Message, close Microsoft Office applications that cause uninstall to fail
		Show-InstallationWelcome -CloseApps 'DATABASECOMPARE,MSACCESS,EXCEL,INFOPATH,SETLANG,MSOUC,ONENOTE,OUTLOOK,POWERPNT,MSPUB,WINWORD,ONENOTEM,SPREADSHEETCOMPARE,COMMUNICATOR'
		
		## Show Progress Message (with the default message)
		Show-InstallationProgress
		
		
		##*===============================================
		##* UNINSTALLATION
		##*===============================================
		[string]$installPhase = 'Uninstallation'
		
		Write-Log -Message "Beginning Office 2016 removal."
		Execute-Process -Path "cscript.exe" -Parameters "`"$dirSupportFiles\OffScrub16.vbs`" ProPlus /S /Q /NoCancel" -WindowStyle Hidden -IgnoreExitCodes '1,2,3'
		if (Get-RegistryKey -Key "HKEY_LOCAL_MACHINE\SOFTWARE\USAF\SDC\Applications" -Value "SdcNiprAndSiprMicrosoftOffice2016Version") {
			Write-Log -Message "Removing versioning information from registry"
			Remove-RegistryKey -Key "HKEY_LOCAL_MACHINE\SOFTWARE\USAF\SDC\Applications" -Name "SdcNiprAndSiprMicrosoftOffice2016Version"
		}
		
		##*===============================================
		##* POST-UNINSTALLATION
		##*===============================================
		[string]$installPhase = 'Post-Uninstallation'
		
		## <Perform Post-Uninstallation tasks here>
	}

	##*===============================================
	##* END SCRIPT BODY
	##*===============================================

	## Call the Exit-Script function to perform final cleanup operations
	Exit-Script -ExitCode $mainExitCode
}
Catch {
	[int32]$mainExitCode = 69008
	[string]$mainErrorMessage = "$(Resolve-Error)"
	Write-Log -Message $mainErrorMessage -Severity 3 -Source $deployAppScriptFriendlyName
	Show-DialogBox -Text $mainErrorMessage -Icon 'Stop'
	Exit-Script -ExitCode $mainExitCode
}
