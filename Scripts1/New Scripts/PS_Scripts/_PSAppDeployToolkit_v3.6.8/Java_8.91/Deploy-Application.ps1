
Param (
	[ValidateSet("Install","Uninstall")] 
	[string]$DeploymentType = "Install",
	[ValidateSet("Interactive","Silent","NonInteractive")]
	[string]$DeployMode = "Interactive",
	[switch] $AllowRebootPassThru = $false,
	[switch] $TerminalServerMode = $false
)

#*===============================================
#* VARIABLE DECLARATION
Try {
#*===============================================

#*===============================================
# Variables: Application

$appx86Name = "jre1.8.0_91_x86.msi"
$appx64Name = "jre1.8.0_91_x64.msi"
$appVersion = [Version]"8.0.910.14"

#*===============================================

$appVendor = "Oracle"
$appArch = ""
$appLang = "EN"
$appRevision = "01"
$appScriptVersion = "1.0.0"
$appScriptDate = "03/30/2016"
$appScriptAuthor = "Mike Stauskas"

#*===============================================
# Variables: Script - Do not modify this section

$deployAppScriptFriendlyName = "Deploy Application"
$deployAppScriptVersion = [version]"3.1.2"
$deployAppScriptDate = "03/30/2016"
$deployAppScriptParameters = $psBoundParameters

# Variables: Environment
$scriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Definition
# Dot source the App Deploy Toolkit Functions
."$scriptDirectory\AppDeployToolkit\AppDeployToolkitMain.ps1"
."$scriptDirectory\SupportFiles\Get-ApplicationInfo.ps1"
."$scriptDirectory\SupportFiles\Get-PendingReboot.ps1"

#*===============================================
#* END VARIABLE DECLARATION
#*===============================================

#*===============================================
#* PRE-INSTALLATION
If ($deploymentType -ne "uninstall") { $installPhase = "Pre-Installation"
#*===============================================

    # Is reboot pending
    
    if ($(Get-PendingReboot).RebootPending) {  
        
            Write-Log "The system is pending reboot from a previous install or uninstall."
        
    }

    # Prompt the user to close the following applications if they are running:
    
    Show-InstallationWelcome -CloseApps "firefox,iexplore,chrome,jusched,jucheck,jqs,winword,excel" -AllowDefer -DeferTimes 3 -CloseAppsCountdown "120"
    
    # Show Progress Message (with the default message)
    
    Show-InstallationProgress 
    
    # Method 1: Use tool-kit method
    
    Write-Log "Uninstall method 1: Use tool-kit method"
    
    # Remove any Java Auto Updater installations
    
    Remove-MSIApplications "Java Auto Updater"
    
    # Remove all java version 6 and Java versions 7
    
    Remove-MSIApplications "Java(TM) 6 Update"
    
    Remove-MSIApplications "Java 6 Update"
    
    Remove-MSIApplications "Java(TM) 5 Update"
    
    Remove-MSIApplications "Java 5 Update"
    
    Remove-MSIApplications "Java(TM) 4 Update"
    
    Remove-MSIApplications "Java 4 Update"

    Remove-MSIApplications "Java 8 Update"
    

    
    # Method 2 If the above process did not succeed. The script will execute the below granular uninstallation procedure. Java versions 6 - n are handled using the below procedure. 
    
    Write-Log " Method 2 - uninstall string method"
    
        Execute-MSI -Action Uninstall -Path "{26A24AE4-039D-4CA4-87B4-2F86416000FF}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{26A24AE4-039D-4CA4-87B4-2F86416001FF}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{26A24AE4-039D-4CA4-87B4-2F86416002FF}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{26A24AE4-039D-4CA4-87B4-2F86416003FF}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{26A24AE4-039D-4CA4-87B4-2F86416004FF}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{26A24AE4-039D-4CA4-87B4-2F86416005FF}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{26A24AE4-039D-4CA4-87B4-2F86416006FF}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{26A24AE4-039D-4CA4-87B4-2F86416007FF}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{26A24AE4-039D-4CA4-87B4-2F86416008FF}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{26A24AE4-039D-4CA4-87B4-2F86416009FF}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{26A24AE4-039D-4CA4-87B4-2F86416010FF}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{26A24AE4-039D-4CA4-87B4-2F86416011FF}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{26A24AE4-039D-4CA4-87B4-2F86416012FF}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{26A24AE4-039D-4CA4-87B4-2F86416013FF}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{26A24AE4-039D-4CA4-87B4-2F86416014FF}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{26A24AE4-039D-4CA4-87B4-2F86416015FF}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{26A24AE4-039D-4CA4-87B4-2F86416016FF}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{26A24AE4-039D-4CA4-87B4-2F86416017FF}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{26A24AE4-039D-4CA4-87B4-2F86416018FF}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{26A24AE4-039D-4CA4-87B4-2F86416019FF}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{26A24AE4-039D-4CA4-87B4-2F86416020FF}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{26A24AE4-039D-4CA4-87B4-2F86416021FF}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{26A24AE4-039D-4CA4-87B4-2F86416022FF}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{26A24AE4-039D-4CA4-87B4-2F86416023FF}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{26A24AE4-039D-4CA4-87B4-2F86416024FF}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{26A24AE4-039D-4CA4-87B4-2F86416025FF}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{26A24AE4-039D-4CA4-87B4-2F86416026FF}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{26A24AE4-039D-4CA4-87B4-2F86416027FF}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{26A24AE4-039D-4CA4-87B4-2F86416028FF}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{26A24AE4-039D-4CA4-87B4-2F86416029FF}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{26A24AE4-039D-4CA4-87B4-2F86416030FF}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{26A24AE4-039D-4CA4-87B4-2F86416031FF}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{26A24AE4-039D-4CA4-87B4-2F86416032FF}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{26A24AE4-039D-4CA4-87B4-2F86416033FF}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{26A24AE4-039D-4CA4-87B4-2F86416034FF}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{26A24AE4-039D-4CA4-87B4-2F86416035FF}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{26A24AE4-039D-4CA4-87B4-2F86416036FF}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{26A24AE4-039D-4CA4-87B4-2F86416037FF}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{7148F0A8-6813-11D6-A77B-00B0D0142000}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{7148F0A8-6813-11D6-A77B-00B0D0142010}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{7148F0A8-6813-11D6-A77B-00B0D0142020}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{7148F0A8-6813-11D6-A77B-00B0D0142030}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{7148F0A8-6813-11D6-A77B-00B0D0142040}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{7148F0A8-6813-11D6-A77B-00B0D0142050}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{7148F0A8-6813-11D6-A77B-00B0D0142060}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{7148F0A8-6813-11D6-A77B-00B0D0142070}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{7148F0A8-6813-11D6-A77B-00B0D0142080}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{7148F0A8-6813-11D6-A77B-00B0D0142090}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{7148F0A8-6813-11D6-A77B-00B0D0142100}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{7148F0A8-6813-11D6-A77B-00B0D0142110}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{7148F0A8-6813-11D6-A77B-00B0D0142120}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{7148F0A8-6813-11D6-A77B-00B0D0142130}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{7148F0A8-6813-11D6-A77B-00B0D0142140}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{7148F0A8-6813-11D6-A77B-00B0D0142150}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{7148F0A8-6813-11D6-A77B-00B0D0142160}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{7148F0A8-6813-11D6-A77B-00B0D0142170}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{7148F0A8-6813-11D6-A77B-00B0D0142180}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{7148F0A8-6813-11D6-A77B-00B0D0142190}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{3248F0A8-6813-11D6-A77B-00B0D0150000}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{3248F0A8-6813-11D6-A77B-00B0D0150010}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{3248F0A8-6813-11D6-A77B-00B0D0150020}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{3248F0A8-6813-11D6-A77B-00B0D0150030}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{3248F0A8-6813-11D6-A77B-00B0D0150040}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{3248F0A8-6813-11D6-A77B-00B0D0150050}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{3248F0A8-6813-11D6-A77B-00B0D0150060}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{3248F0A8-6813-11D6-A77B-00B0D0150070}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{3248F0A8-6813-11D6-A77B-00B0D0150080}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{3248F0A8-6813-11D6-A77B-00B0D0150090}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{3248F0A8-6813-11D6-A77B-00B0D0150100}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{3248F0A8-6813-11D6-A77B-00B0D0150110}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{3248F0A8-6813-11D6-A77B-00B0D0150120}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{3248F0A8-6813-11D6-A77B-00B0D0150130}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{3248F0A8-6813-11D6-A77B-00B0D0150140}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{3248F0A8-6813-11D6-A77B-00B0D0150150}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{3248F0A8-6813-11D6-A77B-00B0D0150160}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{3248F0A8-6813-11D6-A77B-00B0D0150170}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{3248F0A8-6813-11D6-A77B-00B0D0150180}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{3248F0A8-6813-11D6-A77B-00B0D0150190}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{3248F0A8-6813-11D6-A77B-00B0D0150200}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{3248F0A8-6813-11D6-A77B-00B0D0150210}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{3248F0A8-6813-11D6-A77B-00B0D0150220}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{3248F0A8-6813-11D6-A77B-00B0D0160000}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{3248F0A8-6813-11D6-A77B-00B0D0160010}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{3248F0A8-6813-11D6-A77B-00B0D0160020}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{3248F0A8-6813-11D6-A77B-00B0D0160030}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{3248F0A8-6813-11D6-A77B-00B0D0160040}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{3248F0A8-6813-11D6-A77B-00B0D0160050}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{3248F0A8-6813-11D6-A77B-00B0D0160060}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{3248F0A8-6813-11D6-A77B-00B0D0160070}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{3248F0A8-6813-11D6-A77B-00B0D0160080}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{3248F0A8-6813-11D6-A77B-00B0D0160090}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{3248F0A8-6813-11D6-A77B-00B0D0160100}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{3248F0A8-6813-11D6-A77B-00B0D0160110}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{3248F0A8-6813-11D6-A77B-00B0D0160120}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{3248F0A8-6813-11D6-A77B-00B0D0160130}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{3248F0A8-6813-11D6-A77B-00B0D0160140}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{3248F0A8-6813-11D6-A77B-00B0D0160150}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{3248F0A8-6813-11D6-A77B-00B0D0160160}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{3248F0A8-6813-11D6-A77B-00B0D0160170}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{3248F0A8-6813-11D6-A77B-00B0D0160180}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{3248F0A8-6813-11D6-A77B-00B0D0160190}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{3248F0A8-6813-11D6-A77B-00B0D0160200}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{3248F0A8-6813-11D6-A77B-00B0D0160210}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{3248F0A8-6813-11D6-A77B-00B0D0160220}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{26A24AE4-039D-4CA4-87B4-2F83216023FF}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{26A24AE4-039D-4CA4-87B4-2F83216024FF}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{26A24AE4-039D-4CA4-87B4-2F83216025FF}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{26A24AE4-039D-4CA4-87B4-2F83216026FF}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{26A24AE4-039D-4CA4-87B4-2F83216027FF}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{26A24AE4-039D-4CA4-87B4-2F83216028FF}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{26A24AE4-039D-4CA4-87B4-2F83216029FF}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{26A24AE4-039D-4CA4-87B4-2F83216030FF}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{26A24AE4-039D-4CA4-87B4-2F83216031FF}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{26A24AE4-039D-4CA4-87B4-2F83216032FF}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{26A24AE4-039D-4CA4-87B4-2F83216033FF}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{26A24AE4-039D-4CA4-87B4-2F83216034FF}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{26A24AE4-039D-4CA4-87B4-2F83216035FF}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{26A24AE4-039D-4CA4-87B4-2F83216036FF}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        Execute-MSI -Action Uninstall -Path "{26A24AE4-039D-4CA4-87B4-2F83216037FF}" -Parameters "REBOOT=ReallySuppress /QN" -ContinueOnError $true
        
        
    # Method 3: Use powershell and WMI class win32_product
    
    Write-Log "Unistall method 3: Use powershell and WMI class win32_product"
    
    $jUpdateinstalls = Get-WmiObject  -Query "Select * from Win32_Product Where Name like '%Java Auto Updater%'"
    
    if ($jUpdateinstalls) {
    
        foreach ($jUpdateinstall in $jUpdateinstalls) {
        
            Write-Log "Uninstalling $($jUpdateinstall.Name)"
        
            $jUpdateinstall.Uninstall()
        
        }
        
    }
    
    
    
    $jinstalls = Get-WmiObject  -Query "Select * from Win32_Product Where Name like '%Java%Update %'"
    
    if ($jinstalls) {
    
        foreach ($jinstall in $jinstalls) {
        
        $instVersion = [Version]$jinstall.Version
        
            if ($instVersion -lt $appVersion) {
            
                Write-Log "Uninstalling $($jinstall.Name)"
            
                $jinstall.Uninstall()
                
                $RegUninsentries = (Get-ChildItem HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall -Recurse | Get-ItemProperty -name DisplayName -ErrorAction SilentlyContinue |Where-Object {$_.DisplayName -like "*Java*6*update*"}).PSPath
                
                if ($RegUninsentries) {
                
                    if (Test-Path -Path "$RegUninsentries") {Remove-Item "$RegUninsentries" -Force -ErrorAction silentlycontinue}
                    
                }
                
                if (Test-Path -Path "$envProgramFiles\Java\jre6") {Remove-Folder -Path "$envProgramFiles\Java\jre6" -ContinueOnError $true}

                $RegUninsentries86 = (Get-ChildItem HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall -Recurse | Get-ItemProperty -name DisplayName -ErrorAction SilentlyContinue |Where-Object {$_.DisplayName -like "*Java*6*update*"}).PSPath
                
                if ($RegUninsentries86) {
                
                    if (Test-Path -Path "$RegUninsentries86") {Remove-Item "$RegUninsentries86" -Force -ErrorAction silentlycontinue}
                    
                }
                
                if (Test-Path -Path "$envProgramFilesX86\Java\jre6") {Remove-Folder -Path "$envProgramFilesX86\Java\jre6" -ContinueOnError $true}

            
            } else {
            
                Write-Log "An equal or greater version of Java is already installed on this machine. Installed version is $instVersion"
                
                Exit-Script -ExitCode 0
            
            }
        
        }

    }


#*===============================================
#* INSTALLATION 
$installPhase = "Installation"
#*===============================================

    # Install the Java MSI file turn off Autoupdate, disable webstarticon, skip EULA and disable systray icon.


	Write-Log "Installing Java Configuration Files"

	New-Folder -Path "$envWINDIR\SUN\JAVA\DEPLOYMENT" -ContinueOnError $true
	Remove-File -Path "$envWINDIR\SUN\JAVA\DEPLOYMENT" -Recurse -ContinueOnError $true
	Copy-File -Path "$scriptDirectory\JavaConfigFiles" -Destination "$envWINDIR\SUN\JAVA\DEPLOYMENT" -Recurse -ContinueOnError $true
	Start-Sleep -s 10

    
        Write-Log "Installing 64 bit version of java"

#        Execute-Process -FilePath "setupX64.exe" -Arguments "/s" -WindowStyle Hidden -ContinueOnError $true
	Execute-MSI -Action Install -Path "$appx64Name" -Parameters "AUTOUPDATECHECK=0 JAVAUPDATE=0 RebootYesNo=No REBOOT=ReallySuppress /QN" -ContinueOnError $true

        Write-Log "Installing 32 bit version of java"

#        Execute-Process -FilePath "setupX86.exe" -Arguments "/s" -WindowStyle Hidden
	Execute-MSI -Action Install -Path "$appx86Name" -Parameters "AUTOUPDATECHECK=0 JAVAUPDATE=0 RebootYesNo=No REBOOT=ReallySuppress /QN" -ContinueOnError $true


#*===============================================
#* POST-INSTALLATION
$installPhase = "Post-Installation"
#*===============================================

$UnInstallAU = Get-WmiObject -Class win32_product|Where-Object {$_.name -like 'java*Auto*Updater*'}

if ($UnInstallAU) {

    $UnInstallAU.Uninstall()

}


#*===============================================
#* UNINSTALLATION
} ElseIf ($deploymentType -eq "uninstall") { $installPhase = "Uninstallation"
#*===============================================

    # Prompt the user to close the following applications if they are running:
    Show-InstallationWelcome -CloseApps "firefox,iexplore,chrome,jusched,jucheck,jqs,winword,excel" -AllowDefer -DeferTimes 3 -CloseAppsCountdown "120"
    # Show Progress Message (with a message to indicate the application is being uninstalled)
    Show-InstallationProgress -StatusMessage "Uninstalling Application $installTitle. Please Wait..." 
    # Remove this version of Java 7 Update 60
    
	$jinstalls = Get-WmiObject  -Query "Select * from Win32_Product Where Name like '%Java%Update %'"
    
    if ($jinstalls) {
    
        foreach ($jinstall in $jinstalls) {
        
        $instVersion = [Version]$jinstall.Version
        
            if ($instVersion -eq $appVersion) {
            
                Write-Log "Uninstalling $($jinstall.Name)"
            
                $jinstall.Uninstall()

            
            } else {
            
                Write-Log "Found version $instVersion installed. But it will not be uninstalled because it is not equal to $appVersion"  
            
            }
        
        }

    }

#*===============================================
#* END SCRIPT BODY
} } Catch {$exceptionMessage = "$($_.Exception.Message) `($($_.ScriptStackTrace)`)"; Write-Log "$exceptionMessage"; Exit-Script -ExitCode 1} # Catch any errors in this script 
Exit-Script -ExitCode 0 # Otherwise call the Exit-Script function to perform final cleanup operations
#*===============================================