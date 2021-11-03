<#
.Synopsis
   Upgrade-Staging - Ugrades SDC 5.1/5.2 to SDC 5.3.1.
.DESCRIPTION
   This script upgrades USAF SDC 5.x computers to SDC 5.3.1. Though originally intended to be ran in conjuntion with SCCM 2007 in a Task Sequence, this script can also be ran from a file
   share or USB drive. 

   Several Pre-Flight checks are done before setup.exe is launched. Checks include, is the computer on battery power, sufficent disk space, compatiable HBSS versions, current OS version,  
   clone tags, valid source directories and supported computer model. After pre-flights are validated, source files are copied via robocopy to the local disk, a splash screen is started
   warning the user that an OS upgrade has started and setup.exe is executed.

.EXAMPLE
   .\Upgrade-Staging.ps1 -scanonly
   Only runs pre-flight checks and writes output to the screen.
.EXAMPLE
   .\Upgrade-Staging.ps1 -SourcePath \\{Shared Drive}\{Shared Folder}
   In this instance the script is local to the computer being upgraded but the source files are on a network path.

.EXITCODES
1099 : Administrative Rights required
1101 : Source Directories empty or not found
1102 : Target Directory not found
1103 : Robocopy failed
1109 : Device failed pre-flight checks
1200 : Insufficent Disk Space

.CHANGELOG:
1.0.0
- Inital test build

1.0.1
- Updated McAfee version checks from single integer to an array comparison.  See new Check-Version function.
- Added SHB and SDC registry key check to preflight checks 
- Added Clone Tag key check to preflight checks
- Added parameter clean up to address extra "\"
- Updated Write-Verbose structure for readability
- Changed File copy log to include Computer Name and Date
- Added OS Version check to preflight checks

1.0.2
- Updated McAfee paths to use $env:ProgramFile variables
- Updated Check-Version to address comparison bug (9 and 10)
- Added function to check file version (required for VSE and HIPS)
- Removed $sourcepath mandatory requirement.  Added the $scriptdir to support options.  Once set, Push-location allows a .\ reference.
- Aligned Upgrade_OS and OS_Upgrade terminology

1.0.3
-Updated logic Switch statement to evaluate logic statements
-Updated Reg path conditional to evaluate logic statement
-Uncommented original setup.exe command line. New command seemed to be invaild, cause setup.exe to hang

1.0.4
- Added -quiet and -copylogs option to setup.exe
- Moved the -doUpgrade process to only run if -scanOnly is not selected. 

1.0.5
-Removed "or" statments in Switch statment. This cause invalid fall through
-Editted /installDrivers path to the Drivers folder in setup.exe
-Commented out SDC and SHB Key check, still returns false
-Added trim() to sysmodel variable

1.0.6
-Added Splash Screen
-Precheck AC Power applied
-Precheck Local C: Drive Storage space > 15GB
-Tracked Target Destination

1.1.0
-Reorganized Writing Logs to 1 function to do both Host and Log file
-Cleaned up logic holes dealing with null values
-Cleaned up unnessecary code

1.1.1
- Updated McAfee DLP and Agent Checks
- Corrected log entry on line 539 to reflect HIPS vice VSE
- Added $dgReady variable to correct logic in line 577

1.1.2
-Updated env:ProgramFiles and (x86) to hard paths
-Moved SourceFile log logic higher in script
-Modified Task Scheduler to 60 secs
#>

param(
[parameter(Mandatory=$false)][ValidateScript({Test-Path $_ -PathType ‘Container’})][string]$SourcePath,
[parameter(Mandatory=$false)][string]$TargetPath = $env:SystemDrive+"\Upgrade_OS",
[parameter(Mandatory=$false)][string]$myLogPath = $env:SystemDrive+"\Upgrade_OS_Logs",
[parameter(Mandatory=$false)][string]$myLogName = $env:COMPUTERNAME+"_"+(get-date -Uformat %Y%m%d)+"_Upgrade_OS.log",
[parameter(Mandatory=$false)][string]$myCopyLogName = $env:COMPUTERNAME+"_"+(get-date -Uformat %Y%m%d)+"_File_Copy.log",
[parameter(Mandatory=$false)][switch]$doUpgrade =$true,
[parameter(Mandatory=$false)][switch]$scanOnly
)

#Parameter cleanup
#check for and remove last "\" on Source, Target, or Log paths
#May not be needed but helps with string builds later in script.

if ($SourcePath.endsWith("\") -eq $True){
  $SourcePath = $SourcePath-replace ".{1}$"
}
if ($TargetPath.endsWith("\") -eq $True){
  $TargetPath = $TargetPath-replace ".{1}$"
}
if ($myLogPath.endsWith("\") -eq $True){
  $myLogPath = $myLogPath-replace ".{1}$"
}

# Set Script Working Directory
# If no source path isprovided set to script directory
if($SourcePath -eq "")
{
    $scriptPath = $MyInvocation.MyCommand.Path
    $scriptDir = Split-Path $scriptPath
    
}
else
{
    $scriptDir = $SourcePath
}

$callingScript = $myInvocation.MyCommand.Name

# Temporarily change to the script folder
Push-Location $scriptDir

#Check for Administrative rights and exit
If (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))
{
    Write-Warning $callingScript + " : This script requires elevated rights. Please re-run this script as an administrator"
    Pop-Location
    Exit 1099
}

# Define Standard Functions
function WriteLog($log,$errorLevel){
    Write-PSLog -message ($callingScript + " : $log") -component "Main()" -type $errorLevel
    if($errorLevel -eq 3)
    {
        Write-Host "$log" -ForegroundColor "red"
    }
    else
    {
        Write-Host "$log"
    }
}

# EvaluatePSError is a standard process to evaluate errors in PS command lines and log to standard log location
function EvaluatePSError($err){
    if ($err -ne $null){
        WriteLog "Completed with Error: $err" 3
    } else {
        WriteLog "Completed Successfully" 1 
    }
}

function ValidateSource($myPath){
    #Function will check for existence of directory and that it is not empty.
    #This function validates source file structure and exits script if not found.
    if($myPath -eq $null)
    {
           WriteLog "Path is null. Exiting Script." 3
           Pop-Location
           Exit 1101  
    }

    if(Test-Path $myPath){
        if ((Get-ChildItem $myPath) -ne $null){
           WriteLog "Check Passed. The folder $myPath exists and is not empty" 1
        } else {
           WriteLog "Check Failed. The folder $myPath is empty. Exiting Script." 3
           Pop-Location
           Exit 1101
        }
    } else {
        WriteLog "Check Failed. The folder $myPath is does not exist. Exiting Script" 3
        Pop-Location
        Exit 1101
    }
}

function TestFilePath ($myPath){
    #Function to test for existence of a file path.  Returns true or false value.
    $currStep = "Check for "+ $myPath
    $state = Test-Path $myPath -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -ErrorVariable myErr
    EvaluatePSError $myErr
    WriteLog "Check for $myPath is $state" 1  
    Return $state
}

function Get-DriverPath($sysmodel){
if($sysmodel -eq $null)
{
    WriteLog "System model was captured as null" 3
    Pop-location
    Exit 1101
}

Switch ($sysmodel) {
    #The platforms below are supported by the SDC 5.1/5.2 to 5.3.1 Upgrade script
    
    # HP-ProBook 640 G2
    "HP ProBook 640 G2"{
        $driverpath = "$Drivers\HP - ProBook 640"
    }
    #HP-Z240 Workstation
    "HP Z240 Tower Workstation"{
        $driverpath = "$Drivers\HP - Z240"
    }

    # HP- EliteDesk 705 G2 Mini or SFF or G3 Mini
    "HP EliteDesk 705 G2 MINI"{
        $driverpath = "$Drivers\HP - EliteDesk 705"
    }
    "HP EliteDesk 705 G3 DESKTOP MINI"{
        $driverpath = "$Drivers\HP - EliteDesk 705"
    }
    "HP EliteDesk 705 G2 SFF"{
        $driverpath = "$Drivers\HP - EliteDesk 705"
    }


    # HP-Z840 Workstation
    "HP Z840 Workstation"{
        $driverpath = "$Drivers\HP - Z840"
    }
    # HP-EliteDesk 800 G2
    "HP EliteDesk 800 G2 SFF"{
        $driverpath = "$Drivers\HP - 800"
    }

    # HP-Elitebook 840 G2 or G3
    "HP EliteBook 840 G2"{
        $driverpath = "$Drivers\HP - EliteBook 840"
    }
    "HP EliteBook 840 G3"{
        $driverpath = "$Drivers\HP - EliteBook 840"
    }


    # HP-612 G1 Pro X2
    {$sysmodel -like "*612 G1"}{
        $driverpath = "$Drivers\HP - ProBook x2 612"
    }
    # HP-ZBook 15 G3
    "HP ZBook 15 G3"{
        $driverpath = "$Drivers\HP - ZBook 15"
    }
    # GETAC-B300
    "B300G5"{
        $driverpath = "$Drivers\Getac - B300G5"
    }
    # GETAC-V110
    {$sysmodel -like "*V110G2*"}{
        $driverpath = "$Drivers\Getac - V110G2"
    }
    # Surface Pro 3
    "Surface Pro 3"{
        $driverpath = "$Drivers\Microsoft - Surface Pro 3"
    }
    # Surface Pro 4
    "Surface Pro 4"{
        $driverpath = "$Drivers\Microsoft - Surface Pro 4"
    }
    #Surface Book
    "Surface Book"{
        $driverpath = "$Drivers\Microsoft - Surface Book"
    }
    #The platforms below are not supported by the SDC 5.1/5.2 to 5.3.1 upgrade process.  
    #These systems must be W&L with the HVCI kill script and a baseline SDC 5.3.1 installation
    (($sysmodel -like "%705 G1%") -or "HP Z230 Tower Workstation" -or "HP Z420 Workstation" -or "HP Z820 Workstation" -or "HP EliteBook 840 G1" -or "HP ProBook 640 G1" -or "HP ZBook 15 G2" -or "HP ZBook 15" -or "HP ZBook 14" -or "HP ZBook 17" -or "LIFEBOOK U745" -or "ThinkPad W541" -or "ThinkPad L440" -or "Precision T1700"){
        WriteLog "Check Failed. $sysmodel is a not a supported QEB hardware model for the USAF Windows 10 Upgrade Solution. Please see $supportURL for instructions to install SDC 5.3.1." 3
    }
    default{
        WriteLog "Check Failed. $sysmodel is not a supported QEB hardware model.  Please see $supportURL for instructions to install SDC 5.3.1." 3
    }
}

    return $driverpath
}

function File-Copy($Rargs) {
    if($Rargs -eq $null)
    {
        WriteLog "File Copy Args are null" 3
        return 404
    }
    WriteLog "Begin copying $currStep for the SDC 5.3.x Upgrade process. This process may take a significant amount of time. Do not close this window" 1
    if ($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent) {
        $Rargs = $Rargs + " /TEE"
    }
    WriteLog "Copying $currStep. Running robocopy.exe $Rargs" 1
    start-process robocopy -Wait -WindowStyle Hidden -ArgumentList $RArgs
    WriteLog "Completed copying $currStep for the SDC 5.3.x Upgrade process." 1
    return $lastexitcode
}


function Get-FileVersion($filepath){
    #Returns file version number if file exists
    if($filepath -eq $null)
    {
        WriteLog "Filepath is null in Get-FileVersion" 3
        return $null
    }
    if(Test-Path $filepath -PathType Leaf){
        (Get-ItemProperty $filepath ).VersionInfo.FileVersion  
    }
}

function Check-Version($currVer, $minVer){
    #Build arrays from dot values"
    if($currVer -eq $null)
    {
        WriteLog "Current version is null. Please check source is valid." 3
        return $false
    }

    if($minVer -eq $null)
    {
        WriteLog "Minimum version is null. Please check source is valid." 3
        return $false 
    }
    return $currVer -ge $minVer    
}

function CheckACPower(){
    if((Get-WMIObject -Class Win32_ComputerSystem).PCSystemType -eq 2)
        { 
            return (Get-WMIObject -Class BatteryStatus -namespace root\wmi).PowerOnline
        }
    return $true
       
}
function CheckDiskSpace(){
    $PackageSize = 15 * 1gb
    $LocalStorage = (Get-WMIObject -Class Win32_logicaldisk -filter "Drivetype=3").FreeSpace[0]
    return $PackageSize -lt $LocalStorage
}

# Create script Log working directories
if (-not(Test-Path $myLogPath)){
    $currStep = "Creating Script Log Directory."
    New-Item $myLogPath -type Directory -ErrorAction SilentlyContinue -ErrorVariable myErr | Out-Null
}

# Setup Powershell Logging
$currStep = "Import Logging Module"
Import-Module -Name "$scriptDir\Upgrade_OS\PostOOBE\Write-PSLogs\Write-PSLogs.psm1"
Set-PSLogPath -logPath $myLogPath -logName $myLogName
WriteLog "Script Log folder is $myLogPath" 1

#Define standard variables
$supportURL = "https://www.my.af.mil/"
$driverpath = $null
$min_vsever = "8.8.0.1599"
$min_hipver = "8.0.0.3828"
$min_dlpver = "10.0.100.37"
$min_mavver = "5.0.4.283"
$McAfeeReady = $true
$myVer = $PSVersionTable
$SDCKey = "HKLM:\SOFTWARE\USAF"
$SHBKey = "HKLM:\SOFTWARE\DOD"
$ACpower = $true
$InsuffSpace = $true

#Preflight checks
WriteLog "----------------------------------------------------" 1
WriteLog "Begin Preflight checks." 1


#Check Machine Plugged in
if(-not(CheckACPower))
{
    WriteLog "Please plug machine into AC Power" 3
    $ACPower = $false
}
else
{
    WriteLog "AC Power is Plugged in" 1
}
#Check Machine Storage Size
if(-not(CheckDiskSpace))
{
    WriteLog "Please make sure the machine has at least 15GB of Free Disk Space" 3
    $InsuffSpace = $false
}
else
{
    WriteLog "Enough Disk Size is Avaliable" 1
}

#validate Device is supported
$sysmodel = (Get-WmiObject -class win32_computersystem).Model.trim() 
$sysSKU = (Get-WmiObject Win32_ComputerSystem).SystemSKUNumber
$dg = (Get-CimInstance -className Win32_deviceguard -Namespace root\Microsoft\Windows\DeviceGuard).SecurityServicesRunning

WriteLog "System Information" 1 
WriteLog "System Model value = $sysmodel" 1
WriteLog "System SKU value = $sysSKU" 1
WriteLog "Virtual Secure Mode Services value = $dg" 1

WriteLog "**** Checking if device is supported for SDC upgrade" 1

# Set Base Driver Path
$Drivers = $scriptDir + "\Drivers"
$DBGflag = $false

#Get the driverpath for the specified system model
$driverpath = Get-DriverPath($sysmodel)

if ($driverpath -ne $null) {
    WriteLog "Check Passed. $sysmodel is a supported model for USAF Windows 10 Upgrade Solution." 1
    WriteLog "$sysmodel driver path is set to $driverpath" 1
}

#Validate OS Version
WriteLog "**** Checking OS version" 1
$osReady = $false
$myVerBuildVersion = $myVer.BuildVersion.toString()
WriteLog "OS Version Detected: $myVerBuildVersion" 1

if (($myver.BuildVersion.Major -lt 10)){
    WriteLog "Check Failed. OS VERSION IS NOT WINDOWS 10." 3
    $osReady = $false
} else {
    WriteLog "Windows 10 OS detected. Checking build version." 1
    #Windows 10 detected.  Check for build number  
    if (($myver.BuildVersion.Major -eq 10) -and ($myver.BuildVersion.Minor -eq 0) -and ($myver.BuildVersion.Build -eq 10586)){
        WriteLog "Check Passed. Windows 10 OS Version is supported for Upgrade." 1
        $osReady = $true
    } else {
        WriteLog "Check Failed. Windows 10 OS Version is not supported for Upgrade." 3
        $osReady = $false
    }
}

# Check for Clone Tag
WriteLog "**** Checking Clone Tag Values" 1
$cloneTag = $false
if ((Test-Path HKLM:\System\Setup -pathtype Container) -eq $true){
    $prop = Get-ItemProperty -Path HKLM:\System\Setup
    if ($prop){
        $mem = Get-Member -InputObject $prop -Name "CloneTag"
        if ($mem){
	        $cTag = Get-ItemPropertyValue -path HKLM:\System\Setup -Name "CloneTag"
            WriteLog "CloneTag registry entry found with a value of $cTag" 1
            Switch ($cTag){
                "Fri Jul 22 20:46:22 2016" {
                    WriteLog "Check Passed. SDC 5.2 Clone Tag Detected" 1
                    $cloneTag = $true
                }            
                "Wed Mar 23 18:20:12 2016" {
                    WriteLog "Check Passed. SDC 5.1 Clone Tag Detected" 1
                    $cloneTag = $true
                }
                default {
                    WriteLog "Check failed.  Clone Tag value does not match known SDC values." 3
                    $cloneTag = $false
                }
            }
        } else {
            WriteLog "Clone Tag check failed. Value not found.  Unknown Image Congfiguration." 3
            $cloneTag = $false
        }
    }
}

#Collect SHB Keys - Future checks.
<#
Write-Verbose "**** Checking SDC and SHB Registry Keys"
$SHB_Ready = $false
if(Test-Path -LiteralPath $SHBKey -pathtype container){
    Write-PSLog -message ($callingScript +" : SHB Key Found") -component "Main()" -type 1
    Write-Verbose "Check Passed. SHB Key Exists."
    $SHB_Ready = $true
}else{
    Write-PSLog -message ($callingScript +" : No SHB Keys detected. Value: " + $SHBKey) -component "Main()" -type 3
    Write-Verbose "Check Failed.  No SHB Keys detected."
    $SHB_Ready = $false 
}
#>

$SHB_Ready = $true

#Collect SDC Keys
<#
$SDC_Ready = $false
if(Test-Path -LiteralPath $SDCKey -pathtype container){
    Write-PSLog -message ($callingScript +" : SDC Key Found") -component "Main()" -type 1
    Write-Verbose "Check Passed. SDC Key Exists."
    $SDC_Ready = $true
}else{
    Write-PSLog -message ($callingScript +" : No SDC Keys detected.") -component "Main()" -type 3
    Write-Verbose "Check Failed.  No SDC Keys detected."
    $SDC_Ready = $false
}
#>

$SDC_Ready = $true



#Device/Cred Guard Check
WriteLog "**** Checking Virtual Secure Mode Readiness" 1
if ($dg -contains 1){
    WriteLog "Check Passed. VSM is ready for deployment." 1
} else {
    WriteLog "Check Failed. VSM is not ready for deployment - Please check requirements for Device/Credential Guard at $supportURL." 3
}

WriteLog "Checking for existance of McAfee HBSS components" 1
$vsepath = TestFilePath "C:\Program Files (x86)\McAfee\VirusScan Enterprise"
$hipspath = TestFilePath "C:\Program Files\McAfee\Host Intrusion Prevention"
$dlppath = TestFilePath "C:\Program Files\McAfee\DLP\Agent"
$mapath = TestFilePath "C:\Program Files (x86)\McAfee\Common Framework"

WriteLog "****Checking for supported versions of McAfee HBSS products" 1
WriteLog "****Checking McAfee VSE Version" 1
# Validate McAfee VSE Version
If ($vsepath -eq $false) {
    WriteLog "McAfee VSE is not installed.  Ready for Upgrade." 1
} Else {
    #Get McAfee VSE Version
    $VSE = Get-FileVersion "C:\Program Files (x86)\McAfee\VirusScan Enterprise\shstat.exe"
    If ((Check-Version $VSE $min_vsever) -eq $false){
            WriteLog "McAfee VSE Version ($VSE) is NOT supported.  Please contact your INOSC for assistance." 3
            $McAfeeReady = $false
    } Else {
        WriteLog "McAfee VSE Version ($VSE) is supported. Ready for upgrade." 1
    }
}

# Validate McAfee HIPS Version
WriteLog "****Checing McAfee HIPS Version" 1
If ($hipspath -eq $false) {
    WriteLog "McAfee HIPS is not installed. Ready for upgrade." 1
} Else {
    $HIPS = Get-FileVersion "C:\Program Files\McAfee\Host Intrusion Prevention\McAfeeFire.exe"
    If ((Check-Version $HIPS $min_hipver) -eq $false) {
        WriteLog "McAfee HIPS Version ($HIPS) is NOT supported.  Please contact your INOSC for assistance." 3
        $McAfeeReady = $false
    } Else {
        WriteLog "McAfee HIPS Version ($HIPS) is supported.  Ready for Upgrade." 1
    }
}

# Validate McAfee DLP Version
WriteLog "****Checking McAfee DLP Version" 1
If($DLPpath -eq $false) {
    WriteLog "McAfee DLP is not installed. Ready for upgrade." 1
} Else {
    $DLP = Get-FileVersion "C:\Program Files\McAfee\DLP\Agent\fcag.exe"
    If ((Check-Version $DLP $min_dlpver) -eq $false) {
        WriteLog "McAfee DLP Version ($DLP) is NOT supported.  Please contact your INOSC for assistance." 3
        $McAfeeReady = $false
    } Else {
        WriteLog "McAfee DLP Version ($DLP) is supported.  Ready for Upgrade." 1
    }
}

# Validate McAfee Agent version
WriteLog "****Checking McAfee Agent Version" 1
If($mapath -eq $false) {
    WriteLog "McAfee Agent is not installed. Ready for upgrade." 1
} Else {
    $MA = Get-FileVersion "C:\Program Files (x86)\McAfee\Common Framework\masvc.exe"
    if($MA -eq $null)
    {
      $MA = Get-FileVersion "C:\Program Files\McAfee\Agent\masvc.exe"
    } 
    If ((Check-Version $MA $min_mavver) -eq $false) {
        WriteLog "McAfee Agent Version ($MA) is NOT supported.  Please contact your INOSC for assistance." 3
        $McAfeeReady = $false
    } Else {
        WriteLog "McAfee Agent Version ($MA) is supported.  Ready for Upgrade." 1 
    }
}

#Determine if device is ready for upgrade and exit script if needed

If (($SDC_Ready -eq $false) -or ($SHB_Ready -eq $false)-or (-not($dg -contains 1)) -or ($cloneTag -eq $false) -or ($osReady -eq $false) -or ($McAfeeReady -eq $false) -or ($driverpath -eq $null) -or `
    ($ACpower -eq $false) -or ($InsuffSpace -eq $false)){
    WriteLog "Preflight checks have failed. Exiting Script. See $myLogName for details." 3
    WriteLog "----------------------------------------------------" 1
    Pop-Location
    Exit 1109
} Else {
    WriteLog "Preflight checks have been successfully Validated." 1
    WriteLog "--------------------------------------------------" 1
}

if($scanOnly){
    WriteLog "ScanOnly Option was selected.  The pre-stage file process will not be executed." 1
} else {
    WriteLog "Begin process to pre-stage OS installation, driver, and script files." 1
    #Validate Source and Driver file locations exist and are not empty
    WriteLog "Validate Required Source file directory structure" 1

    ValidateSource $scriptDir\Upgrade_OS
    ValidateSource $scriptDir\Upgrade_OS\PostOOBE
    ValidateSource $scriptDir\Upgrade_OS\Setup
    ValidateSource $scriptDir\Upgrade_OS\SDC-Splash
    ValidateSource $scriptDir\Drivers
    ValidateSource $scriptDir\Drivers\Chipset
    ValidateSource $scriptDir\Drivers\Wireless
    ValidateSource $scriptDir\Drivers\SmartCard
    ValidateSource $driverpath

    WriteLog "Required Source file directories exist and are not empty." 1

    #Stage OS Installation Source files.  The files will be copied from the Source directory to the Target Directory under Source folder
    #Create Installation source File folder
    if (-not(Test-Path $TargetPath)){
        New-Item $TargetPath -type Directory -ErrorAction SilentlyContinue -ErrorVariable myErr | Out-Null
        EvaluatePSError $myErr
    }

    if (Test-Path $TargetPath){
       WriteLog "$TargetPath folder location has been found. Begining file staging process." 1
       $currStep = "OS Upgrade and Script files"
       File-Copy "$scriptDir\Upgrade_OS $TargetPath /mir /LOG+:$myLogPath\$myCopyLogName"

	   if ($doUpgrade)
       {
            #Check Machine Plugged in
            if(-not(CheckACPower))
            {
                WriteLog "Please plug machine into AC Power" 3
                Exit 1100
            }

            #Track Target Destination
            WriteLog "Writing SourceFile.log" 1
	        echo $TargetPath > "C:\SourceFile.log"
       
    		#Splash Screen
    		$user = Get-WMIObject -class Win32_ComputerSystem | select username
    		$time = ((Get-Date).AddSeconds(60))
    		$jointime = (-join ($time).hour + ":" + ($time.Minute) + ":" + ($time).Second)
    		$task = "powershell.exe"
    		$argz = "-executionpolicy bypass -windowstyle hidden -file C:\Upgrade_OS\SDC-Splash\SplashScreen.PS1"
    		$tasktime = New-ScheduledTaskTrigger -Once -At $jointime
    		$action = New-ScheduledTaskAction -Execute $task -Argument $argz
    		$principle = New-ScheduledTaskPrincipal -UserID $user.username -LogonType ServiceAccount -RunLevel Highest
    		Register-ScheduledTask -Action $action -Trigger $tasktime -TaskName "Splash Screen" -Description "Fire Splashscreen" -Principal $principle
            WriteLog "Created Scheduled Task: Splash Screen" 1
    	}
        
        $currStep = "Common Chipset Driver files"
        File-Copy "$Drivers\Chipset $TargetPath\Drivers\Chipset /mir /LOG+:$myLogPath\$myCopyLogName"

        $currStep = "Common Wireless Driver files"
        File-Copy "$Drivers\Wireless $TargetPath\Drivers\Wireless /mir /LOG+:$myLogPath\$myCopyLogName"

        $currStep = "Common SmartCard Driver files"
        File-Copy "$Drivers\SmartCard $TargetPath\Drivers\SmartCard /mir /LOG+:$myLogPath\$myCopyLogName"

        $currStep = "$sysmodel Driver files"
        File-Copy "`"$driverpath`" `"$TargetPath\Drivers\$sysmodel`" /mir /LOG+:$myLogPath\$myCopyLogName"
        
        if ($doUpgrade){
        	Unregister-ScheduledTask -TaskName "Splash Screen" -Confirm:$false
            WriteLog "Scheduled Task: Splash Screen was unregistered" 1
    	}
    } else {
        WriteLog "$TargetPath folder location does NOT exist. Exiting Script." 3
        Pop-Location
        Exit 1102
    }

    WriteLog "File Staging process has completed." 1
    WriteLog "--------------------------------------------------" 1
    if ($doUpgrade){
        WriteLog "The OS Upgrade option has been detected.  The OS Upgrade process will now begin." 1
        WriteLog "Begin OS Setup: $TargetPath\Setup\setup.exe /auto upgrade /quiet /copylogs $myLogPath /installDrivers '$TargetPath\Drivers' /PostOOBE '$TargetPath\PostOOBE\setupcomplete.cmd'" 1

        Push-Location "$($TargetPath)\Setup\"
    
        try {
            #Check Machine Plugged in
            if(-not(CheckACPower))
            {
                WriteLog "Please plug machine into AC Power" 3
                Exit 1100
            }
        
            .\setup.exe /auto upgrade /quiet /copylogs $($myLogPath) /installDrivers "$($TargetPath)\Drivers" /PostOOBE "$($TargetPath)\PostOOBE\setupcomplete.cmd"
        }
        catch{
            #Add Exception object log
            WriteLog "Failed to execute: $TargetPath\Setup\setup.exe /auto upgrade /quiet /copylogs $myLogPath /installDrivers '$TargetPath\Drivers\' /PostOOBE '$TargetPath\PostOOBE\setupcomplete.cmd'" 3
            WriteLog "Exception: $($_.Exception.Message)" 3
            Stop-Process -processname mshta
        }  
        Pop-Location

        WriteLog "The Windows Upgrade process has started. The process will take approximately 3 hours and the platform will automatically reboot several times." 1
    } else {
        WriteLog "The OS Upgrade option was not detected.  The OS Upgrade process will not be initiated." 3
	    Stop-Process -processname mshta
    }
}

Pop-Location