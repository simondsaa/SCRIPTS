<#
.Synopsis
   SDCTargeting - Gathers report on machines put into Security Groups for SDC Deployment Solutions
   SDCTargeting - Gathers report on all machines that are ready to upgrade to 5.3.1
   **Needs to be run as administrator
   **Needs rsat tools including functions like Get-ADComputer
.DESCRIPTION
    This script generates a reports on the machines in each of the 2 Security groups:
        SDC_Servicing (Mandatory)
        SDC_Servicing (Automatic)
    This script should be ran on a daily basis to check the current status of machines within the security groups
    
    For the SDC_Servicing groups the report gathers 2 main results:
        Machines that are now on 5.3.1
        Machines that are still on 5.2
            Machines that failed pre-flight checks
            Machines that never recieved advertisment
            Machiens that were interrupted during upgrade phase
    
    Legacy Security Groups are currently unavailable, please wait for next updated script

    For any machine that failed to upgrade, important logs are gathered in host computer to centeralize and process logs to determine reason for failure

.UPDATES
    This script has been modifided to gather all machines ready to upgrade to 5.3.1
    This script has also been modified to create a csv file with all the data presented accordinly

.ARGUMENTS
    BaseName-----------------------------Input base name (ex. Maxwell, Fairchild, Charleston, Scott, etc...)
    Upgradable_SDC_Servicing_Machines----Input whether you want to query all 5.2 machines to see which are ready to upgrade(y/n)
    SDC_Servicing_Mandatory--------------Input whether you want to query this Security Group (y/n)
    SDC_Servicing_Available--------------Input whether you want to query this Security Group (y/n)
    PathToWriteLogs----------------------Input valid path to where important logs will be exported to (ex. C:\Users\Admin\Desktop\Logs)
.EXAMPLE
   .\SDCTargeting.ps1
   Will be prompted for arguments
#>
param(
[parameter(Mandatory=$true)][string]$BaseName,
[parameter(Mandatory=$true)][string]$AFCONUS,
[parameter(Mandatory=$true)][string]$Upgradable_SDC_Servicing_Machines,
[parameter(Mandatory=$true)][string]$Upgradable_Legacy_Machines,
[parameter(Mandatory=$true)][string]$SDC_Servicing_Mandatory,
[parameter(Mandatory=$true)][string]$SDC_Servicing_Available,
[parameter(Mandatory=$true)][ValidateScript({Test-Path $_ -PathType ‘Container’})][string]$PathToWriteLogs
)

Import-Module ActiveDirectory

#Validate input
$SDC_Servicing_Mandatory = $SDC_Servicing_Mandatory.ToLower()
$SDC_Servicing_Available = $SDC_Servicing_Available.ToLower()
$Upgradable_SDC_Servicing_Machines = $Upgradable_SDC_Servicing_Machines.ToLower()
$Upgradable_Legacy_Machines = $Upgradable_Legacy_Machines.ToLower()
$AFCONUS = $AFCONUS.ToUpper()

if(($SDC_Servicing_Mandatory -ne "n" -and $SDC_Servicing_Mandatory -ne "no" -and $SDC_Servicing_Mandatory -ne "y" -and $SDC_Servicing_Mandatory -ne "yes") -or`
   ($SDC_Servicing_Available -ne "n" -and $SDC_Servicing_Available -ne "no" -and $SDC_Servicing_Available -ne "y" -and $SDC_Servicing_Available -ne "yes") -or`
   ($Upgradable_SDC_Servicing_Machines -ne "n" -and $Upgradable_SDC_Servicing_Machines -ne "no" -and $Upgradable_SDC_Servicing_Machines -ne "y" -and $Upgradable_SDC_Servicing_Machines -ne "yes") -or`
   ($Upgradable_Legacy_Machines -ne "n" -and $Upgradable_Legacy_Machines -ne "no" -and $Upgradable_Legacy_Machines -ne "y" -and $Upgradable_Legacy_Machines -ne "yes"))
{
    Write-Host "Invalid input, please enter only (y/n or yes/no) for Upgradeable_Machines/Servicing Groups" -ForegroundColor Red
    exit
}

if($BaseName -like "*AFB*")
{
    Write-Host "Invalid input, please enter only the Basename without AFB" -ForegroundColor Red
    exit
}

if($AFCONUS -ne "EAST" -and $AFCONUS -ne "WEST")
{
    Write-Host "Invalid input, please enter only EAST or WEST based on your base AFCONUS" -ForegroundColor Red
    exit
}

if($SDC_Servicing_Mandatory -eq "y" -or $SDC_Servicing_Mandatory -eq "yes")
{
    $SDC_Servicing_Mandatory = $true
}
else
{
     $SDC_Servicing_Mandatory = $false
}

if($SDC_Servicing_Available -eq "y" -or $SDC_Servicing_Available -eq "yes")
{
    $SDC_Servicing_Available = $true
}
else
{
     $SDC_Servicing_Available = $false
}

if($Upgradable_SDC_Servicing_Machines -eq "y" -or $Upgradable_SDC_Servicing_Machines -eq "yes")
{
    $Upgradable_SDC_Servicing_Machines = $true
}
else
{
     $Upgradable_SDC_Servicing_Machines = $false
}

if($Upgradable_Legacy_Machines -eq "y" -or $Upgradable_Legacy_Machines -eq "yes")
{
    $Upgradable_Legacy_Machines = $true
}
else
{
    $Upgradable_Legacy_Machines = $false
}

if($PathToWriteLogs -eq $null)
{
    $PathToWriteLogs = "C:\Windows\Temp"
}

if ($PathToWriteLogs.endsWith("\") -eq $true)
{
    $PathToWriteLogs = $PathToWriteLogs-replace ".{1}$"
}

function CheckModel($type, $model){
    if($type -eq "Legacy")
    {
        return ($model -eq "HP EliteBook 840 G2" -or $model -eq "HP ProBook 640 G1" -or $model -eq "HP ZBook 15" -or $model -eq "HP Z230 Tower Workstation" -or $model -eq "HP Z840 Workstation" -or`
                $model -eq "HP EliteDesk 705 G1 MT" -or $model -eq "HP EliteDesk 705 G1 SFF" -or $model -eq "HP Z820 Workstation" -or $model -eq "HP EliteBook 840 G1" -or $model -eq "HP ZBook 17" -or $model -eq "HP ZBook 15 G2" -or`
                $model -eq "HP Z420 Workstation" -or $model -eq "HP ZBook 14")
    }
    elseif($type -eq "SDC_Servicing")
    {
        return ($model -eq "HP ProBook 640 G2" -or $model -eq "HP Z240 Tower Workstation" -or $model -eq "HP EliteDesk 705 G2 MINI" -or $model -eq "HP EliteDesk 705 G2 SFF" -or`
                $model -eq "HP EliteDesk 705 G3 DESKTOP MINI" -or $model -eq "HP Z840 Workstation" -or $model -eq "HP EliteDesk 800 G2 SFF" -or $model -eq "HP EliteDesk 800 G2 TWR" -or $model -eq "HP EliteBook 840 G2" -or`
                $model -eq "HP EliteBook 840 G3" -or ($model -like "*612 G1*") -or $model -eq "HP ZBook 15 G3" -or $model -eq "B300G5" -or ($model -like "*V110G2*"))
    }
    return $false
}

function GetModel($type,$name){
    
    if(-not (Test-Connection -ComputerName $name -Count 1 -Quiet))
    {
        $model = "Offline"
    }
    else
    {
        $model = (Get-WMIObject Win32_ComputerSystem -computername $name).model
        if($model -eq $null)
        {
            $model = "Unknown"
        }
        else
        {

            if(-not (CheckModel $type $model))
            {
                $model = "*" + $model
            }
        }
    }
    Write-Host $model
    return $model
}

function CheckACPower($name){
    if((Get-WMIObject -Class Win32_ComputerSystem -ComputerName $name).PCSystemType -eq 2)
        { 
            $status = (Get-WMIObject -Class BatteryStatus -namespace root\wmi -ComputerName $name).PowerOnline[0]
            if($status -eq $null)
            {
                return $false
            }
            else
            {
                return $status
            }
        }
    return $true
       
}

function CheckDiskSpace($name){

    return (Get-WMIObject -Class Win32_logicaldisk -filter "Drivetype=3" -ComputerName $name).FreeSpace[0]
}

function CheckProfileSize($name){
    return (Get-ChildItem \\$name\C$\Users -Recurse | Measure-Object -Property length -sum).sum/1GB
}

function CheckVersion($currVer, $minVer){
    #Build arrays from dot values"
    if($currVer -eq $null)
    {
        return $false
    }

    if($minVer -eq $null)
    {
        return $false 
    }
    return [System.Version]$currVer -ge [System.Version]$minVer 
}

function LogToCSV($status,$group,$connection,$name,$model,$hbss,$vseVersion,$hipVersion,$dlpVersion,$maVersion,$dg,$ct,$disk,$ACPower){
    $report = New-Object psobject

    $report | Add-Member -MemberType NoteProperty -name Base -Value $BaseName
    $report | Add-Member -MemberType NoteProperty -name Group -Value $group
    $report | Add-Member -MemberType NoteProperty -name Connection -Value $connection

    $report | Add-Member -MemberType NoteProperty -name Status -Value $status
    $report | Add-Member -MemberType NoteProperty -name ComputerName -Value $name
    $report | Add-Member -MemberType NoteProperty -name Model -Value $model

    $report | Add-Member -MemberType NoteProperty -name HBSS -Value $hbss
    $report | Add-Member -MemberType NoteProperty -name DeviceGaurdRunning -Value $dg
    $report | Add-Member -MemberType NoteProperty -name ValidCloneTag -Value $ct
    $report | Add-Member -MemberType NoteProperty -name Disk_GB -Value $disk
    $report | Add-Member -MemberType NoteProperty -name ACPower -Value $ACPower

    $report | Add-Member -MemberType NoteProperty -name VSEVersion -Value $vseVersion
    $report | Add-Member -MemberType NoteProperty -name HIPSVersion -Value $hipVersion
    $report | Add-Member -MemberType NoteProperty -name DLPVersion -Value $dlpVersion
    $report | Add-Member -MemberType NoteProperty -name McAfeeAgentVersion -Value $maVersion

    $report | export-csv $PathToWriteLogs\Win10Preflight_$BaseName.csv -NoTypeInformation -Append -NoClobber -Force
}

function LogToLegacy($status,$group, $connection,$name,$model,$OSversion,$profileSize,$disk,$ACPower){
    $report = New-Object psobject

    $report | Add-Member -MemberType NoteProperty -name Base -Value $BaseName
    $report | Add-Member -MemberType NoteProperty -name Group -Value $group
    $report | Add-Member -MemberType NoteProperty -name Connection -Value $connection
    $report | Add-Member -MemberType NoteProperty -name Status -Value $status
    $report | Add-Member -MemberType NoteProperty -name ComputerName -Value $name
    $report | Add-Member -MemberType NoteProperty -name Model -Value $model
    $report | Add-Member -MemberType NoteProperty -name OSVersion -Value $OSversion
    $report | Add-Member -MemberType NoteProperty -name ProfileSize -Value $profileSize
    $report | Add-Member -MemberType NoteProperty -name Disk_GB -Value $disk
    $report | Add-Member -MemberType NoteProperty -name ACPower -Value $ACPower
    $report | export-csv $PathToWriteLogs\Win7_8Preflight_$BaseName.csv -NoTypeInformation -Append -NoClobber -Force
}

function PreflightCheck($group,$strcomputer,$model)
{
    $vsepath = (ls -Path "\\$strcomputer\c$\Program Files (x86)\McAfee\VirusScan Enterprise\" -filter shstat.exe | Get-ItemProperty | Select-Object versioninfo -ExpandProperty versioninfo).fileversion
    $hipspath = (ls -Path "\\$strcomputer\c$\Program Files\McAfee\Host Intrusion Prevention\" -filter McAfeeFire.exe | Get-ItemProperty | Select-Object versioninfo -ExpandProperty versioninfo).fileversion
    $dlppath = (ls -Path "\\$strcomputer\c$\Program Files\McAfee\DLP\Agent\" -filter fcag.exe | Get-ItemProperty | Select-Object versioninfo -ExpandProperty versioninfo).fileversion
    $mapath = (ls -Path "\\$strcomputer\c$\Program Files (x86)\McAfee\Common Framework\" -filter masvc.exe | Get-ItemProperty | Select-Object versioninfo -ExpandProperty versioninfo).fileversion
    if($mapath -eq $null)
    {
        $mapath = (ls -Path "\\$strcomputer\c$\Program Files\McAfee\Agent\" -filter masvc.exe | Get-ItemProperty | Select-Object versioninfo -ExpandProperty versioninfo).fileversion
    } 

    #Check min version
    $min_vsever = "8.8.0.1599"
    $min_hipver = "8.0.0.3828"
    $min_dlpver = "10.0.100.37"
    $min_mavver = "5.0.4.283"

    #Check disk
    $disk = CheckDiskSpace($strcomputer)
    $disk/=1GB
    $diskCheck  = 15 -lt $disk

    #Check if its plugged in
    $ACPower = CheckACPower($strcomputer) 

    #Flags
    $dg = $false
    $ct = $false
    $hbss = $true
                
    if(-not(CheckVersion $vsepath $min_vsever))
    {
        $hbss = $false
    }
    if(-not(CheckVersion $hipspath $min_hipver))
    {
        $hbss = $false
    }
    if(-not(CheckVersion $dlppath $min_dlpver))
    {
        $hbss = $false
    }
    if(-not(CheckVersion $mapath $min_mavver))
    {
        $hbss = $false
    }
            
    #Device Gaurd
    $dg = ((Get-WmiObject -className Win32_deviceguard -Namespace root\Microsoft\Windows\DeviceGuard -ComputerName $strcomputer).SecurityServicesRunning) -contains 1

    $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine,$strcomputer)
    $key = $reg.OpenSubKey('SYSTEM\Setup')
    $CloneTag = $key.GetValue('CloneTag')

    Switch ($CloneTag){
        "Fri Jul 22 20:46:22 2016" {
            $ct = $true
        }            
        "Wed Mar 23 18:20:12 2016" {
            $ct = $true
        }
        default {
            $ct = $false
        }
    }
    if($Model -ne "Unknown")
    {
        if($hbss -and $dg -and $ct -and $dg -and $ACPower)
        {
            $status = "Ready to Upgrade"
        }
        else
        {
            $status = "Failed Preflight"
        }
    }
    else
    {
        $status = "Unknown"
    }
    LogToCSV $status $group "Online" $strcomputer $Model $hbss $vsepath $hipspath $dlppath $mapath $dg $ct $disk $ACPower
}

function CheckMachinesSDC_ServicingReady($computers,$group){
    $ErrorActionPreference = "SilentlyContinue"
    $totalcount = $computers.count
    $x=0
    $starttime = Get-Date

    foreach($strcomputer in $computers)
    {
        $elapsedTime = New-TimeSpan -Start $starttime -End (get-date)
        Write-Progress -Activity "Collecting Upgradable Machine Data ---elapsed time $elapsedtime" -Status $strcomputer -PercentComplete (($x/$totalcount)*100)
        $x+=1

        $strcomputer = $strcomputer.name
        $connection = "Online"
        $status = "Something"
        $SupportedFlag = $true


        $Model = GetModel "SDC_Servicing" $strcomputer

        if($Model -eq "Offline")
        {
            $connection = "Offline"
            $status = "Offline"
            $SupportedFlag = $false
        }elseif($Model[0] -eq "*")
        {
            $status = "Unsupported"
            $Model = $Model.substring(1,$Model.length-1)
            $SupportedFlag = $false
        }


        if($SupportedFlag -eq $true)
        {                   
           PreflightCheck $group $strcomputer $Model
        }
        else
        {
            LogToCSV $status $group $connection $strcomputer $Model $null $null $null $null $null $null $null $null $null
        }
    }

}

function CheckMachinesLegacyReady($computers,$group){
    $ErrorActionPreference = "SilentlyContinue"
    $totalcount = 1$computers.count
    $x=0
    $starttime = Get-Date

    foreach($strcomputer in $computers)
    {
        $elapsedTime = New-TimeSpan -Start $starttime -End (get-date)
        Write-Progress -Activity "Collecting Upgradable Machine Data ---elapsed time $elapsedtime" -Status $strcomputer -PercentComplete (($x/$totalcount)*100)
        $x+=1

        $name = $strcomputer.name
        $connection = "Online"
        $status = "Something"
        $SupportedFlag = $true

        $Model = GetModel "Legacy" $name

        if($Model -eq "Offline")
        {
            $connection = "Offline"
            $status = "Offline"
            $SupportedFlag = $false
        }elseif($Model[0] -eq "*")
        {
            $status = "Unsupported"
            $Model = $Model.substring(1,$Model.length-1)
            $SupportedFlag = $false
        }

        if($SupportedFlag -eq $true)
        {
            #Check disk
            $disk = CheckDiskSpace($name)
            $disk/=1GB
            $diskCheck  = 15 -lt $disk

            #Check if its plugged in
            $ACPower = CheckACPower($name) 

            #CheckProfileSize
            $ProfileSize = CheckProfileSize($name)

            if($Model -ne "Unknown")
            {

                if($diskCheck -and $ACPower)
                {
                    #Passed
                    ReadyLegacy.Add($name)
                    $status = "Ready to Upgrade"
                }
                else
                {
                    $status = "Failed Preflight"
                }
            }
            else
            {
                $status = "Unknown"
            }
            LogToLegacy $status $group $connection $name $Model $strcomputer.OperatingSystemVersion $ProfileSize $disk $ACPower        
        }
        else
        {
            LogToLegacy $status $group $connection $name $Model $strcomputer.OperatingSystemVersion $null $null $null
        } 
    }                  
}

#Gathers machine info from Security Group
function GetSDC_ServicingMachines($computers,$type){
    foreach ($compInfo in $computers){
        $name = $compInfo.Name
        $Model = GetModel "SDC_Servicing" $name
        
        $connection = "Online"
        $status = "Something"
        $SupportedFlag = $true

        if($Model -eq "Offline")
        {
            $connection = "Offline"
            $status = "Offline"
            $SupportedFlag = $false

        }elseif($Model[0] -eq "*")
        {
            $status = "Unsupported"
            $Model = $Model.substring(1,$Model.length-1)
            $SupportedFlag = $false
        }

        if($SupportedFlag -eq $true)
        {

            if($compInfo.OperatingSystemVersion -eq "10.0 (10586)")
            {
                #No advertisement ran
                if(-not (test-path \\$name\c$\upgrade_OS_Logs))
                {
                    LogToCSV "No Advertisement" "SDC_Servicing($type)" $connection $com $Model $null $null $null $null $null $null $null $null $null
                }
                #Pre-Flight Checks Failed
                elseif(-not (test-path \\$name\c$\upgrade_OS))
                {
                    PreflightCheck "SDC_Servicing($type)" $name $Model
                    Copy-Item \\$name\c$\upgrade_OS_Logs\* "$PathToWriteLogs\$name-PreFlightLog.log"
                }
                #Interruptted via Restart
                else
                {
                    LogToCSV "Advertisment Interrupted" "SDC_Servicing($type)" $connection $com $Model $null $null $null $null $null $null $null $null $null
                }
            }
            #Already at 5.3.1
            else
            {
                LogToCSV "Upgraded to 5.3.1" "SDC_Servicing($type)" $connection $com $Model $null $null $null $null $null $null $null $null $null
                #Remove from Group
            }
        }
        else
        {
            LogToCSV $status "SDC_Servicing($type)" $connection $com $Model $null $null $null $null $null $null $null $null $null
        }
    }
}


$date = Get-Date

if($Upgradable_SDC_Servicing_Machines -eq $true -or $SDC_Servicing_Mandatory -eq $true -or $SDC_Servicing_Available -eq $true)
{  
LogToCSV "Initializing new report for $BaseName AFB: $date" $null $null $null $null $null $null $null $null $null $null $null $null $null
LogToCSV '-' '-' '-' '-' '-' '-' '-' '-' '-' '-' '-' '-' '-' '-'   
#Get all computers in bases Active Directory
$compsInfo = Get-ADComputer -SearchBase "OU=$BaseName AFB,OU=AFCONUS$AFCONUS,OU=Bases,DC=AREA52,DC=AFNOAPPS,DC=USAF,DC=MIL" `
-Filter 'OperatingSystem -like "Windows 10*"' -Properties name, OperatingSystemVersion, memberof
}

if($Upgradable_Legacy_Machines -eq $true)
{
LogToLegacy "Initializing new report for $BaseName AFB: $date" $null $null $null $null $null $null $null $null    
LogToLegacy '-' '-' '-' '-' '-' '-' '-' '-' '-'
#Get all computers in bases Active Directory
$comps7Info = Get-ADComputer -SearchBase "OU=$BaseName AFB,OU=AFCONUS$AFCONUS,OU=Bases,DC=AREA52,DC=AFNOAPPS,DC=USAF,DC=MIL" `
-Filter 'OperatingSystem -like "Windows 7*"' -Properties name, OperatingSystemVersion, memberof
}


if($Upgradable_SDC_Servicing_Machines -eq $true)
{
    $AllWin10Computers = $compsInfo | where {$_.OperatingSystemVersion -eq "10.0 (10586)"}
    Write-Host "Gathering all 5.2 machines that are ready to be added to the SDC_Servicing Security Group"
}
if($Upgradable_Legacy_Machines -eq $true)
{
    $AllLegacyComputers = $comps7Info
    Write-Host "Gathering all 3.x/4.x machines that are ready to be added to the Legacy Security Group"
}

if($SDC_Servicing_Mandatory -eq $true)
{
    $SMcompsInfo = $compsInfo | where {$_.memberof -like "*SDC Servicing Script (Mandatory)*" }
    Write-Host Finding Status on 5.1/5.2 Machines in the SDC_Servicing Mandatory Security Group
}

if($SDC_Servicing_Available -eq $true)
{
    $SAcompsInfo = $compsInfo | where {$_.memberof -like "*SDC Servicing Script (Available)*"}
    Write-Host Finding Status on 5.1/5.2 Machines in the SDC_Servicing Available Security Group
}

Write-Host Processing.....

if($Upgradable_SDC_Servicing_Machines -eq $true)
{
    CheckMachinesSDC_ServicingReady $AllWin10Computers "No group"
}

if($Upgradable_Legacy_Machines -eq $true)
{
    CheckMachinesLegacyReady $AllLegacyComputers "No group"
}

if($SDC_Servicing_Mandatory -eq $true)
{
    GetSDC_ServicingMachines $SMcompsInfo "Mandatory"
}

if($SDC_Servicing_Available -eq $true)
{
    GetSDC_ServicingMachines $SAcompsInfo "Available"
}

Write-Host Completed