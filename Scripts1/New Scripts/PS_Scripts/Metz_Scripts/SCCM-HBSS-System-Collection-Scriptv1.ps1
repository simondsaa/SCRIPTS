"Collecting computers from ActiveDirectory"
$starttime = get-date
"Processing started at $starttime"
$ALLcomputers = get-adcomputer -searchbase "OU=Peterson AFB,OU=AFCONUSWEST,OU=Bases,DC=area52,DC=afnoapps,DC=usaf,DC=mil" -filter * 
"AD Computers list obtained"
$computers = $allcomputers.name
$adcompcount = $computers.count
"There are $adcompcount computers"

$maxentries = 250 
$count = 1
$startentry = 0

$ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Definition -ErrorAction Stop

#region Set up Output Folders
$Output_Path = $ScriptPath + "\XML_Output"
    IF (!(Test-Path ($Output_Path) -ErrorAction SilentlyContinue))
        {
		"Attempting to create XML_output"
		$output_path
            New-Item -Path $Output_Path -ItemType Directory -ErrorAction SilentlyContinue | Out-Null -ErrorAction SilentlyContinue
        }
$Output_Error = $ScriptPath + "\Error_Log"
    If (!(Test-Path ($Output_Error) -ErrorAction SilentlyContinue))
        {
            New-Item -Path $Output_Error -ItemType Directory -ErrorAction SilentlyContinue | Out-Null -ErrorAction SilentlyContinue
        }
$Output_Error = "$Output_Error" + "\" + "ERROR_" + "$Output_Time" + ".csv"


do
{
$pipearray = $computers[$startentry..($startentry + $maxentries)]


start-job  -argumentlist $pipearray,$output_path,$Output_error -scriptblock {Param($pipearray, $output_path,$Output_Error)
foreach($Target_Machine in $pipearray)
    {
	$target_machine	

Function OS_Architecture_Bits
    {
        Param(  [Parameter(Position=0,Mandatory=$false)] [string]$Check_TargetFQDN)       

        Try
            {                
                $OS_Architecture = Get-WmiObject Win32_OperatingSystem -ComputerName $Check_TargetFQDN -AsJob -ErrorAction Stop | Wait-Job -Timeout 30
                If ($OS_Architecture.State -like "*Failed*")
                    {
                        $OS_Architecture = Start-Job -ScriptBlock {Get-WmiObject Win32_OperatingSystem -ComputerName $args[0] -ErrorAction Stop} -ArgumentList $Check_TargetFQDN | Wait-Job -Timeout 30
                    }
                If ($OS_Architecture.State -like "Completed")
                    {
                        $Return_Object = Receive-Job $OS_Architecture -ErrorAction Stop
                        $Return_Object = ($Return_Object).OSArchitecture
                        $OS_Architecture | Remove-Job -Force -ErrorAction SilentlyContinue
                        Return $Return_Object
                    }
                Else
                    {
                        $OS_Architecture | Remove-Job -Force -ErrorAction SilentlyContinue
                        $Return_Object = "64-bit"                       
                        Return $Return_Object
                    }                                   
            }
        Catch
            {
                $OS_Architecture = "64-bit"
                Return $OS_Architecture
            }    
    }
$OS_Architecture = OS_Architecture_Bits $Target_Machine
#Registry Paths
$SDCVersion_Key = "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\OEMInformation\Model"
$Pending_Reboot_Array = @(`
"HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending",`
"HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired",`
"HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\PendingFileRenameOperations")
$SCCMManuallyAssignedSiteCode_Key = "HKLM\SOFTWARE\Microsoft\SMS\Mobile Client\AssignedSiteCode"
$SCCMGPAssignedSiteCode_Key = "HKLM\SOFTWARE\Microsoft\SMS\Mobile Client\GPRequestedSiteAssignmentCode"
$GPOWSUSServer_Key = "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer"
If ($OS_Architecture -like "*64-bit*")
    {
        $AVDatDate_Key =  "HKLM\SOFTWARE\Wow6432Node\McAfee\AVEngine\AVDatDate"
        $EPOServerList_Key = "HKLM\SOFTWARE\Wow6432Node\Network Associates\ePolicy Orchestrator\Agent\ePOServerList"
        $EPOAgentGUID_Key = "HKLM\SOFTWARE\Wow6432Node\Network Associates\ePolicy Orchestrator\Agent\AgentGUID"
        $EPOWakeupPort_Key = "HKLM\SOFTWARE\Wow6432Node\Network Associates\ePolicy Orchestrator\Agent\AgentWakeUpPort"
        $EPOPropVersionDate_Key = "HKLM\SOFTWARE\Wow6432Node\Network Associates\ePolicy Orchestrator\Agent\PropsVersion"
        $MAVVersion_Key = "HKLM\SOFTWARE\Wow6432Node\Network Associates\ePolicy Orchestrator\Application Plugins\EPOAGENT3000\Version"
        $VSEVersion_Key = "HKLM\SOFTWARE\Wow6432Node\McAfee\DesktopProtection\szProductVer"
        $DLPVersion_Key = "HKLM\SOFTWARE\Wow6432Node\McAfee\DLP\Agent\AgentVersion"
        $LastASCTime_Key = "HKLM\SOFTWARE\Wow6432Node\Network Associates\ePolicy Orchestrator\Agent\LastASCTime"
        $FWEnabled_Key = "HKLM\SOFTWARE\Wow6432Node\McAfee\HIP\Config\Settings\FW_Enabled"
        $MHIPSVersion_Key = "HKLM\SOFTWARE\Wow6432Node\McAfee\HIP\VERSION"
        #File Paths (WMI)
        $SMSUpdate_Location = "Path='\\Windows\\SysWOW64\\CCM\\Cache\\'"
    }
ElseIf ($OS_Architecture -like "*32-bit")
    {
        $AVDatDate_Key = "HKLM\SOFTWARE\McAfee\AVEngine\AVDatDate"
        $EPOServerList_Key = "HKLM\SOFTWARE\Network Associates\ePolicy Orchestrator\Agent\ePOServerList"
        $EPOAgentGUID_Key = "HKLM\SOFTWARE\Network Associates\ePolicy Orchestrator\Agent\AgentGUID"
        $EPOWakeupPort_Key = "HKLM\SOFTWARE\Network Associates\ePolicy Orchestrator\Agent\AgentWakeUpPort"
        $EPOPropVersionDate_Key = "HKLM\SOFTWARE\Network Associates\ePolicy Orchestrator\Agent\PropsVersion"
        $MAVVersion_Key = "HKLM\SOFTWARE\Network Associates\ePolicy Orchestrator\Application Plugins\EPOAGENT3000\Version"
        $VSEVersion_Key = "HKLM\SOFTWARE\McAfee\DesktopProtection\szProductVer"
        $DLPVersion_Key = "HKLM\SOFTWARE\McAfee\DLP\Agent\AgentVersion"
        $LastASCTime_Key = "HKLM\SOFTWARE\Network Associates\ePolicy Orchestrator\Agent\LastASCTime"
        $FWEnabled_Key = "HKLM\SOFTWARE\McAfee\HIP\Config\Settings\FW_Enabled"
        $MHIPSVersion_Key = "HKLM\SOFTWARE\McAfee\HIP\VERSION"
        #File Paths (WMI)
        $SMSUpdate_Location = "Path='\\Windows\\System32\\CCM\\Cache\\'"
    }
#Service Names
$WindowsUpdate_ServiceName = "wuauserv"
$SMSAgent_ServiceName = "ccmexec"
$BITS_ServiceName = "bits"
$WinMgmt_ServiceName = "WinMgmt"
$RemoteRegistry_ServiceName = "RemoteRegistry"
$GPClient_ServiceName = "gpsvc"
$SMSTSMGR_ServiceName = "smstsmgr"
$McAfeeFramework_ServiceName = "McAfeeFramework"
$McShield_ServiceName = "McShield"
#File Paths (WMI)
$WindowsUpdateAgent_Path = "Name='C:\\Windows\\System32\\wuaueng.dll'"
#File Paths (Powershell)
$McAfeeAgentINI_Location_Array = @(`
"\\$Target_Machine\C$\ProgramData\McAfee\Common Framework\UpdateHistory.ini",`
"\\$Target_Machine\C$\ProgramData\McAfee\Agent\UpdateHistory.ini")

$McAfeeSiteListXML_Location_Array = @(`
"\\$Target_Machine\C$\ProgramData\McAfee\Common Framework\SiteList.xml",`
"\\$Target_Machine\C$\ProgramData\McAfee\Agent\Sitelist.xml",`
"\\$Target_Machine\C$\ProgramData\McAfee\Common Framework\ServerSiteList.xml",`
"\\$Target_Machine\C$\ProgramData\McAfee\Agent\ServerSiteList.xml")

$EPOWakeupPork_Location_Array = @(`
"\\$Target_Machine\C$\ProgramData\McAfee\Agent\manifest.xml")
#endregion Set Values to Query

#region Registry Key Processing Functions
#Function - Extract the Hive Name from a registry key input
Function Extract_RegHiveName
{
    Param(   [Parameter(Position=0,Mandatory=$false)] [string]$KeyToProcess)

    $RegExtract_HiveName = $KeyToProcess
    $RegExtract_HiveName = $RegExtract_HiveName.Remove($RegExtract_HiveName.IndexOf("\"))    
    If ($RegExtract_HiveName -eq "HKLM")
        {
            $RegExtract_HiveName = "LocalMachine"
        }
    If ($RegExtract_HiveName -eq "HKU")
        {
            $RegExtract_HiveName = "Users"
        }
    Return $RegExtract_HiveName
}

#Function - Extract the Key Path from a registry key input
Function Extract_RegKeyPath
{
    Param(   [Parameter(Position=0,Mandatory=$false)] [string]$KeyToProcess)

    $RegExtract_Path = $KeyToProcess.Remove($KeyToProcess.LastIndexOf("\"))
    $RegExtract_Path_HiveToRemove = Extract_RegHiveName $KeyToProcess
    $RegExtract_Path_HiveToRemove = $KeyToProcess.Remove($KeyToProcess.IndexOf("\"))
    $RegExtract_Path = $RegExtract_Path -replace "$RegExtract_Path_HiveToRemove",""
    $RegExtract_Path = $RegExtract_Path.Substring(1)
    return $RegExtract_Path
}

#Function - Extract the Key Name from a registry key input
Function Extract_KeyName
{
    Param(   [Parameter(Position=0,Mandatory=$false)] [string]$KeyToProcess)

    $RegExtract_KeyName = $KeyToProcess.Split("\")[-1]
    return $RegExtract_KeyName
}
#endregion Registry Key Processing Functions

#region Registry Value Collection
Function RegistryCollection_Core
{
Param(  [Parameter(Position=0,Mandatory=$false)] [string]$RegCheck_TargetFQDN,
        [Parameter(Position=1,Mandatory=$false)] [string]$RegCheck_RegHiveName,
        [Parameter(Position=2,Mandatory=$false)] [string]$RegCheck_RegPath,
        [Parameter(Position=3,Mandatory=$false)] [string]$RegCheck_KeyName)

        $RegCol_HiveType = [Microsoft.Win32.RegistryHive]::$RegCheck_RegHiveName
        $RegCol_RegCheckProcess = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($RegCol_HiveType, $RegCheck_TargetFQDN)
        $RegCol_RegCheckProcess = $RegCol_RegCheckProcess.OpenSubKey($RegCheck_RegPath)

        #Prepare for errors by setting a Do Not Proceed variable and setting the returns value which should be changed later
        $RegCol_DNP01 = $false
        $Regcol_Value = "Key_PathNotFound"
        #Checking if we are about to try to process an empty string
        If ($RegCheck_KeyName -eq "")
            {
                $RegCol_Value = "Key_PathNotFound"
                $RegCol_DNP01 = $true
            }
        If ($RegCol_DNP01 -ne $true)
            {
                Try
                    {
                        $Regcol_Value = "Key_KeyNotFound"
                        $Regcol_Value = $RegCol_RegCheckProcess.GetValue($RegCheck_KeyName)
                    }
                Catch
                    {
                        $Regcol_Value = "Key_KeyNotFound"
                    }
            }
        If (($Regcol_Value | Out-String).Trim() -like "")
            {
                $Regcol_Value = "Key_KeyNotInPath"
            }
return (($Regcol_Value | Out-String).Trim())
}

#Wrapper Function
Function RegistryCollection
{
Param(  [Parameter(Position=0,Mandatory=$false)] [string]$RegCheck_TargetFQDN,
        [Parameter(Position=1,Mandatory=$false)] [string]$RegCheck_RegistryKeyToCheck)
        
        $RegCheck_Hive = Extract_RegHiveName $RegCheck_RegistryKeyToCheck
        $RegCheck_Path = Extract_RegKeyPath $RegCheck_RegistryKeyToCheck
        $RegChecK_Key = Extract_KeyName $RegCheck_RegistryKeyToCheck
        
        $RegCheck_Return = RegistryCollection_Core $RegCheck_TargetFQDN $RegCheck_Hive $RegCheck_Path $RegChecK_Key -ErrorAction SilentlyContinue
        If ($RegCheck_Return -like "Key_*")
            {
                $RegCheck_Return = "Unable_To_Collect"
            }
        Return $RegCheck_Return
}

Function RegistryCollection_AllowBlank
{
    Param(  [Parameter(Position=0,Mandatory=$false)] [string]$RegCheck_TargetFQDN,
            [Parameter(Position=1,Mandatory=$false)] [string]$Registry_Key)

    $Return_Value = RegistryCollection $RegCheck_TargetFQDN $Registry_Key -ErrorAction SilentlyContinue
    If ($Return_Value -like "Unable_To_Collect")
        {
            $Return_Value = " "
        }
    Return $Return_Value
}
#endregion Registry Value Collection
#endregion Registry Key Processing Functions

#region FQDN, HostName, IPAddress, and DiskFreeSpace Functions

Function GetFQDN
{
    Param(  [Parameter(Position=0,Mandatory=$false)] [string]$Check_TargetFQDN)

        $FQDN_ReturnedFQDN = [System.Net.Dns]::GetHostByName("$Check_TargetFQDN").HostName.Trim()
        Return $FQDN_ReturnedFQDN
}

Function GetHostName
{
    Param(  [Parameter(Position=0,Mandatory=$false)] [string]$Check_TargetFQDN)

    Try
        {        
            $HostName_ReturnedHostName = Get-WmiObject Win32_ComputerSystem -computer $Check_TargetFQDN -ErrorAction Stop -AsJob | Wait-Job -Timeout 30
            If ($HostName_ReturnedHostName.State -like "*Failed*")
                {
                    $HostName_ReturnedHostName = Start-Job -ScriptBlock {Get-WmiObject Win32_ComputerSystem -computer $args[0] -ErrorAction Stop} -ArgumentList $Check_TargetFQDN | Wait-Job -Timeout 30
                }
            If ($HostName_ReturnedHostName.State -like "*Completed*")
                {
                    $Return_Object = Receive-Job $HostName_ReturnedHostName
                    $Return_Object = $Return_Object | Select-Object -ExpandProperty Name
                    $HostName_ReturnedHostName | Remove-Job -Force -ErrorAction SilentlyContinue
                    Return $Return_Object
                }
            Else
                {
                    $HostName_ReturnedHostName | Remove-Job -Force -ErrorAction SilentlyContinue
                    $Return_Object = "WMI_Failure"                    
                    Return $Return_Object
                }
        }
    Catch 
        {
            $HostName_ReturnedHostName | Remove-Job -Force -ErrorAction SilentlyContinue
            $HostName_ReturnedHostName = "WMI_Failure"            
            Return $HostName_ReturnedHostName
        }    
}

Function GetIPAddress
{
    Param(  [Parameter(Position=0,Mandatory=$false)] [string]$Check_TargetFQDN)
    
    Try
    {
        $Check_ReturnedIPAddress = Test-Connection $Check_TargetFQDN -count 1 -ErrorAction Stop | Select-Object -ExpandProperty IPv4Address
        $Check_ReturnedIPAddress = $Check_ReturnedIPAddress.ToString()
    }
    Catch
    {
        $Check_ReturnedIPAddress = "Connection_Failure"
    }
    Return $Check_ReturnedIPAddress
}

Function GetWindowsVersion
    {    
        Param(  [Parameter(Position=0,Mandatory=$false)] [string]$Check_TargetFQDN)

        Try
            {                
                $WindowsVersion_ReturnedHostName = Get-WmiObject Win32_OperatingSystem -computer $Check_TargetFQDN -ErrorAction Stop -AsJob | Wait-Job -Timeout 30
                If ($WindowsVersion_ReturnedHostName.State -like "*Failed*")
                    {
                        $WindowsVersion_ReturnedHostName = Start-Job -ScriptBlock {Get-WmiObject Win32_OperatingSystem -computer $args[0] -ErrorAction Stop} -ArgumentList $Check_TargetFQDN | Wait-Job -Timeout 30
                    }
                If ($WindowsVersion_ReturnedHostName.State -like "*Completed*")
                    {
                        $Return_Object = Receive-Job $WindowsVersion_ReturnedHostName
                        $Return_Object = $Return_Object.Version
                        $WindowsVersion_ReturnedHostName | Remove-Job -Force -ErrorAction SilentlyContinue
                        Return $Return_Object
                    }
                Else
                    {
                        $WindowsVersion_ReturnedHostName | Remove-Job -Force -ErrorAction SilentlyContinue
                        $Return_Object = "WMI_Failure"                        
                        Return $Return_Object
                    }                    
            }
        Catch 
            {
                $WindowsVersion_ReturnedHostName | Remove-Job -Force -ErrorAction SilentlyContinue
                $WindowsVersion_ReturnedHostName = "WMI_Failure"                
                Return $WindowsVersion_ReturnedHostName
            }        
    }

Function GetOSName
{
    Param(  [Parameter(Position=0,Mandatory=$false)] [string]$Check_TargetFQDN)

    Try
        {
            $Check_OSName = Get-WmiObject Win32_OperatingSystem -ComputerName $Check_TargetFQDN -ErrorAction Stop -AsJob | Wait-Job -Timeout 30
            If ($Check_OSName.State -like "*Failed*")
                {
                    $Check_OSName = Start-Job -ScriptBlock {Get-WmiObject Win32_OperatingSystem -ComputerName $args[0] -ErrorAction Stop} -ArgumentList $Check_TargetFQDN | Wait-Job -Timeout 30
                }
            If ($Check_OSName.State -like "*Completed*")
                {
                    $Return_Object = Receive-Job $Check_OSName
                    $Check_OSName = $Return_Object.Caption
                    $Check_OSName | Remove-Job -Force -ErrorAction SilentlyContinue
                    Return $Check_OSName
                }
            Else
                {
                    $Check_OSName | Remove-Job -Force -ErrorAction SilentlyContinue
                    $Check_OSName = " "                    
                    Return $Check_OSName
                }                         
        }
    Catch
        {
            $Check_OSName | Remove-Job -Force -ErrorAction SilentlyContinue
            $Check_OSName = " "            
            Return $Check_OSName
        }
}

Function GetDiskFreeSpace
{
    Param(  [Parameter(Position=0,Mandatory=$false)] [string]$Check_TargetFQDN)

    Try
        {            
            $Check_C_Drive = Get-WmiObject Win32_LogicalDisk -ComputerName $Check_TargetFQDN -Filter "DeviceID='C:'" -ErrorAction Stop -AsJob | Wait-Job -Timeout 30
            If ($Check_C_Drive.State -like "*Failed*")
                {
                    $Check_C_Drive = Start-Job -ScriptBlock {Get-WmiObject Win32_LogicalDisk -ComputerName $args[0] -Filter "DeviceID='C:'" -ErrorAction Stop} -ArgumentList $Check_TargetFQDN | Wait-Job -Timeout 30
                }
            If ($Check_C_Drive.State -like "*Completed*")
                {
                    $Return_Object = Receive-Job $Check_C_Drive
                    If ($Return_Object.Size -gt 0)
                        {
                            $Check_Disk_PercentFree = [Math]::Round((($Return_Object.FreeSpace/$Return_Object.Size) * 100),2)
                        }                
                    Else
                        {
                            $Check_Disk_PercentFree = 0
                        }
                    $Check_C_Drive | Remove-Job -Force -ErrorAction SilentlyContinue
                    Return $Check_Disk_PercentFree
                }                          
        }
    Catch
        {
            $Check_C_Drive | Remove-Job -Force -ErrorAction SilentlyContinue
            $Check_Disk_PercentFree = "Unable_To_Collect"            
            Return $Check_Disk_PercentFree
        }     
}
#endregion FQDN, HostName, IPAddress, and DiskFreeSpace Functions

#region McAfee and ePo related registry functions

Function EPO_Registry_Processing
{
Param(  [Parameter(Position=0,Mandatory=$false)] [string]$RegCheck_TargetFQDN,
        [Parameter(Position=1,Mandatory=$false)] [string]$RegCheck_RegistryKeyToCheck)

        $EPO_RawReturn = RegistryCollection $RegCheck_TargetFQDN $RegCheck_RegistryKeyToCheck
        IF ($EPO_RawReturn -like "*Unable_To_Collect*")
            {
                $Return_Object = New-Object -TypeName PSObject -Property @{
                    ServerName = "Unable_To_Collect"
                    ServerIP = "Unable_To_Collect"
                    ServerPort = "Unable_To_Collect"
                    }
                Return $Return_Object                
            }
        $EPO_RawReturn = $EPO_RawReturn.Split(";")
        $Return_Object = @()
        ForEach ($Server_Config in $EPO_RawReturn)
            {
                $Server_Config = $Server_Config.Split("|")                
                $Temporary_Object = New-Object -TypeName PSObject -Property @{
                    ServerName = $Server_Config[0]
                    ServerIP = $Server_Config[1]
                    ServerPort = $Server_Config[2]
                    }
                $Return_Object += $Temporary_Object
            }
        Return $Return_Object
}

Function EPO_WakeupPort_Processing
    {        
        Param(  [Parameter(Position=0,Mandatory=$false)] [string]$Check_TargetFQDN,
                [Parameter(Position=1,Mandatory=$false)] [string]$Check_RegistryKeyToCheck,
                [Parameter(Position=2,Mandatory=$false)] $XMLToParse_Array)                
        Try
            {
                $Return_Object = RegistryCollection $Check_TargetFQDN $Check_RegistryKeyToCheck
                If ($Return_Object -like "*Unable_To_Collect*")
                    {
                        ForEach ($XML_File_Location in $XMLToParse_Array)
                            {
                                If (!(Test-Path $XML_File_Location))
                                    {
                                        Continue
                                    }
                                [xml]$XMLToParse = Get-Content $XML_File_Location -ErrorAction Stop
                                If ($XMLToParse.Manifest.AgentMeta.AgentPingPort -notlike "")
                                    {
                                        $Return_Object = ($XMLToParse.Manifest.AgentMeta.AgentPingPort | Out-String).Trim()
                                        Return $Return_Object
                                    }
                            }
                    }
                Else
                    {
                        Return $Return_Object
                    }
            }
        Catch
            {
                $Return_Object = "Unable_To_Collect"
                Return $Return_Object
            }
    }

Function GetPropsVersionDate
    {
        Param(  [Parameter(Position=0,Mandatory=$false)] [string]$Check_TargetFQDN,
                [Parameter(Position=1,Mandatory=$false)] [string]$EPOPropVersionDate_KeyPath)

        $EPOPropsVersion_RawReturn = RegistryCollection $Check_TargetFQDN $EPOPropVersionDate_KeyPath
        If ($EPOPropsVersion_RawReturn -eq "Unable_to_Collect")
            {
                $EPOPropsVersionDate_Return = $EPOPropsVersion_RawReturn
            }
        Else
            {
                Try
                    {
                        $EPOPropsVersionDate_Return = [datetime]::ParseExact("$EPOPropsVersion_RawReturn","yyyyMMddHHmmss",$null)
                        $EPOPropsVersionDate_Return = ((Get-Date $EPOPropsVersionDate_Return -Format "yyyy-MM-dd").ToString()).Trim()
                    }
                Catch
                    {
                        $EPOPropsVersionDate_Return = "Error_Parsing_Data"
                    }
            }
        Return $EPOPropsVersionDate_Return      
    }

Function GetLastASCTime
    {
        Param(  [Parameter(Position=0,Mandatory=$false)] [string]$Check_TargetFQDN,
                [Parameter(Position=1,Mandatory=$false)] [string]$LastASCTime_KeyPath)
        
        $LastASCTime_RawReturn = RegistryCollection $Check_TargetFQDN $LastASCTime_KeyPath

        If ($LastASCTime_RawReturn -eq "Unable_To_Collect")
            {
                $LastASCTime_Return = $LastASCTime_RawReturn
            }
        Else
            {
                $LastASCTime_Return = [TimeZone]::CurrentTimeZone.ToLocalTime(([datetime]'1/1/1970').AddSeconds($LastASCTime_RawReturn))
                $LastASCTime_Return = ((Get-Date $LastASCTime_Return -Format "yyyy-MM-dd").ToString()).Trim()
            }
        Return $LastASCTime_Return        
    }

Function GetFWEnabled
    {
        Param(  [Parameter(Position=0,Mandatory=$false)] [string]$Check_TargetFQDN,
                [Parameter(Position=1,Mandatory=$false)] [string]$FWEnabled_KeyPath)

        $FWEnabled_RawReturn = RegistryCollection $Check_TargetFQDN $FWEnabled_KeyPath
        If ($FWEnabled_RawReturn -eq "1")
            {
                $FWEnabled_Return = "TRUE"
            }
        ElseIf ($FWEnabled_RawReturn -eq "Unable_To_Collect")
            {
                $FWEnabled_Return = $FWEnabled_RawReturn
            }
        Else
            {
                $FWEnabled_Return = "FALSE"
            }
        Return $FWEnabled_Return
    }        

Function GetMcAfeeCatalogVersionDate
    {
        Param(  [Parameter(Position=0,Mandatory=$false)] [string]$Check_TargetFQDN,
                [Parameter(Position=1,Mandatory=$false)] $FileToParse_Location_Array)
        Try
            {
                ForEach ($FileToParse in $FileToParse_Location_Array)
                    {
                        Try
                            {  
                                If (!(Test-Path $FileToParse))
                                    {
                                        Continue
                                    }                              
                                $Return_Value = Get-Content $FileToParse -ErrorAction Stop
                                $Return_Value = $Return_Value | Select-String -Pattern "CatalogVersion="
                                $Return_Value = $Return_Value[0]
                                $Return_Value = $Return_Value.ToString().Split("=")[1]
                                $Return_Value = [datetime]::ParseExact("$Return_Value","yyyyMMddHHmmss",$null)
                                $Return_Value = ((Get-Date $Return_Value -Format "yyyy-MM-dd").ToString()).Trim()
                            }
                        Catch
                            {
                                Continue
                            }
                    }
            }
        Catch
            {
                $Return_Value = "Date_Not_Found"
            }
        Return $Return_Value
    }

#region McAfee Sitelist.xml Function
Function GetMcAfeeSiteXMLInfo
    {
        Param(  [Parameter(Position=0,Mandatory=$false)] [string]$Check_TargetFQDN,
                [Parameter(Position=1,Mandatory=$false)] $XMLToParse_Array)                
        
        $McAfeeSiteNames = @()
        $McAfeeSiteIPs = @()
        $McAfeeSitePorts = @()
        $Return_Object = New-Object -TypeName PSObject -Property @{
            SiteNames = "Unable_To_Collect"
            SiteIPs = "Unable_To_Collect"
            SitePorts = "Unable_To_Collect"
            }
        Try
            {
                ForEach ($XMLToParse_Location in $XMLToParse_Array)
                    {
                        Try
                            {  
                                If (!(Test-Path $XMLToParse_Location -ErrorAction SilentlyContinue))
                                    {
                                        Continue
                                    }                                                                                            
                                [xml]$XMLToParse = Get-Content $XMLToParse_Location -ErrorAction Stop
                                If (($XMLToParse.SiteLists.SiteList.SPipeSite.Server -notlike "") -and ($XMLToParse.SiteLists.SiteList.SPipeSite.ServerIP -notlike "") -and ($XMLToParse.SiteLists.SiteList.SPipeSite.SecurePort -notlike ""))
                                    {
                                        $McAfeeSiteNames = ($XMLToParse.SiteLists.SiteList.SPipeSite.Server).Replace(":80","")
                                        $McAfeeSiteIPs = ($XMLToParse.SiteLists.SiteList.SPipeSite.ServerIP).Replace(":80","")
                                        $McAfeeSitePorts = ($XMLToParse.SiteLists.SiteList.SPipeSite.SecurePort).Replace(":80","")
                                        $Return_Object = New-Object -TypeName PSObject -Property @{
                                            SiteNames = $McAfeeSiteNames
                                            SiteIPs = $McAfeeSiteIPs
                                            SitePorts = $McAfeeSitePorts
                                            }
                                        Return $Return_Object
                                        Break
                                    }            
                            }
                        Catch
                            {
                                Continue
                            }
                    }
            }
        Catch
            {
                Return $Return_Object    
            }          
        Return $Return_Object
    }
#endregion McAfeeSitelist.xml Function
#endregion McAfee and ePo related registry functions
#region Boot and Reboot Status functions

Function GetPendingReboot
{
    Param(  [Parameter(Position=0,Mandatory=$false)] [string]$Check_TargetFQDN,
            [Parameter(Position=1,Mandatory=$false)] $Pending_Reboot_Array)
    
    $RebootPending_Return = "False"
    ForEach ($Registry_Path in $Pending_Reboot_Array)
        {           
            $Pending_Reboot_Result = RegistryCollection $Check_TargetFQDN $Registry_Path
            IF ($Pending_Reboot_Result -ne "Unable_To_Collect")
                {
                    $RebootPending_Return = "True"
                }
        }
    Return $RebootPending_Return
}

Function GetDaysSinceLastBoot
{
    Param(  [Parameter(Position=0,Mandatory=$false)] [string]$Check_TargetFQDN)

    Try
        {
            
            $Check_UpTime = Get-WmiObject Win32_OperatingSystem -ComputerName $Check_TargetFQDN -ErrorAction Stop -AsJob | Wait-Job -Timeout 30
            If ($Check_UpTime.State -like "*Failed*")
                {
                    $Check_UpTime = Start-Job -ScriptBlock {Get-WmiObject Win32_OperatingSystem -ComputerName $args[0] -ErrorAction Stop} -ArgumentList $Check_TargetFQDN | Wait-Job -Timeout 30
                }
            If ($Check_UpTime.State -like "*Completed*")
                {
                    $Return_Object = Receive-Job $Check_UpTime
                    $Check_UpTime | Remove-Job -Force -ErrorAction SilentlyContinue
                    $Return_Object = $Return_Object.LastBootUpTime
                    $Return_Object = [System.Management.ManagementDateTimeConverter]::ToDateTime($Return_Object)
                    $Check_CurrentTime = (Get-Date)
                    $Check_Uptime_Return = New-TimeSpan -Start $Return_Object -End $Check_CurrentTime
                    $Check_Uptime_Return = $Check_Uptime_Return.Days
                    Return $Check_Uptime_Return
                }
            Else
                {
                    $Check_UpTime | Remove-Job -Force -ErrorAction SilentlyContinue
                    $Return_Object = "WMI_Failure"                    
                    Return $Return_Object
                }           
        }
    Catch
        {
            $Check_UpTime | Remove-Job -Force -ErrorAction SilentlyContinue
            $Return_Object = "Unable_To_Collect"            
            Return $Return_Object
        }
}

Function GetMostRecentSMSUpdate
{
    Param(  [Parameter(Position=0,Mandatory=$false)] [string]$Check_TargetFQDN,
            [Parameter(Position=1,Mandatory=$false)] $SMSCache_Path)           

    Try
        {            
            $SMS_Query = Get-WmiObject Win32_Directory -Filter $SMSCache_path -ComputerName $Check_TargetFQDN -ErrorAction Stop -AsJob | Wait-Job -Timeout 30
            If ($SMS_Query.State -like "*Failed*")
                {
                    $SMS_Query = Start-Job -ScriptBlock {Get-WmiObject Win32_Directory -Filter $args[0] -ComputerName $args[1] -ErrorAction Stop} -ArgumentList $SMSCache_Path,$Check_TargetFQDN | Wait-Job -Timeout 30
                }
            If ($SMS_Query.State -like "*Completed*")
                {  
                    $Return_Object = Receive-Job $SMS_Query
                    $Return_Object = $Return_Object.LastModified | Sort | Select -Last 1
                    $SMS_Query | Remove-Job -Force -ErrorAction SilentlyContinue
                    Try
                        {
                            $Return_Object = [System.Management.ManagementDateTimeConverter]::ToDateTime($Return_Object)
                            $Return_Object = $Return_Object.ToString("yyyy-MM-dd")
                        }
                    Catch
                        {
                            $Return_Object = "Unable_To_Collect"        
                        }                    
                    Return $Return_Object
                }
            Else
                {
                    $SMS_Query | Remove-Job -Force -ErrorAction SilentlyContinue
                    $Return_Object = "WMI_Failure"                    
                    Return $Return_Object
                }
        }
    Catch
        {
            $SMS_Query | Remove-Job -Force -ErrorAction SilentlyContinue
            $Return_Object = "WMI_Failure"            
            Return $Return_Object
        }
}
#endregion Boot and Reboot Status functions
#region Site Code Discovery Functions
Function GetWMISiteCode
{
    Param(  [Parameter(Position=0,Mandatory=$false)] [string]$Check_TargetFQDN)
    
        Try
            {                
                $SMS_WMISiteCode = Invoke-WmiMethod -Namespace Root\CCM -Class SMS_Client -Name GetAssignedSite -ComputerName $Check_TargetFQDN -ErrorAction Stop -AsJob | Wait-Job -Timeout 30
                If ($SMS_WMISiteCode.State -like "*Failed*")
                    {
                        $SMS_WMISiteCode = Start-Job -ScriptBlock {Invoke-WmiMethod -Namespace Root\CCM -Class SMS_Client -Name GetAssignedSite -ComputerName $args[0] -ErrorAction Stop} -ArgumentList $Check_TargetFQDN | Wait-Job -Timeout 30
                    }
                IF ($SMS_WMISiteCode.State -like "*Completed*")
                    {
                        $Return_Object = Receive-Job $SMS_WMISiteCode
                        $Return_Object = $Return_Object.sSiteCode
                        $SMS_WMISiteCode | Remove-Job -Force -ErrorAction SilentlyContinue
                        Return $Return_Object
                    }              
            }
        Catch
            {
                $SMS_WMISiteCode | Remove-Job -Force -ErrorAction SilentlyContinue
                $SMS_WMISiteCode = "Unable_To_Collect"                
                Return $SMS_WMISiteCode
            }
}
#endregion Site Code Discovery Functions
#region SMS Certificate Functions
Function GetSMSCerts_RO
{
    Param(   [Parameter(Position=0,Mandatory=$false)] [string]$Check_TargetFQDN)
    
    $RegAccess_RO = [System.Security.Cryptography.X509Certificates.OpenFlags]"ReadOnly"
    $RegAccess_LocalMachine = [System.Security.Cryptography.X509Certificates.StoreLocation]"LocalMachine"
    $RegAccess_Store = New-Object System.Security.Cryptography.X509Certificates.X509Store("\\$Check_TargetFQDN\sms",$RegAccess_LocalMachine)
    $RegAccess_Store.Open($RegAccess_RO)
    $RegAccess_Store.Certificates
}

Function GetSMSEncryptionCert
{
    Param(   [Parameter(Position=0,Mandatory=$false)] [string]$Check_TargetFQDN)

    $GetSMSCertReturn = GetSMSCerts_RO $Check_TargetFQDN.ToString()
    $GetSMSCertReturn = $GetSMSCertReturn | Where-Object {$_.FriendlyName -like "SMS Encryption Certificate"}
    $GetSMSCertReturn = $GetSMSCertReturn.Subject
    Return $GetSMSCertReturn
}

Function GetSMSSigningCert
{
Param(   [Parameter(Position=0,Mandatory=$false)] [string]$Check_TargetFQDN)

$GetSMSCertReturn = GetSMSCerts_RO $Check_TargetFQDN.ToString()
$GetSMSCertReturn = $GetSMSCertReturn | Where-Object {$_.FriendlyName -like "SMS Signing Certificate"}
$GetSMSCertReturn = $GetSMSCertReturn.Subject
Return $GetSMSCertReturn
}

Function EncryptionCertMatchCheck
{
    Param(  [Parameter(Position=0,Mandatory=$false)] [string]$Check_TargetFQDN)
    
    Try
        {
            $CertOutput_HostName = GetSMSEncryptionCert $Check_TargetFQDN
            $CertOutput_HostName = $CertOutput_HostName.Split("=")[-1]    
            $HostName = GetHostName $Check_TargetFQDN
            If ($CertOutput_HostName -like $HostName)
                {
                    $CertNameMatchCheck = "Good"
                }
            Else
                {
                    $CertNameMatchCheck = "Bad"
                }
        }
    Catch
        {
            $CertNameMatchCheck = "Unknown"
        }
    Return $CertNameMatchCheck
}

Function SigningCertMatchCheck
{
    Param(  [Parameter(Position=0,Mandatory=$false)] [string]$Check_TargetFQDN)
    
    Try
        {
            $CertOutput_HostName = GetSMSSigningCert $Check_TargetFQDN
            $CertOutput_HostName = $CertOutput_HostName.Split("=")[-1]    
            $HostName = GetHostName $Check_TargetFQDN
            If ($CertOutput_HostName -like $HostName)
                {
                    $CertNameMatchCheck = "Good"
                }
            Else
                {
                    $CertNameMatchCheck = "Bad"
                }
        }
    Catch
        {
            $CertNameMatchCheck = "Unknown"
        }
    Return $CertNameMatchCheck
}

#endregion SMS Certificate Functions
#region Service Status Information Gathering Functions
Function Service_Status
    {
        Param(  [Parameter(Position=0,Mandatory=$false)] [string]$Check_TargetFQDN,
                [Parameter(Position=1,Mandatory=$false)] [string]$Check_ServiceName)
        Try
            {                           
                $Service_Status = Get-WmiObject Win32_Service -ComputerName $Check_TargetFQDN -Filter "Name='$Check_ServiceName'" -ErrorAction Stop -AsJob | Wait-Job -Timeout 30
                If ($Service_Status.State -like "*Failed*")
                    {
                        $Service_Status = Start-Job -ScriptBlock {Get-WmiObject Win32_Service -ComputerName $args[0] -Filter "Name='$args[1]'" -ErrorAction Stop} -ArgumentList $Check_TargetFQDN,$Check_ServiceName | Wait-Job -Timeout 30
                    }
                If ($Service_Status.State -like "Completed")
                    {
                        Try
                            {
                                $Return_Object = Receive-Job $Service_Status
                                $Service_Status_Object = New-Object PSObject -Property @{
                                    State = $Return_Object.State
                                    StartMode = $Return_Object.StartMode
                                    }
                                $Service_Status | Remove-Job -Force -ErrorAction SilentlyContinue
                                Return $Service_Status_Object
                            }
                        Catch
                            {
                                $Service_Status | Remove-Job -Force -ErrorAction SilentlyContinue
                                $Service_Status_Object = New-Object PSObject -Property @{
                                    State = "Unable_To_Collect"
                                    StartMode = "Unable_To_Collect"
                                    }                                
                                Return $Service_Status_Object
                            }
                    }
                Else
                    {
                        $Service_Status | Remove-Job -Force -ErrorAction SilentlyContinue
                        $Service_Status_Object = New-Object -TypeName PSObject -Property @{
                            State = "Unable_To_Collect"
                            StartMode = "Unable_To_Collect"                    
                            }
                            Return $Service_Status_Object                            
                    }
                }
        Catch
            {
                $Service_Status | Remove-Job -Force -ErrorAction SilentlyContinue
                $Service_Status_Object = New-Object PSObject -Property @{
                    State = "Unable_To_Collect"
                    StartMode = "Unable_To_Collect"
                    }                
                Return $Service_Status_Object
            }  
    }
#endregion Service Status Information Gathering Functions
#region SCCM Information
Function SCCM_Version
{
    Param(  [Parameter(Position=0,Mandatory=$false)] [string]$Check_TargetFQDN)
    Try
    {        
        $SMS_VersionNumber = Get-WmiObject -Namespace Root\CCM -Class CCM_InstalledComponent -ComputerName $Check_TargetFQDN -ErrorAction Stop -AsJob | Wait-Job -Timeout 30
        If ($SMS_VersionNumber.State -like "*Failed*")
            {
                $SMS_VersionNumber = Start-Job -ScriptBlock {Get-WmiObject -Namespace Root\CCM -Class CCM_InstalledComponent -ComputerName $args[0] -ErrorAction Stop} -ArgumentList $Check_TargetFQDN | Wait-Job -Timeout 30
            }
        IF ($SMS_VersionNumber.State -like "*Completed*")
            {
                $Return_Object = Receive-Job $SMS_VersionNumber
                $Return_Object = ($Return_Object | Where-Object {$_.name -like "CcmPolicyAgent"}).version
                $SMS_VersionNumber | Remove-Job -Force -ErrorAction SilentlyContinue
                Return $Return_Object
            }
        Else
            {
                $SMS_VersionNumber | Remove-Job -Force -ErrorAction SilentlyContinue
                $Return_Object = "Unable_To_Collect"                
                Return $Return_Object
            }
    }
    Catch
    {
        $SMS_VersionNumber | Remove-Job -Force -ErrorAction SilentlyContinue
        $Return_Object = "Unable_To_Collect"        
        Return $Return_Object
    }
}

Function SCCM_ManagementPoint
{
    Param(  [Parameter(Position=0,Mandatory=$false)] [string]$Check_TargetFQDN)         
            
    Try
    {        
        $SMSMP_Return = Get-WmiObject -Namespace Root\CCM -Class SMS_Authority -ComputerName $Check_TargetFQDN -ErrorAction Stop -AsJob | Wait-Job -Timeout 30
        If ($SMSMP_Return.State -like "*Failed*")
            {
                $SMSMP_Return = Start-Job -ScriptBlock {Get-WmiObject -Namespace Root\CCM -Class SMS_Authority -ComputerName $args[0] -ErrorAction Stop} -ArgumentList $Check_TargetFQDN | Wait-Job -Timeout 30
            }
        IF ($SMSMP_Return.State -like "*Completed*")
            {
                $Return_Object = Receive-Job $SMSMP_Return
                $Return_Object = $Return_Object.CurrentManagementPoint.Trim()
                $SMSMP_Return | Remove-Job -Force -ErrorAction SilentlyContinue
                Return $Return_Object
            }
        Else
            {
                $SMSMP_Return | Remove-Job -Force -ErrorAction SilentlyContinue
                $Return_Object = "Unable_To_Collect"                
                Return $Return_Object
            }
    }
    Catch
    {
        $SMSMP_Return | Remove-Job -Force -ErrorAction SilentlyContinue
        $Return_Object = "Unable_To_Collect"        
        Return $Return_Object
    }
}
#endregion SCCM_Information
#region WSUS Server Discovery Functions
Function WUA_AgentVersion
{
    Param(  [Parameter(Position=0,Mandatory=$false)] [string]$Check_TargetFQDN,
            [Parameter(Position=1,Mandatory=$false)] [string]$WUA_Path)

    Try 
    {
        $WUA_FileVersion_Return = Get-WmiObject -Class CIM_Datafile -Filter $WUA_Path -ComputerName $Check_TargetFQDN -ErrorAction Stop -AsJob | Wait-Job -Timeout 30
        If ($WUA_FileVersion_Return.State -like "*Failed*")
            {
                $WUA_FileVersion_Return = Start-Job -ScriptBlock {Get-WmiObject -Class CIM_Datafile -Filter $args[0] -ComputerName $args[1] -ErrorAction Stop} -ArgumentList $WUA_Path,$Check_TargetFQDN | Wait-Job -Timeout 30
            }
        If ($WUA_FileVersion_Return.State -like "*Completed*")
            {
                $Return_Object = Receive-Job $WUA_FileVersion_Return
                $Return_Object = ($Return_Object.Version | Out-String).Trim()
                $WUA_FileVersion_Return | Remove-Job -Force -ErrorAction SilentlyContinue
                Return $Return_Object
            }
        Else
            {
                $WUA_FileVersion_Return | Remove-Job -Force -ErrorAction SilentlyContinue
                $Return_Object = "Unable_To_Collect"                
                Return $Return_Object
            }
    }
    Catch
    {
        $WUA_FileVersion_Return | Remove-Job -Force -ErrorAction SilentlyContinue
        $Return_Object = "Unable_To_Collect"        
        Return $Return_Object
    }
}
#endregion WSUS Server Discovery Functions
#region Timestamp
Function TimeStamp
{
    $TimeStamp = (Get-Date -Format "MM/dd/yyyy HH:mm").ToString()
    Return $TimeStamp
}


	Function Connectivity_Test
    {
        Param(      [Parameter(Position=0,Mandatory=$false)] [string]$Target_Machine)
        Try
        {
            If (Test-Connection -ComputerName $Target_Machine -Count 5 -Quiet -ErrorAction Stop)
                {   
                    Try
                        {                            
                            $Not_Used = Get-WmiObject Win32_BIOS -ComputerName $Target_Machine -ErrorAction Stop -AsJob | Wait-Job -Timeout 30
                            If ($Not_Used.State -like "*Completed*")
                                {
                                    Try
                                        {
                                            $Reg_QuickCheck = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $Target_Machine)
                                            $Reg_QuickCheck_Key = $Reg_QuickCheck.OpenSubKey("SYSTEM\CurrentControlSet\Control\ComputerName\ComputerName")
                                            $Output = $Reg_QuickCheck_Key.GetValue("ComputerName")
                                            $Output = "GOOD"
                                        }
                                    Catch [System.UnauthorizedAccessException]
                                        {
                                            $Output = "Registry_Access_Denied"
                                        }
                                    Catch
                                        {
                                            $Output = "Registry_Failure"
                                        }
                                }
                            If ($Not_Used.State -like "*Failed*")
                                {
                                    $Not_Used = Start-Job -ScriptBlock {Get-WmiObject Win32_BIOS -ComputerName $Target_Machine -ErrorAction Stop} | Wait-Job -Timeout 30
                                    IF ($Not_Used.State -like "*Failed*")
                                        {
                                            $Not_Used = Start-Job -ScriptBlock {Get-WmiObject Win32_BIOS -ComputerName $args[0] -ErrorAction Stop} -ArgumentList $Target_Machine | Wait-Job -Timeout 30
                                        }
                                    If ($Not_Used.State -like "*Completed*")
                                        {
                                            Try
                                                {
                                                    $Reg_QuickCheck = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $Target_Machine)
                                                    $Reg_QuickCheck_Key = $Reg_QuickCheck.OpenSubKey("SYSTEM\CurrentControlSet\Control\ComputerName\ComputerName")
                                                    $Output = $Reg_QuickCheck_Key.GetValue("ComputerName")
                                                    $Output = "GOOD"
                                                }
                                            Catch [System.UnauthorizedAccessException]
                                                {
                                                    $Output = "Registry_Access_Denied"
                                                }
                                            Catch
                                                {
                                                    $Output = "Registry_Failure"
                                                }
                                        }
                                    If ($Not_Used.State -like "*Failed*")
                                        {
                                            $Output = "WMI_Failure_002"
                                        }
                                }
                        }
                    Catch [System.UnauthorizedAccessException]
                        {
                            $Output = "WMI_Access_Denied"
                        }
                    Catch [System.Runtime.InteropServices.ExternalException]
                        {
                            $Output = "COM_Access_Failure"
                        }                
                    Catch
                        {
                            $Output = "WMI_Access_Failure"
                        }               
                }
            Else
                {
                    $Output = "Ping_Failure"
                }
            }
            Catch
            {
                $Output = "Ping_Failure"
            }            
        Return $Output
    }


	Function Collection
    {
        Param(  [Parameter(Position=0,Mandatory=$false)] [string]$Target_Machine,
                [Parameter(Position=1,Mandatory=$false)] [string]$Output_Path,
                [Parameter(Position=2,Mandatory=$false)] [string]$Output_Error)    
        
        $OutputTable = @()
        $Output_DNP = Connectivity_Test $Target_Machine
        If ($Output_DNP -like "GOOD")            
            {
                $Output_ComputerFQDN = GetFQDN $Target_Machine
                $Output_ComputerName = GetHostName $Target_Machine
                $Output_IPAddress = GetIPAddress $Target_Machine
                $Output_OSName = GetOSName $Target_Machine
                $Output_SDCVersion = RegistryCollection $Target_Machine $SDCVersion_Key
                $Output_PercentFreespace = GetDiskFreeSpace $Target_Machine
                $Output_PendingReboot = GetPendingReboot $Target_Machine $Pending_Reboot_Array
                $Output_DaysSinceLastBoot = GetDaysSinceLastBoot $Target_Machine
                $Output_MostRecentSMSUPdate = GetMostRecentSMSUpdate $Target_Machine $SMSUpdate_Location
                $Output_ManuallyAssignedSiteCode = RegistryCollection_AllowBlank $Target_Machine $SCCMManuallyAssignedSiteCode_Key
                $Output_GetGPAssignedSiteCode = RegistryCollection $Target_Machine $SCCMGPAssignedSiteCode_Key
                $Output_SMSActualSiteCode = GetWMISiteCode $Target_Machine
                $Output_EncryptionCertSubject = GetSMSEncryptionCert $Target_Machine
                $Output_EncryptionCertMatch = EncryptionCertMatchCheck $Target_Machine
                $Output_SigningCertSubject = GetSMSSigningCert $Target_Machine
                $Output_SigningCertMatch = SigningCertMatchCheck $Target_Machine
                $Intermediate_WindowsUpdateService = Service_Status $Target_Machine $WindowsUpdate_ServiceName
                $Output_WindowsUpdateServiceState = $Intermediate_WindowsUpdateService.State
                $Output_WindowsUpdateServiceStartupMode = $Intermediate_WindowsUpdateService.StartMode
                $Intermediate_SMSAgentServiceState = Service_Status $Target_Machine $SMSAgent_ServiceName
                $Output_SMSAgentServiceState = $Intermediate_SMSAgentServiceState.State
                $Output_SMSAgentStartupMode = $Intermediate_SMSAgentServiceState.StartMode
                $Intermediate_BITSServiceState = Service_Status $Target_Machine $BITS_ServiceName
                $Output_BITSServiceState = $Intermediate_BITSServiceState.State
                $Output_BITSServiceStartupMode = $Intermediate_BITSServiceState.StartMode
                $Intermediate_WinMgmtService = Service_Status $Target_Machine $WinMgmt_ServiceName
                $Output_WinMgmtServiceState = $Intermediate_WinMgmtService.State
                $Output_WinMgmtServiceStartupMode = $Intermediate_WinMgmtService.StartMode
                $Intermediate_RemoteRegistryService = Service_Status $Target_Machine $RemoteRegistry_ServiceName
                $Output_RemoteRegistryServiceState = $Intermediate_RemoteRegistryService.State
                $Output_RemoteRegistryStartupMode = $Intermediate_RemoteRegistryService.StartMode
                $Intermediate_GPClientService = Service_Status $Target_Machine $GPClient_ServiceName
                $Output_GPClientServiceState = $Intermediate_GPClientService.State
                $Output_GPClientServiceStartupMode = $Intermediate_GPClientService.StartMode
                $Intermediate_SMSTSMGRService = Service_Status $Target_Machine $SMSTSMGR_ServiceName
                $Output_SMSTSMGRServiceState = $Intermediate_SMSTSMGRService.State
                $Output_SMSTSMGRServiceStartupMode = $Intermediate_SMSTSMGRService.StartMode
                $Output_SCCMVersion = SCCM_Version $Target_Machine
                $Output_SCCM_ManagementPoint = SCCM_ManagementPoint $Target_Machine
                $Output_WindowsUpdateServer = RegistryCollection $Target_Machine $GPOWSUSServer_Key
                $Output_WUAAgentVersion = WUA_AgentVersion $Target_Machine $WindowsUpdateAgent_Path
                $Output_Timestamp = Timestamp

                #HBSS Related Variables
                $Output_McAfeeFrameworkCatalogVersionDate = GetMcAfeeCatalogVersionDate $Target_Machine $McAfeeAgentINI_Location_Array
                $Intermediate_McAfeeSiteXML = GetMcAfeeSiteXMLInfo $Target_Machine $McAfeeSiteListXML_Location_Array
                $Output_SiteListNames = $Intermediate_McAfeeSiteXML.SiteNames
                $Output_SiteListIPs = $Intermediate_McAfeeSiteXML.SiteIPs
                $Output_SitelistPorts = $Intermediate_McAfeeSiteXML.SitePorts
                
                $Intermediate_McAfeeFrameworkService = Service_Status $Target_Machine $McAfeeFramework_ServiceName    
                $Output_McAfeeFrameworkState = $Intermediate_McAfeeFrameworkService.State
                $Output_McAfeeFrameworkStartupMode = $Intermediate_McAfeeFrameworkService.StartMode
                $Intermediate_McAfeeShieldService = Service_Status $Target_Machine $McShield_ServiceName
                $Output_McAfeeShieldState = $Intermediate_McAfeeShieldService.State
                $Output_McAfeeShieldStartupMode = $Intermediate_McAfeeShieldService.StartMode
                $Output_OSVersion = GetWindowsVersion $Target_Machine                
                $Output_AVDatDate =  RegistryCollection $Target_Machine $AVDatDate_Key
                $Intermediate_EPOServer_Data = EPO_Registry_Processing $Target_Machine $EPOServerList_Key
                $Output_EPOServer_Name = $Intermediate_EPOServer_Data.ServerName
                $Output_EPOServer_IP = $Intermediate_EPOServer_Data.ServerIP
                $Output_AgentGUID = RegistryCollection $Target_Machine $EPOAgentGUID_Key
                $Output_LastASCTime = GetLastASCTime $Target_Machine $LastASCTime_Key
                $Output_PropsVersionDate = GetPropsVersionDate $Target_Machine $EPOPropVersionDate_Key
                $Output_MAVersion = RegistryCollection $Target_Machine $MAVVersion_Key
                $Output_VSEVersion = RegistryCollection $Target_Machine $VSEVersion_Key
                $Output_HIPSVersion = RegistryCollection $Target_Machine $MHIPSVersion_Key
                $Output_DLPVersion = RegistryCollection $Target_Machine $DLPVersion_Key
                $Output_AgentWakeupPort = EPO_WakeupPort_Processing $Target_Machine $EPOWakeupPort_Key $EPOWakeupPork_Location_Array
                $Output_FWEnabled = GetFWEnabled $Target_Machine $FWEnabled_Key
                #End HBSS Related Variables

                $OutputTable = New-Object -TypeName PSObject -Property @{
                ComputerFQDN = $Output_ComputerFQDN
                ComputerName = $Output_ComputerName
                IPAddress = $Output_IPAddress
                OSName = $Output_OSName
                SDCVersion = $Output_SDCVersion
                PercentFreespace = $Output_PercentFreespace
                PendingReboot = $Output_PendingReboot
                DaysSinceLastBoot = $Output_DaysSinceLastBoot
                MostRecentSMSUpdate = $Output_MostRecentSMSUPdate
                ManuallyAssignedSiteCode = $Output_ManuallyAssignedSiteCode
                GPOAssignedSiteCode = $Output_GetGPAssignedSiteCode
                SMSClientReportedSiteCode = $Output_SMSActualSiteCode
                EncryptionCertSubject = $Output_EncryptionCertSubject
                EncryptionCertSubjectMatch = $Output_EncryptionCertMatch
                SigningCertSubject = $Output_SigningCertSubject
                SigningCertSubjectMatch = $Output_SigningCertMatch
                WindowsUpdateServiceState = $Output_WindowsUpdateServiceState
                WindowsUpdateServiceStartupMode = $Output_WindowsUpdateServiceStartupMode
                SMSAgentHostServiceState = $Output_SMSAgentServiceState
                SMSAgentHostServiceStartupMode = $Output_SMSAgentStartupMode
                BITSServiceState = $Output_BITSServiceState
                BITSServiceStartupMode = $Output_BITSServiceStartupMode
                WinMgmtServiceState = $Output_WinMgmtServiceState
                WinMgmtServiceStartupMode = $Output_WinMgmtServiceStartupMode
                RemoteRegistryServiceState = $Output_RemoteRegistryServiceState
                RemoteRegistryServiceStartupMode = $Output_RemoteRegistryStartupMode
                GPClientServiceState = $Output_GPClientServiceState
                GPClientServiceStartupMode = $Output_GPClientServiceStartupMode
                SMSTSMGRServiceState = $Output_SMSTSMGRServiceState
                SMSTSMGRServiceStartupMode = $Output_SMSTSMGRServiceStartupMode
                SCCMVersion = $Output_SCCMVersion
                SCCMManagementPoint = $Output_SCCM_ManagementPoint
                WindowsUpdateServer = $Output_WindowsUpdateServer
                WindowsUpdateAgentVersion = $Output_WUAAgentVersion
                Timestamp = $Output_Timestamp
                SiteListNames = $Output_SiteListNames
                SiteListIPs = $Output_SiteListIPs
                SiteListPorts = $Output_SitelistPorts
                McAfeeFrameworkState = $Output_McAfeeFrameworkState
                McAfeeFrameworkStartupMode = $Output_McAfeeFrameworkStartupMode
                McAfeeShieldState = $Output_McAfeeShieldState
                McAfeeShieldStartupMode = $Output_McAfeeShieldStartupMode
                OSVersion = $Output_OSVersion
                AVDatDate = $Output_AVDatDate
                McAfeeFrameworkCatalogVersion = $Output_McAfeeFrameworkCatalogVersionDate
                EPOServerNames = $Output_EPOServer_Name
                EPOServerIPs = $Output_EPOServer_IP
                AgentGUID = $Output_AgentGUID
                LastASCTime = $Output_LastASCTime
                PropsVersionDate = $Output_PropsVersionDate
                MAVersion = $Output_MAVersion
                VSEVersion = $Output_VSEVersion
                HIPSVersion = $Output_HIPSVersion
                DLPVersion = $Output_DLPVersion
                AgentWakeupPort = $Output_AgentWakeupPort
                FWEnabled = $Output_FWEnabled
                }
                $Output_File = "$Output_Path\$Target_Machine" + ".xml"
                $OutputTable | Select-Object ComputerFQDN,ComputerName,IPAddress,OSName,SDCVersion,PercentFreespace,PendingReboot,DaysSinceLastBoot,MostRecentSMSUPdate,ManuallyAssignedSiteCode,GPOAssignedSiteCode,SMSClientReportedSiteCode,EncryptionCertSubject,EncryptionCertSubjectMatch,SigningCertSubject,SigningCertSubjectMatch,WindowsUpdateServiceState,WindowsUpdateServiceStartupMode,SMSAgentHostServiceState,SMSAgentHostServiceStartupMode,BITSServiceState,BITSServiceStartupMode,WinMgmtServiceState,WinMgmtServiceStartupMode,RemoteRegistryServiceState,RemoteRegistryServiceStartupMode,GPClientServiceState,GPClientServiceStartupMode,SMSTSMGRServiceState,SMSTSMGRServiceStartupMode,SCCMVersion,SCCMManagementPoint,WindowsUpdateServer,WindowsUpdateAgentVersion,Timestamp,SiteListNames,SiteListIPs,SiteListPorts,McAfeeFrameworkState,McAfeeFrameworkStartupMode,McAfeeShieldState,McAfeeShieldStartupMode,OSVersion,AVDatDate,EPOServerNames,EPOServerIPs,AgentGUID,LastASCTime,PropsVersionDate,MAVersion,VSEVersion,HIPSVersion,DLPVersion,AgentWakeupPort,FWEnabled,McAfeeFrameworkCatalogVersion | Export-Clixml $Output_File -Force
                Return
            }
        Else
            {        
                $Output_ComputerFQDN = $Target_Machine
                $Output_ErrorMessage = $Output_DNP
                $Output_Error_Troubleshooting = DNP_TroubleShooting $Output_ErrorMessage
                $OutputTable = New-Object -TypeName PSObject -Property @{
                    TargetMachine = $Output_ComputerFQDN
                    ErrorMessage = $Output_ErrorMessage
                    ErrorTroubleshooting = $Output_Error_Troubleshooting
                    }
                $OutputTable | Select-Object TargetMachine,ErrorMessage,ErrorTroubleshooting | Export-Csv $output_error -Force -NoTypeInformation -Append
                Return                       
            }
        Trap
            {               
                $Output_ComputerFQDN = $Target_Machine
                $OutputTable = New-Object -TypeName PSObject -Property @{
                    TargetMachine = $Output_ComputerFQDN
                    ErrorMessage = "Undocumented_Error"
                    ErrorTroubleshooting = " "
                    }
                $OutputTable | Select-Object TargetMachine,ErrorMessage,ErrorTroubleshooting | Export-Csv $Output_Error -Force -NoTypeInformation -Append 
                Return   
            }
    }



				
	Collection $Target_Machine $output_path $output_error | out-null


    }
}
$startentry = $startentry + $maxentries + 1
$count = $count + 1
}
While($startentry -le $computers.count)

Do
{
	If((get-job | where {$_.state -eq "Running"}).count -ge 1)
	{
		$jobcount = (get-job | where {$_.state -eq "Running"}).count
		"$jobcount Bulk Jobs are still running, waiting 30 Seconds"
		Sleep 30	
	}
}
Until ((get-job | where {$_.state -eq "Running"}).count -lt 1)
"Importing .XML output files for .CSV output"
$allxmlobjfiles = get-childitem $output_path -erroraction stop | ?{$_.Name -like "*.xml"}
$allxmlobjs = @()
push-location $output_path
ForEach ($Report in $allxmlobjfiles)
                    {
                        $ImportedObject = Import-Clixml $Report -ErrorAction SilentlyContinue
                        $allxmlobjs += $ImportedObject                     
                    }
"Exporting SCCM report to $output_path"
$allxmlobjs | Select-Object ComputerFQDN,ComputerName,IPAddress,OSName,SDCVersion,PercentFreespace,PendingReboot,DaysSinceLastBoot,MostRecentSMSUPdate,ManuallyAssignedSiteCode,GPOAssignedSiteCode,SMSClientReportedSiteCode,EncryptionCertSubject,EncryptionCertSubjectMatch,SigningCertSubject,SigningCertSubjectMatch,WindowsUpdateServiceState,WindowsUpdateServiceStartupMode,SMSAgentHostServiceState,SMSAgentHostServiceStartupMode,BITSServiceState,BITSServiceStartupMode,WinMgmtServiceState,WinMgmtServiceStartupMode,RemoteRegistryServiceState,RemoteRegistryServiceStartupMode,GPClientServiceState,GPClientServiceStartupMode,SMSTSMGRServiceState,SMSTSMGRServiceStartupMode,SCCMVersion,SCCMManagementPoint,WindowsUpdateServer,WindowsUpdateAgentVersion,Timestamp | Export-Csv "$output_path\SCCMResults.csv" -Force -NoTypeInformation
"Exporting HBSS report to $output_path"
$allxmlobjs | Select-Object ComputerFQDN,ComputerName,IPAddress,OSName,OSVersion,PendingReboot,DaysSinceLastBoot,RemoteRegistryServiceState,RemoteRegistryServiceStartupMode,McAfeeFrameworkState,McAfeeFrameworkStartupMode,McAfeeShieldState,McAfeeShieldStartupMode,AVDatDate,McAfeeFrameworkCatalogVersion,SiteListNames,SiteListIPs,SiteListPorts,EPOServerNames,EPOServerIPs,AgentGUID,LastASCTime,PropsVersionDate,MAVersion,VSEVersion,HIPSVersion,DLPVersion,AgentWakeupPort,FWEnabled,Timestamp | Export-Csv "$output_path\HBSSReport.csv" -Force -NoTypeInformation

$collectedcompcount = $allxmlobjfiles.count

$endtime = get-date
$hourcount = ($endtime - $starttime).hours
$minutecount = ($endtime - $starttime).minutes
$secondcount = ($endtime - $starttime).seconds
"Collection completed in $hourcount hours $minutecount minutes and $secondcount seconds"
"Data collected from $collectedcompcount of $adcompcount AD computer names"


write-host "Script Complete" -foregroundcolor green