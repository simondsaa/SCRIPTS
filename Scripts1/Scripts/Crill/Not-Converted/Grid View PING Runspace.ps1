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
#$Computername = $AllComputerNames | Sort-Object
$Computername = Get-Content "C:\Users\1394844760A\Desktop\Scripting Test Bed\names.txt"
Measure-Command {
#Begin
#What each runspace will do
$ScriptBlock = {
    Param ($computer,$hash)
    $Ping = test-connection $computer -count 1 -ea 0
    $hash[$Computer]= $Ping
        }

#Setup the runspace
$Script:runspaces = New-Object System.Collections.ArrayList   
# Data table for all of the runspaces
$hash = [hashtable]::Synchronized(@{})
$sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
$runspacepool = [runspacefactory]::CreateRunspacePool(1, 100, $sessionstate, $Host)
$runspacepool.Open() 

#Process
ForEach ($Computer in $Computername) {
    #Create the powershell instance and supply the scriptblock with the other parameters 
    $powershell = [powershell]::Create().AddScript($scriptBlock).AddArgument($computer).AddArgument($hash)

    #Add the runspace into the powershell instance
    $powershell.RunspacePool = $runspacepool

    #Create a temporary collection for each runspace
    $temp = "" | Select-Object PowerShell,Runspace,Computer
    $Temp.Computer = $Computer
    $temp.PowerShell = $powershell

    #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
    $temp.Runspace = $powershell.BeginInvoke()
    Write-Verbose ("Adding {0} collection" -f $temp.Computer)
    $runspaces.Add($temp) | Out-Null               
}

# Wait for all runspaces to finish
#End
Get-RunspaceData -Wait 
$stoptimer = Get-Date 
#Display info, and display in GridView
Write-Host
Write-Host "Availability check complete!" -ForegroundColor Cyan
}

$hash | ogv