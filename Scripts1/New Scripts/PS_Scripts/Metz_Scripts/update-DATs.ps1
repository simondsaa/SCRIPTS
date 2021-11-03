
<#
Param($csvFile)

$datdata = import-csv $csvfile
#>

$datdata = gc "C:\Users\1274873341C\Desktop\Desktop\PS_Scripts\HBSS\dat_targets.txt"
$exe = "mcupdate.exe" 
$switches = "/update /quiet"
$unavailable = @()
$updatenotran = @()
$failedexe = @()
$updateran = @()


Foreach($entry in $datdata."System Name")
{
	$path = ""
	$Ping = new-object system.net.networkinformation.ping
	$reply = $ping.send($entry)
	
    If ($reply.status -eq "Success")
	{
		#Set executable Variables
		If((GWmi win32_operatingsystem -computername $entry).osarchitecture -eq "32-bit"){$Path = "C:\Program files\Mcafee\VirusScan Enterprise"}
		If((GWmi win32_operatingsystem -computername $entry).osarchitecture -eq "64-bit"){$Path = "C:\Program files (x86)\Mcafee\VirusScan Enterprise"}
		
    If(!$Path)
		{
			"ERROR in Setting `$Path Variable"
			$Updatenotran += $entry
		}
	
    	$cmd = "$path\$exe $switches"
		$wmi=([wmiclass]"\\$entry\root\cimv2:win32_process")
		$datprocess = $wmi.create("$cmd")
	
	If($datprocess.returnvalue -ne "0")
		{
			"Dat update failed to start on $entry"
			$failedexe += $entry
		}
		
    If($datprocess.returnvalue -eq "0")
		{
			"Update Process Started on $entry"
			$updateran += $entry
		}
    }
	Else
	{
		"$entry is unavailable"
		$unavailable += $entry
	}

}

$hash = @{}
$hash.unavailable = $unavailable
$hash.failedexe = $failedexe
$hash.updateran = $updateran
$hash.updatenotran = $updatenotran
$hash | export-clixml "Update-dats-results.xml"

			
		





