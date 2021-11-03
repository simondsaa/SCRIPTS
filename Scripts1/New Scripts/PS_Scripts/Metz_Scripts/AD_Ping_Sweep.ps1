"Collecting Tyndall AFB computers from ActiveDirectory"
$starttime = get-date
"Processing started at $starttime"
$ALLcomputers = get-adcomputer -searchbase "OU=Tyndall AFB Computers, OU=Tyndall AFB,OU=AFCONUSEAST,OU=Bases,DC=AREA52,DC=AFNOAPPS,DC=USAF,DC=MIL" -filter * 
"AD Computers list obtained"
$computers = $allcomputers.name
$adcompcount = $computers.count
"There are $adcompcount computers"
cd "C:\Users\1274873341C\Desktop\Desktop\PS_Scripts\Metz_Scripts\Ping_Sweep_Results\"

ri "C:\Users\1274873341C\Desktop\Desktop\PS_Scripts\Metz_Scripts\Ping_Sweep_Results\File*.txt"
ri "C:\Users\1274873341C\Desktop\Desktop\PS_Scripts\Metz_Scripts\Ping_Sweep_Results\master*.txt"


$maxentries = 150
$count = 1
$startentry = 0

DO
{
$pipearray = $computers[$startentry..($startentry + $maxentries)]
#$filepath = ("\\Server\users$\Username\Th1ngz\Remote-Work-parsing\File" + $count + ".txt")
$filepath = ("C:\Users\1274873341C\Desktop\Desktop\PS_Scripts\Metz_Scripts\Ping_Sweep_Results\File" + $count + ".txt")
start-job  -argumentlist $pipearray,$filepath -scriptblock {Param($pipearray, $filepath)
foreach($entry in $pipearray)
    {
    $entry
    $ping = new-object system.net.networkinformation.ping
    $reply = $ping.send($entry)
    if($reply.status -eq "Success")
    {
	   "$entry is online" >> $filepath
    }
       Else
    {
        "$entry is OFFLINE" >> $filepath
    }
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
$endtime = get-date
$minutecount = ($endtime - $starttime).minutes
$secondcount = ($endtime - $starttime).seconds
"Script completed in $minutecount minutes and $secondcount seconds"
foreach($entry in (dir file*.txt)){gc $entry >> masterfile.txt}
$b = gc .\masterfile.txt
$down = $b | ?{$_ -like "*offline*"}
$up = $b | ?{$_ -like "*online*"}
$total = $b.count
$upcount = $up.count
$downcount = $down.count
Write-host "Total number of systems: " -nonewline; write-host $total -foregroundcolor yellow
write-host "Total Online: " -nonewline; write-host $upcount -foregroundcolor yellow
write-host "Total Offline: " -nonewline; write-host $downcount -foregroundcolor yellow
write-host "Script Complete" -foregroundcolor green
