# Issue warning if % free disk space is less than:
$percentWarning = 15;
# Get server list:
$servers = Get-Content "C:\Users\timothy.brady\Desktop\Servers.txt";
foreach($server in $servers)
{	# Get drive info:
	$disks = Get-WmiObject -ComputerName $server -Class Win32_LogicalDisk -filter "DriveType = 3" -ErrorAction SilentlyContinue;
 	foreach($disk in $disks)
	{	$deviceID = $disk.DeviceID;
		[float]$size = $disk.Size;
		[float]$freespace = $disk.FreeSpace;
 
		$percentFree = [Math]::Round(($freespace / $size) * 100, 2);
        $sizeGB = [Math]::Round($size / 1073741824, 1);
		
        $colour = "Green";
		if($percentFree -lt $percentWarning)
		{$colour = "Red";}		
		Write-Host -ForegroundColor $colour "$server $deviceID (Total $sizeGB GB) (Used $usedGB GB) (Free: $freespaceGB GB) Percent Free: $percentFree%";
		}
}