# Issue warning if % free disk space is less than:
$percentWarning = 15
# Get server list:
$servers = Get-Content "C:\Users\timothy.brady\Desktop\Servers.txt"
foreach($server in $servers)
{   # Get drive info:
    $disks = Get-WmiObject -ComputerName $server -Class Win32_LogicalDisk -filter "DriveType = 3" -ErrorAction SilentlyContinue
    foreach($disk in $disks)
    {   $deviceID = $disk.DeviceID
        $size = $disk.Size
        $freespace = $disk.FreeSpace

        $percentFree = [Math]::Round(($freespace/$size) * 100, 1)
        $sizeGB = [Math]::Round($size/1073741824, 1)
        $freespaceGB = [Math]::Round($freespace/1073741824, 1)
        $usedGB = [Math]::Round(($sizeGB - $freespaceGB), 1)
        
        If($percentFree -gt $percentWarning){$colour = "Green"}
        ElseIf($percentFree -lt $percentWarning){$colour = "Red"}
        Write-Host -ForegroundColor $colour "$server $deviceID (Total $sizeGB GB) (Used $usedGB GB) (Free: $freespaceGB GB) Percent Free: $percentFree%"        
        }
}