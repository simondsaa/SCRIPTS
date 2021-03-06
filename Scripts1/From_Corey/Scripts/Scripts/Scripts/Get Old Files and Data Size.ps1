$Date = (Get-Date).AddDays(-1825)
$Directory = "\\XLWU-FS-004\root\325 FW\325 MSG\325 CS"
$Files = Get-ChildItem $Directory -Recurse -ErrorAction SilentlyContinue | Where-Object {$_.LastWriteTime -lt $Date}
$Output = $Files | Select Directory, Name, CreationTime, LastAccessTime, LastWriteTime | Export-CSV C:\Users\timothy.brady\Desktop\Old_Files.csv
$FileSize = $Files | Measure-Object -Property length -Sum
$Name = $Files.Count
$SizeGB = "{0:N1}" -f ($FileSize.sum/1GB)
$SizeMB = "{0:N1}" -f ($FileSize.sum/1MB)
$SizeKB = "{0:N1}" -f ($FileSize.sum/1KB)
Write-Host
Write-Host "Total Files: $Name"
Write-Host "Total Size:  $SizeGB GB"
Write-Host "             $SizeMB MB"
Write-Host "             $SizeKB KB"
Write-Host "File Path:   $Directory"