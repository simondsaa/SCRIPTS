$Directory = "\\xlwu-fs-004\325 MSG"
$FileName = Read-Host "File Type"
$Files = Get-ChildItem $Directory -Recurse -ErrorAction SilentlyContinue | Where-Object {($_.Name -like "*$FileName*")}
$FileSize = $Files | Measure-Object -Property length -Sum
$Name = $Files.Count
$SizeGB = "{0:N1}" -f ($FileSize.sum/1GB)
$SizeMB = "{0:N1}" -f ($FileSize.sum/1MB)
$SizeKB = "{0:N1}" -f ($FileSize.sum/1KB)
Write-Host
Write-Host "Total Files    $Name$FileName(s)"
Write-Host "Total Size     $SizeGB GB"
Write-Host "               $SizeMB MB"
Write-Host "               $SizeKB KB"
Write-Host "Path Searched  $Directory"