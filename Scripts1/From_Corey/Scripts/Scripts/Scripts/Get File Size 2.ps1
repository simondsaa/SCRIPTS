#-----------------------------------------------------------------------------------------#
#                                  Written by SrA Timothy Brady                           #
#                                  Tyndall AFB, Panama City, FL                           #
#-----------------------------------------------------------------------------------------#
$Directory = "\\XLWU-FS-001\root\325 MSG\325 MSS"
$Files = Get-ChildItem $Directory -Recurse -ErrorAction SilentlyContinue
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