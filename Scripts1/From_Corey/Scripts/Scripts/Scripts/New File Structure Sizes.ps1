$Directories = Get-Content "C:\Users\timothy.brady\Desktop\Paths.txt"
ForEach($Directory in $Directories){
$Files = Get-ChildItem $Directory -Recurse -ErrorAction SilentlyContinue
$FileSize = $Files | Measure-Object -Property length -Sum -ErrorAction SilentlyContinue
$Name = $Files.Count
$SizeGB = "{0:N1}" -f ($FileSize.sum/1GB)
$SizeMB = "{0:N1}" -f ($FileSize.sum/1MB)
Write-Host
Write-Host "Total Files: $Name"
Write-Host "Total Size:  $SizeGB GB"
Write-Host "             $SizeMB MB"
Write-Host "File Path:   $Directory"}