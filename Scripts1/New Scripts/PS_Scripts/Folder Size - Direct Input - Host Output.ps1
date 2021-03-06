#-----------------------------------------------------------------------------------------#
#                                  Written by SrA Timothy Brady                           #
#                                  Tyndall AFB, Panama City, FL                           #
#-----------------------------------------------------------------------------------------#
$Directory = "\\XLWU-FS-04pv\Tyndall_325_MSG\325 FSS"
$Directories = Get-ChildItem $Directory -ErrorAction SilentlyContinue

ForEach ($Folder in $Directories)
{
    $Files = Get-ChildItem "$Directory\$Folder" -Recurse -ErrorAction SilentlyContinue
    $FileSize = $Files | Measure-Object -Property length -Sum
    $Count = $Files.Count
    $SizeGB = "{0:N1}" -f ($FileSize.sum/1GB)
    $SizeMB = "{0:N1}" -f ($FileSize.sum/1MB)
    $SizeKB = "{0:N1}" -f ($FileSize.sum/1KB)
    Write-Host
    Write-Host "Folder Name: $Folder"
    Write-Host "Total Files: $Count"
    Write-Host "Total Size:  $SizeGB GB"
    Write-Host "             $SizeMB MB"
    Write-Host "             $SizeKB KB"
}