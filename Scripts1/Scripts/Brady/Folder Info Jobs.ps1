$MaxThreads = 200

$Start = Get-Date

$i = $null

$SearchFolder = "\\xlwu-fs-03pv\root"
$Folders = Get-ChildItem $SearchFolder -Directory

$TotalJobs = $Folders.Count

$ScriptBlock = {

    $Days = 730

    $Date = (Get-Date).AddDays(-$Days)
    $Directory = $args[0].FullName
    $Directories = Get-ChildItem $Directory -Directory -Recurse -ErrorAction SilentlyContinue
    $Files = Get-ChildItem $Directory -File -Recurse -ErrorAction SilentlyContinue
    $Items = $Files | Where-Object {$_.LastAccessTime -lt $Date}

    $FolderCount = $Directories.Count
    $FileCount = $Files.Count
    $FilesFound = $Items.Count

    $DirectorySize = $Files | Measure-Object -Property length -Sum
    $FilesFoundSize = $Items | Measure-Object -Property length -Sum

    $DirectorySizeGB = "{0:N2}" -f ($DirectorySize.sum/1GB)
    $DirectorySizeMB = "{0:N1}" -f ($DirectorySize.sum/1MB)

    $FilesFoundSizeGB = "{0:N2}" -f ($FilesFoundSize.sum/1GB)
    $FilesFoundSizeMB = "{0:N1}" -f ($FilesFoundSize.sum/1MB)
    
    ForEach ($File in $Files)
    {
        $FileLengthKB = "{0:N0}" -f ($File.Length/1KB)
        $FileLengthMB = "{0:N0}" -f ($File.Length/1MB)
    
        $FileDirectory = $File.Directory
        $FileExtension = $File.Extension
        $FileCreationTime = $File.CreationTime
        $FileLastAccessTime = $File.LastAccessTime
        $FileLastWriteTime = $File.LastWriteTime
    
        $Results = [PSCustomObject]@{
            Folder = $args[0]
            Directory = $FileDirectory
            File_Name = $File
            File_Type = $FileExtension
            File_SizeKB = $FileLengthKB
            Created = $FileCreationTime
            Last_Access = $FileLastAccessTime
            Last_Write = $FileLastWriteTime
            }

        $Results | Export-Csv "C:\Users\1180219788A\Desktop\Scans\$($args[0]) Cleanup Scan.csv" -Append -Force
    }

     $Results2 = [PSCustomObject]@{
        Folder = $args[0]
        Total_Folders = $FolderCount
        Total_Files = $FileCount
        Total_SizeGB = $DirectorySizeGB
        Days_Old = $Days
        Old_Files = $FilesFound
        Old_SizeGB = $FilesFoundSizeGB
        }

    $Results2 | Export-Csv "C:\Users\1180219788A\Desktop\Scans\Cleanup Data.csv" -Append -Force
}

ForEach ($Folder in $Folders)
{
    Write-Host "Starting Job on: $Folder" -ForegroundColor Cyan
    $i++
    Write-Host "________________Status: $i / $TotalJobs" -ForegroundColor Yellow

    Start-Job -Name $Folder -ScriptBlock $ScriptBlock -ArgumentList $Folder | Out-Null

    While ($(Get-Job -State Running).Count -ge $MaxThreads)
    {
        Get-Job | Wait-Job -Any | Out-Null
    }
}

While ($(Get-Job -State Running).Count -ne 0)
{
    $JobCount = (Get-Job -State Running).Count
    Start-Sleep -Seconds 10
    Write-Host "Waiting for $JobCount Jobs to complete..." -ForegroundColor DarkYellow
}

$Stop = Get-Date
$TimeS = ($Stop - $Start).Seconds
$TimeM = [Math]::Round(($Stop - $Start).TotalMinutes, 0)
Write-Host
Write-Host "Elapsed Time: $TimeM min $TimeS sec" -ForegroundColor Cyan

Get-Job | Remove-Job -Force

$PopUp = New-Object -Comobject wscript.shell
$Go = $PopUp.popup("The search has completed.",0,"* Folder Info Jobs *",80)