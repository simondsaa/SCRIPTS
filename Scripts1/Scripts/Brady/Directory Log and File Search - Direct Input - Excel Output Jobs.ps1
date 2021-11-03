#-----------------------------------------------------------------------------------------#
#                               Written by SrA Timothy Brady                              #
#                               Tyndall AFB, Panama City, FL                              #
#-----------------------------------------------------------------------------------------#

$MaxThreads = 20

$Start = Get-Date

$i = $null

$SerachFolder = "\\xlwu-fs-04pv\Tyndall_325_MSG\325 CS\SCO\"
$Folders = Get-ChildItem $SerachFolder -Directory

$TotalJobs = $Folders.Count

$ScriptBlock = {

    $a = New-Object -comobject Excel.Application
    $a.visible = $True

    $b = $a.Workbooks.Add()
    $c = $b.Worksheets.Item(1)

    $c.Cells.Item(1,1) = "Directory"
    $c.Cells.Item(1,2) = "File Name"
    $c.Cells.Item(1,3) = "File Type"
    $c.Cells.Item(1,4) = "Size (KB/MB)"
    $c.Cells.Item(1,5) = "Creation Date"
    $c.Cells.Item(1,6) = "Last Access Date"
    $c.Cells.Item(1,7) = "Last Modified Date"
    $c.Cells.Item(1,9) = "Directory Information"

    $d = $c.UsedRange
    $d.Interior.ColorIndex = 19
    $d.Font.ColorIndex = 11
    $d.Font.Bold = $True

    #$MergeCells = $c.Range("H4:H6")
    #$MergeCells.Select() 
    #$MergeCells.MergeCells = $True

    $intRow = 2

    $Days = 180

    $Date = (Get-Date).AddDays(-$Days)
    $Directory = $args[0].FullName
    $Directories = Get-ChildItem $Directory -Directory -Recurse -ErrorAction SilentlyContinue
    $Files = Get-ChildItem $Directory -File -Recurse -ErrorAction SilentlyContinue
    $Items = Get-ChildItem $Directory -File -Recurse -ErrorAction SilentlyContinue | Where-Object {$_.LastWriteTime -lt $Date}
    $DirectorySize = $Files | Measure-Object -Property length -Sum
    $FilesFoundSize = $Items | Measure-Object -Property length -Sum

    $DirectorySizeGB = "{0:N2}" -f ($DirectorySize.sum/1GB)
    $DirectorySizeMB = "{0:N1}" -f ($DirectorySize.sum/1MB)

    $FilesFoundSizeGB = "{0:N2}" -f ($FilesFoundSize.sum/1GB)
    $FilesFoundSizeMB = "{0:N1}" -f ($FilesFoundSize.sum/1MB)

    ForEach ($File in $Items)
    {
        $FileLengthKB = "{0:N0}" -f ($File.Length/1KB)
        $FileLengthMB = "{0:N0}" -f ($File.Length/1MB)
    
        $c.Cells.Item($intRow,1) = $File.Directory
        $c.Cells.Item($intRow,2) = $File
        $c.Cells.Item($intRow,3) = $File.Extension
        $c.Cells.Item($intRow,4) = "$FileLengthKB KB/$FileLengthMB MB"
        $c.Cells.Item($intRow,5) = $File.CreationTime
        $c.Cells.Item($intRow,6) = $File.LastAccessTime
        $c.Cells.Item($intRow,7) = $File.LastWriteTime
    
        $intRow = $intRow + 1
    }

    $FileCount = $Files.Count
    $FolderCount = $Directories.Count
    $FilesFound = $Items.Count

    $c.Cells.Item(2,8) = "$Directory"
    $c.Cells.Item(3,8) = "Total Folders: $FolderCount"
    $c.Cells.Item(4,8) = "Total Files: $FileCount"
    $c.Cells.Item(5,8) = "Total Size: $DirectorySizeGB GB/$DirectorySizeMB MB"
    $c.Cells.Item(6,8) = "Files Older Than: $Days days"
    $c.Cells.Item(7,8) = "Total Files Found: $FilesFound"
    $c.Cells.Item(8,8) = "Total Size: $FilesFoundSizeGB GB/$FilesFoundSizeMB MB"
    #$c.Cells.Item(9,8) = " "

    $d.EntireColumn.AutoFit()

    $b.SaveAs("C:\Users\1392134782A\Desktop\Work_Stuff\File Scans\$($args[0]) Cleanup Scan.xlsx")
    $b.Close()

    $a.Quit()
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
$Go = $PopUp.popup("The search has completed.",0,"* Completed *",80)