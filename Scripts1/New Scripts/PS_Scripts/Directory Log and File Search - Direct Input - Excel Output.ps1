#-----------------------------------------------------------------------------------------#
#                               Written by SrA Timothy Brady                              #
#                               Tyndall AFB, Panama City, FL                              #
#-----------------------------------------------------------------------------------------#

$a = New-Object -comobject Excel.Application
$a.visible = $True

$b = $a.Workbooks.Add()
$c = $b.Worksheets.Item(1)

$c.Cells.Item(1,1) = "Directory"
$c.Cells.Item(1,2) = "Name"
$c.Cells.Item(1,3) = "Size (KB/MB)"
$c.Cells.Item(1,4) = "Creation Date"
$c.Cells.Item(1,5) = "Last Access Date"
$c.Cells.Item(1,6) = "Last Modified Date"
$c.Cells.Item(1,8) = "Directory Information"

$d = $c.UsedRange
$d.Interior.ColorIndex = 19
$d.Font.ColorIndex = 11
$d.Font.Bold = $True

#$MergeCells = $c.Range("H4:H6")
#$MergeCells.Select() 
#$MergeCells.MergeCells = $True

$intRow = 2

$Date = (Get-Date).AddDays(-0)
$FileType = Read-Host "File Type"
$Directory = "\\XLWU-FS-04pv\Tyndall_325_MSG\325 CS\SCO\SCOO"
$Directories = Get-ChildItem $Directory -Directory -Recurse -ErrorAction SilentlyContinue
$Files = Get-ChildItem $Directory -File -Recurse -ErrorAction SilentlyContinue
$Items = Get-ChildItem $Directory -Recurse -ErrorAction SilentlyContinue | Where {$_.Extension -like "*$FileType*"}
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
    $c.Cells.Item($intRow,3) = "$FileLengthKB KB/$FileLengthMB MB"
    $c.Cells.Item($intRow,4) = $File.CreationTime
    $c.Cells.Item($intRow,5) = $File.LastAccessTime
    $c.Cells.Item($intRow,6) = $File.LastWriteTime
    
    $intRow = $intRow + 1
}

$FileCount = $Files.Count
$FolderCount = $Directories.Count
$FilesFound = $Items.Count

$c.Cells.Item(2,8) = "$Directory"
$c.Cells.Item(3,8) = "Total Folders: $FolderCount"
$c.Cells.Item(4,8) = "Total Files: $FileCount"
$c.Cells.Item(5,8) = "Total Size: $DirectorySizeGB GB/$DirectorySizeMB MB"
$c.Cells.Item(6,8) = "File Type Searched: $FileType"
$c.Cells.Item(7,8) = "Total Files Found: $FilesFound"
$c.Cells.Item(8,8) = "Total Size: $FilesFoundSizeGB GB/$FilesFoundSizeMB MB"
#$c.Cells.Item(9,8) = " "

$d.EntireColumn.AutoFit()

$b.SaveAs("C:\Users\1392134782A\Desktop\Work_Stuff\File Scans\File Info.xls")
$b.Close()

$a.Quit()

$PopUp = New-Object -Comobject wscript.shell
$Go = $PopUp.popup("The search has completed.",0,"* Completed *",80)