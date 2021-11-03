#-----------------------------------------------------------------------------------------#
#                       Written by SrA Timothy Brady                           #
#                        Tyndall AFB, Panama City, FL                           #
#----------------------------------------------------------------------------------------#

$a = New-Object -comobject Excel.Application
$a.visible = $True

$b = $a.Workbooks.Add()
$c = $b.Worksheets.Item(1)

$c.Cells.Item(1,1) = "Directory"
$c.Cells.Item(1,2) = "Name"
$c.Cells.Item(1,3) = "Size (MB)"
$c.Cells.Item(1,4) = "Creation Date"
$c.Cells.Item(1,5) = "Last Access Date"
$c.Cells.Item(1,6) = "Last Modified Date"
$c.Cells.Item(1,8) = "Directory Information"

$d = $c.UsedRange
$d.Interior.ColorIndex = 19
$d.Font.ColorIndex = 11
$d.Font.Bold = $True

$intRow = 2

$Directory = "C:\Users\Timothy.Brady\Desktop\Work Stuff"
$Directories = Get-ChildItem $Directory -Directory -Recurse -ErrorAction SilentlyContinue
$Files = Get-ChildItem $Directory -File -Recurse -ErrorAction SilentlyContinue
$Items = Get-ChildItem $Directory -Recurse -ErrorAction SilentlyContinue
$DirectorySize = $Files | Measure-Object -Property length -Sum
$DirectorySizeGB = "{0:N1}" -f ($DirectorySize.sum/1GB)

ForEach ($File in $Items)
{
    $FileLength = "{0:N0}" -f ($File.Length/1MB)
    $c.Cells.Item($intRow,1) = $File.Directory
    $c.Cells.Item($intRow,2) = $File
    $c.Cells.Item($intRow,3) = "$FileLength MB"
    $c.Cells.Item($intRow,4) = $File.CreationTime
    $c.Cells.Item($intRow,5) = $File.LastAccessTime
    $c.Cells.Item($intRow,6) = $File.LastWriteTime
$intRow = $intRow + 1
}

$FileCount = $Files.Count
$FolderCount = $Directories.Count

$c.Cells.Item(2,8) = "$Directory"
$c.Cells.Item(3,8) = "Total Folders: $FolderCount"
$c.Cells.Item(4,8) = "Total Files: $FileCount"
$c.Cells.Item(5,8) = "Total Size: $DirectorySizeGB GB"

$d.EntireColumn.AutoFit()