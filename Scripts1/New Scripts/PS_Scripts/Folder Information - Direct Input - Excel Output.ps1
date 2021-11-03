#----------------------------------------------------------------------------------
#                           Written by SrA Timothy Brady
#                           Tyndall AFB, Panama City, FL
#                             Created February 5, 2014
#----------------------------------------------------------------------------------

#Opening Excel
$a = New-Object -comobject Excel.Application
$a.visible = $True

$b = $a.Workbooks.Add()
$c = $b.Worksheets.Item(1)

#Column titles
$c.Cells.Item(1,1) = "Folder Name"
$c.Cells.Item(1,2) = "Last Write Time"
$c.Cells.Item(1,3) = "Folder Size (MB)"
$c.Cells.Item(1,4) = "Folder Size (GB)"

#Title cell formatting
$d = $c.UsedRange
$d.Interior.ColorIndex = 19
$d.Font.ColorIndex = 11
$d.Font.Bold = $True

$intRow = 2

#Directory you want to search
$Directory = "\\XLWU-FS-001\ANG$\Users"

#Gets all the folders you want the information on
$FolderName = Get-ChildItem $Directory
$Folder = Get-ChildItem $Directory -Recurse

#Measures the total size of all the sub folders/files
$FolderSize = $Folder | Measure-Object -Property length -Sum
    
#Math to get it in a common value GB/MB (N1 = 1 decimal place)
$FolderSizeGB = "{0:N2}" -f ($FolderSize.sum/1GB)
$FolderSizeMB = "{0:N1}" -f ($FolderSize.sum/1MB)

#Populating information into Excel
$c.Cells.Item($intRow,1) = $FolderName
$c.Cells.Item($intRow,2) = $Folder.LastWriteTime
$c.Cells.Item($intRow,3) = $FolderSizeMB
$c.Cells.Item($intRow,4) = $FolderSizeGB
$intRow = $intRow + 1

$d.EntireColumn.AutoFit()