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

#Title cell formatting
$d = $c.UsedRange
$d.Interior.ColorIndex = 19
$d.Font.ColorIndex = 11
$d.Font.Bold = $True

$intRow = 2

#Directory you want to search
$Directory = "\\XLWU-FS-04pv\Tyndall_325_MSG"

#Gets all the folders you want the information on
$Folders = Get-ChildItem $Directory
ForEach ($Folder in $Folders)
{
    #This is where it takes each folder found above, then searches within it and finds every sub folder/file
    $SubFolder = Get-ChildItem $Directory\$Folder -Recurse -ErrorAction SilentlyContinue
    
    #Measures the total size of all the sub folders/files
    $FolderSize = $SubFolder | Measure-Object -Property length -Sum
    
    #Math to get it in a common value GB/MB (N1 = 1 decimal place)
    $FolderSizeGB = "{0:N2}" -f ($FolderSize.sum/1GB)
    $FolderSizeMB = "{0:N1}" -f ($FolderSize.sum/1MB)

    #Populating information into Excel
    $c.Cells.Item($intRow,1) = $Folder
    $c.Cells.Item($intRow,2) = $Folder.LastWriteTime
    $c.Cells.Item($intRow,3) = $FolderSizeMB
    $intRow = $intRow + 1
}

$d.EntireColumn.AutoFit()