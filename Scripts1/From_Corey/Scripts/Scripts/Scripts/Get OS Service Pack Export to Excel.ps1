$erroractionpreference = "SilentlyContinue" 
# Create a New Excel Object for storing Data 
$a = New-Object -comobject Excel.Application 
$a.visible = $True  
$b = $a.Workbooks.Add() 
$c = $b.Worksheets.Item(1) 
# Create the title row 
$c.Cells.Item(1,1) = "Machine Name" 
$c.Cells.Item(1,2) = "OS" 
$c.Cells.Item(1,3) = "Description" 
$c.Cells.Item(1,4) = "Service Pack" 
$d = $c.UsedRange 
$d.Interior.ColorIndex = 23 
$d.Font.ColorIndex = 2 
$d.Font.Bold = $True 
$d.EntireColumn.AutoFit($True) 
$intRow = 2 
$colComputers = get-content C:\Users\timothy.brady\Desktop\Comps.txt
# Run through the Array of Computers 
foreach ($strComputer in $colComputers) 
{$c.Cells.Item($intRow, 1) = $strComputer.ToUpper() 
# Get Operating System Info 
$colOS =Get-WmiObject -class Win32_OperatingSystem -computername $Strcomputer 
foreach($objComp in $colOS) 
{$c.Cells.Item($intRow, 2) = $objComp.Caption 
$c.Cells.Item($intRow, 3) = $objComp.Description 
$c.Cells.Item($intRow, 4) = $objComp.ServicePackMajorVersion} 
$intRow = $intRow + 1} 
$intRow = $intRow + 1 
# Save workbook data 
$b.SaveAs("C:\Users\timothy.brady\Desktop\OS.xlsx") 
# Quit Excel (Remove "#" if you want to quit Excel after the script is completed)