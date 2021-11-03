$Group = Read-Host "Group"

$Path = "C:\Users\1392134782A\Desktop\$Group Membership.xlsx"

$a = New-Object -comobject Excel.Application
$a.visible = $True

$b = $a.Workbooks.Add()
$c = $b.Worksheets.Item(1)

$c.Cells.Item(1,1) = "$Group"
$c.Cells.Item(2,1) = "Computer"
$c.Cells.Item(2,2) = "Organization"

$d = $c.UsedRange
$d.Interior.ColorIndex = 19
$d.Font.ColorIndex = 11
$d.Font.Bold = $True

$intRow = 3

$Names = Get-ADGroupMember -Identity $Group | Select *
ForEach ($Computer in $Names)
{
    $AD = Get-ADComputer -Identity $Computer.name -Properties o
    $Org = (Out-String -InputObject $AD.o).Trim()
    $c.Cells.Item($intRow,1) = $Computer.name
    $c.Cells.Item($intRow,2) = $Org
    $intRow = $intRow + 1
}

$d.EntireColumn.AutoFit()

$b.SaveAs($Path)
$b.Close()

#$a.Quit()