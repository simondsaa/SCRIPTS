$Group = Read-Host "Group"

$Path = "\\xlwu-fs-004\Home\1392134782A\Desktop\$Group Membership.xls"

$a = New-Object -comobject Excel.Application
$a.visible = $True

$b = $a.Workbooks.Add()
$c = $b.Worksheets.Item(1)

$c.Cells.Item(1,1) = "$Group"
$c.Cells.Item(2,1) = "User Display Name"
$c.Cells.Item(2,2) = "User EDI Number"

$d = $c.UsedRange
$d.Interior.ColorIndex = 19
$d.Font.ColorIndex = 11
$d.Font.Bold = $True

$intRow = 3

$Names = Get-ADGroupMember -Identity $Group | Select *
ForEach ($User in $Names)
{
    $DisplayName = (Get-ADUser $User.SamAccountName -Properties DisplayName).DisplayName
    $c.Cells.Item($intRow,1) = $DisplayName
    $c.Cells.Item($intRow,2) = $User.SamAccountName
    $intRow = $intRow + 1
}

$d.EntireColumn.AutoFit()

$b.SaveAs($Path)
$b.Close()

$a.Quit()