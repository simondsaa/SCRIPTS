$Group = Read-Host "Group"

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

$Names = Get-ADGroupMember -Identity $Group
ForEach ($User in $Names)
{
    $c.Cells.Item($intRow,2) = $User.SamAccountName
    $EDIs = $User.SamAccountName
    ForEach ($EDI in $EDIs)
    {
        $UserInfo = Get-ADUser "$EDI" -Properties DisplayName
        $c.Cells.Item($intRow,1) = $UserInfo.DisplayName
    }
    $intRow = $intRow + 1
}

$d.EntireColumn.AutoFit()

$b.SaveAs($Path)
$b.Close()

$a.Quit()