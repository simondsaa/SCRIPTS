$a = New-Object -comobject Excel.Application
$a.visible = $True

$b = $a.Workbooks.Add()
$c = $b.Worksheets.Item(1)

$c.Cells.Item(1,1) = "Display Name"
$c.Cells.Item(1,2) = "EDI Number"
$c.Cells.Item(1,3) = "Organization"

$d = $c.UsedRange
$d.Interior.ColorIndex = 19
$d.Font.ColorIndex = 11
$d.Font.Bold = $True

$intRow = 2

$Users = Get-ADUser -SearchBase "OU=Tyndall AFB Users,OU=Tyndall AFB,OU=AFCONUSEAST,OU=Bases,DC=AREA52,DC=AFNOAPPS,DC=USAF,DC=MIL" -Filter * -Properties DisplayName, SamAccountName | 
Where-Object {($_.DisplayName -match "AOC") -or ($_.DisplayName -match "AFNORTH") -or ($_.DisplayName -match "AFRCC")}
ForEach ($User in $Users)
{
    $Org = (Out-String -InputObject $User.o).Trim()
    $c.Cells.Item($intRow,1) = $User.DisplayName
    $c.Cells.Item($intRow,2) = $User.SamAccountName
    $c.Cells.Item($intRow,3) = $Org

    $intRow = $intRow + 1
}

$d.EntireColumn.AutoFit()

$b.SaveAs("C:\Temp\Tyndall_Users.xls")
$b.Close()

$a.Quit()

