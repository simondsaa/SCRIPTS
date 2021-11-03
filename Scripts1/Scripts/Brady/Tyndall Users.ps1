$AOU = "OU=Tyndall AFB,OU=Administrative Accounts,OU=Administration,DC=AREA52,DC=AFNOAPPS,DC=USAF,DC=MIL"
$OU = "OU=Tyndall AFB Users,OU=Tyndall AFB,OU=AFCONUSEAST,OU=Bases,DC=AREA52,DC=AFNOAPPS,DC=USAF,DC=MIL"

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

$Users = Get-ADUser -SearchBase $OU -Filter * -Properties DisplayName, SamAccountName, o
ForEach ($User in $Users)
{
    $Org = (Out-String -InputObject $User.o).Trim()
    $c.Cells.Item($intRow,1) = $User.DisplayName
    $c.Cells.Item($intRow,2) = $User.SamAccountName
    $c.Cells.Item($intRow,3) = $Org

    $intRow = $intRow + 1
}

$d.EntireColumn.AutoFit()

$b.SaveAs("C:\Users\1252862141.adm\Desktop\Tyndall_Users.xls")
$b.Close()

$a.Quit()

$PopUp = New-Object -Comobject wscript.shell
$Go = $PopUp.popup("The search has completed.",0,"* Completed *",80)