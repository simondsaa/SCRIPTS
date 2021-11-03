$Start = Get-Date

$Path = "C:\Users\1392134782A\Desktop\Disabled Accounts.xlsx"

$a = New-Object -comobject Excel.Application
$a.visible = $True

$b = $a.Workbooks.Add()
$c = $b.Worksheets.Item(1)

$c.Cells.Item(1,1) = "User Display Name"
$c.Cells.Item(1,2) = "User EDI Number"
$c.Cells.Item(1,3) = "Last Logon Date"
$c.Cells.Item(1,4) = "AFDS Comment"

$d = $c.UsedRange
$d.Interior.ColorIndex = 19
$d.Font.ColorIndex = 11
$d.Font.Bold = $True

$intRow = 2

$OU = "OU=Tyndall AFB Users,OU=Tyndall AFB,OU=AFCONUSEAST,OU=Bases,DC=AREA52,DC=AFNOAPPS,DC=USAF,DC=MIL"

$Users = Search-ADAccount -UsersOnly -SearchBase $OU -AccountDisabled
ForEach ($User in $Users)
{
    $Info = Get-ADUser -Identity $User.SamAccountName -Properties DisplayName, SamAccountName, info, LastLogon
    $DisplayName = $Info.DisplayName
    $EDI = $Info.SamAccountName
    $Comment = $Info.info
    $LastLogon = [DateTime]::FromFileTime($Info.LastLogon)

    If ($Info.info -like "*DISABLED LOST DMDC INFORMATION*")
    {
        $c.Cells.Item($intRow,1) = "$DisplayName"
        $c.Cells.Item($intRow,2) = "$EDI"
        $c.Cells.Item($intRow,3) = "$LastLogon"
        $c.Cells.Item($intRow,4) = "$Comment"

        $intRow = $intRow + 1
    }
}
$d.EntireColumn.AutoFit()

$b.SaveAs($Path)
$b.Close()

$a.Quit()

$Stop = Get-Date
$TimeS = ($Stop - $Start).Seconds
$TimeM = [Math]::Round(($Stop - $Start).TotalMinutes, 0)
Write-Host

$Users.Count

Write-Host
Write-Host "Elapsed Time: $TimeM min $TimeS sec" -ForegroundColor Cyan