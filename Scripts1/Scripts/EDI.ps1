$Names = Get-Content C:\temp\EDI.txt

$a = New-Object -comobject Excel.Application
$a.visible = $True

$b = $a.Workbooks.Add()
$c = $b.Worksheets.Item(1)

$c.Cells.Item(1,1) = "User Display Name"
$c.Cells.Item(1,2) = "User EDI Number"
$c.Cells.Item(1,3) = "User Email Address"

$d = $c.UsedRange
$d.Interior.ColorIndex = 19
$d.Font.ColorIndex = 11
$d.Font.Bold = $True

$intRow = 3

ForEach ($Name in $Names)
{
$c.Cells.Item($intRow,2) = $User.SamAccountName
    $EDIs = $User.SamAccountName
    ForEach ($EDI in $EDIs)
    {
        $UserInfo = Get-ADUser "$EDI" -Properties DisplayName, Email
        $c.Cells.Item($intRow,1) = $UserInfo.DisplayName
    }
    $intRow = $intRow + 1
}

$d.EntireColumn.AutoFit()

$b.SaveAs($Path)
$b.Close()

$a.Quit()