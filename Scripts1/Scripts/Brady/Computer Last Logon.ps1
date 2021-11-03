$a = New-Object -comobject Excel.Application
$a.visible = $True

$b = $a.Workbooks.Add()
$c = $b.Worksheets.Item(1)

$c.Cells.Item(1,1) = "Computer Name"
$c.Cells.Item(1,2) = "Last Logon"
$c.Cells.Item(1,3) = "Organization"
$c.Cells.Item(1,4) = "Building"
$c.Cells.Item(1,5) = "Room"

$d = $c.UsedRange
$d.Interior.ColorIndex = 19
$d.Font.ColorIndex = 11
$d.Font.Bold = $True

$intRow = 2

$Computers = Get-Content C:\work\TEST.txt
ForEach ($Computer in $Computers)
{
    $AD = Get-ADComputer $Computer -Properties o, Location, LastLogon
    $Org = ($AD.o | Out-String).Trim()
    $Bldg = $AD.Location.Split(";")[0]
    $Room = $AD.Location.Split(";")[1].TrimStart(" ")
    $Logon = [DateTime]::FromFileTime($AD.LastLogon).ToString('g')

    $c.Cells.Item($intRow,1) = $Computer
    $c.Cells.Item($intRow,2) = $Logon
    $c.Cells.Item($intRow,3) = $Org
    $c.Cells.Item($intRow,4) = $Bldg
    $c.Cells.Item($intRow,5) = $Room
        
    $intRow = $intRow + 1
}

$d.EntireColumn.AutoFit()

$b.SaveAs("C:\work\System_Last_Logon.xlsx")