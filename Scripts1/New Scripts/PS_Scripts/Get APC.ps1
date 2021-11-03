$a = New-Object -comobject Excel.Application
$a.visible = $True

$b = $a.Workbooks.Add()
$c = $b.Worksheets.Item(1)

$c.Cells.Item(1,1) = "User"
$c.Cells.Item(1,2) = "Server"
$c.Cells.Item(1,3) = "APC"

$d = $c.UsedRange
$d.Interior.ColorIndex = 19
$d.Font.ColorIndex = 11
$d.Font.Bold = $True

$intRow = 2

$Users = Get-Content "C:\Users\1392134782A\Desktop\Tyndall Users.txt"

ForEach ($User in $Users)
{
    $Info = Get-ADUser -Identity $User -Properties DisplayName, homeMTA
    $MTA = $Info.homeMTA
    $Name = $Info.DisplayName
    $Server = ($MTA -split "CN=").Split(",")[3]
    
    If ($Server -eq $null)
    {
        $Servername = "No server"
        $APC = "Unavailable"
    }
    
    Else
    {
        $Servername = $Server
        $APC = ((Get-ADComputer -Identity $Server).DistinguishedName -Split "OU=").Split(",")[2]
    }

    $c.Cells.Item($intRow,1) = $Name
    $c.Cells.Item($intRow,2) = $Servername
    $c.Cells.Item($intRow,3) = $APC

    $intRow = $intRow + 1
}

$d.EntireColumn.AutoFit()

$b.SaveAs("C:\Users\1392134782A\Desktop\User_APCs.xlsx")
$b.Close()

$a.Quit()