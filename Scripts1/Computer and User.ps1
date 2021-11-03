$a = New-Object -comobject Excel.Application
$a.visible = $True

$b = $a.Workbooks.Add()
$c = $b.Worksheets.Item(1)

$c.Cells.Item(1,1) = "Computer"
$c.Cells.Item(1,2) = "User"

$d = $c.UsedRange
$d.Interior.ColorIndex = 19
$d.Font.ColorIndex = 11
$d.Font.Bold = $True

$intRow = 2

$Computers = Get-Content "C:\Users\1252862141.adm\Desktop\Scripts1\Enable_Local_Admin.txt"

ForEach ($Computer in $Computers)
{
    $c.Cells.Item($intRow,1) = $Computer
    
    If (Test-Connection $Computer -Quiet -BufferSize 16 -Count 1 -Ea 0)
    {
        $User = Get-WmiObject Win32_ComputerSystem -ComputerName $Computer
        If ($User.UserName -ne $null)
        {
            $EDI = $User.UserName.TrimStart("AREA52\")
            $UserInfo = Get-ADUser "$EDI" -Properties DisplayName
            $c.Cells.Item($intRow,2) = $UserInfo.DisplayName
        }
        Else
        {
            $c.Cells.Item($intRow,2) = "No user"
        }
    }
    Else
    {
        $c.Cells.Item($intRow,2) = "Offline"
        Write-Host "$Computer offline"
    }

    $intRow = $intRow + 1
}

$d.EntireColumn.AutoFit()

$b.SaveAs("C:\work\Contacts_User_Info.xls")