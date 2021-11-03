$Start = Get-Date

$Path = "C:\Users\1392134782A\Desktop\Disabled Accounts.xlsx"

$a = New-Object -comobject Excel.Application
$a.visible = $True

$b = $a.Workbooks.Add()
$c = $b.Worksheets.Item(1)

$c.Cells.Item(1,1) = "User Display Name"
$c.Cells.Item(1,2) = "User EDI Number"
$c.Cells.Item(1,3) = "AFDS Comment"

$d = $c.UsedRange
$d.Interior.ColorIndex = 19
$d.Font.ColorIndex = 11
$d.Font.Bold = $True

$intRow = 2

$Domain = "OU=Tyndall AFB Users,OU=Tyndall AFB,OU=AFCONUSEAST,OU=Bases,DC=AREA52,DC=AFNOAPPS,DC=USAF,DC=MIL"
$ObjDomain = [adsi]("LDAP://" + $domain)
$Search = New-Object System.DirectoryServices.DirectorySearcher
$Search.SearchRoot = $objDomain
$Search.Filter = "(&(objectClass=user)(userAccountControl:1.2.840.113556.1.4.803:=2))"
$Search.SearchScope = "Subtree"
$Search.PageSize = 99999
$Results = $search.FindAll()

ForEach ($Item in $Results)
{
    $ObjUser = $Item.GetDirectoryEntry()
    $DisplayName = $ObjUser.DisplayName
    $EDI = $ObjUser.SamAccountName
    $Comment = $ObjUser.info
    If ($ObjUser.info -like "*AFDS*")
    {
        $c.Cells.Item($intRow,1) = "$DisplayName"
        $c.Cells.Item($intRow,2) = "$EDI"
        $c.Cells.Item($intRow,3) = "$Comment"

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
Write-Host "Elapsed Time: $TimeM min $TimeS sec" -ForegroundColor Cyan