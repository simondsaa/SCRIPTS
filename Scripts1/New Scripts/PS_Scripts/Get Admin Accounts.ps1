$Path = "C:\Users\1392134782A\Desktop\Admin Accounts.xlsx"

$a = New-Object -comobject Excel.Application
$a.visible = $True

$b = $a.Workbooks.Add()
$c = $b.Worksheets.Item(1)

$c.Cells.Item(1,1) = "User Display Name"
$c.Cells.Item(1,2) = "User Account Name"
$c.Cells.Item(1,3) = "User EDI Number"

$d = $c.UsedRange
$d.Interior.ColorIndex = 19
$d.Font.ColorIndex = 11
$d.Font.Bold = $True

$intRow = 2

$Domain = "OU=Tyndall AFB,OU=Administrative Accounts,OU=Administration,DC=AREA52,DC=AFNOAPPS,DC=USAF,DC=MIL"
$ObjDomain = [adsi]("LDAP://" + $domain)
$Search = New-Object System.DirectoryServices.DirectorySearcher
$Search.SearchRoot = $objDomain
$Search.Filter = "(&(objectClass=user))"
$Search.SearchScope = "Subtree"
$Search.PageSize = 99999
$Results = $search.FindAll()

ForEach ($Item in $Results)
{
    $ObjUser = $Item.GetDirectoryEntry()
    $DisplayName = $ObjUser.DisplayName
    $Name = $ObjUser.Name
    $EDI = $ObjUser.SamAccountName
    $c.Cells.Item($intRow,1) = "$DisplayName"
    $c.Cells.Item($intRow,2) = "$Name"
    $c.Cells.Item($intRow,3) = "$EDI"
    $intRow = $intRow + 1
}

$d.EntireColumn.AutoFit()

$b.SaveAs($Path)
$b.Close()

$a.Quit()