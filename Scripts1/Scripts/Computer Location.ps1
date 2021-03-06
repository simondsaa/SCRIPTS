$a = New-Object -comobject Excel.Application
$a.visible = $True

$b = $a.Workbooks.Add()
$c = $b.Worksheets.Item(1)

$c.Cells.Item(1,1) = "Computer"
$c.Cells.Item(1,2) = "User"
$c.Cells.Item(1,3) = "Bldg/Rm"

$d = $c.UsedRange
$d.Interior.ColorIndex = 19
$d.Font.ColorIndex = 11
$d.Font.Bold = $True

$intRow = 2
$Path = Read-Host "PC Names"
$Computers = Get-Content $Path

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
            $Fill = $c.Cells.Item($intRow,3) = $item.$Loc
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

$b.SaveAs("C:\temp\ansdcprob.xls")

$domain = "OU=Tyndall AFB,OU=AFCONUSEAST,OU=Bases,DC=AREA52,DC=AFNOAPPS,DC=USAF,DC=MIL"
$objDomain = [adsi]("LDAP://" + $domain)
ForEach($computer in $Computers)
{
    $search = New-Object System.DirectoryServices.DirectorySearcher
    $search.SearchRoot = $objDomain
    $search.Filter = "(&(objectClass=computer)(cn=*$Computer*))"
    $search.SearchScope = "Subtree"
    $results = $search.FindAll()
    ForEach($item in $results){
        $objComputer = $item.GetDirectoryEntry()
        $Name = $objComputer.cn
        $Loc = $objComputer.Location
        $Fill
            }
        }