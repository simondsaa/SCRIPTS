$domain = "OU=Tyndall AFB,OU=AFCONUSEAST,OU=Bases,DC=AREA52,DC=AFNOAPPS,DC=USAF,DC=MIL"
$objDomain = [adsi]("LDAP://" + $domain)
$Where = Read-Host "PC List"
$Computers = Get-Content $Where
$Path = "C:\Temp\AD Computer Properties.txt"
If (Test-Path $Path){Remove-Item $Path}
ForEach($computer in $Computers){
    $search = New-Object System.DirectoryServices.DirectorySearcher
    $search.SearchRoot = $objDomain
    $search.Filter = "(&(objectClass=computer)(cn=*$Computer*))"
    $search.SearchScope = "Subtree"
    $results = $search.FindAll()
    ForEach($item in $results){
        $objComputer = $item.GetDirectoryEntry()
        $Name = $objComputer.cn
        $Loc = $objComputer.Location
        Write-Output "$Name; $Loc" | Out-File $Path -append
            }
        }
$file = “$Path”
$oXL = New-Object -comobject Excel.Application
$oXL.Visible = $true
$oXL.workbooks.OpenText($file,1,1,1,1,$True,$True,$True,$False,$False,$False)

# 1   Tab = True
# 2   Semicolon = True
# 3   Comma = False
# 4   Space = False
# 5   Other = False