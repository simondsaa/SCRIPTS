$domain = "OU=Tyndall AFB,OU=AFCONUSEAST,OU=Bases,DC=AREA52,DC=AFNOAPPS,DC=USAF,DC=MIL"
$objDomain = [adsi]("LDAP://" + $domain)
$search = New-Object System.DirectoryServices.DirectorySearcher
$search.SearchRoot = $objDomain
$search.Filter = "(&(&(sAMAccountType=805306369)(objectCategory=computer)(objectClass=computer)(operatingSystemVersion=10*)))"
$search.SearchScope = "Subtree"
$search.PageSize = 99999
$results = $search.FindAll()

ForEach($item in $results)
{
    $objComputer = $item.GetDirectoryEntry()
    $Name = $objComputer.cn
    Write-Host "$Name"
    $Name | Out-File -FilePath C:\Temp\BaseComputers.txt -Append -Force
}