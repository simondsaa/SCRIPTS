$Path = "C:\Users\timothy.brady\Desktop\Users.txt"
$domain = "OU=Tyndall AFB,OU=AFCONUSEAST,OU=Bases,DC=AREA52,DC=AFNOAPPS,DC=USAF,DC=MIL"
$objDomain = [adsi]("LDAP://" + $domain)
$search = New-Object System.DirectoryServices.DirectorySearcher
$search.SearchRoot = $objDomain
$search.Filter = "(&(&(&(objectCategory=person)(objectClass=user)(!userAccountControl:1.2.840.113556.1.4.803:=262144))))"
$search.SearchScope = "Subtree"
$results = $search.FindAll()
foreach($item in $results)
{
    $objUser = $item.GetDirectoryEntry()
    $Name = $objUser.displayname
    $Logon = $objUser.sAMAccountName
    $Description = $objUser.description
    $Today=(Get-Date).AddDays(-10)
    Write-Output "$Name
$Logon
$Description
" | Out-File $Path -append
}