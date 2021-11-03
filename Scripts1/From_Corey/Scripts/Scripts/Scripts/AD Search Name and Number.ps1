$EDIPI = Read-Host "EDIPI"
$domain = "OU=!601,OU=Tyndall AFB,OU=AFCONUSEAST,OU=Bases,DC=AREA52,DC=AFNOAPPS,DC=USAF,DC=MIL"
$objDomain = [adsi]("LDAP://" + $domain)
$search = New-Object System.DirectoryServices.DirectorySearcher
$search.SearchRoot = $objDomain
$search.Filter = "(&(objectClass=user)(sAMAccountName=*$EDIPI*))"
$search.SearchScope = "Subtree"
$results = $search.FindAll()
foreach($item in $results)
{
    $objUser = $item.GetDirectoryEntry()
    $Name = $objUser.displayname
    $Number = $objUser.telephoneNumber
    $EDI = $objUser.gigID
    Write-Host "
    $Name
    $Number
    $EDI"
}