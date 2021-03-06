$domain = "DC=tyndall,DC=AETC,DC=DS,DC=AF,DC=MIL"
$objDomain = [adsi]("LDAP://" + $domain)
$search = New-Object System.DirectoryServices.DirectorySearcher
$search.SearchRoot = $objDomain
$search.Filter = "(&(objectClass=user)(employeeType=*)(displayName=*OBoyle*))"
$search.SearchScope = "Subtree"
$results = $search.FindAll()
foreach($item in $results)
{
    $objUser = $item.GetDirectoryEntry()
    $Date = Get-Date -Date ([DateTime]::FromFileTime($objUser.ConvertLargeIntegerToInt64($objUser.lastlogontimestamp[0])))
    $Name = $objUser.displayname
    $Logon = $objUser.gigID
    $LogonName = "AREA52\$Logon"
    $Today=(Get-Date).AddDays(-10)
If ($Date -lt $Today){    
    Write-Host -ForegroundColor Red "$Name - $LogonName has not logged in recently: $Date"}
Else {Write-Host -ForegroundColor Green "$Name - $LogonName has logged in recently: $Date"}
}