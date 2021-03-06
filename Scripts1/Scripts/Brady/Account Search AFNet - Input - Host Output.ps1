$displayname = Read-Host "Name"
$domain = "OU=Tyndall AFB Computers,OU=Tyndall AFB,OU=AFCONUSEAST,OU=Bases,DC=AREA52,DC=AFNOAPPS,DC=USAF,DC=MIL"
$objDomain = [adsi]("LDAP://" + $domain)
$search = New-Object System.DirectoryServices.DirectorySearcher
$search.SearchRoot = $objDomain
$search.Filter = "(&(objectClass=user)(employeeType=*)(displayName=*$displayname*))"
$search.SearchScope = "Subtree"
$results = $search.FindAll()
foreach($item in $results)
{
    $objUser = $item.GetDirectoryEntry()
    #$Date = Get-Date -Date ([DateTime]::FromFileTime($objUser.ConvertLargeIntegerToInt64($objUser.lastlogon[0])))
    $Name = $objUser.displayname
    $L = $objUser.l
    #$Logon = $objUser.gigID
    #$LogonName = "AREA52\$Logon"
    Write-Host "Displayname - $Name"
    Write-Host "L Attribute - $L"
    Write-Host
    #$Today=(Get-Date).AddDays(-10)
#If ($Date -lt $Today){    
#    Write-Host -ForegroundColor Red "$Name has not logged in recently: $Date"}
#Else {Write-Host -ForegroundColor Green "$Name has logged in recently: $Date"}
#}
}