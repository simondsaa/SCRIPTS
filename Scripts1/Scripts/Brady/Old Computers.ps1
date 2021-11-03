$Date = (Get-Date).AddDays(-365)
$domain = "OU=Tyndall AFB,OU=AFCONUSEAST,OU=Bases,DC=AREA52,DC=AFNOAPPS,DC=USAF,DC=MIL"
$objDomain = [adsi]("LDAP://" + $domain)
$search = New-Object System.DirectoryServices.DirectorySearcher
$search.SearchRoot = $objDomain
$search.Filter = "(&(objectClass=computer)(lastlogon=*))"
$search.SearchScope = "Subtree"
$results = $search.FindAll()
$Array = @()
ForEach($item in $results)
{
    $objComputer = $item.GetDirectoryEntry()
    $Name = $objComputer.cn
    $Logon = Get-Date -Date ([DateTime]::FromFileTime($objComputer.ConvertLargeIntegerToInt64($objComputer.lastlogon[0])))
    $Stamp = Get-Date -Date ([DateTime]::FromFileTime($objComputer.ConvertLargeIntegerToInt64($objComputer.lastlogontimestamp[0])))
    If ($Date -gt $Logon)
    {
        $obj = New-Object PSObject
        $obj | Add-Member -Force -MemberType NoteProperty -Name "Computer Name" -Value $Name
        $obj | Add-Member -Force -MemberType NoteProperty -Name "Last Logon" -Value $Logon
        $obj | Add-Member -Force -MemberType NoteProperty -Name "Last Logon Stamp" -Value $Stamp
        $Array += $obj
    }
}
$Array | OGV -Title "Old Computers"