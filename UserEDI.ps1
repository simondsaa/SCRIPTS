$Names = Read-Host "Name"
ForEach ($displayname in $Names)
{
    $domain = "OU=Tyndall AFB,OU=AFCONUSEAST,OU=Bases,DC=AREA52,DC=AFNOAPPS,DC=USAF,DC=MIL"
    $objDomain = [adsi]("LDAP://" + $domain)
    $search = New-Object System.DirectoryServices.DirectorySearcher
    $search.SearchRoot = $objDomain
    $search.Filter = "(&(objectClass=user)(employeeType=*)(displayName=*$displayname*))"
    $search.SearchScope = "Subtree"
    $results = $search.FindAll()
    ForEach($item in $results)
    {
        $objUser = $item.GetDirectoryEntry()
        $Name = $objUser.displayname
        $Logon = $objUser.gigID
        Write-Host $Name $Logon 
    }
}
$EDI = Read-Host "User's EDI"
$UserInfo = Get-ADUser $EDI -Properties DisplayName, City, gigID, EmailAddress, extensionAttribute5, mDBOverHardQuotaLimit, LockedOut, Enabled, OfficePhone, Memberof
    $MailSize = ($UserInfo.mDBOverHardQuotaLimit/1024)
    #$AdGroup = Get-ADObject -SearchBase "OU=Tyndall AFB,OU=AFCONUSEAST,OU=Bases,DC=AREA52,DC=AFNOAPPS,DC=USAF,DC=MIL" -filter "CN=*"
    Write-Host
    Write-Host "Display Name       :" $UserInfo.DisplayName
    Write-Host "Base Name          :" $UserInfo.City
    Write-Host "EDIPI Number       :" $UserInfo.gigID
    Write-Host "Email Address      :" $UserInfo.EmailAddress
    Write-Host "Mail Category      :" $UserInfo.extensionAttribute5
    Write-Host "Box Size Limit     : $MailSize MB"
    Write-Host "Account Locked Out :" $UserInfo.LockedOut
    Write-Host "Account Enabled    :" $UserInfo.Enabled
    Write-Host "Office Phone       :" $UserInfo.OfficePhone
    #Write-Host "Group Membership   :" $UserInfo.memberOf | Select $AdGroup