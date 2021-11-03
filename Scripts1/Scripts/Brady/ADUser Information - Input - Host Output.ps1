$GroupName = ""
$EDI = Read-Host "Pre-Windows 2000 Name"
$UserInfo = Get-ADUser $EDI -Properties DisplayName, City, gigID, EmailAddress, extensionAttribute5, mDBOverHardQuotaLimit, LockedOut, Enabled, OfficePhone, MemberOf
Try 
{
    $Groups = Get-ADPrincipalGroupMembership "$EDI" -ErrorAction SilentlyContinue
    $GroupName = ""
    ForEach ($Group in $Groups)
    {
        $GroupName += $Group.Name + "
                     "
    }
}
Catch
{
    $GroupName = "Error"
}
    $MailSize = ($UserInfo.mDBOverHardQuotaLimit/1024)
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
    Write-Host "Group Membership   :" $GroupName