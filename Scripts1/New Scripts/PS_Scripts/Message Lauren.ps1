$Computer = "XLWUW-430XXV"
$User = Get-WmiObject Win32_ComputerSystem -ComputerName $Computer
If ($User.UserName -ne $null)
{
    $EDI = $User.UserName.TrimStart("AREA52\")
    $UserInfo = (Get-ADUser "$EDI" -Properties DisplayName).DisplayName
}
$UserInfo
$Message = Read-Host "Message"
Msg Console /Server:$Computer $Message