Sleep 17
function NewPassword {
$randomObj = New-Object System.Random
$NewPassword=""
1..48 | ForEach { $NewPassword = $NewPassword + [char]$randomObj.next(33,126) }
return $NewPassword
}
$Password = $NewPassword
$admin=[adsi]"WinNT://$env:COMPUTERNAME/Administrator"
$admin.SetPassword($Password)