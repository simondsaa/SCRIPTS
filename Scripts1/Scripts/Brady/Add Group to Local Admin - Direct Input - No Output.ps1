$Computer = "52XLWUW3-3614LR"
$objUser = [ADSI]("WinNT://AREA52/Tyndall Base Sysadmins")
$objGroup = [ADSI]("WinNT://$Computer/Administrators")
$objGroup.PSBase.Invoke("Add",$objUser.PSBase.Path)