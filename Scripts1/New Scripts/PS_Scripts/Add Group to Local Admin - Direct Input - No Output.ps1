$Computer = "52XLWUW3-3614LR"        #The Domain and Group name you want to add
    $objUser = [ADSI]("WinNT://AREA52/Tyndall Base Sysadmins")
    $objGroup = [ADSI]("WinNT://$Computer/Administrators")
    $objGroup.PSBase.Invoke("Add",$objUser.PSBase.Path)