﻿(Get-WmiObject -computername XLWUL-42093d -Namespace root/hp/instrumentedBIOS -Class HP_BIOSSettingInterface).SetBIOSSetting('Setup Password','<utf-16/>','<utf-16/>password')