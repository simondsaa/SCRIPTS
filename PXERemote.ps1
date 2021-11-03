$MyCred = Get-Credential
SetPCSVDeviceBootConfiguration -TargetAdress xlwul-42093d -ManagementProtocol WSMan -Credential $MyCred -OneTimeBootSource "CIM:Network:1"
