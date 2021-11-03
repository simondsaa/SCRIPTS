Remove-LocalGroupMember -group "Users" -Member "INTERACTIVE", "Authenticated Users", "Domain Users"
Add-LocalGroupMember -group "Users" -Member "1252862141N", "1252862141.adm"