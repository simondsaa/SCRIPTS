New-LocalUser -Name "CST_Admin" -Description "Do NOT Delete. -101 ACOMS CSTs" -NoPassword 
Set-LocalUser CST_Admin -PasswordNeverExpires:$true
Net LocalGroup Administrators CST_Admin /Add
Get-LocalUser -Name "CST_Admin" | Enable-LocalUser
Net User CST_Admin zaq1XSW@zaq1XSW@
