$Comp = Read-Host "Enter Computer Name"
Invoke-command -ComputerName $Comp -ScriptBlock {
New-LocalUser -Name "usaf_admin" -Description "Do NOT Delete. -101 ACOMS CSTs" -NoPassword 
Set-LocalUser usaf_admin -PasswordNeverExpires:$true
Net LocalGroup Administrators usaf_admin /Add
Get-LocalUser -Name "usaf_admin" | Enable-LocalUser
Net User "usaf_admin" /active:yes
Net User usaf_admin 1zqa2xws!ZQA@XWS
}