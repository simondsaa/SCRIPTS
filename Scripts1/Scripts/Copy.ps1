$Computers = Get-Content "C:\Users\1180219788A\Desktop\NCC.txt"
#$Script_Path = "\\xlwu-fs-05pv\Tyndall_PUBLIC\NCC_Admin\Tyndall AFB C2C"

ForEach ($Computer in $Computers)
{ 
#robocopy $Script_Path \\$Computer\c$\ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp c2c.vbs 
Remove-Item "\\$Computer\c$\ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp\c2c.vbs" -Force

}