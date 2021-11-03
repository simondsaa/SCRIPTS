$Computers = Get-Content "C:\Users\1180219788.adm\Desktop\NCC.txt"
#$Computers = Read-Host "Computer Name"

ForEach ($Computer in $Computers)
    {
        schtasks.exe /CREATE /SC ONLOGON /TN "Comply2Connect" /S "$Computer" /TR "cscript.exe '\\xlwu-fs-05pv\Tyndall_PUBLIC\NCC_Admin\Tyndall AFB C2C\c2c.vbs'" /F
        schtasks.exe /CREATE /TN "NOVA" /S $Computer /SC WEEKLY /D SAT /ST 12:00 /RL HIGHEST /RU SYSTEM /TR "powershell.exe -noprofile -File '\\xlwu-fs-05pv\Tyndall_PUBLIC\NCC_Admin\NOVA\Run_CSCC_4.1.1_Installing.ps1'" /F
        #schtasks.exe /DELETE /TN "Comply2Connect" /S "$Computer" /F
        #schtasks.exe /DELETE /TN "NOVA" /S "$Computer" /F
    }