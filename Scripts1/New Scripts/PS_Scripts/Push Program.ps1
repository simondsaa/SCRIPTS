$File = "DISAGVSDodAdminInstaller-win32-TAG_VD_3_5_6_03D.exe"

$Deploy = {
   Set-Location "C:\"
   & msiexec.exe /I GVS /quiet /norestart
}

$Computers = Get-Content "C:\work\TEST.txt"

ForEach ($Computer in $Computers)
{
    Copy-Item -Path "\\XLWU-FS-01pv\Tyndall_ANG\Shared\_14 CST Help (Do NOT Remove)\DISAGVSDodAdminInstaller-win32-TAG_VD_3_5_6_03D.exe" \\$Computer\C$ -Force
    $Session = New-PSSession -ComputerName $Computer
    Invoke-Command -Session $Session -ScriptBlock $Deploy
    Start-Sleep -Seconds 10
    Remove-PSSession -ComputerName $Computer
    Remove-Item -Path \\$Computer\C$\$File -Force
}