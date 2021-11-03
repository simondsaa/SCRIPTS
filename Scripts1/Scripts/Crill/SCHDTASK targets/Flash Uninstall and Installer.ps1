$ErrorActionPreference = 'SilentlyContinue'

Write-host "Uninstalling Flash Player"
WMIC product where "Name LIKE '%%Adobe Flash Player%%'" call uninstall 
Write-host "Installing Flash Player Active X"
Start-Process "\\xlwu-fs-05pv\Tyndall_PUBLIC\ncc admin\flash\install_flash_player_18_active_x.msi" /qn -wait
Write-Host "Installing Flash Player Plugin" 
Start-Process "\\xlwu-fs-05pv\Tyndall_PUBLIC\ncc admin\flash\install_flash_player_18_plugin.msi" /qn -wait

"Pushed $env:computername" | out-file -append "\\xlwu-fs-05pv\Tyndall_PUBLIC\ncc admin\flash\Flash_Install.txt"