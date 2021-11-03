$Path = Read-Host "Path to PCs"
$machines = Get-Content $Path
foreach ($mach in $machines) {
schtasks.exe /create /S $mach /TN Reboot /XML 'E:\Reboot.xml'
}