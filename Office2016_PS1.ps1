$Path = Read-Host "Path to PCs"
$machines = Get-Content $Path
foreach ($comp in $machines) {
Copy-Item -Path 'E:\Office 2016' -Recurse -Destination \\$comp\c$\temp -force
}
foreach ($mach in $machines) {
schtasks.exe /create /S $mach /TN MSOffice16_Install /XML 'E:\Office 2016\Office2016_XML.xml'
}