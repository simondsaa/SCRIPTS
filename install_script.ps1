$machines = Get-Content C:\temp\visiomachines.txt
foreach ($comp in $machines) {
Copy-Item -Path C:\temp\Visio2010 -Recurse -Destination \\$comp\c$\temp -force
}
foreach ($mach in $machines) {
schtasks.exe /create /S $mach /TN Visio_SP /XML 'C:\temp\Visio2010\Visio after hours install.xml'
}