$Path = Read-Host "Path to PCs"
$Copy = Read-Host "Path to software package for install (will be installed in Temp folder on remote PC)"
$TaskName = Read-Host "Name the Task to be created in Task Scheduler"
$XML = Read-Host "Path to XML doc (required)"
$machines = Get-Content $Path
foreach ($comp in $machines) {
Copy-Item -Path $Copy -Recurse -Destination \\$comp\c$\temp -force
}
foreach ($mach in $machines) {
schtasks.exe /create /S $mach /TN $TaskName /RU  