$Path = Read-Host "Path to PCs"
$machines = Get-Content $Path
$Item = Read-Host "Path to file"
foreach ($comp in $machines) {
Get-ChildItem -Path \\$comp\c$\Users\*\Desktop | ?{ $_.PSIsContainer } |%{Copy-Item $Item $_.fullname}
}