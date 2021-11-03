$Path = Read-Host "Path to PCs"
$Computers = Get-Content $Path
foreach ($Comp in $Computers){
New-Item -ItemType "directory" -Path "\\$Comp\C$\Temp\CST_Help"
}