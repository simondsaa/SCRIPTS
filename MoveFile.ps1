$Path = Read-Host "Path to PCs"
$Source = Read-Host "File w/ extension. ex: test.txt"
$Computers = Get-Content $Path
foreach ($Comp in $Computers){
Move-Item -Path \\$Comp\C$\Temp\$Source -Destination \\$Comp\C$\Temp\CST_Help
}