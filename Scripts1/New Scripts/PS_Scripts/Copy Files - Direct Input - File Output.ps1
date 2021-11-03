$Directory = "\\XLWU-FS-002\Tyndall$\Stats\User_Stats"
$Names = Get-Content "C:\Users\timothy.brady\Desktop\Names.txt"
ForEach ($Name in $Names)
{
    $Files = Get-ChildItem -Recurse -Force $Directory -ErrorAction SilentlyContinue | Where-Object {($_.PSIsContainer -eq $false) -and  ($_.Name -like "*$Name*")}
    $Files | Copy-Item -Destination "C:\Users\timothy.brady\Desktop\Test"
}