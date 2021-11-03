$RootPath = "\\XLWU-FS-002\Tyndall$\Stats"
$Win7Path = "$RootPath\Current\Computer_Stats\Win7"
$WinVistaPath = "$RootPath\Current\Computer_Stats\WinVista"
$WinXPPath = "$RootPath\Current\Computer_Stats\WinXP"
$UserPath = "$RootPath\Current\User_Stats"

$Size = 100KB

$Win7Files = Get-ChildItem "$Win7Path" -Filter *.txt -Recurse -ErrorAction SilentlyContinue | Where-Object{$_.length -gt $Size}
ForEach ($Win7File in $Win7Files)
{
    $ArcPath = "$RootPath\Archives\Computer_stats\Win7"
    Move-Item "$Win7Path\$Win7File" "$ArcPath" -Force
}
$WinVistaFiles = Get-ChildItem "$WinVistaPath" -Filter *.txt -Recurse -ErrorAction SilentlyContinue | Where-Object{$_.length -gt $Size}
ForEach ($WinVistaFile in $WinVistaFiles)
{
    $ArcPath = "$RootPath\Archives\Computer_stats\WinVista"
    Move-Item "$WinVistaPath\$WinVistaFile" "$ArcPath" -Force
}
$WinXPFiles = Get-ChildItem "$WinXPPath" -Filter *.txt -Recurse -ErrorAction SilentlyContinue | Where-Object{$_.length -gt $Size}
ForEach ($WinXPFile in $WinXPFiles)
{
    $ArcPath = "$RootPath\Archives\Computer_stats\WinXP"
    Move-Item "$WinXPPath\$WinXPFile" "$ArcPath" -Force
}
$UserFiles = Get-ChildItem "$UserPath" -Filter *.txt -Recurse -ErrorAction SilentlyContinue | Where-Object{$_.length -gt $Size}
ForEach ($UserFile in $UserFiles)
{
    $ArcPath = "$RootPath\Archives\User_Stats"
    Move-Item "$UserPath\$UserFile" "$ArcPath" -Force
}