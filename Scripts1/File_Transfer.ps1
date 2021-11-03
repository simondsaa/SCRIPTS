$VNC = 'C:\Users\1252862141.adm\Desktop\sp81058.exe'
$Computer = 'xlwul-42093d'
$TMP = "\\$Computer\C$\Users\USAF_Admin\Desktop\Temp"

if (!(Test-Path $TMP)) {
    New-Item -Path $TMP -ItemType Directory
}

Copy-Item -LiteralPath $VNC -Destination $TMP -Container -Force -Verbose
Invoke-Command -ScriptBlock {cscript.exe $Using:LocalPath} -Computer $Computer