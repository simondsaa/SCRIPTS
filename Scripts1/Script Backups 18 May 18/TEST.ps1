$VNC = '\\xlwul-42093d\C$\Users\USAF_Admin\desktop\test.vbs'
$Computer = 'xlwuw-491s64'
$TMP = "\\$Computer\c$\TEMP"

if (!(Test-Path $TMP)) {
    New-Item -Path $TMP -ItemType Directory
}

Copy-Item -LiteralPath (Split-Path $VNC -Parent) -Destination $TMP -Container -Recurse -Force -Verbose
$LocalPath = Join-Path 'C:\TEMP' (Join-Path (Split-Path $VNC -Parent | Split-Path -Leaf) (Split-Path $VNC -Leaf))
Invoke-Command -ScriptBlock {cscript.exe $Using:LocalPath} -Computer $Computer
