#---------------------------------------------------------------------------------
#                          Written by SrA Timothy Brady
#                          Tyndall AFB, Panama City, FL
#---------------------------------------------------------------------------------
$c = New-Object -Comobject wscript.shell
$b = $c.popup("Be sure to include \\ when entering the Search Directory",0,"TIP",0)
$body = "<TITLE>Files Found</TITLE><CENTER><H1>Files Found</H1></CENTER>"
$a = "<style>"
$a = $a + "BODY{background-color:palegoldenrod;}"
$a = $a + "TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}"
$a = $a + "TH{border-width: 1px;padding: 1px;border-style: solid;border-color: black;background-color:grey}"
$a = $a + "TD{border-width: 1px;padding: 1px;border-style: solid;border-color: black;background-color:white}"
$a = $a + "</style>"
$filePath = Read-Host "Search Directory"
$fileName = Read-Host "Filename or Filetype (*Computers.txt* or *.jpg*)"
Get-ChildItem -Recurse -Force $filePath -ErrorAction SilentlyContinue |
Where-Object {($_.PSIsContainer -eq $false) -and  ($_.Name -like "*$fileName*")} | 
Select-Object Name,Directory | ConvertTo-HTML -head $body -body $a | Out-File C:\Users\timothy.brady\Desktop\Files.html 