$Computers = Get-Content C:\Temp\ComputersPINGING.txt
foreach ($comp in $Computers){
invoke-command -computername $comp -scriptblock {cmd.exe /c winrm quickconfig -q}}