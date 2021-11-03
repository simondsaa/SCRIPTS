$s = New-PSSession -ComputerName "xlwul-42093d"
Invoke-Command -Session $s -Command {"C:\Temp\jiggler.ps1"}