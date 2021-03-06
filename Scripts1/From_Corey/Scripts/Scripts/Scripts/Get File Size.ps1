Get-ChildItem "\\XLWU-FS-004\root\325 FW" -recurse |
Select Name, @{Name="Kbytes";Expression={"{0:N0}" -f ($_.Length/1Kb)}} |
Measure-Object -property Kbytes -sum | 
Select @{Name="Total Files"; Expression={$_.Count}}, @{Name="Total in KB"; Expression={$_.Sum}} |
Format-List