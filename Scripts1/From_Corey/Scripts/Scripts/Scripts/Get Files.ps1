$date=(Get-Date).AddDays(0)
Get-Childitem -Path "\\XLWU-FS-004\root\325 FW" -Recurse -ErrorAction SilentlyContinue |
Where-Object {$_.LastWriteTime -lt $date} | Select Directory, Name, CreationTime, LastAccessTime, LastWriteTime |
Export-CSV C:\Users\timothy.brady\Desktop\Files.csv