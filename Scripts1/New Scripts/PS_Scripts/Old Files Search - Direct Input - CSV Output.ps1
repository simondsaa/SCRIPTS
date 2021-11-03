$Date=(Get-Date).AddDays(-120)
Get-Childitem -Path "\\XLWU-FS-004\root\325 FW" -Recurse -ErrorAction SilentlyContinue |
Where-Object {$_.LastWriteTime -lt $Date} | Select Directory, Name, CreationTime, LastAccessTime, LastWriteTime |
Export-CSV C:\Users\timothy.brady\Desktop\Files.csv