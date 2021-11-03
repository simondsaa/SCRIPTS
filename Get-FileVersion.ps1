$filename = "\Windows\System32\crypt32.dll" 
 
$obj = New-Object System.Collections.ArrayList 
 
$Path = Read-Host "Path to PCs"
$computernames = Get-Content $Path
foreach ($server in $computernames) 
{ 
$filepath = Test-Path "\\$server\c$\$filename" 
 
if ($filepath -eq "True") { 
$file = Get-Item "\\$server\c$\$filename" 
 
     
        $obj += New-Object psObject -Property @{'Computer'=$server;'FileVersion'=$file.VersionInfo|Select FileVersionRaw;'LastAccessTime'=$file.LastWriteTime} 
        } 
     } 
     
$obj | select computer, FileVersion, lastaccesstime | OGV -Title "FileVersions" Export-Csv -Path 'C:\Temp\File_Results.csv' -NoTypeInformation 