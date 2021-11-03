Write-Host "Results located here:  " -ForegroundColor Yellow -NoNewline
Write-Host "C:\Temp\File_Results_MULTI.csv " -ForegroundColor Green -NoNewline
Write-Host "& " -NoNewline
Write-Host "in a " -NoNewline
Write-Host "pop-up OGV" -ForegroundColor Green
Write-Host "Directory format (Example):  " -ForegroundColor Yellow -NoNewline
Write-Host "\Windows\System32\Name_of_File" -ForegroundColor Green -NoNewline 
Write-Host ""
$filename = "\windows\system32\crypt32.dll" 
 
$obj = New-Object System.Collections.ArrayList 
 
$Path = "C:\temp\ADPull - 30 Sept.txt"
$computers = Get-Content $Path
foreach ($server in $computers) 
    { 
        Invoke-Command -ComputerName $Server -Scriptblock {$filepath = Test-Path "\\$server\c$\$filename"} -AsJob | Get-Job | Receive-Job 
        if ($filepath -eq "True") { 
        $file = Get-Item "\\$server\c$\$filename" 
        $obj += New-Object psObject -Property @{'Computer'=$server;'FileVersion'=$file.VersionInfo|Select FileVersionraw;'LastAccessTime'=$file.LastWriteTime} 
        }else{
                $obj = New-Object PSObject
                $obj | Add-Member -Force -MemberType NoteProperty -Name "Computer" -Value $Server
                $obj | Add-Member -Force -MemberType NoteProperty -Name "FileVersion" -Value "Inaccesible" 
                write-host "$filename " -foregroundcolor green -NoNewline
                write-host "not found on " -NoNewline
                write-host "$Server" -foregroundcolor yellow
            }
    } 
     

$obj | select computer, FileVersion, lastaccesstime | Export-Csv -Path 'C:\Temp\File_Results_MULTI.csv' -NoTypeInformation 
$obj | select computer, FileVersion, lastaccesstime | OGV -Title "File Versions"
#  C:\temp\ADPull - 30 Sept.txt