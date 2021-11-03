#$Computers = Get-Content C:\Users\1180219788A\Desktop\Test.txt
$Computers = Read-Host "Computer Name"
ForEach ($Computer in $Computers)
    {
        If (Test-Connection $Computer -quiet -BufferSize 64 -Ea 0 -count 1)
            { 
                 REG ADD "\\$Computer\HKLM\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings" /v DisableCachingOfSSLPages /t REG_DWORD /d 0 /f
                 #REG DELETE "\\$Computer\HKLM\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings" /v DisableCachingOfSSLPages /f
            } 
            Else
                {
                    Write-Host -ForegroundColor Red "$Computer is not reachable." 
                    $Computer | Out-File -FilePath C:\Users\1180219788A\Desktop\SSL_Cache_Fail.txt -Append -Force
                }  
    }