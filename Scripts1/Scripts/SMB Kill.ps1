#$Computers = Get-Content C:\Users\1180219788A\Desktop\BaseComputers.txt
$Computers = Read-Host "Computer Name"

ForEach ($Computer in $Computers)
    {
        If (Test-Connection $Computer -quiet -BufferSize 64 -Ea 0 -count 1)
            { 
                 REG ADD "\\$Computer\HKLM\SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters" /v SMB1 /t REG_DWORD /d 0 /f
            } 
            Else
                {
                    Write-Host -ForegroundColor Red "$Computer is not reachable." 
                    $Computer | Out-File -FilePath C:\Users\1180219788A\Desktop\SSL_Cache_Fail.txt -Append -Force
                }  
    }
EXIT