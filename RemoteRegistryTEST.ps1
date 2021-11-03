
$Path = Read-Host "Path to List of PCs"
$Computers = Get-Content $Path
$error.clear()
foreach ($computer in $computers)
    {
        {
            try
                {
                    if (Test-Connection -count 1 -computer $computer -quiet)
                        {
                            Write-Host "Updating system" $computer "....." -ForegroundColor Green
                            Set-Service –Name RemoteRegistry –Computer $computer -StartupType Automatic
                            Get-Service -Name RemoteRegistry -Computer $computer | start-service
                            Write-Output $computer | out-file -append -filepath "C:\Temp\G1sRegSuccess_Multi.txt"
                        }
                }
            catch
                {
                        #?
                }
            If (!$error)
                {
                Write-Host "$Computer is not accessible. Logging failed PC here:  C:\Temp\G1sRegFailed_Multi.txt" -ForegroundColor Red
                Write-Output $computer | out-file -append -filepath "C:\Temp\G1sRegFailed_Multi.txt"
                }
        }
    }
