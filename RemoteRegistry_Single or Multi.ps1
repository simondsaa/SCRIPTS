Function RemoteRegistrySingle
{
$error.clear()
    try
        {
            if (Test-Connection -count 1 -computer $computer -quiet)
                {
                    Write-Host "Updating: " $computer -ForegroundColor Green
                    Set-Service –Name RemoteRegistry –Computer $computer -StartupType Automatic
                    Get-Service -Name RemoteRegistry -Computer $computer | start-service
                    Write-Host ""
                    Invoke-Command -ComputerName $Computer -Scriptblock {
                        If ((Get-Service -Name RemoteRegistry).Status -eq 'Running'){
                            Write-Host "Service is running. Log: C:\Temp\G1sRegSuccess_Single.txt" -ForegroundColor Cyan}
                                }
                    Invoke-Command -ComputerName $Computer -ScriptBlock {
                        Write-Output $Computer | out-file -append -filepath "C:\Temp\G1sRegSuccess_Single.txt"
                                }
                    Write-Host ""
                    Stop
                        Function Stop
                            {}
                }
        }
    catch
        {
               $_ | Out-File "C:\Temp\G1sRegFailed_FAILSSS.txt" -Append
        }
                  
    If (!$error)
        {
                Write-Host "$Computer is not accessible. Logging failed PC here:  C:\Temp\G1sRegFailed_Single.txt" -ForegroundColor Yellow
                Write-Output $Computer | out-file -append -filepath "C:\Temp\G1sRegFailed_Single.txt"
        }
}

#=======================================================================================================================================

Function RemoteRegistryMulti
{
$Path = Read-Host "Path to PCs"
$Computers = Get-Content $Path
$error.clear()
foreach ($computer in $computers)
{
    Invoke-Command -Computername $Computer -Scriptblock {psexec.exe \\$Computer -s powershell Enable-PSRemoting -Force
}
foreach ($computer in $computers)
    {
         try
                {
                    if (Test-Connection -count 1 -computer $computer -quiet)
                        {
                            Write-Host ""
                            Write-Host "Updating: " $computer -ForegroundColor Green
                            Write-Host "Logging successfully modified PCs here:  C:\Temp\G1sRegSuccess_Multi.txt" -ForegroundColor Cyan
                            Write-Host "========================================================================" -NoNewline -ForegroundColor Black
                            Write-Host ""
                            Set-Service –Name RemoteRegistry –Computer $computer -StartupType Automatic
                            Get-Service -Name RemoteRegistry -Computer $computer | start-service
                            Write-Output $computer | out-file -append -filepath "C:\Temp\G1sRegSuccess_Multi.txt"
                            Stop
                                Function Stop
                                    {}
                        }
                }

            catch
                {
                        #?
                }
            If (!$error)
                {
                Write-Host "$Computer is not accessible.  Log: C:\Temp\G1sRegFailed_Multi.txt" -ForegroundColor Yellow
                Write-Output $Computer | out-file -append -filepath "C:\Temp\G1sRegFailed_Multi.txt"
                }
    }
}
}
#=======================================================================================================================================

        Write-Host " "
                    $PW = Write-Host "Press" -NoNewline 
                          Write-Host " 1" -ForegroundColor Green -NoNewline
                          Write-Host " to enable a" -NoNewline
                          Write-Host " SINGLE" -ForegroundColor Green -NoNewline
                          Write-Host " machine." -NoNewline
                          Write-Host " Press" -NoNewline 
                          Write-Host " 2" -ForegroundColor Cyan -NoNewline
                          Write-Host " to enable" -NoNewline
                          Write-Host " MULTIPLE" -ForegroundColor Cyan -NoNewline
                          Write-Host " machines:  " 
                    $MW = Read-Host
        If ($MW -eq 1){
            $Computer = Read-Host "Enter PC Name"
            Write-Host ""
            RemoteRegistrySingle
            }
        If ($MW -eq 2){
            Write-Host ""
            RemoteRegistryMulti
            }
