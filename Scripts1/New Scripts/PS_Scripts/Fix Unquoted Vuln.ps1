$LogPath = "C:\temp"

If (-not (Test-Path $LogPath)){New-Item $LogPath -ItemType Directory | Out-Null}

$Computers = Get-Content \\xlwu-fs-004\Home\1392134782A\Desktop\QuotedComputers.txt

ForEach ($Computer in $Computers)
{
    If (Test-Connection $Computer -Quiet -BufferSize 16 -Ea 0 -Count 1)
    {
        "**************************************************" | Out-File "$LogPath\servicesfix.log" -Append
        "Computername: $($Computer)" | Out-File "$LogPath\servicesfix.log" -Append
        "Date: $(Get-date -Format "dd.MM.yyyy HH:mm")" | Out-File "$LogPath\servicesfix.log" -Append

        # Get all services
        $RegPath = "SYSTEM\CurrentControlSet\Services\"
        $Reg = [microsoft.win32.registrykey]::OpenRemoteBaseKey('LocalMachine',$Computer)
        $RegKey = $Reg.OpenSubKey($RegPath)
        $SubKeys = $RegKey.GetSubKeyNames()
        ForEach ($Key in $SubKeys)
        {
            $ThisKey = $RegPath+$Key 
            $ThisSubKey = $Reg.OpenSubKey($ThisKey)
            $ImagePath = $ThisSubKey.GetValue("ImagePath",$null,'DoNotExpandEnvironmentNames')
            #Write-Host $ImagePath
            # Get all services with vulnerability
            If (($ImagePath -like "* *") -and ($ImagePath -notlike '"*"*') -and ($ImagePath -like '*.exe*'))
            { 
                $NewPath = ($ImagePath -split ".exe ")[0]
                $key1 = ($ImagePath -split ".exe ")[1]
                $triger = ($ImagePath -split ".exe ")[2]
    
                # Get all services with vulnerability with key in ImagePath
                If (-not ($triger | Measure-Object).count -ge 1)
                {
                    If (($NewPath -like "* *") -and ($NewPath -notlike "*.exe"))
                    {
        
                        " ***** Old Value $ImagePath" | Out-File "$LogPath\servicesfix.log" -Append
                        Write-Host "`"$NewPath.exe`" $key" | Out-File "$LogPath\servicesfix.log" -Append
                        Invoke-Command -ComputerName $Computer -ScriptBlock {Set-ItemProperty -Path "HKLM:\$($Args[0])" -Name "ImagePath" -Value "`"$($Args[1]).exe`" $($Args[2])"} -ArgumentList $ThisKey,$NewPath,$key1
                    }
                }

                # Get all services with vulnerability with out key in ImagePath
                If (-not ($triger | Measure-Object).count -ge 1)
                {
                    If (($NewPath -like "* *") -and ($NewPath -like "*.exe"))
                    {
        
                        " ***** Old Value $ImagePath" | Out-File "$LogPath\servicesfix.log" -Append
                        Write-Host "`"$NewPath.exe`"" | Out-File "$LogPath\servicesfix.log" -Append
                        Invoke-Command -ComputerName $Computer -ScriptBlock {Set-ItemProperty -Path "HKLM:\$($Args[0])" -Name "ImagePath" -Value "`"$($Args[1])`""} -ArgumentList $ThisKey,$NewPath
                    }
                }
                If (($triger | Measure-Object).count -ge 1)
                {
                    "----- Error Cant parse $ImagePath in registry" | Out-File $LogPath\servicesfix.log -Append
                }
            }
        }
    }
    Else
    {
        "**************************************************" | Out-File "$LogPath\servicesfix.log" -Append
        "Computername: $($Computer)" | Out-File "$LogPath\servicesfix.log" -Append
        "Date: $(Get-date -Format "dd.MM.yyyy HH:mm")" | Out-File "$LogPath\servicesfix.log" -Append
        "System offline..." | Out-File "$LogPath\servicesfix.log" -Append
    }
}