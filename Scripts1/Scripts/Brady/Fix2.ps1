$Computer = $env:COMPUTERNAME
$LogPath = "\\xlwu-fs-05pv\Tyndall_PUBLIC\Stats\Quote Vuln"

"**************************************************" | Out-File "$LogPath\$Computer.log" -Force -Append
"Computername: $($Computer)" | Out-File "$LogPath\$Computer.log" -Append
"Date: $(Get-date -Format "dd.MM.yyyy HH:mm")" | Out-File "$LogPath\$Computer.log" -Force -Append

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
    $KeyPath = "HKLM:\"+$ThisKey

    If (($ImagePath -like "* *") -and ($ImagePath -notlike '"*"*') -and ($ImagePath -like '*.exe*'))
    { 
        $NewPath = ($ImagePath -split ".exe ")[0]
        $key1 = ($ImagePath -split ".exe ")[1]
        $triger = ($ImagePath -split ".exe ")[2]

        If (-not ($triger | Measure-Object).count -ge 1)
        {
            If (($NewPath -like "* *") -and ($NewPath -notlike "*.exe"))
            {
                " ----- Reg Path $KeyPath" | Out-File "$LogPath\$Computer.log" -Force -Append
                "" | Out-File "$LogPath\$Computer.log" -Force -Append
                " ***** Old Value $ImagePath" | Out-File "$LogPath\$Computer.log" -Force -Append
                " ***** New Value `"$NewPath.exe`" $key1" | Out-File "$LogPath\$Computer.log" -Force -Append
                "" | Out-File "$LogPath\$Computer.log" -Force -Append
                Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value "`"$NewPath`" $key1"
            }
        }

        If (-not ($triger | Measure-Object).count -ge 1)
        {
            If (($NewPath -like "* *") -and ($NewPath -like "*.exe"))
            {
                " ----- Reg Path $KeyPath" | Out-File "$LogPath\$Computer.log" -Force -Append
                "" | Out-File "$LogPath\$Computer.log" -Force -Append
                " ***** Old Value $ImagePath" | Out-File "$LogPath\$Computer.log" -Force -Append
                " ***** New Value `"$NewPath`"" | Out-File "$LogPath\$Computer.log" -Force -Append
                "" | Out-File "$LogPath\$Computer.log" -Force -Append
                Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value "`"$NewPath`""
            }
        }
        If (($triger | Measure-Object).count -ge 1)
        {
            "----- Error Cant parse $ImagePath in registry" | Out-File "$LogPath\$Computer.log" -Force -Append
        }
    }
}