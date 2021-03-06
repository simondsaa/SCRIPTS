$RegPath = "SYSTEM\CurrentControlSet\services"        
$Reg = [microsoft.win32.registrykey]::OpenRemoteBaseKey('LocalMachine',$Computer)
$RegKey = $Reg.OpenSubKey($RegPath)
$SubKeys = $RegKey.GetSubKeyNames()
ForEach ($SubKey in $SubKeys)
{
    $Key = $RegPath+"\"+$SubKey 
    $ThisSubKey = $Reg.OpenSubKey($Key)
    $KeyPath = "HKLM:\"+$Key
    $OldPath = $ThisSubKey.GetValue("ImagePath",$null,'DoNotExpandEnvironmentNames')
    If ($Key -like "SYSTEM\CurrentControlSet\Services\Netlogon")
    {
        $NewPath = "%systemroot%\system32\lsass.exe"
        #Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        Write-Host $Key
        Write-Host $OldPath
        Write-Host $NewPath
        Write-Host
    }
        
    If ($Key -like "SYSTEM\CurrentControlSet\Services\Netlogon")
    {
        $NewPath = "%SystemRoot%\system32\lsass.exe"
        #Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        Write-Host $Key
        Write-Host $OldPath
        Write-Host $NewPath
        Write-Host
    }
}