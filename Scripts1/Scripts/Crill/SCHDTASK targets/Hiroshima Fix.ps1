$Message = "This is a mandatory reboot to resolve the Outlook issue, please save your work.

System will restart in 5min"

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
        Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        #Write-Host $KeyPath
        #Write-Host $OldPath
        #Write-Host $NewPath
        #Write-Host
    }
        
    If ($Key -like "SYSTEM\CurrentControlSet\Services\SamSs")
    {
        $NewPath = "%SystemRoot%\system32\lsass.exe"
        Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        #Write-Host $KeyPath
        #Write-Host $OldPath
        #Write-Host $NewPath
        #Write-Host
    }
}
Start-Sleep -Seconds 5
Start-Service -Name Server
Start-Service -Name Netlogon
Shutdown /r /f /c $Message /t 300