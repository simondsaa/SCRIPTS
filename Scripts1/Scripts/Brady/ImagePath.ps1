$Computers = Get-Content \\xlwu-fs-05pv\Tyndall_PUBLIC\Test.txt

ForEach ($Computer in $Computers)
{
    If (Test-Connection $Computer -Quiet -BufferSize 16 -Ea 0 -Count 1)
    {
        $RegPath = "SYSTEM\CurrentControlSet\services"        
        Try
        {
            $Reg = [microsoft.win32.registrykey]::OpenRemoteBaseKey('LocalMachine',$Computer)
            $RegKey = $Reg.OpenSubKey($RegPath)
            $SubKeys = $RegKey.GetSubKeyNames()
            $Array = @()
            ForEach($Key in $SubKeys)
            {
                If ($Key -like "Netlogon")
                {
                
                    $ThisKey = $RegPath+"\"+$Key 
                    $ThisSubKey = $Reg.OpenSubKey($ThisKey)
                    $ImagePath = $thisSubKey.GetValue("ImagePath",$null,'DoNotExpandEnvironmentNames')
                    Write-Host "$Computer - $ImagePath"
                }
            }
        }
        Catch
        {
            "$Computer" | Out-File \\XLWU-FS-05pv\Tyndall_PUBLIC\ImagePathBad.txt -Append -Force
        }
    }
    Else
    {
        Write-Host "$Computer offline"
    }
}