$Computers = Get-Content C:\work\TEST.txt
$Good = @()
ForEach ($Computer in $Computers)
{
    If (Test-Connection $Computer -Quiet -BufferSize 16 -Ea 0 -Count 1)
    {
        $OSInfo = Get-Wmiobject Win32_OperatingSystem -ComputerName $Computer -ErrorAction SilentlyContinue
        If ($OSInfo.OSArchitecture -eq "64-bit"){$RegPath = "Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"}
        ElseIf ($OSInfo.OSArchitecture -eq "32-bit"){$RegPath = "Software\Microsoft\Windows\CurrentVersion\Uninstall"}        
        $Reg = [microsoft.win32.registrykey]::OpenRemoteBaseKey('LocalMachine',$Computer)
        $RegKey = $Reg.OpenSubKey($RegPath)
        $SubKeys = $RegKey.GetSubKeyNames()
        ForEach($Key in $SubKeys)
        {
            If ($Key -like "{40391824-FDD7-4AE0*")
            {
                $ThisKey = $RegPath+"\"+$Key 
                $ThisSubKey = $Reg.OpenSubKey($ThisKey)
                $UC = $thisSubKey.GetValue("DisplayName")

                $obj = New-Object PSObject
                $obj | Add-Member -Force -MemberType NoteProperty -Name "Computer" -Value $Computer
                $Good += $obj
                Write-Host "$Computer - $UC"
            }
        }
    }
}