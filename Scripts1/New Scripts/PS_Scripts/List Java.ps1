#$Computer = "52xlwul3-43002j"
$Good = 0

$Computers = Get-Content C:\Users\1392134782A\Desktop\Java.txt
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
            If ($Key -like "{26A24AE4-039D-4CA4-87B4*")
            {
                $ThisKey = $RegPath+"\"+$Key 
                $ThisSubKey = $Reg.OpenSubKey($ThisKey)
                $Java = $thisSubKey.GetValue("DisplayName")
                If ($Java -like "Java 8*")
                {
                    If ($Java -like "*51*")
                    {
                        Write-Host "$Computer - $Java" -ForegroundColor Green
                        $Good += 1
                    }
                    Else
                    {
                        Write-Host "$Computer - $Java" -ForegroundColor White
                    }
                }
                If ($Java -like "Java 7*")
                {
                    Write-Host "$Computer - $Java" -ForegroundColor Yellow
                }
                If (($Java -like "Java*") -and ($Java -like "* 6 *"))
                {
                     Write-Host "$Computer - $Java" -ForegroundColor DarkYellow
                }
                If (!($Java -like "Java 8 Update 51"))
                {
                    #"$Computer" | Out-File C:\Users\1392134782A\Desktop\JavaUpdate.txt -Append -Force
                }
            }
        }
    }
    Else
    {
        Write-Host "$Computer offline" -ForegroundColor Gray
    }
}

Write-Host "$Good / $Computers.Count"