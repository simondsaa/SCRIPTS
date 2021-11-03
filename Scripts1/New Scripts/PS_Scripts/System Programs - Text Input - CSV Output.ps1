$PCs = Get-Content "C:\Users\1392134782A\Documets\TestBed.txt"

#Type in which Program you're searching for
$Program = ""

#Type in the Version number you want excluded
$Version = ""

#This is the path you want the results to be saved
$Path = "C:\Users\1392134782A\Desktop\Programs.csv"

$sw = [Diagnostics.Stopwatch]::StartNew()
$Array = @()
ForEach ($PC in $PCs)
{   
    If (Test-Connection $PC -quiet -BufferSize 16 -Ea 0 -count 1)
    {
        $OS = Get-WmiObject Win32_OperatingSystem -cn $PC -ErrorAction SilentlyContinue
        If ($OS.OSArchitecture -eq "64-bit"){$RegPath = "Software\\Wow6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall"}
        ElseIf ($OS.OSArchitecture -eq "32-bit"){$RegPath = "Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall"}        
        $Reg = [microsoft.win32.registrykey]::OpenRemoteBaseKey('LocalMachine',$PC)
        $RegKey = $Reg.OpenSubKey($RegPath)
        $SubKeys = $RegKey.GetSubKeyNames()
        ForEach($Key in $SubKeys)
        {
            $ThisKey = $RegPath+"\\"+$Key 
            $ThisSubKey = $Reg.OpenSubKey($ThisKey)
            $obj = New-Object PSObject
            $obj | Add-Member -Force -MemberType NoteProperty -Name "ComputerName" -Value $PC
            $obj | Add-Member -Force -MemberType NoteProperty -Name "DisplayName" -Value $($thisSubKey.GetValue("DisplayName"))
            $obj | Add-Member -Force -MemberType NoteProperty -Name "DisplayVersion" -Value $($thisSubKey.GetValue("DisplayVersion"))
            $obj | Add-Member -Force -MemberType NoteProperty -Name "Publisher" -Value $($thisSubKey.GetValue("Publisher"))
            $obj | Add-Member -Force -MemberType NoteProperty -Name "InstallDate" -Value $($thisSubKey.GetValue("InstallDate"))
            $obj | Add-Member -Force -MemberType NoteProperty -Name "HelpLink" -Value $($thisSubKey.GetValue("HelpLink"))
            $obj | Add-Member -Force -MemberType NoteProperty -Name "UninstallString" -Value $($thisSubKey.GetValue("UninstallString"))
            $Array += $obj
        }
    }
    Else {Write-Host "$PC Not reachable"}
}
$sw.Stop()
$Time = $sw.Elapsed | Select Minutes, Seconds
Write-Host "Execution time:" $Time.Minutes"min"$Time.Seconds"sec"

$Array | Where-Object {($_.DisplayName -like "*$Program*") -and ($_.DisplayVersion -ne "$Version")} | Select ComputerName, DisplayName, DisplayVersion, Publisher, InstallDate, HelpLink, UninstallString | Export-Csv $Path
Import-Csv $Path | OGV -Title "$PC Programs"