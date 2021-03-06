$PC = "52XLWUW3-DKPVV1"

#Type in which Program you're searching for
$Program = "*"

#Type in the Version number you want to set as a baseline
$Version = "*"

#This is the path you want the results to be saved
$Path = "C:\Users\Timothy.Brady\Desktop\Computer Programs\$PC Programs.xml"

$Array = @()
   If (Test-Connection $PC -quiet -BufferSize 16 -Ea 0 -count 1)
   {
        $OS = Get-WmiObject Win32_OperatingSystem -cn $PC
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
            $obj | Add-Member -Force -MemberType NoteProperty -Name "Computer Name" -Value $PC
            $obj | Add-Member -Force -MemberType NoteProperty -Name "Display Name" -Value $($thisSubKey.GetValue("DisplayName"))
            $obj | Add-Member -Force -MemberType NoteProperty -Name "Display Version" -Value $($thisSubKey.GetValue("DisplayVersion"))
            $obj | Add-Member -Force -MemberType NoteProperty -Name "Publisher" -Value $($thisSubKey.GetValue("Publisher"))
            $obj | Add-Member -Force -MemberType NoteProperty -Name "Install Date" -Value $($thisSubKey.GetValue("InstallDate"))
            $obj | Add-Member -Force -MemberType NoteProperty -Name "Help Link" -Value $($thisSubKey.GetValue("HelpLink"))
            $obj | Add-Member -Force -MemberType NoteProperty -Name "Uninstall String" -Value $($thisSubKey.GetValue("UninstallString"))
            $Array += $obj
        }
    }
    Else {Write-Host "$PC Not reachable"}
$Array | Where-Object {($_.DisplayName -like "*$Program*") -and ($_.DisplayVersion -ne "$Version")} | Export-CliXML $Path
Import-CliXML $Path | OGV -Title "$PC Programs"