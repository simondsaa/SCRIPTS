#Type in which Program you're searching for
$Program = "*"

#Type in the Version number you want to set as a baseline
$Version = "*"

#This is the path you want the results to be saved
$Path = "C:\Users\Timothy.Brady\Desktop\Programs.xml"

$Array = @()
#This is the path containing your list of Computers
$Computers = Get-Content "C:\Users\Timothy.Brady\Desktop\Comps.txt"
ForEach($PC in $Computers){
    If (Test-Connection $PC -quiet -count 1){
        $OS = Get-WmiObject Win32_OperatingSystem -cn $PC
        If ($OS.OSArchitecture -eq "64-bit"){$RegPath = "Software\\Wow6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall"}
        ElseIf ($OS.OSArchitecture -eq "32-bit"){$RegPath = "Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall"}        
        $Reg = [microsoft.win32.registrykey]::OpenRemoteBaseKey('LocalMachine',$PC)
        $RegKey = $Reg.OpenSubKey($RegPath)
        $SubKeys = $RegKey.GetSubKeyNames()
        ForEach($Key in $SubKeys){
            $ThisKey = $RegPath+"\\"+$Key 
            $ThisSubKey = $Reg.OpenSubKey($ThisKey) 
            $obj = New-Object PSObject
            $obj | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value $PC
            $obj | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $($thisSubKey.GetValue("DisplayName"))
            $obj | Add-Member -MemberType NoteProperty -Name "DisplayVersion" -Value $($thisSubKey.GetValue("DisplayVersion"))
            $obj | Add-Member -MemberType NoteProperty -Name "Publisher" -Value $($thisSubKey.GetValue("Publisher"))
            $obj | Add-Member -MemberType NoteProperty -Name "InstallDate" -Value $($thisSubKey.GetValue("InstallDate"))
            $obj | Add-Member -MemberType NoteProperty -Name "HelpLink" -Value $($thisSubKey.GetValue("HelpLink"))
            $obj | Add-Member -MemberType NoteProperty -Name "UninstallString" -Value $($thisSubKey.GetValue("UninstallString"))
            $Array += $obj}
    }
Else {Write-Host "$PC Not reachable"}
}
$Array | Where-Object {($_.DisplayName -like "*$Program*") -and ($_.DisplayVersion -ne "$Version")} | Select ComputerName, DisplayName, DisplayVersion, Publisher, InstallDate, HelpLink, UninstallString | Export-CliXML $Path
$Result = Read-Host "Would you like to see your results now? Y or N"
If($Result -eq "Y"){Import-CliXML $Path | OGV} ElseIf($Result -eq "N"){Write-Host "Data has been saved here: $Path"}
ElseIf($Result -eq "N"){Write-Host "Results are saved here: $Path"}