#This is the path you want the results to be saved
$Path = "C:\Temp\test.xml"

#This is the path containing your list of Computers
$Computers = Get-Content "C:\Temp\SDC_Version.txt"
$Array = @()
ForEach($PC in $Computers){
    If (Test-Connection $PC -quiet -count 1){
        $RegPath = "Software\\Microsoft\\Windows\\CurrentVersion"
        $Reg = [microsoft.win32.registrykey]::OpenRemoteBaseKey('LocalMachine',$PC)
        $ThisKey = $RegPath+"\\OEMInformation"
        $ThisSubKey = $Reg.OpenSubKey($ThisKey)
            $obj = New-Object PSObject
            $obj | Add-Member -Force -MemberType NoteProperty -Name "ComputerName" -Value $PC
            $obj | Add-Member -Force -MemberType NoteProperty -Name "SDC" -Value $($ThisSubKey.GetValue("Model"))
            $Array += $obj}
    Else {Write-Host "$PC Not reachable"}
}
$Array | Select ComputerName, SDC | Export-CliXML $Path
$Result = Read-Host "Would you like to see your results now? Y or N"
If($Result -eq "Y"){Import-CliXML $Path | OGV} ElseIf($Result -eq "N"){Write-Host "Data has been saved here: $Path"}