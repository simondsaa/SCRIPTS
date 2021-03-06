#This is the path you want the results to be saved
$Path = "C:\Users\1274873341C\Desktop\Desktop\PS_Scripts\HBSS\Programs.xml"

#This is the path containing your list of Computers
$targetlist = Get-Content "C:\Users\1274873341C\Desktop\Desktop\PS_Scripts\HBSS\dat_targets.txt"
$Array = @()
ForEach($target in $targetlist){
    If (Test-Connection $target -quiet -count 1){
        $RegPath = "Software\\McAfee"
        $Reg = [microsoft.win32.registrykey]::OpenRemoteBaseKey('LocalMachine',$target)
        $ThisKey = $RegPath+"\\AVEngine"
        $ThisSubKey = $Reg.OpenSubKey($ThisKey)
            $obj = New-Object PSObject
            $obj | Add-Member -Force -MemberType NoteProperty -Name "ComputerName" -Value $target
            $obj | Add-Member -Force -MemberType NoteProperty -Name "DatDate" -Value $($ThisSubKey.GetValue("AVDatDate"))
            $obj | Add-Member -Force -MemberType NoteProperty -Name "DatVersion" -Value $($ThisSubKey.GetValue("AVDatVersion"))
            $Array += $obj}
    Else {Write-Host "$target Not reachable"}
}
$Array | Select ComputerName, DatDate, DatVersion | Export-CliXML $Path
$Result = Read-Host "Would you like to see your results now? Y or N"
If($Result -eq "Y"){Import-CliXML $Path | OGV} ElseIf($Result -eq "N"){Write-Host "Data has been saved here: $Path"}